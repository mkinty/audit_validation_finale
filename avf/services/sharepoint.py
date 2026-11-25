"""Audit depuis SharePoint : la version **serveur** des fichiers, pas la copie OneDrive.

Pourquoi : une modification faite dans le dossier OneDrive local peut ne jamais
atteindre SharePoint (conflit de synchronisation, copie « -NOMDUPC », envoi en
échec). Auditer le dossier local montrerait alors un état que personne d'autre
ne voit. Ici, l'utilisateur colle le lien SharePoint du dossier racine et
l'audit lit ce que contient réellement le serveur.

Accès principal — ``api`` : connexion par le navigateur
    « Se connecter » ouvre Microsoft Edge (profil dédié à l'outil) sur le
    dossier : l'utilisateur s'y identifie normalement (MFA comprise). La
    session obtenue sert ensuite à interroger l'**API SharePoint** du site
    (``services.sharepoint_rest``) : liste des ``Dep<XX>/<commune>``,
    téléchargement de chaque ``audit_*.xlsx`` — la version du serveur, sans
    WebDAV, sans synchronisation, sans Excel. Le profil étant conservé, la
    connexion suivante est silencieuse tant que la session SharePoint est valide.

Si aucun navigateur n'est disponible, trois accès de repli, sans inscription
d'application ni droits d'administrateur :

``webdav`` — accès direct (préféré)
    Le lien devient un chemin réseau WebDAV ``\\\\hôte@SSL\\DavWWWRoot\\…``,
    celui qu'utilise « Afficher dans l'Explorateur de fichiers » de SharePoint.
    Windows (service WebClient) s'authentifie avec la session Microsoft 365 du
    poste. Tout l'audit fonctionne alors **à l'identique** : lister les
    ``Dep<XX>/<commune>``, lire chaque ``audit_*.xlsx``… directement sur le serveur.

``excel`` — repli via Excel
    Si WebDAV est refusé, mais que la bibliothèque est synchronisée par OneDrive
    sur le poste : la liste des communes vient du dossier synchronisé (la
    structure des dossiers), mais chaque fichier audit est **ouvert par son
    adresse SharePoint** dans l'instance Excel dédiée (compte Office de
    l'utilisateur), copié dans un dossier temporaire puis lu. Le contenu audité
    est donc celui du serveur. Plus lent (une ouverture Excel par commune).

``url`` — via Excel, sans synchronisation (dernier recours)
    Ni WebDAV ni synchronisation : le dossier ne peut pas être parcouru. Les
    communes (ou ID erreur) sont saisies, et le fichier de chaque commune est
    ouvert par son adresse ``<lien>/Dep<XX>/<commune>/audit_<commune>.xlsx``
    (nommage des fichiers audit). Pas de détection des copies de conflit.

Avec Excel, si Office demande la connexion Microsoft 365 (compte
professionnel), une fenêtre Excel visible s'ouvre une fois pour que
l'utilisateur se connecte ; les ouvertures suivantes sont silencieuses
(``SessionExcel(connexion_interactive=True)``).

La mise à jour, en source SharePoint, passe toujours par Excel (co-édition) :
jamais de réécriture directe du fichier sur le serveur.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
import os
import re
import shutil
import tempfile
import threading
import time
from urllib.parse import parse_qs, quote, unquote, urlsplit

from sna.services.fichiers import detecter_fichiers_conflit, trouver_fichier_audit
from sna.services.onedrive import montages_onedrive, url_sharepoint

from ..models.source import API, EXCEL, URL, WEBDAV, SourceAudit
from . import sharepoint_rest as rest
from .navigateur import (
    NavigateurFerme, NavigateurIndisponible, NavigateurPilote, cookies_pour, trouver_navigateur,
)

_RACINE_WEBDAV = re.compile(r"^\\\\(?P<hote>[^\\@]+)@SSL(?:@\d+)?\\DavWWWRoot(?P<chemin>\\.*)?$", re.I)
# Liens de partage : https://hote/:f:/r/sites/X/... (« r » = chemin réel, résoluble)
_PARTAGE_RESOLUBLE = re.compile(r"^/:[a-z]:/r(/.*)$", re.I)
_PARTAGE_OPAQUE = re.compile(r"^/:[a-z]:/[a-z]/", re.I)


class LienInvalide(ValueError):
    """Le texte saisi n'est pas un lien SharePoint de dossier exploitable."""


class AccesImpossible(RuntimeError):
    """Ni WebDAV ni bibliothèque synchronisée : le dossier SharePoint n'est pas joignable."""


@dataclass(frozen=True)
class LienSharePoint:
    """Dossier SharePoint : hôte et chemin serveur décodé (``/sites/X/Shared Documents/Audit``)."""

    hote: str
    chemin: str

    @property
    def url(self) -> str:
        return f"https://{self.hote}{quote(self.chemin, safe='/')}"

    @property
    def chemin_webdav(self) -> str:
        return f"\\\\{self.hote}@SSL\\DavWWWRoot" + self.chemin.replace("/", "\\")


# ── Analyse du lien saisi ─────────────────────────────────────────
def analyser_lien(texte: str) -> LienSharePoint:
    """Lien collé par l'utilisateur → dossier SharePoint.

    Acceptés :

    - l'adresse du dossier : ``https://x.sharepoint.com/sites/S/Shared%20Documents/Audit`` ;
    - la barre d'adresse du navigateur sur la bibliothèque :
      ``…/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FS%2F…&viewid=…`` ;
    - un lien de partage « personnes de l'organisation » au chemin lisible :
      ``https://x.sharepoint.com/:f:/r/sites/S/Shared%20Documents/Audit?csf=1&web=1&e=…`` ;
    - un chemin WebDAV ``\\\\x.sharepoint.com@SSL\\DavWWWRoot\\sites\\S\\…``.

    Un lien de partage opaque (``/:f:/s/S/EaBc…``) ne contient pas le chemin du
    dossier : il est refusé avec la marche à suivre.
    """
    brut = (texte or "").strip().strip('"').strip()
    if not brut:
        raise LienInvalide("aucun lien saisi")
    webdav = _RACINE_WEBDAV.match(brut)
    if webdav:
        chemin = (webdav.group("chemin") or "").replace("\\", "/")
        return LienSharePoint(webdav.group("hote").lower(), _nettoyer(unquote(chemin)))
    if not re.match(r"^https?://", brut, re.I):
        raise LienInvalide("le lien doit commencer par https:// (adresse SharePoint du dossier racine)")
    morceaux = urlsplit(brut)
    hote = (morceaux.hostname or "").lower()
    if not hote:
        raise LienInvalide("adresse sans nom de serveur")
    chemin = morceaux.path

    identifiant = parse_qs(morceaux.query).get("id") or parse_qs(morceaux.query).get("RootFolder")
    if identifiant:                                  # …/Forms/AllItems.aspx?id=/sites/…
        return LienSharePoint(hote, _nettoyer(unquote(identifiant[0])))
    if _PARTAGE_OPAQUE.match(chemin) and not _PARTAGE_RESOLUBLE.match(chemin):
        raise LienInvalide(
            "lien de partage sans chemin de dossier — ouvrez le dossier dans SharePoint et "
            "copiez l'adresse de la barre du navigateur"
        )
    partage = _PARTAGE_RESOLUBLE.match(chemin)
    if partage:                                      # /:f:/r/sites/…
        chemin = partage.group(1)
    chemin = unquote(chemin)
    if "/Forms/" in chemin:                          # vue de bibliothèque sans « id » : sa racine
        chemin = chemin.split("/Forms/")[0]
    elif chemin.lower().endswith(".aspx"):
        raise LienInvalide("ce lien désigne une page, pas un dossier de documents")
    chemin = _nettoyer(chemin)
    if chemin.count("/") < 2:
        raise LienInvalide("le lien doit désigner un dossier d'une bibliothèque (…/sites/<site>/<bibliothèque>/…)")
    return LienSharePoint(hote, chemin)


def _nettoyer(chemin: str) -> str:
    chemin = "/" + chemin.strip().strip("/")
    return re.sub(r"/{2,}", "/", chemin)


def url_depuis_webdav(chemin: str) -> str | None:
    """``\\\\hôte@SSL\\DavWWWRoot\\…`` → adresse https correspondante (``None`` sinon)."""
    return url_sharepoint(chemin, montages=[])   # sans montage : seul un chemin WebDAV est converti


def dossier_synchronise(lien: LienSharePoint, montages: list[tuple[str, str]] | None = None) -> str | None:
    """Dossier local OneDrive où ce dossier SharePoint est synchronisé, sinon ``None``."""
    montages = montages_onedrive() if montages is None else montages
    cible = f"https://{lien.hote}{lien.chemin}".rstrip("/")
    meilleur: tuple[str, int] | None = None
    for local, url in montages:
        base = unquote(url).rstrip("/")           # le registre peut stocker l'URL encodée ou non
        dedans = cible.casefold() == base.casefold() or cible.casefold().startswith(base.casefold() + "/")
        if dedans and (meilleur is None or len(base) > meilleur[1]):
            meilleur = (local, len(base))
    if meilleur is None:
        return None
    relatif = cible[meilleur[1]:].strip("/")      # casse réelle du lien
    return os.path.join(meilleur[0], *relatif.split("/")) if relatif else meilleur[0]


# ── Connexion par le navigateur (accès principal) ─────────────────
_CLIENTS: dict[str, rest.ClientSharePoint] = {}   # session établie, par hôte SharePoint
_VERROU_CLIENTS = threading.Lock()


def client_pour(lien: str) -> rest.ClientSharePoint:
    """Client de l'API SharePoint ouvert par :func:`connecter_navigateur` pour ce lien."""
    hote = (urlsplit(lien).hostname or "").lower()
    with _VERROU_CLIENTS:
        client = _CLIENTS.get(hote)
    if client is None:
        raise rest.AuthentificationRequise("pas de session SharePoint : cliquez sur « Se connecter »")
    return client


def enregistrer_client(client: rest.ClientSharePoint) -> None:
    with _VERROU_CLIENTS:
        _CLIENTS[client.hote] = client


def oublier_connexion() -> None:
    """Oublie les sessions de l'outil (mémoire et profil du navigateur dédié)."""
    from .navigateur import dossier_profil
    with _VERROU_CLIENTS:
        _CLIENTS.clear()
    shutil.rmtree(dossier_profil(), ignore_errors=True)


def connecter_navigateur(
    texte: str,
    on_etape: Callable[[str], None] | None = None,
    interactif: bool = True,
    delai_s: float = 300,
    delai_silencieux_s: float = 15,
    fabrique: Callable[..., NavigateurPilote] = NavigateurPilote,
    client_factory: Callable[[str, dict], rest.ClientSharePoint] = rest.ClientSharePoint,
) -> SourceAudit:
    """Connexion au dossier SharePoint avec l'identité de l'utilisateur, via le navigateur.

    1. Sans fenêtre : le profil dédié a peut-être encore une session valide.
    2. Sinon (``interactif``) : une fenêtre du navigateur s'ouvre sur le dossier ;
       l'utilisateur s'identifie ; dès que la session ouvre le dossier par l'API,
       la fenêtre se ferme et l'outil est connecté.

    Raises:
        LienInvalide: lien inexploitable.
        NavigateurIndisponible: ni Edge ni Chrome.
        NavigateurFerme: fenêtre fermée avant la fin de la connexion.
        AccesImpossible: délai dépassé, dossier introuvable, ou session sans droits.
        rest.AuthentificationRequise: non interactif et aucune session valide.
    """
    lien = analyser_lien(texte)

    def dire(message: str) -> None:
        if on_etape is not None:
            on_etape(message)

    def essayer(nav: NavigateurPilote) -> rest.ClientSharePoint | None:
        client = client_factory(lien.hote, cookies_pour(nav.cookies(), lien.hote))
        try:
            client.verifier(lien.chemin)
        except rest.AuthentificationRequise:
            return None
        except FileNotFoundError as exc:
            raise AccesImpossible(
                f"connecté à SharePoint, mais le dossier est introuvable : {lien.chemin} "
                "— vérifiez le lien (il doit ouvrir le dossier racine dans le navigateur)"
            ) from exc
        return client

    def attendre(nav: NavigateurPilote, limite_s: float, pas_s: float) -> rest.ClientSharePoint | None:
        fin = time.monotonic() + limite_s
        while time.monotonic() < fin:
            if not nav.ouvert():
                raise NavigateurFerme("fenêtre du navigateur fermée avant la fin de la connexion")
            try:
                client = essayer(nav)
            except NavigateurFerme:
                raise
            except (OSError, RuntimeError) as exc:   # navigateur pas encore prêt, réseau…
                if isinstance(exc, (AccesImpossible, NavigateurIndisponible)):
                    raise
                client = None
            if client is not None:
                return client
            time.sleep(pas_s)
        return None

    dire("vérification d'une session SharePoint existante…")
    with fabrique(lien.url, visible=False) as nav:
        client = attendre(nav, delai_silencieux_s, 1.0)
    if client is None:
        if not interactif:
            raise rest.AuthentificationRequise("connexion SharePoint requise : cliquez sur « Se connecter »")
        dire("🔑 identifiez-vous dans la fenêtre du navigateur (compte professionnel) — "
             "cochez « Rester connecté » pour ne plus avoir à le faire")
        with fabrique(lien.url, visible=True) as nav:
            client = attendre(nav, delai_s, 2.0)
            if client is None:
                raise AccesImpossible(f"connexion SharePoint non terminée en {delai_s / 60:.0f} min")
            dire("✅ identification réussie — fermeture du navigateur")
    enregistrer_client(client)
    return SourceAudit(API, lien.url, lien.url)


def navigateur_disponible() -> bool:
    return trouver_navigateur() is not None


# ── Connexion de repli (sans navigateur) ──────────────────────────
def connecter(
    texte: str,
    delai_s: float = 45,
    montages: list[tuple[str, str]] | None = None,
    est_dossier: Callable[[str], bool] | None = None,
    excel_present: Callable[[], bool] | None = None,
) -> SourceAudit:
    """Trouve un accès au dossier SharePoint : WebDAV, bibliothèque synchronisée,
    sinon ouverture des fichiers par adresse via Excel (mode ``url``).

    ``SourceAudit.diagnostic`` explique pourquoi l'accès direct n'a pas été retenu.

    Raises:
        LienInvalide: lien inexploitable.
        AccesImpossible: aucun accès possible (avec la marche à suivre).
    """
    lien = analyser_lien(texte)
    webdav = lien.chemin_webdav
    est_dossier = est_dossier or _dossier_accessible
    ok, erreur = _avec_delai(lambda: est_dossier(webdav), delai_s)
    if ok:
        return SourceAudit(WEBDAV, webdav, lien.url)
    cause = _cause_webdav(erreur)
    local = dossier_synchronise(lien, montages)
    if local and _sans_erreur(est_dossier, local):
        return SourceAudit(EXCEL, local, lien.url, diagnostic=cause)
    if (excel_present or _excel_present)():
        return SourceAudit(URL, lien.url, lien.url, diagnostic=cause)
    raise AccesImpossible(
        f"dossier SharePoint injoignable — accès direct refusé : {cause}\n\n"
        f"Accès direct (WebDAV) : {webdav}\n"
        "• Vérifiez le lien (il doit ouvrir le dossier racine dans le navigateur).\n"
        "• Ouvrez le dossier dans SharePoint avec Microsoft Edge, puis « Afficher dans "
        "l'Explorateur de fichiers » : Windows mémorise alors la connexion.\n"
        "• Ou synchronisez la bibliothèque avec OneDrive (bouton « Synchroniser » de SharePoint).\n"
        "• Ou installez Excel : les fichiers seraient ouverts sur le serveur par leur adresse."
    )


def _dossier_accessible(chemin: str) -> bool:
    """Lister le dossier, pour obtenir la vraie cause d'un refus (``os.path.isdir`` la tait)."""
    os.listdir(chemin)
    return True


def _sans_erreur(fonction: Callable[[str], bool], chemin: str) -> bool:
    try:
        return bool(fonction(chemin))
    except OSError:
        return False


def _excel_present() -> bool:
    from sna.services.excel_com import excel_disponible
    return excel_disponible()


def _cause_webdav(erreur: str | None) -> str:
    """Message Windows → explication actionnable."""
    if erreur == "delai":
        return "le serveur n'a pas répondu à temps"
    texte = (erreur or "accès refusé ou dossier introuvable").strip()
    if "sites approuvés" in texte or "trusted sites" in texte.lower():
        return ("Windows n'a pas de session SharePoint pour l'accès direct (erreur 224) — "
                "ouvrez le dossier dans Edge puis « Afficher dans l'Explorateur de fichiers »")
    return texte


def _avec_delai(fonction: Callable[[], bool], delai_s: float) -> tuple[bool, str | None]:
    """Un accès WebDAV non authentifié peut rester bloqué longtemps : on ne l'attend pas plus de ``delai_s``."""
    resultat: list = []

    def tache():
        try:
            resultat.append((bool(fonction()), None))
        except OSError as exc:
            resultat.append((False, exc.strerror or str(exc)))
        except Exception as exc:  # noqa: BLE001
            resultat.append((False, str(exc)))

    fil = threading.Thread(target=tache, daemon=True, name="avf-webdav")
    fil.start()
    fil.join(delai_s)
    return resultat[0] if resultat else (False, "delai")


# ── Accès aux fichiers pendant l'audit ────────────────────────────
class AccesFichiers:
    """Comment lire un fichier audit trouvé dans la source, et comment l'afficher.

    Source locale ou WebDAV : le fichier se lit tel quel. L'adresse affichée est
    l'URL SharePoint quand elle est connue.
    """

    def __enter__(self) -> "AccesFichiers":
        return self

    def __exit__(self, *exc) -> bool:
        return False

    def trouver(self, racine: str, commune: str) -> str | None:
        """Fichier audit de la commune (``None`` : aucun)."""
        return trouver_fichier_audit(racine, commune)

    def conflits(self, racine: str, commune: str) -> list[dict]:
        """Copies de conflit OneDrive/SharePoint présentes à côté du fichier."""
        return detecter_fichiers_conflit(racine, commune)

    def lisible(self, chemin: str) -> str:
        """Chemin d'un fichier à lire, au contenu identique à celui de la source."""
        return chemin

    def adresse(self, chemin: str) -> str:
        return url_depuis_webdav(chemin) or chemin


class AccesViaExcel(AccesFichiers):
    """Source ``excel`` : chaque fichier est ouvert **par son URL SharePoint** dans
    une instance Excel dédiée (lecture seule), puis copié dans un dossier
    temporaire supprimé à la fin. Le contenu lu est celui du serveur."""

    def __init__(self, source: SourceAudit, session_factory=None,
                 montages: list[tuple[str, str]] | None = None):
        if session_factory is None:
            session_factory = session_excel_interactive
        self.source = source
        self._montages = montages   # None : bibliothèques synchronisées du registre
        self._fabrique = session_factory
        self._session = None
        self._dossier = ""
        self._compteur = 0
        self._verrou = threading.Lock()

    def __enter__(self) -> "AccesViaExcel":
        self._dossier = tempfile.mkdtemp(prefix="avf_sharepoint_")
        self._session = self._fabrique()
        self._session.__enter__()
        return self

    def __exit__(self, *exc) -> bool:
        try:
            if self._session is not None:
                self._session.__exit__(*exc)
        finally:
            self._session = None
            shutil.rmtree(self._dossier, ignore_errors=True)
        return False

    def adresse(self, chemin: str) -> str:
        return url_sharepoint(chemin, self._montages) or chemin

    def lisible(self, chemin: str) -> str:
        url = self._url(chemin)
        if url is None:
            raise AccesImpossible(f"adresse SharePoint inconnue pour {chemin} (dossier non synchronisé ?)")
        with self._verrou:
            self._compteur += 1
            nom = url.rsplit("/", 1)[-1].replace("%20", " ")
            destination = os.path.join(self._dossier, f"{self._compteur:05d}_{nom}")
        try:
            self._session.copier_version_serveur(url, destination)
        except Exception as exc:
            if _est_introuvable(exc):
                raise FichierIntrouvable(f"fichier introuvable sur SharePoint : {url}") from None
            raise
        return destination

    def _url(self, chemin: str) -> str | None:
        return url_sharepoint(chemin, self._montages)


class AccesParAdresse(AccesViaExcel):
    """Source ``url`` : le dossier ne se parcourt pas ; le fichier d'une commune est
    désigné par son adresse (``SourceAudit.url_fichier``) et ouvert via Excel."""

    def trouver(self, racine: str, commune: str) -> str | None:
        return self.source.url_fichier(commune)

    def conflits(self, racine: str, commune: str) -> list[dict]:
        return []   # non détectables sans parcourir le dossier

    def adresse(self, chemin: str) -> str:
        return chemin

    def _url(self, chemin: str) -> str | None:
        return chemin


class FichierIntrouvable(FileNotFoundError):
    """Le fichier audit n'existe pas à l'adresse attendue sur SharePoint."""


def _est_introuvable(exc: BaseException) -> bool:
    texte = " ".join(str(a) for a in getattr(exc, "args", ()) if a).lower()
    return any(m in texte for m in ("introuvable", "trouvons pas", "n'existe pas", "n’existe pas",
                                     "couldn't find", "could not find", "not found", "does not exist"))


def session_excel_interactive():
    """Session Excel qui montre la fenêtre de connexion Microsoft 365 si Office la demande."""
    from sna.services.excel_com import SessionExcel
    return SessionExcel(connexion_interactive=True)


class AccesApi(AccesFichiers):
    """Source ``api`` : dossiers listés et fichiers téléchargés par l'API SharePoint,
    avec la session du navigateur. Les copies temporaires sont supprimées à la fin."""

    def __init__(self, source: SourceAudit, client: rest.ClientSharePoint | None = None):
        self.source = source
        self.client = client or client_pour(source.lien)
        self.racine = rest.chemin_serveur(source.lien)
        self._dossier = ""
        self._compteur = 0
        self._fichiers: dict[str, list[dict]] = {}
        self._verrou = threading.Lock()

    def __enter__(self) -> "AccesApi":
        self._dossier = tempfile.mkdtemp(prefix="avf_sharepoint_")
        return self

    def __exit__(self, *exc) -> bool:
        shutil.rmtree(self._dossier, ignore_errors=True)
        return False

    def _dossier_commune(self, commune: str) -> str:
        code = commune.strip().upper()
        return f"{self.racine}/Dep{code[:2]}/{code}"

    def _liste(self, commune: str) -> list[dict]:
        dossier = self._dossier_commune(commune)
        with self._verrou:
            if dossier in self._fichiers:
                return self._fichiers[dossier]
        try:
            fichiers = self.client.fichiers(dossier)
        except FileNotFoundError:
            fichiers = []
        with self._verrou:
            self._fichiers[dossier] = fichiers
        return fichiers

    def trouver(self, racine: str, commune: str) -> str | None:
        nom = rest.fichier_audit(self._liste(commune))
        return self.client.url(f"{self._dossier_commune(commune)}/{nom}") if nom else None

    def conflits(self, racine: str, commune: str) -> list[dict]:
        dossier = self._dossier_commune(commune)
        return [{"chemin": self.client.url(f"{dossier}/{nom}"), "date": date}
                for nom, date in rest.copies_de_conflit(self._liste(commune))]

    def existe(self, commune: str) -> bool:
        return self.client.dossier_existe(self._dossier_commune(commune))

    def adresse(self, chemin: str) -> str:
        return chemin

    def lisible(self, chemin: str) -> str:
        if not self._dossier:
            raise RuntimeError("AccesApi s'utilise dans un bloc « with »")
        with self._verrou:
            self._compteur += 1
            nom = chemin.rsplit("/", 1)[-1].replace("%20", " ")
            destination = os.path.join(self._dossier, f"{self._compteur:05d}_{nom}")
        return self.client.telecharger(rest.chemin_serveur(chemin), destination)

    def lister_communes(self) -> list[tuple[str, str, bool]]:
        return rest.lister_communes(self.client, self.racine)


class SessionApi:
    """Mise à jour par l'API SharePoint, avec la session du navigateur (pas d'Excel).

    Même interface que ``sna.services.excel_com.SessionExcel`` : le moteur de
    complete_excel_files (``traiter_communes``) l'utilise sans changement. Pour
    chaque fichier :

    1. téléchargement de la version serveur, avec son ETag ;
    2. application des règles par le moteur (``sna.services.excel.appliquer_regles``,
       l'écriture openpyxl de complete_excel_files : règles, filtres retirés,
       mise en forme reprise, vraies dates) ;
    3. renvoi au serveur **seulement si le fichier est encore à cette version**
       (``If-Match``) ; sinon, quelqu'un l'a enregistré entre-temps : on repart
       de sa version (aucune modification d'autrui n'est écrasée) ;
    4. relecture de la version serveur : les règles n'y ont plus rien à changer,
       sinon les ID sont « non confirmés » (rejoints aux ID en attente).

    Un fichier extrait ou verrouillé est signalé « verrouillé » sans rien écrire.
    """

    TENTATIVES = 3

    def __init__(self, client: rest.ClientSharePoint):
        self.client = client
        self._dossier = ""
        self._compteur = 0
        self._verrou = threading.Lock()

    def __enter__(self) -> "SessionApi":
        self._dossier = tempfile.mkdtemp(prefix="avf_maj_sharepoint_")
        return self

    def __exit__(self, *exc) -> bool:
        shutil.rmtree(self._dossier, ignore_errors=True)
        return False

    def _copie(self, adresse: str) -> str:
        with self._verrou:
            self._compteur += 1
            return os.path.join(self._dossier, f"{self._compteur:05d}_{adresse.rsplit('/', 1)[-1].replace('%20', ' ')}")

    def appliquer_regles(self, fichier: str, config, ids_erreur, on_etape=None):
        from sna.services.excel import appliquer_regles

        def dire(message: str) -> None:
            if on_etape is not None:
                on_etape(message)

        ids = list(ids_erreur)
        chemin = rest.chemin_serveur(fichier)
        etat = self.client.etat_fichier(chemin)
        if etat["CheckOutType"] != 2:                  # 2 = non extrait
            from sna.services.excel import BilanFichier
            bilan = BilanFichier(comptes={r.code: 0 for r in config.actives})
            bilan.verrouille = etat["CheckedOutBy"] or "un autre utilisateur (fichier extrait)"
            dire(f"⚠ fichier extrait par {bilan.verrouille} — rien écrit")
            return bilan
        for essai in range(1, self.TENTATIVES + 1):
            local = self._copie(fichier)
            etag = self.client.telecharger_version(chemin, local)
            dire(f"version serveur téléchargée (API SharePoint, essai {essai})")
            bilan = appliquer_regles(local, config, ids, on_etape=on_etape)
            bilan.ouverture = "API SharePoint (session du navigateur)"
            bilan.tentatives = essai
            if not bilan.enregistre:
                return bilan                           # rien à écrire : valeurs déjà en place
            with open(local, "rb") as f:
                contenu = f.read()
            try:
                dire("envoi sur SharePoint (si personne ne l'a modifié entre-temps)…")
                self.client.remplacer(chemin, contenu, etag)
            except rest.VersionModifiee:
                dire("↻ fichier enregistré par quelqu'un d'autre entre-temps — nouvelle lecture")
                continue
            except rest.FichierVerrouille as exc:
                bilan.verrouille = str(exc)
                bilan.enregistre = False
                bilan.comptes = {k: 0 for k in bilan.comptes}
                dire(f"⚠ {exc} — rien écrit")
                return bilan
            self._verifier(chemin, config, bilan, dire)
            return bilan
        raise rest.VersionModifiee(
            f"fichier enregistré par d'autres à chaque tentative ({self.TENTATIVES}) : rejouez plus tard")

    def _verifier(self, chemin: str, config, bilan, dire) -> None:
        """Relit la version serveur : les règles ne doivent plus rien y changer."""
        from sna.services.excel import appliquer_regles
        relu = self._copie(chemin)
        try:
            self.client.telecharger(chemin, relu)
            reste = appliquer_regles(relu, config, sorted(bilan.ids_trouves), simuler=True)
        except Exception as exc:  # noqa: BLE001 — l'envoi a réussi : on ne met pas la commune en échec
            bilan.ids_non_verifies = sorted(bilan.ids_trouves)
            dire(f"⚠ relecture impossible après l'envoi ({exc}) — ID remis en attente")
            return
        if reste.nb_modifications:
            bilan.ids_non_verifies = sorted(bilan.ids_trouves)
            dire("⚠ valeurs non retrouvées après l'envoi — ID remis en attente")
        else:
            dire("✓ valeurs vérifiées sur SharePoint")


def acces_pour(source: SourceAudit, session_factory=None,
               montages: list[tuple[str, str]] | None = None) -> AccesFichiers:
    """Accès aux fichiers adapté à la source (contexte à ouvrir avec ``with``)."""
    if source.mode == API:
        return AccesApi(source)
    if source.mode == EXCEL:
        return AccesViaExcel(source, session_factory, montages)
    if source.mode == URL:
        return AccesParAdresse(source, session_factory, montages)
    return AccesFichiers()
