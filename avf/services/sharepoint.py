"""Audit depuis SharePoint : la version **serveur** des fichiers, pas la copie OneDrive.

Pourquoi : une modification faite dans le dossier OneDrive local peut ne jamais
atteindre SharePoint (conflit de synchronisation, copie « -NOMDUPC », envoi en
échec). Auditer le dossier local montrerait alors un état que personne d'autre
ne voit. Ici, l'utilisateur colle le lien SharePoint du dossier racine et
l'audit lit ce que contient réellement le serveur.

Deux accès, sans inscription d'application ni droits d'administrateur :

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
from urllib.parse import parse_qs, quote, unquote, urlsplit

from sna.services.onedrive import montages_onedrive, url_sharepoint

from ..models.source import SourceAudit

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


# ── Connexion ─────────────────────────────────────────────────────
def connecter(
    texte: str,
    delai_s: float = 45,
    montages: list[tuple[str, str]] | None = None,
    est_dossier: Callable[[str], bool] = os.path.isdir,
) -> SourceAudit:
    """Trouve un accès au dossier SharePoint : WebDAV, sinon bibliothèque synchronisée.

    Raises:
        LienInvalide: lien inexploitable.
        AccesImpossible: aucun accès possible (avec la marche à suivre).
    """
    lien = analyser_lien(texte)
    webdav = lien.chemin_webdav
    ok, erreur = _avec_delai(lambda: est_dossier(webdav), delai_s)
    if ok:
        return SourceAudit("webdav", webdav, lien.url)
    local = dossier_synchronise(lien, montages)
    if local and est_dossier(local):
        return SourceAudit("excel", local, lien.url)
    cause = "délai dépassé" if erreur == "delai" else (erreur or "accès refusé ou dossier introuvable")
    raise AccesImpossible(
        f"dossier SharePoint injoignable ({cause}).\n\n"
        f"Accès direct (WebDAV) : {webdav}\n"
        "• Vérifiez le lien (il doit ouvrir le dossier racine dans le navigateur).\n"
        "• Connectez-vous à SharePoint dans Microsoft Edge en cochant « Rester connecté », "
        "puis réessayez — ou ouvrez une fois le dossier via « Afficher dans l'Explorateur "
        "de fichiers ».\n"
        "• Le service Windows « WebClient » doit être démarré.\n"
        "• À défaut, synchronisez la bibliothèque avec OneDrive : l'audit lira alors chaque "
        "fichier sur le serveur via Excel."
    )


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
            from sna.services.excel_com import SessionExcel
            session_factory = SessionExcel
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
        url = url_sharepoint(chemin, self._montages)
        if url is None:
            raise AccesImpossible(f"adresse SharePoint inconnue pour {chemin} (dossier non synchronisé ?)")
        with self._verrou:
            self._compteur += 1
            destination = os.path.join(self._dossier, f"{self._compteur:05d}_{os.path.basename(chemin)}")
        self._session.copier_version_serveur(url, destination)
        return destination


def acces_pour(source: SourceAudit, session_factory=None,
               montages: list[tuple[str, str]] | None = None) -> AccesFichiers:
    """Accès aux fichiers adapté à la source (contexte à ouvrir avec ``with``)."""
    if source.mode == "excel":
        return AccesViaExcel(source, session_factory, montages)
    return AccesFichiers()
