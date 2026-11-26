"""Écriture des mises à jour : cellule par cellule, en co-édition, même si le fichier est ouvert.

Le problème : un fichier audit ouvert par un collègue (dans Excel, Excel Online ou
Teams) n'est pas modifiable « en bloc ». SharePoint refuse de remplacer un fichier
en cours de co-édition (verrou partagé), et réécrire tout le classeur ferait perdre
le travail en cours des autres.

La solution est celle d'Office lui-même : la **co-édition**. Excel ouvre le fichier
par son adresse SharePoint et rejoint la session des autres utilisateurs. Seules les
cellules des lignes visées, celles des **ID imm**, sont écrites, puis SharePoint
fusionne avec le travail des autres (``sna.services.excel_com.SessionExcel``).
C'est la seule voie disponible : l'API Excel de Microsoft 365, qui écrit des
cellules sans Excel, exige une application déclarée par l'informatique (« OAuth
only », vérifié sur le SharePoint réel).

:class:`SessionEcriture`, par fichier :

1. **co-édition Excel** quand Excel est disponible, cellule par cellule ;
2. sinon, **repli** :
   - connexion navigateur : envoi par l'API SharePoint, refusé si le fichier est
     ouvert par d'autres (``SessionApi``) ;
   - fichier local : réécriture par openpyxl (moteur complete_excel_files), après
     avoir vérifié qu'il n'est pas ouvert. Excel laisse un fichier verrou ``~$…``
     qui donne le nom de la personne.

Chaque refus est expliqué : qui a le fichier, pourquoi, et comment débloquer.
"""

from __future__ import annotations

from collections.abc import Callable
import os

from sna.services.excel import BilanFichier, appliquer_regles as appliquer_openpyxl


class SessionEcriture:
    """Même interface que ``SessionExcel`` (``appliquer_regles``) : utilisable par le moteur.

    Args:
        excel_factory: fabrique d'une session Excel (co-édition) ; ``None`` = Excel absent.
        repli: session de repli (``SessionApi``) ; ``None`` = réécriture locale openpyxl.
        prefere_excel: Excel pour **chaque** fichier. Sinon, Excel n'est démarré qu'à la
            demande, pour un fichier que le repli ne peut pas écrire parce qu'il est ouvert
            (verrou ``~$``, refus 423 de SharePoint) : aucun fichier ouvert n'est laissé de
            côté tant qu'Excel peut le co-éditer.
        raison_sans_excel: pourquoi Excel n'est pas préféré (affiché si le repli est utilisé).
    """

    def __init__(self, excel_factory: Callable[[], object] | None = None, repli=None,
                 raison_sans_excel: str = "", prefere_excel: bool | None = None):
        self._excel_factory = excel_factory
        self.repli = repli
        self.excel = None
        self.prefere_excel = (excel_factory is not None) if prefere_excel is None else prefere_excel
        self._excel_essaye = False
        self.raison_sans_excel = raison_sans_excel or (
            "Excel / pywin32 introuvables sur ce poste" if excel_factory is None else "")
        self.excel_utilise = False
        self.repli_utilise = False

    def __enter__(self) -> "SessionEcriture":
        if self.repli is not None and hasattr(self.repli, "__enter__"):
            self.repli.__enter__()
        if self.prefere_excel:
            self._excel_pret()
        return self

    def __exit__(self, *exc) -> bool:
        try:
            if self.excel is not None:
                self.excel.__exit__(*exc)
        finally:
            if self.repli is not None and hasattr(self.repli, "__exit__"):
                self.repli.__exit__(*exc)
        return False

    def _excel_pret(self) -> bool:
        """Démarre Excel au premier besoin ; un seul essai par session."""
        if self.excel is not None:
            return True
        if self._excel_factory is None or self._excel_essaye:
            return False
        self._excel_essaye = True
        try:
            session = self._excel_factory()
            session.__enter__()
            self.excel = session
            return True
        except Exception as exc:  # noqa: BLE001 — Excel absent, bloqué… : repli
            self.raison_sans_excel = f"Excel n'a pas pu démarrer : {_raison(exc)}"
            return False

    def _par_excel(self, fichier, config, ids, on_etape, dire) -> BilanFichier | None:
        """Écriture en co-édition ; ``None`` si Excel est hors service (repli)."""
        try:
            bilan = self.excel.appliquer_regles(fichier, config, ids, on_etape=on_etape)
        except Exception as exc:  # noqa: BLE001
            if not _excel_hors_service(exc):
                raise                       # erreur propre à ce fichier : la commune échoue, expliquée
            self.raison_sans_excel = f"co-édition Excel impossible : {_raison(exc)}"
            dire(f"⚠ {self.raison_sans_excel} — écriture de repli")
            try:
                self.excel.__exit__(None, None, None)
            except Exception:  # noqa: BLE001
                pass
            self.excel = None               # ne pas réessayer Excel pour chaque fichier
            return None
        self.excel_utilise = True
        if bilan.verrouille and "co-édition" not in bilan.verrouille:
            bilan.verrouille = (f"{bilan.verrouille} — classeur ouvert en mode exclusif (pas en "
                                "co-édition), Excel ne l'obtient qu'en lecture seule : il doit être fermé, "
                                "ou ouvert dans Excel Online / Excel Microsoft 365 avec l'enregistrement "
                                "automatique (co-édition)")
        return bilan

    def appliquer_regles(self, fichier: str, config, ids_erreur, on_etape=None) -> BilanFichier:
        def dire(message: str) -> None:
            if on_etape is not None:
                on_etape(message)

        ids = list(ids_erreur)
        if self.prefere_excel and self._excel_pret():
            bilan = self._par_excel(fichier, config, ids, on_etape, dire)
            if bilan is not None:
                return bilan

        self.repli_utilise = True
        if self.repli is not None:
            bilan = self.repli.appliquer_regles(fichier, config, ids, on_etape=on_etape)
            qui = bilan.verrouille if bilan.verrouille and "EXTRAIT" not in bilan.verrouille else None
        else:
            qui = proprietaire_verrou(fichier)
            bilan = None
        if qui and not self.prefere_excel and self._excel_pret():
            dire(f"fichier ouvert par {qui} : écriture cellule par cellule via Excel (co-édition)")
            par_excel = self._par_excel(fichier, config, ids, on_etape, dire)
            if par_excel is not None:
                return par_excel
        if bilan is None:
            if qui:
                bilan = BilanFichier(comptes={r.code: 0 for r in config.actives})
                bilan.verrouille = qui
                dire(f"⚠ fichier ouvert par {qui} — rien écrit")
            else:
                bilan = appliquer_openpyxl(fichier, config, ids, on_etape=on_etape)
        if bilan.verrouille and "EXTRAIT" not in bilan.verrouille:
            bilan.verrouille = (f"{bilan.verrouille} (fichier ouvert) — l'écriture dans un fichier ouvert "
                                f"passe par la co-édition Excel, indisponible ({self.raison_sans_excel})")
        return bilan


# ── Fichier ouvert sur ce poste ou un partage : le fichier verrou d'Excel ─
def proprietaire_verrou(chemin: str) -> str | None:
    """Nom de la personne qui a ``chemin`` ouvert dans Excel, ou ``None`` s'il est libre.

    Excel crée ``~$<nom du fichier>`` à côté du classeur ouvert ; il contient le nom de
    l'utilisateur. À défaut de ce fichier, un fichier qu'on ne peut pas ouvrir en écriture
    est aussi considéré comme ouvert.
    """
    if not chemin or chemin.lower().startswith(("http://", "https://")) or not os.path.isfile(chemin):
        return None
    dossier, nom = os.path.split(chemin)
    for candidat in (f"~${nom}", f"~${nom[2:]}"):
        verrou = os.path.join(dossier, candidat)
        if os.path.isfile(verrou):
            return _nom_dans_verrou(verrou) or "un autre utilisateur"
    try:
        with open(chemin, "r+b"):
            return None
    except PermissionError:
        return "un autre utilisateur (fichier ouvert ou verrouillé)"
    except OSError:
        return None


def _nom_dans_verrou(chemin: str) -> str:
    """Le fichier verrou commence par la longueur du nom puis le nom (ANSI)."""
    try:
        with open(chemin, "rb") as f:
            donnees = f.read(165)
    except OSError:
        return ""
    if not donnees:
        return ""
    longueur = donnees[0]
    nom = donnees[1:1 + longueur].decode("cp1252", errors="ignore").strip("\x00 ").strip()
    return nom if nom.isprintable() else ""


def _excel_hors_service(exc: BaseException) -> bool:
    """Excel lui-même est inutilisable (absent, connexion refusée…), pas seulement ce fichier."""
    try:
        from sna.services.excel_com import ConnexionRequise, ExcelIndisponible
    except ImportError:
        return False
    return isinstance(exc, (ConnexionRequise, ExcelIndisponible))


def _raison(exc: BaseException) -> str:
    return (str(exc) or exc.__class__.__name__).split("\n")[0][:200]
