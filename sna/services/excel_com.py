"""Écriture des règles via Excel (COM), dans la session de co-édition SharePoint.

Pourquoi ce module existe : ``services.excel`` modifie le fichier en réécrivant
le classeur entier avec openpyxl, puis laisse OneDrive renvoyer ce nouveau
fichier au serveur. Si un autre utilisateur a le classeur ouvert — ou si
OneDrive est en train de synchroniser — OneDrive ne sait pas fusionner une
version produite hors d'Office : conflit de fusion, copie « -NOMDUPC », ou
simplement rien envoyé au serveur (« Certaines de vos modifications n'ont
peut-être pas pu être affichées »). La mise à jour a l'air d'avoir réussi en
local, mais les autres utilisateurs ne la voient jamais sur SharePoint.

Ici, Excel ouvre le classeur (par son adresse SharePoint si le dossier est
synchronisé par OneDrive, sinon par son chemin local), rejoint la co-édition
comme n'importe quel utilisateur, n'écrit que les cellules concernées par les
règles, puis enregistre : c'est SharePoint qui fusionne avec le travail des
autres, pas notre programme. Chaque valeur écrite est **relue après
l'enregistrement** pour confirmer qu'elle est bien en place ; un ID dont
l'écriture n'est pas confirmée n'est jamais compté comme traité.

Si le classeur est ouvert en mode exclusif ailleurs, Excel l'ouvre en lecture
seule : rien n'est écrit, et c'est signalé (verrouillé) plutôt qu'un faux
succès.

Garde-fou anti-blocage
----------------------
Un classeur partagé en co-édition peut faire apparaître une boîte de dialogue
SharePoint (« Nous rechargeons votre classeur pour terminer la fusion ») —
modale, donc jamais validée puisqu'Excel tourne invisible : l'appel Excel en
cours (souvent ``Save()``) ne revient jamais. Pour ne jamais geler l'outil,
chaque appel tourne sur un fil dédié, propriétaire de l'appartement COM pour
toute la session (:class:`_FilExcel`). Ce fil donne des **signes de vie** à
chaque étape (attente d'Excel, chaque ligne écrite ou relue) ; s'il n'en donne
plus pendant ``delai_blocage_s`` (``delai_ouverture_s`` pour l'ouverture du
classeur), le PROCESSUS Excel de cette session est tué (jamais celui de
l'utilisateur), un neuf est relancé et le classeur est retenté. Le délai porte
sur l'**inactivité**, pas sur la durée totale : une grosse commune qui
travaille normalement n'est jamais prise pour un blocage.

``AutoSaveOn`` et ``Interactive`` sont aussi désactivés pour réduire la
fréquence de ces fusions, sans l'éliminer : c'est le garde-fou ci-dessus qui
garantit l'absence de blocage, pas ces réglages.

Vérification après enregistrement
---------------------------------
Elle ne repose que sur la **relecture des cellules** (le drapeau global
``classeur.Saved`` est remis à faux par la moindre activité de co-édition sans
rapport avec nos cellules : s'y fier donnait de faux « écriture non
confirmée »). La relecture repart d'objets COM neufs et retrouve chaque ID par
sa valeur, pas par son ancien numéro de ligne : une fusion qui recharge le
classeur (objets invalidés, erreur 0x800A01A8) ou un collègue qui insère des
lignes ne faussent donc pas le résultat. Si la relecture est impossible, les
ID concernés sont « non confirmés » (rejoués plus tard) — la commune n'est pas
mise en échec pour autant, puisque l'enregistrement a réussi.

Concurrence
-----------
:meth:`SessionExcel.appliquer_regles` peut être appelée depuis plusieurs
threads (cf. ``services.traitement``, option ``paralleles``) : les appels sont
sérialisés par un verrou, et le délai anti-blocage ne démarre qu'au moment où
la tâche s'exécute réellement sur le fil Excel — l'attente de son tour n'est
jamais prise pour un blocage.

Ne dépend que de ``pywin32`` (absent de toute autre partie du programme) :
sans lui, ou sans Excel installé, :func:`excel_disponible` renvoie faux et
l'appelant retombe sur ``services.excel`` (réécriture directe du fichier).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
import csv
from datetime import date, datetime, timedelta
import math
import queue
import subprocess
import threading
import time
import traceback

from ..models.regles import ReglesConfig
from .colonnes_excel import NOM_ONGLET_CIBLE, OngletIntrouvable, normaliser, resoudre_colonne, table_depuis_entetes
from .evaluation import regle_sapplique, valeur_pour_action
from .excel import BilanFichier, ColonneIntrouvable, _plan_regle
from .identifiants import normaliser_id
from .onedrive import url_sharepoint
from .valeurs import est_format_date

_EN = 1033          # LCID anglais : NumberFormat est refusé par un Excel en français sinon.
_EPOQUE = datetime(1899, 12, 30)
_XL_CALC_MANUEL = -4135

# HRESULT renvoyés quand Excel est occupé (chargement, synchro co-édition, recalcul, enregistrement).
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
_REJETS = (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER)

# Processus Excel mort ou proxy coupé de son serveur : il faut une instance neuve.
RPC_E_DISCONNECTED = -2147417848         # 0x80010108 « déconnecté de ses clients »
RPC_S_SERVER_UNAVAILABLE = -2147023174   # 0x800706BA
RPC_S_CALL_FAILED = -2147023170          # 0x800706BE
CO_E_OBJNOTCONNECTED = -2147220995       # 0x800401FD
_DECONNEXIONS = (RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED, CO_E_OBJNOTCONNECTED)

# Objet COM invalidé par un rechargement du classeur (fusion de co-édition) :
# l'instance est saine, il suffit de rouvrir le classeur.
XL_OBJET_INVALIDE = -2146827864          # 0x800A01A8 « Objet requis »

# Signature du rapporteur d'étapes (journal en temps réel).
Etape = Callable[[str], None]


class ExcelIndisponible(Exception):
    """pywin32 absent, ou Excel non installé / impossible à démarrer."""


class ExcelOccupe(Exception):
    """Excel a refusé les appels trop longtemps (occupé par une synchronisation, un calcul…)."""


class ExcelBloque(Exception):
    """Excel n'a plus donné signe de vie dans le délai imparti : probablement une
    boîte de dialogue de fusion SharePoint restée sans réponse (fenêtre « Ne
    répond pas »). Le processus de cette session est tué et un neuf est relancé."""


class ConnexionRequise(ExcelBloque):
    """Excel affiche la fenêtre de connexion Microsoft 365 (compte professionnel).

    Dans une instance invisible, personne ne peut y répondre : sans cette
    détection, l'ouverture d'un fichier SharePoint resterait bloquée jusqu'au
    délai anti-blocage, à chaque tentative. Ajout de la copie vendorisée.
    """


# Classes de fenêtre de la connexion Microsoft 365 ouverte par Office (WebView de
# l'authentification « moderne »). Leur apparition pendant un appel = connexion demandée.
_CLASSES_CONNEXION = {"BasicEmbeddedBrowser"}


def fenetre_connexion(pid: int | None) -> bool:
    """Vrai si le processus ``pid`` affiche une fenêtre de connexion Microsoft 365."""
    if not pid:
        return False
    try:
        import win32gui
        import win32process
    except ImportError:
        return False
    trouvee = []

    def examiner(fenetre, _):
        try:
            if (win32process.GetWindowThreadProcessId(fenetre)[1] == pid
                    and win32gui.GetClassName(fenetre) in _CLASSES_CONNEXION):
                trouvee.append(fenetre)
                return False
        except Exception:  # noqa: BLE001
            pass
        return True

    try:
        win32gui.EnumWindows(examiner, None)
    except Exception:  # noqa: BLE001 — EnumWindows « échoue » quand le rappel l'arrête
        pass
    return bool(trouvee)


def excel_disponible() -> bool:
    """Vrai si pywin32 est installé et qu'Excel est enregistré sur le poste."""
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer"))
        return True
    except Exception:  # noqa: BLE001 — n'importe quel souci → repli sur l'écriture directe
        return False


def _pythoncom():
    import pythoncom
    return pythoncom


# ── Classement des erreurs COM ─────────────────────────────────────
def _codes_com(exc) -> set[int]:
    """HRESULT d'une ``com_error`` : code de l'appel et, s'il y en a, code renvoyé par Excel."""
    args = getattr(exc, "args", None) or ()
    codes: set[int] = set()
    if args and isinstance(args[0], int):
        codes.add(args[0])
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 5 and isinstance(args[2][5], int):
        codes.add(args[2][5])
    return codes


def _est_rejet(exc) -> bool:
    return bool(getattr(exc, "args", None)) and exc.args[0] in _REJETS


def _est_deconnexion(exc) -> bool:
    return bool(_codes_com(exc) & set(_DECONNEXIONS))


def _est_objet_invalide(exc) -> bool:
    return XL_OBJET_INVALIDE in _codes_com(exc)


# ── Appel COM forcé en anglais (sinon un Excel en français refuse NumberFormat) ──
def _put_en(obj, propriete: str, valeur) -> None:
    dispid = obj._oleobj_.GetIDsOfNames(propriete)
    obj._oleobj_.Invoke(dispid, _EN, _pythoncom().DISPATCH_PROPERTYPUT, 0, valeur)


def _get_en(obj, propriete: str):
    dispid = obj._oleobj_.GetIDsOfNames(propriete)
    return obj._oleobj_.Invoke(dispid, _EN, _pythoncom().DISPATCH_PROPERTYGET, 1)


# ── Conversions de valeurs ─────────────────────────────────────────
def _serie(valeur) -> float | None:
    """``datetime``/``date`` → numéro de série Excel (jours depuis 30/12/1899)."""
    if isinstance(valeur, datetime):
        return (valeur.replace(tzinfo=None) - _EPOQUE).total_seconds() / 86400
    if isinstance(valeur, date):
        return float((valeur - _EPOQUE.date()).days)
    return None


def _plat(valeurs):
    """``Range.Value`` d'une ligne ou d'une colonne → liste plate (scalaire si une seule cellule)."""
    if not isinstance(valeurs, tuple):
        return [valeurs]
    return [x for rangee in valeurs for x in rangee]


# ── Processus Excel ────────────────────────────────────────────────
def _pid_de(xl):
    """PID de l'instance Excel pilotée (via sa fenêtre principale)."""
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(xl.Hwnd)[1] or None
    except Exception:  # noqa: BLE001
        return None


def _pids_excel() -> set[int]:
    """PID de tous les processus EXCEL.EXE du poste (repli si la fenêtre ne donne pas le PID)."""
    try:
        sortie = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
    except Exception:  # noqa: BLE001
        return set()
    return {int(champs[1]) for champs in csv.reader(sortie.splitlines())
            if len(champs) > 1 and champs[1].strip().isdigit()}


def _terminer_si_toujours_ouvert(pid: int | None, attente_s: float = 15) -> None:
    """Ferme le PROCESSUS de CETTE session s'il ne s'est pas terminé seul (jamais celui de l'utilisateur)."""
    if not pid:
        return
    try:
        import win32api
        import win32con
        import win32event
        poignee = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:  # noqa: BLE001
        return
    try:
        if win32event.WaitForSingleObject(poignee, int(attente_s * 1000)) == win32event.WAIT_TIMEOUT:
            win32api.TerminateProcess(poignee, 1)
            win32event.WaitForSingleObject(poignee, 5000)
    finally:
        win32api.CloseHandle(poignee)


def _terminer_de_force(pid: int | None, attente_s: float = 5) -> None:
    """Termine le PROCESSUS de cette session (jamais celui de l'utilisateur),
    sans délai de grâce — à utiliser seulement après un blocage confirmé
    (``ExcelBloque``). Attend ensuite que le processus ait *réellement*
    disparu : rouvrir aussitôt le même fichier, avant que Windows n'ait fini
    de relâcher ses verrous, peut renvoyer un objet COM à moitié initialisé."""
    if not pid:
        return
    try:
        import win32api
        import win32con
        import win32event
        poignee = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:  # noqa: BLE001
        return
    try:
        win32api.TerminateProcess(poignee, 1)
        win32event.WaitForSingleObject(poignee, int(attente_s * 1000))
    except Exception:  # noqa: BLE001
        pass
    finally:
        win32api.CloseHandle(poignee)


class _FiltreMessages:
    """Filtre de messages COM (``IMessageFilter``) pour le thread courant.

    Sans lui, tout appel envoyé pendant qu'Excel est occupé échoue aussitôt
    avec ``RPC_E_CALL_REJECTED``. Avec lui, Windows réessaie l'appel toutes
    les 250 ms jusqu'à ``delai_max_ms``. pywin32 n'exposant pas
    ``CoRegisterMessageFilter``, l'interface est construite avec ``ctypes``
    (table virtuelle de 6 méthodes).
    """

    _ATTENTE_MS = 250

    def __init__(self, delai_max_ms: int = 120_000):
        import ctypes
        import uuid
        from ctypes import wintypes

        self.delai_max_ms = delai_max_ms
        self._ancien = ctypes.c_void_p()
        self._installe = False
        HRESULT, DWORD, ULONG, LPVOID = ctypes.c_long, wintypes.DWORD, ctypes.c_ulong, ctypes.c_void_p

        class GUID(ctypes.Structure):
            _fields_ = [("d", ctypes.c_ubyte * 16)]

        def guid(s):
            return GUID.from_buffer_copy(uuid.UUID(s).bytes_le)

        connus = {
            bytes(guid("00000000-0000-0000-C000-000000000046").d),   # IUnknown
            bytes(guid("00000016-0000-0000-C000-000000000046").d),   # IMessageFilter
        }
        F = ctypes.WINFUNCTYPE

        def query_interface(this, riid, ppv):
            if bytes(riid.contents.d) in connus:
                ppv[0] = this
                return 0
            ppv[0] = None
            return -2147467262                                      # E_NOINTERFACE

        def handle_incoming(this, call_type, task, tick, info):
            return 0                                                 # SERVERCALL_ISHANDLED

        def retry_rejected(this, task, tick, reject_type):
            # tick = ms écoulées depuis le 1er refus de CET appel ; 0xFFFFFFFF = abandon.
            return self._ATTENTE_MS if tick < self.delai_max_ms else 0xFFFFFFFF

        def message_pending(this, task, tick, pending_type):
            return 2                                                 # PENDINGMSG_WAITDEFPROCESS

        self._fns = [
            F(HRESULT, LPVOID, ctypes.POINTER(GUID), ctypes.POINTER(LPVOID))(query_interface),
            F(ULONG, LPVOID)(lambda this: 1),                       # AddRef (objet statique)
            F(ULONG, LPVOID)(lambda this: 1),                       # Release
            F(DWORD, LPVOID, DWORD, LPVOID, DWORD, LPVOID)(handle_incoming),
            F(DWORD, LPVOID, LPVOID, DWORD, DWORD)(retry_rejected),
            F(DWORD, LPVOID, LPVOID, DWORD, DWORD)(message_pending),
        ]

        class VTable(ctypes.Structure):
            _fields_ = [(f"f{i}", type(f)) for i, f in enumerate(self._fns)]

        class Objet(ctypes.Structure):
            _fields_ = [("vtbl", ctypes.POINTER(VTable))]

        self._vtbl = VTable(*self._fns)
        self._objet = Objet(ctypes.pointer(self._vtbl))
        self._ole32 = ctypes.windll.ole32
        self._ole32.CoRegisterMessageFilter.argtypes = [LPVOID, ctypes.POINTER(LPVOID)]

    def installer(self) -> bool:
        import ctypes
        hr = self._ole32.CoRegisterMessageFilter(ctypes.addressof(self._objet), ctypes.byref(self._ancien))
        self._installe = hr == 0
        return self._installe

    def retirer(self) -> None:
        if self._installe:
            self._ole32.CoRegisterMessageFilter(self._ancien, None)
            self._installe = False


# ── Résolution de l'onglet « Audit » (COM) ─────────────────────────
def _choisir_feuille_com(classeur, nom: str = NOM_ONGLET_CIBLE):
    """Équivalent COM de ``colonnes_excel.choisir_feuille`` : même nom, même tolérance."""
    cible = normaliser(nom)
    noms = [str(classeur.Worksheets(i).Name) for i in range(1, classeur.Worksheets.Count + 1)]
    for index, nom_reel in enumerate(noms, start=1):
        if normaliser(nom_reel) == cible:
            return classeur.Worksheets(index)
    raise OngletIntrouvable(
        f"onglet « {nom} » introuvable — onglets disponibles : " + ", ".join(noms)
    )


def _retirer_filtres_com(feuille) -> dict:
    """Retire le filtre automatique et réaffiche lignes/colonnes masquées (COM).

    Même effet que ``services.excel._afficher_tout`` (retrait complet, flèches
    comprises) pour un filtre de feuille normal. Un tableau Excel
    (``ListObject``) garde ses flèches — elles lui sont intrinsèques — mais
    perd son critère de filtre actif. Sur une feuille protégée, un élément qui
    ne peut pas être démasqué est listé dans ``"protege"`` plutôt que de lever
    une erreur — la protection n'est jamais retirée.

    Returns:
        ``{"filtres": bool, "lignes": bool, "colonnes": [...], "protege": [...]}``
    """
    resultat = {"filtres": False, "lignes": False, "colonnes": [], "protege": []}
    protegee = bool(feuille.ProtectContents)

    try:
        if feuille.AutoFilterMode:
            feuille.AutoFilterMode = False     # retire le filtre entier, flèches comprises
            resultat["filtres"] = True
    except Exception:  # noqa: BLE001 — pas bloquant pour la mise à jour
        pass

    for i in range(1, feuille.ListObjects.Count + 1):
        try:
            filtre_tableau = feuille.ListObjects(i).AutoFilter
            if filtre_tableau is not None and filtre_tableau.FilterMode:
                filtre_tableau.ShowAllData()
                resultat["filtres"] = True
        except Exception:  # noqa: BLE001
            pass

    utile = feuille.UsedRange
    for c in range(utile.Column, utile.Column + utile.Columns.Count):
        if not feuille.Columns(c).Hidden:
            continue
        nom = str(feuille.Cells(1, c).Value or c)
        if protegee:
            resultat["protege"].append(f"colonne {nom}")
            continue
        feuille.Columns(c).Hidden = False
        resultat["colonnes"].append(nom)

    fonctions = feuille.Application.WorksheetFunction
    if fonctions.Subtotal(103, utile) < fonctions.CountA(utile):   # 103 = cellules remplies VISIBLES
        if protegee:
            resultat["protege"].append("lignes masquées")
        else:
            utile.EntireRow.Hidden = False
            resultat["lignes"] = True

    return resultat


def affichage_modifie(resultat: dict | None) -> bool:
    return bool(resultat and (resultat["filtres"] or resultat["lignes"] or resultat["colonnes"]))


# ── Lecture / écriture d'une cellule (COM) ─────────────────────────
def _valeur_cellule_com(feuille, ligne: int, colonne: int):
    """Valeur native d'une cellule — une date reste une date, pas un numéro de série.

    ``Value2`` ne convertit jamais en ``datetime`` (c'est sa raison d'être :
    éviter le décalage UTC que ``Value`` introduit) : une cellule au format
    date renvoie un simple nombre à virgule. On le reconvertit nous-mêmes dès
    que le format d'affichage de la cellule est une date.
    """
    cellule = feuille.Cells(ligne, colonne)
    valeur = cellule.Value2
    if isinstance(valeur, (int, float)) and est_format_date(_get_en(cellule, "NumberFormat")):
        return _EPOQUE + timedelta(days=float(valeur))
    return valeur


def _ecrire_cellule_com(cellule, valeur) -> None:
    if valeur is None or valeur == "":
        cellule.ClearContents()
        return
    serie = _serie(valeur)
    if serie is not None:
        cellule.Value2 = serie
        if not est_format_date(_get_en(cellule, "NumberFormat")):
            _put_en(cellule, "NumberFormat", "dd/mm/yyyy")
        return
    cellule.Value2 = valeur
    if isinstance(valeur, str) and not isinstance(cellule.Value2, str):
        cellule.Value2 = "'" + valeur     # Excel a converti le texte en nombre : forcer le texte


def _conforme_valeur(lu, attendu) -> bool:
    """Vrai si ``lu`` (``Value2`` brut d'une cellule) correspond à la valeur ``attendu``."""
    if attendu is None or attendu == "":
        return lu in (None, "")
    serie = _serie(attendu)
    if serie is not None:
        return isinstance(lu, (int, float)) and not isinstance(lu, bool) and abs(float(lu) - serie) < 1e-6
    if isinstance(attendu, bool):
        return lu == attendu
    if isinstance(attendu, (int, float)):
        return isinstance(lu, (int, float)) and math.isclose(float(lu), float(attendu), rel_tol=1e-9, abs_tol=1e-9)
    return str(lu if lu is not None else "").replace("\r\n", "\n") == str(attendu).replace("\r\n", "\n")


def _conforme_com(cellule, attendu) -> bool:
    return _conforme_valeur(cellule.Value2, attendu)


def _brut_apres_ecriture(valeur):
    """Ce que ``Value2`` renverra après :func:`_ecrire_cellule_com` (pour le cache)."""
    if valeur is None or valeur == "":
        return None
    serie = _serie(valeur)
    return serie if serie is not None else valeur


class _CacheColonnes:
    """Colonnes utiles d'une plage de lignes, lues en **un appel COM par colonne**.

    Remplace des milliers d'appels unitaires (``Cells(l, c).Value2`` puis
    ``NumberFormat`` pour chaque référence de chaque règle de chaque ligne) :
    c'est l'essentiel du temps passé quand Excel est piloté hors processus.
    Le format d'affichage n'est demandé que pour les valeurs numériques (dates
    possibles), en un appel pour toute la colonne quand il est uniforme.
    Les écritures faites par les règles précédentes y sont reportées : une
    condition qui lit une colonne déjà modifiée voit la nouvelle valeur, comme
    avec la lecture cellule par cellule.
    """

    def __init__(self, feuille, lignes: Iterable[int], colonnes: Iterable[int]):
        lignes = sorted(set(lignes))
        self._feuille = feuille
        self._valeurs: dict[tuple[int, int], object] = {}
        self._formats: dict[tuple[int, int], str | None] = {}
        self._format_colonne: dict[int, str | None] = {}
        if not lignes:
            return
        haut, bas = lignes[0], lignes[-1]
        voulues = set(lignes)
        for colonne in sorted(set(colonnes)):
            plage = feuille.Range(feuille.Cells(haut, colonne), feuille.Cells(bas, colonne))
            for decalage, valeur in enumerate(_plat(plage.Value2)):
                if haut + decalage in voulues:
                    self._valeurs[(haut + decalage, colonne)] = valeur
            try:
                fmt = _get_en(plage, "NumberFormat")
                self._format_colonne[colonne] = fmt if isinstance(fmt, str) else None
            except Exception:  # noqa: BLE001 — formats mélangés : demandés cellule par cellule
                self._format_colonne[colonne] = None

    def brut(self, ligne: int, colonne: int):
        if (ligne, colonne) not in self._valeurs:
            self._valeurs[(ligne, colonne)] = self._feuille.Cells(ligne, colonne).Value2
        return self._valeurs[(ligne, colonne)]

    def _format(self, ligne: int, colonne: int) -> str | None:
        if (ligne, colonne) in self._formats:
            return self._formats[(ligne, colonne)]
        fmt = self._format_colonne.get(colonne)
        if fmt is None:
            fmt = _get_en(self._feuille.Cells(ligne, colonne), "NumberFormat")
        self._formats[(ligne, colonne)] = fmt
        return fmt

    def valeur(self, ligne: int, colonne: int):
        """Valeur native (date reconvertie), comme :func:`_valeur_cellule_com`."""
        brut = self.brut(ligne, colonne)
        if isinstance(brut, (int, float)) and not isinstance(brut, bool) and est_format_date(self._format(ligne, colonne)):
            return _EPOQUE + timedelta(days=float(brut))
        return brut

    def lecteur(self, ligne: int, refs: dict[str, int]):
        def lire(reference: str):
            index = refs.get(reference)
            return None if index is None else self.valeur(ligne, index)
        return lire

    def ecrire(self, ligne: int, colonne: int, valeur) -> None:
        self._valeurs[(ligne, colonne)] = _brut_apres_ecriture(valeur)
        if _serie(valeur) is not None and not est_format_date(self._format(ligne, colonne)):
            self._formats[(ligne, colonne)] = "dd/mm/yyyy"


# ── Fil dédié, propriétaire de l'appartement COM ───────────────────
class _FilExcel:
    """Un fil unique qui possède Excel pour toute la session.

    Un objet COM (``Application``, ``Workbook``...) est lié au thread qui l'a
    créé : on ne peut pas répartir les appels sur n'importe quel fil. Celui-ci
    démarre Excel sur son propre fil, puis exécute une à une les tâches qu'on
    lui soumet via :meth:`executer`. La tâche en cours repousse son échéance à
    chaque signe de vie (:meth:`signe`, :meth:`accorder`) ; l'appelant abandonne
    la tâche dès que l'échéance est dépassée, sans jamais rester bloqué.
    """

    def __init__(self, visible: bool, delai_occupe_s: float, delai_demarrage_s: float = 60):
        self.xl = None
        self.pid: int | None = None
        self._filtre: _FiltreMessages | None = None
        self._taches: queue.Queue = queue.Queue()
        self._delai_courant = 0.0
        self._echeance: float | None = None
        pret = threading.Event()
        self._erreur_demarrage: BaseException | None = None
        self._fil = threading.Thread(
            target=self._boucle, args=(visible, delai_occupe_s, pret), daemon=True, name="sna-excel",
        )
        self._fil.start()
        if not pret.wait(delai_demarrage_s):
            # S'il finit par démarrer, le fil trouvera cette pilule et refermera
            # Excel aussitôt, au lieu de laisser une instance invisible orpheline.
            self._taches.put(None)
            raise ExcelIndisponible(f"Excel n'a pas démarré en {delai_demarrage_s:.0f} s")
        if self._erreur_demarrage is not None:
            raise ExcelIndisponible(
                f"impossible de démarrer Excel : {self._erreur_demarrage}"
            ) from self._erreur_demarrage

    def _boucle(self, visible: bool, delai_occupe_s: float, pret: threading.Event) -> None:
        pythoncom = _pythoncom()
        pythoncom.CoInitialize()
        try:
            import win32com.client as wc
            self._filtre = _FiltreMessages(delai_max_ms=int(delai_occupe_s * 1000))
            self._filtre.installer()
            avant = _pids_excel()
            self.xl = wc.DispatchEx("Excel.Application")   # nouvelle instance, n'interfère pas avec l'Excel ouvert
            self.pid = _pid_de(self.xl)
            if self.pid is None:
                nouveaux = _pids_excel() - avant
                self.pid = nouveaux.pop() if len(nouveaux) == 1 else None
            self.xl.Visible = bool(visible)
            self.xl.DisplayAlerts = False
            self.xl.AskToUpdateLinks = False
            self.xl.ScreenUpdating = bool(visible)
            self.xl.Interactive = False   # réduit les boîtes de dialogue nécessitant une réponse
        except Exception as exc:  # noqa: BLE001 — transmis au fil appelant
            self._erreur_demarrage = exc
            pret.set()
            pythoncom.CoUninitialize()
            return
        pret.set()

        while True:
            tache = self._taches.get()
            if tache is None:   # pilule empoisonnée : fin de session demandée
                break
            tache()

        try:
            for i in range(self.xl.Workbooks.Count, 0, -1):
                self.xl.Workbooks(i).Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            self.xl.Quit()
        except Exception:  # noqa: BLE001
            pass
        self.xl = None
        if self._filtre is not None:
            self._filtre.retirer()
        pythoncom.CoUninitialize()

    # ── Signes de vie (appelés depuis le fil Excel, par la tâche en cours) ──
    def signe(self) -> None:
        """La tâche progresse : l'échéance repart pour le délai de la tâche."""
        self._echeance = time.monotonic() + self._delai_courant

    def accorder(self, delai_s: float) -> None:
        """Le prochain appel peut légitimement durer ``delai_s`` (ouverture d'un classeur…)."""
        self._echeance = time.monotonic() + max(delai_s, self._delai_courant)

    def executer(self, fonction: Callable[[], object], delai_s: float, detecter_connexion: bool = True):
        """Exécute ``fonction`` sur le fil Excel ; lève :class:`ExcelBloque` si
        elle ne donne aucun signe de vie pendant ``delai_s`` secondes (ou si
        elle n'a pas pu démarrer dans ce délai), et :class:`ConnexionRequise`
        dès qu'Excel affiche la fenêtre de connexion Microsoft 365."""
        resultat: list = []
        erreur: list[BaseException] = []
        fini = threading.Event()
        demarree = threading.Event()

        def tache() -> None:
            self._delai_courant = delai_s
            self.signe()
            demarree.set()
            try:
                resultat.append(fonction())
            except BaseException as exc:  # noqa: BLE001 — transmis à l'appelant
                _liberer_objets_com(exc)
                erreur.append(exc)
            finally:
                self._echeance = None
                fini.set()

        limite_demarrage = time.monotonic() + delai_s
        prochain_controle = time.monotonic() + 1.0
        self._taches.put(tache)
        while not fini.wait(0.1):
            maintenant = time.monotonic()
            if not demarree.is_set():
                if maintenant > limite_demarrage:
                    raise ExcelBloque(f"le fil Excel n'a pas pris la tâche en {delai_s:.0f} s")
                continue
            if detecter_connexion and maintenant >= prochain_controle:
                prochain_controle = maintenant + 1.0
                if fenetre_connexion(self.pid):
                    raise ConnexionRequise(
                        "Excel demande une connexion à Microsoft 365 (compte professionnel)"
                    )
            echeance = self._echeance
            if echeance is not None and maintenant > echeance:
                raise ExcelBloque(
                    f"Excel ne répond plus depuis plus de {delai_s:.0f} s — probablement une "
                    "boîte de dialogue de fusion SharePoint sans personne pour y répondre"
                )
        if erreur:
            raise erreur[0]
        return resultat[0]

    def arreter_proprement(self) -> None:
        """Fin de session normale : demande au fil de fermer Excel et de sortir."""
        self._taches.put(None)


# ── Session Excel ──────────────────────────────────────────────────
class SessionExcel:
    """Une instance Excel dédiée (jamais celle de l'utilisateur), fermée à la sortie.

        with SessionExcel() as session:
            bilan = session.appliquer_regles(chemin, config, ids, on_etape=...)

    Args:
        delai_occupe_s: attente maximale d'un Excel occupé (``Ready`` faux,
            appels refusés) avant de retenter le classeur.
        delai_blocage_s: inactivité maximale (aucun signe de vie) avant de
            considérer Excel comme bloqué, de le tuer et de le relancer.
        delai_ouverture_s: idem pour l'ouverture d'un classeur, plus longue
            (téléchargement depuis SharePoint).
    """

    TENTATIVES = 3   # essais complets d'un classeur si Excel reste occupé, bloqué ou déconnecté

    def __init__(self, visible: bool = False, delai_occupe_s: float = 120, delai_blocage_s: float = 90,
                 delai_ouverture_s: float = 180, connexion_interactive: bool = False,
                 delai_connexion_s: float = 300):
        self.visible = visible
        # Fenêtre de connexion Microsoft 365 : la montrer (Excel visible) pour que
        # l'utilisateur se connecte une fois, puis reprendre — plutôt qu'échouer.
        self.connexion_interactive = connexion_interactive
        self.delai_connexion_s = delai_connexion_s
        self.delai_occupe_s = delai_occupe_s
        self.delai_blocage_s = delai_blocage_s
        self.delai_ouverture_s = delai_ouverture_s
        self._fil: _FilExcel | None = None
        self._verrou = threading.Lock()
        self.redemarrages = 0

    @property
    def xl(self):
        return self._fil.xl if self._fil is not None else None

    def __enter__(self) -> "SessionExcel":
        self._fil = _FilExcel(self.visible, self.delai_occupe_s)
        return self

    def __exit__(self, *exc) -> bool:
        with self._verrou:
            if self._fil is not None:
                pid = self._fil.pid
                self._fil.arreter_proprement()
                self._fil = None
                # Une référence COM qui traînerait encore empêcherait Excel de se fermer,
                # le laissant invisible à verrouiller des fichiers : on termine CETTE
                # instance précise (identifiée par son PID), jamais celle de l'utilisateur.
                _terminer_si_toujours_ouvert(pid, attente_s=15)
        return False

    def _redemarrer(self, on_etape: Etape | None = None) -> None:
        """Tue le PROCESSUS de cette session et en relance un neuf sur un fil neuf.

        Le fil précédent, resté coincé dans l'appel bloqué, n'est jamais
        rejoint ; il reçoit quand même une pilule empoisonnée pour se nettoyer
        (``CoUninitialize``) s'il se débloque un jour. Si la nouvelle instance
        ne démarre pas, la session reste sans Excel et la prochaine commune
        retentera de le démarrer (:class:`ExcelIndisponible` sinon).
        """
        ancien, self._fil = self._fil, None
        if ancien is not None:
            if ancien.pid is None:
                _dire(on_etape, "⚠ processus Excel non identifié : il n'a pas pu être arrêté")
            _terminer_de_force(ancien.pid)
            ancien.arreter_proprement()
        self.redemarrages += 1
        self._fil = _FilExcel(self.visible, self.delai_occupe_s)

    # ── Application des règles ──────────────────────────────────────
    def appliquer_regles(
        self, chemin_fichier: str, config: ReglesConfig, ids_erreur: Iterable[str],
        on_etape: Etape | None = None,
    ) -> BilanFichier:
        """Applique les règles actives via une session Excel, avec relecture de vérification.

        Même contrat que ``services.excel.appliquer_regles`` ; le ``BilanFichier``
        porte en plus ``ouverture``, ``verrouille``, ``ids_non_verifies`` et
        ``tentatives``. Le classeur entier est retenté (l'écriture est
        idempotente : mêmes valeurs aux mêmes cellules) si Excel reste occupé,
        si un objet est invalidé par une fusion, si Excel se déconnecte ou s'il
        ne donne plus signe de vie — dans ces deux derniers cas, après avoir
        remplacé son processus. Sûr depuis plusieurs threads (appels sérialisés).
        """
        ids_erreur = list(ids_erreur)
        bilan, essai = self._avec_reprises(
            lambda fil, etape: self._appliquer_une_fois(fil, chemin_fichier, config, ids_erreur, etape),
            on_etape, adresse=url_sharepoint(chemin_fichier) or chemin_fichier,
        )
        bilan.tentatives = essai
        return bilan

    def copier_version_serveur(self, adresse: str, destination: str, on_etape: Etape | None = None) -> str:
        """Enregistre dans ``destination`` une copie du classeur ``adresse`` tel qu'Excel l'ouvre.

        Ajout de la copie vendorisée (cf. PROVENANCE.md), pour l'audit depuis
        SharePoint : ``adresse`` est l'URL SharePoint du fichier, ouvert **en
        lecture seule** avec le compte Office de l'utilisateur — c'est donc la
        version du serveur qui est copiée, pas la copie OneDrive locale. Mêmes
        garde-fous et reprises que :meth:`appliquer_regles`.
        """
        def copier(fil: _FilExcel, etape: Etape) -> str:
            fil.accorder(self.delai_ouverture_s)
            classeur = fil.xl.Workbooks.Open(
                adresse, 0, True, None, None, None, True, None, None, False, False, None, False,
            )
            try:
                fil.signe()
                self._attendre_pret(fil)
                classeur.SaveCopyAs(destination)
                fil.signe()
                return destination
            finally:
                try:
                    classeur.Close(False)
                except Exception:  # noqa: BLE001
                    pass

        return self._avec_reprises(copier, on_etape, adresse=adresse)[0]

    def _connexion_interactive(self, adresse: str, on_etape: Etape | None) -> None:
        """Ouvre ``adresse`` dans un Excel **visible** pour que l'utilisateur se connecte.

        Office mémorise la connexion (gestionnaire de comptes Windows) : les
        instances invisibles suivantes ouvrent alors les fichiers sans rien demander.
        """
        _dire(on_etape, "🔑 Microsoft 365 demande une connexion : une fenêtre Excel s'ouvre — "
                        "connectez-vous avec votre compte professionnel (elle se fermera ensuite)")
        fil = _FilExcel(True, self.delai_occupe_s)
        try:
            def ouvrir() -> None:
                try:
                    fil.xl.WindowState = -4137   # xlMaximized, au premier plan
                except Exception:  # noqa: BLE001
                    pass
                classeur = fil.xl.Workbooks.Open(adresse, 0, True)
                classeur.Close(False)

            fil.executer(ouvrir, delai_s=self.delai_connexion_s, detecter_connexion=False)
            _dire(on_etape, "🔑 Connexion Microsoft 365 effectuée — reprise du traitement")
        except ExcelBloque as exc:
            raise ConnexionRequise(
                f"connexion Microsoft 365 non effectuée en {self.delai_connexion_s:.0f} s"
            ) from exc
        finally:
            if not _fil_libre(fil):          # connexion abandonnée : l'appel est encore en cours
                _terminer_de_force(fil.pid)
            fil.arreter_proprement()
            _terminer_si_toujours_ouvert(fil.pid, attente_s=10)

    def _avec_reprises(self, travail: Callable[["_FilExcel", Etape], object], on_etape: Etape | None,
                       adresse: str | None = None):
        """Exécute ``travail(fil, etape)`` sur le fil Excel, avec reprises ; renvoie ``(résultat, essai)``.

        ``adresse`` : fichier ouvert par ``travail``, pour la connexion Microsoft 365
        interactive si Excel la demande (cf. ``connexion_interactive``).
        """
        connexion_faite = False
        with self._verrou:
            for essai in range(1, self.TENTATIVES + 1):
                if self._fil is None:
                    self._fil = _FilExcel(self.visible, self.delai_occupe_s)
                fil = self._fil
                etape = self._etape_du_fil(fil, on_etape)
                try:
                    resultat = fil.executer(lambda f=fil, e=etape: travail(f, e), delai_s=self.delai_blocage_s)
                    return resultat, essai
                except ConnexionRequise as exc:
                    self._redemarrer(on_etape)      # l'instance invisible reste bloquée sur la fenêtre
                    if not self.connexion_interactive or connexion_faite or not adresse:
                        raise ConnexionRequise(
                            f"{exc} — ouvrez une fois un fichier du site SharePoint dans Excel et "
                            "connectez-vous avec votre compte professionnel, puis relancez"
                        ) from None
                    self._connexion_interactive(adresse, on_etape)
                    connexion_faite = True
                except ExcelBloque as exc:
                    _dire(on_etape, f"⛔ {exc}")
                    self._redemarrer(on_etape)
                    if essai == self.TENTATIVES:
                        raise
                    _dire(on_etape, f"Excel redémarré — nouvelle tentative ({essai + 1}/{self.TENTATIVES})…")
                except Exception as exc:
                    if _est_deconnexion(exc):
                        if essai == self.TENTATIVES:
                            self._redemarrer(on_etape)
                            raise
                        _dire(on_etape, f"⚠ Excel s'est déconnecté — redémarrage puis nouvelle "
                                        f"tentative ({essai + 1}/{self.TENTATIVES})…")
                        self._redemarrer(on_etape)
                    elif _est_objet_invalide(exc):
                        if essai == self.TENTATIVES:
                            raise
                        _dire(on_etape, f"⚠ classeur rechargé par SharePoint pendant le traitement — "
                                        f"nouvelle tentative ({essai + 1}/{self.TENTATIVES})…")
                        time.sleep(2 * essai)
                    elif _est_rejet(exc) or isinstance(exc, ExcelOccupe):
                        if essai == self.TENTATIVES:
                            raise ExcelOccupe(
                                f"Excel est resté occupé ({self.TENTATIVES} essais de {self.delai_occupe_s} s) "
                                "— synchronisation ou co-édition en cours ? Les ID sont mis en attente, "
                                "rejouez-les plus tard"
                            ) from exc
                        _dire(on_etape, f"Excel occupé — nouvelle tentative ({essai}/{self.TENTATIVES})…")
                        time.sleep(5 * essai)
                    else:
                        raise
        raise AssertionError("inaccessible")  # pragma: no cover

    def _etape_du_fil(self, fil: _FilExcel, on_etape: Etape | None) -> Etape:
        """Rapporteur qui se tait dès que ``fil`` a été abandonné : une tentative
        débloquée après coup ne mélange pas ses messages à ceux de la suivante."""
        def etape(message: str) -> None:
            if on_etape is not None and self._fil is fil:
                on_etape(message)
        return etape

    def _attendre_pret(self, fil: _FilExcel) -> None:
        """Attend qu'Excel ait fini (chargement, synchro co-édition, recalcul, envoi)."""
        limite = time.monotonic() + self.delai_occupe_s
        while time.monotonic() < limite:
            try:
                if fil.xl.Ready:
                    fil.signe()
                    return
            except Exception as exc:
                if not _est_rejet(exc):
                    raise
            fil.signe()   # Excel répond (même « pas prêt ») : ce n'est pas un blocage
            _pythoncom().PumpWaitingMessages()
            time.sleep(0.25)
        raise ExcelOccupe(f"Excel n'est pas redevenu disponible après {self.delai_occupe_s} s")

    def _ouvrir(self, fil: _FilExcel, chemin: str):
        """Ouvre par l'adresse SharePoint si le fichier est synchronisé, sinon par le chemin local."""
        url = url_sharepoint(chemin)
        local = "co-édition SharePoint" if chemin.lower().startswith(("http://", "https://")) else \
            "fichier local OneDrive"
        erreur = None
        for cible, mode in ((url, "co-édition SharePoint"), (chemin, local)):
            if not cible:
                continue
            fil.accorder(self.delai_ouverture_s)
            try:
                classeur = fil.xl.Workbooks.Open(
                    cible, 0, False, None, None, None, True, None, None, False, False, None, False,
                )
                fil.signe()
                return classeur, mode
            except Exception as exc:  # noqa: BLE001
                erreur = exc
                if _est_deconnexion(exc):
                    break   # Excel mort : inutile d'essayer l'autre chemin avec lui
        raise erreur

    def _appliquer_une_fois(
        self, fil: _FilExcel, chemin_fichier: str, config: ReglesConfig, ids_erreur: list[str],
        on_etape: Etape | None,
    ) -> BilanFichier:
        regles = config.actives
        attendus = {normaliser_id(i) for i in ids_erreur} - {""}
        bilan = BilanFichier(comptes={r.code: 0 for r in regles})
        if not regles or not attendus:
            return bilan

        xl = fil.xl
        _dire(on_etape, "connexion à Excel…")
        classeur, bilan.ouverture = self._ouvrir(fil, chemin_fichier)
        calcul = None
        externe = False   # classeur déjà ouvert dans l'Excel de l'utilisateur : ne jamais le fermer
        try:
            self._attendre_pret(fil)
            if classeur.ReadOnly:
                try:
                    qui = classeur.WriteReservedBy
                except Exception:  # noqa: BLE001
                    qui = ""
                # Ajout (cf. PROVENANCE.md) : ouvert en écriture par l'utilisateur lui-même,
                # dans son propre Excel, notre instance ne l'a qu'en lecture seule ; on écrit
                # alors dans SON classeur ouvert (cellules seulement), qui reste ouvert.
                ouvert = _classeur_ouvert_ailleurs([url_sharepoint(chemin_fichier), chemin_fichier], fil.xl)
                if ouvert is not None:
                    try:
                        classeur.Close(False)
                    except Exception:  # noqa: BLE001
                        pass
                    classeur, externe, xl = ouvert, True, ouvert.Application
                    bilan.ouverture = "classeur déjà ouvert dans votre Excel"
                    _dire(on_etape, "classeur déjà ouvert dans votre Excel : écriture dans celui-ci "
                                    "(il reste ouvert ; vos modifications en cours sont enregistrées avec)")
                else:
                    bilan.verrouille = qui or "un autre utilisateur"
                    _dire(on_etape, f"⚠ classeur verrouillé par {bilan.verrouille} — rien écrit")
                    return bilan

            if not externe:   # le réglage de l'utilisateur n'est pas touché
                try:
                    classeur.AutoSaveOn = False   # un enregistrement explicite, pas une fusion continue
                except Exception:  # noqa: BLE001 — absent hors SharePoint, sans effet
                    pass

            feuille = _choisir_feuille_com(classeur)
            _dire(on_etape, f"ouverture {bilan.ouverture} — onglet « {feuille.Name} »")

            affichage = _retirer_filtres_com(feuille)
            bilan.visibilite_restauree = affichage_modifie(affichage)
            if bilan.visibilite_restauree:
                _dire(on_etape, "🔓 filtre retiré, lignes/colonnes masquées démasquées")
            fil.signe()

            entetes = self._lire_entetes(feuille, config.ligne_entete)
            table = table_depuis_entetes(entetes)
            index_id = _resoudre(config.colonne_id, table, "colonne des ID erreur")
            plans = {r.code: _plan_regle(r, table) for r in regles}
            lignes_par_id = self._index_lignes(feuille, index_id, config.ligne_entete)
            bilan.ids_trouves = attendus & set(lignes_par_id)
            bilan.nb_lignes = sum(len(lignes_par_id[i]) for i in bilan.ids_trouves)
            fil.signe()

            # Toutes les lignes d'un ID, comme le chemin openpyxl (services.excel) :
            # la version d'origine ne traitait que la première occurrence.
            cibles = sorted(
                (ligne, id_norm) for id_norm in bilan.ids_trouves for ligne in lignes_par_id[id_norm]
            )
            colonnes = {p.cible for p in plans.values()} | {i for p in plans.values() for i in p.refs.values()}
            cache = _CacheColonnes(feuille, (ligne for ligne, _ in cibles), colonnes)
            fil.signe()

            calcul = xl.Calculation
            xl.Calculation = _XL_CALC_MANUEL
            ecrits: list[tuple[str, str, int, int, object]] = []   # (règle, id, ligne, colonne, valeur)
            for ligne, id_norm in cibles:
                fil.signe()
                for regle in regles:
                    plan = plans[regle.code]
                    lire = cache.lecteur(ligne, plan.refs)
                    if not regle_sapplique(regle, lire):
                        continue
                    en_date = _colonne_est_date(entetes.get(plan.cible, ""))
                    valeur = valeur_pour_action(regle.action, lire, en_date=en_date)
                    if _conforme_valeur(cache.brut(ligne, plan.cible), valeur):
                        continue
                    _ecrire_cellule_com(feuille.Cells(ligne, plan.cible), valeur)
                    cache.ecrire(ligne, plan.cible, valeur)
                    ecrits.append((regle.code, id_norm, ligne, plan.cible, valeur))
            xl.Calculation = calcul
            calcul = None

            if not ecrits and not bilan.visibilite_restauree:
                _dire(on_etape, "valeurs déjà en place — fichier inchangé")
                return bilan

            self._attendre_pret(fil)
            _dire(on_etape, f"enregistrement de {len(ecrits)} modification(s)…")
            fil.signe()
            classeur.Save()
            fil.signe()
            self._attendre_pret(fil)
            bilan.enregistre = True

            self._verifier(fil, classeur, config, entetes, index_id, lignes_par_id, ecrits, bilan, on_etape)
            return bilan
        finally:
            if calcul is not None:
                try:
                    xl.Calculation = calcul
                except Exception:  # noqa: BLE001
                    pass
            if not externe:
                try:
                    classeur.Close(False)
                except Exception:  # noqa: BLE001
                    pass

    # ── Vérification après enregistrement ────────────────────────────
    def _verifier(self, fil, classeur, config, entetes, index_id, lignes_avant, ecrits, bilan, on_etape) -> None:
        """Relit chaque valeur écrite, à partir d'objets COM neufs, en retrouvant
        chaque ID par sa valeur (cf. docstring du module). Ne lève jamais."""
        if not ecrits:
            return
        try:
            feuille = _choisir_feuille_com(classeur)
            nouvelles = self._lire_entetes(feuille, config.ligne_entete)
            correspondance = _correspondance_colonnes(entetes, nouvelles)
            colonne_id = correspondance.get(index_id, index_id)
            lignes_apres = self._index_lignes(feuille, colonne_id, config.ligne_entete)
            fil.signe()

            emplacements = []
            for code, id_norm, ligne, colonne, valeur in ecrits:
                avant, apres = lignes_avant.get(id_norm, []), lignes_apres.get(id_norm, [])
                rang = avant.index(ligne) if ligne in avant else -1
                nouvelle = apres[rang] if 0 <= rang < len(apres) and len(apres) == len(avant) else None
                emplacements.append((code, id_norm, nouvelle, correspondance.get(colonne, colonne), valeur))
            relu = _CacheColonnes(
                feuille, (l for _, _, l, _, _ in emplacements if l is not None),
                {c for _, _, l, c, _ in emplacements if l is not None},
            )
            fil.signe()
        except Exception as exc:  # noqa: BLE001 — l'enregistrement a réussi : on ne met pas la commune en échec
            for _, id_norm, *_ in ecrits:
                if id_norm not in bilan.ids_non_verifies:
                    bilan.ids_non_verifies.append(id_norm)
            _dire(on_etape, f"⚠ relecture impossible après enregistrement ({_libelle(exc)}) — "
                            f"{len(bilan.ids_non_verifies)} ID remis en attente")
            return

        for code, id_norm, ligne, colonne, valeur in emplacements:
            if ligne is not None and _conforme_valeur(relu.brut(ligne, colonne), valeur):
                bilan.comptes[code] += 1
            elif id_norm not in bilan.ids_non_verifies:
                bilan.ids_non_verifies.append(id_norm)
                _dire(on_etape, f"⚠ {id_norm} : écriture non confirmée après enregistrement")

    # ── Lecture via COM ──────────────────────────────────────────────
    @staticmethod
    def _lire_entetes(feuille, ligne_entete: int) -> dict[int, str]:
        utile = feuille.UsedRange
        derniere_colonne = utile.Column + utile.Columns.Count - 1
        valeurs = _plat(feuille.Range(
            feuille.Cells(ligne_entete, 1), feuille.Cells(ligne_entete, derniere_colonne),
        ).Value)
        return {
            i: str(v).strip() for i, v in enumerate(valeurs, start=1)
            if v is not None and str(v).strip()
        }

    @staticmethod
    def _index_lignes(feuille, index_colonne: int, ligne_entete: int) -> dict[str, list[int]]:
        """``{ID erreur normalisé: [numéros de ligne]}`` — toutes les occurrences."""
        utile = feuille.UsedRange
        derniere_ligne = utile.Row + utile.Rows.Count - 1
        if derniere_ligne <= ligne_entete:
            return {}
        valeurs = _plat(feuille.Range(
            feuille.Cells(ligne_entete + 1, index_colonne), feuille.Cells(derniere_ligne, index_colonne),
        ).Value)
        lignes: dict[str, list[int]] = {}
        for decalage, valeur in enumerate(valeurs):
            identifiant = normaliser_id(valeur)
            if identifiant:
                lignes.setdefault(identifiant, []).append(ligne_entete + 1 + decalage)
        return lignes

    @staticmethod
    def _lecteur(feuille, ligne: int, refs: dict[str, int]):
        """Accesseur ``référence -> valeur`` pour une ligne donnée (dates reconverties)."""
        def lire(reference: str):
            index = refs.get(reference)
            if index is None:
                return None
            return _valeur_cellule_com(feuille, ligne, index)

        return lire


# ── Helpers ────────────────────────────────────────────────────────
def _fil_libre(fil: "_FilExcel") -> bool:
    """Vrai si le fil Excel n'a aucune tâche en cours (échéance levée)."""
    return fil._echeance is None


def _liberer_objets_com(exc: BaseException | None) -> None:
    """Vide, ici sur le fil Excel, les variables locales des cadres de la trace de ``exc``.

    Une exception levée pendant le traitement garde dans sa trace les cadres
    d'exécution, donc leurs variables locales : ``classeur``, ``feuille``,
    cellules… La trace forme un cycle (l'exception est rangée dans une liste
    que le cadre courant référence) : ces objets COM seraient libérés plus tard
    par le ramasse-miettes, **depuis n'importe quel thread** — y compris le fil
    Excel neuf d'une session redémarrée, en pleine écriture. Libérer un proxy
    COM hors de son appartement perturbe alors les appels en cours (feuille
    « à moitié initialisée » : ``AttributeError: <unknown>.Cells``,
    « Invalid number of arguments »). Le type et le message de l'exception
    sont conservés.
    """
    vues: set[int] = set()
    while exc is not None and id(exc) not in vues:
        vues.add(id(exc))
        if exc.__traceback__ is not None:
            traceback.clear_frames(exc.__traceback__)   # ignore le cadre en cours d'exécution
        exc = exc.__cause__ or exc.__context__


def _correspondance_colonnes(avant: dict[int, str], apres: dict[int, str]) -> dict[int, int]:
    """``{ancien index: nouvel index}`` d'après les libellés d'en-tête (colonne insérée par un collègue)."""
    if avant == apres:
        return {}
    par_nom = {}
    for index, nom in apres.items():
        par_nom.setdefault(normaliser(nom), index)
    correspondance = {}
    for index, nom in avant.items():
        if apres.get(index) == nom:
            continue
        nouveau = par_nom.get(normaliser(nom))
        if nouveau is not None:
            correspondance[index] = nouveau
    return correspondance


def _cle_classeur(nom: str) -> str:
    """Forme comparable d'un chemin ou d'une adresse de classeur."""
    from urllib.parse import unquote
    return unquote(str(nom or "")).replace("\\", "/").strip().rstrip("/").lower()


def _classeur_ouvert_ailleurs(cibles, notre_excel=None):
    """Classeur ``cibles`` (adresse SharePoint ou chemin) ouvert **en écriture** dans
    une autre instance d'Excel de cette session Windows (celle de l'utilisateur),
    trouvé dans la table des objets COM en cours ; ``None`` sinon. Ne lève jamais.
    """
    cles = {_cle_classeur(c) for c in cibles if c}
    if not cles:
        return None
    try:
        pythoncom = _pythoncom()
        import win32com.client as wc
        rot = pythoncom.GetRunningObjectTable()
        contexte = pythoncom.CreateBindCtx(0)
        notre = None
        try:
            notre = notre_excel.Hwnd if notre_excel is not None else None
        except Exception:  # noqa: BLE001
            pass
        for moniker in rot:
            try:
                if _cle_classeur(moniker.GetDisplayName(contexte, None)) not in cles:
                    continue
                classeur = wc.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
                if _cle_classeur(classeur.FullName) not in cles or classeur.ReadOnly:
                    continue
                if notre is not None and classeur.Application.Hwnd == notre:
                    continue
                return classeur
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        return None
    return None


def _libelle(exc: BaseException) -> str:
    codes = _codes_com(exc)
    if codes:
        return ", ".join(f"0x{c & 0xFFFFFFFF:08X}" for c in sorted(codes))
    return str(exc) or exc.__class__.__name__


def _colonne_est_date(nom_entete: str) -> bool:
    from .colonnes_excel import entete_est_date
    return entete_est_date(nom_entete)


def _resoudre(reference: str, table: dict[str, int], contexte: str) -> int:
    index = resoudre_colonne(reference, table)
    if index is None:
        raise ColonneIntrouvable(f"{contexte} : colonne « {reference} » introuvable dans le fichier")
    return index


def _dire(on_etape: Etape | None, message: str) -> None:
    if on_etape is not None:
        on_etape(message)
