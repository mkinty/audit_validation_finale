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

Ne dépend que de ``pywin32`` (absent de toute autre partie du programme) :
sans lui, ou sans Excel installé, :func:`excel_disponible` renvoie faux et
l'appelant retombe sur ``services.excel`` (réécriture directe du fichier).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from datetime import datetime, timedelta
import math
import time

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

# Signature du rapporteur d'étapes (journal en temps réel).
Etape = Callable[[str], None]


class ExcelIndisponible(Exception):
    """pywin32 absent, ou Excel non installé / impossible à démarrer."""


class ExcelOccupe(Exception):
    """Excel a refusé les appels trop longtemps (occupé par une synchronisation, un calcul…)."""


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


def _est_rejet(exc) -> bool:
    return bool(getattr(exc, "args", None)) and exc.args[0] in _REJETS


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
    from datetime import date as _date
    if isinstance(valeur, _date):
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


def _conforme_com(cellule, attendu) -> bool:
    lu = cellule.Value2
    if attendu is None or attendu == "":
        return lu in (None, "")
    serie = _serie(attendu)
    if serie is not None:
        return isinstance(lu, (int, float)) and abs(float(lu) - serie) < 1e-6
    if isinstance(attendu, bool):
        return lu == attendu
    if isinstance(attendu, (int, float)):
        return isinstance(lu, (int, float)) and math.isclose(float(lu), float(attendu), rel_tol=1e-9, abs_tol=1e-9)
    return str(lu if lu is not None else "").replace("\r\n", "\n") == str(attendu).replace("\r\n", "\n")


# ── Session Excel ──────────────────────────────────────────────────
class SessionExcel:
    """Une instance Excel dédiée (jamais celle de l'utilisateur), fermée à la sortie.

        with SessionExcel() as session:
            bilan = session.appliquer_regles(chemin, config, ids, on_etape=...)
    """

    TENTATIVES = 3   # essais complets d'un classeur si Excel reste occupé au-delà du délai

    def __init__(self, visible: bool = False, delai_occupe_s: float = 120):
        self.visible = visible
        self.delai_occupe_s = delai_occupe_s
        self.xl = None
        self._filtre = None
        self._pid = None

    def __enter__(self) -> "SessionExcel":
        try:
            pythoncom = _pythoncom()
            import win32com.client as wc
        except ImportError as exc:
            raise ExcelIndisponible("pywin32 n'est pas installé") from exc
        pythoncom.CoInitialize()
        self._filtre = _FiltreMessages(delai_max_ms=int(self.delai_occupe_s * 1000))
        self._filtre.installer()
        try:
            self.xl = wc.DispatchEx("Excel.Application")    # nouvelle instance, n'interfère pas avec l'Excel ouvert
            self._pid = _pid_de(self.xl)
            self.xl.Visible = bool(self.visible)
            self.xl.DisplayAlerts = False
            self.xl.AskToUpdateLinks = False
            self.xl.ScreenUpdating = bool(self.visible)
        except Exception as exc:
            self.__exit__(None, None, None)
            raise ExcelIndisponible(f"impossible de démarrer Excel : {exc}") from exc
        return self

    def __exit__(self, *exc) -> bool:
        if self.xl is not None:
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
            self._filtre = None
        _pythoncom().CoUninitialize()
        # Une référence COM qui traînerait encore empêcherait Excel de se fermer,
        # le laissant invisible à verrouiller des fichiers : on termine CETTE
        # instance précise (identifiée par son PID), jamais celle de l'utilisateur.
        _terminer_si_toujours_ouvert(self._pid, attente_s=15)
        self._pid = None
        return False

    def _attendre_pret(self) -> None:
        """Attend qu'Excel ait fini (chargement, synchro co-édition, recalcul, envoi)."""
        limite = time.time() + self.delai_occupe_s
        while time.time() < limite:
            try:
                if self.xl.Ready:
                    return
            except Exception as exc:
                if not _est_rejet(exc):
                    raise
            _pythoncom().PumpWaitingMessages()
            time.sleep(0.25)
        raise ExcelOccupe(f"Excel n'est pas redevenu disponible après {self.delai_occupe_s} s")

    def _ouvrir(self, chemin: str):
        """Ouvre par l'adresse SharePoint si le fichier est synchronisé, sinon par le chemin local."""
        url = url_sharepoint(chemin)
        erreur = None
        for cible, mode in ((url, "co-édition SharePoint"), (chemin, "fichier local OneDrive")):
            if not cible:
                continue
            try:
                classeur = self.xl.Workbooks.Open(
                    cible, 0, False, None, None, None, True, None, None, False, False, None, False,
                )
                return classeur, mode
            except Exception as exc:  # noqa: BLE001
                erreur = exc
        raise erreur

    # ── Application des règles ──────────────────────────────────────
    def appliquer_regles(
        self, chemin_fichier: str, config: ReglesConfig, ids_erreur: Iterable[str],
        on_etape: Etape | None = None,
    ) -> BilanFichier:
        """Applique les règles actives via une session Excel, avec relecture de vérification.

        Même contrat que ``services.excel.appliquer_regles`` ; le ``BilanFichier``
        porte en plus ``ouverture``, ``verrouille`` et ``ids_non_verifies``. Si
        Excel reste occupé au-delà du délai, le classeur entier est retenté
        (l'écriture est idempotente : mêmes valeurs aux mêmes cellules).
        """
        ids_erreur = list(ids_erreur)
        for essai in range(1, self.TENTATIVES + 1):
            try:
                bilan = self._appliquer_une_fois(chemin_fichier, config, ids_erreur, on_etape)
                bilan.tentatives = essai
                return bilan
            except Exception as exc:
                if not (_est_rejet(exc) or isinstance(exc, ExcelOccupe)):
                    raise
                if essai == self.TENTATIVES:
                    raise ExcelOccupe(
                        f"Excel est resté occupé ({self.TENTATIVES} essais de {self.delai_occupe_s} s) "
                        "— synchronisation ou co-édition en cours ? Les ID sont mis en attente, "
                        "rejouez-les plus tard"
                    ) from exc
                _dire(on_etape, f"Excel occupé — nouvelle tentative ({essai}/{self.TENTATIVES})…")
                time.sleep(5 * essai)

    def _appliquer_une_fois(
        self, chemin_fichier: str, config: ReglesConfig, ids_erreur: list[str],
        on_etape: Etape | None,
    ) -> BilanFichier:
        regles = config.actives
        attendus = {normaliser_id(i) for i in ids_erreur} - {""}
        bilan = BilanFichier(comptes={r.code: 0 for r in regles})
        if not regles or not attendus:
            return bilan

        _dire(on_etape, "connexion à Excel…")
        classeur, bilan.ouverture = self._ouvrir(chemin_fichier)
        calcul = None
        try:
            self._attendre_pret()
            if classeur.ReadOnly:
                try:
                    qui = classeur.WriteReservedBy
                except Exception:  # noqa: BLE001
                    qui = ""
                bilan.verrouille = qui or "un autre utilisateur"
                _dire(on_etape, f"⚠ classeur verrouillé par {bilan.verrouille} — rien écrit")
                return bilan

            feuille = _choisir_feuille_com(classeur)
            _dire(on_etape, f"ouverture {bilan.ouverture} — onglet « {feuille.Name} »")

            affichage = _retirer_filtres_com(feuille)
            bilan.visibilite_restauree = affichage_modifie(affichage)
            if bilan.visibilite_restauree:
                _dire(on_etape, "🔓 filtre retiré, lignes/colonnes masquées démasquées")

            entetes = self._lire_entetes(feuille, config.ligne_entete)
            table = table_depuis_entetes(entetes)
            index_id = _resoudre(config.colonne_id, table, "colonne des ID erreur")
            plans = {r.code: _plan_regle(r, table) for r in regles}
            lignes_par_id = self._index_lignes(feuille, index_id, config.ligne_entete)
            bilan.ids_trouves = attendus & set(lignes_par_id)
            bilan.nb_lignes = sum(len(lignes_par_id[i]) for i in bilan.ids_trouves)

            calcul = self.xl.Calculation
            self.xl.Calculation = _XL_CALC_MANUEL
            ecrits: dict[str, list] = {}   # code règle -> [(id, cellule, valeur), ...]
            # Toutes les lignes d'un ID, comme le chemin openpyxl (services.excel) :
            # la version d'origine ne traitait que la première occurrence.
            cibles = sorted(
                (ligne, id_norm) for id_norm in bilan.ids_trouves for ligne in lignes_par_id[id_norm]
            )
            for ligne, id_norm in cibles:
                for regle in regles:
                    plan = plans[regle.code]
                    lire = self._lecteur(feuille, ligne, plan.refs)
                    if not regle_sapplique(regle, lire):
                        continue
                    en_date = _colonne_est_date(entetes.get(plan.cible, ""))
                    valeur = valeur_pour_action(regle.action, lire, en_date=en_date)
                    cellule = feuille.Cells(ligne, plan.cible)
                    if _conforme_com(cellule, valeur):
                        continue
                    _ecrire_cellule_com(cellule, valeur)
                    ecrits.setdefault(regle.code, []).append((id_norm, cellule, valeur))
            self.xl.Calculation = calcul
            calcul = None

            total_ecrits = sum(len(v) for v in ecrits.values())
            if not total_ecrits and not bilan.visibilite_restauree:
                _dire(on_etape, "valeurs déjà en place — fichier inchangé")
                return bilan

            self._attendre_pret()
            _dire(on_etape, f"enregistrement de {total_ecrits} modification(s)…")
            classeur.Save()
            self._attendre_pret()

            touches: set[str] = set()
            for code, faits in ecrits.items():
                for id_norm, cellule, valeur in faits:
                    touches.add(id_norm)
                    if classeur.Saved and _conforme_com(cellule, valeur):
                        bilan.comptes[code] += 1
                    elif id_norm not in bilan.ids_non_verifies:
                        bilan.ids_non_verifies.append(id_norm)
                        _dire(on_etape, f"⚠ {id_norm} : écriture non confirmée après enregistrement")
            bilan.enregistre = True
            return bilan
        finally:
            if calcul is not None:
                try:
                    self.xl.Calculation = calcul
                except Exception:  # noqa: BLE001
                    pass
            try:
                classeur.Close(False)
            except Exception:  # noqa: BLE001
                pass

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
        """``{ID erreur normalisé: [numéros de ligne]}`` (toutes les occurrences)."""
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
