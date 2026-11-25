"""Robustesse de l'écriture via Excel — améliorations de la copie vendorisée.

Cf. ``sna/PROVENANCE.md`` : doublons, cache de colonnes, garde-fou par
inactivité, appels concurrents, relecture par ID, erreurs COM transitoires.
Ces tests pilotent un vrai Excel (ignorés sans Excel / pywin32).
"""

from __future__ import annotations

import threading
import time

import pytest
from openpyxl import load_workbook

from sna.models.regles import Condition
from sna.services.excel import BilanFichier
from sna.services.excel_com import (
    SessionExcel, _FilExcel, _terminer_de_force, _terminer_si_toujours_ouvert, excel_disponible,
)

from .test_excel_com import _classeur, _config, _regle

pytestmark = pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")


def _com_error(hresult, scode=None):
    import pywintypes
    info = (0, "Microsoft Excel", "erreur simulée", None, 0, scode) if scode is not None else None
    return pywintypes.com_error(hresult, "erreur simulée", info, None)


def test_doublon_toutes_les_lignes_ecrites(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"], ["13001_2", "b"], ["13001_1", "c"]])
    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1"])
    assert bilan.nb_lignes == 2 and bilan.comptes == {"R1": 2} and bilan.ids_non_verifies == []
    ws = load_workbook(chemin)["Audit"]
    assert [ws.cell(r, 2).value for r in (2, 3, 4)] == ["CORRIGE", "b", "CORRIGE"]


def test_conditions_lisent_les_valeurs_ecrites_par_les_regles_precedentes(tmp_path):
    """Le cache de colonnes reporte les écritures : R2 voit la valeur posée par R1."""
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Audit", "Validation finale"], [["13001_1", None, None]])
    r1 = _regle(colonne="Audit", valeur="OK", code="R1")
    r2 = _regle(colonne="Validation finale", valeur="OK", code="R2",
                conditions=[Condition("Audit", "=", droite="OK")])
    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, _config(r1, r2), ["13001_1"])
    assert bilan.comptes == {"R1": 1, "R2": 1}
    ws = load_workbook(chemin)["Audit"]
    assert (ws["B2"].value, ws["C2"].value) == ("OK", "OK")


def test_beaucoup_de_lignes_et_dates(tmp_path):
    """Lecture groupée par colonne : valeurs et dates identiques à la lecture cellule par cellule."""
    from datetime import datetime
    chemin = str(tmp_path / "audit.xlsx")
    lignes = [[f"13001_{n}", datetime(2026, 1, 1 + n % 28), None] for n in range(1, 401)]
    _classeur(chemin, ["ID erreur", "Date audit", "Statut"], lignes)
    regle = _regle(colonne="Statut", valeur="OK",
                   conditions=[Condition("Date audit", ">", droite="15/01/2026")])
    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, _config(regle), [l[0] for l in lignes])
    attendus = sum(1 for l in lignes if l[1] > datetime(2026, 1, 15))
    assert bilan.comptes == {"R1": attendus} and bilan.ids_non_verifies == []
    ws = load_workbook(chemin)["Audit"]
    assert sum(1 for r in range(2, 402) if ws.cell(r, 3).value == "OK") == attendus


def test_travail_long_mais_actif_n_est_pas_un_blocage():
    """Le délai porte sur l'inactivité : une tâche qui donne signe de vie peut durer."""
    fil = _FilExcel(visible=False, delai_occupe_s=30)
    try:
        def longue():
            for _ in range(10):
                time.sleep(0.1)
                fil.signe()
            return "fini"
        assert fil.executer(longue, delai_s=0.4) == "fini"
    finally:
        fil.arreter_proprement()
        _terminer_si_toujours_ouvert(fil.pid, attente_s=10)


def test_appels_concurrents_ne_sont_pas_pris_pour_un_blocage(tmp_path):
    """Plusieurs threads sur la même session : l'attente de son tour ne compte pas."""
    chemins = []
    for n in range(3):
        chemin = str(tmp_path / f"audit_{n}.xlsx")
        _classeur(chemin, ["ID erreur", "Statut"], [[f"1300{n}_1", "x"]])
        chemins.append(chemin)
    bilans, erreurs = {}, []

    with SessionExcel(delai_blocage_s=4) as session:
        def travail(n):
            try:
                bilans[n] = session.appliquer_regles(chemins[n], _config(_regle()), [f"1300{n}_1"])
            except Exception as exc:  # noqa: BLE001
                erreurs.append(exc)
        fils = [threading.Thread(target=travail, args=(n,)) for n in range(3)]
        for f in fils:
            f.start()
        for f in fils:
            f.join(120)
        assert session.redemarrages == 0

    assert erreurs == []
    assert all(bilans[n].comptes == {"R1": 1} for n in range(3))


def test_verification_retrouve_l_id_apres_insertion_de_lignes(tmp_path):
    """Un collègue insère des lignes pendant l'enregistrement : la relecture suit l'ID."""
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"], ["13001_2", "b"]])

    class SessionCollegue(SessionExcel):
        def _verifier(self, fil, classeur, *args, **kwargs):
            feuille = classeur.Worksheets("Audit")
            feuille.Rows(2).Insert()
            feuille.Rows(2).Insert()
            feuille.Cells(2, 1).Value2 = "AUTRE_1"
            return super()._verifier(fil, classeur, *args, **kwargs)

    with SessionCollegue() as session:
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_2"])
    assert bilan.comptes == {"R1": 1} and bilan.ids_non_verifies == []


def test_relecture_impossible_met_les_id_en_attente_sans_echec(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"], ["13001_2", "b"]])
    appels = []

    class SessionRelectureCassee(SessionExcel):
        def _lire_entetes(self, feuille, ligne_entete):
            appels.append(1)
            if len(appels) > 1:   # le 2e appel est celui de la relecture
                raise _com_error(-2147352567, -2146827864)
            return super()._lire_entetes(feuille, ligne_entete)

    messages = []
    with SessionRelectureCassee() as session:
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1", "13001_2"],
                                         on_etape=messages.append)
    assert bilan.enregistre
    assert sorted(bilan.ids_non_verifies) == ["13001_1", "13001_2"]
    assert bilan.comptes == {"R1": 0}
    assert any("relecture impossible" in m and "0x800A01A8" in m for m in messages)
    assert load_workbook(chemin)["Audit"]["B2"].value == "CORRIGE"   # l'écriture a bien eu lieu


def test_objet_invalide_par_une_fusion_est_retente_sans_redemarrer(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"]])
    with SessionExcel() as session:
        original, appels = session._appliquer_une_fois, []

        def une_fois(*args, **kwargs):
            appels.append(1)
            if len(appels) == 1:
                raise _com_error(-2147352567, -2146827864)
            return original(*args, **kwargs)

        session._appliquer_une_fois = une_fois
        messages = []
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1"], on_etape=messages.append)
        assert session.redemarrages == 0
    assert bilan.tentatives == 2 and bilan.comptes == {"R1": 1}
    assert any("rechargé par SharePoint" in m for m in messages)


def test_deconnexion_d_excel_redemarre_l_instance(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"]])
    with SessionExcel() as session:
        pid_initial = session._fil.pid
        original, appels = session._appliquer_une_fois, []

        def une_fois(*args, **kwargs):
            appels.append(1)
            if len(appels) == 1:
                raise _com_error(-2147417848)   # RPC_E_DISCONNECTED
            return original(*args, **kwargs)

        session._appliquer_une_fois = une_fois
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1"])
        assert session.redemarrages == 1 and session._fil.pid != pid_initial
    assert bilan.tentatives == 2 and bilan.comptes == {"R1": 1}


def test_session_sans_instance_en_redemarre_une(tmp_path):
    """Un redémarrage qui a échoué laisse la session sans Excel : la commune suivante en relance un."""
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"]])
    with SessionExcel() as session:
        _terminer_de_force(session._fil.pid)
        session._fil.arreter_proprement()
        session._fil = None
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1"])
    assert bilan.comptes == {"R1": 1}


def test_messages_d_une_tentative_abandonnee_sont_ignores(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "a"]])
    messages = []
    with SessionExcel(delai_blocage_s=2) as session:
        original, appels = session._appliquer_une_fois, []

        def parfois_bloque(fil, chemin_, config, ids, etape):
            appels.append(1)
            if len(appels) == 1:
                time.sleep(4)
                etape("MESSAGE PERIME")
                return BilanFichier(comptes={})
            return original(fil, chemin_, config, ids, etape)

        session._appliquer_une_fois = parfois_bloque
        bilan = session.appliquer_regles(chemin, _config(_regle()), ["13001_1"], on_etape=messages.append)
        time.sleep(3)   # laisse la tentative abandonnée se réveiller et tenter d'écrire au journal
    assert bilan.comptes == {"R1": 1}
    assert "MESSAGE PERIME" not in messages
