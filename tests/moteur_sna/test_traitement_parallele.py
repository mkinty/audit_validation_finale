"""Traitement en parallèle et arrêt — améliorations de la copie vendorisée (cf. sna/PROVENANCE.md)."""

from __future__ import annotations

import threading
import time

from openpyxl import load_workbook

from sna.models.base import ResultatCommune, Statut
from sna.services import attente
from sna.services.traitement import traiter_communes

from .test_traitement import CONFIG, VALEUR, _preparer_commune

COMMUNES = ["13001", "13002", "2A004", "75020", "75021", "69001"]


def _lot(tmp_path):
    groupes = {}
    for code in COMMUNES:
        ids = [f"{code}_{n}" for n in range(1, 6)]
        _preparer_commune(tmp_path, code, ids)
        groupes[code] = ids[:3]
    groupes["99999"] = ["99999_1"]   # introuvable
    return groupes


def _resume(resultats):
    return [(r.commune, r.statut, r.nb_modifications, r.nb_lignes, r.ids_absents) for r in resultats]


def test_parallele_identique_au_sequentiel(tmp_path):
    seq_dir, par_dir = tmp_path / "seq", tmp_path / "par"
    seq_dir.mkdir()
    par_dir.mkdir()
    groupes = _lot(seq_dir)
    _lot(par_dir)

    sequentiel = traiter_communes(groupes, CONFIG, str(seq_dir))
    progression, terminees = [], []
    parallele = traiter_communes(
        groupes, CONFIG, str(par_dir), paralleles=4,
        on_progression=lambda f, t, l: progression.append((f, t)), on_resultat=terminees.append,
    )

    assert _resume(parallele) == _resume(sequentiel)
    assert [r.commune for r in parallele] == list(groupes)          # ordre demandé, pas ordre de fin
    assert sorted(r.commune for r in terminees) == sorted(groupes)
    assert progression[0] == (0, 7) and progression[-1] == (7, 7)
    assert sorted(f for f, _ in progression[1:-1]) == list(range(1, 8))
    for code in COMMUNES:
        ws = load_workbook(par_dir / f"Dep{code[:2]}" / code / f"audit_{code}.xlsx")["Audit"]
        assert [ws[f"Z{i}"].value for i in range(2, 7)] == [VALEUR] * 3 + ["à revoir"] * 2


def test_callbacks_jamais_simultanes(tmp_path):
    groupes = _lot(tmp_path)
    en_cours, chevauchements = [0], []
    verrou = threading.Lock()

    def on_resultat(_):
        with verrou:
            en_cours[0] += 1
            chevauchements.append(en_cours[0])
        time.sleep(0.02)
        with verrou:
            en_cours[0] -= 1

    traiter_communes(groupes, CONFIG, str(tmp_path), paralleles=4, on_resultat=on_resultat)
    assert max(chevauchements) == 1


def test_arret_sequentiel_entre_deux_communes(tmp_path):
    groupes = _lot(tmp_path)
    appels = iter([False, True])
    resultats = traiter_communes(groupes, CONFIG, str(tmp_path), doit_arreter=lambda: next(appels))
    assert [r.commune for r in resultats] == ["13001"]
    ws = load_workbook(tmp_path / "Dep13" / "13002" / "audit_13002.xlsx")["Audit"]
    assert ws["Z2"].value == "à revoir"   # commune non commencée : intacte


def test_arret_en_parallele(tmp_path):
    groupes = _lot(tmp_path)
    arret = threading.Event()
    resultats = traiter_communes(
        groupes, CONFIG, str(tmp_path), paralleles=2, doit_arreter=arret.is_set,
        on_resultat=lambda r: arret.set(),
    )
    assert 1 <= len(resultats) < len(groupes)
    assert all(r.statut is Statut.SUCCES for r in resultats)
    ordre = list(groupes)
    assert [r.commune for r in resultats] == sorted((r.commune for r in resultats), key=ordre.index)


def test_session_excel_partagee_appelee_en_parallele(tmp_path):
    """Avec une session Excel, la préparation est parallèle et chaque commune lui est bien transmise."""
    from sna.services.excel import appliquer_regles

    groupes = _lot(tmp_path)

    class SessionFactice:
        def __init__(self):
            self.appels, self.simultanes, self._actifs = [], [], 0
            self._verrou = threading.Lock()

        def appliquer_regles(self, fichier, config, ids, on_etape=None):
            with self._verrou:   # comme SessionExcel : appels sérialisés
                self.appels.append(fichier)
                return appliquer_regles(fichier, config, ids, on_etape=on_etape)

    session = SessionFactice()
    resultats = traiter_communes(groupes, CONFIG, str(tmp_path), session_excel=session, paralleles=4)
    assert len(session.appels) == len(COMMUNES)
    assert [r.commune for r in resultats] == list(groupes)
    assert sum(r.nb_modifications for r in resultats) == 3 * len(COMMUNES)


def test_non_verifie_reconnu_quelle_que_soit_la_saisie():
    """Correctif : ``ids_non_verifies`` est normalisé, la saisie ne l'est pas forcément."""
    resultat = ResultatCommune(commune="2A004", statut=Statut.SUCCES, ids_non_verifies=["2A004_1"])
    echecs, reussis = attente.calculer_echecs_et_reussites({"2A004": ["2a004_1 ", "2A004_2"]}, [resultat])
    assert list(echecs) == ["2a004_1 "] and reussis == ["2A004_2"]


def test_sans_prereperage_la_session_est_toujours_appelee(tmp_path):
    """Ajout vendorisé : ``prereperage=False`` ouvre Excel même si la copie parcourue n'a pas l'ID."""
    from sna.services.excel import BilanFichier

    _preparer_commune(tmp_path, "13001", ["13001_1"])

    class Session:
        appels = []

        def appliquer_regles(self, fichier, config, ids, on_etape=None):
            self.appels.append(list(ids))
            return BilanFichier(comptes={"R1": 0}, ids_trouves=set())

    session = Session()
    traiter_communes({"13001": ["13001_9"]}, CONFIG, str(tmp_path), session_excel=session)
    assert session.appels == []                       # défaut : repérage local, Excel non ouvert
    traiter_communes({"13001": ["13001_9"]}, CONFIG, str(tmp_path), session_excel=session, prereperage=False)
    assert session.appels == [["13001_9"]]
