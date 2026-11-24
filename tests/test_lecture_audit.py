import os
from datetime import datetime

import pytest

from avf.models.colonnes import ColonnesAudit
from avf.services import lecture_audit
from avf.services.lecture_audit import ColonneAbsente, formater_valeur, lire_audit
from sna.services.colonnes_excel import OngletIntrouvable
from sna.services.lecture_rapide import LectureImpossible

from .conftest import ENTETES, ecrire_audit


@pytest.fixture
def forcer_openpyxl(monkeypatch):
    def refuser(*args, **kwargs):
        raise LectureImpossible("forcé pour le test")

    monkeypatch.setattr(lecture_audit, "_lire_rapide", refuser)


def _chemin(racine, dep, commune):
    return os.path.join(racine, f"Dep{dep}", commune, f"audit_{commune}.xlsx")


def _resume(lecture):
    return [
        (l.ligne, l.id_erreur, l.validation, l.audit, l.date_audit, l.auditeur)
        for l in lecture.lignes
    ]


ATTENDU_13001 = [
    (2, "13001_1", "OK", "OK", "12/03/2026", "Alice"),
    (3, "13001_2", "", "", "", ""),
    (4, "13001_3", "", "KO", "", "Bob"),
    (5, "13001_4", "Validé", "OK", "05/01/2026", "Alice"),
    (6, "13001_5", "OK", "", "", ""),
    (7, "13001_5", "", "", "", ""),
]


def test_lecture_rapide(racine):
    lecture = lire_audit(_chemin(racine, "13", "13001"), "ID erreur", ColonnesAudit())
    assert lecture.mode == "rapide"
    assert _resume(lecture) == ATTENDU_13001
    assert lecture.nb_lignes_lues == 6  # dernière ligne portant un ID
    assert lecture.avertissements == []


def test_lecture_openpyxl_equivalente(racine, forcer_openpyxl):
    lecture = lire_audit(_chemin(racine, "13", "13001"), "ID erreur", ColonnesAudit())
    assert lecture.mode == "openpyxl"
    assert _resume(lecture) == ATTENDU_13001


def test_validation_espaces_seuls_lue_vide(racine):
    lecture = lire_audit(_chemin(racine, "13", "13001"), "ID erreur", ColonnesAudit())
    assert not lecture.lignes[2].validee  # « 13001_3 » : « ␣␣␣ »


def test_validation_numerique(racine):
    lecture = lire_audit(_chemin(racine, "13", "13002"), "ID erreur", ColonnesAudit())
    assert [l.validation for l in lecture.lignes] == ["OK", "1"]


@pytest.mark.parametrize("forcer", [False, True])
def test_onglet_audit_par_nom_pas_le_premier(racine, request, forcer):
    if forcer:
        request.getfixturevalue("forcer_openpyxl")
    lecture = lire_audit(_chemin(racine, "2A", "2A004"), "ID erreur", ColonnesAudit())
    assert [l.id_erreur for l in lecture.lignes] == ["2A004_1", "2a004_2"]


def test_onglet_absent(tmp_path):
    chemin = ecrire_audit(str(tmp_path / "audit_x.xlsx"), [["1_1"]], nom_feuille="Feuil1")
    with pytest.raises(OngletIntrouvable):
        lire_audit(chemin, "ID erreur", ColonnesAudit())


def test_colonne_validation_absente(racine):
    with pytest.raises(ColonneAbsente, match="Validation finale"):
        lire_audit(_chemin(racine, "75", "75020"), "ID erreur", ColonnesAudit())


def test_colonne_id_absente(racine):
    with pytest.raises(ColonneAbsente, match="Identifiant"):
        lire_audit(_chemin(racine, "13", "13001"), "Identifiant", ColonnesAudit())


def test_colonne_informative_absente_avertit(tmp_path):
    chemin = ecrire_audit(
        str(tmp_path / "audit_x.xlsx"), [["13001_1", "OK"]],
        entetes=["ID erreur", "Validation finale"],
    )
    lecture = lire_audit(chemin, "ID erreur", ColonnesAudit())
    assert len(lecture.avertissements) == 3
    assert lecture.lignes[0].validation == "OK"
    assert lecture.lignes[0].audit == ""


def test_noms_de_colonnes_personnalises_et_casse(tmp_path):
    chemin = ecrire_audit(
        str(tmp_path / "audit_x.xlsx"), [["13001_1", "", "Moi"]],
        entetes=["id ERREUR", "VF", "Contrôleur"],
    )
    colonnes = ColonnesAudit(validation="vf", auditeur="contrôleur")
    lecture = lire_audit(chemin, "ID erreur", colonnes)
    assert lecture.lignes[0].auditeur == "Moi"
    assert not lecture.lignes[0].validee


def test_libelle_court_absent_pas_pris_pour_une_lettre(tmp_path):
    """« VF » absent ne doit pas être lu comme la colonne physique VF."""
    chemin = ecrire_audit(str(tmp_path / "audit_x.xlsx"), [["13001_1", "OK"]],
                          entetes=["ID erreur", "Validation finale"])
    with pytest.raises(ColonneAbsente):
        lire_audit(chemin, "ID erreur", ColonnesAudit(validation="Vf"))


def test_lettre_explicite_acceptee(tmp_path):
    chemin = ecrire_audit(str(tmp_path / "audit_x.xlsx"), [["13001_1", "OK"]],
                          entetes=["ID erreur", "Validation finale"])
    lecture = lire_audit(chemin, "A", ColonnesAudit(validation="B"))
    assert lecture.lignes[0].validation == "OK"


def test_ligne_entete_decalee(tmp_path):
    chemin = ecrire_audit(
        str(tmp_path / "audit_x.xlsx"), [ENTETES, ["13001_1", "a", None, None, None, "OK"]],
        entetes=["Titre du rapport"],
    )
    lecture = lire_audit(chemin, "ID erreur", ColonnesAudit(), ligne_entete=2)
    assert [(l.ligne, l.validation) for l in lecture.lignes] == [(3, "OK")]


def test_gros_fichier(tmp_path):
    lignes = [[f"13001_{i}", "a", None, None, None, "OK" if i % 3 else None] for i in range(5000)]
    chemin = ecrire_audit(str(tmp_path / "audit_x.xlsx"), lignes)
    lecture = lire_audit(chemin, "ID erreur", ColonnesAudit())
    assert len(lecture.lignes) == 5000
    assert sum(1 for l in lecture.lignes if not l.validee) == len(range(0, 5000, 3))


@pytest.mark.parametrize("valeur, en_date, brut, attendu", [
    (None, False, False, ""),
    (datetime(2026, 3, 12), False, False, "12/03/2026"),
    (datetime(2026, 3, 12, 8, 30), False, False, "12/03/2026 08:30"),
    ("46093", True, True, "12/03/2026"),
    ("46093.5", True, True, "12/03/2026 12:00"),
    ("12", True, True, "12"),            # trop petit pour une date : inchangé
    ("texte", True, True, "texte"),
    (3.0, False, False, "3"),
    ("3.0", False, True, "3"),
    (True, False, False, "VRAI"),
    ("  OK ", False, False, "OK"),
])
def test_formater_valeur(valeur, en_date, brut, attendu):
    assert formater_valeur(valeur, en_date=en_date, brut=brut) == attendu
