"""Exports Excel (.xlsx) et CSV du rapport."""

import csv
import io
import os
import re

import pytest
from openpyxl import load_workbook

from avf.models.audit import NON_VALIDES, VALIDEES
from avf.services.audit import auditer_communes
from avf.services.export import exporter, exporter_csv, exporter_xlsx, nom_fichier_defaut
from avf.views.rapport_tableur import generer_csv

from .conftest import imm
from .test_audit import TOUTES


@pytest.fixture
def rapport(racine):
    return auditer_communes(TOUTES, racine)


def _lire_csv(chemin):
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f, delimiter=";"))


def test_csv_une_ligne_par_ligne_auditee(rapport, tmp_path):
    chemin = exporter_csv(rapport, str(tmp_path / "r.csv"))
    assert open(chemin, "rb").read(3) == b"\xef\xbb\xbf"          # BOM : Excel lit les accents
    lignes = _lire_csv(chemin)
    assert lignes[0] == ["Département", "Commune", "ID imm", "ID erreur", "Ligne", "Statut", "Audit",
                         "Date audit", "Auditeur", "Validation finale", "Motifs", "Fichier"]
    corps = lignes[1:]
    assert len(corps) == rapport.nb_lignes + len(rapport.en_echec)
    assert sum(1 for l in corps if l[5] == "Non valide") == 8
    assert sum(1 for l in corps if l[5] == "Validée") == 3
    l5 = next(l for l in corps if l[3] == "13001_4")
    assert l5[2:11] == [imm("13001", 5), "13001_4", "5", "Non valide", "OK", "05/01/2026", "Alice", "Validé",
                        "Validation finale « Validé » ≠ OK"]
    echec = next(l for l in corps if l[1] == "75020")
    assert echec[5] == "Erreur de lecture" and "Validation finale" in echec[10]


def test_csv_par_categorie(rapport):
    nv = list(csv.reader(io.StringIO(generer_csv(rapport, NON_VALIDES)), delimiter=";"))[1:]
    assert {l[5] for l in nv if l[2]} == {"Non valide"}
    ok = list(csv.reader(io.StringIO(generer_csv(rapport, VALIDEES)), delimiter=";"))[1:]
    assert sorted(l[2] for l in ok if l[2]) == [imm("13001", 2), imm("13002", 2), imm("13002", 3)]


def test_xlsx_onglets_et_contenu(rapport, tmp_path):
    chemin = exporter_xlsx(rapport, str(tmp_path / "r.xlsx"))
    wb = load_workbook(chemin)
    assert wb.sheetnames == ["Synthèse", "Communes", "Lignes", "ID imm non valides"]

    lignes = list(wb["Lignes"].iter_rows(values_only=True))
    assert lignes[0][2:4] == ("ID imm", "ID erreur")                    # ID imm avant ID erreur
    assert lignes[0][9] == "Validation finale" and lignes[0][10] == "Motifs"
    assert len(lignes) - 1 == rapport.nb_lignes + len(rapport.en_echec)
    assert wb["Lignes"].auto_filter.ref and wb["Lignes"].freeze_panes == "A2"

    communes = {r[1]: r for r in wb["Communes"].iter_rows(min_row=2, values_only=True)}
    assert communes["13001"][2:9] == ("À corriger", 6, 5, 6, 1, 5, pytest.approx(1 / 6))
    assert communes["13002"][2] == "Complète"

    feuille = wb["ID imm non valides"]
    assert [c.value for c in feuille[1]] == ["Département", "Commune", "ID imm", "ID erreur",
                                             "Lignes non valides", "Motifs"]
    ids = [r[2] for r in feuille.iter_rows(min_row=2, values_only=True)]
    assert ids == rapport.ids_de(NON_VALIDES)                           # les ID imm à copier
    ligne_7 = next(r for r in feuille.iter_rows(min_row=2, values_only=True) if r[2] == imm("13001", 7))
    assert ligne_7[3:5] == ("13001_5", "7") and "Audit vide" in ligne_7[5]

    rangees = [r for r in wb["Synthèse"].iter_rows(values_only=True) if r and r[0]]
    synthese = {r[0]: r[1:] for r in rangees}
    assert synthese["Lignes auditées"][0] == 11 and synthese["ID imm"][0] == 11
    assert [r[1] for r in rangees if r[0] == "  non valides"] == [8, 8]   # lignes, puis ID imm
    assert synthese["Date audit = une date"][0] == 7


def test_exporter_selon_le_format_et_nom(rapport, tmp_path):
    for format in ("html", "xlsx", "csv"):
        nom = nom_fichier_defaut(rapport, format)
        assert re.fullmatch(rf"rapport_audit_vf_\d{{8}}_\d{{4}}\.{format}", nom)
        chemin = exporter(rapport, str(tmp_path / nom), format)
        assert os.path.getsize(chemin) > 0 and not os.path.exists(chemin + ".tmp")


def test_export_impossible_ne_laisse_pas_de_temporaire(rapport, tmp_path):
    cible = tmp_path / "dossier"
    cible.mkdir()   # un dossier à la place du fichier : le remplacement échoue
    with pytest.raises(OSError):
        exporter_xlsx(rapport, str(cible))
    assert not (tmp_path / "dossier.tmp").exists()
