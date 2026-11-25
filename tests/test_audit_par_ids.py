"""Audit restreint à des ID erreur saisis (au lieu de communes entières)."""

import csv
import io

from openpyxl import load_workbook

from avf.config import Parametres, charger, enregistrer
from avf.services.audit import auditer_communes
from avf.services.detection import verifier_ids
from avf.services.export import exporter_xlsx
from avf.services.mise_a_jour import groupes_depuis_rapport
from avf.views import presentation as p
from avf.views.rapport_html import generer_html
from avf.views.rapport_tableur import generer_csv

IDS = {"13001": ["13001_2", "13001_5", "13001_99"], "2A004": ["2a004_2"]}


def test_saisie_libre_regroupee_par_commune(racine):
    s = verifier_ids(racine, "13001_2, 13001_5\n2a004_2 ; 99999_1 13002  13001_2")
    assert s.groupes == {"13001": ["13001_2", "13001_5"], "2A004": ["2A004_2"]}
    assert s.absents == {"99999": ["99999_1"]}      # pas de dossier Dep99/99999
    assert s.codes_sans_id == ["13002"]               # code INSEE sans numéro : ignoré
    assert s.nb_ids == 3 and s.ids == ["13001_2", "13001_5", "2A004_2"]
    assert [c.code for c in s.trouvees] == ["13001", "2A004"]


def test_saisie_vide():
    s = verifier_ids("x", "rien d'exploitable ici")
    assert s.groupes == {} and s.nb_ids == 0


def test_seules_les_lignes_des_ids_saisis_sont_auditees(racine):
    r = auditer_communes(["13001", "2A004"], racine, ids=IDS)
    c = r.commune("13001")
    assert [(l.id_erreur, l.ligne) for l in c.lignes] == [("13001_2", 3), ("13001_5", 6), ("13001_5", 7)]
    assert c.nb_ids == 2 and c.ids_demandes == ["13001_2", "13001_5", "13001_99"]
    assert c.ids_introuvables == ["13001_99"]
    assert any("introuvable(s) dans le fichier : 13001_99" in a for a in c.avertissements)
    assert r.commune("2A004").ids_non_valides == ["2A004_2"]     # 2A004_1 non demandé : hors audit
    assert r.par_ids and r.nb_ids_demandes == 4
    assert r.ids_introuvables() == ["13001_99"]
    assert r.nb_lignes == 4


def test_commune_hors_saisie_auditee_entierement(racine):
    r = auditer_communes(["13001", "13002"], racine, ids={"13001": ["13001_1"]})
    assert r.commune("13001").nb_lignes == 1
    assert r.commune("13002").nb_lignes == 2 and r.commune("13002").ids_demandes is None


def test_commune_illisible_ses_ids_sont_introuvables(racine):
    r = auditer_communes(["13003"], racine, ids={"13003": ["13003_1"]})
    assert r.ids_introuvables() == ["13003_1"]


def test_audit_parallele_par_ids(racine):
    seq = auditer_communes(["13001", "2A004"], racine, ids=IDS)
    par = auditer_communes(["13001", "2A004"], racine, ids=IDS, paralleles=4)
    assert [(c.commune, c.ids_non_valides, c.ids_introuvables) for c in par.communes] == \
        [(c.commune, c.ids_non_valides, c.ids_introuvables) for c in seq.communes]


def test_tous_valides_dans_le_perimetre(racine):
    r = auditer_communes(["13001"], racine, ids={"13001": ["13001_1"]})
    assert r.commune("13001").complete           # seul l'ID saisi compte


def test_mise_a_jour_limitee_aux_ids_saisis(racine):
    r = auditer_communes(["13001", "2A004"], racine, ids=IDS)
    assert groupes_depuis_rapport(r) == {"13001": ["13001_2", "13001_5"], "2A004": ["2A004_2"]}


def test_perimetre_dans_le_resume_et_les_rapports(racine, tmp_path):
    r = auditer_communes(["13001", "2A004"], racine, ids=IDS)
    assert p.resume(r)[-1] == "🎯 Audit limité à 4 ID erreur saisi(s) — 1 introuvable(s) : 13001_99"
    assert "Audit limité à 4 ID erreur saisi(s)" in generer_html(r)
    lignes = list(csv.reader(io.StringIO(generer_csv(r)), delimiter=";"))
    introuvable = next(l for l in lignes if l[2] == "13001_99")
    assert introuvable[4] == "ID introuvable" and introuvable[9] == "absent du fichier audit"
    wb = load_workbook(exporter_xlsx(r, str(tmp_path / "r.xlsx")))
    synthese = {l[0]: l[1] for l in wb["Synthèse"].iter_rows(values_only=True) if l and l[0]}
    assert synthese["Périmètre"].startswith("Audit limité à 4 ID erreur")


def test_audit_complet_sans_perimetre(racine):
    r = auditer_communes(["13001"], racine)
    assert not r.par_ids and r.ids_introuvables() == []
    assert not p.resume(r)[-1].startswith("🎯")


def test_ids_cibles_memorises(tmp_path):
    fichier = str(tmp_path / "cfg.json")
    enregistrer(Parametres(ids_cibles=["13001_2", "2A004_1"]), fichier)
    assert charger(fichier).ids_cibles == ["13001_2", "2A004_1"]
