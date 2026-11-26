"""Audit restreint à des ID saisis — ID erreur ou ID imm — au lieu de communes entières."""

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

from .conftest import ecrire_audit, imm

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
    assert c.nb_ids == 3 and c.ids_demandes == ["13001_2", "13001_5", "13001_99"]   # 13001_5 : 2 ID imm
    assert c.ids_introuvables == ["13001_99"]
    assert any("introuvable(s) dans le fichier : 13001_99" in a for a in c.avertissements)
    assert r.commune("2A004").ids_non_valides == [imm("2A004", 3)]   # 2A004_1 non demandé : hors audit
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
    assert groupes_depuis_rapport(r) == {"13001": [imm("13001", n) for n in (3, 6, 7)],
                                         "2A004": [imm("2A004", 3)]}


def test_perimetre_dans_le_resume_et_les_rapports(racine, tmp_path):
    r = auditer_communes(["13001", "2A004"], racine, ids=IDS)
    assert p.resume(r)[-1] == "🎯 Audit limité à 4 ID erreur saisi(s) — 1 introuvable(s) : 13001_99"
    assert "Audit limité à 4 ID erreur saisi(s)" in generer_html(r)
    lignes = list(csv.reader(io.StringIO(generer_csv(r)), delimiter=";"))
    introuvable = next(l for l in lignes if l[3] == "13001_99")      # colonne ID erreur
    assert introuvable[2] == "" and introuvable[5] == "ID introuvable" and introuvable[10] == "absent du fichier audit"
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



# ── Audit restreint à des ID imm ──────────────────────────────────
def test_insee_d_un_id_imm():
    from avf.services.detection import extraire_ids_imm, grouper_ids_imm, insee_de_id_imm
    assert insee_de_id_imm("ID-58006000005V") == "58006"
    assert insee_de_id_imm("ID-2B030000006D") == "2B030"
    assert insee_de_id_imm("id-2b030000006e ") == "2B030"
    assert insee_de_id_imm("13001_2") == "" and insee_de_id_imm("ID-58") == ""
    texte = "ID imm\nID-58006000005V\nID-58006000005S, id-2B030000006D;ID-58006000005V xID-1300100000 13001_2"
    assert extraire_ids_imm(texte) == ["ID-58006000005V", "ID-58006000005S", "ID-2B030000006D"]
    assert grouper_ids_imm(extraire_ids_imm(texte)) == {
        "58006": ["ID-58006000005V", "ID-58006000005S"], "2B030": ["ID-2B030000006D"]}


def test_saisie_d_id_imm_regroupee_par_commune(racine):
    s = verifier_ids(racine, f"{imm('13001', 3)} {imm('13001', 6)}\n{imm('2A004', 2)} ID-9999900000AB",
                     type_ids="id_imm")
    assert s.groupes == {"13001": [imm("13001", 3), imm("13001", 6)], "2A004": [imm("2A004", 2)]}
    assert s.absents == {"99999": ["ID-9999900000AB"]} and s.codes_sans_id == []


def test_audit_par_id_imm_une_seule_ligne_d_un_id_erreur_partage(racine):
    """13001_5 couvre deux logements : saisir l'ID imm de l'un n'audite que sa ligne."""
    r = auditer_communes(["13001"], racine, ids={"13001": [imm("13001", 7), "ID-1300199999ZZ"]},
                         type_ids="id_imm")
    c = r.commune("13001")
    assert [(l.ligne, l.id_erreur) for l in c.lignes] == [(7, "13001_5")]
    assert c.ids_non_valides == [imm("13001", 7)] and c.type_ids == "id_imm"
    assert c.ids_introuvables == ["ID-1300199999ZZ"]
    assert any("ID imm introuvable(s)" in a and "autre commune" in a for a in c.avertissements)
    assert r.type_ids == "id_imm"
    assert p.resume(r)[-1] == "🎯 Audit limité à 2 ID imm saisi(s) — 1 introuvable(s) : ID-1300199999ZZ"
    lignes = list(csv.reader(io.StringIO(generer_csv(r)), delimiter=";"))
    assert next(l for l in lignes if l[2] == "ID-1300199999ZZ")[5] == "ID introuvable"


def test_id_imm_au_prefixe_d_une_autre_commune(tmp_path):
    """Cas réel (diagnostic) : ID-3803300000FS dans le fichier de 38074. Saisi seul, il est
    cherché dans 38033 ; audité via 38074, il est trouvé et signalé."""
    racine = str(tmp_path)
    ecrire_audit(f"{racine}/Dep38/38074/audit_38074.xlsx",
                 [["38074_56", "a", None, None, None, None, "ID-3803300000FS"]])
    r = auditer_communes(["38074"], racine, ids={"38074": ["ID-3803300000FS"]}, type_ids="id_imm")
    assert r.commune("38074").ids_non_valides == ["ID-3803300000FS"]
    assert any("au code d'une autre commune" in a for a in r.commune("38074").avertissements)
