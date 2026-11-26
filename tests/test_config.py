import json

from avf.config import Parametres, charger, enregistrer
from avf.models.colonnes import ColonnesAudit
from sna.models.regles import ReglesConfig


def test_defauts(tmp_path):
    p = charger(str(tmp_path / "absent.json"))
    assert p.racine == "" and p.date_du_jour and p.reauditer and p.ecrire_via_excel
    assert p.colonnes == ColonnesAudit()
    assert p.regles.to_dict() == ReglesConfig.defaut().to_dict()


def test_aller_retour(tmp_path):
    fichier = str(tmp_path / "cfg.json")
    p = Parametres(racine="D:/audit", communes=["13001"],
                   colonnes=ColonnesAudit(validation="VF"), date_du_jour=False,
                   ecrire_via_excel=False, reauditer=False, dossier_export="D:/r")
    p.regles.regles[2].action.valeur = "Moi"
    p.regles.colonne_id = p.colonnes.id_imm      # identifiant des mises à jour : l'ID imm
    assert enregistrer(p, fichier)
    relu = charger(fichier)
    assert relu.to_dict() == p.to_dict()
    assert relu.regles.regles[2].action.valeur == "Moi"


def test_fichier_corrompu(tmp_path):
    fichier = tmp_path / "cfg.json"
    fichier.write_text("{pas du json", encoding="utf-8")
    assert charger(str(fichier)).to_dict() == Parametres().to_dict()


def test_cles_manquantes_et_regles_corrompues(tmp_path):
    fichier = tmp_path / "cfg.json"
    fichier.write_text(json.dumps({"racine": "X", "regles": {"regles": [{"pas_de_code": 1}]},
                                   "colonnes": {"validation": "  "}}), encoding="utf-8")
    p = charger(str(fichier))
    assert p.racine == "X"
    assert p.regles.to_dict() == {**ReglesConfig.defaut().to_dict(), "colonne_id": "ID imm"}
    assert p.colonnes.validation == "Validation finale"


def test_migration_identifiant_id_imm(tmp_path):
    """Ancienne configuration : les règles identifiaient les lignes par l'ID erreur (colonne
    éventuellement renommée). Désormais : ID imm ; le nom de la colonne ID erreur est conservé."""
    fichier = tmp_path / "cfg.json"
    fichier.write_text(json.dumps({"regles": {"colonne_id": "Identifiant erreur", "regles": []}}),
                       encoding="utf-8")
    p = charger(str(fichier))
    assert p.regles.colonne_id == "ID imm" and p.colonnes.id_erreur == "Identifiant erreur"
    fichier.write_text(json.dumps({"colonnes": {"id_imm": "N° imm", "id_erreur": "Err"},
                                   "type_ids": "id_imm"}), encoding="utf-8")
    p = charger(str(fichier))
    assert (p.regles.colonne_id, p.colonnes.id_erreur, p.type_ids) == ("N° imm", "Err", "id_imm")


def test_enregistrement_impossible(tmp_path):
    assert not enregistrer(Parametres(), str(tmp_path / "inexistant" / "cfg.json"))


def test_ancienne_config_avec_departements(tmp_path):
    """La clé « departements » de la version précédente est simplement ignorée."""
    fichier = tmp_path / "cfg.json"
    fichier.write_text(json.dumps({"racine": "X", "departements": ["13"], "communes": ["13001"]}),
                       encoding="utf-8")
    p = charger(str(fichier))
    assert p.communes == ["13001"]
    assert "departements" not in p.to_dict()


def test_paralleles_borne_et_par_defaut(tmp_path):
    fichier = tmp_path / "cfg.json"
    for valeur, attendu in ((0, 1), (99, 16), ("x", 4), (3, 3)):
        fichier.write_text(json.dumps({"paralleles": valeur}), encoding="utf-8")
        assert charger(str(fichier)).paralleles == attendu
    assert Parametres().paralleles == 4


def test_source_et_lien_sharepoint(tmp_path):
    fichier = tmp_path / "cfg.json"
    fichier.write_text(json.dumps({"source": "sharepoint", "lien_sharepoint": " https://x/sites/a/b "}),
                       encoding="utf-8")
    p = charger(str(fichier))
    assert (p.source, p.lien_sharepoint) == ("sharepoint", "https://x/sites/a/b")
    fichier.write_text(json.dumps({"source": "nimporte"}), encoding="utf-8")
    assert charger(str(fichier)).source == "local"
    assert Parametres().source == "local"
