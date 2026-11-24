"""Persistance de la configuration : racine et règles."""

import json

from sna.config import (
    ROOT_DEFAULT, charger_racine, charger_regles, load_config, save_config,
    stocker_racine, stocker_regles,
)
from sna.models.regles import Action, RegleModification, ReglesConfig


def test_defauts_quand_le_fichier_est_absent(tmp_path):
    cfg = load_config(str(tmp_path / "absent.json"))

    assert charger_racine(cfg) == ROOT_DEFAULT
    assert charger_regles(cfg) == ReglesConfig.defaut()


def test_config_par_defaut_livre_les_regles_d_audit(tmp_path):
    """Un config.json neuf est écrit avec les règles d'audit par défaut."""
    fichier = str(tmp_path / "config.json")

    cfg = load_config(fichier)          # fichier absent → défauts
    save_config(cfg, fichier)

    regles = charger_regles(load_config(fichier))
    attendus = ReglesConfig.defaut().regles
    assert [r.code for r in regles.regles] == [r.code for r in attendus]
    assert len(regles.regles) == len(attendus) > 0
    assert regles == ReglesConfig.defaut()


def test_racine_et_regles_persistent(tmp_path):
    fichier = str(tmp_path / "config.json")
    cfg = load_config(fichier)
    regles = ReglesConfig(
        regles=[RegleModification(
            code="R1", libelle="statut", action=Action(colonne="B", valeur="y")
        )],
        ligne_entete=2,
        colonne_id="Identifiant erreur",
    )

    stocker_racine(cfg, r"D:\audits")
    stocker_regles(cfg, regles)
    assert save_config(cfg, fichier) is True

    relu = load_config(fichier)
    assert charger_racine(relu) == r"D:\audits"
    assert charger_regles(relu) == regles


def test_les_regles_configurees_sont_relues_telles_quelles(tmp_path):
    """Ce que l'utilisateur a configuré (ou supprimé) est conservé au chargement."""
    fichier = tmp_path / "config.json"
    fichier.write_text(
        json.dumps({"root_path_be": "X", "regles": {"regles": [
            {"code": "R1", "libelle": "Audit", "colonne": "Audit", "valeur": "OK"}
        ]}}),
        encoding="utf-8",
    )

    regles = charger_regles(load_config(str(fichier)))

    assert [r.colonne for r in regles.regles] == ["Audit"]
    assert regles.regles[0].valeur == "OK"


def test_json_corrompu_retombe_sur_les_defauts(tmp_path):
    fichier = tmp_path / "config.json"
    fichier.write_text("{ ceci n'est pas du JSON", encoding="utf-8")

    assert charger_racine(load_config(str(fichier))) == ROOT_DEFAULT


def test_regles_corrompues_retombent_sur_les_defauts(tmp_path):
    fichier = tmp_path / "config.json"
    fichier.write_text(
        json.dumps({"root_path_be": "X", "regles": {"regles": [{"inattendu": 1}]}}),
        encoding="utf-8",
    )
    cfg = load_config(str(fichier))

    assert charger_regles(cfg) == ReglesConfig.defaut()


def test_ancienne_configuration_sans_colonne_id(tmp_path):
    """Un config.json antérieur reste lisible : la colonne ID prend son défaut."""
    fichier = tmp_path / "config.json"
    fichier.write_text(
        json.dumps({
            "root_path_be": "X",
            "regles": {"ligne_entete": 1, "regles": [
                {"code": "R1", "libelle": "l", "colonne": "Z", "valeur": "v"}
            ]},
        }),
        encoding="utf-8",
    )

    regles = charger_regles(load_config(str(fichier)))

    assert regles.colonne_id == "ID erreur"
    assert regles.regles[0].colonne == "Z"


def test_cles_manquantes_sont_completees(tmp_path):
    fichier = tmp_path / "config.json"
    fichier.write_text(json.dumps({"root_path_be": "X"}), encoding="utf-8")

    cfg = load_config(str(fichier))

    assert charger_racine(cfg) == "X"
    assert charger_regles(cfg) == ReglesConfig.defaut()


def test_ancienne_regle_action_condition_est_reprise(tmp_path):
    """Un config.json de la version précédente conserve colonne et valeur."""
    fichier = tmp_path / "config.json"
    fichier.write_text(
        json.dumps({
            "root_path_be": "X",
            "regles": {"ligne_entete": 1, "regles": [{
                "code": "R1",
                "libelle": "ancienne règle",
                "actif": True,
                "ligne_debut": 2,
                "exact": True,
                "action": {"colonne": "Z", "source": "valeur", "valeur": "nouvelle"},
                "conditions": [{
                    "colonne": "Z", "operateur": "=",
                    "type_droite": "valeur", "droite": "ancienne",
                }],
            }]},
        }),
        encoding="utf-8",
    )

    regle = charger_regles(load_config(str(fichier))).regles[0]

    assert regle.colonne == "Z"
    assert regle.valeur == "nouvelle"
