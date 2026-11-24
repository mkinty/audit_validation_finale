"""Lecture / écriture de la configuration (dossier racine ET règles).

Tout ce qui peut varier d'un poste à l'autre est regroupé dans un unique
``config.json``, modifiable depuis l'interface :

    root_path_be → dossier racine des fichiers audit
    regles       → règles de modification, ligne d'en-tête et colonne des
                   ID erreur, construites par l'utilisateur

Les valeurs ci-dessous ne servent que de repli au tout premier lancement.
"""

from __future__ import annotations

import json
import os

from .chemins import CONFIG_FILE
from .models.regles import ReglesConfig

# Chemin proposé si config.json est absent ou illisible.
ROOT_DEFAULT = r"C:\Users\u269775\Altice Campus SFR\Swap Adresse - Etude Cible\audit_SNA"

CLE_RACINE = "root_path_be"
CLE_REGLES = "regles"

DEFAULT_CFG = {
    CLE_RACINE: ROOT_DEFAULT,
    CLE_REGLES: ReglesConfig.defaut().to_dict(),
}


# ── Fichier de configuration ──────────────────────────────────────
def load_config(config_file: str = CONFIG_FILE) -> dict:
    """Charge la configuration ; complète les clés manquantes par les défauts."""
    if os.path.isfile(config_file):
        try:
            with open(config_file, encoding="utf-8") as f:
                cfg = json.load(f)
            for cle, valeur in DEFAULT_CFG.items():
                cfg.setdefault(cle, valeur)
            return cfg
        except (OSError, json.JSONDecodeError):
            pass
    return json.loads(json.dumps(DEFAULT_CFG))  # copie profonde


def save_config(cfg: dict, config_file: str = CONFIG_FILE) -> bool:
    """Enregistre la configuration. Retourne False si l'écriture a échoué."""
    try:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        return True
    except OSError:
        return False


# ── Racine ────────────────────────────────────────────────────────
def charger_racine(cfg: dict | None = None) -> str:
    """Dossier racine des fichiers audit."""
    cfg = load_config() if cfg is None else cfg
    return str(cfg.get(CLE_RACINE) or ROOT_DEFAULT)


def stocker_racine(cfg: dict, racine: str) -> None:
    """Écrit la racine dans le dict de configuration (à sauvegarder ensuite)."""
    cfg[CLE_RACINE] = str(racine).strip()


# ── Règles ────────────────────────────────────────────────────────
def charger_regles(cfg: dict) -> ReglesConfig:
    """Extrait les règles du dict de configuration.

    Les règles enregistrées sont relues **telles quelles** : ce que
    l'utilisateur a configuré (ou supprimé) est conservé d'un démarrage à
    l'autre. En l'absence de données exploitables, la configuration livrée
    (16 règles d'audit par défaut) est retournée.
    """
    data = cfg.get(CLE_REGLES)
    if not data:
        return ReglesConfig.defaut()
    try:
        return ReglesConfig.from_dict(data)
    except Exception:  # noqa: BLE001 — données corrompues → retour aux défauts
        return ReglesConfig.defaut()


def stocker_regles(cfg: dict, regles: ReglesConfig) -> None:
    """Écrit les règles dans le dict de configuration (à sauvegarder ensuite)."""
    cfg[CLE_REGLES] = regles.to_dict()
