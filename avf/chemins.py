"""Emplacement des fichiers de configuration et de données de l'outil.

    - développement (script Python)      → à la racine du projet ;
    - application packagée (PyInstaller) → dossier de données utilisateur.

Ces fichiers sont **distincts** de ceux de complete_excel_files : les deux
outils peuvent tourner sur le même poste sans se marcher dessus.
"""

from __future__ import annotations

import os
import sys

APP_NAME = "audit_validation_finale"

RACINE_PROJET = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def dossier_donnees() -> str:
    if not getattr(sys, "frozen", False):
        return RACINE_PROJET
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    dossier = os.path.join(base, APP_NAME)
    os.makedirs(dossier, exist_ok=True)
    return dossier


CONFIG_FILE = os.path.join(dossier_donnees(), "config_audit_vf.json")

# ID erreur non appliqués lors d'une mise à jour (cf. sna.services.attente).
IDS_EN_ATTENTE_FILE = os.path.join(dossier_donnees(), "ids_en_attente.json")
