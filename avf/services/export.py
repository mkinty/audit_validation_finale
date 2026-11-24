"""Export du rapport d'audit en fichier HTML, et ouverture dans le navigateur."""

from __future__ import annotations

import os
import webbrowser
from pathlib import Path

from ..models.audit import RapportAudit
from ..views.rapport_html import generer_html


def nom_fichier_defaut(rapport: RapportAudit) -> str:
    """Ex. ``rapport_validation_finale_20261124_1432.html``."""
    return f"rapport_validation_finale_{rapport.date:%Y%m%d_%H%M}.html"


def exporter_html(rapport: RapportAudit, chemin: str) -> str:
    """Écrit le rapport (UTF-8) et retourne son chemin absolu.

    L'écriture passe par un fichier temporaire remplacé d'un coup : un rapport
    ouvert dans le navigateur n'est jamais vu à moitié écrit.
    """
    cible = Path(chemin).expanduser().resolve()
    cible.parent.mkdir(parents=True, exist_ok=True)
    temporaire = cible.with_name(cible.name + ".tmp")
    temporaire.write_text(generer_html(rapport), encoding="utf-8")
    os.replace(temporaire, cible)
    return str(cible)


def ouvrir_dans_navigateur(chemin: str) -> bool:
    """Ouvre le rapport dans le navigateur par défaut."""
    return webbrowser.open(Path(chemin).resolve().as_uri())
