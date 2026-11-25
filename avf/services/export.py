"""Export du rapport d'audit (HTML, Excel, CSV) et ouverture du fichier produit.

Chaque écriture passe par un fichier temporaire remplacé d'un coup : un
rapport ouvert dans le navigateur ou dans Excel n'est jamais vu à moitié écrit.
"""

from __future__ import annotations

import os
from pathlib import Path
import webbrowser

from ..models.audit import TOUTES, RapportAudit
from ..views.rapport_html import generer_html
from ..views.rapport_tableur import generer_csv, generer_xlsx

FORMATS = {
    "html": ("Page HTML", "*.html"),
    "xlsx": ("Classeur Excel", "*.xlsx"),
    "csv": ("CSV (Excel, séparateur « ; »)", "*.csv"),
}


def nom_fichier_defaut(rapport: RapportAudit, extension: str = "html") -> str:
    """Ex. ``rapport_audit_vf_20261125_1432.xlsx``."""
    return f"rapport_audit_vf_{rapport.date:%Y%m%d_%H%M}.{extension}"


def exporter(rapport: RapportAudit, chemin: str, format: str = "html") -> str:
    """Écrit le rapport au ``format`` demandé et retourne le chemin absolu."""
    if format == "xlsx":
        return exporter_xlsx(rapport, chemin)
    if format == "csv":
        return exporter_csv(rapport, chemin)
    return exporter_html(rapport, chemin)


def exporter_html(rapport: RapportAudit, chemin: str) -> str:
    """Rapport HTML autonome (UTF-8)."""
    return _ecrire(chemin, lambda tmp: tmp.write_text(generer_html(rapport), encoding="utf-8"))


def exporter_csv(rapport: RapportAudit, chemin: str, categorie: str = TOUTES) -> str:
    """Détail CSV (UTF-8 avec BOM, « ; ») : une ligne par ligne auditée."""
    def ecrire(tmp: Path) -> None:
        with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
            f.write(generer_csv(rapport, categorie))
    return _ecrire(chemin, ecrire)


def exporter_xlsx(rapport: RapportAudit, chemin: str) -> str:
    """Classeur Excel : synthèse, communes, lignes, ID non valides."""
    return _ecrire(chemin, lambda tmp: generer_xlsx(rapport).save(tmp))


def _ecrire(chemin: str, ecrire) -> str:
    cible = Path(chemin).expanduser().resolve()
    cible.parent.mkdir(parents=True, exist_ok=True)
    temporaire = cible.with_name(cible.name + ".tmp")
    try:
        ecrire(temporaire)
        os.replace(temporaire, cible)
    finally:
        if temporaire.exists():
            temporaire.unlink()
    return str(cible)


def ouvrir_dans_navigateur(chemin: str) -> bool:
    """Ouvre le rapport HTML dans le navigateur par défaut."""
    return webbrowser.open(Path(chemin).resolve().as_uri())


def ouvrir_fichier(chemin: str) -> bool:
    """Ouvre un fichier avec l'application associée (Excel pour .xlsx / .csv)."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(chemin)  # type: ignore[attr-defined]
            return True
        return webbrowser.open(Path(chemin).resolve().as_uri())
    except OSError:
        return False
