"""Localisation des fichiers audit sur le disque.

Arborescence attendue :
    <racine>/Dep<XX>/<commune>/audit_*.xlsx

où ``XX`` correspond aux deux premiers caractères du code commune.
"""

from __future__ import annotations

from datetime import datetime
import os
import re

PREFIXE_DEP = "Dep"
PREFIXE_AUDIT = "audit_"
EXTENSION = ".xlsx"

# Une date au format AAAA-MM-JJ ou JJ/MM/AAAA (/ ou -) dans le nom du fichier,
# telle qu'OneDrive/SharePoint en ajoute aux copies de conflit
# (« audit_13001-DESKTOP-AB12CD3.xlsx », « audit_13001 (copie en conflit 2026-11-24).xlsx »…).
_MOTIF_DATE_NOM = re.compile(r"\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{4}")


def dossier_commune(racine: str, commune: str) -> str:
    """Chemin attendu du dossier d'une commune (existant ou non)."""
    commune = str(commune).strip().upper()
    return os.path.join(racine, f"{PREFIXE_DEP}{commune[:2]}", commune)


def trouver_fichier_audit(racine: str, commune: str) -> str | None:
    """Retourne le chemin du fichier audit d'une commune, ou ``None``.

    Le premier fichier commençant par ``audit_`` et finissant par ``.xlsx``
    dans le dossier de la commune est retenu.
    """
    dossier = dossier_commune(racine, commune)
    if not os.path.isdir(dossier):
        return None
    return _premier_audit(dossier)


def detecter_fichiers_conflit(racine: str, commune: str) -> list[dict]:
    """Fichiers ``audit_*.xlsx`` surnuméraires dans le dossier d'une commune.

    Le dossier d'une commune ne devrait contenir qu'un seul fichier audit.
    Un deuxième (ou plus) est le signe probable d'une copie de conflit laissée
    par OneDrive/SharePoint après une synchronisation concurrente (le fichier
    retenu par :func:`trouver_fichier_audit` n'est jamais inclus). Pour chaque
    fichier supplémentaire, la date retenue est celle présente dans son nom si
    elle y figure, sinon sa date de dernière modification sur le disque.

    Retourne une liste de ``{"chemin": str, "date": str}``, vide si aucune
    commune, aucun dossier ou aucun doublon.
    """
    dossier = dossier_commune(racine, commune)
    principal = _premier_audit(dossier)
    if principal is None:
        return []
    try:
        noms = sorted(os.listdir(dossier))
    except OSError:
        return []
    return [
        {"chemin": chemin, "date": _date_conflit(nom, chemin)}
        for nom in noms
        if nom.startswith(PREFIXE_AUDIT) and nom.endswith(EXTENSION)
        for chemin in (os.path.join(dossier, nom),)
        if chemin != principal
    ]


def _date_conflit(nom: str, chemin: str) -> str:
    correspondance = _MOTIF_DATE_NOM.search(nom)
    if correspondance:
        return correspondance.group(0)
    try:
        return datetime.fromtimestamp(os.path.getmtime(chemin)).strftime("%d/%m/%Y %H:%M")
    except OSError:
        return "date inconnue"


def lister_communes(racine: str) -> list[str]:
    """Codes communes présents sous ``<racine>/Dep*/``, triés.

    Ne lève jamais : un dossier illisible est simplement ignoré.
    """
    communes: list[str] = []
    for departement in _sous_dossiers(racine, prefixe=PREFIXE_DEP):
        communes.extend(os.path.basename(d) for d in _sous_dossiers(departement))
    return sorted(set(communes))


def trouver_fichier_exemple(racine: str) -> str | None:
    """Premier fichier audit trouvé sous la racine, tous départements confondus.

    Sert à peupler automatiquement les listes déroulantes de colonnes du
    panneau de règles, sans demander de fichier à l'utilisateur.
    """
    for departement in _sous_dossiers(racine, prefixe=PREFIXE_DEP):
        for commune in _sous_dossiers(departement):
            fichier = _premier_audit(commune)
            if fichier:
                return fichier
    return None


# ── Helpers ───────────────────────────────────────────────────────
def _sous_dossiers(chemin: str, prefixe: str = "") -> list[str]:
    """Sous-dossiers directs, triés, sans jamais lever d'erreur d'accès."""
    try:
        entrees = sorted(os.listdir(chemin))
    except OSError:
        return []
    return [
        os.path.join(chemin, nom)
        for nom in entrees
        if nom.startswith(prefixe) and os.path.isdir(os.path.join(chemin, nom))
    ]


def _premier_audit(dossier: str) -> str | None:
    try:
        noms = sorted(os.listdir(dossier))
    except OSError:
        return None
    return next(
        (
            os.path.join(dossier, nom)
            for nom in noms
            if nom.startswith(PREFIXE_AUDIT) and nom.endswith(EXTENSION)
        ),
        None,
    )
