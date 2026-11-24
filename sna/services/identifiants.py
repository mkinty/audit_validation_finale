"""Extraction des « ID erreur » saisis par l'utilisateur, et regroupement par commune.

Un **ID erreur** identifie une anomalie précise à l'intérieur d'une commune. Il
s'écrit ``<insee>_<numéro>`` :

    13001_1     75020_27     2A004_3

Le code INSEE en préfixe suffit à retrouver le fichier audit de la commune
(``<racine>/Dep<XX>/<insee>/audit_*.xlsx``), tandis que l'identifiant complet
sert à sélectionner les lignes à modifier dans ce fichier, via la colonne
« ID erreur ».

Le regroupement produit un dictionnaire ``{insee: [ID erreur, …]}`` : chaque
fichier n'est donc ouvert qu'une fois, quel que soit le nombre d'ID erreur
demandés pour cette commune.
"""

from __future__ import annotations

from collections.abc import Iterable
import re

# Code INSEE : 5 chiffres (métropole/DOM) ou format corse 2A/2B + 3 chiffres.
_INSEE = r"(?:2[AB]\d{3}|\d{5})"

# ID erreur complet : code INSEE, un souligné, puis le numéro de l'anomalie.
_MOTIF_ID = re.compile(rf"\b({_INSEE})_([A-Z0-9]+)\b")

# Code INSEE isolé, c'est-à-dire *non* suivi d'un souligné (« _ » étant un
# caractère de mot, la limite \b finale l'exclut d'elle-même).
_MOTIF_INSEE_SEUL = re.compile(rf"\b{_INSEE}\b")


def normaliser_id(valeur) -> str:
    """Forme comparable d'un ID erreur : espaces supprimés, majuscules.

    Utilisée des deux côtés — saisie utilisateur et cellules du classeur — pour
    que ``« 13001_1 »``, ``« 13001_1 »`` et ``« 2a004_3 »`` se rejoignent.
    """
    if valeur is None:
        return ""
    return "".join(str(valeur).split()).upper()


def extraire_ids(texte: str) -> list[str]:
    """Retourne les ID erreur valides du texte, sans doublons, dans l'ordre.

    Le texte est mis en majuscules au préalable (pour accepter ``2a004_3``),
    puis filtré par le motif ``<insee>_<numéro>``.
    """
    ids = [f"{insee}_{numero}" for insee, numero in _MOTIF_ID.findall(texte.upper())]
    # dict.fromkeys : supprime les doublons tout en préservant l'ordre
    return list(dict.fromkeys(ids))


def codes_sans_id(texte: str) -> list[str]:
    """Codes INSEE saisis **sans** numéro d'erreur, sans doublons.

    Ils ne sont pas traités : sans identifiant, aucune ligne ne peut être
    sélectionnée. Les remonter évite qu'une saisie incomplète passe inaperçue.
    """
    return list(dict.fromkeys(_MOTIF_INSEE_SEUL.findall(texte.upper())))


def insee_de(id_erreur: str) -> str:
    """Code commune porté par un ID erreur (``13001_7`` → ``13001``)."""
    return normaliser_id(id_erreur).split("_", 1)[0]


def grouper_par_commune(ids: Iterable[str]) -> dict[str, list[str]]:
    """Regroupe des ID erreur par code INSEE.

    Returns:
        ``{insee: [ID erreur, …]}``, communes et identifiants dans leur ordre
        d'apparition, sans doublons — soit exactement un fichier audit à ouvrir
        par entrée du dictionnaire.
    """
    groupes: dict[str, list[str]] = {}
    for identifiant in ids:
        cle = normaliser_id(identifiant)
        if not cle:
            continue
        liste = groupes.setdefault(insee_de(cle), [])
        if cle not in liste:
            liste.append(cle)
    return groupes


def groupes_depuis_texte(texte: str) -> dict[str, list[str]]:
    """Raccourci ``texte saisi → {insee: [ID erreur, …]}``."""
    return grouper_par_commune(extraire_ids(texte))
