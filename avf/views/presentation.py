"""Formats d'affichage partagés par le rapport HTML et la fenêtre du rapport.

Les deux vues affichent les mêmes chiffres de la même façon : nombres groupés
par milliers, pourcentages à une décimale, libellés de statut avec icône (la
couleur ne porte jamais seule l'information).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass

from ..models.audit import AuditCommune, RapportAudit, StatutAudit

# Espace fine insécable : séparateur de milliers et avant « % » (typographie française).
ESPACE_FINE = " "

SEPARATEURS_COPIE = {
    "ligne": "\n",
    "virgule": ", ",
    "point-virgule": ";",
    "espace": " ",
}


def nombre(valeur: int) -> str:
    """``1284`` → ``« 1 284 »``."""
    return f"{valeur:,}".replace(",", ESPACE_FINE)


def pourcent(taux: float | None) -> str:
    """``0.875`` → ``« 87,5 % »`` ; ``None`` → ``« — »``."""
    if taux is None:
        return "—"
    valeur = round(taux * 100, 1)
    texte = f"{valeur:.0f}" if valeur.is_integer() else f"{valeur:.1f}".replace(".", ",")
    return f"{texte}{ESPACE_FINE}%"


def duree(secondes: float) -> str:
    if secondes < 60:
        return f"{secondes:.1f} s".replace(".", ",")
    minutes, reste = divmod(int(round(secondes)), 60)
    return f"{minutes} min {reste:02d} s"


@dataclass(frozen=True)
class Etat:
    """État d'une commune dans le rapport."""

    cle: str        # complete | a-completer | echec | vide
    icone: str
    libelle: str


ETAT_COMPLETE = Etat("complete", "✅", "Complète")
ETAT_A_COMPLETER = Etat("a-completer", "⏳", "À compléter")
ETAT_SANS_ID = Etat("vide", "∅", "Aucun ID erreur")
ETAT_INTROUVABLE = Etat("echec", "⚠️", "Fichier introuvable")
ETAT_ERREUR = Etat("echec", "❌", "Erreur de lecture")


def etat(commune: AuditCommune) -> Etat:
    if commune.statut is StatutAudit.INTROUVABLE:
        return ETAT_INTROUVABLE
    if commune.statut is StatutAudit.ERREUR:
        return ETAT_ERREUR
    if not commune.nb_ids:
        return ETAT_SANS_ID
    return ETAT_COMPLETE if commune.nb_vides == 0 else ETAT_A_COMPLETER


def texte_ids(ids: Iterable[str], separateur: str = "\n") -> str:
    """Liste d'ID prête à coller (un par ligne par défaut)."""
    return separateur.join(ids)


def resume(rapport: RapportAudit) -> list[str]:
    """Bilan en quelques lignes (journal, en-tête de la fenêtre du rapport)."""
    lignes = [
        f"{nombre(rapport.nb_communes)} commune(s) auditée(s) — "
        f"{nombre(len(rapport.communes_completes))} complète(s), "
        f"{nombre(len(rapport.communes_a_traiter))} à compléter, "
        f"{nombre(len(rapport.en_echec))} en échec",
        f"{nombre(rapport.nb_ids)} ID erreur — "
        f"{nombre(rapport.nb_renseignes)} avec Validation finale ({pourcent(rapport.taux)}), "
        f"{nombre(rapport.nb_vides)} sans",
    ]
    return lignes
