"""Structures de base du domaine (statut d'un traitement, résultat par commune).

Ces objets ne dépendent d'aucune bibliothèque externe : ils décrivent *ce que*
le programme manipule, pas *comment* il le fait.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class Statut(Enum):
    """Issue du traitement d'une commune."""

    SUCCES = "succes"            # Fichier trouvé et traité
    INTROUVABLE = "introuvable"  # Aucun fichier audit pour cette commune
    ERREUR = "erreur"            # Une exception est survenue pendant le traitement


@dataclass
class ResultatCommune:
    """Résultat du traitement d'une seule commune.

    ``detail`` associe le code de chaque règle au nombre de cellules qu'elle a
    modifiées, ce qui permet d'afficher un bilan règle par règle.

    ``ids_absents`` liste les ID erreur demandés pour cette commune mais
    introuvables dans le fichier : sans cette remontée, une saisie erronée
    passerait pour un traitement réussi.

    ``ids_non_verifies`` (écriture via Excel uniquement, cf.
    ``services.excel_com``) liste les ID dont la cellule a été écrite
    localement mais dont la relecture après enregistrement ne confirme pas la
    valeur — l'enregistrement vers SharePoint n'a pas pu être garanti.
    """

    commune: str
    statut: Statut
    nb_modifications: int = 0
    message: str = ""
    detail: dict[str, int] = field(default_factory=dict)
    nb_ids_demandes: int = 0
    nb_lignes: int = 0
    ids_absents: list[str] = field(default_factory=list)
    ids_non_verifies: list[str] = field(default_factory=list)
    duree: float = 0.0

    def resume_detail(self) -> str:
        """Ex. « R1: 12, R2: 3 » (règles sans effet omises)."""
        return ", ".join(f"{code}: {n}" for code, n in self.detail.items() if n)

    def resume_duree(self) -> str:
        """Durée du traitement, arrondie à la milliseconde près (« 2,4 s »)."""
        return f"{self.duree:.1f} s".replace(".", ",")

    def resume_ids(self) -> str:
        """Ex. « 4/5 ID erreur trouvés » (identifiants manquants signalés)."""
        trouves = self.nb_ids_demandes - len(self.ids_absents)
        return f"{trouves}/{self.nb_ids_demandes} ID erreur trouvés"
