"""Noms des colonnes lues dans l'onglet « Audit » de chaque fichier.

La colonne des ID erreur n'est pas portée ici : elle est partagée avec les
règles de mise à jour (``sna.models.regles.ReglesConfig.colonne_id``), pour que
l'audit et la mise à jour sélectionnent toujours les lignes de la même façon.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ColonnesAudit:
    """Libellés d'en-tête recherchés (casse et espaces ignorés à la lecture)."""

    validation: str = "Validation finale"
    audit: str = "Audit"
    date_audit: str = "Date audit"
    auditeur: str = "Auditeur"

    def complementaires(self) -> dict[str, str]:
        """Colonnes affichées à titre d'information : ``{clé: libellé}``."""
        return {"audit": self.audit, "date_audit": self.date_audit, "auditeur": self.auditeur}

    def to_dict(self) -> dict:
        return {
            "validation": self.validation,
            "audit": self.audit,
            "date_audit": self.date_audit,
            "auditeur": self.auditeur,
        }

    @staticmethod
    def from_dict(data: dict | None) -> "ColonnesAudit":
        defaut = ColonnesAudit()
        data = data or {}

        def lire(cle: str) -> str:
            return str(data.get(cle) or getattr(defaut, cle)).strip() or getattr(defaut, cle)

        return ColonnesAudit(
            validation=lire("validation"),
            audit=lire("audit"),
            date_audit=lire("date_audit"),
            auditeur=lire("auditeur"),
        )
