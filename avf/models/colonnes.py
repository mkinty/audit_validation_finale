"""Colonnes contrôlées dans l'onglet « Audit », et valeurs qu'elles doivent contenir.

La colonne des ID erreur n'est pas portée ici : elle est partagée avec les
règles de mise à jour (``sna.models.regles.ReglesConfig.colonne_id``), pour que
l'audit et la mise à jour sélectionnent toujours les lignes de la même façon.

Critères d'une ligne **validée** (cf. ``services.controle``) :

    Validation finale  ∈ valeurs_validation   (défaut : OK)
    Audit              ∈ valeurs_audit        (défaut : OK, KO)
    Date audit         renseignée et lisible comme une date (même saisie en texte)
    Auditeur           renseigné

Les valeurs sont comparées sans tenir compte de la casse ni des espaces autour.
"""

from __future__ import annotations

from dataclasses import dataclass

VALEURS_VALIDATION_DEFAUT = ("OK",)
VALEURS_AUDIT_DEFAUT = ("OK", "KO")


@dataclass
class ColonnesAudit:
    """Libellés d'en-tête recherchés (casse et espaces ignorés à la lecture)."""

    validation: str = "Validation finale"
    audit: str = "Audit"
    id_erreur: str = "ID erreur"
    id_imm: str = "ID imm"     # identifiant des mises à jour : unique par ligne (ID erreur ne l'est pas)
    date_audit: str = "Date audit"
    auditeur: str = "Auditeur"
    valeurs_validation: tuple[str, ...] = VALEURS_VALIDATION_DEFAUT
    valeurs_audit: tuple[str, ...] = VALEURS_AUDIT_DEFAUT

    def controlees(self) -> dict[str, str]:
        """Les quatre colonnes contrôlées : ``{clé: libellé}``, dans l'ordre d'affichage."""
        return {
            "audit": self.audit, "date_audit": self.date_audit,
            "auditeur": self.auditeur, "validation": self.validation,
        }

    def to_dict(self) -> dict:
        return {
            "validation": self.validation,
            "audit": self.audit,
            "date_audit": self.date_audit,
            "auditeur": self.auditeur,
            "id_erreur": self.id_erreur,
            "id_imm": self.id_imm,
            "valeurs_validation": list(self.valeurs_validation),
            "valeurs_audit": list(self.valeurs_audit),
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
            id_erreur=lire("id_erreur"),
            id_imm=lire("id_imm"),
            valeurs_validation=parser_valeurs(data.get("valeurs_validation")) or VALEURS_VALIDATION_DEFAUT,
            valeurs_audit=parser_valeurs(data.get("valeurs_audit")) or VALEURS_AUDIT_DEFAUT,
        )


def parser_valeurs(valeur) -> tuple[str, ...]:
    """``"OK, KO"`` ou ``["OK", "KO"]`` → ``("OK", "KO")`` (vides et doublons retirés)."""
    if valeur is None:
        return ()
    morceaux = valeur if isinstance(valeur, (list, tuple)) else str(valeur).replace(";", ",").split(",")
    vus: dict[str, str] = {}
    for m in morceaux:
        texte = str(m).strip()
        if texte and texte.casefold() not in vus:
            vus[texte.casefold()] = texte
    return tuple(vus.values())
