"""Résultats de l'audit : ligne, ID erreur, commune, rapport complet.

Règle métier : un **ID erreur** est « renseigné » quand sa cellule
« Validation finale » contient une valeur (espaces seuls = vide). Si un même
ID apparaît sur plusieurs lignes, il n'est renseigné que si **toutes** ses
lignes le sont — la mise à jour ciblant toutes les lignes d'un ID, c'est la
seule définition qui ne laisse jamais de ligne vide derrière elle.

Ces objets ne dépendent d'aucune bibliothèque externe ni de l'interface.
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum

from .colonnes import ColonnesAudit


class StatutAudit(Enum):
    """Issue de la lecture d'un fichier audit."""

    OK = "ok"                    # Fichier lu, colonnes trouvées
    INTROUVABLE = "introuvable"  # Aucun fichier audit_*.xlsx pour la commune
    ERREUR = "erreur"            # Onglet / colonne absent, fichier illisible…


def est_vide(valeur) -> bool:
    """Vrai si une cellule est considérée comme non renseignée."""
    return valeur is None or str(valeur).strip() == ""


@dataclass
class LigneAudit:
    """Une ligne de l'onglet « Audit » portant un ID erreur."""

    ligne: int
    id_erreur: str
    validation: str = ""
    audit: str = ""
    date_audit: str = ""
    auditeur: str = ""

    @property
    def validee(self) -> bool:
        return not est_vide(self.validation)


@dataclass
class IdErreurAudit:
    """Un ID erreur et la ou les lignes qui le portent."""

    id_erreur: str
    lignes: list[LigneAudit] = field(default_factory=list)

    @property
    def renseigne(self) -> bool:
        """Vrai si toutes les lignes de l'ID ont une « Validation finale »."""
        return bool(self.lignes) and all(l.validee for l in self.lignes)

    @property
    def doublon(self) -> bool:
        return len(self.lignes) > 1

    @property
    def reference(self) -> LigneAudit:
        """Ligne affichée : la première non validée, à défaut la première."""
        return next((l for l in self.lignes if not l.validee), self.lignes[0])


@dataclass
class AuditCommune:
    """Résultat de l'audit d'une commune (un fichier audit)."""

    commune: str
    departement: str
    statut: StatutAudit = StatutAudit.OK
    chemin: str = ""
    message: str = ""
    ids: list[IdErreurAudit] = field(default_factory=list)
    avertissements: list[str] = field(default_factory=list)
    nb_lignes_lues: int = 0
    duree: float = 0.0

    # ── Comptages ─────────────────────────────────────────────────
    @property
    def nb_ids(self) -> int:
        return len(self.ids)

    @property
    def nb_renseignes(self) -> int:
        return sum(1 for i in self.ids if i.renseigne)

    @property
    def nb_vides(self) -> int:
        return self.nb_ids - self.nb_renseignes

    @property
    def nb_lignes(self) -> int:
        return sum(len(i.lignes) for i in self.ids)

    @property
    def nb_doublons(self) -> int:
        return sum(1 for i in self.ids if i.doublon)

    @property
    def taux(self) -> float | None:
        """Part des ID renseignés (0 → 1), ``None`` sans aucun ID."""
        return self.nb_renseignes / self.nb_ids if self.nb_ids else None

    @property
    def ids_vides(self) -> list[str]:
        return [i.id_erreur for i in self.ids if not i.renseigne]

    @property
    def details_vides(self) -> list[IdErreurAudit]:
        return [i for i in self.ids if not i.renseigne]

    @property
    def lisible(self) -> bool:
        return self.statut is StatutAudit.OK

    @property
    def complete(self) -> bool:
        """Fichier lu et aucun ID sans validation finale."""
        return self.lisible and self.nb_vides == 0


@dataclass
class RapportAudit:
    """Rapport complet d'un audit, toutes communes confondues."""

    racine: str
    communes: list[AuditCommune] = field(default_factory=list)
    colonnes: ColonnesAudit = field(default_factory=ColonnesAudit)
    colonne_id: str = "ID erreur"
    date: datetime = field(default_factory=datetime.now)
    duree: float = 0.0

    # ── Totaux ────────────────────────────────────────────────────
    @property
    def lues(self) -> list[AuditCommune]:
        return [c for c in self.communes if c.lisible]

    @property
    def en_echec(self) -> list[AuditCommune]:
        return [c for c in self.communes if not c.lisible]

    @property
    def nb_communes(self) -> int:
        return len(self.communes)

    @property
    def nb_ids(self) -> int:
        return sum(c.nb_ids for c in self.communes)

    @property
    def nb_renseignes(self) -> int:
        return sum(c.nb_renseignes for c in self.communes)

    @property
    def nb_vides(self) -> int:
        return sum(c.nb_vides for c in self.communes)

    @property
    def taux(self) -> float | None:
        return self.nb_renseignes / self.nb_ids if self.nb_ids else None

    @property
    def communes_a_traiter(self) -> list[AuditCommune]:
        """Communes lues ayant au moins un ID sans validation finale."""
        return [c for c in self.lues if c.nb_vides]

    @property
    def communes_completes(self) -> list[AuditCommune]:
        return [c for c in self.lues if c.complete]

    def ids_vides_par_commune(self) -> dict[str, list[str]]:
        """``{commune: [ID erreur sans validation finale]}`` — entrée de la mise à jour."""
        return {c.commune: c.ids_vides for c in self.communes if c.lisible and c.nb_vides}

    def tous_ids_vides(self) -> list[str]:
        return [i for c in self.communes for i in (c.ids_vides if c.lisible else [])]

    def par_departement(self) -> "OrderedDict[str, dict[str, int]]":
        """Synthèse ``{département: {communes, ids, renseignes, vides, erreurs}}``."""
        synthese: OrderedDict[str, dict[str, int]] = OrderedDict()
        for c in sorted(self.communes, key=lambda x: (x.departement, x.commune)):
            ligne = synthese.setdefault(
                c.departement,
                {"communes": 0, "ids": 0, "renseignes": 0, "vides": 0, "erreurs": 0},
            )
            ligne["communes"] += 1
            ligne["ids"] += c.nb_ids
            ligne["renseignes"] += c.nb_renseignes
            ligne["vides"] += c.nb_vides
            ligne["erreurs"] += 0 if c.lisible else 1
        return synthese

    def commune(self, code: str) -> AuditCommune | None:
        return next((c for c in self.communes if c.commune == code), None)

    def remplacer(self, nouveaux: list[AuditCommune]) -> None:
        """Remplace (ou ajoute) les communes ré-auditées, en gardant l'ordre."""
        index = {c.commune: c for c in nouveaux}
        vus = set()
        resultat = []
        for c in self.communes:
            if c.commune in index:
                resultat.append(index[c.commune])
                vus.add(c.commune)
            else:
                resultat.append(c)
        resultat.extend(c for c in nouveaux if c.commune not in vus)
        self.communes = resultat
        self.date = datetime.now()
