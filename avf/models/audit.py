"""Résultats de l'audit : ligne, ID erreur, commune, rapport complet.

Règle métier : chaque **ligne** portant un ID erreur est contrôlée sur ses
quatre colonnes (Audit, Date audit, Auditeur, Validation finale — critères
dans ``models.colonnes`` et ``services.controle``). Une ligne qui les respecte
toutes est **validée** ; sinon elle est **non valide**, avec le motif de chaque
critère non respecté.

Un ID erreur pouvant apparaître sur plusieurs lignes, il est **validé** si
toutes ses lignes le sont, **non valide** dès qu'une de ses lignes ne l'est
pas : la mise à jour réécrivant toutes les lignes d'un ID, c'est la seule
définition qui ne laisse jamais de ligne non valide derrière elle.

Ces objets ne dépendent d'aucune bibliothèque externe ni de l'interface.
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum

from .colonnes import ColonnesAudit
from .source import SourceAudit

# Critères contrôlés, dans l'ordre d'affichage.
CRITERES = ("audit", "date_audit", "auditeur", "validation")

# Ce que l'on affiche ou copie : les lignes non valides, les validées, ou toutes.
NON_VALIDES, VALIDEES, TOUTES = "non_valides", "validees", "toutes"
CATEGORIES = (NON_VALIDES, VALIDEES, TOUTES)


class StatutAudit(Enum):
    """Issue de la lecture d'un fichier audit."""

    OK = "ok"                    # Fichier lu, colonnes trouvées
    INTROUVABLE = "introuvable"  # Aucun fichier audit_*.xlsx pour la commune
    ERREUR = "erreur"            # Onglet / colonne absent, fichier illisible…


def est_vide(valeur) -> bool:
    """Vrai si une cellule est considérée comme non renseignée (espaces seuls = vide)."""
    return valeur is None or str(valeur).strip() == ""


@dataclass(frozen=True)
class Anomalie:
    """Un critère non respecté par une ligne."""

    critere: str   # l'un de CRITERES
    message: str   # ex. « Audit « Fait » ≠ OK / KO »


@dataclass
class LigneAudit:
    """Une ligne de l'onglet « Audit » portant un ID erreur."""

    ligne: int
    id_erreur: str
    validation: str = ""
    audit: str = ""
    date_audit: str = ""
    auditeur: str = ""
    anomalies: list[Anomalie] = field(default_factory=list)

    @property
    def valide(self) -> bool:
        return not self.anomalies

    @property
    def motifs(self) -> str:
        """Motifs de non-validité, lisibles (vide pour une ligne validée)."""
        return " · ".join(a.message for a in self.anomalies)

    def dans(self, categorie: str) -> bool:
        return categorie == TOUTES or (self.valide if categorie == VALIDEES else not self.valide)


@dataclass
class IdErreurAudit:
    """Un ID erreur et la ou les lignes qui le portent."""

    id_erreur: str
    lignes: list[LigneAudit] = field(default_factory=list)

    @property
    def valide(self) -> bool:
        """Vrai si toutes les lignes de l'ID sont validées."""
        return bool(self.lignes) and all(l.valide for l in self.lignes)

    @property
    def doublon(self) -> bool:
        return len(self.lignes) > 1

    @property
    def lignes_non_valides(self) -> list[LigneAudit]:
        return [l for l in self.lignes if not l.valide]

    @property
    def reference(self) -> LigneAudit:
        """Ligne mise en avant : la première non valide, à défaut la première."""
        return next((l for l in self.lignes if not l.valide), self.lignes[0])

    def dans(self, categorie: str) -> bool:
        return categorie == TOUTES or (self.valide if categorie == VALIDEES else not self.valide)


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

    # ── Lignes ────────────────────────────────────────────────────
    @property
    def lignes(self) -> list[LigneAudit]:
        """Toutes les lignes portant un ID erreur, dans l'ordre du fichier."""
        return sorted((l for i in self.ids for l in i.lignes), key=lambda l: l.ligne)

    def lignes_de(self, categorie: str = NON_VALIDES) -> list[LigneAudit]:
        return [l for l in self.lignes if l.dans(categorie)]

    @property
    def nb_lignes(self) -> int:
        return sum(len(i.lignes) for i in self.ids)

    @property
    def nb_lignes_valides(self) -> int:
        return sum(1 for i in self.ids for l in i.lignes if l.valide)

    @property
    def nb_lignes_non_valides(self) -> int:
        return self.nb_lignes - self.nb_lignes_valides

    # ── ID erreur ─────────────────────────────────────────────────
    @property
    def nb_ids(self) -> int:
        return len(self.ids)

    @property
    def nb_ids_valides(self) -> int:
        return sum(1 for i in self.ids if i.valide)

    @property
    def nb_ids_non_valides(self) -> int:
        return self.nb_ids - self.nb_ids_valides

    @property
    def nb_doublons(self) -> int:
        return sum(1 for i in self.ids if i.doublon)

    def ids_de(self, categorie: str = NON_VALIDES) -> list[str]:
        """ID erreur de la catégorie (``NON_VALIDES`` : au moins une ligne non valide)."""
        return [i.id_erreur for i in self.ids if i.dans(categorie)] if self.lisible else []

    @property
    def ids_non_valides(self) -> list[str]:
        return self.ids_de(NON_VALIDES)

    # ── Synthèse ──────────────────────────────────────────────────
    @property
    def taux(self) -> float | None:
        """Part des lignes validées (0 → 1), ``None`` sans aucune ligne."""
        return self.nb_lignes_valides / self.nb_lignes if self.nb_lignes else None

    def par_critere(self) -> dict[str, int]:
        """Nombre de lignes ne respectant pas chaque critère."""
        comptes = dict.fromkeys(CRITERES, 0)
        for i in self.ids:
            for l in i.lignes:
                for critere in {a.critere for a in l.anomalies}:
                    comptes[critere] = comptes.get(critere, 0) + 1
        return comptes

    @property
    def lisible(self) -> bool:
        return self.statut is StatutAudit.OK

    @property
    def complete(self) -> bool:
        """Fichier lu et toutes ses lignes validées."""
        return self.lisible and self.nb_lignes_non_valides == 0


@dataclass
class RapportAudit:
    """Rapport complet d'un audit, toutes communes confondues."""

    racine: str
    communes: list[AuditCommune] = field(default_factory=list)
    colonnes: ColonnesAudit = field(default_factory=ColonnesAudit)
    colonne_id: str = "ID erreur"
    date: datetime = field(default_factory=datetime.now)
    duree: float = 0.0
    source: SourceAudit | None = None   # None : dossier local (rapports antérieurs, tests)

    @property
    def libelle_source(self) -> str:
        return self.source.libelle if self.source else "Dossier local"

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
    def nb_lignes(self) -> int:
        return sum(c.nb_lignes for c in self.communes)

    @property
    def nb_lignes_valides(self) -> int:
        return sum(c.nb_lignes_valides for c in self.communes)

    @property
    def nb_lignes_non_valides(self) -> int:
        return sum(c.nb_lignes_non_valides for c in self.communes)

    @property
    def nb_ids(self) -> int:
        return sum(c.nb_ids for c in self.communes)

    @property
    def nb_ids_valides(self) -> int:
        return sum(c.nb_ids_valides for c in self.communes)

    @property
    def nb_ids_non_valides(self) -> int:
        return sum(c.nb_ids_non_valides for c in self.communes)

    @property
    def taux(self) -> float | None:
        """Part des lignes validées."""
        return self.nb_lignes_valides / self.nb_lignes if self.nb_lignes else None

    @property
    def communes_a_corriger(self) -> list[AuditCommune]:
        """Communes lues ayant au moins une ligne non valide."""
        return [c for c in self.lues if c.nb_lignes_non_valides]

    @property
    def communes_completes(self) -> list[AuditCommune]:
        return [c for c in self.lues if c.complete]

    def ids_non_valides_par_commune(self) -> dict[str, list[str]]:
        """``{commune: [ID erreur non valides]}`` — entrée de la mise à jour."""
        return {c.commune: c.ids_non_valides for c in self.communes if c.ids_non_valides}

    def ids_de(self, categorie: str = NON_VALIDES) -> list[str]:
        return [i for c in self.communes for i in c.ids_de(categorie)]

    def par_critere(self) -> dict[str, int]:
        """Lignes non valides par critère non respecté (une ligne peut en cumuler plusieurs)."""
        total = dict.fromkeys(CRITERES, 0)
        for c in self.communes:
            for critere, n in c.par_critere().items():
                total[critere] = total.get(critere, 0) + n
        return total

    def par_departement(self) -> "OrderedDict[str, dict[str, int]]":
        """Synthèse ``{département: {communes, ids, lignes, valides, non_valides, erreurs}}`` (lignes)."""
        synthese: OrderedDict[str, dict[str, int]] = OrderedDict()
        for c in sorted(self.communes, key=lambda x: (x.departement, x.commune)):
            ligne = synthese.setdefault(
                c.departement,
                {"communes": 0, "ids": 0, "lignes": 0, "valides": 0, "non_valides": 0, "erreurs": 0},
            )
            ligne["communes"] += 1
            ligne["ids"] += c.nb_ids
            ligne["lignes"] += c.nb_lignes
            ligne["valides"] += c.nb_lignes_valides
            ligne["non_valides"] += c.nb_lignes_non_valides
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
