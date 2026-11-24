"""Règles de modification paramétrables par l'utilisateur.

Le choix des lignes reste dicté par les **ID erreur** saisis dans la fenêtre
principale : une ligne n'est envisagée que si la valeur de sa colonne
« ID erreur » figure parmi les identifiants demandés pour la commune traitée.

Sur ces lignes, une RÈGLE — reprise du constructeur de règles du projet
``excel_rule_editor`` — se compose de :

    zéro, une ou plusieurs CONDITIONS combinées en ET   +   une ACTION

Les CONDITIONS sont un **filtre supplémentaire facultatif** : sans condition,
la règle s'applique à toutes les lignes des ID erreur ; avec conditions, elle
ne s'applique qu'à celles qui les vérifient toutes.

Chaque CONDITION compare une colonne (choisie dans la liste réelle du fichier
Excel) à une VALEUR littérale ou à une autre COLONNE, via un opérateur adapté
au type de la colonne.

L'ACTION écrit dans une colonne (celle de la 1re condition par défaut) une
VALEUR littérale, le contenu d'une autre COLONNE, ou vide la cellule.

``ReglesConfig`` porte en plus trois réglages transverses au classeur :

    ligne_entete     → ligne des libellés de colonnes (1 par défaut) ;
    colonne_id       → colonne contenant les ID erreur (« ID erreur » par défaut) ;
    harmoniser_style → reprise de la police / du format de la colonne écrite.

La configuration de départ (``REGLES_DEFAUT``, 16 règles d'audit) est écrite
dans ``config.json`` au premier lancement ; ensuite, ce que l'utilisateur a
configuré est conservé d'un démarrage à l'autre. Le bouton « Réinitialiser
(config par défaut) » du panneau y ramène à tout moment.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from enum import Enum
from datetime import date

# Types de colonnes déduits du libellé et des valeurs (cf. services.colonnes_excel).
# Ils informent l'utilisateur et adaptent la saisie : une colonne de dates
# propose une saisie au format JJ/MM/AAAA et des opérateurs de comparaison.
TYPE_NUM = "numérique"
TYPE_TEXTE = "texte"
TYPE_DATE = "date"

# Libellé de colonne recherché par défaut pour les identifiants d'erreur.
COLONNE_ID_DEFAUT = "ID erreur"


# ── Opérateurs ────────────────────────────────────────────────────
# Opérateurs ne prenant pas d'opérande de droite.
OPERATEURS_UNAIRES = ("vide", "non vide")

OPERATEURS_NUM = ["=", "≠", ">", "<", "≥", "≤", *OPERATEURS_UNAIRES]
OPERATEURS_DATE = ["=", "≠", ">", "<", "≥", "≤", *OPERATEURS_UNAIRES]
OPERATEURS_TEXTE = [
    "=", "≠", "contient", "ne contient pas", "commence par", "finit par",
    *OPERATEURS_UNAIRES,
]

def get_today_date():
    """Return today's date as a string in DD/MM/YYYY format."""
    return date.today().strftime("%d/%m/%Y")


def operateurs_pour_type(type_colonne: str) -> list[str]:
    """Opérateurs autorisés selon le type de la colonne de gauche."""
    if type_colonne == TYPE_NUM:
        return OPERATEURS_NUM
    if type_colonne == TYPE_DATE:
        return OPERATEURS_DATE
    return OPERATEURS_TEXTE


class TypeDroite(str, Enum):
    """Nature de l'opérande de droite d'une condition."""

    VALEUR = "valeur"
    COLONNE = "colonne"


class SourceValeur(str, Enum):
    """Nature de la valeur écrite par une action."""

    VALEUR = "valeur"    # une valeur saisie par l'utilisateur
    COLONNE = "colonne"  # recopie du contenu d'une autre colonne
    VIDE = "vide"        # efface la cellule


LIBELLE_SOURCE = {
    SourceValeur.VALEUR: "valeur",
    SourceValeur.COLONNE: "colonne",
    SourceValeur.VIDE: "effacer",
}
SOURCE_PAR_LIBELLE = {v: k for k, v in LIBELLE_SOURCE.items()}


@dataclass
class Condition:
    """Une condition élémentaire : ``<colonne> <opérateur> <valeur|colonne>``.

    ``colonne`` et ``droite`` acceptent indifféremment un nom d'en-tête
    (« Audit - Street View ») ou une lettre de colonne (« Z »).
    """

    colonne: str
    operateur: str = "="
    type_droite: TypeDroite = TypeDroite.VALEUR
    droite: str = ""

    def libelle(self) -> str:
        if self.operateur in OPERATEURS_UNAIRES:
            return f"{self.colonne} {self.operateur}"
        cible = self.droite if self.type_droite == TypeDroite.VALEUR else f"[{self.droite}]"
        return f"{self.colonne} {self.operateur} « {cible} »"


@dataclass
class Action:
    """Ce qui est écrit quand toutes les conditions sont vraies.

    ``colonne`` vide signifie « la colonne de la première condition », ce qui
    couvre le cas courant de la correction sur place. Une action ``VALEUR``
    sans valeur n'est pas appliquée : c'est l'état des règles préparées
    d'avance, en attente d'être renseignées. Vider une cellule se demande
    explicitement (``VIDE``), pour qu'un formulaire laissé blanc n'efface
    jamais une colonne par inadvertance.
    """

    colonne: str = ""
    source: SourceValeur = SourceValeur.VALEUR
    valeur: str = ""

    def libelle(self) -> str:
        cible = self.colonne or "(colonne testée)"
        if self.source == SourceValeur.VIDE:
            return f"effacer {cible}"
        quoi = f"[{self.valeur}]" if self.source == SourceValeur.COLONNE else f"« {self.valeur} »"
        return f"{cible} ← {quoi}"


@dataclass
class RegleModification:
    """Conditions en ET (facultatives) + action, sur les lignes retenues.

    Args:
        code: identifiant court affiché dans le journal (R1, R2, …).
        libelle: description libre saisie par l'utilisateur.
        action: modification à appliquer.
        actif: une règle inactive est ignorée sans être supprimée.
        conditions: filtre supplémentaire — toutes doivent être vraies (ET).
            Vide : la règle s'applique à toutes les lignes des ID erreur.
        exact: comparaison texte sensible à la casse et aux espaces.
            Décocher rend la comparaison tolérante (``strip`` + casse ignorée).
    """

    code: str
    libelle: str = ""
    action: Action = field(default_factory=Action)
    actif: bool = True
    conditions: list[Condition] = field(default_factory=list)
    exact: bool = True

    # ── Alias de lecture (compatibilité et confort) ───────────────
    @property
    def colonne(self) -> str:
        """Colonne désignée par l'action (« » si laissée au repli)."""
        return self.action.colonne

    @property
    def valeur(self) -> str:
        """Valeur (ou colonne source) portée par l'action."""
        return self.action.valeur

    @property
    def effacer(self) -> bool:
        """Vrai si l'action vide la cellule."""
        return self.action.source is SourceValeur.VIDE

    @property
    def colonne_cible(self) -> str:
        """Colonne réellement écrite (repli sur la 1re condition)."""
        if self.action.colonne:
            return self.action.colonne
        return self.conditions[0].colonne if self.conditions else ""

    @property
    def renseignee(self) -> bool:
        """Vrai si la règle a quelque chose à faire (écrire ou effacer)."""
        if self.action.source is SourceValeur.VIDE:
            return True
        return str(self.action.valeur).strip() != ""

    def resume(self) -> str:
        etat = "actif" if self.actif else "INACTIF"
        cible = self.colonne_cible or "(colonne non définie)"
        if self.action.source is SourceValeur.VIDE:
            quoi = "(effacer)"
        elif not self.renseignee:
            quoi = "(aucune valeur — règle en attente)"
        elif self.action.source is SourceValeur.COLONNE:
            quoi = f"[{self.action.valeur}]"
        else:
            quoi = f"« {self.action.valeur} »"
        affectation = f"{cible} ← {quoi}"
        if self.conditions:
            conds = " ET ".join(c.libelle() for c in self.conditions)
            return f"{self.code} ({etat}) : SI {conds} ALORS {affectation}"
        return f"{self.code} ({etat}) : {affectation}"


# ── Configuration livrée : 16 règles d'audit par défaut ───────────
# Cette liste EST la configuration de départ du programme : elle est écrite
# dans ``config.json`` au premier lancement, et le bouton « Réinitialiser
# (config par défaut) » du panneau la restaure à l'identique (mêmes colonnes,
# mêmes valeurs, mêmes opérateurs). Pour changer la configuration livrée, il
# suffit de modifier cette liste — rien d'autre dans le code n'est à toucher.
_ANL = "Analyse nombre de Logement"

REGLES_DEFAUT: list[RegleModification] = [
    RegleModification(
        code="R1", libelle="Audit",
        action=Action(colonne="Audit", source=SourceValeur.VALEUR, valeur="OK"),
    ),
    RegleModification(
        code="R2", libelle="Date audit",
        action=Action(colonne="Date audit", source=SourceValeur.VALEUR, valeur=get_today_date()),
    ),
    RegleModification(
        code="R3", libelle="Auditeur",
        action=Action(colonne="Auditeur", source=SourceValeur.VALEUR, valeur="Validation en masse"),
    ),
    RegleModification(
        code="R4", libelle="Validation finale",
        action=Action(colonne="Validation finale", source=SourceValeur.VALEUR, valeur="OK"),
    ),
]


@dataclass
class ReglesConfig:
    """Liste ordonnée de règles + repères de lecture du classeur.

    ``ligne_entete`` sert à détecter les noms de colonnes et à résoudre les
    références des règles ; ``colonne_id`` désigne la colonne des ID erreur.

    ``harmoniser_style`` donne aux cellules écrites la police, la taille et le
    format d'affichage des valeurs déjà présentes dans leur colonne, pour que
    les valeurs saisies ne détonnent pas dans le classeur.
    """

    regles: list[RegleModification] = field(default_factory=list)
    ligne_entete: int = 1
    colonne_id: str = COLONNE_ID_DEFAUT
    harmoniser_style: bool = True

    # ── Configuration livrée ──────────────────────────────────────
    @staticmethod
    def defaut() -> "ReglesConfig":
        """Configuration de départ : les 16 règles d'audit de ``REGLES_DEFAUT``.

        Écrite dans ``config.json`` au premier lancement, puis modifiable
        comme n'importe quelle configuration ; les modifications de
        l'utilisateur sont conservées d'un démarrage à l'autre. Le bouton
        « Réinitialiser (config par défaut) » du panneau restaure cette
        liste à l'identique.

        Une copie profonde est renvoyée : l'appelant peut modifier le résultat
        sans altérer le modèle partagé.
        """
        return ReglesConfig(
            regles=deepcopy(REGLES_DEFAUT),
            ligne_entete=1,
            colonne_id=COLONNE_ID_DEFAUT,
            harmoniser_style=True,
        )

    # ── Accès ─────────────────────────────────────────────────────
    @property
    def actives(self) -> list[RegleModification]:
        """Règles cochées, visant une colonne, et ayant quelque chose à écrire."""
        return [
            r for r in self.regles
            if r.actif and r.colonne_cible.strip() and r.renseignee
        ]

    @property
    def en_attente(self) -> list[RegleModification]:
        """Règles cochées mais sans valeur : préparées, pas encore utilisables."""
        return [
            r for r in self.regles
            if r.actif and r.colonne_cible.strip() and not r.renseignee
        ]

    def code_unique(self, prefixe: str = "R") -> str:
        """Génère un code de règle non utilisé (R2, R3, …)."""
        existants = {r.code for r in self.regles}
        i = 1
        while f"{prefixe}{i}" in existants:
            i += 1
        return f"{prefixe}{i}"

    def resume(self) -> list[str]:
        """Résumé lisible, une ligne par règle (journal, confirmation)."""
        return [r.resume() for r in self.regles]

    # ── Sérialisation JSON ────────────────────────────────────────
    def to_dict(self) -> dict:
        return {
            "ligne_entete": self.ligne_entete,
            "colonne_id": self.colonne_id,
            "harmoniser_style": self.harmoniser_style,
            "regles": [_regle_to_dict(r) for r in self.regles],
        }

    @staticmethod
    def from_dict(data: dict) -> "ReglesConfig":
        return ReglesConfig(
            regles=[_regle_from_dict(d) for d in data.get("regles", [])],
            ligne_entete=int(data.get("ligne_entete", 1) or 1),
            colonne_id=str(data.get("colonne_id") or COLONNE_ID_DEFAUT),
            harmoniser_style=bool(data.get("harmoniser_style", True)),
        )


# ── Sérialisation d'une règle ─────────────────────────────────────
def _regle_to_dict(regle: RegleModification) -> dict:
    return {
        "code": regle.code,
        "libelle": regle.libelle,
        "actif": regle.actif,
        "exact": regle.exact,
        "action": {
            "colonne": regle.action.colonne,
            "source": regle.action.source.value,
            "valeur": regle.action.valeur,
        },
        "conditions": [
            {
                "colonne": c.colonne,
                "operateur": c.operateur,
                "type_droite": c.type_droite.value,
                "droite": c.droite,
            }
            for c in regle.conditions
        ],
    }


def _regle_from_dict(d: dict) -> RegleModification:
    """Reconstruit une règle, y compris depuis un ``config.json`` antérieur.

    Trois formes sont acceptées :

    - la forme actuelle : ``action`` + ``conditions`` ;
    - la forme « à plat » d'une version intermédiaire : ``colonne`` /
      ``valeur`` / ``effacer`` — transposée en une action sans condition ;
    - la toute première forme (``action`` + ``conditions`` + ``ligne_debut``) :
      relue telle quelle, ``ligne_debut`` étant simplement ignoré puisque les
      lignes sont désormais choisies par les ID erreur.
    """
    code = d["code"]
    libelle = d.get("libelle", code)
    actif = bool(d.get("actif", True))
    exact = bool(d.get("exact", True))

    action_brute = d.get("action")
    if isinstance(action_brute, dict):
        action = Action(
            colonne=str(action_brute.get("colonne", "") or ""),
            source=_source(action_brute.get("source")),
            valeur=str(action_brute.get("valeur", "") or ""),
        )
        conditions = [
            _condition_from_dict(c) for c in (d.get("conditions") or [])
        ]
    else:
        colonne = str(d.get("colonne", "") or "")
        valeur = str(d.get("valeur", "") or "")
        source = SourceValeur.VIDE if d.get("effacer") else SourceValeur.VALEUR
        action = Action(colonne=colonne, source=source, valeur=valeur)
        conditions = []

    return RegleModification(
        code=code, libelle=libelle, action=action, actif=actif,
        conditions=conditions, exact=exact,
    )


def _condition_from_dict(c: dict) -> Condition:
    return Condition(
        colonne=str(c.get("colonne", "") or ""),
        operateur=str(c.get("operateur", "=") or "="),
        type_droite=_type_droite(c.get("type_droite")),
        droite=str(c.get("droite", "") or ""),
    )


def _source(valeur) -> SourceValeur:
    try:
        return SourceValeur(valeur)
    except ValueError:
        return SourceValeur.VALEUR


def _type_droite(valeur) -> TypeDroite:
    try:
        return TypeDroite(valeur)
    except ValueError:
        return TypeDroite.VALEUR
