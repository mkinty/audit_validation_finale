"""Modèles de données du domaine.

    sna.models.base    → Statut, ResultatCommune
    sna.models.regles  → RegleModification, ReglesConfig, Condition, Action

Tout est ré-exporté ici pour que ``from sna.models import ...`` reste possible.
"""

from .base import ResultatCommune, Statut
from .regles import (
    COLONNE_ID_DEFAUT,
    REGLES_DEFAUT,
    Action,
    Condition,
    RegleModification,
    ReglesConfig,
    SourceValeur,
    TYPE_DATE,
    TYPE_NUM,
    TYPE_TEXTE,
    TypeDroite,
    operateurs_pour_type,
)

__all__ = [
    "COLONNE_ID_DEFAUT",
    "REGLES_DEFAUT",
    "Action",
    "Condition",
    "RegleModification",
    "ReglesConfig",
    "ResultatCommune",
    "SourceValeur",
    "Statut",
    "TYPE_DATE",
    "TYPE_NUM",
    "TYPE_TEXTE",
    "TypeDroite",
    "operateurs_pour_type",
]
