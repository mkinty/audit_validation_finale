"""Évaluation des conditions d'une règle et de la valeur de son action.

Le moteur ne connaît ni openpyxl ni l'interface : il reçoit un accesseur
``lire(reference) -> valeur`` et raisonne uniquement sur des valeurs Python.
Il est donc testable sans fichier Excel.

Sémantique de comparaison (reprise du constructeur de règles du projet
``excel_rule_editor``) :
    - si les deux opérandes sont des nombres → comparaison numérique ;
    - sinon, s'ils s'interprètent tous deux comme des dates → comparaison de
      dates ;
    - sinon → comparaison texte. En mode ``exact`` la casse et les espaces
      comptent ; sinon les deux côtés sont nettoyés (``strip`` + casse ignorée) ;
    - les inégalités (> < ≥ ≤) sur du texte pur renvoient Faux.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from ..models.regles import (
    Action, Condition, OPERATEURS_UNAIRES, RegleModification, SourceValeur,
    TypeDroite,
)
from .colonnes_excel import to_num
from .valeurs import analyser_date, valeur_a_ecrire

# Accesseur : reçoit une référence de colonne, retourne la valeur de la cellule.
Lecteur = Callable[[str], Any]


def regle_sapplique(regle: RegleModification, lire: Lecteur) -> bool:
    """Vrai si TOUTES les conditions de la règle sont vraies (ET).

    Une règle **sans condition** s'applique à toutes les lignes retenues : dans
    ce programme, les lignes sont d'abord choisies par les ID erreur, les
    conditions n'étant qu'un filtre supplémentaire facultatif.
    """
    return all(condition_vraie(c, lire, regle.exact) for c in regle.conditions)


def condition_vraie(condition: Condition, lire: Lecteur, exact: bool = True) -> bool:
    """Évalue une condition élémentaire sur une ligne."""
    gauche = lire(condition.colonne)

    if condition.operateur in OPERATEURS_UNAIRES:
        vide = gauche is None or str(gauche).strip() == ""
        return vide if condition.operateur == "vide" else not vide

    if condition.type_droite == TypeDroite.VALEUR:
        droite = condition.droite
    else:
        droite = lire(condition.droite)

    return comparer(gauche, droite, condition.operateur, exact)


def valeur_pour_action(action: Action, lire: Lecteur, en_date: bool = False) -> Any:
    """Valeur que l'action doit inscrire dans la cellule cible.

    ``en_date`` s'applique à une valeur littérale visant une colonne de dates :
    la saisie est alors interprétée comme une date. La recopie d'une colonne
    reprend telle quelle la valeur lue.
    """
    if action.source is SourceValeur.VIDE:
        return None
    if action.source is SourceValeur.COLONNE:
        return lire(action.valeur)
    return valeur_a_ecrire(action.valeur, en_date=en_date)


def comparer(gauche: Any, droite: Any, operateur: str, exact: bool = True) -> bool:
    """Compare deux valeurs selon l'opérateur demandé."""
    gn, dn = to_num(gauche), to_num(droite)
    if gn is not None and dn is not None:
        return {
            "=": gn == dn, "≠": gn != dn, ">": gn > dn,
            "<": gn < dn, "≥": gn >= dn, "≤": gn <= dn,
        }.get(operateur, False)

    gd, dd = analyser_date(gauche), analyser_date(droite)
    if gd is not None and dd is not None:
        return {
            "=": gd == dd, "≠": gd != dd, ">": gd > dd,
            "<": gd < dd, "≥": gd >= dd, "≤": gd <= dd,
        }.get(operateur, False)

    g = "" if gauche is None else str(gauche)
    d = "" if droite is None else str(droite)
    if not exact:
        g, d = g.strip().casefold(), d.strip().casefold()

    if operateur == "=":
        return g == d
    if operateur == "≠":
        return g != d
    if operateur == "contient":
        return d in g
    if operateur == "ne contient pas":
        return d not in g
    if operateur == "commence par":
        return g.startswith(d)
    if operateur == "finit par":
        return g.endswith(d)
    return False  # inégalités non pertinentes sur du texte
