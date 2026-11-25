"""Lecture des colonnes auditées dans l'onglet « Audit » d'un fichier.

Seules cinq colonnes sont utiles (ID erreur, Validation finale, Audit, Date
audit, Auditeur) : plutôt que de charger tout le classeur avec openpyxl, on
réutilise la lecture en flux du XML du moteur ``sna``
(``sna.services.lecture_rapide``), qui ne décode que la colonne demandée. Une
structure de classeur inattendue fait basculer sur openpyxl en lecture seule —
même résultat, plus lent.

Comme pour la mise à jour, c'est toujours l'onglet nommé « Audit » qui est lu
(casse et espaces ignorés), jamais l'onglet actif.

Les cinq colonnes sont **obligatoires** : les quatre colonnes contrôlées
(Audit, Date audit, Auditeur, Validation finale) font chacune partie des
critères de validité d'une ligne. Une colonne absente fait échouer la lecture
(:class:`ColonneAbsente`, qui les cite toutes) plutôt que de déclarer à tort
toutes les lignes non valides.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
import re
import warnings

from openpyxl import load_workbook

from sna.services.colonnes_excel import (
    choisir_feuille, entete_est_date, entetes_depuis_feuille, resoudre_colonne,
    table_depuis_entetes, to_num,
)
from sna.services.lecture_rapide import LectureImpossible, lire_colonne, lire_entetes

from ..models.audit import LigneAudit
from ..models.colonnes import ColonnesAudit

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_EPOQUE_EXCEL = datetime(1899, 12, 30)
_LETTRE = re.compile(r"[A-Z]{1,3}")

# Bornes raisonnables d'un numéro de série Excel interprété comme une date
# (1954 → 2119) : en deçà ou au-delà, la valeur reste affichée telle quelle.
_SERIE_MIN, _SERIE_MAX = 20_000, 80_000


class ColonneAbsente(LookupError):
    """Une colonne obligatoire (ID erreur ou colonne contrôlée) est absente."""


@dataclass
class LectureAudit:
    """Contenu utile d'un fichier audit."""

    lignes: list[LigneAudit] = field(default_factory=list)
    avertissements: list[str] = field(default_factory=list)
    nb_lignes_lues: int = 0
    mode: str = "rapide"   # « rapide » (XML en flux) ou « openpyxl » (repli)


def lire_audit(
    chemin: str, colonne_id: str, colonnes: ColonnesAudit, ligne_entete: int = 1,
) -> LectureAudit:
    """Lit les lignes portant un ID erreur et leurs colonnes d'audit.

    Raises:
        ColonneAbsente: colonne des ID erreur ou colonne contrôlée absente.
        sna.services.colonnes_excel.OngletIntrouvable: aucun onglet « Audit ».
    """
    try:
        return _lire_rapide(chemin, colonne_id, colonnes, ligne_entete)
    except LectureImpossible:
        lecture = _lire_openpyxl(chemin, colonne_id, colonnes, ligne_entete)
        lecture.mode = "openpyxl"
        return lecture


# ── Chemin rapide : XML en flux ───────────────────────────────────
def _lire_rapide(chemin, colonne_id, colonnes, ligne_entete) -> LectureAudit:
    entetes = lire_entetes(chemin, ligne_entete)
    if not entetes:
        raise LectureImpossible("ligne d'en-tête vide")
    index_id, index_info, avertissements = _resoudre(entetes, colonne_id, colonnes)

    ids: dict[int, str] = {}
    derniere = ligne_entete
    for ligne, valeur in lire_colonne(chemin, index_id, ligne_entete + 1):
        derniere = max(derniere, ligne)
        if str(valeur).strip():
            ids[ligne] = str(valeur).strip()

    valeurs = {}
    for cle, (index, en_date) in index_info.items():
        valeurs[cle] = _colonne(chemin, index, ligne_entete, ids, en_date=en_date)

    return LectureAudit(
        lignes=_assembler(ids, valeurs),
        avertissements=avertissements,
        nb_lignes_lues=max(derniere - ligne_entete, 0),
    )


def _colonne(chemin, index, ligne_entete, ids, en_date=False) -> dict[int, str]:
    """Valeurs d'une colonne, restreintes aux lignes portant un ID erreur."""
    return {
        ligne: formater_valeur(valeur, en_date=en_date, brut=True)
        for ligne, valeur in lire_colonne(chemin, index, ligne_entete + 1)
        if ligne in ids
    }


# ── Repli : openpyxl en lecture seule ─────────────────────────────
def _lire_openpyxl(chemin, colonne_id, colonnes, ligne_entete) -> LectureAudit:
    classeur = load_workbook(chemin, read_only=True, data_only=True)
    try:
        feuille = choisir_feuille(classeur)
        entetes = entetes_depuis_feuille(feuille, ligne_entete)
        index_id, index_info, avertissements = _resoudre(entetes, colonne_id, colonnes)
        utiles = [index_id, *(i for i, _ in index_info.values())]
        col_min, col_max = min(utiles), max(utiles)

        ids: dict[int, str] = {}
        valeurs: dict[str, dict[int, str]] = {k: {} for k in index_info}
        ligne = ligne_entete
        for ligne, rangee in enumerate(
            feuille.iter_rows(
                min_row=ligne_entete + 1, min_col=col_min, max_col=col_max, values_only=True,
            ),
            start=ligne_entete + 1,
        ):
            def lire(index):
                rang = index - col_min
                return rangee[rang] if rang < len(rangee) else None

            identifiant = lire(index_id)
            if identifiant is None or not str(identifiant).strip():
                continue
            ids[ligne] = str(identifiant).strip()
            for cle, (index, en_date) in index_info.items():
                valeurs[cle][ligne] = formater_valeur(lire(index), en_date=en_date)

        return LectureAudit(
            lignes=_assembler(ids, valeurs),
            avertissements=avertissements,
            nb_lignes_lues=max(ligne - ligne_entete, 0),
        )
    finally:
        classeur.close()


# ── Résolution des colonnes ───────────────────────────────────────
def _resoudre(entetes: dict[int, str], colonne_id: str, colonnes: ColonnesAudit):
    """Index des colonnes ; lève pour une colonne obligatoire absente."""
    table = table_depuis_entetes(entetes)
    noms = {i: n for i, n in entetes.items()}

    def trouver(reference: str) -> int | None:
        index = resoudre_colonne(reference, table)
        if index is None or index in noms:
            return index
        # Colonne sans en-tête : retenue seulement si l'utilisateur a saisi
        # explicitement une lettre (« Z »), jamais pour un libellé absent.
        return index if _LETTRE.fullmatch(reference.strip()) else None

    index_id = trouver(colonne_id)
    if index_id is None:
        raise ColonneAbsente(f"colonne « {colonne_id} » (ID erreur) introuvable")

    index_info: dict[str, tuple[int, bool]] = {}
    absentes = []
    for cle, libelle in colonnes.controlees().items():
        index = trouver(libelle)
        if index is None:
            absentes.append(f"« {libelle} »")
            continue
        en_date = cle == "date_audit" or entete_est_date(noms.get(index, libelle))
        index_info[cle] = (index, en_date)
    if absentes:
        pluriel = "s" if len(absentes) > 1 else ""
        raise ColonneAbsente(
            f"colonne{pluriel} {', '.join(absentes)} introuvable{pluriel} — "
            "les quatre colonnes contrôlées sont obligatoires"
        )
    return index_id, index_info, []


def _assembler(ids: dict[int, str], valeurs: dict[str, dict[int, str]]) -> list[LigneAudit]:
    return [
        LigneAudit(
            ligne=ligne,
            id_erreur=identifiant,
            validation=valeurs.get("validation", {}).get(ligne, ""),
            audit=valeurs.get("audit", {}).get(ligne, ""),
            date_audit=valeurs.get("date_audit", {}).get(ligne, ""),
            auditeur=valeurs.get("auditeur", {}).get(ligne, ""),
        )
        for ligne, identifiant in sorted(ids.items())
    ]


# ── Mise en forme des valeurs lues ────────────────────────────────
def formater_valeur(valeur, en_date: bool = False, brut: bool = False) -> str:
    """Valeur de cellule → texte affichable.

    ``brut`` : la valeur vient du XML (une date y est un numéro de série) ;
    dans une colonne de dates, ce numéro est reconverti en ``JJ/MM/AAAA``.
    """
    if valeur is None:
        return ""
    if isinstance(valeur, datetime):
        return valeur.strftime("%d/%m/%Y" if valeur.time() == time() else "%d/%m/%Y %H:%M")
    if isinstance(valeur, date):
        return valeur.strftime("%d/%m/%Y")
    if isinstance(valeur, bool):
        return "VRAI" if valeur else "FAUX"
    if en_date:
        serie = to_num(valeur)
        if serie is not None and _SERIE_MIN <= serie <= _SERIE_MAX:
            moment = _EPOQUE_EXCEL + timedelta(days=serie)
            return formater_valeur(moment.replace(microsecond=0))
    if isinstance(valeur, float) and valeur.is_integer():
        return str(int(valeur))
    texte = str(valeur).strip()
    if brut and texte.endswith(".0") and to_num(texte) is not None:
        return texte[:-2]
    return texte
