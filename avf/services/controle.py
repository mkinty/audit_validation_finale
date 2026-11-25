"""Contrôle d'une ligne sur ses quatre colonnes : validée, ou non valide (avec motifs).

    Validation finale  égale à une des valeurs attendues (défaut : OK)
    Audit              égal à une des valeurs attendues (défaut : OK ou KO)
    Date audit         renseignée et lisible comme une date, même saisie en texte
    Auditeur           renseigné

Les valeurs sont comparées sans tenir compte de la casse ni des espaces autour
(« ok », « OK  » valent « OK ») ; une cellule faite d'espaces est vide.

Dates acceptées (en plus des vraies dates Excel, que la lecture restitue en
``JJ/MM/AAAA``) :

    12/03/2026  12-03-2026  12.03.2026  12/03/26  (jour/mois, à la française)
    2026-03-12  2026/03/12  (ISO), éventuellement suivies d'une heure
    12 mars 2026  1er mars 2026  12 sept. 2026  (mois en lettres)
    un numéro de série Excel (date stockée comme nombre sans format date)

Une date impossible (31/02/2026) n'est pas une date.
"""

from __future__ import annotations

from datetime import datetime
import re
import unicodedata

from ..models.audit import Anomalie, LigneAudit, est_vide
from ..models.colonnes import ColonnesAudit

# Numéro de série Excel plausible pour une date d'audit (1954 → 2119), comme à la lecture.
_SERIE_MIN, _SERIE_MAX = 20_000, 80_000

_HEURE = r"(?:[ t]+\d{1,2}[:h]\d{2}(?::\d{2}(?:\.\d+)?)?)?"   # texte déjà en minuscules
_JMA = re.compile(rf"(\d{{1,2}})[/.\-](\d{{1,2}})[/.\-](\d{{4}}|\d{{2}}){_HEURE}")
_AMJ = re.compile(rf"(\d{{4}})[/.\-](\d{{1,2}})[/.\-](\d{{1,2}}){_HEURE}")
_LETTRES = re.compile(r"(\d{1,2})(?:er)?\s+([a-z]+)\.?\s+(\d{4})" + _HEURE)
_SERIE = re.compile(r"\d{5}(?:[.,]\d+)?")

_MOIS = {
    "janvier": 1, "janv": 1, "jan": 1, "fevrier": 2, "fevr": 2, "fev": 2, "mars": 3, "mar": 3,
    "avril": 4, "avr": 4, "mai": 5, "juin": 6, "juillet": 7, "juil": 7, "aout": 8,
    "septembre": 9, "sept": 9, "sep": 9, "octobre": 10, "oct": 10, "novembre": 11, "nov": 11,
    "decembre": 12, "dec": 12,
}


def est_date(valeur) -> bool:
    """Vrai si ``valeur`` (texte lu dans la cellule) représente une date réelle."""
    if isinstance(valeur, datetime):
        return True
    texte = _sans_accents(str(valeur or "")).strip().lower()
    if not texte:
        return False
    if m := _JMA.fullmatch(texte):
        jour, mois, annee = (int(x) for x in m.groups())
        return _existe(annee + 2000 if annee < 100 else annee, mois, jour)
    if m := _AMJ.fullmatch(texte):
        annee, mois, jour = (int(x) for x in m.groups())
        return _existe(annee, mois, jour)
    if m := _LETTRES.fullmatch(texte):
        mois = _MOIS.get(m.group(2))
        return mois is not None and _existe(int(m.group(3)), mois, int(m.group(1)))
    if _SERIE.fullmatch(texte):
        return _SERIE_MIN <= float(texte.replace(",", ".")) <= _SERIE_MAX
    return False


def controler_ligne(ligne: LigneAudit, colonnes: ColonnesAudit) -> list[Anomalie]:
    """Critères non respectés par la ligne (liste vide : ligne validée)."""
    anomalies = []
    attendus_audit = colonnes.valeurs_audit
    if est_vide(ligne.audit):
        anomalies.append(Anomalie("audit", f"{colonnes.audit} vide"))
    elif not _parmi(ligne.audit, attendus_audit):
        anomalies.append(Anomalie(
            "audit", f"{colonnes.audit} « {ligne.audit.strip()} » ≠ {' / '.join(attendus_audit)}",
        ))
    if est_vide(ligne.date_audit):
        anomalies.append(Anomalie("date_audit", f"{colonnes.date_audit} vide"))
    elif not est_date(ligne.date_audit):
        anomalies.append(Anomalie(
            "date_audit", f"{colonnes.date_audit} « {ligne.date_audit.strip()} » n'est pas une date",
        ))
    if est_vide(ligne.auditeur):
        anomalies.append(Anomalie("auditeur", f"{colonnes.auditeur} vide"))
    attendus_vf = colonnes.valeurs_validation
    if est_vide(ligne.validation):
        anomalies.append(Anomalie("validation", f"{colonnes.validation} vide"))
    elif not _parmi(ligne.validation, attendus_vf):
        anomalies.append(Anomalie(
            "validation",
            f"{colonnes.validation} « {ligne.validation.strip()} » ≠ {' / '.join(attendus_vf)}",
        ))
    return anomalies


def controler(lignes: list[LigneAudit], colonnes: ColonnesAudit) -> list[LigneAudit]:
    """Renseigne les anomalies de chaque ligne (en place) et renvoie la liste."""
    for ligne in lignes:
        ligne.anomalies = controler_ligne(ligne, colonnes)
    return lignes


def _parmi(valeur: str, attendus) -> bool:
    cle = str(valeur).strip().casefold()
    return any(cle == str(a).strip().casefold() for a in attendus)


def _existe(annee: int, mois: int, jour: int) -> bool:
    try:
        datetime(annee, mois, jour)
        return True
    except ValueError:
        return False


def _sans_accents(texte: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texte) if unicodedata.category(c) != "Mn")
