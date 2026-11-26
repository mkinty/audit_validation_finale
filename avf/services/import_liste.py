"""Import d'une liste de communes (codes INSEE) ou d'ID erreur depuis un fichier.

Formats : ``.txt`` (un code par ligne, ou séparés par espaces / virgules /
points-virgules), ``.csv`` (séparateur détecté), ``.xlsx`` / ``.xlsm``.

Dans un CSV ou un classeur, la colonne est cherchée par son en-tête (casse,
espaces, accents et ``_`` ignorés) parmi des libellés usuels — « INSEE »,
« Code INSEE », « Commune »… ou « ID erreur », « ID_erreur », « Identifiant »… —
dans les 20 premières lignes de chaque feuille. Sans colonne reconnue, tout le
contenu est lu : les codes sont ensuite extraits par le même motif que la saisie
(``sna.services.identifiants``), donc un en-tête ou une colonne de texte ne gêne pas.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
import io
import os
import unicodedata

from openpyxl import load_workbook

EXTENSIONS = (".txt", ".csv", ".xlsx", ".xlsm")

COLONNES_INSEE = ("insee", "codeinsee", "codecommune", "commune", "codeinseecommune", "inseecommune", "code")
COLONNES_ID = ("iderreur", "iderreurs", "identifiant", "identifianterreur", "id")
COLONNES_ID_IMM = ("idimm", "idimmeuble", "identifiantimmeuble", "idimmeubles", "immeuble", "imm")

_LIGNES_ENTETE = 20


class ImportImpossible(ValueError):
    """Fichier illisible ou de format non pris en charge."""


@dataclass
class ListeImportee:
    valeurs: list[str] = field(default_factory=list)
    origine: str = ""             # « colonne « INSEE » (feuille Communes) » ou « tout le fichier »

    @property
    def texte(self) -> str:
        return "\n".join(self.valeurs)


def normaliser_entete(valeur) -> str:
    texte = unicodedata.normalize("NFD", str(valeur or ""))
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return "".join(c for c in texte.casefold() if c.isalnum())


def lire_liste(chemin: str, colonnes: tuple[str, ...]) -> ListeImportee:
    """Valeurs de la colonne reconnue (``colonnes``), sinon tout le contenu du fichier."""
    extension = os.path.splitext(chemin)[1].lower()
    try:
        if extension in (".xlsx", ".xlsm"):
            return _lire_classeur(chemin, colonnes)
        if extension == ".csv":
            return _lire_tableau(_lignes_csv(_texte(chemin)), colonnes, "")
        if extension == ".txt":
            return ListeImportee(_texte(chemin).split(), "tout le fichier")
    except ImportImpossible:
        raise
    except Exception as exc:  # noqa: BLE001 — fichier corrompu, verrouillé…
        raise ImportImpossible(f"fichier illisible : {exc}") from exc
    raise ImportImpossible(f"format non pris en charge ({extension or 'sans extension'}) — "
                           f"utilisez {', '.join(EXTENSIONS)}")


def _texte(chemin: str) -> str:
    brut = open(chemin, "rb").read()
    for encodage in ("utf-8-sig", "cp1252"):
        try:
            return brut.decode(encodage)
        except UnicodeDecodeError:
            continue
    return brut.decode("latin-1")


def _lignes_csv(texte: str) -> list[list[str]]:
    echantillon = texte[:20000]
    try:
        dialecte = csv.Sniffer().sniff(echantillon, delimiters=";,\t|")
    except csv.Error:
        dialecte = csv.excel
        dialecte.delimiter = ";" if echantillon.count(";") >= echantillon.count(",") else ","
    return [ligne for ligne in csv.reader(io.StringIO(texte), dialecte)]


def _lire_classeur(chemin: str, colonnes: tuple[str, ...]) -> ListeImportee:
    classeur = load_workbook(chemin, read_only=True, data_only=True)
    try:
        feuilles = [(f.title, [list(r) for r in f.iter_rows(values_only=True)]) for f in classeur.worksheets]
    finally:
        classeur.close()
    for titre, lignes in feuilles:                  # 1re feuille ayant la colonne cherchée
        resultat = _colonne(lignes, colonnes, titre)
        if resultat is not None:
            return resultat
    valeurs = [_cellule(v) for _, lignes in feuilles for ligne in lignes for v in ligne if _cellule(v)]
    return ListeImportee(valeurs, "tout le fichier (aucune colonne reconnue)")


def _lire_tableau(lignes: list[list], colonnes: tuple[str, ...], titre: str) -> ListeImportee:
    resultat = _colonne(lignes, colonnes, titre)
    if resultat is not None:
        return resultat
    valeurs = [_cellule(v) for ligne in lignes for v in ligne if _cellule(v)]
    return ListeImportee(valeurs, "tout le fichier (aucune colonne reconnue)")


def _colonne(lignes: list[list], colonnes: tuple[str, ...], titre: str) -> ListeImportee | None:
    voulues = {normaliser_entete(c) for c in colonnes}
    for rang, ligne in enumerate(lignes[:_LIGNES_ENTETE]):
        entetes = [normaliser_entete(v) for v in ligne]
        for priorite in colonnes:                   # l'ordre de ``colonnes`` départage
            cle = normaliser_entete(priorite)
            if cle in entetes and cle in voulues:
                index = entetes.index(cle)
                valeurs = [_cellule(l[index]) for l in lignes[rang + 1:] if index < len(l) and _cellule(l[index])]
                libelle = ligne[index]
                ou = f" (feuille {titre})" if titre else ""
                return ListeImportee(valeurs, f"colonne « {libelle} »{ou}")
    return None


def _cellule(valeur) -> str:
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    if isinstance(valeur, int):
        # Code INSEE saisi comme nombre dans Excel : le zéro de tête a été perdu (4006 → 04006).
        return f"{valeur:05d}" if 0 < valeur < 100000 else str(valeur)
    return str(valeur).strip()
