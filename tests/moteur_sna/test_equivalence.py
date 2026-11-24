"""Équivalence avec une implémentation de référence de la spécification.

Le moteur optimisé (résolution des colonnes une seule fois, écriture
conditionnelle, exclusion des lignes fantômes) doit produire **exactement** le
même classeur qu'une transcription naïve et directe de la règle métier :

    pour chaque ligne de données,
        si la colonne « ID erreur » porte un identifiant demandé,
            chaque règle active écrit sa valeur dans sa colonne.

Ce test rejoue cette version naïve sur une matrice de cas et compare les deux
classeurs **cellule par cellule**, ainsi que le nombre de modifications.
"""

import pytest
from openpyxl import Workbook, load_workbook

from sna.models.regles import Action, RegleModification, ReglesConfig, SourceValeur
from sna.services.excel import appliquer_regles
from sna.services.valeurs import analyser_date

COLONNE_ID = "C"
ENTETES = {"A": "Adresse", "B": "Nb", "C": "ID erreur", "Z": "Statut"}

# (ids du fichier, valeurs de Z, valeurs de A, ids demandés, règles [(lettre, valeur)])
CAS = [
    # Cas nominal : deux lignes sur trois sont visées.
    (["13001_1", "13001_2", "13001_3"], ["a", "b", "c"], ["x", "y", "z"],
     ["13001_1", "13001_3"], [("Z", "corrigé")]),
    # Aucun identifiant demandé n'est présent dans le fichier.
    (["13001_1"], ["a"], ["x"], ["13001_9"], [("Z", "corrigé")]),
    # Valeur déjà en place : aucune écriture ne doit être comptée.
    (["13001_1"], ["corrigé"], ["x"], ["13001_1"], [("Z", "corrigé")]),
    # Deux règles sur deux colonnes différentes.
    (["13001_1", "13001_2"], ["a", "b"], ["x", "y"], ["13001_2"],
     [("Z", "corrigé"), ("A", "12 rue de la Paix")]),
    # Deux règles sur la même colonne : la dernière l'emporte.
    (["13001_1"], ["a"], ["x"], ["13001_1"], [("Z", "premier"), ("Z", "second")]),
    # Valeur vide : la cellule est effacée.
    (["13001_1", "13001_2"], ["a", "b"], ["x", "y"], ["13001_1"], [("Z", "")]),
    # Valeur numérique : écrite comme un nombre.
    (["13001_1"], ["a"], ["x"], ["13001_1"], [("B", "42")]),
    # Cellules d'origine vides ou numériques.
    (["13001_1", "13001_2"], [None, 12], ["x", None], ["13001_1", "13001_2"],
     [("Z", "corrigé")]),
    # Identifiants du fichier avec espaces et casse différentes.
    ([" 2a004_3 ", "2A004_4"], ["a", "b"], ["x", "y"], ["2A004_3"], [("Z", "corrigé")]),
    # Colonne ID erreur partiellement vide.
    ([None, "13001_1", ""], ["a", "b", "c"], ["x", "y", "z"], ["13001_1"],
     [("Z", "corrigé")]),
    # Fichier sans aucune ligne de données.
    ([], [], [], ["13001_1"], [("Z", "corrigé")]),
]


def _ecrire(chemin, ids, valeurs_z, valeurs_a):
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    for lettre, libelle in ENTETES.items():
        ws[f"{lettre}1"] = libelle
    for i, identifiant in enumerate(ids, start=2):
        ws[f"{COLONNE_ID}{i}"] = identifiant
    for i, valeur in enumerate(valeurs_z, start=2):
        ws[f"Z{i}"] = valeur
    for i, valeur in enumerate(valeurs_a, start=2):
        ws[f"A{i}"] = valeur
    wb.save(chemin)


def _implementation_reference(chemin, ids_demandes, regles):
    """Transcription directe de la règle métier, sans aucune optimisation."""
    attendus = {"".join(str(i).split()).upper() for i in ids_demandes}
    classeur = load_workbook(chemin)
    feuille = classeur.active
    nb = 0
    try:
        for ligne in range(2, feuille.max_row + 1):
            brut = feuille[f"{COLONNE_ID}{ligne}"].value
            cle = "" if brut is None else "".join(str(brut).split()).upper()
            if cle not in attendus:
                continue
            for lettre, valeur in regles:
                cellule = feuille[f"{lettre}{ligne}"]
                nouvelle = _typer(valeur)
                if cellule.value == nouvelle and type(cellule.value) is type(nouvelle):
                    continue
                cellule.value = nouvelle
                nb += 1
        if nb:
            classeur.save(chemin)
        return nb
    finally:
        classeur.close()


def _typer(valeur: str):
    if valeur == "":
        return None
    return int(valeur) if valeur.isdigit() else valeur


def _toutes_les_cellules(chemin):
    ws = load_workbook(chemin).active
    return [
        [ws.cell(ligne, colonne).value for colonne in range(1, 27)]
        for ligne in range(1, (ws.max_row or 1) + 1)
    ]


@pytest.mark.parametrize("ids,valeurs_z,valeurs_a,demandes,regles", CAS)
def test_meme_resultat_que_l_implementation_de_reference(
    tmp_path, ids, valeurs_z, valeurs_a, demandes, regles
):
    reference = str(tmp_path / "reference.xlsx")
    moteur = str(tmp_path / "moteur.xlsx")
    _ecrire(reference, ids, valeurs_z, valeurs_a)
    _ecrire(moteur, ids, valeurs_z, valeurs_a)

    config = ReglesConfig(
        regles=[
            RegleModification(
                code=f"R{i}",
                action=Action(
                    colonne=lettre, valeur=valeur,
                    source=SourceValeur.VIDE if valeur == "" else SourceValeur.VALEUR,
                ),
            )
            for i, (lettre, valeur) in enumerate(regles, start=1)
        ],
        ligne_entete=1,
        colonne_id="ID erreur",
    )

    nb_reference = _implementation_reference(reference, demandes, regles)
    bilan = appliquer_regles(moteur, config, demandes)

    assert bilan.nb_modifications == nb_reference
    assert _toutes_les_cellules(moteur) == _toutes_les_cellules(reference)


# En-têtes de toutes les colonnes citées par les 16 règles d'audit livrées
# (cibles, conditions et colonnes comparées / recopiées).
_ENTETES_DEFAUT = [
    "ID erreur", "Audit", "Validation finale", "Date audit", "Auditeur",
    "Date Analyse", "CDP", "Analyse nombre de Logement",
    "OK", "NOK", "SNA", "01/01/2026", "31/12/2026", "X BAN",
]


def _ecrire_audit_complet(chemin, id_erreur="13001_1", validation_finale=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    for col, libelle in enumerate(_ENTETES_DEFAUT, start=1):
        ws.cell(1, col, libelle)
    ws.cell(2, 1, id_erreur)   # colonne « ID erreur »
    if validation_finale:
        index = _ENTETES_DEFAUT.index("Validation finale") + 1
        ws.cell(2, index, validation_finale)
    wb.save(chemin)


def test_la_config_livree_s_applique_sans_erreur(tmp_path):
    """Les règles d'audit par défaut s'appliquent proprement et sont idempotentes."""
    chemin = str(tmp_path / "audit.xlsx")
    _ecrire_audit_complet(chemin, validation_finale="Terminé")
    defaut = ReglesConfig.defaut()

    bilan = appliquer_regles(chemin, defaut, ["13001_1"])
    assert bilan.nb_modifications >= 1

    ws = load_workbook(chemin).active
    valeurs = {ws.cell(1, c).value: ws.cell(2, c).value for c in range(1, len(_ENTETES_DEFAUT) + 1)}
    for regle in defaut.actives:
        attendu = regle.action.valeur
        moment = analyser_date(attendu)
        assert valeurs[regle.colonne_cible] == (moment if moment is not None else attendu)

    # Deuxième passage : plus rien à corriger.
    bilan2 = appliquer_regles(chemin, defaut, ["13001_1"])
    assert bilan2.nb_modifications == 0


# ── Équivalence des deux chemins de lecture ───────────────────────
@pytest.mark.parametrize("ids,valeurs_z,valeurs_a,demandes,regles", CAS)
def test_lecture_rapide_et_openpyxl_donnent_le_meme_resultat(
    tmp_path, monkeypatch, ids, valeurs_z, valeurs_a, demandes, regles
):
    """Le repli openpyxl doit produire exactement le même classeur.

    Le repérage rapide lit le XML directement ; c'est un chemin distinct de
    celui d'openpyxl, avec ses propres pièges (entités, chaînes partagées).
    Cette comparaison croisée est le garde-fou de cette optimisation.
    """
    from sna.services import excel as module_excel
    from sna.services.lecture_rapide import LectureImpossible

    rapide = str(tmp_path / "rapide.xlsx")
    repli = str(tmp_path / "repli.xlsx")
    _ecrire(rapide, ids, valeurs_z, valeurs_a)
    _ecrire(repli, ids, valeurs_z, valeurs_a)

    config = ReglesConfig(
        regles=[
            RegleModification(
                code=f"R{i}",
                action=Action(
                    colonne=lettre, valeur=valeur,
                    source=SourceValeur.VIDE if valeur == "" else SourceValeur.VALEUR,
                ),
            )
            for i, (lettre, valeur) in enumerate(regles, start=1)
        ],
        ligne_entete=1,
        colonne_id="ID erreur",
    )

    bilan_rapide = appliquer_regles(rapide, config, demandes)

    def refuser(*_args, **_kwargs):
        raise LectureImpossible("repli forcé")

    monkeypatch.setattr(module_excel, "lire_entetes", refuser)
    bilan_repli = appliquer_regles(repli, config, demandes)

    assert bilan_repli.comptes == bilan_rapide.comptes
    assert bilan_repli.ids_trouves == bilan_rapide.ids_trouves
    assert bilan_repli.nb_lignes == bilan_rapide.nb_lignes
    assert _toutes_les_cellules(repli) == _toutes_les_cellules(rapide)
