"""Colonnes de dates et reprise de la mise en forme du classeur.

Deux exigences d'apparence cosmétique mais lourdes de conséquences :
une date écrite comme du texte n'est ni triable ni filtrable dans Excel, et une
valeur écrite dans une autre police que sa colonne se repère immédiatement à
l'œil dans un livrable.
"""

from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

from sna.models.regles import Action, RegleModification, ReglesConfig
from sna.services.excel import appliquer_regles
from sna.services.valeurs import FORMAT_DATE_DEFAUT


def _config(*regles, harmoniser_style=True):
    return ReglesConfig(
        regles=list(regles), ligne_entete=1, colonne_id="ID erreur",
        harmoniser_style=harmoniser_style,
    )


def _regle(colonne, valeur, code="R1"):
    return RegleModification(
        code=code, libelle=code, action=Action(colonne=colonne, valeur=valeur)
    )


# ── Colonnes de dates ─────────────────────────────────────────────
def test_colonne_reconnue_par_son_entete(creer_classeur):
    """Tout en-tête contenant le mot « date » impose une écriture datée."""
    chemin = creer_classeur(
        {"C": ["13001_1"], "A": ["à remplir"]},
        entetes={"A": "Date de visite", "C": "ID erreur"},
    )

    appliquer_regles(chemin, _config(_regle("Date de visite", "12/03/2026")), ["13001_1"])

    cellule = load_workbook(chemin).active["A2"]
    assert cellule.value == datetime(2026, 3, 12)
    assert cellule.number_format == FORMAT_DATE_DEFAUT


def test_colonne_reconnue_par_ses_valeurs(creer_classeur):
    """Une colonne dont les valeurs sont des dates l'est aussi, sans le dire."""
    chemin = creer_classeur(
        {"C": ["13001_1", "13001_2"], "A": [datetime(2025, 1, 1), datetime(2025, 2, 2)]},
        entetes={"A": "Visite", "C": "ID erreur"},
    )

    appliquer_regles(chemin, _config(_regle("Visite", "12/03/2026")), ["13001_2"])

    assert load_workbook(chemin).active["A3"].value == datetime(2026, 3, 12)


def test_format_d_affichage_existant_conserve(creer_classeur):
    """Un classeur qui affiche « 12 mars 2026 » doit continuer de le faire."""
    chemin = creer_classeur(
        {"C": ["13001_1", "13001_2"], "A": [datetime(2025, 1, 1), None]},
        entetes={"A": "Date de visite", "C": "ID erreur"},
    )
    classeur = load_workbook(chemin)
    classeur.active["A2"].number_format = "DD MMMM YYYY"
    classeur.save(chemin)

    appliquer_regles(chemin, _config(_regle("Date de visite", "12/03/2026")), ["13001_2"])

    assert load_workbook(chemin).active["A3"].number_format == "DD MMMM YYYY"


def test_saisie_non_datee_reste_du_texte(creer_classeur):
    chemin = creer_classeur(
        {"C": ["13001_1"], "A": ["x"]}, entetes={"A": "Date de visite", "C": "ID erreur"}
    )

    appliquer_regles(chemin, _config(_regle("Date de visite", "à planifier")), ["13001_1"])

    assert load_workbook(chemin).active["A2"].value == "à planifier"


def test_hors_colonne_de_dates_rien_ne_change(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    appliquer_regles(chemin, _config(_regle("Z", "12/03/2026")), ["13001_1"])

    assert load_workbook(chemin).active["Z2"].value == "12/03/2026"


# ── Reprise de la mise en forme ───────────────────────────────────
def _preparer_police(chemin, cellules: dict, police: Font, alignement: Alignment = None):
    classeur = load_workbook(chemin)
    feuille = classeur.active
    for reference in cellules:
        feuille[reference].font = police
        if alignement is not None:
            feuille[reference].alignment = alignement
    classeur.save(chemin)


def test_police_de_la_colonne_reprise(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1", "13001_2"], "Z": ["repère", "à changer"]})
    _preparer_police(
        chemin, {"Z2"}, Font(name="Calibri", size=14, bold=True, italic=True),
        Alignment(horizontal="center"),
    )

    appliquer_regles(chemin, _config(_regle("Z", "corrigé")), ["13001_2"])

    cellule = load_workbook(chemin).active["Z3"]
    assert cellule.value == "corrigé"
    assert (cellule.font.name, cellule.font.size) == ("Calibri", 14)
    assert cellule.font.bold and cellule.font.italic
    assert cellule.alignment.horizontal == "center"


def test_le_modele_n_est_pas_une_ligne_modifiee(creer_classeur):
    """La forme reprise est celle du fichier, jamais celle qu'on vient d'écrire."""
    chemin = creer_classeur({"C": ["13001_1", "13001_2"], "Z": ["repère", "autre"]})
    _preparer_police(chemin, {"Z2"}, Font(name="Calibri", size=14))
    _preparer_police(chemin, {"Z3"}, Font(name="Wingdings", size=6))

    appliquer_regles(chemin, _config(_regle("Z", "corrigé")), ["13001_2"])

    assert load_workbook(chemin).active["Z3"].font.name == "Calibri"


def test_option_desactivee_laisse_la_forme_intacte(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1", "13001_2"], "Z": ["repère", "à changer"]})
    _preparer_police(chemin, {"Z2"}, Font(name="Calibri", size=14))
    _preparer_police(chemin, {"Z3"}, Font(name="Wingdings", size=6))

    appliquer_regles(
        chemin, _config(_regle("Z", "corrigé"), harmoniser_style=False), ["13001_2"]
    )

    assert load_workbook(chemin).active["Z3"].font.name == "Wingdings"


# ── Colonnes vides : police dominante du classeur ─────────────────
def test_colonne_vide_reprend_la_police_dominante(creer_classeur):
    """Cas des colonnes d'audit : vides, donc sans modèle dans leur colonne.

    Sans ce relais, openpyxl écrirait en Calibri 11 dans un classeur composé
    en Calibri 10, et la valeur se repérerait à l'œil.
    """
    chemin = creer_classeur(
        {"C": ["13001_1", "13001_2"], "A": ["rue A", "rue B"], "Z": [None, None]}
    )
    _preparer_police(chemin, {"A2", "A3", "C2", "C3"}, Font(name="Calibri", size=10))

    appliquer_regles(chemin, _config(_regle("Z", "OK")), ["13001_1"])

    cellule = load_workbook(chemin).active["Z2"]
    assert cellule.value == "OK"
    assert (cellule.font.name, cellule.font.size) == ("Calibri", 10)


def test_police_dominante_l_emporte_sur_une_exception(creer_classeur):
    """C'est la police la plus employée qui fait référence, pas la première."""
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2", "13001_3"],
        "A": ["rue A", "rue B", "rue C"],
        "Z": [None, None, None],
    })
    _preparer_police(chemin, {"A2"}, Font(name="Arial", size=18))
    _preparer_police(chemin, {"A3", "A4", "C2", "C3", "C4"}, Font(name="Calibri", size=10))

    appliquer_regles(chemin, _config(_regle("Z", "OK")), ["13001_1"])

    assert load_workbook(chemin).active["Z2"].font.name == "Calibri"


def test_la_colonne_prime_sur_la_police_dominante(creer_classeur):
    """Une colonne qui a sa propre forme la conserve."""
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2"], "A": ["rue A", "rue B"], "Z": ["repère", None],
    })
    _preparer_police(chemin, {"A2", "A3", "C2", "C3"}, Font(name="Calibri", size=10))
    _preparer_police(chemin, {"Z2"}, Font(name="Consolas", size=8))

    appliquer_regles(chemin, _config(_regle("Z", "OK")), ["13001_2"])

    assert load_workbook(chemin).active["Z3"].font.name == "Consolas"


def test_option_desactivee_ignore_aussi_la_police_dominante(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "A": ["rue A"], "Z": [None]})
    _preparer_police(chemin, {"A2", "C2"}, Font(name="Calibri", size=10))
    _preparer_police(chemin, {"Z2"}, Font(name="Wingdings", size=6))

    appliquer_regles(
        chemin, _config(_regle("Z", "OK"), harmoniser_style=False), ["13001_1"]
    )

    assert load_workbook(chemin).active["Z2"].font.name == "Wingdings"


def test_classeur_sans_donnees_ne_bloque_pas(creer_classeur):
    """Aucune police à relever : la cellule garde simplement la sienne."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": [None]})

    bilan = appliquer_regles(chemin, _config(_regle("Z", "OK")), ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["Z2"].value == "OK"


# ── Effets de bord d'openpyxl ─────────────────────────────────────
def test_aucune_ligne_fantome_ajoutee(creer_classeur):
    """Consulter une cellule la crée : le classeur ne doit pas gonfler pour autant."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    appliquer_regles(chemin, _config(_regle("Z", "OK")), ["13001_1"])

    feuille = load_workbook(chemin).active
    assert feuille.max_row == 2
    assert feuille.max_column == 26
