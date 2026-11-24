"""Application des règles sur un classeur, filtrée par les ID erreur."""

import os

import pytest
from openpyxl import Workbook, load_workbook

from sna.models.regles import (
    Action, Condition, RegleModification, ReglesConfig, SourceValeur, TypeDroite,
)
from sna.services.colonnes_excel import OngletIntrouvable
from sna.services.excel import ColonneIntrouvable, appliquer_regles

VALEUR = "Vérification RT impossible nb log"


def _config(*regles, ligne_entete=1, colonne_id="ID erreur"):
    return ReglesConfig(
        regles=list(regles), ligne_entete=ligne_entete, colonne_id=colonne_id
    )


def _regle(colonne="Z", valeur=VALEUR, code="R1", actif=True,
           source=SourceValeur.VALEUR, conditions=None, exact=True):
    return RegleModification(
        code=code, libelle=code,
        action=Action(colonne=colonne, source=source, valeur=valeur),
        actif=actif, conditions=conditions or [], exact=exact,
    )


# ── Filtrage par ID erreur ────────────────────────────────────────
def test_seules_les_lignes_des_ids_demandes_sont_modifiees(creer_classeur):
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2", "13001_3"],
        "Z": ["à revoir", "à revoir", "à revoir"],
    })

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1", "13001_3"])

    assert bilan.comptes == {"R1": 2}
    assert bilan.nb_lignes == 2
    ws = load_workbook(chemin).active
    assert ws["Z2"].value == VALEUR
    assert ws["Z3"].value == "à revoir"   # ID non demandé : ligne intacte
    assert ws["Z4"].value == VALEUR


def test_ids_trouves_et_absents_sont_remontes(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1", "13001_9"])

    assert bilan.ids_trouves == {"13001_1"}


def test_correspondance_tolerante_a_la_casse_et_aux_espaces(creer_classeur):
    chemin = creer_classeur({"C": [" 2a004_3 "], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle()), ["2A004_3"])

    assert bilan.comptes == {"R1": 1}


def test_aucun_id_demande_ne_modifie_rien(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle()), [])

    assert bilan.nb_modifications == 0
    assert load_workbook(chemin).active["Z2"].value == "x"


def test_l_entete_n_est_jamais_modifie(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    ws = load_workbook(chemin).active
    assert ws["Z1"].value == "Audit - Street View"


def test_plusieurs_lignes_portant_le_meme_id(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1", "13001_1"], "Z": ["a", "b"]})

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert bilan.comptes == {"R1": 2} and bilan.nb_lignes == 2


# ── Désignation des colonnes ──────────────────────────────────────
def test_colonne_designee_par_sa_lettre(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle(colonne="Z")), ["13001_1"])

    assert bilan.comptes == {"R1": 1}


def test_colonne_id_designee_par_sa_lettre(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle(), colonne_id="C"), ["13001_1"])

    assert bilan.comptes == {"R1": 1}


def test_nom_de_colonne_tolerant_aux_espaces_et_a_la_casse(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})
    config = _config(_regle(colonne="audit-streetview"), colonne_id="iderreur")

    assert appliquer_regles(chemin, config, ["13001_1"]).comptes == {"R1": 1}


def test_colonne_de_regle_inconnue_leve_une_erreur(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    with pytest.raises(ColonneIntrouvable):
        appliquer_regles(chemin, _config(_regle(colonne="Colonne fantôme")), ["13001_1"])


def test_colonne_id_absente_leve_une_erreur(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    with pytest.raises(ColonneIntrouvable) as erreur:
        appliquer_regles(chemin, _config(_regle(), colonne_id="Absente"), ["13001_1"])

    assert "ID erreur" in str(erreur.value)


# ── Règles multiples ──────────────────────────────────────────────
def test_toutes_les_regles_actives_sont_appliquees(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"], "A": ["y"]})
    config = _config(
        _regle(colonne="Z", valeur="statut", code="R1"),
        _regle(colonne="Adresse", valeur="rue corrigée", code="R2"),
    )

    bilan = appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.comptes == {"R1": 1, "R2": 1}
    ws = load_workbook(chemin).active
    assert ws["Z2"].value == "statut"
    assert ws["A2"].value == "rue corrigée"


def test_regle_inactive_ignoree(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle(actif=False)), ["13001_1"])

    assert bilan.comptes == {}
    assert load_workbook(chemin).active["Z2"].value == "x"


# ── Valeurs écrites ───────────────────────────────────────────────
def test_valeur_numerique_ecrite_comme_nombre(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "B": ["à chiffrer"]})

    appliquer_regles(chemin, _config(_regle(colonne="B", valeur="42")), ["13001_1"])

    assert load_workbook(chemin).active["B2"].value == 42


def test_effacement_explicite_vide_la_cellule(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["à effacer"]})
    regle = _regle(valeur="", source=SourceValeur.VIDE)

    bilan = appliquer_regles(chemin, _config(regle), ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["Z2"].value is None


def test_regle_sans_valeur_ne_touche_a_rien(creer_classeur):
    """Une règle « valeur » laissée vide reste sans effet."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["à garder"]})

    bilan = appliquer_regles(chemin, _config(_regle(valeur="")), ["13001_1"])

    assert bilan.comptes == {}
    assert load_workbook(chemin).active["Z2"].value == "à garder"


def test_valeur_deja_presente_n_est_pas_recomptee(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": [VALEUR]})

    assert appliquer_regles(chemin, _config(_regle()), ["13001_1"]).comptes == {"R1": 0}


# ── Écriture sur disque ───────────────────────────────────────────
def test_aucune_modification_ne_reecrit_pas_le_fichier(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": [VALEUR]})
    mtime_avant = os.path.getmtime(chemin)

    appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert os.path.getmtime(chemin) == mtime_avant


def test_simulation_ne_modifie_pas_le_fichier(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"], simuler=True)

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["Z2"].value == "x"


# ── Ligne d'en-tête ───────────────────────────────────────────────
def test_ligne_entete_personnalisee(creer_classeur):
    chemin = creer_classeur(
        {"A": ["ID erreur", "13001_1"], "B": ["Statut", "x"]}, entetes={}
    )
    config = _config(
        _regle(colonne="Statut", valeur="ok"), ligne_entete=2, colonne_id="ID erreur"
    )

    bilan = appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["B3"].value == "ok"


# ── Repli sur openpyxl ────────────────────────────────────────────
def test_repli_openpyxl_donne_le_meme_resultat(creer_classeur, monkeypatch):
    """Un classeur illisible par le lecteur rapide reste traité correctement."""
    from sna.services import excel as module_excel
    from sna.services.lecture_rapide import LectureImpossible

    def refuser(*_args, **_kwargs):
        raise LectureImpossible("structure simulée")

    chemin = creer_classeur({
        "C": ["13001_1", "13001_2", "13001_3"],
        "Z": ["a", "b", "c"],
    })
    monkeypatch.setattr(module_excel, "lire_entetes", refuser)
    messages = []

    bilan = appliquer_regles(
        chemin, _config(_regle()), ["13001_1", "13001_3"], on_etape=messages.append
    )

    assert bilan.comptes == {"R1": 2}
    assert bilan.ids_trouves == {"13001_1", "13001_3"}
    assert any("repli openpyxl" in m for m in messages)
    ws = load_workbook(chemin).active
    assert [ws[f"Z{ligne}"].value for ligne in (2, 3, 4)] == [VALEUR, "b", VALEUR]


# ── Étapes journalisées ───────────────────────────────────────────
def test_les_etapes_sont_rapportees_au_fil_du_traitement(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})
    messages = []

    appliquer_regles(chemin, _config(_regle()), ["13001_1"], on_etape=messages.append)

    assert messages[0].startswith("analyse")
    assert any("chargement" in m for m in messages)
    assert any("enregistrement" in m for m in messages)


def test_fichier_sans_correspondance_n_est_pas_charge(creer_classeur):
    """Le message doit dire explicitement que le classeur n'a pas été ouvert."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})
    messages = []

    appliquer_regles(chemin, _config(_regle()), ["99999_1"], on_etape=messages.append)

    assert any("non ouvert en écriture" in m for m in messages)
    assert not any("chargement" in m for m in messages)


def test_colonne_absente_detectee_avant_tout_chargement(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})
    messages = []

    with pytest.raises(ColonneIntrouvable):
        appliquer_regles(
            chemin, _config(_regle(colonne="Absente")), ["13001_1"],
            on_etape=messages.append,
        )

    assert not any("chargement" in m for m in messages)


def test_valeur_ressemblant_a_un_id_reste_du_texte(creer_classeur):
    """Écrire « 13001_2 » ne doit pas produire le nombre 130012."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    appliquer_regles(chemin, _config(_regle(valeur="13001_2")), ["13001_1"])

    assert load_workbook(chemin).active["Z2"].value == "13001_2"


# ── Conditions : filtre supplémentaire sur les lignes des ID erreur ──
def _cond(colonne, operateur, droite="", type_droite=TypeDroite.VALEUR):
    return Condition(colonne=colonne, operateur=operateur,
                     type_droite=type_droite, droite=droite)


def test_condition_restreint_les_lignes_modifiees(creer_classeur):
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2"],
        "Z": ["à revoir", "à revoir"],
        "B": [1, 5],
    })
    regle = _regle(colonne="Z", valeur="corrigé",
                   conditions=[_cond("Audit - Nb BAL", ">", "3")])

    bilan = appliquer_regles(chemin, _config(regle), ["13001_1", "13001_2"])

    assert bilan.comptes == {"R1": 1}
    ws = load_workbook(chemin).active
    assert ws["Z2"].value == "à revoir"   # B=1, condition fausse
    assert ws["Z3"].value == "corrigé"    # B=5, condition vraie


def test_condition_vide_sur_une_colonne(creer_classeur):
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2"], "A": ["déjà", None], "Z": ["x", "y"],
    })
    regle = _regle(colonne="Adresse", valeur="À COMPLÉTER",
                   conditions=[_cond("Adresse", "vide")])

    appliquer_regles(chemin, _config(regle), ["13001_1", "13001_2"])

    ws = load_workbook(chemin).active
    assert ws["A2"].value == "déjà"
    assert ws["A3"].value == "À COMPLÉTER"


def test_condition_comparee_a_une_autre_colonne(creer_classeur):
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2"],
        "A": ["même", "diff"],
        "Z": ["même", "autre"],
    })
    regle = _regle(colonne="Z", valeur="=A",
                   conditions=[_cond("Adresse", "=", "Audit - Street View",
                                     TypeDroite.COLONNE)])

    bilan = appliquer_regles(chemin, _config(regle), ["13001_1", "13001_2"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["Z2"].value == "=A"


def test_recopie_d_une_colonne_source(creer_classeur):
    chemin = creer_classeur({
        "C": ["13001_1"], "A": ["12 rue de la Paix"], "Z": ["à corriger"],
    })
    regle = _regle(colonne="Z", valeur="Adresse", source=SourceValeur.COLONNE)

    bilan = appliquer_regles(chemin, _config(regle), ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin).active["Z2"].value == "12 rue de la Paix"


def test_colonne_de_condition_absente_leve_une_erreur(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})
    regle = _regle(colonne="Z", valeur="v",
                   conditions=[_cond("Colonne fantôme", "=", "x")])

    with pytest.raises(ColonneIntrouvable):
        appliquer_regles(chemin, _config(regle), ["13001_1"])


def test_condition_faussse_partout_ne_charge_rien_en_ecriture(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"], "B": [1]})
    regle = _regle(colonne="Z", valeur="v",
                   conditions=[_cond("Audit - Nb BAL", ">", "100")])

    bilan = appliquer_regles(chemin, _config(regle), ["13001_1"])

    assert bilan.comptes == {"R1": 0}
    assert load_workbook(chemin).active["Z2"].value == "x"


# ── Sélection de l'onglet « Audit » ───────────────────────────────
def _classeur_multi_onglets(chemin, noms_feuilles, actif=0, ids=None, valeurs_z=None):
    """Classeur à plusieurs onglets ; seul celui nommé « Audit » porte des
    données. Les autres restent vides, pour qu'une lecture du mauvais onglet
    échoue plutôt que de réussir silencieusement.
    """
    wb = Workbook()
    wb.remove(wb.active)
    for nom in noms_feuilles:
        ws = wb.create_sheet(nom)
        ws["A1"], ws["C1"], ws["Z1"] = "Adresse", "ID erreur", "Audit - Street View"
    cible = wb[noms_feuilles[-1]]
    for i, (identifiant, valeur) in enumerate(zip(ids or [], valeurs_z or []), start=2):
        cible[f"C{i}"] = identifiant
        cible[f"Z{i}"] = valeur
    wb.active = actif
    wb.save(chemin)


def test_onglet_audit_choisi_meme_si_ce_n_est_pas_l_onglet_actif(tmp_path):
    """L'onglet actif à l'enregistrement (ici « Autre », en 1re position et
    actif) est ignoré : seul l'onglet nommé « Audit » est lu et modifié."""
    chemin = str(tmp_path / "multi.xlsx")
    _classeur_multi_onglets(
        chemin, ["Autre", "Audit"], actif=0, ids=["13001_1"], valeurs_z=["à revoir"],
    )

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    classeur = load_workbook(chemin)
    assert classeur["Audit"]["Z2"].value == VALEUR
    assert classeur["Autre"]["Z2"].value is None


def test_onglet_audit_insensible_a_la_casse_et_aux_espaces(tmp_path):
    chemin = str(tmp_path / "casse.xlsx")
    _classeur_multi_onglets(chemin, [" AUDIT "], ids=["13001_1"], valeurs_z=["x"])

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert bilan.comptes == {"R1": 1}


def test_onglet_audit_absent_leve_une_erreur(tmp_path):
    chemin = str(tmp_path / "sans_audit.xlsx")
    _classeur_multi_onglets(chemin, ["Feuil1", "Feuil2"], ids=["13001_1"], valeurs_z=["x"])

    with pytest.raises(OngletIntrouvable):
        appliquer_regles(chemin, _config(_regle()), ["13001_1"])


# ── Filtre automatique et lignes/colonnes masquées ────────────────
def test_filtre_et_masquage_retires_avant_ecriture(creer_classeur):
    """Une ligne filtrée (masquée) qui vient d'être modifiée doit rester visible."""
    chemin = creer_classeur({
        "C": ["13001_1", "13001_2"], "Z": ["à revoir", "à revoir"],
    })
    classeur = load_workbook(chemin)
    feuille = classeur.active
    feuille.auto_filter.ref = "A1:Z3"
    feuille.row_dimensions[3].hidden = True   # ligne de 13001_2, filtrée
    feuille.column_dimensions["Z"].hidden = True
    classeur.save(chemin)

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1", "13001_2"])

    assert bilan.comptes == {"R1": 2}
    assert bilan.visibilite_restauree is True
    relu = load_workbook(chemin)
    feuille_relue = relu.active
    assert feuille_relue.auto_filter.ref is None
    assert feuille_relue.row_dimensions[3].hidden is False
    assert feuille_relue.column_dimensions["Z"].hidden is False
    assert feuille_relue["Z3"].value == VALEUR  # la ligne masquée a bien été mise à jour


def test_classeur_sans_filtre_ni_masquage_n_est_pas_signale(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["x"]})

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert bilan.visibilite_restauree is False


def test_filtre_retire_meme_sans_aucune_cellule_a_changer(creer_classeur):
    """Le filtre doit disparaître même si les valeurs étaient déjà correctes."""
    chemin = creer_classeur({"C": ["13001_1"], "Z": [VALEUR]})
    classeur = load_workbook(chemin)
    feuille = classeur.active
    feuille.auto_filter.ref = "A1:Z2"
    classeur.save(chemin)

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"])

    assert bilan.comptes == {"R1": 0}
    assert bilan.visibilite_restauree is True
    assert bilan.enregistre is True
    assert load_workbook(chemin).active.auto_filter.ref is None


def test_simulation_ne_retire_pas_le_filtre_sur_le_disque(creer_classeur):
    chemin = creer_classeur({"C": ["13001_1"], "Z": ["à revoir"]})
    classeur = load_workbook(chemin)
    classeur.active.auto_filter.ref = "A1:Z2"
    classeur.save(chemin)

    bilan = appliquer_regles(chemin, _config(_regle()), ["13001_1"], simuler=True)

    assert bilan.visibilite_restauree is True       # constaté...
    assert bilan.enregistre is False
    assert load_workbook(chemin).active.auto_filter.ref == "A1:Z2"  # ... mais pas écrit
