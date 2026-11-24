"""ID erreur non appliqués, conservés d'un traitement à l'autre pour être rejoués."""

from sna.models.base import ResultatCommune, Statut
from sna.services import attente


# ── Lecture / écriture du fichier ─────────────────────────────────
def test_fichier_absent_donne_une_liste_vide(tmp_path):
    assert attente.charger(str(tmp_path / "absent.json")) == {}


def test_fichier_corrompu_donne_une_liste_vide(tmp_path):
    fichier = tmp_path / "attente.json"
    fichier.write_text("{ ceci n'est pas du JSON", encoding="utf-8")

    assert attente.charger(str(fichier)) == {}


def test_vider_efface_la_liste(tmp_path):
    fichier = str(tmp_path / "attente.json")
    attente.mettre_a_jour(fichier, {"13001_1": "fichier audit introuvable"}, [])
    assert attente.charger(fichier) != {}

    attente.vider(fichier)

    assert attente.charger(fichier) == {}


# ── Mise à jour ────────────────────────────────────────────────────
def test_les_echecs_sont_ajoutes_avec_leur_raison(tmp_path):
    fichier = str(tmp_path / "attente.json")

    liste = attente.mettre_a_jour(
        fichier, {"13001_1": "fichier audit introuvable", "75020_4": "ID erreur introuvable"}, [],
    )

    assert set(liste) == {"13001_1", "75020_4"}
    assert liste["13001_1"]["raison"] == "fichier audit introuvable"
    assert liste["13001_1"]["date"]  # une date a été enregistrée


def test_les_reussis_sont_retires(tmp_path):
    fichier = str(tmp_path / "attente.json")
    attente.mettre_a_jour(fichier, {"13001_1": "introuvable", "75020_4": "introuvable"}, [])

    liste = attente.mettre_a_jour(fichier, {}, ["13001_1"])

    assert set(liste) == {"75020_4"}


def test_un_nouvel_echec_remplace_l_ancien(tmp_path):
    fichier = str(tmp_path / "attente.json")
    attente.mettre_a_jour(fichier, {"13001_1": "ancienne raison"}, [])

    liste = attente.mettre_a_jour(fichier, {"13001_1": "nouvelle raison"}, [])

    assert liste["13001_1"]["raison"] == "nouvelle raison"


def test_persiste_sur_disque_entre_deux_appels(tmp_path):
    """Un second appel à charger() relit bien ce qu'un appel précédent a écrit."""
    fichier = str(tmp_path / "attente.json")
    attente.mettre_a_jour(fichier, {"13001_1": "introuvable"}, [])

    relu = attente.charger(fichier)

    assert list(relu) == ["13001_1"]


# ── Répartition échecs / réussites d'un traitement ────────────────
def _resultat(commune, statut, message="", ids_absents=None, ids_non_verifies=None):
    return ResultatCommune(
        commune=commune, statut=statut, message=message, ids_absents=ids_absents or [],
        ids_non_verifies=ids_non_verifies or [],
    )


def test_fichier_introuvable_met_tous_ses_ids_en_echec():
    groupes = {"13001": ["13001_1", "13001_2"]}
    resultats = [_resultat("13001", Statut.INTROUVABLE)]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)

    assert echecs == {
        "13001_1": "fichier audit introuvable", "13001_2": "fichier audit introuvable",
    }
    assert reussis == []


def test_erreur_met_tous_ses_ids_en_echec_avec_le_message():
    groupes = {"13001": ["13001_1"]}
    resultats = [_resultat("13001", Statut.ERREUR, message="colonne « X » introuvable")]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)

    assert echecs == {"13001_1": "colonne « X » introuvable"}
    assert reussis == []


def test_succes_separe_les_ids_absents_des_reussis():
    groupes = {"13001": ["13001_1", "13001_2", "13001_3"]}
    resultats = [_resultat("13001", Statut.SUCCES, ids_absents=["13001_2"])]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)

    assert echecs == {"13001_2": "ID erreur introuvable dans le fichier"}
    assert sorted(reussis) == ["13001_1", "13001_3"]


def test_ecriture_non_verifiee_est_un_echec():
    """Écrit via Excel (services.excel_com) mais pas confirmé après enregistrement."""
    groupes = {"13001": ["13001_1", "13001_2"]}
    resultats = [_resultat("13001", Statut.SUCCES, ids_non_verifies=["13001_1"])]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)

    assert echecs == {"13001_1": "écriture non confirmée dans SharePoint après enregistrement"}
    assert reussis == ["13001_2"]


def test_plusieurs_communes_combinees():
    groupes = {"13001": ["13001_1"], "75020": ["75020_4"], "99999": ["99999_1"]}
    resultats = [
        _resultat("13001", Statut.SUCCES),
        _resultat("75020", Statut.INTROUVABLE),
        _resultat("99999", Statut.SUCCES, ids_absents=["99999_1"]),
    ]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)

    assert echecs == {
        "75020_4": "fichier audit introuvable",
        "99999_1": "ID erreur introuvable dans le fichier",
    }
    assert reussis == ["13001_1"]


def test_bout_en_bout_echec_puis_succes(tmp_path):
    """Un ID en échec, rejoué avec succès, disparaît de la liste."""
    fichier = str(tmp_path / "attente.json")
    groupes = {"13001": ["13001_1"]}

    echecs, reussis = attente.calculer_echecs_et_reussites(
        groupes, [_resultat("13001", Statut.INTROUVABLE)],
    )
    liste = attente.mettre_a_jour(fichier, echecs, reussis)
    assert "13001_1" in liste

    echecs2, reussis2 = attente.calculer_echecs_et_reussites(
        groupes, [_resultat("13001", Statut.SUCCES)],
    )
    liste2 = attente.mettre_a_jour(fichier, echecs2, reussis2)

    assert liste2 == {}
