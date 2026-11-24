"""Extraction et regroupement des ID erreur saisis par l'utilisateur."""

from sna.services.identifiants import (
    codes_sans_id, extraire_ids, grouper_par_commune, groupes_depuis_texte,
    insee_de, normaliser_id,
)


# ── Extraction ────────────────────────────────────────────────────
def test_extrait_les_ids_au_format_insee_numero():
    assert extraire_ids("13001_1 75020_27") == ["13001_1", "75020_27"]


def test_ids_corses_mis_en_majuscules():
    assert extraire_ids("2a004_3 2B033_10") == ["2A004_3", "2B033_10"]


def test_supprime_doublons_en_preservant_ordre():
    assert extraire_ids("75020_1 13001_2 75020_1") == ["75020_1", "13001_2"]


def test_extrait_dans_du_texte_libre():
    texte = "Erreurs à corriger : 13001_1, 13001_2 ; puis 2A004_7."
    assert extraire_ids(texte) == ["13001_1", "13001_2", "2A004_7"]


def test_ignore_les_codes_sans_numero_d_erreur():
    assert extraire_ids("13001 75020") == []


def test_ignore_les_formats_hors_norme():
    assert extraire_ids("bonjour 123_1 1234567_1 abcde_2") == []


# ── Codes incomplets ──────────────────────────────────────────────
def test_signale_les_codes_insee_sans_numero():
    assert codes_sans_id("13001_1 75020 2a004") == ["75020", "2A004"]


def test_un_id_complet_n_est_pas_signale_comme_incomplet():
    assert codes_sans_id("13001_1 13001_2") == []


# ── Décomposition ─────────────────────────────────────────────────
def test_insee_extrait_du_prefixe():
    assert insee_de("13001_27") == "13001"
    assert insee_de(" 2a004_3 ") == "2A004"


def test_normalisation_tolere_espaces_et_casse():
    assert normaliser_id(" 2a004_3 ") == "2A004_3"
    assert normaliser_id("13001 _ 1") == "13001_1"
    assert normaliser_id(None) == ""


# ── Regroupement ──────────────────────────────────────────────────
def test_groupe_les_ids_par_commune_sans_doublon():
    ids = ["13001_1", "75020_4", "13001_2", "13001_1"]

    assert grouper_par_commune(ids) == {
        "13001": ["13001_1", "13001_2"],
        "75020": ["75020_4"],
    }


def test_ordre_d_apparition_preserve():
    groupes = grouper_par_commune(["75020_1", "13001_1"])

    assert list(groupes) == ["75020", "13001"]


def test_raccourci_depuis_le_texte_saisi():
    assert groupes_depuis_texte("13001_1 ; 13001_2 ; 75020_9") == {
        "13001": ["13001_1", "13001_2"],
        "75020": ["75020_9"],
    }


def test_texte_sans_id_donne_un_dictionnaire_vide():
    assert groupes_depuis_texte("aucun identifiant ici") == {}
