import os

from avf.services.detection import (
    departement_de, lister_communes, parser_communes, verifier_communes,
)


def test_lister_communes_de_toute_la_racine(racine):
    communes = lister_communes(racine)
    assert [c.code for c in communes] == ["13001", "13002", "13003", "2A004", "75020", "75021"]
    assert [c.a_fichier for c in communes] == [True, True, False, True, True, True]
    assert communes[3].departement == "2A"
    assert os.path.isdir(communes[0].dossier)


def test_dossiers_hors_dep_ignores(racine):
    os.makedirs(os.path.join(racine, "Divers", "99999"))
    assert "99999" not in [c.code for c in lister_communes(racine)]


def test_racine_absente(tmp_path):
    assert lister_communes(str(tmp_path / "nulle_part")) == []


def test_casse_du_prefixe_dep(tmp_path):
    os.makedirs(tmp_path / "DEP13" / "13001")
    assert [c.code for c in lister_communes(str(tmp_path))] == ["13001"]


def test_parser_communes():
    assert parser_communes("13001, 2a004\n13001 ;75020") == ["13001", "2A004", "75020"]


def test_departement_de():
    assert departement_de("2a004") == "2A"
    assert departement_de("13001") == "13"


def test_verifier_communes_deduit_le_dossier_dep(racine):
    v = verifier_communes(racine, ["13001", "99999", "2A004", "13003"])
    assert [c.code for c in v.trouvees] == ["13001", "2A004", "13003"]
    assert [c.departement for c in v.trouvees] == ["13", "2A", "13"]
    assert [c.a_fichier for c in v.trouvees] == [True, True, False]
    assert v.absentes == ["99999"]
