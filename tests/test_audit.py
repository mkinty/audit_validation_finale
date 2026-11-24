from avf.models.audit import AuditCommune, IdErreurAudit, LigneAudit, RapportAudit, StatutAudit
from avf.services.audit import auditer_commune, auditer_communes, regrouper_par_id

TOUTES = ["13001", "13002", "13003", "2A004", "75020", "75021", "99999"]


def test_commune_complete_et_comptages(racine):
    c = auditer_commune("13001", racine)
    assert c.statut is StatutAudit.OK
    assert c.departement == "13"
    assert c.nb_ids == 5
    assert c.nb_renseignes == 2          # 13001_1, 13001_4
    assert c.nb_vides == 3               # 13001_2, 13001_3 (espaces), 13001_5 (1 ligne sur 2)
    assert c.ids_vides == ["13001_2", "13001_3", "13001_5"]
    assert c.nb_lignes == 6
    assert c.nb_doublons == 1
    assert c.taux == 2 / 5
    assert not c.complete


def test_doublon_reference_premiere_ligne_non_validee(racine):
    c = auditer_commune("13001", racine)
    doublon = next(i for i in c.ids if i.id_erreur == "13001_5")
    assert doublon.doublon
    assert doublon.reference.ligne == 7


def test_commune_complete(racine):
    c = auditer_commune("13002", racine)
    assert c.complete and c.nb_vides == 0 and c.taux == 1


def test_ids_normalises_comme_le_moteur(racine):
    c = auditer_commune("2a004", racine)
    assert c.commune == "2A004" and c.departement == "2A"
    assert c.ids_vides == ["2A004_1", "2A004_2"]


def test_sans_fichier(racine):
    c = auditer_commune("13003", racine)
    assert c.statut is StatutAudit.INTROUVABLE
    assert c.ids_vides == []


def test_dossier_absent(racine):
    assert auditer_commune("99999", racine).statut is StatutAudit.INTROUVABLE


def test_colonne_absente_en_erreur(racine):
    c = auditer_commune("75020", racine)
    assert c.statut is StatutAudit.ERREUR
    assert "Validation finale" in c.message


def test_conflit_signale(racine):
    c = auditer_commune("75021", racine)
    assert c.statut is StatutAudit.OK
    assert any("conflit" in a for a in c.avertissements)
    assert c.nb_vides == 1


def test_auditer_communes_rapport(racine):
    progression, etapes, resultats = [], [], []
    rapport = auditer_communes(
        TOUTES + ["13001", " "], racine,
        on_progression=lambda f, t, l: progression.append((f, t)),
        on_etape=lambda c, m: etapes.append(c),
        on_resultat=resultats.append,
    )
    assert [c.commune for c in rapport.communes] == TOUTES   # doublons et vides retirés
    assert len(resultats) == len(TOUTES)
    assert progression[0] == (0, 7) and progression[-1] == (7, 7)
    assert set(etapes) >= {"13001", "2A004"}
    assert rapport.nb_ids == 5 + 2 + 2 + 1
    assert rapport.nb_renseignes == 4
    assert rapport.nb_vides == 6
    assert [c.commune for c in rapport.en_echec] == ["13003", "75020", "99999"]
    assert [c.commune for c in rapport.communes_completes] == ["13002"]
    assert rapport.ids_vides_par_commune() == {
        "13001": ["13001_2", "13001_3", "13001_5"],
        "2A004": ["2A004_1", "2A004_2"],
        "75021": ["75021_1"],
    }
    assert len(rapport.tous_ids_vides()) == 6
    deps = rapport.par_departement()
    assert list(deps) == ["13", "2A", "75", "99"]
    assert deps["13"] == {"communes": 3, "ids": 7, "renseignes": 4, "vides": 3, "erreurs": 1}


def test_arret_demande(racine):
    appels = iter([False, True])
    rapport = auditer_communes(TOUTES, racine, doit_arreter=lambda: next(appels))
    assert [c.commune for c in rapport.communes] == ["13001"]


def test_regrouper_par_id_ignore_les_vides():
    groupes = regrouper_par_id([
        LigneAudit(2, " 1_1 "), LigneAudit(3, ""), LigneAudit(4, "1_1", validation="OK"),
    ])
    assert [(g.id_erreur, len(g.lignes)) for g in groupes] == [("1_1", 2)]
    assert not groupes[0].renseigne


def test_rapport_remplacer_garde_ordre():
    a, b, c = (AuditCommune(x, x[:2]) for x in ("13001", "13002", "13003"))
    rapport = RapportAudit(racine="r", communes=[a, b])
    nouveau_b = AuditCommune("13002", "13", ids=[IdErreurAudit("13002_1", [LigneAudit(2, "13002_1")])])
    rapport.remplacer([nouveau_b, c])
    assert [x.commune for x in rapport.communes] == ["13001", "13002", "13003"]
    assert rapport.communes[1] is nouveau_b


def test_rapport_vide():
    r = RapportAudit(racine="r")
    assert r.taux is None and r.nb_ids == 0 and r.tous_ids_vides() == []
