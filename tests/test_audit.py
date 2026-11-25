from avf.models.audit import (
    NON_VALIDES, TOUTES as TOUTES_LIGNES, VALIDEES, Anomalie, AuditCommune, IdErreurAudit, LigneAudit,
    RapportAudit, StatutAudit,
)
from avf.services.audit import auditer_commune, auditer_communes, regrouper_par_id

TOUTES = ["13001", "13002", "13003", "2A004", "75020", "75021", "99999"]


def _motifs(c, ligne):
    return next(l for l in c.lignes if l.ligne == ligne).motifs


def test_lignes_validees_et_non_valides(racine):
    c = auditer_commune("13001", racine)
    assert c.statut is StatutAudit.OK and c.departement == "13"
    assert c.nb_lignes == 6 and c.nb_lignes_valides == 1 and c.nb_lignes_non_valides == 5
    assert [l.ligne for l in c.lignes_de(VALIDEES)] == [2]
    assert [l.ligne for l in c.lignes_de(NON_VALIDES)] == [3, 4, 5, 6, 7]
    assert len(c.lignes_de(TOUTES_LIGNES)) == 6
    assert c.taux == 1 / 6
    assert not c.complete


def test_ids_valides_si_toutes_leurs_lignes_le_sont(racine):
    c = auditer_commune("13001", racine)
    assert c.nb_ids == 5 and c.nb_ids_valides == 1 and c.nb_ids_non_valides == 4
    assert c.ids_non_valides == ["13001_2", "13001_3", "13001_4", "13001_5"]
    assert c.ids_de(VALIDEES) == ["13001_1"]
    assert c.nb_doublons == 1


def test_motifs_par_critere(racine):
    c = auditer_commune("13001", racine)
    assert _motifs(c, 3) == "Audit vide · Date audit vide · Auditeur vide · Validation finale vide"
    assert _motifs(c, 4) == "Date audit vide · Validation finale vide"   # KO accepté, VF « ␣␣␣ » vide
    assert _motifs(c, 5) == "Validation finale « Validé » ≠ OK"
    assert _motifs(c, 2) == ""
    assert c.par_critere() == {"audit": 3, "date_audit": 4, "auditeur": 3, "validation": 4}


def test_doublon_reference_premiere_ligne_non_valide(racine):
    c = auditer_commune("13001", racine)
    doublon = next(i for i in c.ids if i.id_erreur == "13001_5")
    assert doublon.doublon and not doublon.valide
    assert doublon.reference.ligne == 6
    assert [l.ligne for l in doublon.lignes_non_valides] == [6, 7]


def test_commune_complete_date_texte_et_casse(racine):
    c = auditer_commune("13002", racine)
    assert c.complete and c.nb_lignes_non_valides == 0 and c.taux == 1
    assert c.lignes[1].date_audit == "15/02/2026" and c.lignes[1].valide


def test_ids_normalises_comme_le_moteur(racine):
    c = auditer_commune("2a004", racine)
    assert c.commune == "2A004" and c.departement == "2A"
    assert c.ids_non_valides == ["2A004_1", "2A004_2"]


def test_sans_fichier(racine):
    c = auditer_commune("13003", racine)
    assert c.statut is StatutAudit.INTROUVABLE
    assert c.ids_non_valides == [] and c.ids_de(VALIDEES) == []


def test_dossier_absent(racine):
    assert auditer_commune("99999", racine).statut is StatutAudit.INTROUVABLE


def test_colonne_controlee_absente_en_erreur(racine):
    c = auditer_commune("75020", racine)
    assert c.statut is StatutAudit.ERREUR
    assert "Validation finale" in c.message and "obligatoires" in c.message


def test_conflit_signale(racine):
    c = auditer_commune("75021", racine)
    assert c.statut is StatutAudit.OK
    assert any("conflit" in a for a in c.avertissements)
    assert c.nb_ids_non_valides == 1


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
    assert (rapport.nb_lignes, rapport.nb_lignes_valides, rapport.nb_lignes_non_valides) == (11, 3, 8)
    assert (rapport.nb_ids, rapport.nb_ids_valides, rapport.nb_ids_non_valides) == (10, 3, 7)
    assert rapport.taux == 3 / 11
    assert [c.commune for c in rapport.en_echec] == ["13003", "75020", "99999"]
    assert [c.commune for c in rapport.communes_completes] == ["13002"]
    assert [c.commune for c in rapport.communes_a_corriger] == ["13001", "2A004", "75021"]
    assert rapport.ids_non_valides_par_commune() == {
        "13001": ["13001_2", "13001_3", "13001_4", "13001_5"],
        "2A004": ["2A004_1", "2A004_2"],
        "75021": ["75021_1"],
    }
    assert len(rapport.ids_de(NON_VALIDES)) == 7
    assert rapport.ids_de(VALIDEES) == ["13001_1", "13002_1", "13002_2"]
    assert rapport.par_critere() == {"audit": 6, "date_audit": 7, "auditeur": 6, "validation": 7}
    deps = rapport.par_departement()
    assert list(deps) == ["13", "2A", "75", "99"]
    assert deps["13"] == {"communes": 3, "ids": 7, "lignes": 8, "valides": 3, "non_valides": 5, "erreurs": 1}


def test_arret_demande(racine):
    appels = iter([False, True])
    rapport = auditer_communes(TOUTES, racine, doit_arreter=lambda: next(appels))
    assert [c.commune for c in rapport.communes] == ["13001"]


def test_regrouper_par_id_ignore_les_vides():
    groupes = regrouper_par_id([
        LigneAudit(2, " 1_1 "), LigneAudit(3, ""),
        LigneAudit(4, "1_1", anomalies=[Anomalie("audit", "Audit vide")]),
    ])
    assert [(g.id_erreur, len(g.lignes)) for g in groupes] == [("1_1", 2)]
    assert not groupes[0].valide


def test_rapport_remplacer_garde_ordre():
    a, b, c = (AuditCommune(x, x[:2]) for x in ("13001", "13002", "13003"))
    rapport = RapportAudit(racine="r", communes=[a, b])
    nouveau_b = AuditCommune("13002", "13", ids=[IdErreurAudit("13002_1", [LigneAudit(2, "13002_1")])])
    rapport.remplacer([nouveau_b, c])
    assert [x.commune for x in rapport.communes] == ["13001", "13002", "13003"]
    assert rapport.communes[1] is nouveau_b


def test_rapport_vide():
    r = RapportAudit(racine="r")
    assert r.taux is None and r.nb_ids == 0 and r.ids_de(NON_VALIDES) == []


def _signature(rapport):
    return [(c.commune, c.statut, c.nb_ids, c.nb_lignes_non_valides, c.ids_non_valides, c.message)
            for c in rapport.communes]


def test_audit_parallele_identique_au_sequentiel(racine):
    progression, resultats = [], []
    sequentiel = auditer_communes(TOUTES, racine)
    parallele = auditer_communes(
        TOUTES, racine, paralleles=4,
        on_progression=lambda f, t, l: progression.append((f, t)), on_resultat=resultats.append,
    )
    assert _signature(parallele) == _signature(sequentiel)
    assert [c.commune for c in parallele.communes] == TOUTES   # ordre demandé
    assert sorted(r.commune for r in resultats) == sorted(TOUTES)
    assert progression[0] == (0, 7) and progression[-1] == (7, 7)


def test_arret_en_parallele(racine):
    import threading
    arret = threading.Event()
    rapport = auditer_communes(TOUTES, racine, paralleles=2, doit_arreter=arret.is_set,
                               on_resultat=lambda r: arret.set())
    assert 1 <= len(rapport.communes) < len(TOUTES)
    codes = [c.commune for c in rapport.communes]
    assert codes == sorted(codes, key=TOUTES.index)
