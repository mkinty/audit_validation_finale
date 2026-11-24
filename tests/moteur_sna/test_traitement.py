"""Orchestration d'un lot de communes déduit des ID erreur."""

from openpyxl import Workbook, load_workbook

from sna.models.base import Statut
from sna.models.regles import Action, RegleModification, ReglesConfig
from sna.services.excel import BilanFichier
from sna.services.traitement import traiter_communes

VALEUR = "Vérification RT impossible nb log"

CONFIG = ReglesConfig(
    regles=[RegleModification(
        code="R1", libelle="statut", action=Action(colonne="Z", valeur=VALEUR)
    )],
    colonne_id="ID erreur",
)


def _preparer_commune(racine, commune, ids, valeurs_z=None):
    """Crée ``<racine>/Dep<XX>/<insee>/audit_<insee>.xlsx`` avec une ligne par ID."""
    dossier = racine / f"Dep{commune[:2]}" / commune
    dossier.mkdir(parents=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    ws["C1"] = "ID erreur"
    ws["Z1"] = "Audit - Street View"
    valeurs_z = valeurs_z or ["à revoir"] * len(ids)
    for i, (identifiant, valeur) in enumerate(zip(ids, valeurs_z), start=2):
        ws[f"C{i}"] = identifiant
        ws[f"Z{i}"] = valeur
    chemin = dossier / f"audit_{commune}.xlsx"
    wb.save(chemin)
    return chemin


def test_traite_une_commune_par_groupe(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1", "13001_2", "13001_3"])
    _preparer_commune(tmp_path, "75020", ["75020_5"])

    resultats = traiter_communes(
        {"13001": ["13001_1", "13001_3"], "75020": ["75020_5"]},
        CONFIG, str(tmp_path),
    )

    assert [r.commune for r in resultats] == ["13001", "75020"]
    assert resultats[0].nb_modifications == 2
    assert resultats[1].nb_modifications == 1
    assert all(r.statut is Statut.SUCCES for r in resultats)


def test_ids_absents_du_fichier_sont_signales(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])

    resultat = traiter_communes(
        {"13001": ["13001_1", "13001_42"]}, CONFIG, str(tmp_path)
    )[0]

    assert resultat.statut is Statut.SUCCES
    assert resultat.ids_absents == ["13001_42"]
    assert resultat.nb_ids_demandes == 2
    assert resultat.resume_ids() == "1/2 ID erreur trouvés"


def test_commune_sans_fichier(tmp_path):
    resultat = traiter_communes({"99999": ["99999_1"]}, CONFIG, str(tmp_path))[0]

    assert resultat.statut is Statut.INTROUVABLE
    assert resultat.nb_ids_demandes == 1


def test_detail_par_regle(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])

    resultat = traiter_communes({"13001": ["13001_1"]}, CONFIG, str(tmp_path))[0]

    assert resultat.detail == {"R1": 1}
    assert resultat.resume_detail() == "R1: 1"


def test_colonne_absente_donne_une_erreur_isolee(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    _preparer_commune(tmp_path, "75020", ["75020_1"])
    config = ReglesConfig(
        regles=[RegleModification(
            code="R1", action=Action(colonne="Colonne absente", valeur="x")
        )],
        colonne_id="ID erreur",
    )

    resultats = traiter_communes(
        {"13001": ["13001_1"], "75020": ["75020_1"]}, config, str(tmp_path)
    )

    assert all(r.statut is Statut.ERREUR for r in resultats)
    assert "introuvable" in resultats[0].message


def test_simulation_n_ecrit_rien(tmp_path):
    fichier = _preparer_commune(tmp_path, "13001", ["13001_1"])

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), simuler=True
    )[0]

    assert resultat.nb_modifications == 1
    assert load_workbook(fichier).active["Z2"].value == "à revoir"


def test_callbacks_sont_appeles(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    progressions, resultats = [], []

    traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path),
        on_progression=lambda f, t, lib: progressions.append((f, t)),
        on_resultat=resultats.append,
    )

    assert (0, 1) in progressions and (1, 1) in progressions
    assert len(resultats) == 1 and resultats[0].commune == "13001"


def test_les_etapes_sont_transmises_avec_la_commune(tmp_path):
    """Le journal doit s'alimenter pendant le traitement, pas seulement après."""
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    etapes = []

    traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path),
        on_etape=lambda commune, message: etapes.append((commune, message)),
    )

    assert etapes[0] == ("13001", "recherche du fichier audit…")
    assert any("audit_13001.xlsx" in message for _, message in etapes)
    assert all(commune == "13001" for commune, _ in etapes)


def test_duree_mesuree_par_commune(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])

    resultat = traiter_communes({"13001": ["13001_1"]}, CONFIG, str(tmp_path))[0]

    assert resultat.duree > 0
    assert resultat.resume_duree().endswith(" s")


def test_progression_annoncee_avec_le_nombre_d_ids(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1", "13001_2"])
    libelles = []

    traiter_communes(
        {"13001": ["13001_1", "13001_2"]}, CONFIG, str(tmp_path),
        on_progression=lambda fait, total, libelle: libelles.append(libelle),
    )

    assert any("2 ID erreur" in libelle for libelle in libelles)


# ── Écriture via Excel (session_excel) ────────────────────────────
class _FausseSession:
    """Remplace services.excel_com.SessionExcel : pas besoin d'un vrai Excel ici."""

    def __init__(self, bilan: BilanFichier | None = None):
        self.bilan = bilan
        self.appels: list[tuple] = []

    def appliquer_regles(self, chemin_fichier, config, ids_erreur, on_etape=None):
        self.appels.append((chemin_fichier, tuple(ids_erreur)))
        return self.bilan if self.bilan is not None else BilanFichier(comptes={})


def test_session_excel_est_utilisee_quand_une_ligne_correspond(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    session = _FausseSession(BilanFichier(comptes={"R1": 1}, ids_trouves={"13001_1"}, nb_lignes=1))

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), session_excel=session,
    )[0]

    assert len(session.appels) == 1
    assert resultat.statut is Statut.SUCCES and resultat.nb_modifications == 1


def test_session_excel_non_sollicitee_si_aucun_id_local(tmp_path):
    """Pas d'ID correspondant localement → inutile d'ouvrir Excel pour rien."""
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    session = _FausseSession()

    resultat = traiter_communes(
        {"13001": ["13001_9"]}, CONFIG, str(tmp_path), session_excel=session,
    )[0]

    assert session.appels == []
    assert resultat.statut is Statut.SUCCES
    assert resultat.ids_absents == ["13001_9"]


def test_session_excel_ignoree_en_simulation(tmp_path):
    """En simulation, on retombe sur la réécriture directe (rien n'est écrit)."""
    fichier = _preparer_commune(tmp_path, "13001", ["13001_1"])
    session = _FausseSession()

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), simuler=True, session_excel=session,
    )[0]

    assert session.appels == []
    assert resultat.nb_modifications == 1
    assert load_workbook(fichier).active["Z2"].value == "à revoir"


def test_classeur_verrouille_devient_une_erreur_avec_le_nom(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    session = _FausseSession(BilanFichier(comptes={}, verrouille="Jean Dupont"))

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), session_excel=session,
    )[0]

    assert resultat.statut is Statut.ERREUR
    assert "Jean Dupont" in resultat.message


def test_copie_en_conflit_bloque_la_commune_sans_rien_ecrire(tmp_path):
    """Une copie de conflit OneDrive/SharePoint fait échouer la commune, sans
    jamais ouvrir ni le fichier principal ni Excel."""
    fichier = _preparer_commune(tmp_path, "13001", ["13001_1"])
    conflit = fichier.with_name("audit_13001_DESKTOP-AB12CD3.xlsx")
    conflit.write_bytes(fichier.read_bytes())
    session = _FausseSession()

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), session_excel=session,
    )[0]

    assert resultat.statut is Statut.ERREUR
    assert "conflit" in resultat.message
    assert "DESKTOP-AB12CD3" in resultat.message
    assert session.appels == []
    assert load_workbook(fichier).active["Z2"].value == "à revoir"


def test_ids_non_verifies_sont_reportes_dans_le_resultat(tmp_path):
    _preparer_commune(tmp_path, "13001", ["13001_1"])
    session = _FausseSession(BilanFichier(
        comptes={"R1": 0}, ids_trouves={"13001_1"}, nb_lignes=1,
        ids_non_verifies=["13001_1"],
    ))

    resultat = traiter_communes(
        {"13001": ["13001_1"]}, CONFIG, str(tmp_path), session_excel=session,
    )[0]

    assert resultat.statut is Statut.SUCCES
    assert resultat.ids_non_verifies == ["13001_1"]
