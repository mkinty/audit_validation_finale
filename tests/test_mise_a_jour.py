"""Mise à jour de bout en bout : audit → moteur sna (complete_excel_files) → ré-audit."""

from datetime import date, datetime
import json
import os

import pytest
from openpyxl import load_workbook

from avf.services.audit import auditer_communes
from avf.services.mise_a_jour import groupes_depuis_rapport, mettre_a_jour, preparer_regles
from sna.models.base import Statut
from sna.models.regles import Action, Condition, RegleModification, ReglesConfig, SourceValeur
from sna.services.excel import appliquer_regles

COMMUNES = ["13001", "13002", "2A004", "75021"]


def _valeurs(chemin, feuille="Audit"):
    wb = load_workbook(chemin)
    ws = next(wb[n] for n in wb.sheetnames if "".join(n.split()).lower() == feuille.lower())
    entetes = [c.value for c in ws[1]]
    lignes = [dict(zip(entetes, [c.value for c in r])) for r in ws.iter_rows(min_row=2)]
    wb.close()
    return lignes


# ── Règles par défaut : celles de complete_excel_files ────────────
def test_regles_par_defaut_identiques_a_complete_excel_files():
    regles = ReglesConfig.defaut()
    assert [(r.colonne_cible, r.action.source, r.action.valeur) for r in regles.regles] == [
        ("Audit", SourceValeur.VALEUR, "OK"),
        ("Date audit", SourceValeur.VALEUR, date.today().strftime("%d/%m/%Y")),
        ("Auditeur", SourceValeur.VALEUR, "Validation en masse"),
        ("Validation finale", SourceValeur.VALEUR, "OK"),
    ]
    assert regles.colonne_id == "ID erreur"


# ── preparer_regles ───────────────────────────────────────────────
def _regles_datees(valeur="01/01/2020"):
    regles = ReglesConfig.defaut()
    regles.regles[1].action = Action("Date audit", SourceValeur.VALEUR, valeur)
    return regles


def test_preparer_regles_date_du_jour():
    regles = _regles_datees()
    prete = preparer_regles(regles, "Date audit", True, aujourd_hui=date(2026, 11, 24))
    assert prete.regles[1].action.valeur == "24/11/2026"
    assert regles.regles[1].action.valeur == "01/01/2020"  # original intact
    assert [r.action.valeur for r in prete.regles if r.code != "R2"] == ["OK", "Validation en masse", "OK"]


def test_preparer_regles_sans_option():
    prete = preparer_regles(_regles_datees(), "Date audit", False)
    assert prete.regles[1].action.valeur == "01/01/2020"


def test_preparer_regles_casse_et_espaces_de_colonne():
    prete = preparer_regles(_regles_datees(), " date AUDIT ", True, aujourd_hui=date(2026, 1, 2))
    assert prete.regles[1].action.valeur == "02/01/2026"


def test_preparer_regles_ne_touche_pas_recopie_ni_effacement():
    regles = ReglesConfig(regles=[
        RegleModification("R1", action=Action("Date audit", SourceValeur.COLONNE, "Autre date")),
        RegleModification("R2", action=Action("Date audit", SourceValeur.VIDE, "")),
    ])
    prete = preparer_regles(regles, "Date audit", True)
    assert [r.action.valeur for r in prete.regles] == ["Autre date", ""]


# ── Groupes ───────────────────────────────────────────────────────
def test_groupes_depuis_rapport(racine):
    rapport = auditer_communes(COMMUNES, racine)
    assert groupes_depuis_rapport(rapport) == {
        "13001": ["13001_2", "13001_3", "13001_4", "13001_5"],
        "2A004": ["2A004_1", "2A004_2"],
        "75021": ["75021_1"],
    }
    assert list(groupes_depuis_rapport(rapport, ["2A004", "13002"])) == ["2A004"]


# ── De bout en bout ───────────────────────────────────────────────
def test_mise_a_jour_corrige_les_ids_non_valides(racine, tmp_path):
    rapport = auditer_communes(COMMUNES, racine)
    groupes = groupes_depuis_rapport(rapport)
    regles = preparer_regles(ReglesConfig.defaut(), "Date audit", True, aujourd_hui=date(2026, 11, 24))
    attente_json = str(tmp_path / "attente.json")
    etapes, resultats = [], []

    bilan = mettre_a_jour(
        groupes, regles, racine, fichier_attente=attente_json,
        on_etape=lambda c, m: etapes.append(c), on_resultat=resultats.append,
    )

    par_commune = {r.commune: r for r in bilan.resultats}
    assert par_commune["13001"].statut is Statut.SUCCES
    assert par_commune["2A004"].statut is Statut.SUCCES
    # Copie de conflit : refus, exactement comme complete_excel_files
    assert par_commune["75021"].statut is Statut.ERREUR
    assert "conflit" in par_commune["75021"].message
    assert bilan.echecs == {"75021_1": par_commune["75021"].message}
    assert sorted(bilan.reussis) == ["13001_2", "13001_3", "13001_4", "13001_5", "2A004_1", "2A004_2"]
    assert bilan.nb_en_attente == 1
    assert list(json.load(open(attente_json, encoding="utf-8"))) == ["75021_1"]
    assert len(resultats) == 3 and etapes

    lignes = _valeurs(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"))
    par_id = {}
    for l in lignes:
        par_id.setdefault(l["ID erreur"], []).append(l)
    # ID validé : inchangé
    assert par_id["13001_1"][0]["Auditeur"] == "Alice"
    assert par_id["13001_1"][0]["Date audit"] == datetime(2026, 3, 12)
    # ID non valides (dont 13001_4, Validation finale « Validé ») : les 4 colonnes
    # réécrites, sur toutes leurs lignes
    for ident in ("13001_2", "13001_3", "13001_4", "13001_5"):
        for l in par_id[ident]:
            assert l["Audit"] == "OK"
            assert l["Date audit"] == datetime(2026, 11, 24)   # vraie date Excel
            assert l["Auditeur"] == "Validation en masse"
            assert l["Validation finale"] == "OK"
    # La ligne sans ID n'est pas touchée
    assert lignes[-1]["Audit"] is None and lignes[-1]["Validation finale"] is None

    # L'onglet « AUDIT  » de 2A004 a été modifié, pas le premier onglet
    assert all(l["Validation finale"] == "OK" for l in _valeurs(
        os.path.join(racine, "Dep2A", "2A004", "audit_2A004.xlsx"), "audit"))

    # Ré-audit : plus aucun ID non valide hors conflit
    apres = auditer_communes(COMMUNES, racine)
    assert apres.ids_non_valides_par_commune() == {"75021": ["75021_1"]}
    assert apres.commune("13001").complete


def test_simulation_n_ecrit_rien(racine, tmp_path):
    chemin = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    avant = os.path.getmtime(chemin), open(chemin, "rb").read()
    rapport = auditer_communes(["13001"], racine)
    attente_json = str(tmp_path / "attente.json")
    bilan = mettre_a_jour(
        groupes_depuis_rapport(rapport), ReglesConfig.defaut(), racine,
        simuler=True, fichier_attente=attente_json,
    )
    # 13001_2 : 4 cellules · 13001_3 : 4 (KO → OK) · 13001_4 : 3 (Audit déjà OK)
    # · 13001_5 : 3 (Validation finale déjà OK) + 4
    assert bilan.simulation and bilan.nb_modifications == 18
    assert open(chemin, "rb").read() == avant[1]
    assert not os.path.exists(attente_json)
    assert auditer_communes(["13001"], racine).nb_ids_non_valides == 4


def test_meme_resultat_que_le_moteur_appele_directement(racine, tmp_path):
    """La mise à jour n'ajoute rien au moteur : mêmes cellules que appliquer_regles."""
    import shutil
    copie = str(tmp_path / "copie.xlsx")
    original = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    shutil.copy(original, copie)
    regles = preparer_regles(ReglesConfig.defaut(), "Date audit", True)
    ids = groupes_depuis_rapport(auditer_communes(["13001"], racine))["13001"]

    appliquer_regles(copie, regles, ids)
    mettre_a_jour({"13001": ids}, regles, racine)
    assert _valeurs(copie) == _valeurs(original)


def test_regles_personnalisees_avec_condition(racine):
    """Réglage utilisateur : ne poser « Audit » que s'il est vide, auditeur nommé."""
    regles = ReglesConfig(regles=[
        RegleModification("R1", action=Action("Audit", SourceValeur.VALEUR, "OK"),
                          conditions=[Condition("Audit", "vide")]),
        RegleModification("R2", action=Action("Auditeur", SourceValeur.VALEUR, "Zoé")),
        RegleModification("R3", action=Action("Validation finale", SourceValeur.VALEUR, "VALIDE")),
    ])
    rapport = auditer_communes(["13001"], racine)
    mettre_a_jour(groupes_depuis_rapport(rapport), regles, racine)
    lignes = _valeurs(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"))
    ligne_13001_2, ligne_13001_3 = lignes[1], lignes[2]
    assert ligne_13001_3["Audit"] == "KO"                 # non vide : conservé
    assert ligne_13001_3["Auditeur"] == "Zoé"
    assert ligne_13001_2["Audit"] == "OK"
    assert ligne_13001_2["Validation finale"] == "VALIDE"
    assert lignes[0]["Auditeur"] == "Alice"               # 13001_1 déjà validé : intact


class _SessionFactice:
    """Remplace SessionExcel : vérifie que le chemin « via Excel » est emprunté."""

    def __init__(self):
        self.appels = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def appliquer_regles(self, fichier, config, ids, on_etape=None):
        self.appels.append((os.path.basename(fichier), list(ids)))
        return appliquer_regles(fichier, config, ids, on_etape=on_etape)


def test_via_excel_utilise_la_session(racine):
    session = _SessionFactice()
    rapport = auditer_communes(["13001", "2A004"], racine)
    bilan = mettre_a_jour(
        groupes_depuis_rapport(rapport), ReglesConfig.defaut(), racine,
        via_excel=True, session_factory=lambda: session,
    )
    assert bilan.via_excel
    assert [a[0] for a in session.appels] == ["audit_13001.xlsx", "audit_2A004.xlsx"]
    assert bilan.nb_communes_ok == 2


def test_via_excel_ignore_en_simulation(racine):
    session = _SessionFactice()
    rapport = auditer_communes(["13001"], racine)
    bilan = mettre_a_jour(
        groupes_depuis_rapport(rapport), ReglesConfig.defaut(), racine,
        simuler=True, via_excel=True, session_factory=lambda: session,
    )
    assert not bilan.via_excel and session.appels == []


def test_aucun_groupe(racine):
    bilan = mettre_a_jour({}, ReglesConfig.defaut(), racine)
    assert bilan.resultats == [] and bilan.nb_modifications == 0


def test_ids_absents_entre_audit_et_mise_a_jour(racine):
    """Un ID disparu du fichier entre l'audit et la mise à jour rejoint les échecs."""
    bilan = mettre_a_jour({"13001": ["13001_2", "13001_999"]}, ReglesConfig.defaut(), racine)
    assert bilan.echecs == {"13001_999": "ID erreur introuvable dans le fichier"}
    assert bilan.reussis == ["13001_2"]


# ── Écriture via une vraie instance Excel (co-édition SharePoint) ─
from sna.services.excel_com import excel_disponible  # noqa: E402


@pytest.mark.excel
@pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")
def test_mise_a_jour_via_excel_reel(racine):
    rapport = auditer_communes(["13001"], racine)
    regles = preparer_regles(ReglesConfig.defaut(), "Date audit", True, aujourd_hui=date(2026, 11, 24))
    bilan = mettre_a_jour(groupes_depuis_rapport(rapport), regles, racine, via_excel=True)
    assert bilan.via_excel
    resultat = bilan.resultats[0]
    assert resultat.statut is Statut.SUCCES, resultat.message
    assert resultat.ids_non_verifies == []
    lignes = _valeurs(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"))
    assert [l["Validation finale"] for l in lignes[:6]] == ["OK"] * 6
    assert lignes[1]["Date audit"] == datetime(2026, 11, 24)
    assert auditer_communes(["13001"], racine).commune("13001").complete


def test_mise_a_jour_parallele_identique(racine, tmp_path):
    import shutil
    copie = str(tmp_path / "copie")
    shutil.copytree(racine, copie)
    regles = preparer_regles(ReglesConfig.defaut(), aujourd_hui=date(2026, 11, 24))
    seq = mettre_a_jour(groupes_depuis_rapport(auditer_communes(COMMUNES, racine)), regles, racine)
    par = mettre_a_jour(groupes_depuis_rapport(auditer_communes(COMMUNES, copie)), regles, copie,
                        paralleles=4)
    assert [(r.commune, r.statut, r.nb_modifications) for r in par.resultats] == \
           [(r.commune, r.statut, r.nb_modifications) for r in seq.resultats]
    assert par.echecs.keys() == seq.echecs.keys() and sorted(par.reussis) == sorted(seq.reussis)
    assert not par.interrompu
    assert auditer_communes(COMMUNES, copie).ids_non_valides_par_commune() == {"75021": ["75021_1"]}


def test_arret_de_la_mise_a_jour(racine, tmp_path):
    attente_json = str(tmp_path / "attente.json")
    groupes = groupes_depuis_rapport(auditer_communes(COMMUNES, racine))
    appels = iter([False, True, True, True])
    bilan = mettre_a_jour(groupes, ReglesConfig.defaut(), racine, fichier_attente=attente_json,
                          doit_arreter=lambda: next(appels))
    assert bilan.communes_traitees == ["13001"] and bilan.interrompu
    assert bilan.nb_communes_demandees == 3
    # Les communes non commencées ne sont ni en échec ni réussies : intactes et hors attente
    assert bilan.echecs == {} and bilan.nb_en_attente == 0
    assert set(auditer_communes(COMMUNES, racine).ids_non_valides_par_commune()) == {"2A004", "75021"}
