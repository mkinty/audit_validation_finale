"""Écriture des mises à jour dans des fichiers ouverts par d'autres (services.ecriture)."""

from __future__ import annotations

import os

import pytest
from openpyxl import load_workbook

from avf.models.audit import RapportAudit
from avf.services.audit import auditer_communes
from avf.services.ecriture import SessionEcriture, proprietaire_verrou
from avf.services.mise_a_jour import groupes_depuis_rapport, mettre_a_jour
from avf.views import presentation as p
from sna.models.base import Statut
from sna.models.regles import ReglesConfig
from sna.services import excel_com
from sna.services.excel import BilanFichier, appliquer_regles

from .conftest import ecrire_audit, imm

IMM = "ID imm"


def _config():
    return ReglesConfig.from_dict({**ReglesConfig.defaut().to_dict(), "colonne_id": IMM})


def _verrou(chemin: str, nom: str = "Alice Martin") -> str:
    """Fichier verrou d'Excel ``~$…`` : longueur du nom, puis le nom (ANSI), complété à 54 octets."""
    dossier, base = os.path.split(chemin)
    verrou = os.path.join(dossier, "~$" + base[2:])
    donnees = bytes([len(nom)]) + nom.encode("cp1252")
    with open(verrou, "wb") as f:
        f.write(donnees + b"\x00" * (165 - len(donnees)))
    return verrou


class _Excel:
    """Session Excel factice : écrit avec openpyxl, ou simule un classeur exclusif / Excel en panne."""

    def __init__(self, journal, panne=None, verrouille=None):
        self.journal, self.panne, self.verrouille = journal, panne, verrouille
        self.ouvert = self.ferme = False

    def __enter__(self):
        if self.panne == "demarrage":
            raise excel_com.ExcelIndisponible("Excel introuvable")
        self.ouvert = True
        return self

    def __exit__(self, *exc):
        self.ferme = True
        return False

    def appliquer_regles(self, fichier, config, ids, on_etape=None):
        self.journal.append(("excel", os.path.basename(fichier)))
        if self.panne == "connexion":
            raise excel_com.ConnexionRequise("connectez-vous à Office")
        if self.panne == "fichier":
            raise ValueError("onglet Audit introuvable")
        if self.verrouille:
            bilan = BilanFichier(comptes={r.code: 0 for r in config.actives})
            bilan.verrouille = self.verrouille
            return bilan
        return appliquer_regles(fichier, config, ids, on_etape=on_etape)


class _Api:
    """Repli « API SharePoint » factice : refus 423 si ``verrouille``."""

    def __init__(self, journal, verrouille=None):
        self.journal, self.verrouille = journal, verrouille

    def appliquer_regles(self, fichier, config, ids, on_etape=None):
        self.journal.append(("api", os.path.basename(fichier)))
        bilan = BilanFichier(comptes={r.code: 0 for r in config.actives})
        bilan.verrouille = self.verrouille
        return bilan


@pytest.fixture
def fichier(tmp_path):
    return ecrire_audit(str(tmp_path / "Dep13" / "13001" / "audit_13001.xlsx"),
                        [["13001_1", "a"], ["13001_2", "b"], ["13001_2", "c"]])


def _vf(chemin):
    wb = load_workbook(chemin)
    ws = wb["Audit"]
    entetes = [c.value for c in ws[1]]
    valeurs = {r[entetes.index(IMM)]: r[entetes.index("Validation finale")]
               for r in ws.iter_rows(min_row=2, values_only=True)}
    wb.close()
    return valeurs


# ── Fichier verrou d'Excel ───────────────────────────────────────────
def test_proprietaire_verrou_lit_le_nom(fichier):
    assert proprietaire_verrou(fichier) is None
    _verrou(fichier, "Bérénice Durand")
    assert proprietaire_verrou(fichier) == "Bérénice Durand"


def test_proprietaire_verrou_ignore_les_adresses():
    assert proprietaire_verrou("https://x.sharepoint.com/a.xlsx") is None
    assert proprietaire_verrou("") is None


# ── Ordre des voies d'écriture ───────────────────────────────────────
def test_excel_prefere_ecrit_cellule_par_cellule(fichier):
    journal = []
    excel = _Excel(journal)
    with SessionEcriture(lambda: excel, repli=_Api(journal)) as ecriture:
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 3)])
    assert journal == [("excel", "audit_13001.xlsx")]          # pas d'API : Excel suffit
    assert ecriture.excel_utilise and not ecriture.repli_utilise and excel.ferme
    vf = _vf(fichier)
    assert vf[imm("13001", 3)] == "OK" and vf[imm("13001", 4)] is None   # même ID erreur, autre ID imm : intact
    assert not bilan.verrouille


def test_fichier_ouvert_sans_preference_excel_passe_par_excel(fichier):
    """Option décochée : réécriture directe, sauf pour un fichier ouvert → co-édition Excel à la demande."""
    journal, etapes = [], []
    excel = _Excel(journal)
    _verrou(fichier)
    with SessionEcriture(lambda: excel, prefere_excel=False) as ecriture:
        assert not excel.ouvert                                    # démarré seulement au besoin
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)], on_etape=etapes.append)
    assert journal == [("excel", "audit_13001.xlsx")] and not bilan.verrouille
    assert any("ouvert par Alice Martin" in e and "co-édition" in e for e in etapes)
    assert _vf(fichier)[imm("13001", 2)] == "OK"


def test_fichier_ouvert_sans_excel_rien_ecrit_et_raison(fichier):
    _verrou(fichier)
    avant = open(fichier, "rb").read()
    with SessionEcriture(None) as ecriture:
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])
    assert open(fichier, "rb").read() == avant
    assert bilan.verrouille.startswith("Alice Martin — l'écriture dans un fichier ouvert")
    assert "co-édition Excel, indisponible (Excel / pywin32 introuvables" in bilan.verrouille


def test_refus_api_423_rattrape_par_excel(fichier):
    journal = []
    with SessionEcriture(lambda: _Excel(journal), repli=_Api(journal, verrouille="Bob"),
                         prefere_excel=False) as ecriture:
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])
    assert journal == [("api", "audit_13001.xlsx"), ("excel", "audit_13001.xlsx")]
    assert not bilan.verrouille and ecriture.excel_utilise


def test_fichier_extrait_jamais_tente_via_excel(fichier):
    journal = []
    extrait = "Bob — fichier EXTRAIT (check-out) : réarchivez-le"
    with SessionEcriture(lambda: _Excel(journal), repli=_Api(journal, verrouille=extrait),
                         prefere_excel=False) as ecriture:
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])
    assert journal == [("api", "audit_13001.xlsx")] and bilan.verrouille == extrait


@pytest.mark.parametrize("panne", ["demarrage", "connexion"])
def test_excel_hors_service_repli_explique(fichier, panne):
    journal, etapes = [], []
    with SessionEcriture(lambda: _Excel(journal, panne=panne), repli=_Api(journal)) as ecriture:
        ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)], on_etape=etapes.append)
        ecriture.appliquer_regles(fichier, _config(), [imm("13001", 3)], on_etape=etapes.append)
    assert journal.count(("api", "audit_13001.xlsx")) == 2
    assert journal.count(("excel", "audit_13001.xlsx")) == (1 if panne == "connexion" else 0)   # un seul essai
    assert ecriture.repli_utilise and ("Excel" in ecriture.raison_sans_excel)


def test_erreur_propre_au_fichier_remontee(fichier):
    with SessionEcriture(lambda: _Excel([], panne="fichier")) as ecriture:
        with pytest.raises(ValueError):
            ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])


def test_classeur_exclusif_explique(fichier):
    with SessionEcriture(lambda: _Excel([], verrouille="Carole")) as ecriture:
        bilan = ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])
    assert bilan.verrouille.startswith("Carole — classeur ouvert en mode exclusif")


# ── De bout en bout : mise_a_jour ────────────────────────────────────
def test_mise_a_jour_fichier_ouvert_raison_par_id_imm(racine):
    """Fichier ouvert, pas d'Excel : rien écrit, chaque ID imm porte la raison et part en attente."""
    rapport = auditer_communes(["13001", "2A004"], racine)
    groupes = groupes_depuis_rapport(rapport)
    chemin = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    _verrou(chemin)
    avant = open(chemin, "rb").read()
    excel = _Excel([], panne="demarrage")
    bilan = mettre_a_jour(groupes, ReglesConfig.defaut(), racine, session_factory=lambda: excel,
                          colonne_id=IMM)
    assert open(chemin, "rb").read() == avant
    res = {r.commune: r for r in bilan.resultats}
    assert res["13001"].statut is not Statut.SUCCES and res["2A004"].statut is Statut.SUCCES
    assert set(bilan.echecs) == set(groupes["13001"])
    assert all("Alice Martin" in r and "fichier ouvert" in r for r in bilan.echecs.values())
    journal = "\n".join(p.refus_mise_a_jour(bilan.echecs))
    assert f"⛔ {len(groupes['13001'])} ID imm non mis à jour" in journal and "Alice Martin" in journal
    assert "💡" in journal


def test_mise_a_jour_fichier_ouvert_ecrit_via_excel(racine):
    rapport = auditer_communes(["13001"], racine)
    groupes = groupes_depuis_rapport(rapport)
    _verrou(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"))
    journal = []
    bilan = mettre_a_jour(groupes, ReglesConfig.defaut(), racine, session_factory=lambda: _Excel(journal),
                          colonne_id=IMM)
    assert bilan.echecs == {} and bilan.via_excel and journal == [("excel", "audit_13001.xlsx")]


def test_dossier_onedrive_force_excel(racine, monkeypatch):
    """« Parcourir » sur un dossier synchronisé par OneDrive : co-édition Excel même option décochée."""
    from avf.services import mise_a_jour as module
    monkeypatch.setattr(module, "_synchronise", lambda r: True)
    rapport = auditer_communes(["13001"], racine)
    journal = []
    bilan = mettre_a_jour(groupes_depuis_rapport(rapport), ReglesConfig.defaut(), racine,
                          via_excel=False, session_factory=lambda: _Excel(journal), colonne_id=IMM)
    assert bilan.via_excel and journal == [("excel", "audit_13001.xlsx")]


# ── Journal ──────────────────────────────────────────────────────────
def test_refus_groupes_par_raison_et_conseils():
    lignes = p.refus_mise_a_jour({
        "ID-1": "absent de la colonne « ID imm » du fichier",
        "ID-2": "absent de la colonne « ID imm » du fichier",
        "ID-3": "fichier audit introuvable",
    })
    assert lignes[0].startswith("⛔ 2 ID imm non mis à jour — absent de la colonne")
    assert "ID-1, ID-2" in lignes[0] and "faute de frappe" in lignes[0]
    assert "ID-3" in lignes[1] and "Dep<XX>" in lignes[1]


def test_refus_liste_tronquee():
    lignes = p.refus_mise_a_jour({f"ID-{i:02d}": "x" for i in range(20)}, limite=5)
    assert "… (+15)" in lignes[0]


def test_encore_non_valides_apres_reaudit(racine):
    rapport = auditer_communes(["13001"], racine)
    non_valides = [i.identifiant for c in rapport.lues for i in c.ids if not i.valide]
    lignes = p.encore_non_valides(rapport, non_valides[:2] + ["ID-INCONNU"])
    assert lignes[0].startswith("⚠ 2 ID imm mis à jour mais encore non valide")
    assert len(lignes) == 3 and all(" — " in l for l in lignes[1:])
    assert p.encore_non_valides(RapportAudit(racine=racine, communes=[]), non_valides) == []


# ── Moteur : classeur déjà ouvert dans l'Excel de l'utilisateur ──────
def test_cle_classeur_compare_chemins_et_adresses():
    assert excel_com._cle_classeur(r"C:\A\audit 1.xlsx") == excel_com._cle_classeur("c:/a/audit%201.xlsx")
    assert excel_com._classeur_ouvert_ailleurs([None, ""]) is None


def test_erreur_de_structure_jamais_rattrapee_par_l_api(fichier):
    from sna.services.excel import ColonneIntrouvable

    class _ExcelColonne(_Excel):
        def appliquer_regles(self, fichier, config, ids, on_etape=None):
            raise ColonneIntrouvable("colonne « ID imm » absente")

    journal = []
    with SessionEcriture(lambda: _ExcelColonne(journal), repli=_Api(journal)) as ecriture:
        with pytest.raises(ColonneIntrouvable):
            ecriture.appliquer_regles(fichier, _config(), [imm("13001", 2)])
    assert journal == []


# ── Diagnostic (lecture seule) ───────────────────────────────────────
class _ExcelCopie:
    def __init__(self, erreur=None):
        self.erreur = erreur

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def copier_version_serveur(self, adresse, destination, on_etape=None):
        if self.erreur:
            raise self.erreur
        open(destination, "wb").write(b"x")
        return destination


def test_diagnostic_local_fichier_ouvert_excel_ok(racine):
    from avf.models.source import LOCAL, SourceAudit
    from avf.services.diagnostic_ecriture import KO, OK, diagnostiquer_ecriture
    chemin = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    _verrou(chemin)
    avant = open(chemin, "rb").read()
    lignes = diagnostiquer_ecriture(SourceAudit(LOCAL, racine), "13001", session_factory=_ExcelCopie,
                                    excel_present=lambda: True)
    assert any(e == KO and "ouvert par Alice Martin" in t for e, t in lignes)
    assert any(e == OK and "co-édition est possible" in t for e, t in lignes)
    assert lignes[-1] == (OK, lignes[-1][1]) and "CONCLUSION" in lignes[-1][1]
    assert open(chemin, "rb").read() == avant


def test_diagnostic_excel_refuse(racine):
    from avf.models.source import URL, SourceAudit
    from avf.services.diagnostic_ecriture import KO, diagnostiquer_ecriture
    source = SourceAudit(URL, "https://x.sharepoint.com/sites/S/Docs/audit", "https://x.sharepoint.com/sites/S/Docs/audit")
    lignes = diagnostiquer_ecriture(source, "13001", session_factory=lambda: _ExcelCopie(OSError("accès refusé")),
                                    excel_present=lambda: True)
    assert any("Dep13/13001/audit_13001.xlsx" in t for _, t in lignes)
    assert any(e == KO and "Excel ne peut pas ouvrir ce fichier : accès refusé" in t for e, t in lignes)
    assert lignes[-1][1].startswith("CONCLUSION : aucune voie")


def test_diagnostic_commune_absente(racine):
    from avf.models.source import LOCAL, SourceAudit
    from avf.services.diagnostic_ecriture import diagnostiquer_ecriture
    lignes = diagnostiquer_ecriture(SourceAudit(LOCAL, racine), "99999", excel_present=lambda: True)
    assert lignes[-1][1] == "fichier audit introuvable : rien à diagnostiquer"


def test_reglages_des_reprises():
    from avf.config import Parametres
    assert (Parametres().reprises_verrou, Parametres().delai_reprise_s) == (3, 60)
    p2 = Parametres.from_dict({"reprises_verrou": 99, "delai_reprise_s": "x"})
    assert (p2.reprises_verrou, p2.delai_reprise_s) == (20, 60)
    assert Parametres.from_dict(p2.to_dict()).reprises_verrou == 20


def test_conseil_excel_et_verrou():
    assert "Fichier › Compte" in p.conseil_refus("Excel n'a pas pu ouvrir ce fichier : x")
    assert "🩺" in p.conseil_refus("classeur verrouillé par Bob")


# ── Démarrage d'Excel : réponse française de tasklist ────────────────
@pytest.mark.parametrize("sortie", [
    b"Information\xff: aucune t\x83che en service ne correspond aux crit\x8ares sp\x82cifi\x82s.\r\n",  # réel
    None,
])
def test_pids_excel_reponse_francaise_sans_excel(monkeypatch, sortie):
    """Régression : « 'NoneType' object has no attribute 'splitlines' » empêchait Excel de démarrer."""
    import subprocess

    class _Resultat:
        stdout = sortie
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Resultat())
    assert excel_com._pids_excel() == set()


def test_pids_excel_lit_les_pid(monkeypatch):
    import subprocess

    class _Resultat:
        stdout = b'"EXCEL.EXE","4242","Console","1","95\xff000 Ko"\r\n"EXCEL.EXE","77","Console","1","1 Ko"\r\n'
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Resultat())
    assert excel_com._pids_excel() == {4242, 77}
