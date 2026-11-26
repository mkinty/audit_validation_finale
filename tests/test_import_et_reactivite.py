"""Import de listes (txt, csv, xlsx) et réactivité de l'interface avec de longues listes.

Jeu réel : ``tests/donnees/communes_exemple.txt`` (320 codes INSEE, doublons et
codes corses compris), fourni par l'utilisateur.
"""

import http.server
import os
import threading
import time

import pytest
from openpyxl import Workbook

from avf.models.source import API, SourceAudit
from avf.services import sharepoint
from avf.services import sharepoint_rest as rest
from avf.services.detection import parser_communes
from avf.services.import_liste import COLONNES_ID, COLONNES_INSEE, ImportImpossible, lire_liste
from sna.services.identifiants import extraire_ids
from tests.conftest import ecrire_audit
from tests.test_sharepoint_api import LIEN, _FauxSharePoint
from tests.test_ui import DialogueSaisie, app, attendre, dialogues, sans_vrai_navigateur  # noqa: F401

LISTE = os.path.join(os.path.dirname(__file__), "donnees", "communes_exemple.txt")
CODES = open(LISTE, encoding="utf-8").read().split()


# ── Import de fichiers ────────────────────────────────────────────
def test_liste_reelle_txt():
    liste = lire_liste(LISTE, COLONNES_INSEE)
    assert len(liste.valeurs) == 320 and liste.origine == "tout le fichier"
    codes = parser_communes(liste.texte)
    assert len(set(codes)) == 315 and "2B334" in codes and "04006" in codes   # doublons retirés à l'usage


def test_csv_colonne_reconnue_parmi_d_autres(tmp_path):
    chemin = tmp_path / "communes.csv"
    lignes = ["Département;Code INSEE;Nom"] + [f"{c[:2]};{c};Commune {i}" for i, c in enumerate(CODES[:50])]
    chemin.write_text("\n".join(lignes), encoding="cp1252")
    liste = lire_liste(str(chemin), COLONNES_INSEE)
    assert liste.valeurs == CODES[:50] and liste.origine == "colonne « Code INSEE »"


def test_csv_virgule_et_id_erreur(tmp_path):
    chemin = tmp_path / "ids.csv"
    chemin.write_text("commune,ID_erreur,statut\n13001,13001_2,KO\n13001,13001_5,KO\n", encoding="utf-8-sig")
    assert lire_liste(str(chemin), COLONNES_ID).valeurs == ["13001_2", "13001_5"]


def test_xlsx_colonne_sur_une_autre_feuille_et_zero_perdu(tmp_path):
    chemin = str(tmp_path / "liste.xlsx")
    wb = Workbook()
    wb.active.title = "Notes"
    wb.active["A1"] = "rien d'utile"
    ws = wb.create_sheet("Communes")
    ws.append(["Titre du suivi"])
    ws.append(["Nom", "INSEE"])
    ws.append(["Aiglun", 4006])              # saisi comme nombre : zéro de tête perdu
    ws.append(["Ajaccio", "2B334"])
    ws.append(["", None])
    wb.save(chemin)
    liste = lire_liste(chemin, COLONNES_INSEE)
    assert liste.valeurs == ["04006", "2B334"]
    assert liste.origine == "colonne « INSEE » (feuille Communes)"


def test_xlsx_sans_colonne_reconnue_lit_tout(tmp_path):
    chemin = str(tmp_path / "ids.xlsx")
    wb = Workbook()
    wb.active.append(["Liste à reprendre"])
    wb.active.append(["13001_2", "voir 13001_5"])
    wb.save(chemin)
    liste = lire_liste(chemin, COLONNES_ID)
    assert extraire_ids(liste.texte) == ["13001_2", "13001_5"] and "aucune colonne" in liste.origine


def test_format_refuse(tmp_path):
    chemin = tmp_path / "liste.pdf"
    chemin.write_bytes(b"%PDF")
    with pytest.raises(ImportImpossible, match="format non pris en charge"):
        lire_liste(str(chemin), COLONNES_INSEE)


def test_dialogue_charge_le_fichier(app):
    d = DialogueSaisie(app, "t", "c", "e", colonnes_fichier=COLONNES_INSEE)
    assert d.charger_fichier(LISTE)
    assert d.zone.get("1.0", "end").split() == CODES
    assert "320 valeur(s)" in d.lbl_compte["text"]
    d.destroy()


# ── Réactivité : jamais de requête réseau sur le fil de l'interface ─
@pytest.fixture
def sharepoint_lent(tmp_path):
    """Faux SharePoint aux latences réalistes, avec un dossier pour 305 des 315 communes."""
    racine = str(tmp_path / "racine")
    for code in set(CODES[:-15]):
        ecrire_audit(os.path.join(racine, f"Dep{code[:2]}", code, f"audit_{code}.xlsx"), [[f"{code}_1"]])

    class Lent(_FauxSharePoint):
        def do_GET(self):
            time.sleep(0.05)
            return _FauxSharePoint.do_GET(self)

    gestion = type("F", (Lent,), {"racine": racine, "appels": [], "versions": {}, "verrouilles": set(),
                                   "extraits": {}})
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), gestion)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    client = rest.ClientSharePoint("contoso.sharepoint.com", {"FedAuth": "valide"},
                                   base=f"http://127.0.0.1:{srv.server_port}")
    sharepoint.enregistrer_client(client)
    yield gestion, client
    sharepoint._CLIENTS.clear()
    srv.shutdown()


def _connecte(app):
    app.var_source.set("sharepoint")
    app._changer_source()
    app._source_sp = SourceAudit(API, LIEN, LIEN)


def test_coller_la_liste_reelle_ne_fige_pas_l_interface(app, sharepoint_lent, monkeypatch):
    gestion, _ = sharepoint_lent
    _connecte(app)
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: "\n".join(CODES))
    debut = time.monotonic()
    app._renseigner_communes()
    assert time.monotonic() - debut < 0.5          # avant : 18 s d'interface figée
    assert app.occupe                              # le travail continue en arrière-plan
    attendre(app, timeout=60)
    assert len(app.liste_com.codes) == len(set(CODES[:-15]))
    assert len(gestion.appels) <= 30               # un scan (1 + une requête par Dep), pas 315


def test_coller_des_centaines_d_id_ne_fige_pas_l_interface(app, sharepoint_lent, monkeypatch, dialogues):
    gestion, _ = sharepoint_lent
    _connecte(app)
    ids = " ".join(f"{c}_{n}" for c in CODES for n in (1, 2))
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: ids)
    debut = time.monotonic()
    app._renseigner_ids()
    assert time.monotonic() - debut < 0.5
    attendre(app, timeout=60)
    assert len(app.ids_cibles) == len(set(CODES[:-15]))
    avertissement = next(m for t, m in dialogues if t == "ID erreur")
    absentes = set(CODES[-15:]) - set(CODES[:-15])       # 12 : trois des 15 derniers sont des doublons
    assert f"{len(absentes)} commune(s) sans dossier" in avertissement and len(avertissement) < 1500
    assert len(gestion.appels) <= 30


def test_deuxieme_saisie_sans_aucune_requete(app, sharepoint_lent, monkeypatch):
    gestion, _ = sharepoint_lent
    _connecte(app)
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: " ".join(CODES[:100]))
    app._renseigner_communes()
    attendre(app, timeout=60)
    avant = len(gestion.appels)
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: " ".join(CODES[100:]))
    app._renseigner_communes()
    attendre(app, timeout=60)
    assert len(gestion.appels) == avant            # arborescence gardée en mémoire


def test_detection_rafraichit_l_arborescence(app, sharepoint_lent):
    gestion, client = sharepoint_lent
    _connecte(app)
    client.arborescence(rest.chemin_serveur(LIEN))
    avant = len(gestion.appels)
    app._detecter_communes()
    attendre(app, timeout=60)
    assert len(gestion.appels) > avant             # « Détecter » refait toujours le scan
    assert len(app.liste_com.codes) == len(set(CODES[:-15]))


def test_restauration_au_lancement_en_arriere_plan(tmp_path, sharepoint_lent, dialogues):
    """Après connexion, recharger les communes de la dernière fois ne fige pas non plus."""
    import json
    from avf.ui import application as module_app
    cfg = str(tmp_path / "cfg.json")
    json.dump({"source": "sharepoint", "communes": CODES}, open(cfg, "w", encoding="utf-8"))
    a = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    a.withdraw()
    try:
        debut = time.monotonic()
        a._fin_connexion(SourceAudit(API, LIEN, LIEN))
        assert time.monotonic() - debut < 0.5
        attendre(a, timeout=60)
        assert len(a.liste_com.codes) == len(set(CODES[:-15]))
    finally:
        a.destroy()


def test_interface_explique_l_acces_conditionnel(app, dialogues):
    from avf.models.audit import AuditCommune, RapportAudit, StatutAudit
    message = ("téléchargement interdit depuis cet appareil par les stratégies de l'organisation (…)")
    rapport = RapportAudit(racine=LIEN, communes=[
        AuditCommune(c, c[:2], statut=StatutAudit.ERREUR, message=message) for c in ("04006", "04017")])
    app._fin_audit(rapport, apres=None)
    avertissements = [m for t, m in dialogues if "organisation" in t]
    assert len(avertissements) == 1 and "poste de l'entreprise" in avertissements[0]
    app.traiter_file()
    assert "Téléchargement des fichiers SharePoint interdit" in app.txt_log.get("1.0", "end")


def test_milliers_de_communes_affichees_vite(app):
    """4 022 communes sur le vrai SharePoint : la liste se remplit sans figer l'interface."""
    from avf.services.detection import CommuneDetectee
    communes = [CommuneDetectee(f"{d:02d}{n:03d}", f"{d:02d}", "", n % 7 != 0) for d in range(1, 49) for n in range(84)]
    voulues = [c.code for c in communes if c.a_fichier]
    debut = time.monotonic()
    app._remplir_communes(communes, selection=voulues)
    assert time.monotonic() - debut < 1.0
    assert len(app.liste_com.codes) == 4032 and app._communes_choisies() == voulues
