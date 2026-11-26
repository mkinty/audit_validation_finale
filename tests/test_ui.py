"""Interface Tkinter pilotée de bout en bout (boîtes de dialogue neutralisées)."""

from __future__ import annotations

import gc
import os
import time
import tkinter as tk

import pytest

from avf.config import charger
from avf.ui import application as module_app
from avf.models.audit import NON_VALIDES

from .conftest import imm
from avf.ui import fenetre_rapport
from avf.ui.composants import DialogueColonnes, DialogueSaisie
from openpyxl import load_workbook


def _tk_disponible() -> bool:
    try:
        racine = tk.Tk()
        racine.destroy()
        return True
    except tk.TclError:
        return False


pytestmark = pytest.mark.skipif(not _tk_disponible(), reason="pas d'affichage Tk")


@pytest.fixture(autouse=True)
def sans_vrai_navigateur(monkeypatch):
    """Jamais de vrai navigateur lancé par les tests : la connexion utilise l'accès de
    repli (``sharepoint.connecter``), sauf si un test simule le navigateur lui-même."""
    from avf.services import sharepoint
    monkeypatch.setattr(sharepoint, "navigateur_disponible", lambda: False)

    def interdit(*a, **k):
        raise AssertionError("un test a voulu lancer un vrai navigateur")
    monkeypatch.setattr(sharepoint, "connecter_navigateur", interdit)


@pytest.fixture
def dialogues(monkeypatch):
    """Réponses automatiques aux boîtes de dialogue, et trace des messages."""
    traces = []

    def repondre(valeur):
        def f(titre, message="", **kw):
            traces.append((titre, message))
            return valeur
        return f

    for nom in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(module_app.messagebox, nom, repondre(None))
    monkeypatch.setattr(module_app.messagebox, "askyesno", repondre(True))
    monkeypatch.setattr(module_app, "ouvrir_dans_navigateur", lambda chemin: True)
    return traces


@pytest.fixture
def app(tmp_path, dialogues):
    a = module_app.Application(
        fichier_config=str(tmp_path / "cfg.json"), fichier_attente=str(tmp_path / "attente.json"),
    )
    a.withdraw()
    attendre(a)   # sélection restaurée en arrière-plan
    a.var_excel.set(False)  # tests : réécriture directe, jamais Excel
    yield a
    a.destroy()
    # Variables Tk libérées ici, sur le fil principal, plutôt que par un
    # ramasse-miettes déclenché plus tard dans un thread de travail.
    gc.collect()


def attendre(app, timeout=30):
    """Laisse tourner les threads de fond et la file jusqu'à la fin du traitement."""
    fin = time.time() + timeout
    while time.time() < fin:
        app.update()
        app.traiter_file()
        if not app.occupe:
            app.update()
            app.traiter_file()
            if not app.occupe:
                return
        time.sleep(0.02)
    raise TimeoutError("traitement trop long")


def preparer(app, racine, communes=None):
    """Racine renseignée, communes détectées ; ``communes`` restreint la sélection."""
    app.var_racine.set(racine)
    app._enregistrer_racine()
    app._detecter_communes()
    if communes is not None:
        app.liste_com.selection_clear(0, tk.END)
        for rang, code in enumerate(app.liste_com.codes):
            if code in communes:
                app.liste_com.selection_set(rang)
        app._sauver_selection()
        app._maj_etat()


def test_etat_initial_sans_racine(app):
    assert str(app.btn_det_com["state"]) == "disabled"
    assert str(app.btn_saisie_com["state"]) == "disabled"
    assert str(app.btn_audit["state"]) == "disabled"
    assert str(app.btn_maj["state"]) == "disabled"
    assert "Audit ← « OK »" in app.lbl_regles["text"]
    assert not hasattr(app, "liste_dep")


def test_racine_renseignee_active_les_communes(app, racine):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    assert str(app.btn_det_com["state"]) == "normal"
    assert str(app.btn_saisie_com["state"]) == "normal"


def test_detection_de_toutes_les_communes(app, racine):
    preparer(app, racine)
    assert app.liste_com.codes == ["13001", "13002", "13003", "2A004", "75020", "75021"]
    # Les communes sans fichier audit ne sont pas présélectionnées
    assert app._communes_choisies() == ["13001", "13002", "2A004", "75020", "75021"]
    assert "Dep2A" in app.liste_com.get(3)
    assert str(app.btn_audit["state"]) == "normal"


def test_renseigner_des_communes(app, racine, dialogues, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: "13001 2a004, 88888")
    app._renseigner_communes()
    attendre(app)
    assert app.liste_com.codes == ["13001", "2A004"]
    assert app._communes_choisies() == ["13001", "2A004"]
    assert any("88888" in m for _, m in dialogues)


def test_saisie_annulee(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: None)
    app._renseigner_communes()
    attendre(app)
    assert app.liste_com.codes == []


def _ligne_detail(f, iid):
    return dict(zip([c[0] for c in fenetre_rapport._DETAIL], f.detail.item(iid, "values")))


def test_audit_rapport_copie_mise_a_jour_reaudit_export(app, racine, tmp_path):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)

    r = app.rapport
    assert (r.nb_lignes_valides, r.nb_lignes_non_valides, r.nb_ids_non_valides) == (3, 8, 8)
    assert "8 ID non valides" in app.btn_maj["text"]
    journal = app.txt_log.get("1.0", tk.END)
    assert "13001 — 6 ligne(s) · 1 validée(s) · 5 non valide(s)" in journal and "75020" in journal
    assert "Ligne validée si : Audit = OK / KO" in journal

    # Popup ouverte automatiquement après l'audit
    f = app.fenetre_rapport
    assert f is not None and f.winfo_exists()
    f.update()
    assert f.tuiles["valides"][0]["text"] == "3" and f.tuiles["non_valides"][0]["text"] == "8"
    assert "Date audit = une date : 7" in f.lbl_motifs["text"]
    assert set(f.arbre.get_children()) == {"13001", "13002", "2A004", "75020", "75021"}

    # Détail de la commune : lignes non valides par défaut, avec Validation finale et motifs
    f.arbre.selection_set("13001")
    f.update()
    assert f.detail["columns"][:2] == ("imm", "id")                 # ID imm avant ID erreur
    assert f.detail["columns"][7] == "vf" and f.detail.heading("vf")["text"] == "Validation finale"
    lignes = [_ligne_detail(f, iid) for iid in f.detail.get_children()]
    assert [(l["imm"], l["id"], l["ligne"]) for l in lignes] == [
        (imm("13001", 3), "13001_2", "3"), (imm("13001", 4), "13001_3", "4"), (imm("13001", 5), "13001_4", "5"),
        (imm("13001", 6), "13001_5", "6"), (imm("13001", 7), "13001_5", "7"),
    ]
    assert lignes[2]["vf"] == "Validé" and lignes[2]["motifs"] == "Validation finale « Validé » ≠ OK"
    assert lignes[2]["statut"] == "⚠️ Non valide"
    f._copier_commune()                                           # les ID imm, pas les ID erreur
    assert f.clipboard_get() == "\n".join(imm("13001", n) for n in (3, 4, 5, 6, 7))
    # Les deux lignes de 13001_5 sélectionnées : leurs deux ID imm
    f.detail.selection_set([iid for iid in f.detail.get_children()
                            if _ligne_detail(f, iid)["id"] == "13001_5"])
    f._copier_selection_detail()
    assert f.clipboard_get() == f"{imm('13001', 6)}\n{imm('13001', 7)}"
    f.var_separateur.set("virgule")
    f._copier_affiches()
    assert f.clipboard_get().split(", ") == r.ids_de(NON_VALIDES)

    # Case « Afficher aussi les validées » (comme audit_ppts) : la ligne validée s'ajoute
    assert f.case_tous["text"] == "Afficher aussi les validées" and not f.var_tous.get()
    f.case_tous.invoke()
    lignes = [_ligne_detail(f, i) for i in f.detail.get_children()]
    assert [l["ligne"] for l in lignes] == ["2", "3", "4", "5", "6", "7"]
    assert lignes[0]["imm"] == imm("13001", 2) and lignes[0]["statut"] == "✅ Validée"
    assert f.detail.item(f.detail.get_children()[0], "tags") == ("validee",)
    f.var_separateur.set("ligne")
    f._copier_commune()                                   # la copie reste celle des non valides
    assert f.clipboard_get() == "\n".join(imm("13001", n) for n in (3, 4, 5, 6, 7))
    assert "(8)" in f.btn_copier_affiches["text"]
    f.case_tous.invoke()
    assert len(f.detail.get_children()) == 5
    # Commune complète : rien par défaut, ses lignes avec la case cochée
    f.arbre.selection_set("13002")
    f.update()
    assert f.detail.get_children() == () and str(f.btn_copier_commune["state"]) == "disabled"
    f.case_tous.invoke()
    assert len(f.detail.get_children()) == 2
    f.case_tous.invoke()
    f.arbre.selection_set("13001")
    f.update()

    # Filtre texte sur un ID erreur, filtre par état
    f.var_recherche.set("2a004_2")
    f.update()
    assert f.arbre.get_children() == ("2A004",)
    f.var_recherche.set("")
    f.var_filtre.set("❌ En échec")
    f._remplir()
    assert f.arbre.get_children() == ("75020",)
    f.var_filtre.set("⚠️ À corriger")
    f._remplir()
    assert f.arbre.get_children() == ("13001", "2A004", "75021")
    f.var_filtre.set("Toutes")
    f._remplir()

    # Tri par nombre de lignes non valides décroissant
    f._trier("nv")
    f._trier("nv")
    assert f.arbre.get_children()[0] == "13001"

    # Mise à jour de toutes les communes depuis la popup, puis ré-audit automatique
    f._maj_tout()
    attendre(app)
    attendre(app)  # le ré-audit est relancé à la fin de la mise à jour
    assert app.rapport.ids_non_valides_par_commune() == {"75021": [imm("75021", 2)]}  # conflit refusé
    assert f.tuiles["non_valides"][0]["text"] == "1"
    journal = app.txt_log.get("1.0", tk.END)
    assert "Ré-audit" in journal and "conflit" in journal

    # Exports HTML, Excel, CSV
    chemin = app._exporter("html", str(tmp_path / "rapport.html"), ouvrir=False)
    assert "75021_1" in open(chemin, encoding="utf-8").read()
    chemin = app._exporter("xlsx", str(tmp_path / "rapport.xlsx"), ouvrir=False)
    assert load_workbook(chemin).sheetnames == ["Synthèse", "Communes", "Lignes", "ID imm non valides"]
    chemin = app._exporter("csv", str(tmp_path / "rapport.csv"), ouvrir=False)
    assert "75021_1" in open(chemin, encoding="utf-8-sig").read()
    assert charger(app.fichier_config).dossier_export == str(tmp_path)


def test_export_ouvre_excel_ou_le_navigateur(app, racine, tmp_path, monkeypatch):
    preparer(app, racine, communes=["13001"])
    app._lancer_audit()
    attendre(app)
    ouverts = []
    monkeypatch.setattr(module_app, "ouvrir_fichier", lambda c: ouverts.append(("excel", c)))
    monkeypatch.setattr(module_app, "ouvrir_dans_navigateur", lambda c: ouverts.append(("web", c)))
    app._exporter("xlsx", str(tmp_path / "a.xlsx"), ouvrir=True)
    app._exporter("html", str(tmp_path / "a.html"), ouvrir=True)
    assert [o[0] for o in ouverts] == ["excel", "web"]


def test_export_impossible_signale(app, racine, tmp_path, dialogues):
    preparer(app, racine, communes=["13001"])
    app._lancer_audit()
    attendre(app)
    (tmp_path / "pris").mkdir()
    assert app._exporter("xlsx", str(tmp_path / "pris"), ouvrir=False) is None
    assert any("n'a pas pu être écrit" in m for _, m in dialogues)


def test_mise_a_jour_d_une_seule_commune(app, racine):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    f = app.fenetre_rapport
    f.arbre.selection_set("2A004")
    f.update()
    f._maj_commune()
    attendre(app)
    attendre(app)
    assert set(app.rapport.ids_non_valides_par_commune()) == {"13001", "75021"}


def test_simulation_ne_modifie_rien(app, racine):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    app.var_simulation.set(True)
    app._lancer_mise_a_jour(None)
    attendre(app)
    assert "simulation" in app.txt_log.get("1.0", tk.END)
    app._lancer_audit(["13001"])
    attendre(app)
    assert app.rapport.commune("13001").nb_ids_non_valides == 5


def test_refus_de_confirmation(app, racine, monkeypatch):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    monkeypatch.setattr(module_app.messagebox, "askyesno", lambda *a, **k: False)
    app._lancer_mise_a_jour(None)
    assert not app.occupe
    app._lancer_audit(["13001"])
    attendre(app)
    assert app.rapport.commune("13001").nb_ids_non_valides == 5


def test_colonnes_auditees(app, racine, monkeypatch):
    from avf.models.colonnes import ColonnesAudit
    monkeypatch.setattr(DialogueColonnes, "montrer",
                        lambda self: (ColonnesAudit(validation="Inexistante"), "ID erreur", 1))
    app._configurer_colonnes()
    preparer(app, racine, communes=["13001", "13002"])
    app._lancer_audit()
    attendre(app)
    assert all(not c.lisible for c in app.rapport.communes)


def test_options_et_selection_persistees(tmp_path, racine, dialogues):
    cfg = str(tmp_path / "cfg.json")
    a = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    a.withdraw()
    attendre(a)   # sélection restaurée en arrière-plan
    preparer(a, racine)
    a.liste_com.selection_clear(0, tk.END)
    a.liste_com.selection_set(1)  # 13002
    a._sauver_selection()
    a.var_date_du_jour.set(False)
    a.var_reauditer.set(False)
    a._sauver_options()
    a.destroy()

    b = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    b.withdraw()
    attendre(b)   # sélection restaurée en arrière-plan
    try:
        assert b.var_racine.get() == racine
        assert b.liste_com.codes == ["13002"]
        assert b._communes_choisies() == ["13002"]
        assert not b.var_date_du_jour.get() and not b.var_reauditer.get()
    finally:
        b.destroy()
        gc.collect()


def test_regles_par_defaut_conservent_la_colonne_id(app):
    app.params.regles.colonne_id = "Identifiant"
    app.params.regles.regles = []
    app._regles_par_defaut()
    assert app.params.regles.colonne_id == "Identifiant"
    assert len(app.params.regles.regles) == 4


def test_panneau_regles_s_ouvre(app, racine):
    preparer(app, racine)
    app._ouvrir_regles()
    app.update()
    panneaux = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel"]
    assert panneaux
    for w in panneaux:
        w.destroy()


def test_arret_pendant_la_mise_a_jour(app, racine, monkeypatch):
    """« Arrêter » est transmis au moteur : les communes non commencées restent intactes."""
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    vrai_mettre_a_jour, recus = module_app.mettre_a_jour, {}

    def mettre_a_jour_puis_arreter(*args, **kwargs):
        recus.update(kwargs)
        app._arret.set()   # l'utilisateur clique « Arrêter » dès le début
        return vrai_mettre_a_jour(*args, **kwargs)

    monkeypatch.setattr(module_app, "mettre_a_jour", mettre_a_jour_puis_arreter)
    app.var_reauditer.set(False)
    app._lancer_mise_a_jour(None)
    attendre(app)
    assert recus["paralleles"] == app.params.paralleles
    assert "Arrêt demandé" in app.txt_log.get("1.0", tk.END)
    app._lancer_audit()
    attendre(app)
    assert app.rapport.nb_ids_non_valides == 8   # rien n'a été écrit


def test_dialogue_colonnes_valeurs_acceptees(app):
    from avf.models.colonnes import ColonnesAudit
    d = DialogueColonnes(app, ColonnesAudit(), "ID erreur", 1)
    assert d.vars["valeurs_audit"].get() == "OK, KO"
    d.vars["valeurs_validation"].set("OK; Validé")
    d.vars["auditeur"].set("Contrôleur")
    d._valider()
    colonnes, colonne_id, ligne = d.resultat
    assert colonnes.valeurs_validation == ("OK", "Validé") and colonnes.auditeur == "Contrôleur"
    assert (colonne_id, ligne) == ("ID erreur", 1)


# ── Source SharePoint ─────────────────────────────────────────────
LIEN = "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Audit"


def _connecter(app, monkeypatch, source=None, erreur=None):
    from avf.services import sharepoint

    def connecter(lien):
        if erreur:
            raise erreur
        return source
    monkeypatch.setattr(module_app.sharepoint, "connecter", connecter)
    app.var_source.set("sharepoint")
    app._changer_source()
    app.var_lien.set(LIEN)
    app._connecter_sharepoint()
    attendre(app)
    return sharepoint


def test_choix_de_la_source_affiche_la_bonne_saisie(app):
    assert app.ligne_local.winfo_manager() == "pack" and not app.ligne_sp.winfo_manager()
    app.var_source.set("sharepoint")
    app._changer_source()
    assert app.ligne_sp.winfo_manager() == "pack" and not app.ligne_local.winfo_manager()
    assert "lien SharePoint" in app.lbl_source["text"]
    assert str(app.btn_det_com["state"]) == "disabled"          # pas encore connecté
    assert charger(app.fichier_config).source == "sharepoint"


def test_connexion_sharepoint_puis_audit_de_la_version_serveur(app, racine, monkeypatch):
    from avf.models.source import WEBDAV, SourceAudit
    source = SourceAudit(WEBDAV, racine, LIEN)
    _connecter(app, monkeypatch, source)
    assert app._source() == source
    assert "Connecté" in app.lbl_source["text"] and "WebDAV" in app.lbl_source["text"]
    assert charger(app.fichier_config).lien_sharepoint == LIEN
    app._detecter_communes()
    assert "13001" in app.liste_com.codes
    app._lancer_audit()
    attendre(app)
    r = app.rapport
    assert r.source == source and r.racine == LIEN
    assert "Source : SharePoint — accès direct (WebDAV)" in app.txt_log.get("1.0", tk.END)
    assert "SharePoint" in app.fenetre_rapport.lbl_meta["text"]


def test_echec_de_connexion_explique(app, monkeypatch, dialogues):
    from avf.services.sharepoint import AccesImpossible
    _connecter(app, monkeypatch, erreur=AccesImpossible("dossier SharePoint injoignable (accès refusé).\n\nDétails"))
    assert app._source() is None and not app.occupe
    assert "Non connecté" in app.lbl_source["text"]
    assert any("injoignable" in m and "Détails" in m for _, m in dialogues)
    assert str(app.btn_audit["state"]) == "disabled"


def test_lien_invalide_refuse_sans_connexion(app, monkeypatch, dialogues):
    appels = []
    monkeypatch.setattr(module_app.sharepoint, "connecter", lambda lien: appels.append(lien))
    app.var_source.set("sharepoint")
    app._changer_source()
    app.var_lien.set("https://contoso.sharepoint.com/:f:/s/SNA/EaBcD")
    app._connecter_sharepoint()
    assert appels == [] and any("Lien inutilisable" in m for _, m in dialogues)


def test_mise_a_jour_refusee_si_le_rapport_vient_d_une_autre_source(app, racine, monkeypatch, dialogues):
    from avf.models.source import WEBDAV, SourceAudit
    preparer(app, racine, communes=["13001"])
    app._lancer_audit()
    attendre(app)                                    # rapport du dossier local
    _connecter(app, monkeypatch, SourceAudit(WEBDAV, racine + "_serveur", LIEN))
    app._lancer_mise_a_jour(None)
    assert not app.occupe
    assert any("autre source" in m for _, m in dialogues)


def test_mise_a_jour_sharepoint_force_excel(app, racine, monkeypatch):
    from avf.models.source import WEBDAV, SourceAudit
    _connecter(app, monkeypatch, SourceAudit(WEBDAV, racine, LIEN))
    app._detecter_communes()
    app._lancer_audit(["13001"])
    attendre(app)
    recus = {}
    monkeypatch.setattr(module_app, "mettre_a_jour", lambda *a, **k: recus.update(k) or (_ for _ in ()).throw(
        RuntimeError("arrêt du test")))
    app._excel = True
    app.var_excel.set(False)                         # même décoché : SharePoint passe par Excel
    app._lancer_mise_a_jour(None)
    attendre(app)
    assert recus["via_excel"] is True and recus["source"].mode == WEBDAV


def test_source_sharepoint_retrouvee_au_lancement(tmp_path, racine, monkeypatch, dialogues):
    from avf.models.source import WEBDAV, SourceAudit
    monkeypatch.setattr(module_app.sharepoint, "connecter", lambda lien: SourceAudit(WEBDAV, racine, lien))
    cfg = str(tmp_path / "cfg.json")
    import json
    json.dump({"source": "sharepoint", "lien_sharepoint": LIEN, "communes": ["13002"]},
              open(cfg, "w", encoding="utf-8"))
    a = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    a.withdraw()
    attendre(a)   # sélection restaurée en arrière-plan
    try:
        fin = time.time() + 10
        while a._source() is None and time.time() < fin:   # connexion automatique, en arrière-plan
            a.update()
            a.traiter_file()
            time.sleep(0.02)
        attendre(a)
        assert a._source().lien == LIEN and a.liste_com.codes == ["13002"]
        assert dialogues == []                              # silencieuse au lancement
    finally:
        a.destroy()
        gc.collect()


# ── Audit restreint à des ID erreur ───────────────────────────────
def _saisir_ids(app, monkeypatch, texte):
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: texte)
    app._renseigner_ids()
    attendre(app)


def test_audit_par_ids_puis_mise_a_jour_et_reaudit(app, racine, monkeypatch, dialogues):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2 13001_5 13001_99\n2a004_2")
    assert app.liste_com.codes == ["13001", "2A004"] and app._communes_choisies() == ["13001", "2A004"]
    assert app.cadre_perimetre.winfo_manager() == "pack"
    assert "4 ID erreur sur 2 commune(s)" in app.lbl_perimetre["text"]
    assert charger(app.fichier_config).ids_cibles == ["13001_2", "13001_5", "13001_99", "2A004_2"]

    app._lancer_audit()
    attendre(app)
    r = app.rapport
    assert r.par_ids and r.nb_lignes == 4 and r.ids_introuvables() == ["13001_99"]
    journal = app.txt_log.get("1.0", tk.END)
    assert "Limité à 4 ID erreur" in journal and "introuvable(s) : 13001_99" in journal
    assert "4 ID saisis" in app.fenetre_rapport.lbl_meta["text"]

    # Mise à jour des ID non valides du périmètre, ré-audit limité aux mêmes ID
    app.fenetre_rapport._maj_tout()
    attendre(app)
    attendre(app)
    r = app.rapport
    assert r.par_ids and r.nb_lignes == 4 and r.nb_ids_non_valides == 0
    # Hors périmètre, rien n'a été touché : 2A004_1 reste non valide
    complet = module_app.auditer_communes(["2A004"], racine)
    assert complet.commune("2A004").ids_non_valides == [imm("2A004", 2)]


def test_saisie_ids_signale_les_problemes(app, racine, monkeypatch, dialogues):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2 99999_1 13002")
    messages = " ".join(m for _, m in dialogues)
    assert "1 commune(s) sans dossier" in messages and "99999" in messages
    assert "sans numéro" in messages and "13002" in messages
    assert app.liste_com.codes == ["13001"]


def test_saisie_ids_sans_rien_d_exploitable(app, racine, monkeypatch, dialogues):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "bonjour")
    assert any("Aucun ID erreur exploitable" in m for _, m in dialogues)
    assert app.ids_cibles == {}


def test_retour_aux_communes_entieres(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2")
    app._effacer_ids()
    assert app.ids_cibles == {} and not app.cadre_perimetre.winfo_manager()
    assert charger(app.fichier_config).ids_cibles == []
    _saisir_ids(app, monkeypatch, "13001_2")
    app._detecter_communes()                         # détecter les communes : communes entières
    assert app.ids_cibles == {}
    app._lancer_audit(["13001"])
    attendre(app)
    assert not app.rapport.par_ids and app.rapport.nb_lignes == 6


def test_ids_saisis_retrouves_au_lancement(tmp_path, racine, dialogues):
    import json
    cfg = str(tmp_path / "cfg.json")
    json.dump({"racine": racine, "ids_cibles": ["13001_2", "2A004_1"], "communes": ["2A004"]},
              open(cfg, "w", encoding="utf-8"))
    a = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    a.withdraw()
    attendre(a)   # sélection restaurée en arrière-plan
    try:
        assert a.ids_cibles == {"13001": ["13001_2"], "2A004": ["2A004_1"]}
        assert a.liste_com.codes == ["13001", "2A004"] and a._communes_choisies() == ["2A004"]
        assert "2 ID erreur" in a.lbl_perimetre["text"]
    finally:
        a.destroy()
        gc.collect()


# ── Mode « via Excel, sans synchronisation » ──────────────────────
def test_mode_adresse_saisie_sans_parcours(app, monkeypatch, dialogues):
    from avf.models.source import URL, SourceAudit
    source = SourceAudit(URL, LIEN, LIEN, diagnostic="Windows n'a pas de session SharePoint (erreur 224)")
    _connecter(app, monkeypatch, source)
    texte = app.lbl_source["text"]
    assert "sans synchronisation" in texte and "erreur 224" in texte and "renseignez les communes" in texte
    assert str(app.btn_det_com["state"]) == "disabled"
    app._detecter_communes()
    assert any("ne peut pas être parcouru" in m for _, m in dialogues)
    _saisir_ids(app, monkeypatch, "38153_1 38153_7")
    assert app.liste_com.codes == ["38153"]              # retenue sans dossier vérifiable
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: "38153 38154")
    app._renseigner_communes()
    attendre(app)
    assert app.liste_com.codes == ["38153", "38154"]


# ── Défilement de la fenêtre ──────────────────────────────────────
def test_fenetre_defilante_sur_petit_ecran(app):
    app.deiconify()
    app.geometry("900x420")
    app.update()
    assert app.zone.defile() and app.zone.barre.winfo_ismapped()
    app.zone.canvas.yview_moveto(0)
    app.zone.voir(app.btn_maj)                           # étape 4, hors de vue au départ
    app.update()
    assert app.zone.canvas.yview()[0] > 0
    app.geometry("1100x1400")
    app.update()
    assert not app.zone.defile() and not app.zone.barre.winfo_ismapped()
    app.withdraw()


def test_molette_fait_defiler_la_page_sauf_sur_le_journal(app):
    app.deiconify()
    app.geometry("900x420")
    app.update()
    app.zone.canvas.yview_moveto(0)

    class Evenement:
        def __init__(self, widget, delta):
            self.widget, self.delta = widget, delta

    assert app.zone._molette(Evenement(app.lbl_bilan, -120)) == "break"
    app.update()
    assert app.zone.canvas.yview()[0] > 0
    for i in range(80):
        app.journaliser(f"ligne {i}")
    app.traiter_file()
    app.txt_log.yview_moveto(0)
    assert app.zone._molette(Evenement(app.txt_log, -120)) is None    # le journal défile lui-même
    app.withdraw()


# ── Connexion par le navigateur + API SharePoint ──────────────────
from .test_sharepoint_api import LIEN as LIEN_API, client, serveur  # noqa: E402,F401  (fixtures)


def test_connexion_navigateur_detection_et_audit_par_l_api(app, racine, client, monkeypatch):
    from avf.models.source import API, SourceAudit
    from avf.services import sharepoint
    appels = {}

    def connecter_navigateur(lien, on_etape=None, interactif=True):
        appels["interactif"] = interactif
        on_etape("🔑 identifiez-vous dans la fenêtre du navigateur")
        sharepoint.enregistrer_client(client)
        return SourceAudit(API, LIEN_API, LIEN_API)

    monkeypatch.setattr(module_app.sharepoint, "navigateur_disponible", lambda: True)
    monkeypatch.setattr(module_app.sharepoint, "connecter_navigateur", connecter_navigateur)
    app.var_source.set("sharepoint")
    app._changer_source()
    app.var_lien.set(LIEN_API)
    app._connecter_sharepoint()
    attendre(app)
    try:
        assert appels == {"interactif": True}
        assert "connexion par le navigateur" in app.lbl_source["text"]
        assert "identifiez-vous" in app.txt_log.get("1.0", tk.END)
        assert str(app.btn_det_com["state"]) == "normal"
        app._detecter_communes()                     # par l'API, en arrière-plan
        attendre(app)
        assert app.liste_com.codes == ["13001", "13002", "13003", "2A004", "75020", "75021"]
        app._lancer_audit()
        attendre(app)
        r = app.rapport
        assert r.source.mode == API and r.nb_ids_non_valides == 8
        assert r.commune("13001").chemin == LIEN_API + "/Dep13/13001/audit_13001.xlsx"
        # Saisie d'ID : l'existence du dossier de commune est vérifiée par l'API
        monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: "13001_2 99999_1")
        app._renseigner_ids()
        attendre(app)
        assert app.liste_com.codes == ["13001"]
    finally:
        sharepoint._CLIENTS.clear()


def test_oublier_la_connexion_sharepoint(app, monkeypatch, dialogues):
    from avf.models.source import API, SourceAudit
    oublis = []
    monkeypatch.setattr(module_app.sharepoint, "oublier_connexion", lambda: oublis.append(1))
    app.var_source.set("sharepoint")
    app._changer_source()
    app._source_sp = SourceAudit(API, LIEN, LIEN)
    app._oublier_connexion()
    assert oublis == [1] and app._source() is None and "Collez le lien" in app.lbl_source["text"]


def test_mise_a_jour_avec_la_connexion_navigateur_sans_excel(app, racine, client, monkeypatch, dialogues):
    """Le cas signalé : en connexion navigateur, la mise à jour ne doit pas dépendre d'Excel."""
    from avf.models.source import API, SourceAudit
    from avf.services import sharepoint
    sharepoint.enregistrer_client(client)
    try:
        app.var_source.set("sharepoint")
        app._changer_source()
        app._source_sp = SourceAudit(API, LIEN_API, LIEN_API)
        app._excel = False                                   # « impossible de démarrer Excel »
        monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: "13001_2 13001_5")
        app._renseigner_ids()
        attendre(app)
        app._lancer_audit()
        attendre(app)
        assert app.rapport.nb_ids_non_valides == 3          # 13001_2 + 13001_5 (deux logements)
        app.fenetre_rapport._maj_tout()
        attendre(app)
        attendre(app)                                        # ré-audit par l'API
        journal = app.txt_log.get("1.0", tk.END)
        assert "API SharePoint avec votre connexion (navigateur)" in journal
        assert "valeurs vérifiées sur SharePoint" in journal
        assert app.rapport.nb_ids_non_valides == 0
        confirmation = next(m for t, m in dialogues if t == "Confirmer la mise à jour")
        assert "ligne par ligne (ID imm) : API SharePoint avec votre connexion (navigateur)" in confirmation
        assert not any("Excel" in t for t, _ in dialogues)
    finally:
        sharepoint._CLIENTS.clear()


# ── Vider les communes ou les ID erreur ───────────────────────────
def test_vider_les_communes(app, racine):
    preparer(app, racine)
    assert str(app.btn_vider_com["state"]) == "normal"
    assert str(app.btn_vider_ids["state"]) == "disabled"          # aucun ID saisi
    app._vider_communes()
    assert app.liste_com.codes == [] and charger(app.fichier_config).communes == []
    assert str(app.btn_vider_com["state"]) == "disabled" and str(app.btn_audit["state"]) == "disabled"
    app.traiter_file()
    assert "Liste des communes vidée (6 commune(s))" in app.txt_log.get("1.0", tk.END)


def test_vider_les_id_erreur(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2 13001_5 2a004_2")
    attendre(app)
    assert str(app.btn_vider_ids["state"]) == "normal"
    app._vider_ids()
    assert app.ids_cibles == {} and app.liste_com.codes == []
    assert not app.cadre_perimetre.winfo_manager()
    cfg = charger(app.fichier_config)
    assert cfg.ids_cibles == [] and cfg.communes == []
    assert str(app.btn_vider_ids["state"]) == "disabled"
    app.traiter_file()
    assert "3 ID erreur saisi(s) vidé(s)" in app.txt_log.get("1.0", tk.END)


def test_vider_les_communes_oublie_aussi_les_id(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2")
    attendre(app)
    app._vider_communes()
    assert app.ids_cibles == {} and charger(app.fichier_config).ids_cibles == []


def test_communes_entieres_desactive_vider_les_id(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_ids(app, monkeypatch, "13001_2")
    attendre(app)
    app._effacer_ids()                         # « ✖ Communes entières » : garde les communes
    assert app.liste_com.codes == ["13001"]
    assert str(app.btn_vider_ids["state"]) == "disabled" and str(app.btn_vider_com["state"]) == "normal"


# ── 🆔 ID imm : saisie, audit ciblé, mise à jour ligne par ligne ───
def _saisir_imm(app, monkeypatch, texte):
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: texte)
    app._renseigner_ids("id_imm")
    attendre(app)


def test_id_imm_audit_et_mise_a_jour_d_un_seul_logement(app, racine, monkeypatch, dialogues):
    """13001_5 couvre deux logements : on ne met à jour que celui de la ligne 7."""
    app.var_racine.set(racine)
    app._enregistrer_racine()
    _saisir_imm(app, monkeypatch, f"ID imm\n{imm('13001', 7)}\nID-9999900000AB")
    assert app.liste_com.codes == ["13001"] and app.type_cibles == "id_imm"
    assert "🆔 Audit limité à 1 ID imm sur 1 commune(s)" in app.lbl_perimetre["text"]
    assert any("1 commune(s) sans dossier" in m and "99999" in m for _, m in dialogues)
    cfg = charger(app.fichier_config)
    assert cfg.ids_cibles == [imm("13001", 7)] and cfg.type_ids == "id_imm"

    app._lancer_audit()
    attendre(app)
    r = app.rapport
    assert r.type_ids == "id_imm" and r.nb_lignes == 1 and r.ids_de(NON_VALIDES) == [imm("13001", 7)]
    app.fenetre_rapport.arbre.selection_set("13001")
    app.fenetre_rapport.update()
    app.fenetre_rapport._copier_commune()
    assert app.fenetre_rapport.clipboard_get() == imm("13001", 7)

    app.fenetre_rapport._maj_tout()
    attendre(app)
    attendre(app)
    assert app.rapport.nb_ids_non_valides == 0
    complet = module_app.auditer_communes(["13001"], racine).commune("13001")
    assert imm("13001", 7) not in complet.ids_non_valides
    assert imm("13001", 6) in complet.ids_non_valides       # l'autre logement de 13001_5 : intact
    journal = app.txt_log.get("1.0", tk.END)
    assert "identifiant des lignes : « ID imm »" in journal


def test_id_imm_retrouves_au_lancement(tmp_path, racine, dialogues):
    import json
    cfg = str(tmp_path / "cfg.json")
    json.dump({"racine": racine, "ids_cibles": [imm("13001", 3), imm("2A004", 2)], "type_ids": "id_imm"},
              open(cfg, "w", encoding="utf-8"))
    a = module_app.Application(fichier_config=cfg, fichier_attente=str(tmp_path / "att.json"))
    a.withdraw()
    attendre(a)
    try:
        assert a.type_cibles == "id_imm"
        assert a.ids_cibles == {"13001": [imm("13001", 3)], "2A004": [imm("2A004", 2)]}
        assert "2 ID imm" in a.lbl_perimetre["text"]
    finally:
        a.destroy()
        gc.collect()


def test_dialogue_colonnes_id_imm_et_id_erreur(app):
    from avf.models.colonnes import ColonnesAudit
    d = DialogueColonnes(app, ColonnesAudit(), "ID imm", 1)
    assert d.vars["colonne_id"].get() == "ID imm" and d.vars["id_erreur"].get() == "ID erreur"
    d.vars["colonne_id"].set("N° immeuble")
    d.vars["id_erreur"].set("Erreur")
    d._valider()
    colonnes, colonne_id, _ = d.resultat
    assert (colonnes.id_imm, colonnes.id_erreur, colonne_id) == ("N° immeuble", "Erreur", "N° immeuble")
