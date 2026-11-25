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
    assert app.liste_com.codes == ["13001", "2A004"]
    assert app._communes_choisies() == ["13001", "2A004"]
    assert any("88888" in m for _, m in dialogues)


def test_saisie_annulee(app, racine, monkeypatch):
    app.var_racine.set(racine)
    app._enregistrer_racine()
    monkeypatch.setattr(DialogueSaisie, "montrer", lambda self: None)
    app._renseigner_communes()
    assert app.liste_com.codes == []


def _ligne_detail(f, iid):
    return dict(zip([c[0] for c in fenetre_rapport._DETAIL], f.detail.item(iid, "values")))


def test_audit_rapport_copie_mise_a_jour_reaudit_export(app, racine, tmp_path):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)

    r = app.rapport
    assert (r.nb_lignes_valides, r.nb_lignes_non_valides, r.nb_ids_non_valides) == (3, 8, 7)
    assert "7 ID non valides" in app.btn_maj["text"]
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
    assert f.detail["columns"][6] == "vf" and f.detail.heading("vf")["text"] == "Validation finale"
    lignes = [_ligne_detail(f, iid) for iid in f.detail.get_children()]
    assert [(l["id"], l["ligne"]) for l in lignes] == [
        ("13001_2", "3"), ("13001_3", "4"), ("13001_4", "5"), ("13001_5", "6"), ("13001_5", "7"),
    ]
    assert lignes[2]["vf"] == "Validé" and lignes[2]["motifs"] == "Validation finale « Validé » ≠ OK"
    assert lignes[2]["statut"] == "⚠️ Non valide"
    f._copier_commune()
    assert f.clipboard_get() == "13001_2\n13001_3\n13001_4\n13001_5"
    # Lignes de 13001_5 sélectionnées : l'ID n'est copié qu'une fois
    f.detail.selection_set([iid for iid in f.detail.get_children() if iid.startswith("13001_5")])
    f._copier_selection_detail()
    assert f.clipboard_get() == "13001_5"
    f.var_separateur.set("virgule")
    f._copier_affiches()
    assert f.clipboard_get().split(", ") == r.ids_de(NON_VALIDES)

    # Case « Afficher aussi les validées » (comme audit_ppts) : la ligne validée s'ajoute
    assert f.case_tous["text"] == "Afficher aussi les validées" and not f.var_tous.get()
    f.case_tous.invoke()
    lignes = [_ligne_detail(f, i) for i in f.detail.get_children()]
    assert [l["ligne"] for l in lignes] == ["2", "3", "4", "5", "6", "7"]
    assert lignes[0]["id"] == "13001_1" and lignes[0]["statut"] == "✅ Validée"
    assert f.detail.item(f.detail.get_children()[0], "tags") == ("validee",)
    f.var_separateur.set("ligne")
    f._copier_commune()                                   # la copie reste celle des non valides
    assert f.clipboard_get() == "13001_2\n13001_3\n13001_4\n13001_5"
    assert "(7)" in f.btn_copier_affiches["text"]
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
    assert app.rapport.ids_non_valides_par_commune() == {"75021": ["75021_1"]}  # conflit refusé
    assert f.tuiles["non_valides"][0]["text"] == "1"
    journal = app.txt_log.get("1.0", tk.END)
    assert "Ré-audit" in journal and "conflit" in journal

    # Exports HTML, Excel, CSV
    chemin = app._exporter("html", str(tmp_path / "rapport.html"), ouvrir=False)
    assert "75021_1" in open(chemin, encoding="utf-8").read()
    chemin = app._exporter("xlsx", str(tmp_path / "rapport.xlsx"), ouvrir=False)
    assert load_workbook(chemin).sheetnames == ["Synthèse", "Communes", "Lignes", "ID non valides"]
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
    assert app.rapport.commune("13001").nb_ids_non_valides == 4


def test_refus_de_confirmation(app, racine, monkeypatch):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    monkeypatch.setattr(module_app.messagebox, "askyesno", lambda *a, **k: False)
    app._lancer_mise_a_jour(None)
    assert not app.occupe
    app._lancer_audit(["13001"])
    attendre(app)
    assert app.rapport.commune("13001").nb_ids_non_valides == 4


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
    assert app.rapport.nb_ids_non_valides == 7   # rien n'a été écrit


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
