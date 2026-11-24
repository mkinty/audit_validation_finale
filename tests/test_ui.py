"""Interface Tkinter pilotée de bout en bout (boîtes de dialogue neutralisées)."""

from __future__ import annotations

import os
import time
import tkinter as tk

import pytest

from avf.config import charger
from avf.ui import application as module_app
from avf.ui.composants import DialogueColonnes, DialogueSaisie


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


def test_audit_rapport_copie_mise_a_jour_reaudit_export(app, racine, tmp_path):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)

    r = app.rapport
    assert r.nb_vides == 6 and r.nb_renseignes == 4
    assert "6 ID" in app.btn_maj["text"]
    journal = app.txt_log.get("1.0", tk.END)
    assert "13001 — 5 ID" in journal and "75020" in journal

    # Popup ouverte automatiquement après l'audit
    f = app.fenetre_rapport
    assert f is not None and f.winfo_exists()
    f.update()
    assert f.tuiles["vides"][0]["text"] == "6"
    assert set(f.arbre.get_children()) == {"13001", "13002", "2A004", "75020", "75021"}

    # Sélection d'une commune → détail et copie
    f.arbre.selection_set("13001")
    f.update()
    assert list(f.detail.get_children()) == ["13001_2", "13001_3", "13001_5"]
    f._copier_commune()
    assert f.clipboard_get() == "13001_2\n13001_3\n13001_5"
    f.var_separateur.set("virgule")
    f._copier_affiches()
    assert f.clipboard_get().split(", ") == r.tous_ids_vides()

    # Filtre texte sur un ID erreur
    f.var_recherche.set("2a004_2")
    f.update()
    assert f.arbre.get_children() == ("2A004",)
    f.var_recherche.set("")
    f.var_filtre.set("❌ En échec")
    f._remplir()
    assert f.arbre.get_children() == ("75020",)
    f.var_filtre.set("Toutes")
    f._remplir()

    # Tri par nombre de vides décroissant
    f._trier("vides")
    f._trier("vides")
    assert f.arbre.get_children()[0] == "13001"

    # Mise à jour de toutes les communes depuis la popup, puis ré-audit automatique
    f._maj_tout()
    attendre(app)
    attendre(app)  # le ré-audit est relancé à la fin de la mise à jour
    assert app.rapport.ids_vides_par_commune() == {"75021": ["75021_1"]}  # conflit refusé
    assert f.tuiles["vides"][0]["text"] == "1"
    journal = app.txt_log.get("1.0", tk.END)
    assert "Ré-audit" in journal and "conflit" in journal

    # Export HTML
    chemin = app._exporter(str(tmp_path / "rapport.html"), ouvrir=False)
    assert os.path.isfile(chemin)
    assert "75021_1" in open(chemin, encoding="utf-8").read()
    assert charger(app.fichier_config).dossier_export == str(tmp_path)


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
    assert set(app.rapport.ids_vides_par_commune()) == {"13001", "75021"}


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
    assert app.rapport.nb_vides == 3


def test_refus_de_confirmation(app, racine, monkeypatch):
    preparer(app, racine)
    app._lancer_audit()
    attendre(app)
    monkeypatch.setattr(module_app.messagebox, "askyesno", lambda *a, **k: False)
    app._lancer_mise_a_jour(None)
    assert not app.occupe
    app._lancer_audit(["13001"])
    attendre(app)
    assert app.rapport.nb_vides == 3


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
