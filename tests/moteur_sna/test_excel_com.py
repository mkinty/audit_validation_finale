"""Écriture via Excel (COM). Ignorés si Excel / pywin32 ne sont pas disponibles.

Ces tests pilotent un vrai Excel : plus lents que le reste de la suite, et
chacun ouvre puis referme sa propre instance (jamais celle de l'utilisateur).
"""

from __future__ import annotations

from contextlib import contextmanager
import time

import pytest
from openpyxl import Workbook, load_workbook

from sna.models.regles import Action, Condition, RegleModification, ReglesConfig, SourceValeur
from sna.services.colonnes_excel import OngletIntrouvable
from sna.services.excel import BilanFichier, ColonneIntrouvable
from sna.services.excel_com import (
    ExcelBloque, SessionExcel, _FilExcel, _pid_de, _terminer_de_force,
    _terminer_si_toujours_ouvert, excel_disponible,
)

pytestmark = pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")


def _config(*regles, ligne_entete=1, colonne_id="ID erreur"):
    return ReglesConfig(regles=list(regles), ligne_entete=ligne_entete, colonne_id=colonne_id)


def _regle(colonne="Statut", valeur="CORRIGE", code="R1", source=SourceValeur.VALEUR, conditions=None):
    return RegleModification(
        code=code, action=Action(colonne=colonne, source=source, valeur=valeur),
        conditions=conditions or [],
    )


def _classeur(chemin, entetes, lignes, feuille="Audit"):
    wb = Workbook()
    ws = wb.active
    ws.title = feuille
    for col, nom in enumerate(entetes, start=1):
        ws.cell(1, col, nom)
    for i, ligne in enumerate(lignes, start=2):
        for col, valeur in enumerate(ligne, start=1):
            ws.cell(i, col, valeur)
    wb.save(chemin)


@contextmanager
def _collegue_ouvre(chemin: str):
    """Simule un collègue qui garde ``chemin`` ouvert, le temps du bloc.

    Fermé par PID à la sortie plutôt que via les méthodes COM (``Close``,
    ``Quit``) : le proxy d'une session ouverte avant une autre ``SessionExcel``
    imbriquée sur le même thread se retrouve déconnecté une fois celle-ci
    sortie (``CoUninitialize`` partage l'appartement du thread) — fermer le
    *processus* par son identifiant, comme le fait déjà
    ``SessionExcel.__exit__`` pour elle-même, reste fiable dans tous les cas.
    """
    import pythoncom
    import win32com.client as wc

    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        pid = _pid_de(xl)
        xl.Workbooks.Open(
            chemin, 0, False, None, None, None, True, None, None, False, False, None, False,
        )
        yield
    finally:
        _terminer_si_toujours_ouvert(pid, attente_s=1)
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001 — l'apartement a pu être recyclé entre-temps
            pass


# ── Écriture et vérification ───────────────────────────────────────
def test_ecriture_et_verification(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(
        chemin, ["ID erreur", "Statut", "Validation finale"],
        [["13001_1", "à revoir", "OK"], ["13001_2", "autre", ""]],
    )
    config = _config(_regle(conditions=[Condition("Validation finale", "non vide")]))

    messages = []
    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, config, ["13001_1", "13001_2", "13001_9"], on_etape=messages.append)

    assert bilan.comptes == {"R1": 1}            # seul 13001_1 a « Validation finale » renseignée
    assert bilan.ids_trouves == {"13001_1", "13001_2"}
    assert bilan.verrouille is None
    assert bilan.ids_non_verifies == []
    assert bilan.ouverture == "fichier local OneDrive"      # pas de montage SharePoint en test
    assert any("enregistrement" in m for m in messages)

    ws = load_workbook(chemin)["Audit"]
    assert ws["B2"].value == "CORRIGE"
    assert ws["B3"].value == "autre"             # condition non vérifiée : inchangée


def test_recopie_d_une_colonne(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(
        chemin, ["ID erreur", "Statut", "Adresse validée"],
        [["13001_1", "ancien", "12 rue de la Paix"]],
    )
    config = _config(_regle(colonne="Statut", valeur="Adresse validée", source=SourceValeur.COLONNE))

    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    assert load_workbook(chemin)["Audit"]["B2"].value == "12 rue de la Paix"


def test_date_ecrite_et_formatee(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Date audit"], [["13001_1", None]])
    config = _config(_regle(colonne="Date audit", valeur="12/03/2026"))

    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.comptes == {"R1": 1}
    cellule = load_workbook(chemin)["Audit"]["B2"]
    assert cellule.value.strftime("%d/%m/%Y") == "12/03/2026"
    assert "d" in cellule.number_format.lower() and "y" in cellule.number_format.lower()


def test_deuxieme_passage_ne_reecrit_rien(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "à revoir"]])
    config = _config(_regle())

    with SessionExcel() as session:
        session.appliquer_regles(chemin, config, ["13001_1"])
        bilan2 = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan2.comptes == {"R1": 0}


def test_filtre_et_masquage_retires_avant_ecriture(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "à revoir"], ["13001_2", "x"]])
    wb = load_workbook(chemin)
    ws = wb["Audit"]
    ws.auto_filter.ref = "A1:B3"
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["B"].hidden = True
    wb.save(chemin)
    config = _config(_regle())

    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.visibilite_restauree is True
    relu = load_workbook(chemin)["Audit"]
    assert relu.auto_filter.ref is None
    assert relu.row_dimensions[3].hidden is False
    assert relu.column_dimensions["B"].hidden is False


# ── Verrou (classeur ouvert ailleurs) ──────────────────────────────
def test_fichier_ouvert_dans_un_autre_excel_est_ecrit_dedans(tmp_path):
    """Ajout (cf. PROVENANCE.md) : le classeur est déjà ouvert en écriture dans une autre instance
    d'Excel de la session (l'Excel de l'utilisateur). Notre instance ne l'obtient qu'en lecture
    seule : les cellules sont alors écrites dans le classeur ouvert, qui est enregistré."""
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "à revoir"]])
    config = _config(_regle())

    with _collegue_ouvre(chemin):
        with SessionExcel() as session:
            bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert not bilan.verrouille and bilan.comptes == {"R1": 1}
    assert bilan.ouverture == "classeur déjà ouvert dans votre Excel"
    assert bilan.ids_non_verifies == []


def test_fichier_ouvert_ailleurs_introuvable_est_signale_sans_rien_ecrire(tmp_path, monkeypatch):
    """Ouvert par quelqu'un d'autre (autre poste, introuvable ici) : verrouillé, rien écrit."""
    from sna.services import excel_com
    monkeypatch.setattr(excel_com, "_classeur_ouvert_ailleurs", lambda *a, **k: None)
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "à revoir"]])
    avant = open(chemin, "rb").read()
    config = _config(_regle())

    with _collegue_ouvre(chemin):
        with SessionExcel() as session:
            bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.verrouille and bilan.comptes == {"R1": 0}
    assert open(chemin, "rb").read() == avant


# ── Erreurs structurelles ──────────────────────────────────────────
def test_colonne_id_absente_leve_une_erreur(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["Autre colonne", "Statut"], [["x", "y"]])
    config = _config(_regle())

    with SessionExcel() as session, pytest.raises(ColonneIntrouvable):
        session.appliquer_regles(chemin, config, ["13001_1"])


def test_onglet_audit_absent_leve_une_erreur(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "x"]], feuille="Feuil1")
    config = _config(_regle())

    with SessionExcel() as session, pytest.raises(OngletIntrouvable):
        session.appliquer_regles(chemin, config, ["13001_1"])


def test_aucune_regle_active_n_ouvre_pas_le_classeur(tmp_path):
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "x"]])
    config = _config()   # aucune règle

    with SessionExcel() as session:
        bilan = session.appliquer_regles(chemin, config, ["13001_1"])

    assert bilan.comptes == {} and bilan.ouverture is None


# ── Garde-fou anti-blocage (boîte de dialogue de fusion sans réponse) ──
def test_fil_excel_signale_un_blocage_sans_attendre_la_fonction():
    fil = _FilExcel(visible=False, delai_occupe_s=30)
    try:
        debut = time.monotonic()
        with pytest.raises(ExcelBloque):
            fil.executer(lambda: time.sleep(1), delai_s=0.2)
        assert time.monotonic() - debut < 1    # abandonné bien avant la seconde du faux blocage
    finally:
        _terminer_de_force(fil.pid)
        # Laisse la fonction artificiellement lente se terminer, puis demande au
        # fil de se nettoyer (CoUninitialize) plutôt que de le laisser bloqué à
        # jamais sur sa file — un appartement COM qui ne se libère jamais finit
        # par perturber les tests suivants dans le même processus.
        time.sleep(1.5)
        fil.arreter_proprement()


def test_blocage_redemarre_excel_et_reessaie(tmp_path):
    """Un appel qui dépasse le délai ne gèle jamais l'appelant : Excel est tué
    et remplacé, et la tentative suivante repart sur une instance saine.

    Marge large entre le faux blocage (6 s) et ``delai_blocage_s`` (3 s) :
    la tentative suivante doit, elle, avoir largement le temps de faire du
    vrai travail Excel (ouvrir, écrire, enregistrer) sans être elle-même
    confondue avec un blocage.
    """
    chemin = str(tmp_path / "audit.xlsx")
    _classeur(chemin, ["ID erreur", "Statut"], [["13001_1", "à revoir"]])
    config = _config(_regle())

    with SessionExcel(delai_blocage_s=3) as session:
        pid_initial = session._fil.pid
        original = session._appliquer_une_fois
        appels = []

        def parfois_bloque(*args, **kwargs):
            appels.append(1)
            if len(appels) == 1:
                time.sleep(6)   # dépasse delai_blocage_s : déclenche ExcelBloque côté appelant
                # Ne jamais rappeler ``original`` ici : par le temps que ce premier
                # appel se réveille, l'appelant a déjà abandonné et redémarré Excel
                # sur un fil différent — toucher du COM ici reviendrait à appeler
                # l'appartement d'un AUTRE fil depuis celui-ci.
                return BilanFichier(comptes={})
            return original(*args, **kwargs)

        session._appliquer_une_fois = parfois_bloque
        messages = []
        bilan = session.appliquer_regles(chemin, config, ["13001_1"], on_etape=messages.append)

        assert len(appels) == 2
        assert bilan.tentatives == 2
        assert bilan.comptes == {"R1": 1}
        assert session._fil.pid != pid_initial     # Excel a bien été redémarré
        assert any("redémarré" in m or "Excel" in m for m in messages)
