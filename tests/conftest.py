"""Fixtures : arborescence d'audits ``<racine>/Dep<XX>/<commune>/audit_*.xlsx``."""

from __future__ import annotations

from datetime import datetime
import os

import pytest
from openpyxl import Workbook

ENTETES = ["ID erreur", "Adresse", "Audit", "Date audit", "Auditeur", "Validation finale", "ID imm"]


def imm(id_erreur_ou_insee: str, ligne: int) -> str:
    """ID imm généré par les fixtures pour la ligne Excel ``ligne`` : ``ID-<INSEE><ligne sur 7 chiffres>``."""
    return f"ID-{str(id_erreur_ou_insee)[:5].upper()}{ligne:07d}"


def ecrire_audit(
    chemin: str, lignes: list[list], entetes: list[str] = ENTETES, nom_feuille: str = "Audit",
    autres_feuilles: tuple[str, ...] = (),
) -> str:
    """Écrit un classeur audit : ``lignes`` = valeurs dans l'ordre des en-têtes.

    Si la colonne « ID imm » est attendue mais absente d'une ligne, elle est complétée par
    un ID imm unique (cf. :func:`imm`) quand la ligne porte un ID erreur.
    """
    os.makedirs(os.path.dirname(chemin), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = nom_feuille
    ws.append(entetes)
    rang_imm = entetes.index("ID imm") if "ID imm" in entetes else None
    for numero, ligne in enumerate(lignes, start=2):
        ligne = list(ligne)
        if rang_imm is not None and len(ligne) <= rang_imm:
            ligne += [None] * (rang_imm - len(ligne))
            ligne.append(imm(ligne[0], numero) if ligne and ligne[0] else None)
        ws.append(ligne)
    for nom in autres_feuilles:
        wb.create_sheet(nom)["A1"] = "ne pas lire"
    wb.save(chemin)
    return chemin


@pytest.fixture(autouse=True)
def sans_vrai_excel_a_la_demande(request, monkeypatch):
    """La mise à jour démarre Excel à la demande pour un fichier ouvert ou refusé (423) :
    hors tests marqués ``excel``, jamais de vrai Excel lancé contre un faux serveur ou un
    fichier de test (une fabrique factice passée en ``session_factory`` reste utilisée)."""
    if request.node.get_closest_marker("excel"):
        return
    from avf.services import mise_a_jour
    monkeypatch.setattr(mise_a_jour, "excel_disponible", lambda: False)


@pytest.fixture
def racine(tmp_path):
    """Arborescence de test :

    Dep13/13001  : 5 ID, 6 lignes — 1 ligne validée (13001_1), 5 non valides :
                   13001_2 tout vide · 13001_3 Audit KO mais date et VF (espaces) vides
                   · 13001_4 VF « Validé » ≠ OK · 13001_5 sur deux lignes, aucune valide
    Dep13/13002  : 2 ID, tous validés (dont une date saisie en texte, « ok  »)
    Dep13/13003  : dossier sans fichier audit
    Dep2A/2A004  : 2 ID non valides, onglet « AUDIT » qui n'est pas le premier
    Dep75/75020  : colonne « Validation finale » absente
    Dep75/75021  : deux fichiers audit (copie de conflit)
    """
    base = tmp_path / "audit_SNA"
    r = str(base)
    ecrire_audit(
        os.path.join(r, "Dep13", "13001", "audit_13001.xlsx"),
        [
            ["13001_1", "1 rue A", "OK", datetime(2026, 3, 12), "Alice", "OK"],
            ["13001_2", "2 rue A", None, None, None, None],
            ["13001_3", "3 rue A", "KO", None, "Bob", "   "],
            ["13001_4", "4 rue A", "OK", datetime(2026, 1, 5), "Alice", "Validé"],
            ["13001_5", "5 rue A", None, None, None, "OK"],
            ["13001_5", "5 bis rue A", None, None, None, None],
            [None, "ligne sans ID", None, None, None, None],
        ],
    )
    ecrire_audit(
        os.path.join(r, "Dep13", "13002", "audit_13002.xlsx"),
        [
            ["13002_1", "x", "OK", datetime(2026, 2, 1), "Carl", "OK"],
            ["13002_2", "y", "ok", "15/02/2026", "Carl", "ok "],   # date en texte, casse/espaces
        ],
    )
    os.makedirs(os.path.join(r, "Dep13", "13003"))
    wb_2a = os.path.join(r, "Dep2A", "2A004", "audit_2A004.xlsx")
    os.makedirs(os.path.dirname(wb_2a))
    wb = Workbook()
    wb.active.title = "Synthèse"
    wb.active["A1"] = "ID erreur"
    ws = wb.create_sheet("AUDIT ")
    ws.append(ENTETES)
    ws.append(["2A004_1", "a", None, None, None, None, imm("2A004", 2)])
    ws.append(["2a004_2", "b", None, None, None, "", imm("2A004", 3)])
    wb.save(wb_2a)
    ecrire_audit(
        os.path.join(r, "Dep75", "75020", "audit_75020.xlsx"),
        [["75020_1", "z", "OK", None, "Dan"]],
        entetes=ENTETES[:5] + ["ID imm"],
    )
    ecrire_audit(
        os.path.join(r, "Dep75", "75021", "audit_75021.xlsx"),
        [["75021_1", "z", None, None, None, None]],
    )
    ecrire_audit(
        os.path.join(r, "Dep75", "75021", "audit_75021-DESKTOP-AB12CD3.xlsx"),
        [["75021_1", "z", None, None, None, None]],
    )
    os.makedirs(os.path.join(r, "Divers"))  # ignoré : pas un dossier Dep
    return r
