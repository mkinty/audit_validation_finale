"""Fixtures : arborescence d'audits ``<racine>/Dep<XX>/<commune>/audit_*.xlsx``."""

from __future__ import annotations

from datetime import datetime
import os

import pytest
from openpyxl import Workbook

ENTETES = ["ID erreur", "Adresse", "Audit", "Date audit", "Auditeur", "Validation finale"]


def ecrire_audit(
    chemin: str, lignes: list[list], entetes: list[str] = ENTETES, nom_feuille: str = "Audit",
    autres_feuilles: tuple[str, ...] = (),
) -> str:
    """Écrit un classeur audit : ``lignes`` = valeurs dans l'ordre des en-têtes."""
    os.makedirs(os.path.dirname(chemin), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = nom_feuille
    ws.append(entetes)
    for ligne in lignes:
        ws.append(ligne)
    for nom in autres_feuilles:
        wb.create_sheet(nom)["A1"] = "ne pas lire"
    wb.save(chemin)
    return chemin


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
    ws.append(["2A004_1", "a", None, None, None, None])
    ws.append(["2a004_2", "b", None, None, None, ""])
    wb.save(wb_2a)
    ecrire_audit(
        os.path.join(r, "Dep75", "75020", "audit_75020.xlsx"),
        [["75020_1", "z", "OK", None, "Dan"]],
        entetes=ENTETES[:5],
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
