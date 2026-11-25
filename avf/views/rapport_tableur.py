"""Exports tableur du rapport : CSV (une ligne par ligne auditée) et classeur Excel.

CSV — comme l'outil audit_ppts : séparateur « ; » et UTF-8 avec BOM, ce
qu'attend Excel en paramètres régionaux français pour ouvrir le fichier sans
assistant d'importation, accents compris. Une ligne par ligne d'audit (un ID en
double y figure donc deux fois), et une ligne par commune non auditée.

XLSX — quatre onglets, prêts à filtrer :

    Synthèse        indicateurs, critères, motifs de non-validité, départements
    Communes        une ligne par commune (état, comptes, taux, fichier)
    Lignes          une ligne par ligne auditée (statut, 4 colonnes, motifs)
    ID non valides  un ID par ligne, avec ses motifs — à coller dans la mise à jour

Les valeurs sont écrites comme lues (texte) : l'export montre ce qui a été
contrôlé, sans réinterprétation.
"""

from __future__ import annotations

import csv
import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..models.audit import NON_VALIDES, TOUTES, AuditCommune, LigneAudit, RapportAudit
from . import presentation as p


def entetes_lignes(rapport: RapportAudit) -> list[str]:
    c = rapport.colonnes
    return [
        "Département", "Commune", rapport.colonne_id, "Ligne", "Statut",
        c.audit, c.date_audit, c.auditeur, c.validation, "Motifs", "Fichier",
    ]


def lignes_detail(rapport: RapportAudit, categorie: str = TOUTES) -> list[list]:
    """Lignes du détail : une par ligne auditée de la catégorie, une par commune en échec."""
    lignes = []
    for c in rapport.communes:
        if not c.lisible:
            e = p.etat(c)
            lignes.append([f"Dep{c.departement}", c.commune, "", None, e.libelle,
                           "", "", "", "", c.message, c.chemin])
            continue
        for l in c.lignes_de(categorie):
            lignes.append(_ligne(c, l))
        for i in c.ids_introuvables:          # audit par ID : ID saisi absent du fichier
            lignes.append([f"Dep{c.departement}", c.commune, i, None, "ID introuvable",
                           "", "", "", "", "absent du fichier audit", c.chemin])
    return lignes


def _ligne(c: AuditCommune, l: LigneAudit) -> list:
    return [
        f"Dep{c.departement}", c.commune, l.id_erreur, l.ligne, p.etat_ligne(l).libelle,
        l.audit, l.date_audit, l.auditeur, l.validation, l.motifs, c.chemin,
    ]


# ── CSV ───────────────────────────────────────────────────────────
def generer_csv(rapport: RapportAudit, categorie: str = TOUTES) -> str:
    tampon = io.StringIO()
    ecrivain = csv.writer(tampon, delimiter=";", lineterminator="\r\n")
    ecrivain.writerow(entetes_lignes(rapport))
    for ligne in lignes_detail(rapport, categorie):
        ecrivain.writerow(["" if v is None else v for v in ligne])
    return tampon.getvalue()


# ── XLSX ──────────────────────────────────────────────────────────
_GRAS = Font(bold=True)
_ENTETE = PatternFill("solid", fgColor="E8E8E4")
_VERT = PatternFill("solid", fgColor="E3F4E3")
_AMBRE = PatternFill("solid", fgColor="FFF1D6")
_ROUGE = PatternFill("solid", fgColor="FBE4E4")
_TITRE = Font(bold=True, size=14)


def generer_xlsx(rapport: RapportAudit) -> Workbook:
    """Classeur du rapport (non enregistré : cf. ``services.export.exporter_xlsx``)."""
    wb = Workbook()
    _synthese(wb.active, rapport)
    _communes(wb.create_sheet("Communes"), rapport)
    _lignes(wb.create_sheet("Lignes"), rapport)
    _ids_non_valides(wb.create_sheet("ID non valides"), rapport)
    return wb


def _synthese(ws, r: RapportAudit) -> None:
    ws.title = "Synthèse"
    ws["A1"] = f"Audit des fichiers — {r.colonnes.validation}, {r.colonnes.audit}, " \
               f"{r.colonnes.date_audit}, {r.colonnes.auditeur}"
    ws["A1"].font = _TITRE
    lignes = [
        ("Racine", r.racine), ("Source", r.libelle_source),
        *([("Périmètre", p.perimetre(r).removeprefix("🎯 "))] if r.par_ids else []),
        ("Généré le", f"{r.date:%d/%m/%Y %H:%M}"), ("Durée", p.duree(r.duree)),
        None,
        ("Critères d'une ligne validée", None),
        *((f"  {libelle}", None) for libelle in p.libelles_criteres(r.colonnes).values()),
        None,
        ("Indicateur", "Valeur", "Part"),
        ("Communes auditées", r.nb_communes, None),
        ("  complètes", len(r.communes_completes), None),
        ("  à corriger", len(r.communes_a_corriger), None),
        ("  en échec", len(r.en_echec), None),
        ("Lignes auditées", r.nb_lignes, None),
        ("  validées", r.nb_lignes_valides, _part(r.nb_lignes_valides, r.nb_lignes)),
        ("  non valides", r.nb_lignes_non_valides, _part(r.nb_lignes_non_valides, r.nb_lignes)),
        ("ID erreur", r.nb_ids, None),
        ("  validés", r.nb_ids_valides, _part(r.nb_ids_valides, r.nb_ids)),
        ("  non valides", r.nb_ids_non_valides, _part(r.nb_ids_non_valides, r.nb_ids)),
        None,
        ("Motif de non-validité", "Lignes", "Part des non valides"),
        *((libelle, n, _part(n, r.nb_lignes_non_valides))
          for libelle, n in zip(p.libelles_criteres(r.colonnes).values(), r.par_critere().values())),
        None,
        ("Département", "Communes", "ID erreur", "Lignes", "Validées", "Non valides", "En échec", "Taux validé"),
        *((f"Dep{dep}", s["communes"], s["ids"], s["lignes"], s["valides"], s["non_valides"],
           s["erreurs"], _part(s["valides"], s["lignes"])) for dep, s in r.par_departement().items()),
    ]
    rang = 3
    for valeurs in lignes:
        rang += 1
        if valeurs is None:
            continue
        for col, v in enumerate(valeurs, start=1):
            cellule = ws.cell(rang, col, v)
            if isinstance(v, float):
                cellule.number_format = "0.0%"
        if valeurs[0] in ("Indicateur", "Motif de non-validité", "Département", "Critères d'une ligne validée"):
            for col in range(1, len(valeurs) + 1):
                ws.cell(rang, col).font = _GRAS
                ws.cell(rang, col).fill = _ENTETE
    _largeurs(ws, [34, 14, 20, 12, 12, 12, 12, 12])


def _communes(ws, r: RapportAudit) -> None:
    entetes = ["Département", "Commune", "État", "ID erreur", "ID non valides", "Lignes", "Validées",
               "Non valides", "Taux validé", *p.libelles_criteres(r.colonnes).values(), "Fichier", "Message"]
    ws.append(entetes)
    for c in r.communes:
        e = p.etat(c)
        criteres = c.par_critere()
        ws.append([
            f"Dep{c.departement}", c.commune, e.libelle, c.nb_ids, c.nb_ids_non_valides, c.nb_lignes,
            c.nb_lignes_valides, c.nb_lignes_non_valides, c.taux,
            *criteres.values(), c.chemin, c.message or " · ".join(c.avertissements),
        ])
        ws.cell(ws.max_row, 3).fill = {"complete": _VERT, "a-corriger": _AMBRE, "echec": _ROUGE}.get(e.cle, _ENTETE)
        ws.cell(ws.max_row, 9).number_format = "0.0%"
    _tableau(ws, [12, 10, 18, 10, 13, 9, 10, 11, 11, 20, 20, 20, 24, 60, 50])


def _lignes(ws, r: RapportAudit) -> None:
    ws.append(entetes_lignes(r))
    for valeurs in lignes_detail(r, TOUTES):
        ws.append(valeurs)
        statut = valeurs[4]
        ws.cell(ws.max_row, 5).fill = (
            _VERT if statut == p.LIGNE_VALIDEE.libelle
            else _AMBRE if statut == p.LIGNE_NON_VALIDE.libelle else _ROUGE
        )
    _tableau(ws, [12, 10, 16, 7, 18, 10, 13, 22, 16, 70, 60])


def _ids_non_valides(ws, r: RapportAudit) -> None:
    ws.append(["Département", "Commune", r.colonne_id, "Lignes non valides", "Motifs"])
    for c in r.communes:
        for i in c.ids:
            if not c.lisible or not i.dans(NON_VALIDES):
                continue
            fautives = i.lignes_non_valides
            motifs = list(dict.fromkeys(a.message for l in fautives for a in l.anomalies))
            ws.append([f"Dep{c.departement}", c.commune, i.id_erreur,
                       ", ".join(str(l.ligne) for l in fautives), " · ".join(motifs)])
    _tableau(ws, [12, 10, 16, 18, 90])


def _tableau(ws, largeurs: list[int]) -> None:
    """En-tête en gras, volet figé, filtre automatique, largeurs de colonnes."""
    for cellule in ws[1]:
        cellule.font = _GRAS
        cellule.fill = _ENTETE
        cellule.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    _largeurs(ws, largeurs)


def _largeurs(ws, largeurs: list[int]) -> None:
    for rang, largeur in enumerate(largeurs, start=1):
        ws.column_dimensions[get_column_letter(rang)].width = largeur


def _part(n: int, total: int) -> float | None:
    return n / total if total else None
