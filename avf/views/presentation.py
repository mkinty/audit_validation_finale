"""Formats d'affichage partagés par le rapport HTML et la fenêtre du rapport.

Les deux vues affichent les mêmes chiffres de la même façon : nombres groupés
par milliers, pourcentages à une décimale, libellés de statut avec icône (la
couleur ne porte jamais seule l'information).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass

from ..models.audit import AuditCommune, LigneAudit, RapportAudit, StatutAudit
from ..models.colonnes import ColonnesAudit

# Espace fine insécable : séparateur de milliers et avant « % » (typographie française).
ESPACE_FINE = " "

SEPARATEURS_COPIE = {
    "ligne": "\n",
    "virgule": ", ",
    "point-virgule": ";",
    "espace": " ",
}


def nombre(valeur: int) -> str:
    """``1284`` → ``« 1 284 »``."""
    return f"{valeur:,}".replace(",", ESPACE_FINE)


def pourcent(taux: float | None) -> str:
    """``0.875`` → ``« 87,5 % »`` ; ``None`` → ``« — »``."""
    if taux is None:
        return "—"
    valeur = round(taux * 100, 1)
    texte = f"{valeur:.0f}" if valeur.is_integer() else f"{valeur:.1f}".replace(".", ",")
    return f"{texte}{ESPACE_FINE}%"


def duree(secondes: float) -> str:
    if secondes < 60:
        return f"{secondes:.1f} s".replace(".", ",")
    minutes, reste = divmod(int(round(secondes)), 60)
    return f"{minutes} min {reste:02d} s"


@dataclass(frozen=True)
class Etat:
    """État d'une commune, ou d'une ligne, dans le rapport."""

    cle: str        # complete | a-corriger | echec | vide  —  validee | non-valide
    icone: str
    libelle: str


ETAT_COMPLETE = Etat("complete", "✅", "Complète")
ETAT_A_CORRIGER = Etat("a-corriger", "⚠️", "À corriger")
ETAT_SANS_ID = Etat("vide", "∅", "Aucune ligne")
ETAT_INTROUVABLE = Etat("echec", "⛔", "Fichier introuvable")
ETAT_ERREUR = Etat("echec", "❌", "Erreur de lecture")

LIGNE_VALIDEE = Etat("validee", "✅", "Validée")
LIGNE_NON_VALIDE = Etat("non-valide", "⚠️", "Non valide")


def etat(commune: AuditCommune) -> Etat:
    if commune.statut is StatutAudit.INTROUVABLE:
        return ETAT_INTROUVABLE
    if commune.statut is StatutAudit.ERREUR:
        return ETAT_ERREUR
    if not commune.nb_ids:
        return ETAT_SANS_ID
    return ETAT_COMPLETE if commune.complete else ETAT_A_CORRIGER


def etat_ligne(ligne: LigneAudit) -> Etat:
    return LIGNE_VALIDEE if ligne.valide else LIGNE_NON_VALIDE


def libelles_criteres(colonnes: ColonnesAudit) -> dict[str, str]:
    """``{critère: libellé du critère}`` — ex. ``Audit = OK / KO``."""
    return {
        "audit": f"{colonnes.audit} = {' / '.join(colonnes.valeurs_audit)}",
        "date_audit": f"{colonnes.date_audit} = une date",
        "auditeur": f"{colonnes.auditeur} renseigné",
        "validation": f"{colonnes.validation} = {' / '.join(colonnes.valeurs_validation)}",
        "id_imm": f"{colonnes.id_imm} renseigné",
    }


def texte_ids(ids: Iterable[str], separateur: str = "\n") -> str:
    """Liste d'ID prête à coller (un par ligne par défaut)."""
    return separateur.join(ids)


def resume(rapport: RapportAudit) -> list[str]:
    """Bilan en quelques lignes (journal, en-tête de la fenêtre du rapport)."""
    return [
        f"{nombre(rapport.nb_communes)} commune(s) auditée(s) — "
        f"{nombre(len(rapport.communes_completes))} complète(s), "
        f"{nombre(len(rapport.communes_a_corriger))} à corriger, "
        f"{nombre(len(rapport.en_echec))} en échec",
        f"{nombre(rapport.nb_lignes)} ligne(s) — "
        f"{nombre(rapport.nb_lignes_valides)} validée(s) ({pourcent(rapport.taux)}), "
        f"{nombre(rapport.nb_lignes_non_valides)} non valide(s) · "
        f"{nombre(rapport.nb_ids_non_valides)} {rapport.colonnes.id_imm} non valide(s) sur {nombre(rapport.nb_ids)}",
    ] + ([perimetre(rapport)] if rapport.par_ids else [])


def perimetre(rapport: RapportAudit) -> str:
    """« 🎯 12 ID imm saisis — 2 introuvable(s) : … » (audit restreint à des ID saisis)."""
    introuvables = rapport.ids_introuvables()
    nature = rapport.colonnes.id_imm if rapport.type_ids == "id_imm" else rapport.colonnes.id_erreur
    texte = f"🎯 Audit limité à {nombre(rapport.nb_ids_demandes)} {nature} saisi(s)"
    if introuvables:
        apercu = ", ".join(introuvables[:10]) + ("…" if len(introuvables) > 10 else "")
        texte += f" — {nombre(len(introuvables))} introuvable(s) : {apercu}"
    return texte


# ── Journal de la mise à jour : pourquoi un ID imm n'a pas été mis à jour ─
_CONSEILS = (
    ("absent de la colonne", "vérifiez l'ID imm : faute de frappe, ligne supprimée ou renumérotée, ou "
                             "ligne rangée dans le fichier d'une autre commune"),
    ("ID erreur introuvable", "vérifiez l'ID : faute de frappe, ligne supprimée, ou autre commune"),
    ("fichier audit introuvable", "le fichier audit_<commune>.xlsx n'est pas à l'emplacement attendu "
                                  "(Dep<XX>/<commune>/)"),
    ("non confirmée", "Excel a enregistré mais la relecture n'a pas retrouvé la valeur (fusion de "
                      "co-édition en cours) : ID remis en attente, rejoué à la prochaine mise à jour"),
    ("conflit", "une copie de conflit OneDrive existe : fusionnez-la ou supprimez-la, puis relancez"),
    ("verrouillé", "ID remis en attente : relancez la mise à jour une fois le fichier libéré"),
)


def conseil_refus(raison: str) -> str:
    """Que faire pour débloquer un ID refusé pour ``raison`` (``""`` si rien à ajouter)."""
    bas = raison.lower()
    return next((c for cle, c in _CONSEILS if cle.lower() in bas), "")


def refus_mise_a_jour(echecs: dict[str, str], limite: int = 15) -> list[str]:
    """Lignes de journal : les ID imm non mis à jour, regroupés par raison (la plus fréquente d'abord)."""
    par_raison: dict[str, list[str]] = {}
    for identifiant, raison in echecs.items():
        par_raison.setdefault(raison or "raison inconnue", []).append(identifiant)
    lignes = []
    for raison, ids in sorted(par_raison.items(), key=lambda e: (-len(e[1]), e[0])):
        ids = sorted(ids)
        extrait = ", ".join(ids[:limite]) + (f" … (+{len(ids) - limite})" if len(ids) > limite else "")
        texte = f"⛔ {nombre(len(ids))} ID imm non mis à jour — {raison}\n   ↳ {extrait}"
        conseil = conseil_refus(raison)
        if conseil:
            texte += f"\n   💡 {conseil}"
        lignes.append(texte)
    return lignes


def encore_non_valides(rapport: RapportAudit, ids: Iterable[str], limite: int = 30) -> list[str]:
    """Après le ré-audit : les ID imm écrits qui restent non valides, avec leurs motifs."""
    from sna.services.identifiants import normaliser_id

    cles = {normaliser_id(i) for i in ids} - {""}
    trouves = []
    for commune in rapport.lues:
        for groupe in commune.ids:
            if groupe.identifiant and not groupe.valide and normaliser_id(groupe.identifiant) in cles:
                motifs = "; ".join(dict.fromkeys(l.motifs for l in groupe.lignes_non_valides))
                trouves.append(f"   ↳ {commune.commune} · {groupe.identifiant} — {motifs}")
    if not trouves:
        return []
    entete = (f"⚠ {nombre(len(trouves))} ID imm mis à jour mais encore non valide(s) au ré-audit — la "
              "règle ne s'applique pas à ces lignes (condition non remplie) ou une autre colonne reste à corriger :")
    return [entete, *trouves[:limite]] + ([f"   … (+{len(trouves) - limite})"] if len(trouves) > limite else [])
