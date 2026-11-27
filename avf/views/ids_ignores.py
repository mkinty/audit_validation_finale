"""Liste texte des ID saisis qui n'ont pas pu être audités, pour les rechercher à part.

Quatre raisons, une section chacune :

1. la commune de l'ID n'a pas de dossier ``<racine>/Dep<XX>/<commune>`` (refusé dès la saisie) ;
2. code INSEE saisi sans numéro d'ID (ID erreur seulement) ;
3. le fichier audit de la commune n'a pas pu être lu (absent, onglet manquant…) ;
4. l'ID est absent du fichier audit de sa commune.

Chaque ID est sur sa propre ligne : ``ID <tab> commune <tab> précision <tab> fichier``. Pour un
ID absent de son fichier, la précision est un diagnostic : présent dans l'autre colonne
d'identifiant, ID proche (faute de frappe probable), trouvé dans le fichier d'une autre
commune, longueur inhabituelle, ou simplement absent. Le fichier
se colle tel quel dans Excel, ou dans « 🆔 Saisir des ID imm » pour relancer une
recherche (les lignes de titre ne contiennent aucun ID).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from ..models.audit import RapportAudit


@dataclass
class IdsIgnores:
    """ID écartés à la saisie (avant tout audit)."""

    nom: str = "ID"                                                # « ID imm » ou « ID erreur »
    sans_dossier: dict[str, list[str]] = field(default_factory=dict)   # commune → ID
    codes_sans_id: list[str] = field(default_factory=list)

    @property
    def nombre(self) -> int:
        return sum(len(v) for v in self.sans_dossier.values()) + len(self.codes_sans_id)


def sections(ignores: IdsIgnores | None, rapport: RapportAudit | None) -> list[tuple[str, list[tuple[str, ...]]]]:
    """``[(titre, [(id, commune, précision)…])…]`` — sections vides omises."""
    resultat: list[tuple[str, list[tuple[str, ...]]]] = []
    if ignores and ignores.sans_dossier:
        resultat.append(("Commune sans dossier <racine>/Dep<XX>/<commune> (ignorés à la saisie)",
                         [(i, c, f"dossier Dep{c[:2]}/{c} introuvable")
                          for c, ids in sorted(ignores.sans_dossier.items()) for i in ids]))
    if ignores and ignores.codes_sans_id:
        resultat.append(("Code INSEE saisi sans numéro d'ID (format <insee>_<numéro>)",
                         [(c, c, "pas de numéro") for c in ignores.codes_sans_id]))
    if rapport is not None and rapport.par_ids:
        illisibles = [(i, c.commune, c.message, c.chemin) for c in rapport.communes if not c.lisible
                      for i in (c.ids_demandes or [])]
        if illisibles:
            resultat.append(("Fichier audit de la commune illisible ou absent", illisibles))
        absents = [(i, c.commune, c.diagnostic(i), c.chemin) for c in rapport.communes if c.lisible
                   for i in c.ids_introuvables]
        if absents:
            resultat.append(("Absent du fichier audit de sa commune (peut être dans le fichier d'une autre "
                             "commune, ou mal saisi)", absents))
    return resultat


def nombre(ignores: IdsIgnores | None, rapport: RapportAudit | None) -> int:
    return sum(len(lignes) for _, lignes in sections(ignores, rapport))


def generer_texte(ignores: IdsIgnores | None, rapport: RapportAudit | None,
                  racine: str = "", date: datetime | None = None) -> str:
    nom = (ignores.nom if ignores else None) or (
        rapport.colonnes.id_imm if rapport is not None and rapport.type_ids == "id_imm" else "ID")
    blocs = sections(ignores, rapport)
    total = sum(len(l) for _, l in blocs)
    entete = [f"{nom} non audités — {total} au total",
              f"Généré le {(date or datetime.now()):%d/%m/%Y à %H:%M}"]
    if racine:
        entete.append(f"Racine : {racine}")
    entete += [f"  · {len(l)} : {titre}" for titre, l in blocs]
    lignes = entete + [""]
    for titre, contenu in blocs:
        lignes += [f"=== {titre} — {len(contenu)} ===", f"{nom}\tCommune\tPrécision\tFichier"]
        lignes += ["\t".join(str(v or "") for v in ligne).rstrip("\t") for ligne in contenu]
        lignes.append("")
    return "\r\n".join(lignes)


def nom_fichier_defaut(date: datetime | None = None) -> str:
    return f"ids_non_audites_{(date or datetime.now()):%Y%m%d_%H%M}.txt"
