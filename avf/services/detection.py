"""Détection et saisie des communes à auditer.

Arborescence :

    <racine>/Dep<XX>/<commune>/audit_*.xlsx

où ``XX`` correspond aux deux premiers caractères du code INSEE (chiffres ou
lettres : ``Dep13``, ``Dep2A``, ``Dep97``…).

L'utilisateur ne choisit **que des communes**, de deux façons :

    - **détecter** : toutes les communes présentes sous la racine, quel que
      soit leur dossier ``Dep<XX>`` ;
    - **renseigner** : des codes INSEE saisis à la main, normalisés puis
      confrontés au disque (codes introuvables signalés, jamais ignorés en
      silence). Le dossier ``Dep<XX>`` est déduit du code.

Aucune fonction ne lève sur un dossier illisible : il est simplement ignoré.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
import os
import re

from sna.services.fichiers import PREFIXE_DEP, dossier_commune
from sna.services.identifiants import codes_sans_id, groupes_depuis_texte

# Séparateurs acceptés dans une saisie : virgule, point-virgule, espaces, retours.
_SEPARATEURS = re.compile(r"[\s,;|]+")


@dataclass(frozen=True)
class CommuneDetectee:
    """Une commune présente sur le disque."""

    code: str
    departement: str
    dossier: str
    a_fichier: bool = True


def departement_de(commune: str) -> str:
    """Département déduit d'un code INSEE : ``"13001"`` → ``"13"``."""
    return str(commune or "").strip().upper()[:2]


def normaliser_commune(saisie: str) -> str:
    return "".join(str(saisie or "").split()).upper()


def parser_communes(texte: str) -> list[str]:
    """Codes communes saisis, normalisés, sans doublons, dans l'ordre."""
    codes = (normaliser_commune(m) for m in _SEPARATEURS.split(texte or ""))
    return list(dict.fromkeys(c for c in codes if c))


def lister_communes(racine: str) -> list[CommuneDetectee]:
    """Toutes les communes ``<racine>/Dep<XX>/<commune>``, triées par code.

    Les dossiers de la racine qui ne commencent pas par ``Dep`` (casse
    ignorée) sont ignorés. ``a_fichier`` indique si le dossier de la commune
    contient au moins un ``audit_*.xlsx``.
    """
    communes: dict[str, CommuneDetectee] = {}
    for dossier_dep in _dossiers_departements(racine):
        for nom in _lister(dossier_dep):
            chemin = os.path.join(dossier_dep, nom)
            code = normaliser_commune(nom)
            if os.path.isdir(chemin) and code not in communes:
                communes[code] = CommuneDetectee(
                    code=code, departement=departement_de(code),
                    dossier=chemin, a_fichier=_contient_audit(chemin),
                )
    return sorted(communes.values(), key=lambda c: c.code)


@dataclass
class VerificationCommunes:
    """Résultat de la confrontation d'une saisie de communes au disque."""

    trouvees: list[CommuneDetectee] = field(default_factory=list)
    absentes: list[str] = field(default_factory=list)


def verifier_communes(racine: str, codes: Iterable[str],
                       existe: Callable[[str], bool] = os.path.isdir) -> VerificationCommunes:
    """Confronte des codes communes saisis au disque.

    Une commune est trouvée si ``<racine>/Dep<XX>/<commune>`` existe, ``XX``
    étant les deux premiers caractères du code (cf. complete_excel_files).
    """
    resultat = VerificationCommunes()
    for code in codes:
        dossier = dossier_commune(racine, code)
        if existe(dossier):
            resultat.trouvees.append(CommuneDetectee(
                code=code, departement=departement_de(code),
                dossier=dossier, a_fichier=_contient_audit(dossier),
            ))
        else:
            resultat.absentes.append(code)
    return resultat


@dataclass
class SaisieIds:
    """ID erreur saisis, répartis par commune et confrontés au disque."""

    groupes: dict[str, list[str]] = field(default_factory=dict)    # communes trouvées → ID
    absents: dict[str, list[str]] = field(default_factory=dict)    # communes sans dossier → ID
    codes_sans_id: list[str] = field(default_factory=list)         # « 13001 » saisi sans « _n »
    trouvees: list[CommuneDetectee] = field(default_factory=list)

    @property
    def nb_ids(self) -> int:
        return sum(len(v) for v in self.groupes.values())

    @property
    def ids(self) -> list[str]:
        return [i for v in self.groupes.values() for i in v]


def verifier_ids(racine: str, texte: str, existe: Callable[[str], bool] = os.path.isdir) -> SaisieIds:
    """Texte libre (ID séparés par espaces, virgules, retours à la ligne…) → ID par commune.

    Même lecture que complete_excel_files (``sna.services.identifiants``) : un ID
    s'écrit ``<insee>_<numéro>`` (``13001_2``, ``2a004_1``) et son préfixe INSEE
    désigne la commune, donc le dossier ``<racine>/Dep<XX>/<insee>``. Les codes
    INSEE saisis seuls sont remontés à part : sans numéro, aucune ligne à auditer.
    """
    saisie = SaisieIds(codes_sans_id=codes_sans_id(texte or ""))
    groupes = groupes_depuis_texte(texte or "")
    verification = verifier_communes(racine, list(groupes), existe)
    saisie.trouvees = verification.trouvees
    saisie.groupes = {c.code: groupes[c.code] for c in verification.trouvees}
    saisie.absents = {code: groupes[code] for code in verification.absentes}
    return saisie


# ── Helpers ───────────────────────────────────────────────────────
def _lister(chemin: str) -> list[str]:
    try:
        return sorted(os.listdir(chemin))
    except OSError:
        return []


def _dossiers_departements(racine: str) -> list[str]:
    """Dossiers ``Dep<XX>`` de la racine (casse du préfixe tolérée)."""
    return [
        os.path.join(racine, nom)
        for nom in _lister(racine)
        if nom[: len(PREFIXE_DEP)].casefold() == PREFIXE_DEP.casefold()
        and len(nom) > len(PREFIXE_DEP)
        and os.path.isdir(os.path.join(racine, nom))
    ]


def _contient_audit(dossier: str) -> bool:
    return any(n.startswith("audit_") and n.endswith(".xlsx") for n in _lister(dossier))
