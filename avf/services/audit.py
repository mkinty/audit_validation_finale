"""Orchestration de l'audit : une liste de communes → un :class:`RapportAudit`.

Même découplage que le traitement du moteur ``sna`` : l'interface est informée
par trois callbacks facultatifs (progression, étape, résultat), et la fonction
reste testable sans environnement graphique. Une commune en échec n'interrompt
jamais l'audit des suivantes.

Le fichier audit d'une commune est localisé exactement comme le fait la mise à
jour (``sna.services.fichiers.trouver_fichier_audit``) : l'audit et la mise à
jour portent donc toujours sur le même fichier. Une copie de conflit
OneDrive/SharePoint dans le dossier est signalée dès l'audit, puisque la mise
à jour la refusera tant qu'elle n'est pas résolue.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
import os
import time

from sna.services.fichiers import detecter_fichiers_conflit, trouver_fichier_audit
from sna.services.identifiants import normaliser_id

from ..models.audit import AuditCommune, IdErreurAudit, LigneAudit, RapportAudit, StatutAudit
from ..models.colonnes import ColonnesAudit
from .detection import departement_de, normaliser_commune
from .lecture_audit import lire_audit

CallbackProgression = Callable[[int, int, str], None]   # (fait, total, libellé)
CallbackEtape = Callable[[str, str], None]              # (commune, message)
CallbackResultat = Callable[[AuditCommune], None]


def auditer_communes(
    communes: Iterable[str],
    racine: str,
    colonne_id: str = "ID erreur",
    colonnes: ColonnesAudit | None = None,
    ligne_entete: int = 1,
    on_progression: CallbackProgression | None = None,
    on_etape: CallbackEtape | None = None,
    on_resultat: CallbackResultat | None = None,
    doit_arreter: Callable[[], bool] | None = None,
) -> RapportAudit:
    """Audite chaque commune et retourne le rapport complet.

    ``doit_arreter`` (facultatif) est consulté avant chaque commune : il permet
    à l'interface d'interrompre proprement un long audit.
    """
    colonnes = colonnes or ColonnesAudit()
    codes = list(dict.fromkeys(normaliser_commune(c) for c in communes if str(c).strip()))
    rapport = RapportAudit(racine=racine, colonnes=colonnes, colonne_id=colonne_id)
    depart = time.perf_counter()
    total = len(codes)

    _progresser(on_progression, 0, total, "Démarrage…")
    for rang, code in enumerate(codes, start=1):
        if doit_arreter is not None and doit_arreter():
            break
        _progresser(on_progression, rang, total, code)
        resultat = auditer_commune(
            code, racine, colonne_id, colonnes, ligne_entete, on_etape=on_etape,
        )
        rapport.communes.append(resultat)
        if on_resultat is not None:
            on_resultat(resultat)

    rapport.duree = time.perf_counter() - depart
    _progresser(on_progression, len(rapport.communes), total, "Terminé")
    return rapport


def auditer_commune(
    commune: str,
    racine: str,
    colonne_id: str = "ID erreur",
    colonnes: ColonnesAudit | None = None,
    ligne_entete: int = 1,
    on_etape: CallbackEtape | None = None,
) -> AuditCommune:
    """Audite une seule commune (ne lève jamais)."""
    colonnes = colonnes or ColonnesAudit()
    code = normaliser_commune(commune)
    resultat = AuditCommune(commune=code, departement=departement_de(code))

    def etape(message: str) -> None:
        if on_etape is not None:
            on_etape(code, message)

    depart = time.perf_counter()
    etape("recherche du fichier audit…")
    fichier = trouver_fichier_audit(racine, code)
    if fichier is None:
        resultat.statut = StatutAudit.INTROUVABLE
        resultat.message = "aucun fichier audit_*.xlsx dans le dossier de la commune"
        return resultat
    resultat.chemin = fichier

    for conflit in detecter_fichiers_conflit(racine, code):
        resultat.avertissements.append(
            f"copie en conflit : {os.path.basename(conflit['chemin'])} ({conflit['date']}) "
            "— la mise à jour sera refusée tant qu'elle n'est pas résolue"
        )

    etape(f"lecture de {os.path.basename(fichier)}…")
    try:
        lecture = lire_audit(fichier, colonne_id, colonnes, ligne_entete)
    except Exception as exc:  # noqa: BLE001 — l'échec d'une commune est isolé
        resultat.statut = StatutAudit.ERREUR
        resultat.message = str(exc) or exc.__class__.__name__
        resultat.duree = time.perf_counter() - depart
        return resultat

    resultat.ids = regrouper_par_id(lecture.lignes)
    resultat.avertissements.extend(lecture.avertissements)
    resultat.nb_lignes_lues = lecture.nb_lignes_lues
    resultat.duree = time.perf_counter() - depart
    etape(
        f"{resultat.nb_ids} ID erreur · {resultat.nb_renseignes} renseigné(s) · "
        f"{resultat.nb_vides} vide(s)"
    )
    return resultat


def regrouper_par_id(lignes: Iterable[LigneAudit]) -> list[IdErreurAudit]:
    """Regroupe les lignes par ID erreur (forme normalisée), ordre d'apparition.

    La normalisation est celle de la mise à jour (``normaliser_id`` : espaces
    supprimés, majuscules) : un ID vu « 13001_1 » ici sera bien retrouvé par
    le moteur ``sna`` au moment d'écrire.
    """
    groupes: dict[str, IdErreurAudit] = {}
    for ligne in lignes:
        cle = normaliser_id(ligne.id_erreur)
        if not cle:
            continue
        groupes.setdefault(cle, IdErreurAudit(id_erreur=cle)).lignes.append(ligne)
    return list(groupes.values())


def _progresser(callback, fait, total, libelle) -> None:
    if callback is not None:
        callback(fait, total, libelle)
