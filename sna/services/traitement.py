"""Orchestration : traitement d'un lot de communes.

L'entrée est le dictionnaire ``{insee: [ID erreur, …]}`` construit à partir de
la saisie de l'utilisateur : une entrée = un fichier audit à ouvrir.

Le cœur métier est totalement découplé de l'interface : il communique via trois
callbacks optionnels (progression, étape, résultat) et retourne la liste des
résultats. Appelé sans callbacks, il reste parfaitement testable.

Le callback ``on_etape`` est appelé **pendant** le traitement d'une commune
(recherche du fichier, analyse, chargement, écriture, enregistrement) : c'est
lui qui évite à l'utilisateur d'attendre devant un journal vide.

``session_excel`` (facultatif) fait écrire les règles via une session Excel
réelle plutôt que réécrire le fichier avec openpyxl — seule façon fiable de ne
rien perdre quand un autre utilisateur a le classeur ouvert, ou pendant une
synchronisation SharePoint (cf. ``services.excel_com``). Le repérage rapide
(``services.excel.reperer_ids``) reste utilisé en amont dans tous les cas :
inutile d'ouvrir Excel pour un fichier dont aucun ID demandé n'est présent.

Avant tout traitement d'une commune, une copie de conflit OneDrive/SharePoint
restée dans son dossier (cf. ``services.fichiers.detecter_fichiers_conflit``)
fait échouer la commune explicitement, sans rien ouvrir ni écrire : laquelle
des deux versions garder reste un choix humain, pas quelque chose à deviner.

``paralleles`` (facultatif, 1 par défaut : traitement séquentiel, identique à
complete_excel_files) traite plusieurs communes à la fois sur un pool de
threads. Les communes sont indépendantes (un fichier chacune) ; le gain vient
des entrées/sorties qui se recouvrent (fichiers OneDrive à rapatrier, lecture
rapide du XML, décompression). Avec ``session_excel``, la préparation de
chaque commune (recherche du fichier, copies de conflit, repérage rapide) se
fait en parallèle pendant qu'Excel écrit la précédente ; les écritures Excel,
elles, restent une à une (une seule instance, cf. ``excel_com.SessionExcel``).
``on_progression`` et ``on_resultat`` sont alors appelés depuis les threads du
pool, un à la fois, au fil des communes terminées ; ``on_etape`` peut l'être
depuis plusieurs threads. La liste renvoyée garde l'ordre des communes demandées.

``doit_arreter`` (facultatif) est consulté avant chaque commune : les communes
non commencées sont abandonnées (absentes des résultats), celles en cours vont
à leur terme — un fichier n'est jamais laissé à moitié écrit.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
import os
import threading
import time

from ..models.base import ResultatCommune, Statut
from ..models.regles import ReglesConfig
from .excel import appliquer_regles, reperer_ids
from .fichiers import detecter_fichiers_conflit, trouver_fichier_audit
from .identifiants import normaliser_id

# Signatures des callbacks
CallbackProgression = Callable[[int, int, str], None]   # (fait, total, libellé)
CallbackEtape = Callable[[str, str], None]              # (commune, message)
CallbackResultat = Callable[[ResultatCommune], None]


def traiter_communes(
    groupes: Mapping[str, Sequence[str]],
    config: ReglesConfig,
    racine: str,
    simuler: bool = False,
    on_progression: CallbackProgression | None = None,
    on_resultat: CallbackResultat | None = None,
    on_etape: CallbackEtape | None = None,
    session_excel=None,
    paralleles: int = 1,
    doit_arreter: Callable[[], bool] | None = None,
    prereperage: bool = True,
    localiser: Callable[[str, str], str | None] | None = None,
    conflits: Callable[[str, str], list[dict]] | None = None,
) -> list[ResultatCommune]:
    """Applique les règles aux lignes des ID erreur, commune par commune.

    Args:
        groupes: ``{insee: [ID erreur, …]}`` — cf. ``services.identifiants``.
        config: règles de modification à appliquer.
        racine: dossier racine des fichiers audit.
        simuler: n'écrit rien sur le disque, compte seulement les modifications.
        on_progression: appelé avant chaque commune (facultatif).
        on_etape: appelé à chaque étape interne d'une commune (facultatif).
        on_resultat: appelé après chaque commune avec son résultat (facultatif).
        session_excel: une ``services.excel_com.SessionExcel`` déjà ouverte,
            pour écrire via Excel plutôt que par réécriture directe du
            fichier. ``None`` (défaut) ou ``simuler=True`` → réécriture directe.
        paralleles: nombre de communes traitées en même temps (1 : séquentiel).
        doit_arreter: consulté avant chaque commune pour interrompre le lot.
        prereperage: avec ``session_excel``, vérifier d'abord dans le fichier
            (lecture rapide, sans Excel) qu'un des ID y figure. ``False`` quand
            le fichier parcouru n'est pas celui qu'Excel ouvrira (audit depuis
            SharePoint via une bibliothèque synchronisée : la copie locale peut
            différer du serveur) — Excel est alors toujours ouvert.
        localiser: ``(racine, commune) → fichier`` ; par défaut
            ``trouver_fichier_audit``. Permet de désigner un fichier par son
            adresse SharePoint quand le dossier ne peut pas être parcouru.
        conflits: ``(racine, commune) → copies de conflit`` ; par défaut
            ``detecter_fichiers_conflit``.

    Returns:
        La liste des :class:`ResultatCommune`, dans l'ordre des communes demandées.
    """
    communes = list(groupes)
    total = len(communes)
    _notifier_progression(on_progression, 0, total, "Démarrage...")

    if paralleles > 1 and total > 1:
        resultats = _traiter_en_parallele(
            communes, groupes, config, racine, simuler, on_progression, on_resultat, on_etape,
            session_excel, paralleles, doit_arreter, prereperage, localiser, conflits,
        )
    else:
        resultats = []
        for index, commune in enumerate(communes, start=1):
            if doit_arreter is not None and doit_arreter():
                break
            ids = list(groupes[commune])
            _notifier_progression(
                on_progression, index, total,
                f"{commune} ({len(ids)} ID erreur)",
            )
            resultat = _traiter_une_commune(
                commune, ids, config, racine, simuler, on_etape, session_excel, prereperage,
                localiser, conflits,
            )
            resultats.append(resultat)
            if on_resultat is not None:
                on_resultat(resultat)

    _notifier_progression(on_progression, total, total, "Terminé")
    return resultats


def _traiter_en_parallele(
    communes, groupes, config, racine, simuler, on_progression, on_resultat, on_etape,
    session_excel, paralleles, doit_arreter, prereperage=True, localiser=None, conflits=None,
) -> list[ResultatCommune]:
    """Pool de threads ; résultats dans l'ordre des communes demandées."""
    total = len(communes)
    resultats: list[ResultatCommune | None] = [None] * total
    verrou = threading.Lock()
    faits = 0

    def traiter(rang: int, commune: str) -> None:
        nonlocal faits
        if doit_arreter is not None and doit_arreter():
            return
        ids = list(groupes[commune])
        resultat = _traiter_une_commune(
            commune, ids, config, racine, simuler, on_etape, session_excel, prereperage,
            localiser, conflits,
        )
        with verrou:   # callbacks un à la fois, dans l'ordre de fin
            faits += 1
            resultats[rang] = resultat
            _notifier_progression(on_progression, faits, total, f"{commune} ({len(ids)} ID erreur)")
            if on_resultat is not None:
                on_resultat(resultat)

    with ThreadPoolExecutor(max_workers=min(paralleles, total), thread_name_prefix="sna-commune") as pool:
        for future in [pool.submit(traiter, rang, c) for rang, c in enumerate(communes)]:
            future.result()   # _traiter_une_commune ne lève pas : seule une erreur de callback remonte
    return [r for r in resultats if r is not None]


def _traiter_une_commune(
    commune: str,
    ids: list[str],
    config: ReglesConfig,
    racine: str,
    simuler: bool,
    on_etape: CallbackEtape | None,
    session_excel,
    prereperage: bool = True,
    localiser=None,
    conflits=None,
) -> ResultatCommune:
    """Traite une seule commune et renvoie son résultat (ne lève jamais)."""
    def etape(message: str) -> None:
        if on_etape is not None:
            on_etape(commune, message)

    etape("recherche du fichier audit…")
    try:   # localiser / conflits fournis par l'appelant peuvent passer par le réseau
        fichier = (localiser or trouver_fichier_audit)(racine, commune)
        conflits = (conflits or detecter_fichiers_conflit)(racine, commune) if fichier else []
    except Exception as exc:  # noqa: BLE001 — l'échec d'une commune est isolé
        return ResultatCommune(commune, Statut.ERREUR, message=str(exc) or exc.__class__.__name__,
                               nb_ids_demandes=len(ids))
    if fichier is None:
        return ResultatCommune(
            commune, Statut.INTROUVABLE,
            message="fichier introuvable", nb_ids_demandes=len(ids),
        )

    if conflits:
        detail = ", ".join(
            f"{os.path.basename(c['chemin'])} (conflit du {c['date']})" for c in conflits
        )
        etape(f"copie(s) en conflit détectée(s) — rien n'est écrit : {detail}")
        return ResultatCommune(
            commune, Statut.ERREUR,
            message=f"copie(s) en conflit à résoudre manuellement : {detail}",
            nb_ids_demandes=len(ids),
        )

    etape(f"{os.path.basename(fichier)} ({_taille(fichier)})")
    depart = time.perf_counter()
    try:
        if session_excel is not None and not simuler:
            bilan = _appliquer_via_excel(session_excel, fichier, config, ids, etape, prereperage)
        else:
            bilan = appliquer_regles(fichier, config, ids, simuler=simuler, on_etape=etape)

        if bilan.verrouille:
            return ResultatCommune(
                commune, Statut.ERREUR, message=f"classeur verrouillé par {bilan.verrouille}",
                nb_ids_demandes=len(ids), duree=time.perf_counter() - depart,
            )
        absents = [i for i in ids if normaliser_id(i) not in bilan.ids_trouves]
        return ResultatCommune(
            commune,
            Statut.SUCCES,
            nb_modifications=bilan.nb_modifications,
            detail=bilan.comptes,
            nb_ids_demandes=len(ids),
            nb_lignes=bilan.nb_lignes,
            ids_absents=absents,
            ids_non_verifies=list(bilan.ids_non_verifies),
            duree=time.perf_counter() - depart,
        )
    except Exception as exc:  # noqa: BLE001 — on isole l'échec d'une commune
        return ResultatCommune(
            commune, Statut.ERREUR, message=str(exc), nb_ids_demandes=len(ids),
            duree=time.perf_counter() - depart,
        )


def _appliquer_via_excel(session_excel, fichier, config, ids, etape, prereperage=True):
    """Écrit via Excel, après une pré-vérification bon marché (sans l'ouvrir).

    Si aucun des ID demandés n'apparaît dans le fichier (repérage rapide,
    comme pour la réécriture directe), inutile de payer le coût d'une session
    Excel pour ne rien y faire : le résultat du repérage est renvoyé tel quel.
    """
    if not prereperage:
        return session_excel.appliquer_regles(fichier, config, ids, on_etape=etape)
    pre_bilan = reperer_ids(fichier, config, ids, on_etape=etape)
    if not pre_bilan.ids_trouves:
        etape(f"aucune ligne concernée sur {pre_bilan.nb_lignes_lues} lue(s) — Excel non ouvert")
        return pre_bilan
    return session_excel.appliquer_regles(fichier, config, ids, on_etape=etape)


def _taille(chemin: str) -> str:
    """Taille lisible d'un fichier, pour situer la durée annoncée."""
    try:
        octets = os.path.getsize(chemin)
    except OSError:
        return "taille inconnue"
    if octets >= 1024 * 1024:
        return f"{octets / (1024 * 1024):.1f} Mo"
    return f"{max(octets // 1024, 1)} Ko"


def _notifier_progression(
    callback: CallbackProgression | None, fait: int, total: int, libelle: str
) -> None:
    if callback is not None:
        callback(fait, total, libelle)
