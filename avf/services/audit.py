"""Orchestration de l'audit : une liste de communes → un :class:`RapportAudit`.

Même découplage que le traitement du moteur ``sna`` : l'interface est informée
par trois callbacks facultatifs (progression, étape, résultat), et la fonction
reste testable sans environnement graphique. Une commune en échec n'interrompt
jamais l'audit des suivantes.

Chaque ligne lue est contrôlée sur ses quatre colonnes (``services.controle``) :
validée, ou non valide avec ses motifs.

Le fichier audit d'une commune est localisé exactement comme le fait la mise à
jour (``sna.services.fichiers.trouver_fichier_audit``) : l'audit et la mise à
jour portent donc toujours sur le même fichier. Une copie de conflit
OneDrive/SharePoint dans le dossier est signalée dès l'audit, puisque la mise
à jour la refusera tant qu'elle n'est pas résolue.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
import os
import threading
import time

from sna.services.identifiants import normaliser_id

from ..models.audit import (
    PAR_ID_ERREUR, PAR_ID_IMM, AuditCommune, IdAudit, LigneAudit, RapportAudit, StatutAudit,
)
from ..models.colonnes import ColonnesAudit
from ..models.source import SourceAudit
from .controle import controler
from .sharepoint import AccesFichiers
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
    paralleles: int = 1,
    source: SourceAudit | None = None,
    acces: AccesFichiers | None = None,
    ids: Mapping[str, Sequence[str]] | None = None,
    type_ids: str = PAR_ID_ERREUR,
) -> RapportAudit:
    """Audite chaque commune et retourne le rapport complet.

    ``doit_arreter`` (facultatif) est consulté avant chaque commune : il permet
    à l'interface d'interrompre proprement un long audit.

    ``paralleles`` > 1 lit plusieurs fichiers à la fois (pool de threads) : la
    lecture est dominée par les entrées/sorties (rapatriement OneDrive,
    décompression), qui se recouvrent. Les communes du rapport restent dans
    l'ordre demandé ; ``on_progression`` et ``on_resultat`` sont appelés un à
    la fois, au fil des communes terminées.

    ``source`` (facultatif) décrit d'où viennent les fichiers — dossier local
    ou SharePoint, cf. ``services.sharepoint`` — et ``acces`` comment les lire
    (déjà ouvert par l'appelant ; par défaut : lecture directe).

    ``ids`` (facultatif) restreint l'audit à des ID saisis — **ID erreur** ou **ID imm**
    selon ``type_ids`` :
    ``{commune: [ID erreur]}`` (cf. ``sna.services.identifiants.groupes_depuis_texte``).
    Seules leurs lignes sont contrôlées ; un ID absent du fichier est signalé
    (``AuditCommune.ids_introuvables``). Une commune de ``communes`` absente de
    ``ids`` est auditée entièrement.
    """
    colonnes = colonnes or ColonnesAudit()
    codes = list(dict.fromkeys(normaliser_commune(c) for c in communes if str(c).strip()))
    rapport = RapportAudit(racine=source.affichage if source else racine, colonnes=colonnes,
                           colonne_id=colonne_id, source=source)
    depart = time.perf_counter()
    total = len(codes)

    cibles = {normaliser_commune(c): list(v) for c, v in ids.items()} if ids else {}

    def auditer(code: str) -> AuditCommune:
        return auditer_commune(code, racine, colonne_id, colonnes, ligne_entete, on_etape=on_etape,
                               acces=acces, ids=cibles.get(code), type_ids=type_ids)

    _progresser(on_progression, 0, total, "Démarrage…")
    if paralleles > 1 and total > 1:
        rapport.communes = _auditer_en_parallele(
            codes, auditer, paralleles, on_progression, on_resultat, doit_arreter,
        )
    else:
        for rang, code in enumerate(codes, start=1):
            if doit_arreter is not None and doit_arreter():
                break
            _progresser(on_progression, rang, total, code)
            resultat = auditer(code)
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
    acces: AccesFichiers | None = None,
    ids: Sequence[str] | None = None,
    type_ids: str = PAR_ID_ERREUR,
) -> AuditCommune:
    """Audite une seule commune (ne lève jamais).

    ``acces`` : comment lire le fichier trouvé (version serveur en source
    SharePoint via Excel) et quelle adresse afficher.

    ``ids`` : seulement ces ID erreur (toutes leurs lignes) ; ``None`` : toute la commune.
    """
    colonnes = colonnes or ColonnesAudit()
    acces = acces or AccesFichiers()
    code = normaliser_commune(commune)
    resultat = AuditCommune(commune=code, departement=departement_de(code))
    if ids is not None:
        resultat.ids_demandes = list(dict.fromkeys(normaliser_id(i) for i in ids if normaliser_id(i)))
        resultat.type_ids = type_ids

    def etape(message: str) -> None:
        if on_etape is not None:
            on_etape(code, message)

    depart = time.perf_counter()
    etape("recherche du fichier audit…")
    try:   # en source SharePoint, la recherche passe par le réseau : elle peut échouer
        fichier = acces.trouver(racine, code)
        conflits = acces.conflits(racine, code) if fichier is not None else []
    except Exception as exc:  # noqa: BLE001 — l'échec d'une commune est isolé
        resultat.statut = StatutAudit.ERREUR
        resultat.message = str(exc) or exc.__class__.__name__
        return resultat
    if fichier is None:
        resultat.statut = StatutAudit.INTROUVABLE
        resultat.message = "aucun fichier audit_*.xlsx dans le dossier de la commune"
        return resultat
    resultat.chemin = acces.adresse(fichier)

    for conflit in conflits:
        resultat.avertissements.append(
            f"copie en conflit : {os.path.basename(conflit['chemin'])} ({conflit['date']}) "
            "— la mise à jour sera refusée tant qu'elle n'est pas résolue"
        )

    etape(f"lecture de {os.path.basename(fichier)}…")
    try:
        lecture = lire_audit(acces.lisible(fichier), colonne_id, colonnes, ligne_entete)
    except FileNotFoundError as exc:      # mode URL : rien à cette adresse sur SharePoint
        resultat.statut = StatutAudit.INTROUVABLE
        resultat.message = str(exc)
        resultat.duree = time.perf_counter() - depart
        return resultat
    except Exception as exc:  # noqa: BLE001 — l'échec d'une commune est isolé
        resultat.statut = StatutAudit.ERREUR
        resultat.message = str(exc) or exc.__class__.__name__
        resultat.duree = time.perf_counter() - depart
        return resultat

    lignes = lecture.lignes
    if resultat.ids_demandes is not None:
        voulus = set(resultat.ids_demandes)
        cle = (lambda l: normaliser_id(l.id_imm)) if type_ids == PAR_ID_IMM else (lambda l: normaliser_id(l.id_erreur))
        lignes = [l for l in lignes if cle(l) in voulus]
        trouves = {cle(l) for l in lignes}
        resultat.ids_introuvables = [i for i in resultat.ids_demandes if i not in trouves]
        if resultat.ids_introuvables:
            nature = colonnes.id_imm if type_ids == PAR_ID_IMM else colonnes.id_erreur
            resultat.avertissements.append(
                f"{len(resultat.ids_introuvables)} {nature} introuvable(s) dans le fichier : "
                + ", ".join(resultat.ids_introuvables)
                + (" — un ID imm peut se trouver dans le fichier d'une autre commune que celle de son "
                   "préfixe" if type_ids == PAR_ID_IMM else "")
            )
    resultat.ids = regrouper_par_id(controler(lignes, colonnes))
    resultat.avertissements.extend(_qualite_identifiants(resultat, colonnes))
    resultat.avertissements.extend(lecture.avertissements)
    resultat.nb_lignes_lues = lecture.nb_lignes_lues
    resultat.duree = time.perf_counter() - depart
    etape(
        f"{resultat.nb_lignes} ligne(s) · {resultat.nb_lignes_valides} validée(s) · "
        f"{resultat.nb_lignes_non_valides} non valide(s) ({resultat.nb_ids_non_valides} ID erreur)"
    )
    return resultat


def _qualite_identifiants(c: AuditCommune, colonnes: ColonnesAudit) -> list[str]:
    """Anomalies d'identifiants qui comptent pour la mise à jour par ID imm."""
    from .detection import insee_de_id_imm
    messages = []
    doublons = [i.identifiant for i in c.identifiables if i.doublon]
    if doublons:
        messages.append(f"{len(doublons)} {colonnes.id_imm} en double (la mise à jour modifiera chacune "
                        f"de leurs lignes) : " + ", ".join(doublons[:20]) + ("…" if len(doublons) > 20 else ""))
    sans = [l.ligne for l in c.lignes if not normaliser_id(l.id_imm)]
    if sans:
        messages.append(f"{len(sans)} ligne(s) sans {colonnes.id_imm} (non ciblables par la mise à jour) : "
                        "ligne(s) " + ", ".join(map(str, sans[:20])) + ("…" if len(sans) > 20 else ""))
    ailleurs = [i.identifiant for i in c.identifiables
                if insee_de_id_imm(i.identifiant) and insee_de_id_imm(i.identifiant) != c.commune]
    if ailleurs:
        messages.append(f"{len(ailleurs)} {colonnes.id_imm} au code d'une autre commune (à saisir via leur "
                        "commune réelle) : " + ", ".join(ailleurs[:20]) + ("…" if len(ailleurs) > 20 else ""))
    return messages


def regrouper_par_id(lignes: Iterable[LigneAudit]) -> list[IdAudit]:
    """Regroupe les lignes par **ID imm** (forme normalisée), dans l'ordre d'apparition.

    L'ID imm est l'identifiant unique d'une ligne : c'est lui que l'on copie et que la
    mise à jour cible. Un ID imm en double (anomalie) regroupe ses lignes. Les lignes
    sans ID imm forment un groupe sans identifiant : affichées et comptées, mais jamais
    copiées ni mises à jour. La normalisation est celle de la mise à jour
    (``normaliser_id`` : espaces supprimés, majuscules).
    """
    groupes: dict[str, IdAudit] = {}
    for ligne in lignes:
        if not normaliser_id(ligne.id_imm) and not normaliser_id(ligne.id_erreur):
            continue
        cle = normaliser_id(ligne.id_imm)
        groupes.setdefault(cle, IdAudit(identifiant=cle)).lignes.append(ligne)
    return list(groupes.values())


def _auditer_en_parallele(codes, auditer, paralleles, on_progression, on_resultat, doit_arreter):
    """Pool de threads ; résultats dans l'ordre de ``codes`` (communes non commencées omises)."""
    resultats: list[AuditCommune | None] = [None] * len(codes)
    verrou = threading.Lock()
    faits = 0

    def tache(rang: int, code: str) -> None:
        nonlocal faits
        if doit_arreter is not None and doit_arreter():
            return
        resultat = auditer(code)
        with verrou:
            faits += 1
            resultats[rang] = resultat
            _progresser(on_progression, faits, len(codes), code)
            if on_resultat is not None:
                on_resultat(resultat)

    with ThreadPoolExecutor(max_workers=min(paralleles, len(codes)), thread_name_prefix="avf-audit") as pool:
        for future in [pool.submit(tache, rang, code) for rang, code in enumerate(codes)]:
            future.result()
    return [r for r in resultats if r is not None]


def _progresser(callback, fait, total, libelle) -> None:
    if callback is not None:
        callback(fait, total, libelle)
