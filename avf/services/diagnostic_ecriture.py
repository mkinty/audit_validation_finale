"""Diagnostic, en lecture seule, de la mise à jour d'un fichier audit.

Pourquoi un fichier ouvert par un collègue n'est-il pas mis à jour sur ce poste ?
Le diagnostic essaie chaque voie d'écriture **sans rien écrire** et dit laquelle
fonctionne :

1. le fichier : où il est, s'il est extrait (check-out) ou ouvert, et par qui ;
2. l'**API SharePoint** (connexion du navigateur) : session valide, téléchargement
   autorisé par les stratégies de l'organisation. Elle ne peut pas écrire dans un
   fichier ouvert ;
3. **Excel** : il ouvre le fichier par son adresse, en lecture seule, avec le compte
   Office du poste. C'est la condition de la co-édition, seule écriture possible
   dans un fichier ouvert par d'autres.

Il se termine par une conclusion : ce qui se passera pour un fichier ouvert, et
comment le débloquer.
"""

from __future__ import annotations

from collections.abc import Callable
import os
import shutil
import tempfile

from sna.services.excel_com import SessionExcel, excel_disponible
from sna.services.fichiers import trouver_fichier_audit
from sna.services.onedrive import url_sharepoint

from ..models.source import API, LOCAL, URL, SourceAudit
from .ecriture import proprietaire_verrou

OK, KO, INFO = "ok", "ko", "info"


def diagnostiquer_ecriture(
    source: SourceAudit,
    commune: str,
    session_factory: Callable[[], object] | None = None,
    excel_present: Callable[[], bool] | None = None,
    client=None,
) -> list[tuple[str, str]]:
    """``[(état, texte)]`` — ``état`` : ``ok``, ``ko`` ou ``info``. Ne lève jamais, n'écrit rien."""
    lignes: list[tuple[str, str]] = []

    def dire(etat: str, texte: str) -> None:
        lignes.append((etat, texte))

    commune = commune.strip().upper()
    dire(INFO, f"Commune {commune} — source : {source.libelle}")
    adresse, local = _localiser(source, commune, client, dire)
    if not adresse and not local:
        dire(KO, "fichier audit introuvable : rien à diagnostiquer")
        return lignes

    # ── 1. Le fichier ───────────────────────────────────────────────
    api_ok = None
    if source.mode == API:
        api_ok = _diagnostic_api(source, adresse, client, dire)
    if local:
        qui = proprietaire_verrou(local)
        dire(INFO if qui is None else KO, f"fichier local : {local}" + (f" — ouvert par {qui}" if qui else
                                                                         " — pas ouvert sur ce poste"))

    # ── 2. Excel : la co-édition ────────────────────────────────────
    excel_ok = _diagnostic_excel(adresse or local, session_factory, excel_present or excel_disponible, dire)

    # ── 3. Conclusion ───────────────────────────────────────────────
    if excel_ok:
        dire(OK, "CONCLUSION : un fichier ouvert par d'autres sera mis à jour via Excel en co-édition, "
                 "cellule par cellule (sauf classeur ouvert en mode exclusif ou extrait).")
    elif api_ok:
        dire(KO, "CONCLUSION : seuls les fichiers FERMÉS peuvent être mis à jour (API SharePoint). Un fichier "
                 "ouvert par quelqu'un est refusé par SharePoint, retenté, puis mis en attente. Pour écrire "
                 "dans les fichiers ouverts, Excel doit pouvoir ouvrir les fichiers SharePoint sur ce poste "
                 "(compte de l'organisation connecté dans Excel › Fichier › Compte, poste autorisé par "
                 "l'informatique).")
    else:
        dire(KO, "CONCLUSION : aucune voie d'écriture SharePoint ne fonctionne sur ce poste — voir les "
                 "lignes ci-dessus.")
    return lignes


def _localiser(source: SourceAudit, commune: str, client, dire) -> tuple[str, str]:
    """``(adresse SharePoint, chemin local)`` du fichier de la commune (l'un ou l'autre peut manquer)."""
    if source.mode == API:
        try:
            from .sharepoint import AccesApi
            adresse = AccesApi(source, client).trouver(source.racine, commune) or ""
        except Exception as exc:  # noqa: BLE001
            dire(KO, f"connexion SharePoint (navigateur) : {exc}")
            return "", ""
        if adresse:
            dire(INFO, f"fichier sur SharePoint : {adresse}")
        return adresse, ""
    if source.mode == URL:
        adresse = source.url_fichier(commune)
        dire(INFO, f"adresse déduite du nommage : {adresse}")
        return adresse, ""
    try:
        local = trouver_fichier_audit(source.racine, commune) or ""
    except Exception as exc:  # noqa: BLE001
        dire(KO, f"dossier inaccessible : {exc}")
        return "", ""
    adresse = (url_sharepoint(local) or "") if local else ""
    if local and source.mode == LOCAL:
        dire(INFO, "dossier synchronisé par OneDrive : adresse SharePoint " + adresse if adresse
             else "dossier local (non synchronisé par OneDrive)")
    return adresse, local


def _diagnostic_api(source: SourceAudit, adresse: str, client, dire) -> bool:
    from . import sharepoint
    from . import sharepoint_rest as rest
    try:
        client = client or sharepoint.client_pour(source.lien)
        chemin = rest.chemin_serveur(adresse)
        etat = client.etat_fichier(chemin)
    except Exception as exc:  # noqa: BLE001
        dire(KO, f"API SharePoint : {exc}")
        return False
    if etat["CheckOutType"] != 2:
        dire(KO, f"fichier EXTRAIT (check-out) par {etat['CheckedOutBy'] or 'un autre utilisateur'} : "
                 "personne ne peut le modifier avant son réarchivage (check-in)")
    qui = client.verrou_par(chemin)
    dire(KO if qui else OK, f"fichier ouvert en ce moment par {qui}" if qui else
         "SharePoint n'indique personne qui ait le fichier ouvert")
    dossier = tempfile.mkdtemp(prefix="avf_diag_")
    try:
        client.telecharger_version(chemin, os.path.join(dossier, "copie.xlsx"))
        dire(OK, "API SharePoint : session valide, téléchargement autorisé — écriture possible des fichiers "
                 "fermés (jamais par-dessus une version enregistrée entre-temps)")
        return True
    except rest.TelechargementInterdit as exc:
        dire(KO, f"API SharePoint : {exc} — ce poste n'est pas autorisé (accès conditionnel) : "
                 "ni lecture ni écriture par l'API")
        return False
    except Exception as exc:  # noqa: BLE001
        dire(KO, f"API SharePoint : téléchargement impossible ({exc})")
        return False
    finally:
        shutil.rmtree(dossier, ignore_errors=True)


def _diagnostic_excel(cible: str, session_factory, excel_present, dire) -> bool:
    if not excel_present():
        dire(KO, "Excel / pywin32 introuvables sur ce poste : pas de co-édition possible")
        return False
    fabrique = session_factory or (lambda: SessionExcel(connexion_interactive=True))
    dossier = tempfile.mkdtemp(prefix="avf_diag_")
    try:
        with fabrique() as session:
            session.copier_version_serveur(cible, os.path.join(dossier, "copie.xlsx"))
        quoi = "l'adresse SharePoint" if cible.lower().startswith("http") else "le fichier"
        dire(OK, f"Excel ouvre {quoi} avec le compte Office du poste (lecture seule, rien écrit) : "
                 "la co-édition est possible")
        return True
    except Exception as exc:  # noqa: BLE001
        texte = (str(exc) or exc.__class__.__name__).split("\n")[0][:300]
        if "démarrer Excel" in texte or exc.__class__.__name__ == "ExcelIndisponible":
            dire(KO, f"Excel n'a pas pu DÉMARRER depuis l'outil : {texte} — le compte Office n'est pas en "
                     "cause ; fermez les fenêtres Excel bloquées (Gestionnaire des tâches) et réessayez")
            return False
        dire(KO,f"Excel ne peut pas ouvrir ce fichier : {texte} — vérifiez Excel › Fichier › Compte (compte "
                 "de l'organisation) ; si le poste n'est pas autorisé par l'informatique, la co-édition "
                 "depuis ce poste est impossible")
        return False
    finally:
        shutil.rmtree(dossier, ignore_errors=True)
