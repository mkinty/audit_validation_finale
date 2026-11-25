"""Mise à jour des ID erreur non valides — via le moteur ``sna``.

Le processus est **celui de l'outil complete_excel_files**, appelé tel quel :

    sna.services.traitement.traiter_communes(groupes, règles, racine, …)

donc, pour chaque commune : recherche du fichier, refus si une copie de
conflit est présente, repérage rapide des lignes, application des règles
(conditions + action), retrait du filtre / démasquage, reprise de la mise en
forme, enregistrement — par openpyxl ou via une instance Excel réelle
(co-édition SharePoint, relecture après enregistrement). Les ID non appliqués
rejoignent la liste « ID en attente » (``sna.services.attente``).

Ce module ne fait que trois choses en amont :

1. construire ``{commune: [ID erreur]}`` à partir du rapport d'audit — ce sont
   les ID non valides (au moins une ligne ne respecte pas les critères d'audit) ;
2. **dater du jour** les règles qui écrivent la colonne « Date audit »
   (option, activée par défaut) : la règle par défaut de complete_excel_files
   fige la date au jour où la configuration a été créée ; ici elle suit le jour
   de la mise à jour ;
3. ouvrir la session Excel si l'écriture via Excel est demandée.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from contextlib import ExitStack, nullcontext
from dataclasses import dataclass, field, replace
from datetime import date
import time

from sna.models.base import ResultatCommune, Statut
from sna.models.regles import ReglesConfig, SourceValeur
from sna.services import attente
from sna.services.colonnes_excel import normaliser
from sna.services.excel_com import SessionExcel, excel_disponible
from sna.services.traitement import traiter_communes

from ..models.audit import RapportAudit
from ..models.source import API, EXCEL, URL, SourceAudit


@dataclass
class BilanMiseAJour:
    """Résultat global d'une mise à jour."""

    resultats: list[ResultatCommune] = field(default_factory=list)
    echecs: dict[str, str] = field(default_factory=dict)
    reussis: list[str] = field(default_factory=list)
    simulation: bool = False
    via_excel: bool = False
    duree: float = 0.0
    nb_en_attente: int = 0
    nb_communes_demandees: int = 0

    @property
    def interrompu(self) -> bool:
        """Arrêt demandé : des communes n'ont pas été commencées (rien n'y a été écrit)."""
        return self.nb_communes < self.nb_communes_demandees

    @property
    def nb_communes(self) -> int:
        return len(self.resultats)

    @property
    def nb_communes_ok(self) -> int:
        return sum(1 for r in self.resultats if r.statut is Statut.SUCCES)

    @property
    def nb_modifications(self) -> int:
        return sum(r.nb_modifications for r in self.resultats)

    @property
    def communes_traitees(self) -> list[str]:
        return [r.commune for r in self.resultats]


# ── Préparation ───────────────────────────────────────────────────
def groupes_depuis_rapport(
    rapport: RapportAudit, communes: Sequence[str] | None = None,
) -> dict[str, list[str]]:
    """``{commune: [ID non valides]}``, éventuellement restreint."""
    groupes = rapport.ids_non_valides_par_commune()
    if communes is None:
        return groupes
    voulues = set(communes)
    return {c: ids for c, ids in groupes.items() if c in voulues}


def preparer_regles(
    regles: ReglesConfig,
    colonne_date: str = "Date audit",
    date_du_jour: bool = True,
    aujourd_hui: date | None = None,
) -> ReglesConfig:
    """Copie des règles prête à appliquer (l'original n'est jamais modifié).

    Avec ``date_du_jour``, chaque règle écrivant une **valeur** dans
    ``colonne_date`` reçoit la date du jour (``JJ/MM/AAAA``). Les autres règles
    — recopie de colonne, effacement, autres colonnes — sont inchangées.
    """
    copie = ReglesConfig.from_dict(regles.to_dict())
    if not date_du_jour or not colonne_date.strip():
        return copie
    jour = (aujourd_hui or date.today()).strftime("%d/%m/%Y")
    cible = normaliser(colonne_date)
    for regle in copie.regles:
        if regle.action.source is SourceValeur.VALEUR and normaliser(regle.colonne_cible) == cible:
            regle.action = replace(regle.action, valeur=jour)
    return copie


# ── Exécution ─────────────────────────────────────────────────────
def mettre_a_jour(
    groupes: Mapping[str, Sequence[str]],
    regles: ReglesConfig,
    racine: str,
    simuler: bool = False,
    via_excel: bool = False,
    fichier_attente: str | None = None,
    on_progression: Callable[[int, int, str], None] | None = None,
    on_etape: Callable[[str, str], None] | None = None,
    on_resultat: Callable[[ResultatCommune], None] | None = None,
    session_factory: Callable[[], object] = SessionExcel,
    paralleles: int = 1,
    doit_arreter: Callable[[], bool] | None = None,
    source: SourceAudit | None = None,
) -> BilanMiseAJour:
    """Applique les règles aux ID erreur des communes, comme complete_excel_files.

    Args:
        groupes: ``{commune: [ID erreur]}`` — cf. :func:`groupes_depuis_rapport`.
        regles: règles **déjà préparées** (cf. :func:`preparer_regles`).
        simuler: aucune écriture sur le disque, décompte seulement.
        via_excel: écrire via une instance Excel (ignoré en simulation ou si
            Excel / pywin32 sont indisponibles).
        fichier_attente: liste « ID en attente » à tenir à jour (``None`` :
            non tenue, utile en test).
        session_factory: fabrique de la session Excel (remplaçable en test).
        paralleles: communes traitées en même temps (threads). Via Excel, la
            préparation des communes suivantes se fait pendant que la session
            unique écrit la commune en cours.
        doit_arreter: consulté avant chaque commune (bouton « Arrêter ») ; une
            commune commencée va toujours à son terme.
        source: source de l'audit. En **SharePoint**, l'écriture passe
            toujours par Excel (co-édition), jamais par une réécriture directe
            du fichier sur le serveur ; et via une bibliothèque synchronisée,
            Excel est ouvert même si la copie locale ne contient pas l'ID (elle
            peut différer du serveur).

    Raises:
        sna.services.excel_com.ExcelIndisponible: Excel n'a pas pu démarrer.
        EcritureSharePointImpossible: source SharePoint sans Excel disponible.
    """
    depart = time.perf_counter()
    sharepoint = source is not None and source.sharepoint
    if sharepoint and not simuler:
        if not _excel_disponible(session_factory):
            raise EcritureSharePointImpossible(
                "la mise à jour d'un dossier SharePoint passe par Excel (co-édition), "
                "et Excel / pywin32 ne sont pas disponibles sur ce poste"
            )
        via_excel = True
    ecrire_via_excel = via_excel and not simuler and _excel_disponible(session_factory)
    if session_factory is SessionExcel:
        # Si Office demande la connexion Microsoft 365, la montrer plutôt qu'échouer.
        contexte = SessionExcel(connexion_interactive=True) if ecrire_via_excel else nullcontext()
    else:
        contexte = session_factory() if ecrire_via_excel else nullcontext()

    # Dossier non parcourable localement : le fichier de chaque commune est désigné
    # par son adresse SharePoint (trouvé par l'API, ou déduit du nommage).
    # En simulation, rien n'est écrit : le décompte se fait sur une copie téléchargée
    # de la version serveur (openpyxl ne lit pas une adresse https).
    acces = None
    localiser = None
    pile = ExitStack()
    if source is not None and source.mode in (URL, API):
        from .sharepoint import acces_pour
        acces = acces_pour(source)
        localiser = acces.trouver
        if simuler:
            pile.enter_context(acces)

            def localiser(racine_, commune):
                adresse = acces.trouver(racine_, commune)
                return acces.lisible(adresse) if adresse else None

    with pile, contexte as session:
        resultats = traiter_communes(
            groupes, regles, racine, simuler=simuler,
            on_progression=on_progression, on_resultat=on_resultat, on_etape=on_etape,
            session_excel=session if ecrire_via_excel else None,
            paralleles=paralleles, doit_arreter=doit_arreter,
            prereperage=not (source is not None and source.mode in (EXCEL, URL, API)),
            localiser=localiser,
            conflits=acces.conflits if acces is not None else None,
        )

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)
    bilan = BilanMiseAJour(
        resultats=resultats, echecs=echecs, reussis=reussis, simulation=simuler,
        via_excel=ecrire_via_excel, duree=time.perf_counter() - depart,
        nb_communes_demandees=len(groupes),
    )
    if fichier_attente and not simuler:
        bilan.nb_en_attente = len(attente.mettre_a_jour(fichier_attente, echecs, reussis))
    return bilan


class EcritureSharePointImpossible(RuntimeError):
    """Source SharePoint : la mise à jour exige Excel (co-édition)."""


def _excel_disponible(session_factory) -> bool:
    """La vraie session exige Excel ; une fabrique de test est toujours « disponible »."""
    return excel_disponible() if session_factory is SessionExcel else True
