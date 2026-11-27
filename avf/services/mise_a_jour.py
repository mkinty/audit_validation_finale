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
from .ecriture import SessionEcriture


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
    via_api: bool = False   # écriture par l'API SharePoint (session du navigateur)
    sans_excel: str = ""    # pourquoi la co-édition Excel n'a pas été utilisée (repli)

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
    """``{commune: [ID imm non valides]}``, éventuellement restreint (à mettre à jour avec
    ``colonne_id=rapport.colonnes.id_imm``)."""
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
    colonne_id: str | None = None,
    reprises_verrou: int = 0,
    delai_reprise_s: float = 60,
    attendre: Callable[[float], None] | None = None,
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
        reprises_verrou: nouveaux essais des fichiers refusés parce qu'ouverts par
            quelqu'un (hors fichier extrait), espacés de ``delai_reprise_s`` secondes ;
            ``doit_arreter`` interrompt l'attente.
        colonne_id: colonne qui identifie les lignes de ``groupes`` (l'**ID imm** pour un
            rapport d'audit), à la place de celle des règles.
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
    if colonne_id:
        # Identifiant des lignes à modifier : l'ID imm (unique), pas l'ID erreur (qui peut
        # couvrir plusieurs logements : toutes ses lignes seraient modifiées ensemble).
        regles = ReglesConfig.from_dict({**regles.to_dict(), "colonne_id": colonne_id})
    # Écriture cellule par cellule en co-édition Excel (le fichier peut être ouvert par
    # d'autres), avec repli expliqué si Excel est indisponible — cf. services.ecriture.
    sharepoint = source is not None and source.sharepoint
    par_api = source is not None and source.mode == API
    excel_ok = _excel_disponible(session_factory)
    if sharepoint and not simuler and not par_api and not excel_ok:
        raise EcritureSharePointImpossible(
            "la mise à jour d'un dossier SharePoint passe par Excel (co-édition), "
            "et Excel / pywin32 ne sont pas disponibles sur ce poste"
        )
    # Dossier « Parcourir » synchronisé par OneDrive : ses fichiers sont ceux de SharePoint.
    # Une réécriture locale entière y crée des copies en conflit si un collègue a le
    # fichier ouvert (les verrous ``~$`` des autres ne sont pas synchronisés) : Excel aussi.
    synchronise = not sharepoint and _synchronise(racine)
    # SharePoint hors API et OneDrive : Excel est la seule écriture sûre ; ailleurs,
    # l'option de l'utilisateur — et Excel reste appelé à la demande pour tout fichier ouvert.
    veut_excel = excel_ok and (via_excel or synchronise or (sharepoint and not par_api))
    if session_factory is SessionExcel:
        fabrique_excel = lambda: SessionExcel(connexion_interactive=True)  # noqa: E731
    else:
        fabrique_excel = session_factory
    if simuler:
        contexte = nullcontext()
    else:
        repli = None
        if par_api:
            from .sharepoint import SessionApi, client_pour
            repli = SessionApi(client_pour(source.lien))
        raison = "" if veut_excel else (
            "Excel / pywin32 introuvables sur ce poste" if not excel_ok
            else "option « Écrire via Excel (co-édition) » décochée")
        contexte = SessionEcriture(fabrique_excel if excel_ok else None, repli=repli,
                                   raison_sans_excel=raison, prefere_excel=veut_excel)

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
        def traiter(lot):
            return traiter_communes(
                lot, regles, racine, simuler=simuler,
                on_progression=on_progression, on_resultat=on_resultat, on_etape=on_etape,
                session_excel=session if not simuler else None,
                paralleles=paralleles, doit_arreter=doit_arreter,
                prereperage=not (source is not None and source.mode in (EXCEL, URL, API)),
                localiser=localiser,
                conflits=acces.conflits if acces is not None else None,
            )

        resultats = traiter(groupes)
        for tour in range(1, (0 if simuler else reprises_verrou) + 1):
            ouvertes = [r.commune for r in resultats if _ouvert_ailleurs(r)]
            if not ouvertes or (doit_arreter and doit_arreter()):
                break
            if on_etape is not None:
                on_etape("⏳", f"{len(ouvertes)} fichier(s) ouvert(s) par d'autres ({', '.join(ouvertes[:6])}"
                               f"{' …' if len(ouvertes) > 6 else ''}) : nouvel essai {tour}/{reprises_verrou} "
                               f"dans {int(delai_reprise_s)} s — « Arrêter » pour ne pas attendre")
            if not _patienter(delai_reprise_s, doit_arreter, attendre):
                break
            nouveaux = {r.commune: r for r in traiter({c: groupes[c] for c in ouvertes})}
            resultats = [nouveaux.get(r.commune, r) for r in resultats]

    echecs, reussis = attente.calculer_echecs_et_reussites(groupes, resultats)
    if colonne_id:   # message du moteur écrit pour les ID erreur : préciser la colonne réelle
        echecs = {i: (f"absent de la colonne « {colonne_id} » du fichier"
                      if r == "ID erreur introuvable dans le fichier" else r) for i, r in echecs.items()}
    ecriture = contexte if isinstance(contexte, SessionEcriture) else None
    bilan = BilanMiseAJour(
        resultats=resultats, echecs=echecs, reussis=reussis, simulation=simuler,
        via_excel=bool(ecriture and ecriture.excel_utilise),
        via_api=bool(par_api and not simuler and (ecriture is None or ecriture.repli_utilise)),
        duree=time.perf_counter() - depart, nb_communes_demandees=len(groupes),
        sans_excel=ecriture.raison_sans_excel if ecriture and ecriture.repli_utilise else "",
    )
    if fichier_attente and not simuler:
        bilan.nb_en_attente = len(attente.mettre_a_jour(fichier_attente, echecs, reussis))
    return bilan


class EcritureSharePointImpossible(RuntimeError):
    """Source SharePoint : la mise à jour exige Excel (co-édition)."""


def _ouvert_ailleurs(resultat) -> bool:
    """Commune refusée parce que son fichier est ouvert (à retenter), pas extrait ni en erreur."""
    message = resultat.message or ""
    return (resultat.statut is Statut.ERREUR and message.startswith("classeur verrouillé")
            and "EXTRAIT" not in message)


def _patienter(delai_s: float, doit_arreter, attendre=None) -> bool:
    """Attend ``delai_s`` secondes ; ``False`` si l'arrêt est demandé entre-temps."""
    if attendre is not None:
        attendre(delai_s)
        return not (doit_arreter and doit_arreter())
    fin = time.monotonic() + delai_s
    while time.monotonic() < fin:
        if doit_arreter and doit_arreter():
            return False
        time.sleep(min(0.5, max(0.0, fin - time.monotonic())))
    return True


def _synchronise(racine: str) -> bool:
    """``racine`` est dans une bibliothèque SharePoint synchronisée par OneDrive."""
    try:
        from sna.services.onedrive import url_sharepoint
        return bool(racine) and url_sharepoint(racine) is not None
    except Exception:  # noqa: BLE001
        return False


def _excel_disponible(session_factory) -> bool:
    """La vraie session exige Excel ; une fabrique de test est toujours « disponible »."""
    return excel_disponible() if session_factory is SessionExcel else True
