"""Application des règles de modification à un classeur Excel.

L'onglet lu et modifié est toujours celui nommé **« Audit »**
(``colonnes_excel.NOM_ONGLET_CIBLE``, casse et espaces ignorés) — jamais
l'onglet actif du classeur (celui affiché au dernier enregistrement dans
Excel). Un classeur sans onglet « Audit » échoue explicitement
(:class:`OngletIntrouvable`) plutôt que de modifier le mauvais onglet.

Avant d'écrire quoi que ce soit, le filtre automatique de la feuille (s'il y
en a un) est retiré, et toute ligne ou colonne masquée est démasquée. Un
filtre ou un masquage ne change rien à ce que le programme lit ou écrit —
openpyxl ignore l'affichage — mais ils peuvent cacher à l'utilisateur la ligne
qui vient d'être modifiée, au point de le faire douter que le traitement a eu
lieu. Les retirer systématiquement garantit que tout reste visible à la
réouverture du classeur.

Le traitement se fait en **deux temps**, pour ne payer le prix fort que
lorsqu'il y a réellement quelque chose à écrire :

1. **Repérage** — le XML de la feuille est parcouru en flux sur la seule
   colonne « ID erreur » (cf. ``services.lecture_rapide``, avec repli sur
   openpyxl si le classeur a une structure inattendue). Cette passe ne construit
   aucun objet cellule et détermine les numéros de ligne concernés. Les colonnes
   des règles — cible, conditions, recopie — sont résolues ici : une colonne
   absente échoue donc avant tout chargement coûteux.
2. **Écriture** — le classeur n'est chargé entièrement que si au moins une ligne
   correspond. Sur chaque ligne retenue, les conditions de la règle (si elle en
   a) sont évaluées ; seules les lignes qui les vérifient toutes sont touchées.

Un fichier dont aucun ID erreur n'est présent n'est ainsi jamais chargé en
mémoire ; c'est le gain le plus important sur les gros classeurs.

Attention : openpyxl réécrit intégralement le classeur à la sauvegarde.
Certains éléments (graphiques, images, mises en forme conditionnelles avancées)
peuvent ne pas être préservés. Testez toujours sur une copie avant un
traitement en masse. Le fichier n'est réécrit que si au moins une cellule a
réellement changé.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import date, datetime
import warnings

from openpyxl import load_workbook

from ..models.regles import RegleModification, ReglesConfig, SourceValeur, TypeDroite
from .colonnes_excel import (
    OngletIntrouvable, choisir_feuille, entete_est_date, entetes_depuis_feuille,
    resoudre_colonne, table_depuis_entetes,
)
from .evaluation import regle_sapplique, valeur_pour_action
from .identifiants import normaliser_id
from .lecture_rapide import LectureImpossible, lire_colonne, lire_entetes
from .valeurs import assurer_format_date, harmoniser_forme, memes_valeurs

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Signature du rapporteur d'étapes (journal en temps réel).
Etape = Callable[[str], None]

# Périodicité des points d'avancement pendant le repérage des lignes.
_PAS_AVANCEMENT = 20_000

# Nombre de lignes explorées au maximum pour trouver une cellule de référence
# de mise en forme.
_PROFONDEUR_MODELE = 500

# Étendue de l'échantillon servant à déterminer la police dominante du classeur.
_LIGNES_POLICE = 200
_COLONNES_POLICE = 40


class ColonneIntrouvable(LookupError):
    """Une règle — ou la colonne des ID erreur — est absente du classeur."""


@dataclass
class _Plan:
    """Colonnes d'une règle résolues une fois pour toutes sur un classeur.

    ``refs`` traduit chaque référence textuelle citée par la règle (cible,
    colonnes des conditions, colonne recopiée) en numéro de colonne : le
    lecteur de ligne s'en sert sans refaire de recherche.
    """

    cible: int
    refs: dict[str, int] = field(default_factory=dict)


@dataclass
class BilanFichier:
    """Ce qu'a produit le traitement d'un fichier.

    ``ids_trouves`` permet à l'appelant de signaler les identifiants demandés
    qui n'apparaissent nulle part dans le classeur. ``visibilite_restauree``
    signale qu'un filtre a été retiré, ou une ligne/colonne démasquée.

    Les quatre derniers champs ne sont renseignés que par
    ``services.excel_com`` (écriture via Excel) : ``ouverture`` (co-édition
    SharePoint ou fichier local), ``verrouille`` (qui bloque le classeur, le
    cas échéant — rien n'est alors écrit), ``ids_non_verifies`` (écrits
    localement mais dont la relecture après enregistrement ne confirme pas la
    valeur) et ``tentatives`` (essais nécessaires face à un Excel occupé).
    """

    comptes: dict[str, int] = field(default_factory=dict)
    nb_lignes: int = 0
    ids_trouves: set[str] = field(default_factory=set)
    nb_lignes_lues: int = 0
    enregistre: bool = False
    visibilite_restauree: bool = False
    ouverture: str | None = None
    verrouille: str | None = None
    ids_non_verifies: list[str] = field(default_factory=list)
    tentatives: int = 0

    @property
    def nb_modifications(self) -> int:
        return sum(self.comptes.values())


def appliquer_regles(
    chemin_fichier: str,
    config: ReglesConfig,
    ids_erreur: Iterable[str],
    simuler: bool = False,
    on_etape: Etape | None = None,
) -> BilanFichier:
    """Applique les règles actives aux lignes des ID erreur demandés.

    Args:
        chemin_fichier: classeur à modifier (modifié sur place).
        config: règles à appliquer, ligne d'en-tête et colonne des ID erreur.
        ids_erreur: identifiants ``<insee>_<numéro>`` à traiter dans ce fichier.
        simuler: si vrai, rien n'est écrit sur le disque — le décompte retourné
            indique ce qui *serait* modifié.
        on_etape: appelé à chaque étape avec un message court, pour alimenter
            le journal pendant le traitement plutôt qu'à la fin.

    Returns:
        Un :class:`BilanFichier` : modifications par règle, nombre de lignes
        retenues, et identifiants effectivement rencontrés.

    Raises:
        ColonneIntrouvable: si la colonne des ID erreur ou une colonne citée
            par une règle est absente du classeur.
    """
    regles = config.actives
    attendus = {normaliser_id(i) for i in ids_erreur} - {""}
    bilan = BilanFichier(comptes={r.code: 0 for r in regles})
    if not regles or not attendus:
        return bilan

    _dire(on_etape, "analyse du fichier…")
    lignes, entetes, plans = _reperer(chemin_fichier, config, regles, attendus, bilan, on_etape)

    if not lignes:
        _dire(on_etape, f"aucune ligne concernée sur {bilan.nb_lignes_lues} lue(s) — non ouvert en écriture")
        return bilan

    _dire(on_etape, f"{len(lignes)} ligne(s) à modifier — chargement du classeur…")
    classeur = load_workbook(chemin_fichier)
    try:
        feuille = choisir_feuille(classeur)

        bilan.visibilite_restauree = _afficher_tout(feuille)
        if bilan.visibilite_restauree:
            _dire(on_etape, "🔓 filtre retiré, lignes/colonnes masquées démasquées")

        secours = _PoliceSecours(feuille, config.ligne_entete)
        contextes = {
            regle.code: _contexte_colonne(
                feuille, plans[regle.code], entetes, config, set(lignes), regle, secours,
            )
            for regle in regles
        }

        _dire(on_etape, "application des règles…")
        for ligne in lignes:
            for regle in regles:
                contexte = contextes[regle.code]
                lire = _lecteur(feuille, ligne, contexte.refs)
                if not regle_sapplique(regle, lire):
                    continue

                valeur = valeur_pour_action(regle.action, lire, en_date=contexte.en_date)
                cellule = feuille.cell(ligne, contexte.index)
                if memes_valeurs(cellule.value, valeur):
                    continue

                format_initial = cellule.number_format
                cellule.value = valeur
                if config.harmoniser_style:
                    harmoniser_forme(cellule, contexte.modele, contexte.police)
                if isinstance(valeur, (datetime, date)):
                    assurer_format_date(
                        cellule, contexte.format_impose(config, format_initial)
                    )
                bilan.comptes[regle.code] += 1

        a_ecrire = bilan.nb_modifications or bilan.visibilite_restauree
        if a_ecrire and not simuler:
            _dire(on_etape, f"enregistrement de {bilan.nb_modifications} modification(s)…")
            classeur.save(chemin_fichier)
            bilan.enregistre = True
        elif simuler:
            _dire(on_etape, "simulation — aucune écriture")
        else:
            _dire(on_etape, "valeurs déjà en place — fichier inchangé")
        return bilan
    finally:
        classeur.close()


def reperer_ids(
    chemin_fichier: str, config: ReglesConfig, ids_erreur: Iterable[str],
    on_etape: Etape | None = None,
) -> BilanFichier:
    """Repérage seul (passe 1), sans jamais ouvrir le classeur en écriture.

    Sert de pré-vérification bon marché avant d'ouvrir une session Excel
    (``services.excel_com``) : si aucun des ID demandés n'apparaît dans le
    fichier, inutile de payer le coût d'une session COM pour ne rien y faire.
    """
    regles = config.actives
    attendus = {normaliser_id(i) for i in ids_erreur} - {""}
    bilan = BilanFichier(comptes={r.code: 0 for r in regles})
    if not regles or not attendus:
        return bilan
    _reperer(chemin_fichier, config, regles, attendus, bilan, on_etape)
    return bilan


# ── Passe 1 : repérage ────────────────────────────────────────────
def _reperer(
    chemin: str,
    config: ReglesConfig,
    regles: list,
    attendus: set[str],
    bilan: BilanFichier,
    on_etape: Etape | None,
) -> tuple[list[int], dict[int, str], dict[str, _Plan]]:
    """Numéros de lignes concernées, en-têtes, et plans de règles résolus.

    La lecture directe du XML est tentée d'abord ; tout classeur de structure
    inattendue bascule sur openpyxl, plus lent mais universel. Le résultat est
    identique dans les deux cas.
    """
    try:
        return _reperer_rapide(chemin, config, regles, attendus, bilan, on_etape)
    except LectureImpossible as exc:
        _dire(on_etape, f"lecture rapide indisponible ({exc}) — repli openpyxl")
        return _reperer_openpyxl(chemin, config, regles, attendus, bilan, on_etape)


def _reperer_rapide(
    chemin: str, config: ReglesConfig, regles: list, attendus: set[str],
    bilan: BilanFichier, on_etape: Etape | None,
) -> tuple[list[int], dict[int, str], dict[str, _Plan]]:
    """Repérage par lecture directe du XML : une seule colonne est décodée."""
    entetes = lire_entetes(chemin, config.ligne_entete)
    if not entetes:
        raise LectureImpossible("ligne d'en-tête vide")

    colonne_id, plans = _resoudre_colonnes(entetes, config, regles)

    lignes: list[int] = []
    trouves: set[str] = set()
    derniere = config.ligne_entete
    for ligne, valeur in lire_colonne(chemin, colonne_id, config.ligne_entete + 1):
        derniere = ligne
        identifiant = normaliser_id(valeur)
        if identifiant in attendus:
            lignes.append(ligne)
            trouves.add(identifiant)
        if ligne % _PAS_AVANCEMENT == 0:
            _dire(on_etape, _avancement(ligne))

    return _cloturer(bilan, lignes, trouves, derniere, config), entetes, plans


def _reperer_openpyxl(
    chemin: str, config: ReglesConfig, regles: list, attendus: set[str],
    bilan: BilanFichier, on_etape: Etape | None,
) -> tuple[list[int], dict[int, str], dict[str, _Plan]]:
    """Repérage de repli, en lecture seule openpyxl."""
    classeur = load_workbook(chemin, read_only=True)
    try:
        feuille = choisir_feuille(classeur)
        entetes = entetes_depuis_feuille(feuille, config.ligne_entete)
        colonne_id, plans = _resoudre_colonnes(entetes, config, regles)

        lignes: list[int] = []
        trouves: set[str] = set()
        ligne = config.ligne_entete
        for cellules in feuille.iter_rows(
            min_row=config.ligne_entete + 1,
            min_col=colonne_id,
            max_col=colonne_id,
            values_only=True,
        ):
            ligne += 1
            identifiant = normaliser_id(cellules[0] if cellules else None)
            if identifiant in attendus:
                lignes.append(ligne)
                trouves.add(identifiant)
            if ligne % _PAS_AVANCEMENT == 0:
                _dire(on_etape, _avancement(ligne))

        return _cloturer(bilan, lignes, trouves, ligne, config), entetes, plans
    finally:
        classeur.close()


def _resoudre_colonnes(
    entetes: dict[int, str], config: ReglesConfig, regles: list
) -> tuple[int, dict[str, _Plan]]:
    """Colonne des ID erreur et plan de chaque règle, résolus avant tout chargement.

    Faire cette résolution ici garantit qu'une colonne absente est signalée sans
    avoir chargé le classeur en entier.
    """
    table = table_depuis_entetes(entetes)
    colonne_id = _resoudre(config.colonne_id, table, "colonne des ID erreur")
    plans = {regle.code: _plan_regle(regle, table) for regle in regles}
    return colonne_id, plans


def _plan_regle(regle: RegleModification, table: dict[str, int]) -> _Plan:
    """Résout toutes les colonnes citées par une règle (cible, conditions, recopie)."""
    refs: dict[str, int] = {}

    def resoudre(reference: str, contexte: str) -> int:
        index = _resoudre(reference, table, contexte)
        refs[reference] = index
        return index

    for condition in regle.conditions:
        if condition.colonne.strip():
            resoudre(condition.colonne, f"règle {regle.code}")
        if condition.type_droite is TypeDroite.COLONNE and condition.droite.strip():
            resoudre(condition.droite, f"règle {regle.code}")

    if regle.action.source is SourceValeur.COLONNE and str(regle.action.valeur).strip():
        resoudre(regle.action.valeur, f"règle {regle.code} (recopie)")

    cible = resoudre(regle.colonne_cible, f"règle {regle.code}")
    return _Plan(cible=cible, refs=refs)


def _lecteur(feuille, ligne: int, refs: dict[str, int]):
    """Accesseur ``référence -> valeur`` pour une ligne donnée, sans créer de cellule."""
    def lire(reference: str):
        index = refs.get(reference)
        if index is None:
            return None
        cellule = _cellule_existante(feuille, ligne, index)
        return cellule.value if cellule is not None else None

    return lire


def _cloturer(
    bilan: BilanFichier, lignes: list[int], trouves: set[str],
    derniere: int, config: ReglesConfig,
) -> list[int]:
    """Reporte le résultat du repérage dans le bilan (une fois la passe finie)."""
    bilan.ids_trouves = trouves
    bilan.nb_lignes = len(lignes)
    bilan.nb_lignes_lues = max(derniere - config.ligne_entete, 0)
    return lignes


def _avancement(ligne: int) -> str:
    return f"{ligne:,} lignes analysées…".replace(",", " ")


# ── Filtre et masquage ────────────────────────────────────────────
def _afficher_tout(feuille) -> bool:
    """Retire le filtre automatique et démasque les lignes et colonnes.

    Un filtre actif ou des colonnes masquées n'empêchent ni de lire ni
    d'écrire une cellule — openpyxl ignore totalement l'affichage. Mais une
    ligne filtrée (donc masquée) qui vient d'être modifiée resterait invisible
    à la réouverture du classeur, ce qui peut faire croire à l'utilisateur que
    rien n'a été fait. On les retire donc avant d'écrire la moindre valeur.

    Returns:
        Vrai si quelque chose a été changé (filtre retiré, ligne ou colonne
        démasquée) — sert à décider si le classeur doit être enregistré même
        quand aucune cellule n'a changé de valeur.
    """
    change = False

    if feuille.auto_filter.ref:
        feuille.auto_filter.ref = None
        change = True

    for dimension in feuille.row_dimensions.values():
        if dimension.hidden:
            dimension.hidden = False
            change = True

    for dimension in feuille.column_dimensions.values():
        if dimension.hidden:
            dimension.hidden = False
            change = True

    return change


# ── Passe 2 : contexte d'écriture d'une colonne ───────────────────
@dataclass
class _Contexte:
    """Tout ce qu'il faut savoir pour écrire dans une colonne, calculé une fois."""

    index: int
    refs: dict[str, int]
    en_date: bool
    modele: object = None
    police: object = None

    def format_impose(self, config: ReglesConfig, format_initial: str) -> str:
        """Format d'affichage réellement voulu pour la cellule écrite.

        C'est celui du modèle quand la mise en forme est reprise, sinon celui
        que la cellule portait déjà.
        """
        if config.harmoniser_style and self.modele is not None:
            return self.modele.number_format
        return format_initial


def _contexte_colonne(
    feuille, plan: _Plan, entetes: dict[int, str], config: ReglesConfig,
    lignes_modifiees: set[int], regle: RegleModification, secours: "_PoliceSecours",
) -> _Contexte:
    """Résout la mise en forme à reprendre et le mode de saisie de la colonne cible.

    Le modèle est cherché parmi les lignes **non modifiées** : on recopie ainsi
    la forme d'origine du fichier, jamais celle d'une cellule qu'on vient
    d'écrire. Il sert aussi à confirmer qu'une colonne contient des dates
    lorsque son en-tête ne le dit pas.

    Colonne vide — le cas des colonnes d'audit à remplir — : la police dominante
    du classeur prend le relais.
    """
    modele = _cellule_modele(feuille, plan.cible, config.ligne_entete, lignes_modifiees)
    en_date = entete_est_date(entetes.get(plan.cible, "")) or (
        modele is not None and isinstance(modele.value, (datetime, date))
    )
    return _Contexte(
        index=plan.cible,
        refs=plan.refs,
        en_date=en_date,
        modele=modele,
        police=None if modele is not None else secours.police(),
    )


def _cellule_modele(feuille, index: int, ligne_entete: int, exclues: set[int]):
    """Première cellule renseignée de la colonne, hors lignes modifiées.

    L'exploration est bornée : au-delà de quelques centaines de lignes sans
    aucune valeur, la colonne est considérée comme vide.
    """
    derniere = min(feuille.max_row or ligne_entete, ligne_entete + _PROFONDEUR_MODELE)
    for ligne in range(ligne_entete + 1, derniere + 1):
        if ligne in exclues:
            continue
        cellule = _cellule_existante(feuille, ligne, index)
        if cellule is not None and cellule.value not in (None, ""):
            return cellule
    return None


class _PoliceSecours:
    """Police dominante du classeur, calculée une seule fois et à la demande.

    Beaucoup d'audits sont composés en Calibri 10 alors qu'openpyxl écrit en
    Calibri 11 par défaut : une valeur inscrite dans une colonne vide se
    repérerait immédiatement à l'œil. On relève donc la police la plus employée
    dans les données du classeur, et c'est elle qu'on applique.
    """

    def __init__(self, feuille, ligne_entete: int):
        self._feuille = feuille
        self._ligne_entete = ligne_entete
        self._calculee = False
        self._police = None

    def police(self):
        if not self._calculee:
            self._police = self._relever()
            self._calculee = True
        return self._police

    def _relever(self):
        feuille, entete = self._feuille, self._ligne_entete
        derniere = min(feuille.max_row or entete, entete + _LIGNES_POLICE)
        colonnes = min(feuille.max_column or 1, _COLONNES_POLICE)

        frequences: Counter = Counter()
        exemples: dict = {}
        for ligne in range(entete + 1, derniere + 1):
            for colonne in range(1, colonnes + 1):
                cellule = _cellule_existante(feuille, ligne, colonne)
                if cellule is None or cellule.value in (None, ""):
                    continue
                signature = (cellule.font.name, cellule.font.size)
                frequences[signature] += 1
                exemples.setdefault(signature, cellule.font)

        if not frequences:
            return None
        return exemples[frequences.most_common(1)[0][0]]


def _cellule_existante(feuille, ligne: int, colonne: int):
    """Cellule déjà présente dans la feuille, sans en créer une nouvelle.

    ``feuille.cell(...)`` crée la cellule — et donc sa ligne — dès qu'on la
    consulte : inspecter la mise en forme ajouterait ainsi des lignes vides au
    classeur enregistré. On interroge donc directement la table interne, en se
    rabattant sur l'accès normal si openpyxl venait à changer.
    """
    cellules = getattr(feuille, "_cells", None)
    if cellules is None:  # pragma: no cover — filet de sécurité
        return feuille.cell(ligne, colonne)
    return cellules.get((ligne, colonne))


# ── Helpers ───────────────────────────────────────────────────────
def _resoudre(reference: str, table: dict[str, int], contexte: str) -> int:
    """Traduit une référence de colonne en numéro, ou échoue explicitement."""
    index = resoudre_colonne(reference, table)
    if index is None:
        raise ColonneIntrouvable(
            f"{contexte} : colonne « {reference} » introuvable dans le fichier"
        )
    return index


def _dire(on_etape: Etape | None, message: str) -> None:
    if on_etape is not None:
        on_etape(message)
