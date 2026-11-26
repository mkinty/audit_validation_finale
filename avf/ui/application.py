"""Fenêtre principale de l'outil (Tkinter).

Déroulé, de haut en bas :

    1. dossier racine : dossier local (parcourir), ou lien SharePoint du dossier
       racine — l'audit lit alors la version serveur des fichiers ;
    2. communes — toutes celles de ``<racine>/Dep<XX>/<commune>`` détectées,
       codes INSEE renseignés (le dossier ``Dep<XX>`` en est déduit), ou
       seulement des ID erreur (``13001_2``…) : l'audit ne contrôle que leurs lignes ;
    3. audit des 4 colonnes de chaque ligne (threads de fond, journal en temps réel),
       rapport en popup, export HTML / Excel / CSV ;
    4. mise à jour des ID non valides — moteur et règles de
       complete_excel_files, réglables dans le même constructeur de règles.

Toute la logique métier est dans ``avf.services`` (et ``sna.services`` pour la
mise à jour). Les threads de fond ne touchent jamais aux widgets : ils
déposent leurs messages et rappels dans une file vidée périodiquement sur le
thread principal.
"""

from __future__ import annotations

from collections.abc import Callable
import os
import queue
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from sna.models.base import ResultatCommune, Statut
from sna.models.regles import ReglesConfig
from sna.services.excel_com import ExcelIndisponible, excel_disponible
from sna.services.fichiers import trouver_fichier_exemple
from sna.ui.panneau_regles import PanneauRegles
from sna.ui.theme import STYLE_STATUT, Couleurs, Polices

from .. import chemins
from ..config import Parametres, charger, enregistrer
from ..models.audit import PAR_ID_ERREUR, PAR_ID_IMM, AuditCommune, RapportAudit
from ..models.source import API, EXCEL, LOCAL, URL, SourceAudit
from ..services import detection, sharepoint, sharepoint_rest
from ..services.import_liste import COLONNES_ID, COLONNES_ID_IMM, COLONNES_INSEE
from sna.services.identifiants import grouper_par_commune
from ..services.audit import auditer_communes
from ..services.export import FORMATS, exporter, nom_fichier_defaut, ouvrir_dans_navigateur, ouvrir_fichier
from ..services.mise_a_jour import (
    EcritureSharePointImpossible, groupes_depuis_rapport, mettre_a_jour, preparer_regles,
)
from ..views import presentation as p
from .composants import (
    DialogueColonnes, DialogueSaisie, ZoneDefilante, bouton, carte, case, champ, configurer_styles,
)
from .fenetre_rapport import FenetreRapport

SEPARATEUR = "─" * 64
PERIODE_FILE = 100  # ms


class Application(tk.Tk):
    def __init__(self, fichier_config: str = chemins.CONFIG_FILE,
                 fichier_attente: str = chemins.IDS_EN_ATTENTE_FILE):
        super().__init__()
        self.fichier_config = fichier_config
        self.fichier_attente = fichier_attente
        self.params: Parametres = charger(fichier_config)
        self.rapport: RapportAudit | None = None
        self.fenetre_rapport: FenetreRapport | None = None
        self.occupe = False
        self._arret = threading.Event()
        self._file: queue.Queue[Callable[[], None]] = queue.Queue()
        self._excel = excel_disponible()

        self.title("Audit de la Validation finale — Projet SNA")
        hauteur = min(940, max(self.winfo_screenheight() - 90, 480))
        self.geometry(f"1100x{hauteur}")
        self.minsize(760, 420)
        self.configure(bg=Couleurs.FOND)
        configurer_styles(self)

        self.var_racine = tk.StringVar(value=self.params.racine)
        self.var_source = tk.StringVar(value=self.params.source)
        self.var_lien = tk.StringVar(value=self.params.lien_sharepoint)
        self._source_sp: SourceAudit | None = None   # accès SharePoint établi
        self.ids_cibles: dict[str, list[str]] = {}    # audit restreint à ces ID ({} : communes entières)
        self.type_cibles = self.params.type_ids       # PAR_ID_ERREUR ou PAR_ID_IMM
        self.var_simulation = tk.BooleanVar(value=False)
        self.var_date_du_jour = tk.BooleanVar(value=self.params.date_du_jour)
        self.var_excel = tk.BooleanVar(value=self.params.ecrire_via_excel and self._excel)
        self.var_reauditer = tk.BooleanVar(value=self.params.reauditer)

        self._construire()
        self._restaurer_selection()
        self._maj_resume_regles()
        self._maj_etat()
        self.after(PERIODE_FILE, self._vider_file)
        if self.var_source.get() == "sharepoint" and self.var_lien.get().strip():
            self.after(300, lambda: self._connecter_sharepoint(silencieux=True))

    # ══ Construction ══════════════════════════════════════════════
    def _construire(self) -> None:
        entete = tk.Frame(self, bg=Couleurs.FOND_HEADER, pady=10)
        entete.pack(fill="x")
        tk.Label(entete, text="  🔎  Audit de la Validation finale", font=Polices.TITRE,
                 bg=Couleurs.FOND_HEADER, fg=Couleurs.TEXTE).pack(side="left")
        tk.Label(entete, text="<racine>/Dep<XX>/<commune>/audit_*.xlsx  ", font=Polices.MONO,
                 bg=Couleurs.FOND_HEADER, fg=Couleurs.TEXTE_SECONDAIRE).pack(side="right")

        # Tout le contenu défile : sur un petit écran, aucune étape n'est hors d'atteinte.
        self.zone = ZoneDefilante(self, bg=Couleurs.FOND)
        self.zone.pack(fill="both", expand=True)
        corps = tk.Frame(self.zone.interieur, bg=Couleurs.FOND)
        corps.pack(fill="both", expand=True, padx=16, pady=10)

        self._construire_racine(corps)
        self._construire_communes(corps)
        self._construire_audit(corps)
        self._construire_mise_a_jour(corps)
        self._construire_journal(corps)

    def _construire_racine(self, parent) -> None:
        c = carte(parent, "1 · Dossier racine")
        c.pack(fill="x")
        choix = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        choix.pack(fill="x", pady=(0, 6))
        self.choix_source = {}
        for valeur, texte in (("local", "📁 Dossier local (parcourir)"),
                              ("sharepoint", "🌐 Lien SharePoint (version serveur)")):
            bouton_radio = tk.Radiobutton(
                choix, text=texte, value=valeur, variable=self.var_source, command=self._changer_source,
                bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE, selectcolor=Couleurs.FOND_PANNEAU2,
                activebackground=Couleurs.FOND_CARTE, activeforeground=Couleurs.TEXTE,
                font=Polices.LABEL, cursor="hand2",
            )
            bouton_radio.pack(side="left", padx=(0, 22))
            self.choix_source[valeur] = bouton_radio

        self.ligne_local = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        champ(self.ligne_local, self.var_racine).pack(side="left", fill="x", expand=True)
        bouton(self.ligne_local, "📁 Parcourir", self._parcourir_racine).pack(side="left", padx=(6, 0))
        bouton(self.ligne_local, "Enregistrer", self._enregistrer_racine, "succes").pack(side="left", padx=(6, 0))

        self.ligne_sp = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        self.champ_lien = champ(self.ligne_sp, self.var_lien)
        self.champ_lien.pack(side="left", fill="x", expand=True)
        self.champ_lien.bind("<Return>", lambda e: self._connecter_sharepoint())
        self.btn_connecter = bouton(self.ligne_sp, "🔗 Se connecter", self._connecter_sharepoint, "succes")
        self.btn_connecter.pack(side="left", padx=(6, 0))
        self.btn_oublier = bouton(self.ligne_sp, "🚪 Oublier", self._oublier_connexion, "discret")
        self.btn_oublier.pack(side="left", padx=(6, 0))

        self.lbl_source = tk.Label(c.interieur, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
                                   font=Polices.PETIT, anchor="w", justify="left")
        self.lbl_source.pack(fill="x", pady=(4, 0), side="bottom")
        self._changer_source(initial=True)

    def _construire_communes(self, parent) -> None:
        c = carte(parent, "2 · Communes", "codes INSEE — <racine>/Dep<XX>/<commune>")
        c.pack(fill="both", pady=(8, 0))
        boutons = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        boutons.pack(fill="x")
        self.btn_det_com = bouton(boutons, "🔍 Détecter toutes les communes",
                                  self._detecter_communes, "principal")
        self.btn_det_com.pack(side="left")
        self.btn_saisie_com = bouton(boutons, "✏️ Renseigner…", self._renseigner_communes)
        self.btn_saisie_com.pack(side="left", padx=6)
        self.btn_saisie_ids = bouton(boutons, "🎯 ID erreur…", self._renseigner_ids)
        self.btn_saisie_ids.pack(side="left")
        self.btn_saisie_imm = bouton(boutons, "🆔 ID imm…", lambda: self._renseigner_ids(PAR_ID_IMM))
        self.btn_saisie_imm.pack(side="left", padx=6)
        self.cadre_perimetre = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        self.lbl_perimetre = tk.Label(self.cadre_perimetre, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.INFO,
                                      font=Polices.PETIT, anchor="w", justify="left", wraplength=820)
        self.lbl_perimetre.pack(side="left", fill="x", expand=True)
        self.btn_effacer_ids = bouton(self.cadre_perimetre, "✖ Communes entières", self._effacer_ids, "discret")
        self.btn_effacer_ids.pack(side="right")
        self.liste_com = self._liste(c.interieur, colonnes=4)
        self.liste_com.bind("<<ListboxSelect>>", lambda e: (self._sauver_selection(), self._maj_etat()))
        self.lbl_com = self._ligne_selection(c.interieur, self.liste_com)
        ligne = self.lbl_com.master
        self.btn_vider_ids = bouton(ligne, "🗑 ID erreur", self._vider_ids, "discret", padx=8, pady=2)
        self.btn_vider_ids.pack(side="right", padx=(0, 12))
        self.btn_vider_com = bouton(ligne, "🗑 Communes", self._vider_communes, "discret", padx=8, pady=2)
        self.btn_vider_com.pack(side="right", padx=4)

    def _liste(self, parent, colonnes: int = 1) -> tk.Listbox:
        cadre = tk.Frame(parent, bg=Couleurs.FOND_CARTE)
        cadre.pack(fill="both", expand=True, pady=(8, 4))
        liste = tk.Listbox(
            cadre, selectmode="extended", height=5, bg=Couleurs.FOND_JOURNAL, fg=Couleurs.TEXTE,
            font=Polices.MONO, relief="flat", highlightthickness=0, activestyle="none",
            selectbackground=Couleurs.BLEU, selectforeground=Couleurs.BLANC, exportselection=False,
        )
        defil = ttk.Scrollbar(cadre, orient="vertical", command=liste.yview)
        liste.configure(yscrollcommand=defil.set)
        defil.pack(side="right", fill="y")
        liste.pack(fill="both", expand=True)
        liste.codes = []  # type: ignore[attr-defined] — code associé à chaque ligne
        return liste

    def _ligne_selection(self, parent, liste: tk.Listbox) -> tk.Label:
        ligne = tk.Frame(parent, bg=Couleurs.FOND_CARTE)
        ligne.pack(fill="x")
        lbl = tk.Label(ligne, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
                       font=Polices.PETIT)
        lbl.pack(side="left")
        bouton(ligne, "Aucun", lambda: self._selectionner(liste, False), "discret",
               padx=8, pady=2).pack(side="right")
        bouton(ligne, "Tout", lambda: self._selectionner(liste, True), "discret",
               padx=8, pady=2).pack(side="right", padx=4)
        return lbl

    def _construire_audit(self, parent) -> None:
        c = carte(parent, "3 · Audit et rapport")
        c.pack(fill="x", pady=(8, 0))
        ligne = c.interieur
        self.btn_audit = bouton(ligne, "▶  Lancer l'audit", self._lancer_audit, "principal",
                                font=Polices.BOUTON, padx=18, pady=7)
        self.btn_audit.pack(side="left")
        self.btn_arret = bouton(ligne, "■ Arrêter", self._arreter, "discret", state="disabled")
        self.btn_arret.pack(side="left", padx=6)
        self.btn_rapport = bouton(ligne, "📊 Afficher le rapport", self._afficher_rapport, "succes")
        self.btn_rapport.pack(side="left", padx=(18, 0))
        self.btn_export = bouton(ligne, "💾 HTML", lambda: self._exporter("html"))
        self.btn_export.pack(side="left", padx=(6, 0))
        self.btn_export_xlsx = bouton(ligne, "📊 Excel", lambda: self._exporter("xlsx"))
        self.btn_export_xlsx.pack(side="left", padx=(6, 0))
        self.btn_export_csv = bouton(ligne, "📑 CSV", lambda: self._exporter("csv"))
        self.btn_export_csv.pack(side="left", padx=6)
        bouton(ligne, "⚙️ Colonnes auditées", self._configurer_colonnes, "discret").pack(side="right")
        self.lbl_bilan = tk.Label(c, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
                                  font=Polices.LABEL, justify="left", anchor="w")
        self.lbl_bilan.pack(fill="x", padx=12, pady=(0, 8))

    def _construire_mise_a_jour(self, parent) -> None:
        c = carte(parent, "4 · Mise à jour des ID non valides",
                  "même processus que complete_excel_files")
        c.pack(fill="x", pady=(8, 0))
        haut = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        haut.pack(fill="x")
        bouton(haut, "⚙️ Réglage des valeurs (règles)", self._ouvrir_regles).pack(side="left")
        bouton(haut, "↺ Valeurs par défaut", self._regles_par_defaut, "discret").pack(side="left", padx=6)
        self.btn_maj = bouton(haut, "🛠️  Mettre à jour les ID non valides",
                              lambda: self._lancer_mise_a_jour(None), "alerte", padx=16, pady=6)
        self.btn_maj.pack(side="right")

        self.lbl_regles = tk.Label(c.interieur, text="", bg=Couleurs.FOND_CARTE,
                                   fg=Couleurs.BLEU_CLAIR, font=Polices.JOURNAL, justify="left", anchor="w",
                                   wraplength=1000)
        self.lbl_regles.bind("<Configure>",
                             lambda e: self.lbl_regles.config(wraplength=max(e.width - 10, 200)))
        self.lbl_regles.pack(fill="x", pady=(6, 4))

        options = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        options.pack(fill="x")
        case(options, "Simulation (aucune écriture)", self.var_simulation, Couleurs.AMBRE).pack(side="left")
        case(options, "📅 « Date audit » = date du jour", self.var_date_du_jour,
             command=self._sauver_options).pack(side="left", padx=12)
        case(options, "🔐 Écrire via Excel (co-édition SharePoint)", self.var_excel, Couleurs.VERT,
             command=self._sauver_options,
             state="normal" if self._excel else "disabled").pack(side="left", padx=12)
        case(options, "🔁 Ré-auditer après mise à jour", self.var_reauditer,
             command=self._sauver_options).pack(side="left", padx=12)
        if not self._excel:
            tk.Label(c.interieur, text="Excel / pywin32 indisponible sur ce poste → réécriture directe "
                     "du fichier (à risque si quelqu'un l'a ouvert ailleurs).", bg=Couleurs.FOND_CARTE,
                     fg=Couleurs.AMBRE, font=Polices.PETIT).pack(anchor="w")

    def _construire_journal(self, parent) -> None:
        self.progress = ttk.Progressbar(parent, style="avf.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", pady=(10, 2))
        self.lbl_prog = tk.Label(parent, text="", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                                 font=Polices.LABEL, anchor="w")
        self.lbl_prog.pack(fill="x")
        self.txt_log = tk.Text(parent, height=6, bg=Couleurs.FOND_JOURNAL, fg=Couleurs.SUCCES,
                               font=Polices.JOURNAL, relief="flat", state="disabled", wrap="word")
        self.txt_log.pack(fill="both", expand=True, pady=(4, 0))

    # ══ Racine : dossier local ou SharePoint ══════════════════════
    def _source(self) -> SourceAudit | None:
        """Source courante : dossier local saisi, ou accès SharePoint établi (``None`` sinon)."""
        if self.var_source.get() == "sharepoint":
            return self._source_sp
        racine = self.var_racine.get().strip()
        return SourceAudit(LOCAL, racine) if racine else None

    def _racine(self) -> str:
        """Dossier parcouru pour trouver les communes (local, WebDAV ou dossier synchronisé)."""
        source = self._source()
        return source.racine if source else ""

    def _racine_ok(self) -> bool:
        if self.var_source.get() == "sharepoint":
            return self._source_sp is not None
        return os.path.isdir(self._racine())

    def _changer_source(self, initial: bool = False) -> None:
        sharepoint_choisi = self.var_source.get() == "sharepoint"
        (self.ligne_local if sharepoint_choisi else self.ligne_sp).pack_forget()
        (self.ligne_sp if sharepoint_choisi else self.ligne_local).pack(fill="x")
        if not initial:
            if self.params.source != self.var_source.get():
                self.params.source = self.var_source.get()
                self._sauver()
                self._remplir_communes([])     # les communes listées viennent de l'autre source
            self._restaurer_selection()
        self._maj_lbl_source()
        if not initial:
            self._maj_etat()

    def _maj_lbl_source(self, texte: str | None = None, couleur: str | None = None) -> None:
        if texte is None:
            if self.var_source.get() == "local":
                texte, couleur = ("Les fichiers sont lus dans ce dossier (copie OneDrive locale le cas échéant).",
                                  Couleurs.TEXTE_SECONDAIRE)
            elif self._source_sp is None:
                texte, couleur = ("Collez le lien SharePoint du dossier racine (barre d'adresse du navigateur), "
                                  "puis « Se connecter ». L'audit lira la version serveur des fichiers.",
                                  Couleurs.TEXTE_SECONDAIRE)
            else:
                s = self._source_sp
                texte = f"✅ Connecté — {s.libelle}\n   {s.racine}"
                if s.diagnostic:
                    texte += f"\n   Accès direct impossible : {s.diagnostic}"
                if s.mode == URL:
                    texte += ("\n   Le dossier ne peut pas être parcouru : renseignez les communes ou les "
                              "ID erreur (étape 2). Chaque fichier est ouvert sur le serveur par Excel.")
                if s.mode == EXCEL:
                    texte += "\n   (WebDAV refusé : chaque fichier est ouvert sur le serveur par Excel — plus lent)"
                couleur = Couleurs.SUCCES
        self.lbl_source.config(text=texte, fg=couleur or Couleurs.TEXTE_SECONDAIRE)

    def _connecter_sharepoint(self, silencieux: bool = False) -> None:
        """Établit l'accès au dossier SharePoint (thread de fond : WebDAV peut être lent à répondre)."""
        if self.occupe:
            return
        lien = self.var_lien.get().strip()
        try:
            sharepoint.analyser_lien(lien)
        except sharepoint.LienInvalide as exc:
            if not silencieux:
                messagebox.showwarning("Lien SharePoint", f"Lien inutilisable : {exc}.", parent=self)
            self._maj_lbl_source(f"❌ Lien inutilisable : {exc}", Couleurs.ERREUR)
            return
        if self.params.lien_sharepoint != lien:
            self.params.lien_sharepoint = lien
            self._sauver()
            self._remplir_communes([])
        self._source_sp = None
        self._debut_traitement()
        self._maj_lbl_source("⏳ Connexion au dossier SharePoint…", Couleurs.INFO)
        self.journaliser(f"🌐 Connexion à SharePoint : {lien}", Couleurs.INFO)
        threading.Thread(target=self._executer_connexion, args=(lien, silencieux), daemon=True).start()

    def _executer_connexion(self, lien: str, silencieux: bool) -> None:
        try:
            if sharepoint.navigateur_disponible():
                # Identification dans le navigateur, puis API SharePoint avec cette session.
                source = sharepoint.connecter_navigateur(
                    lien, on_etape=lambda m: self.journaliser(f"   · {m}", Couleurs.TEXTE_SECONDAIRE),
                    interactif=not silencieux,
                )
            else:
                source = sharepoint.connecter(lien)
        except Exception as exc:  # noqa: BLE001 — LienInvalide, AccesImpossible, ou imprévu
            # Le message est lu ici : « exc » n'existe plus quand le rappel s'exécute.
            message = str(exc) or exc.__class__.__name__
            self.planifier(lambda: self._echec_connexion(message, silencieux))
            return
        self.planifier(lambda: self._fin_connexion(source))

    def _fin_connexion(self, source: SourceAudit) -> None:
        self._source_sp = source
        self._fin_traitement()
        self.journaliser(f"✅ SharePoint connecté — {source.libelle} : {source.racine}", Couleurs.SUCCES)
        if source.mode == EXCEL and not self._excel:
            self.journaliser("⚠️ Excel est nécessaire pour lire les fichiers sur le serveur dans ce mode.",
                             Couleurs.AVERTISSEMENT)
        self._maj_lbl_source()
        self._restaurer_selection()
        self._maj_etat()

    def _echec_connexion(self, message: str, silencieux: bool) -> None:
        self._fin_traitement()
        premiere = message.split("\n")[0]
        self._maj_lbl_source(f"❌ Non connecté — {premiere}", Couleurs.ERREUR)
        self.journaliser(f"❌ SharePoint : {premiere}", Couleurs.ERREUR)
        if not silencieux:
            messagebox.showerror("Connexion SharePoint", message, parent=self)
        self._maj_etat()

    def _parcourir_racine(self) -> None:
        depart = self._racine() if os.path.isdir(self._racine()) else None
        chemin = filedialog.askdirectory(title="Dossier racine des audits", initialdir=depart, parent=self)
        if chemin:
            self.var_racine.set(os.path.normpath(chemin))
            self._enregistrer_racine()

    def _enregistrer_racine(self) -> None:
        racine = self._racine()
        if not racine:
            messagebox.showwarning("Dossier racine", "Indiquez un dossier racine.", parent=self)
            return
        self.params.racine = racine
        self._sauver()
        if os.path.isdir(racine):
            self.journaliser(f"📂 Dossier racine : {racine}", Couleurs.INFO)
        else:
            self.journaliser(f"⚠️ Dossier racine introuvable : {racine}", Couleurs.AVERTISSEMENT)
        self._maj_etat()

    def _verifier_racine(self) -> bool:
        if self.var_source.get() == "sharepoint":
            if self._source_sp is not None:
                return True
            messagebox.showwarning("Dossier SharePoint", "Connectez-vous d'abord au dossier SharePoint "
                                   "(étape 1 : lien puis « 🔗 Se connecter »).", parent=self)
            return False
        if os.path.isdir(self._racine()):
            if self.params.racine != self._racine():
                self.params.racine = self._racine()
                self._sauver()
            return True
        messagebox.showwarning("Dossier racine", f"Dossier introuvable :\n{self._racine() or '(vide)'}",
                               parent=self)
        return False

    # ══ Communes ══════════════════════════════════════════════════
    def _oublier_connexion(self) -> None:
        """Oublie la session SharePoint de l'outil (pour changer de compte, par exemple)."""
        if self.occupe:
            return
        if not messagebox.askyesno("Oublier la connexion", "Oublier la connexion SharePoint de l'outil ?\n"
                                   "Il faudra vous identifier de nouveau dans le navigateur.", parent=self):
            return
        sharepoint.oublier_connexion()
        self._source_sp = None
        self._remplir_communes([])
        self._maj_lbl_source()
        self._maj_etat()
        self.journaliser("🚪 Connexion SharePoint oubliée", Couleurs.INFO)

    def _existe(self, dossier: str) -> bool:
        return self._existe_pour(self._source())(dossier)

    @staticmethod
    def _existe_pour(source: SourceAudit | None) -> Callable[[str], bool]:
        """Test d'existence d'un dossier de commune, **utilisable hors du fil de l'interface**.

        Connexion navigateur : par l'arborescence SharePoint gardée en mémoire (un scan
        rapide, puis aucune requête quel que soit le nombre de communes). Sans parcours
        possible (mode URL) : la commune est retenue, son fichier sera cherché à son adresse.
        """
        if source is not None and source.mode == API:
            acces = sharepoint.AccesApi(source)
            return lambda dossier: acces.existe(os.path.basename(dossier.rstrip("\\/")))
        if source is not None and not source.listable:
            return lambda dossier: True
        return os.path.isdir

    def _en_fond(self, libelle: str, travail: Callable[[], object], fin: Callable[[object], None]) -> bool:
        """Exécute ``travail()`` dans un thread, puis ``fin(résultat)`` sur le fil de l'interface.

        Toute opération qui peut attendre le réseau (SharePoint, WebDAV, disque
        partagé) passe par ici : l'interface ne se fige jamais (« Ne répond pas »).
        """
        if self.occupe:
            return False
        self._debut_traitement()
        if libelle:
            self.journaliser(libelle, Couleurs.INFO)

        def tache() -> None:
            try:
                resultat = travail()
            except Exception as exc:  # noqa: BLE001 — remonté au journal, jamais un plantage
                message = str(exc) or exc.__class__.__name__
                self.planifier(lambda: (self._fin_traitement(),
                                        self.journaliser(f"❌ {message}", Couleurs.ERREUR)))
                return
            self.planifier(lambda: (self._fin_traitement(), fin(resultat)))

        threading.Thread(target=tache, daemon=True).start()
        return True

    def _detecter_communes(self) -> None:
        if not self._verifier_racine():
            return
        if self._source().mode == API:
            self._detecter_communes_api()
            return
        if not self._source().listable:
            messagebox.showinfo(
                "Communes", "Le dossier SharePoint ne peut pas être parcouru dans ce mode d'accès : "
                "renseignez les communes (✏️) ou les ID erreur (🎯).", parent=self)
            return
        self._effacer_ids(silencieux=True)
        communes = detection.lister_communes(self._racine())
        if not communes:
            messagebox.showinfo("Communes", "Aucune commune trouvée sous la racine "
                                "(<racine>/Dep<XX>/<commune>).", parent=self)
        avec_fichier = [c.code for c in communes if c.a_fichier]
        self._remplir_communes(communes, selection=avec_fichier)
        sans = len(communes) - len(avec_fichier)
        self.journaliser(
            f"🔍 {len(communes)} commune(s) détectée(s)"
            + (f" — {sans} sans fichier audit (non sélectionnée(s))" if sans else ""), Couleurs.INFO)

    def _detecter_communes_api(self) -> None:
        """Liste les communes par l'API SharePoint (thread de fond : une requête par Dep<XX>)."""
        if self.occupe:
            return
        self._effacer_ids(silencieux=True)
        self._debut_traitement()
        self.journaliser("🔍 Recherche des communes sur SharePoint…", Couleurs.INFO)
        source = self._source()

        def travail() -> None:
            try:
                trouvees = sharepoint.AccesApi(source).lister_communes()
            except Exception as exc:  # noqa: BLE001
                message = str(exc) or exc.__class__.__name__
                self.planifier(lambda: (self._fin_traitement(),
                                        self.journaliser(f"❌ Communes non listées : {message}", Couleurs.ERREUR)))
                return
            communes = [detection.CommuneDetectee(code, dep, source.url_fichier(code).rsplit("/", 1)[0], a)
                        for code, dep, a in trouvees]
            self.planifier(lambda: (self._fin_traitement(), self._communes_detectees(communes)))

        threading.Thread(target=travail, daemon=True).start()

    def _communes_detectees(self, communes) -> None:
        if not communes:
            messagebox.showinfo("Communes", "Aucune commune trouvée sous la racine "
                                "(<racine>/Dep<XX>/<commune>).", parent=self)
        avec_fichier = [c.code for c in communes if c.a_fichier]
        self._remplir_communes(communes, selection=avec_fichier)
        sans = len(communes) - len(avec_fichier)
        self.journaliser(
            f"🔍 {len(communes)} commune(s) détectée(s)"
            + (f" — {sans} sans fichier audit (non sélectionnée(s))" if sans else ""), Couleurs.INFO)

    def _renseigner_communes(self) -> None:
        if not self._verifier_racine():
            return
        texte = DialogueSaisie(
            self, "Renseigner des communes",
            "Saisissez ou collez des codes INSEE (séparés par des virgules, espaces ou retours à "
            "la ligne), ou chargez un fichier (colonne « INSEE »). Le dossier Dep<XX> est déduit "
            "des deux premiers caractères du code.",
            "13001, 13055  2A004",
            texte="\n".join(self._codes(self.liste_com, selection=True)),
            colonnes_fichier=COLONNES_INSEE,
        ).montrer()
        if texte is None:
            return
        codes = list(dict.fromkeys(detection.parser_communes(texte)))
        if not codes:
            messagebox.showinfo("Communes", "Aucun code INSEE dans la saisie.", parent=self)
            return
        self._effacer_ids(silencieux=True)
        racine, existe = self._racine(), self._existe_pour(self._source())

        def fin(verif) -> None:
            if verif.absentes:
                messagebox.showwarning(
                    "Communes introuvables",
                    f"{len(verif.absentes)} commune(s) sans dossier <racine>/Dep<XX>/<commune> — ignorée(s) :\n"
                    + _apercu(verif.absentes), parent=self)
            self._remplir_communes(verif.trouvees, selection=[c.code for c in verif.trouvees])
            self.journaliser(f"✏️ {len(verif.trouvees)} commune(s) renseignée(s)"
                             + (f", {len(verif.absentes)} introuvable(s)" if verif.absentes else ""),
                             Couleurs.INFO)

        self._en_fond(f"🔎 Vérification de {len(codes)} commune(s)…",
                      lambda: detection.verifier_communes(racine, codes, existe), fin)

    # ── Audit restreint à des ID erreur ───────────────────────────
    def _renseigner_ids(self, type_ids: str = PAR_ID_ERREUR) -> None:
        """Audit restreint à des ID saisis : ID erreur (🎯) ou ID imm (🆔)."""
        if not self._verifier_racine():
            return
        par_imm = type_ids == PAR_ID_IMM
        nom = self.params.colonnes.id_imm if par_imm else self.params.colonnes.id_erreur
        precedents = self.params.ids_cibles if self.params.type_ids == type_ids else []
        texte = DialogueSaisie(
            self, f"Auditer seulement des {nom}",
            (f"Collez les {nom} à auditer (séparés par des virgules, espaces ou retours à la ligne), "
             f"ou chargez un fichier (colonne « {nom} »). "
             + ("La commune est donnée par les 5 caractères après « ID- » (ID-58006000005V → 58006) ; "
                if par_imm else "La commune est déduite du code INSEE en préfixe ; ")
             + "seules les lignes de ces ID sont contrôlées, puis mises à jour."),
            "ID-58006000005V  ID-2B030000006D" if par_imm else "13001_2, 13001_15  2A004_1",
            texte="\n".join(precedents),
            colonnes_fichier=COLONNES_ID_IMM if par_imm else COLONNES_ID,
        ).montrer()
        if texte is None:
            return
        racine, existe = self._racine(), self._existe_pour(self._source())

        def fin(saisie) -> None:
            problemes = []
            if saisie.absents:
                ids = [i for v in saisie.absents.values() for i in v]
                problemes.append(f"{len(saisie.absents)} commune(s) sans dossier <racine>/Dep<XX>/<commune> — "
                                 f"leurs {len(ids)} {nom} sont ignorés :\n" + _apercu(list(saisie.absents)))
            if saisie.codes_sans_id:
                problemes.append("Codes INSEE saisis sans numéro d'ID (format <insee>_<numéro>) — ignorés :\n"
                                 + _apercu(saisie.codes_sans_id))
            if problemes:
                messagebox.showwarning(nom, "\n\n".join(problemes), parent=self)
            if not saisie.groupes:
                format_ = "ID-58006000005V" if par_imm else "13001_2"
                messagebox.showinfo(nom, f"Aucun {nom} exploitable (format attendu : {format_}).", parent=self)
                return
            self.ids_cibles = saisie.groupes
            self.type_cibles = type_ids
            self.params.ids_cibles = saisie.ids
            self.params.type_ids = type_ids
            self._remplir_communes(saisie.trouvees, selection=list(saisie.groupes))
            self._maj_perimetre()
            self.journaliser(f"{'🆔' if par_imm else '🎯'} {saisie.nb_ids} {nom} saisi(s) sur "
                             f"{len(saisie.groupes)} commune(s) — seules leurs lignes seront auditées",
                             Couleurs.INFO)

        self._en_fond(f"🔎 Vérification des {nom} saisis…",
                      lambda: detection.verifier_ids(racine, texte, existe, type_ids), fin)

    def _vider_communes(self) -> None:
        """Vide la liste des communes (codes INSEE) — et les ID saisis qui s'y rapportaient."""
        if self.occupe or not (self.liste_com.codes or self.ids_cibles):
            return
        nombre = len(self.liste_com.codes)
        self._effacer_ids(silencieux=True)
        self._remplir_communes([])
        self.journaliser(f"🗑 Liste des communes vidée ({nombre} commune(s))", Couleurs.INFO)

    def _vider_ids(self) -> None:
        """Oublie les ID erreur saisis et les communes qui en venaient (repartir de zéro)."""
        if self.occupe or not self.ids_cibles:
            return
        nombre = sum(len(v) for v in self.ids_cibles.values())
        self._effacer_ids(silencieux=True)
        self._remplir_communes([])
        self.journaliser(f"🗑 {nombre} {self._nom_ids()} saisi(s) vidé(s), avec leurs communes", Couleurs.INFO)

    def _effacer_ids(self, silencieux: bool = False) -> None:
        """Revient à l'audit des communes entières."""
        if not self.ids_cibles and not self.params.ids_cibles:
            return
        self.ids_cibles = {}
        self.params.ids_cibles = []
        self._sauver()
        self._maj_perimetre()
        self._maj_etat()
        if not silencieux:
            self.journaliser("🗂 Audit des communes entières (ID erreur saisis oubliés)", Couleurs.INFO)

    def _maj_perimetre(self) -> None:
        if not self.ids_cibles:
            self.cadre_perimetre.pack_forget()
            return
        ids = [i for v in self.ids_cibles.values() for i in v]
        apercu = ", ".join(ids[:10]) + (f"… (+{len(ids) - 10})" if len(ids) > 10 else "")
        self.lbl_perimetre.config(
            text=f"{'🆔' if self.type_cibles == PAR_ID_IMM else '🎯'} Audit limité à {len(ids)} "
                 f"{self._nom_ids()} sur {len(self.ids_cibles)} commune(s) : {apercu}")
        self.cadre_perimetre.pack(fill="x", pady=(6, 0), after=self.btn_det_com.master)

    def _nom_ids(self) -> str:
        """« ID imm » ou « ID erreur » : nature des ID saisis."""
        return self.params.colonnes.id_imm if self.type_cibles == PAR_ID_IMM else self.params.colonnes.id_erreur

    def _ids_pour(self, communes: list[str]) -> dict[str, list[str]] | None:
        """ID erreur à auditer par commune, ou ``None`` pour les communes entières."""
        if not self.ids_cibles:
            return None
        return {c: self.ids_cibles[c] for c in communes if c in self.ids_cibles}

    def _remplir_communes(self, communes, selection: list[str] | None = None) -> None:
        self._remplir(self.liste_com, [
            (c.code, f"{c.code:<7} Dep{c.departement:<3}" + ("" if c.a_fichier else "  (sans fichier audit)"))
            for c in communes
        ], selection)
        self._sauver_selection()
        self._maj_etat()

    def _communes_choisies(self) -> list[str]:
        return self._codes(self.liste_com, selection=True)

    # ── Listes ────────────────────────────────────────────────────
    @staticmethod
    def _remplir(liste: tk.Listbox, elements: list[tuple[str, str]], selection=None) -> None:
        # Un seul appel Tk pour remplir, et une sélection par plages contiguës :
        # des milliers de communes (4 000 sur un vrai SharePoint) s'affichent sans attendre.
        liste.delete(0, tk.END)
        liste.codes = [code for code, _ in elements]
        if elements:
            liste.insert(tk.END, *(texte for _, texte in elements))
        voulus = set(liste.codes if selection is None else selection)
        debut = None
        for rang, code in enumerate(liste.codes + [None]):
            if code is not None and code in voulus:
                debut = rang if debut is None else debut
            elif debut is not None:
                liste.selection_set(debut, rang - 1)
                debut = None

    @staticmethod
    def _codes(liste: tk.Listbox, selection: bool = False) -> list[str]:
        if not selection:
            return list(liste.codes)
        return [liste.codes[i] for i in liste.curselection()]

    def _selectionner(self, liste: tk.Listbox, tout: bool) -> None:
        if tout:
            liste.selection_set(0, tk.END)
        else:
            liste.selection_clear(0, tk.END)
        self._maj_etat()

    def _restaurer_selection(self) -> None:
        """Recharge les communes sélectionnées au dernier lancement (si elles existent toujours).

        En arrière-plan : la vérification peut passer par le réseau (SharePoint).
        """
        if not self._racine_ok():
            return
        racine, existe = self._racine(), self._existe_pour(self._source())
        ids_cibles, communes = list(self.params.ids_cibles), list(self.params.communes)
        if ids_cibles:                             # audit restreint à des ID erreur
            type_ids = self.params.type_ids
            groupes = (detection.grouper_ids_imm(ids_cibles) if type_ids == PAR_ID_IMM
                       else grouper_par_commune(ids_cibles))

            def fin_ids(verif) -> None:
                self.ids_cibles = {c.code: groupes[c.code] for c in verif.trouvees}
                self.type_cibles = type_ids
                voulues = [c for c in communes if c in self.ids_cibles] or list(self.ids_cibles)
                self._remplir_communes(verif.trouvees, selection=voulues)
                self._maj_perimetre()

            self._en_fond("", lambda: detection.verifier_communes(racine, list(groupes), existe), fin_ids)
            return
        if not communes:
            return
        self._en_fond("", lambda: detection.verifier_communes(racine, communes, existe),
                      lambda verif: self._remplir_communes(verif.trouvees))

    def _lancer_audit(self, communes: list[str] | None = None, apres: Callable | None = None) -> None:
        if self.occupe or not self._verifier_racine():
            return
        communes = communes if communes is not None else self._communes_choisies()
        if not communes:
            messagebox.showinfo("Audit", "Sélectionnez au moins une commune (étape 2).", parent=self)
            return
        self._sauver_selection()
        self._debut_traitement()
        if apres is None:
            self._vider_journal_ui()
        source = self._source()
        ids = self._ids_pour(communes)
        self.journaliser(f"🚀 Audit de {len(communes)} commune(s)", Couleurs.INFO)
        if ids is not None:
            self.journaliser(f"🎯 Limité à {sum(len(v) for v in ids.values())} {self._nom_ids()} saisi(s)",
                             Couleurs.INFO)
        self.journaliser(f"🗂 Source : {source.libelle} — {source.affichage}", Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(
            f"🔎 Ligne validée si : {' · '.join(p.libelles_criteres(self.params.colonnes).values())}"
            f" — identifiant des lignes : « {self.params.colonnes.id_imm} »", Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        params = self.params
        threading.Thread(target=self._executer_audit, args=(communes, source, params, apres, ids,
                                                             self.type_cibles), daemon=True).start()

    def _executer_audit(self, communes, source: SourceAudit, params: Parametres, apres, ids=None,
                        type_ids: str = PAR_ID_ERREUR) -> None:
        try:
            if source.mode in (EXCEL, URL) and not excel_disponible():
                raise ExcelIndisponible("Excel est nécessaire pour lire les fichiers sur le serveur "
                                        "(WebDAV refusé)")
            with sharepoint.acces_pour(source) as acces:
                rapport = auditer_communes(
                    communes, source.racine, colonne_id=params.colonnes.id_erreur, colonnes=params.colonnes,
                    ligne_entete=params.regles.ligne_entete,
                    on_progression=self._progresser,
                    on_resultat=lambda r: self.journaliser(*_ligne_audit(r)),
                    doit_arreter=self._arret.is_set,
                    paralleles=params.paralleles,
                    source=source, acces=acces, ids=ids, type_ids=type_ids,
                )
        except Exception as exc:  # noqa: BLE001 — filet de sécurité
            self.journaliser(f"❌ Audit interrompu : {exc}", Couleurs.ERREUR)
            self.planifier(self._fin_traitement)
            return
        self.planifier(lambda: self._fin_audit(rapport, apres))

    def _fin_audit(self, rapport: RapportAudit, apres) -> None:
        interrompu = self._arret.is_set()
        if apres is not None and self.rapport is not None:
            self.rapport.remplacer(rapport.communes)
            self.rapport.duree += rapport.duree
        else:
            self.rapport = rapport
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        for ligne in p.resume(self.rapport):
            self.journaliser(f"📊 {ligne}", Couleurs.SUCCES)
        if interrompu:
            self.journaliser("⏹ Audit arrêté à la demande — rapport partiel", Couleurs.AVERTISSEMENT)
        self._fin_traitement()
        self._maj_bilan()
        self._signaler_telechargement_interdit(c.message for c in rapport.communes)
        if self.fenetre_rapport is not None and self.fenetre_rapport.winfo_exists():
            self.fenetre_rapport.rafraichir(self.rapport)
        if apres is not None:
            apres()
        elif not interrompu:
            self._afficher_rapport()

    def _arreter(self) -> None:
        self._arret.set()
        self.btn_arret.config(state="disabled")

    # ══ Rapport ═══════════════════════════════════════════════════
    def _afficher_rapport(self) -> None:
        if self.rapport is None:
            return
        if self.fenetre_rapport is not None and self.fenetre_rapport.winfo_exists():
            self.fenetre_rapport.rafraichir(self.rapport)
            self.fenetre_rapport.lift()
            self.fenetre_rapport.focus_force()
            return
        self.fenetre_rapport = FenetreRapport(
            self, self.rapport, on_exporter=self._exporter, on_mettre_a_jour=self._lancer_mise_a_jour,
        )

    def _exporter(self, format: str = "html", chemin: str | None = None,
                  ouvrir: bool | None = None) -> str | None:
        """Exporte le rapport en ``html``, ``xlsx`` ou ``csv`` (``chemin`` : sans boîte de dialogue)."""
        if self.rapport is None:
            return None
        parent = self.fenetre_rapport if (self.fenetre_rapport and self.fenetre_rapport.winfo_exists()) else self
        libelle, motif = FORMATS.get(format, FORMATS["html"])
        if chemin is None:
            dossier = self.params.dossier_export if os.path.isdir(self.params.dossier_export) else self._racine()
            chemin = filedialog.asksaveasfilename(
                parent=parent, title=f"Exporter le rapport — {libelle}", defaultextension=f".{format}",
                initialdir=dossier if os.path.isdir(dossier) else None,
                initialfile=nom_fichier_defaut(self.rapport, format),
                filetypes=[(libelle, motif), ("Tous les fichiers", "*.*")],
            )
            if not chemin:
                return None
        try:
            chemin = exporter(self.rapport, chemin, format)
        except OSError as exc:
            messagebox.showerror("Export", f"Le rapport n'a pas pu être écrit :\n{exc}\n\n"
                                 "S'il est ouvert dans Excel, fermez-le puis réessayez.", parent=parent)
            return None
        self.params.dossier_export = os.path.dirname(chemin)
        self._sauver()
        self.journaliser(f"💾 Rapport exporté ({libelle}) : {chemin}", Couleurs.INFO)
        application = "le navigateur" if format == "html" else "Excel"
        if ouvrir is None:
            ouvrir = messagebox.askyesno("Export", f"Rapport enregistré :\n{chemin}\n\n"
                                         f"L'ouvrir dans {application} ?", parent=parent)
        if ouvrir:
            (ouvrir_dans_navigateur if format == "html" else ouvrir_fichier)(chemin)
        return chemin

    # ══ Colonnes et règles ════════════════════════════════════════
    def _configurer_colonnes(self) -> None:
        resultat = DialogueColonnes(self, self.params.colonnes, self.params.colonnes.id_imm,
                                    self.params.regles.ligne_entete).montrer()
        if resultat is None:
            return
        self.params.colonnes, self.params.regles.colonne_id, self.params.regles.ligne_entete = resultat
        self._sauver()
        self._maj_resume_regles()
        self.journaliser("⚙️ Colonnes auditées mises à jour — relancez l'audit pour en tenir compte",
                         Couleurs.INFO)

    def _ouvrir_regles(self) -> None:
        PanneauRegles(self, self.params.regles, on_save=self._enregistrer_regles,
                      fichier_exemple=self._fichier_exemple())

    def _fichier_exemple(self) -> str:
        if self.rapport is not None:
            premier = next((c.chemin for c in self.rapport.lues), "")
            if premier:
                return premier
        if not os.path.isdir(self._racine()):
            return ""
        return trouver_fichier_exemple(self._racine()) or ""

    def _enregistrer_regles(self, regles: ReglesConfig) -> None:
        self.params.regles = regles
        self._sauver()
        self._maj_resume_regles()
        self.journaliser("⚙️ Règles de mise à jour enregistrées", Couleurs.INFO)

    def _regles_par_defaut(self) -> None:
        if messagebox.askyesno(
            "Valeurs par défaut",
            "Restaurer les règles par défaut de complete_excel_files ?\n\n"
            + "\n".join(r.resume() for r in ReglesConfig.defaut().regles), parent=self,
        ):
            courant = self.params.regles
            defaut = ReglesConfig.defaut()
            defaut.colonne_id, defaut.ligne_entete = courant.colonne_id, courant.ligne_entete
            self._enregistrer_regles(defaut)

    def _regles_effectives(self) -> ReglesConfig:
        return preparer_regles(self.params.regles, self.params.colonnes.date_audit,
                               self.var_date_du_jour.get())

    def _maj_resume_regles(self) -> None:
        regles = self._regles_effectives()
        resumes = [r.resume().replace(" (actif) :", " :") for r in regles.regles]
        self.lbl_regles.config(text="   ·   ".join(resumes) or "Aucune règle configurée.",
                               fg=Couleurs.BLEU_CLAIR if regles.actives else Couleurs.AVERTISSEMENT)

    # ══ Mise à jour ═══════════════════════════════════════════════
    def _lancer_mise_a_jour(self, communes: list[str] | None = None) -> None:
        if self.occupe:
            return
        parent = self.fenetre_rapport if (self.fenetre_rapport and self.fenetre_rapport.winfo_exists()) else self
        if self.rapport is None or not self.rapport.nb_ids_non_valides:
            messagebox.showinfo("Mise à jour", "Aucun ID erreur non valide dans le rapport : "
                                "lancez d'abord l'audit.", parent=parent)
            return
        if not self._verifier_racine():
            return
        source = self._source()
        if not _meme_source(self.rapport.source, source):
            messagebox.showwarning(
                "Mise à jour", "Le rapport a été produit sur une autre source que celle choisie à "
                f"l'étape 1 :\n{self.rapport.racine}\n\nRelancez l'audit avant de mettre à jour.", parent=parent)
            return
        groupes = groupes_depuis_rapport(self.rapport, communes)
        if not groupes:
            messagebox.showinfo("Mise à jour", "Aucun ID à mettre à jour pour cette sélection.", parent=parent)
            return
        regles = self._regles_effectives()
        if not regles.actives:
            messagebox.showwarning("Mise à jour", "Aucune règle applicable : ouvrez « ⚙️ Réglage des "
                                   "valeurs » pour les renseigner.", parent=parent)
            return

        simulation = self.var_simulation.get()
        via_excel = self.var_excel.get() and self._excel and not simulation
        par_api = source.mode == API and not simulation
        if source.sharepoint and not simulation and not par_api:
            if not self._excel:
                messagebox.showerror("Mise à jour", "La mise à jour d'un dossier SharePoint passe par Excel "
                                     "(co-édition) : Excel / pywin32 sont introuvables sur ce poste.", parent=parent)
                return
            via_excel = True   # jamais de réécriture directe d'un fichier sur le serveur
        nb_ids = sum(len(v) for v in groupes.values())
        conflits = [c.commune for c in self.rapport.communes
                    if c.commune in groupes and any("conflit" in a for a in c.avertissements)]
        texte = (
            f"Mettre à jour {p.nombre(nb_ids)} ID imm non valide(s) "
            f"sur {len(groupes)} commune(s) ?\n\n"
            + "\n".join(r.resume() for r in regles.actives)
            + "\n\n" + ("🔍 Mode simulation : aucun fichier ne sera modifié." if simulation else
                        ("⚠️ Les fichiers du dossier SharePoint seront modifiés" if source.sharepoint
                         else "⚠️ Les fichiers audit d'origine seront modifiés")
                        + ", ligne par ligne (ID imm) : " + _mode_ecriture(par_api, via_excel, self._excel))
        )
        if conflits:
            texte += ("\n\n⚠ Copie de conflit présente — ces communes seront refusées : "
                      + ", ".join(conflits))
        if not messagebox.askyesno("Confirmer la mise à jour", texte, parent=parent):
            return

        self._debut_traitement()
        self._vider_journal_ui()
        self.journaliser(f"🛠️ Mise à jour de {p.nombre(nb_ids)} ID imm sur {len(groupes)} commune(s)"
                         + (" (simulation)" if simulation else ""), Couleurs.INFO)
        for regle in regles.actives:
            self.journaliser(f"   ⚙️ {regle.resume()}", Couleurs.TEXTE_SECONDAIRE)
        if not simulation:
            self.journaliser("✍ " + _mode_ecriture(par_api, via_excel, self._excel), Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        threading.Thread(
            target=self._executer_mise_a_jour,
            args=(groupes, regles, source, simulation, via_excel, self.var_reauditer.get()),
            daemon=True,
        ).start()

    def _executer_mise_a_jour(self, groupes, regles, source, simulation, via_excel, reauditer) -> None:
        try:
            bilan = mettre_a_jour(
                groupes, regles, source.racine, simuler=simulation, via_excel=via_excel, source=source,
                colonne_id=self.params.colonnes.id_imm,
                fichier_attente=self.fichier_attente,
                on_progression=self._progresser,
                on_etape=lambda c, m: self.journaliser(f"   · {c} — {m}", Couleurs.TEXTE_SECONDAIRE),
                on_resultat=lambda r: self.journaliser(*_ligne_maj(r)),
                paralleles=self.params.paralleles,
                doit_arreter=self._arret.is_set,
            )
        except EcritureSharePointImpossible as exc:
            self.journaliser(f"❌ {exc} — aucune écriture faite.", Couleurs.ERREUR)
            self.planifier(self._fin_traitement)
            return
        except ExcelIndisponible as exc:
            self.journaliser(f"❌ {exc} — aucune écriture faite. Décochez « Écrire via Excel » pour "
                             "utiliser la réécriture directe.", Couleurs.ERREUR)
            self.planifier(self._fin_traitement)
            return
        except Exception as exc:  # noqa: BLE001
            self.journaliser(f"❌ Mise à jour interrompue : {exc}", Couleurs.ERREUR)
            self.planifier(self._fin_traitement)
            return
        self.planifier(lambda: self._fin_mise_a_jour(bilan, reauditer))

    def _signaler_telechargement_interdit(self, messages) -> bool:
        """Accès conditionnel de l'organisation : une explication claire, une seule fois."""
        if not any(TELECHARGEMENT_INTERDIT in (m or "") for m in messages):
            return False
        self.journaliser("⛔ Téléchargement des fichiers SharePoint interdit depuis cet appareil "
                         "(stratégie de l'organisation)", Couleurs.ERREUR)
        messagebox.showwarning("Accès SharePoint limité par votre organisation",
                               sharepoint_rest.TelechargementInterdit.EXPLICATION, parent=self)
        return True

    def _fin_mise_a_jour(self, bilan, reauditer: bool) -> None:
        self._signaler_telechargement_interdit(r.message for r in bilan.resultats)
        mode = " (simulation)" if bilan.simulation else ""
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        self.journaliser(
            f"🎉 {bilan.nb_communes_ok}/{bilan.nb_communes} fichier(s) traité(s) — "
            f"{p.nombre(bilan.nb_modifications)} modification(s){mode} — {p.duree(bilan.duree)}",
            Couleurs.SUCCES)
        if bilan.interrompu:
            self.journaliser(f"■ Arrêt demandé : {bilan.nb_communes_demandees - bilan.nb_communes} "
                             "commune(s) non commencée(s), aucun fichier laissé à moitié écrit",
                             Couleurs.AVERTISSEMENT)
        if bilan.sans_excel and not bilan.simulation:
            self.journaliser(f"ℹ Co-édition Excel non utilisée pour certains fichiers : {bilan.sans_excel}",
                             Couleurs.AVERTISSEMENT)
        if bilan.echecs:
            self.journaliser("Pourquoi ces ID imm n'ont pas été mis à jour :", Couleurs.AVERTISSEMENT)
            for ligne in p.refus_mise_a_jour(bilan.echecs):
                self.journaliser(ligne, Couleurs.AVERTISSEMENT)
            self.journaliser(f"⏳ {len(bilan.echecs)} ID non appliqué(s) → liste « ID en attente » "
                             f"({bilan.nb_en_attente} au total, {self.fichier_attente})",
                             Couleurs.AVERTISSEMENT)
        self._fin_traitement()

        def bilan_final():
            if reauditer and not bilan.simulation and self.rapport is not None:
                for ligne in p.encore_non_valides(self.rapport, bilan.reussis):
                    self.journaliser(ligne, Couleurs.AVERTISSEMENT)
            details = (f"\n\n⚠️ {len(bilan.echecs)} ID non appliqué(s) — raisons dans le journal."
                       if bilan.echecs else "")
            if bilan.interrompu:
                details += (f"\n\n■ Arrêté : {bilan.nb_communes_demandees - bilan.nb_communes} "
                            "commune(s) non commencée(s).")
            if bilan.simulation:
                details += "\n\n(simulation : aucun fichier modifié)"
            messagebox.showinfo(
                "Mise à jour terminée",
                f"✅ {bilan.nb_communes_ok}/{bilan.nb_communes} fichier(s) traité(s)\n"
                f"{p.nombre(bilan.nb_modifications)} modification(s){details}",
                parent=self.fenetre_rapport if (self.fenetre_rapport and self.fenetre_rapport.winfo_exists()) else self,
            )

        if reauditer and not bilan.simulation and bilan.communes_traitees:
            self.journaliser("🔁 Ré-audit des communes mises à jour…", Couleurs.INFO)
            self._lancer_audit(bilan.communes_traitees, apres=bilan_final)
        else:
            bilan_final()

    # ══ État de l'interface ═══════════════════════════════════════
    def _debut_traitement(self) -> None:
        self.occupe = True
        self._arret.clear()
        self._maj_etat()

    def _fin_traitement(self) -> None:
        self.occupe = False
        self._maj_etat()

    def _maj_etat(self) -> None:
        racine_ok = self._racine_ok()
        self.btn_connecter.config(state="normal" if not self.occupe else "disabled")
        self.btn_oublier.config(state="normal" if not self.occupe else "disabled")
        for bouton_radio in self.choix_source.values():
            bouton_radio.config(state="normal" if not self.occupe else "disabled")
        nb_com, nb_com_sel = len(self.liste_com.codes), len(self.liste_com.curselection())
        self.lbl_com.config(text=f"{nb_com_sel}/{nb_com} sélectionnée(s)")
        libre = not self.occupe

        def etat(condition: bool) -> str:
            return "normal" if condition and libre else "disabled"

        source = self._source()
        self.btn_det_com.config(state=etat(racine_ok and (source is None or source.listable)))
        self.btn_saisie_com.config(state=etat(racine_ok))
        self.btn_saisie_ids.config(state=etat(racine_ok))
        self.btn_saisie_imm.config(state=etat(racine_ok))
        self.btn_effacer_ids.config(state=etat(True))
        self.btn_vider_com.config(state=etat(bool(self.liste_com.codes or self.ids_cibles)))
        self.btn_vider_ids.config(state=etat(bool(self.ids_cibles)))
        self.btn_audit.config(state=etat(racine_ok and nb_com_sel > 0))
        self.btn_arret.config(state="normal" if self.occupe else "disabled")
        a_rapport = self.rapport is not None
        self.btn_rapport.config(state=etat(a_rapport))
        for b in (self.btn_export, self.btn_export_xlsx, self.btn_export_csv):
            b.config(state=etat(a_rapport))
        self.btn_maj.config(state=etat(a_rapport and bool(self.rapport and self.rapport.nb_ids_non_valides)))
        if a_rapport:
            self.btn_maj.config(text=f"🛠️  Mettre à jour les {p.nombre(self.rapport.nb_ids_non_valides)} "
                                     "ID non valides")

    def _maj_bilan(self) -> None:
        if self.rapport is not None:
            self.lbl_bilan.config(text="\n".join(p.resume(self.rapport)))

    # ══ Persistance ═══════════════════════════════════════════════
    def _sauver(self) -> None:
        if not enregistrer(self.params, self.fichier_config):
            self.journaliser("⚠️ Configuration non enregistrée (fichier inaccessible)", Couleurs.AVERTISSEMENT)

    def _sauver_selection(self) -> None:
        self.params.communes = self._communes_choisies()
        self._sauver()

    def _sauver_options(self) -> None:
        self.params.date_du_jour = self.var_date_du_jour.get()
        self.params.reauditer = self.var_reauditer.get()
        if self._excel:
            self.params.ecrire_via_excel = self.var_excel.get()
        self._sauver()
        self._maj_resume_regles()

    # ══ File thread → interface ═══════════════════════════════════
    def planifier(self, rappel: Callable[[], None]) -> None:
        """Exécute ``rappel`` sur le thread principal (appelable de n'importe quel thread)."""
        self._file.put(rappel)

    def journaliser(self, message: str, couleur: str = Couleurs.SUCCES) -> None:
        self.planifier(lambda: self._ecrire_journal(message, couleur))

    def _vider_file(self) -> None:
        try:
            for _ in range(500):  # borne : l'interface reste réactive
                self._file.get_nowait()()
        except queue.Empty:
            pass
        self.after(PERIODE_FILE, self._vider_file)

    def traiter_file(self) -> None:
        """Vide la file immédiatement (utilisé par les tests)."""
        while True:
            try:
                self._file.get_nowait()()
            except queue.Empty:
                return

    def _ecrire_journal(self, message: str, couleur: str) -> None:
        tag = f"c{couleur.replace('#', '')}"
        self.txt_log.config(state="normal")
        self.txt_log.tag_config(tag, foreground=couleur)
        self.txt_log.insert(tk.END, message + "\n", tag)
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def _vider_journal_ui(self) -> None:
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state="disabled")

    def _progresser(self, fait: int, total: int, libelle: str = "") -> None:
        def maj():
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = fait
            self.lbl_prog.config(text=f"{fait}/{total} — {libelle}")
        self.planifier(maj)


def _meme_source(du_rapport: SourceAudit | None, courante: SourceAudit | None) -> bool:
    """Le rapport porte-t-il sur la source choisie (même mode d'accès et même dossier) ?"""
    if du_rapport is None or courante is None:
        return True
    return (du_rapport.sharepoint, os.path.normcase(du_rapport.racine)) == \
        (courante.sharepoint, os.path.normcase(courante.racine))


TELECHARGEMENT_INTERDIT = "téléchargement interdit depuis cet appareil"


def _apercu(elements: list[str], nombre: int = 40) -> str:
    """Liste lisible dans une boîte de dialogue, même pour des centaines d'éléments."""
    texte = ", ".join(elements[:nombre])
    return texte + (f"… (+{len(elements) - nombre})" if len(elements) > nombre else "")


# ── Lignes de journal ─────────────────────────────────────────────
def _ligne_audit(c: AuditCommune) -> tuple[str, str]:
    e = p.etat(c)
    if not c.lisible:
        return f"{e.icone} {c.commune} — {c.message}", Couleurs.ERREUR
    texte = (f"{e.icone} {c.commune} — {p.nombre(c.nb_lignes)} ligne(s) · "
             f"{p.nombre(c.nb_lignes_valides)} validée(s) · {p.nombre(c.nb_lignes_non_valides)} non valide(s)"
             f" ({p.nombre(c.nb_ids_non_valides)}/{p.nombre(c.nb_ids)} ID) · {p.duree(c.duree)}")
    for a in c.avertissements:
        texte += f"\n   ⚠ {a}"
    couleur = Couleurs.AVERTISSEMENT if (c.nb_lignes_non_valides or c.avertissements) else Couleurs.SUCCES
    return texte, couleur


def _mode_ecriture(par_api: bool, via_excel: bool, excel: bool) -> str:
    """Comment les cellules seront écrites (confirmation et journal)."""
    if via_excel:
        texte = ("cellules écrites via Excel en co-édition, même si le fichier est ouvert par d'autres "
                 "(valeurs relues après enregistrement)")
        return texte + (" ; repli : API SharePoint (navigateur)" if par_api else "")
    if par_api:
        base = ("API SharePoint avec votre connexion (navigateur), jamais par-dessus une version enregistrée "
                "entre-temps")
    else:
        base = "réécriture directe du fichier (openpyxl)"
    return base + (" ; un fichier ouvert par quelqu'un est écrit via Excel en co-édition" if excel
                   else " ; Excel absent : un fichier ouvert par quelqu'un est refusé (raison au journal)")


def _ligne_maj(r: ResultatCommune) -> tuple[str, str]:
    couleur, icone = STYLE_STATUT[r.statut]
    if r.statut is not Statut.SUCCES:
        return f"{icone} {r.commune} — {r.message}", couleur
    texte = (f"{icone} {r.commune} — {r.nb_modifications} modification(s) sur {r.nb_lignes} ligne(s) · "
             f"{r.resume_ids()} · {r.resume_duree()}")
    if r.resume_detail():
        texte += f"  [{r.resume_detail()}]"
    if r.ids_absents:
        couleur = Couleurs.AVERTISSEMENT
        texte += "\n   ↳ introuvables : " + ", ".join(r.ids_absents[:8])
    if r.ids_non_verifies:
        couleur = Couleurs.AVERTISSEMENT
        texte += "\n   ↳ écriture non confirmée dans SharePoint : " + ", ".join(r.ids_non_verifies[:8])
    return texte, couleur
