"""Fenêtre principale de l'outil (Tkinter).

Déroulé, de haut en bas :

    1. dossier racine ;
    2. communes — toutes celles de ``<racine>/Dep<XX>/<commune>`` détectées,
       ou codes INSEE renseignés (le dossier ``Dep<XX>`` en est déduit) ;
    3. audit (thread de fond, journal en temps réel), rapport en popup,
       export HTML ;
    4. mise à jour des ID sans Validation finale — moteur et règles de
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
from ..models.audit import AuditCommune, RapportAudit
from ..services import detection
from ..services.audit import auditer_communes
from ..services.export import exporter_html, nom_fichier_defaut, ouvrir_dans_navigateur
from ..services.mise_a_jour import groupes_depuis_rapport, mettre_a_jour, preparer_regles
from ..views import presentation as p
from .composants import (
    DialogueColonnes, DialogueSaisie, bouton, carte, case, champ, configurer_styles,
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
        self.geometry("1100x940")
        self.minsize(960, 780)
        self.configure(bg=Couleurs.FOND)
        configurer_styles(self)

        self.var_racine = tk.StringVar(value=self.params.racine)
        self.var_simulation = tk.BooleanVar(value=False)
        self.var_date_du_jour = tk.BooleanVar(value=self.params.date_du_jour)
        self.var_excel = tk.BooleanVar(value=self.params.ecrire_via_excel and self._excel)
        self.var_reauditer = tk.BooleanVar(value=self.params.reauditer)

        self._construire()
        self._restaurer_selection()
        self._maj_resume_regles()
        self._maj_etat()
        self.after(PERIODE_FILE, self._vider_file)

    # ══ Construction ══════════════════════════════════════════════
    def _construire(self) -> None:
        entete = tk.Frame(self, bg=Couleurs.FOND_HEADER, pady=10)
        entete.pack(fill="x")
        tk.Label(entete, text="  🔎  Audit de la Validation finale", font=Polices.TITRE,
                 bg=Couleurs.FOND_HEADER, fg=Couleurs.TEXTE).pack(side="left")
        tk.Label(entete, text="<racine>/Dep<XX>/<commune>/audit_*.xlsx  ", font=Polices.MONO,
                 bg=Couleurs.FOND_HEADER, fg=Couleurs.TEXTE_SECONDAIRE).pack(side="right")

        corps = tk.Frame(self, bg=Couleurs.FOND)
        corps.pack(fill="both", expand=True, padx=16, pady=10)

        self._construire_racine(corps)
        self._construire_communes(corps)
        self._construire_audit(corps)
        self._construire_mise_a_jour(corps)
        self._construire_journal(corps)

    def _construire_racine(self, parent) -> None:
        c = carte(parent, "1 · Dossier racine")
        c.pack(fill="x")
        ligne = c.interieur
        champ(ligne, self.var_racine).pack(side="left", fill="x", expand=True)
        bouton(ligne, "📁 Parcourir", self._parcourir_racine).pack(side="left", padx=(6, 0))
        bouton(ligne, "Enregistrer", self._enregistrer_racine, "succes").pack(side="left", padx=(6, 0))

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
        self.liste_com = self._liste(c.interieur, colonnes=4)
        self.liste_com.bind("<<ListboxSelect>>", lambda e: (self._sauver_selection(), self._maj_etat()))
        self.lbl_com = self._ligne_selection(c.interieur, self.liste_com)

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
        self.btn_export = bouton(ligne, "💾 Exporter en HTML", self._exporter)
        self.btn_export.pack(side="left", padx=6)
        bouton(ligne, "⚙️ Colonnes auditées", self._configurer_colonnes, "discret").pack(side="right")
        self.lbl_bilan = tk.Label(c, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
                                  font=Polices.LABEL, justify="left", anchor="w")
        self.lbl_bilan.pack(fill="x", padx=12, pady=(0, 8))

    def _construire_mise_a_jour(self, parent) -> None:
        c = carte(parent, "4 · Mise à jour des ID sans Validation finale",
                  "même processus que complete_excel_files")
        c.pack(fill="x", pady=(8, 0))
        haut = tk.Frame(c.interieur, bg=Couleurs.FOND_CARTE)
        haut.pack(fill="x")
        bouton(haut, "⚙️ Réglage des valeurs (règles)", self._ouvrir_regles).pack(side="left")
        bouton(haut, "↺ Valeurs par défaut", self._regles_par_defaut, "discret").pack(side="left", padx=6)
        self.btn_maj = bouton(haut, "🛠️  Mettre à jour les ID sans Validation finale",
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

    # ══ Racine ════════════════════════════════════════════════════
    def _racine(self) -> str:
        return self.var_racine.get().strip()

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
        if os.path.isdir(self._racine()):
            if self.params.racine != self._racine():
                self.params.racine = self._racine()
                self._sauver()
            return True
        messagebox.showwarning("Dossier racine", f"Dossier introuvable :\n{self._racine() or '(vide)'}",
                               parent=self)
        return False

    # ══ Communes ══════════════════════════════════════════════════
    def _detecter_communes(self) -> None:
        if not self._verifier_racine():
            return
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

    def _renseigner_communes(self) -> None:
        if not self._verifier_racine():
            return
        texte = DialogueSaisie(
            self, "Renseigner des communes",
            "Saisissez un ou plusieurs codes INSEE (séparés par des virgules, espaces ou retours à "
            "la ligne). Le dossier Dep<XX> est déduit des deux premiers caractères du code.",
            "13001, 13055  2A004",
            texte=" ".join(self._codes(self.liste_com, selection=True)),
        ).montrer()
        if texte is None:
            return
        verif = detection.verifier_communes(self._racine(), detection.parser_communes(texte))
        if verif.absentes:
            messagebox.showwarning(
                "Communes introuvables",
                "Ces communes n'ont pas de dossier <racine>/Dep<XX>/<commune> et sont ignorées :\n"
                + ", ".join(verif.absentes), parent=self)
        self._remplir_communes(verif.trouvees, selection=[c.code for c in verif.trouvees])
        if verif.trouvees:
            self.journaliser(f"✏️ {len(verif.trouvees)} commune(s) renseignée(s)", Couleurs.INFO)

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
        liste.delete(0, tk.END)
        liste.codes = [code for code, _ in elements]
        for _, texte in elements:
            liste.insert(tk.END, texte)
        voulus = set(liste.codes if selection is None else selection)
        for rang, code in enumerate(liste.codes):
            if code in voulus:
                liste.selection_set(rang)

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
        """Recharge les communes sélectionnées au dernier lancement (si elles existent toujours)."""
        if not os.path.isdir(self._racine()) or not self.params.communes:
            return
        verif = detection.verifier_communes(self._racine(), self.params.communes)
        self._remplir_communes(verif.trouvees)

    # ══ Audit ═════════════════════════════════════════════════════
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
        self.journaliser(f"🚀 Audit de {len(communes)} commune(s)", Couleurs.INFO)
        self.journaliser(
            f"🔎 Colonnes : « {self.params.regles.colonne_id} », « {self.params.colonnes.validation} »"
            f" (+ {', '.join(self.params.colonnes.complementaires().values())})",
            Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        racine, params = self._racine(), self.params
        threading.Thread(target=self._executer_audit, args=(communes, racine, params, apres),
                         daemon=True).start()

    def _executer_audit(self, communes, racine, params: Parametres, apres) -> None:
        try:
            rapport = auditer_communes(
                communes, racine, colonne_id=params.regles.colonne_id, colonnes=params.colonnes,
                ligne_entete=params.regles.ligne_entete,
                on_progression=self._progresser,
                on_resultat=lambda r: self.journaliser(*_ligne_audit(r)),
                doit_arreter=self._arret.is_set,
            )
        except Exception as exc:  # noqa: BLE001 — filet de sécurité, jamais attendu
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

    def _exporter(self, chemin: str | None = None, ouvrir: bool | None = None) -> str | None:
        if self.rapport is None:
            return None
        parent = self.fenetre_rapport if (self.fenetre_rapport and self.fenetre_rapport.winfo_exists()) else self
        if chemin is None:
            dossier = self.params.dossier_export if os.path.isdir(self.params.dossier_export) else self._racine()
            chemin = filedialog.asksaveasfilename(
                parent=parent, title="Exporter le rapport HTML", defaultextension=".html",
                initialdir=dossier if os.path.isdir(dossier) else None,
                initialfile=nom_fichier_defaut(self.rapport),
                filetypes=[("Page HTML", "*.html"), ("Tous les fichiers", "*.*")],
            )
            if not chemin:
                return None
        try:
            chemin = exporter_html(self.rapport, chemin)
        except OSError as exc:
            messagebox.showerror("Export", f"Le rapport n'a pas pu être écrit :\n{exc}", parent=parent)
            return None
        self.params.dossier_export = os.path.dirname(chemin)
        self._sauver()
        self.journaliser(f"💾 Rapport exporté : {chemin}", Couleurs.INFO)
        if ouvrir is None:
            ouvrir = messagebox.askyesno("Export", f"Rapport enregistré :\n{chemin}\n\n"
                                         "L'ouvrir dans le navigateur ?", parent=parent)
        if ouvrir:
            ouvrir_dans_navigateur(chemin)
        return chemin

    # ══ Colonnes et règles ════════════════════════════════════════
    def _configurer_colonnes(self) -> None:
        resultat = DialogueColonnes(self, self.params.colonnes, self.params.regles.colonne_id,
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
        if self.rapport is None or not self.rapport.nb_vides:
            messagebox.showinfo("Mise à jour", "Aucun ID sans Validation finale dans le rapport : "
                                "lancez d'abord l'audit.", parent=parent)
            return
        if not self._verifier_racine():
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
        nb_ids = sum(len(v) for v in groupes.values())
        conflits = [c.commune for c in self.rapport.communes
                    if c.commune in groupes and any("conflit" in a for a in c.avertissements)]
        texte = (
            f"Mettre à jour {p.nombre(nb_ids)} ID erreur sans Validation finale "
            f"sur {len(groupes)} commune(s) ?\n\n"
            + "\n".join(r.resume() for r in regles.actives)
            + "\n\n" + ("🔍 Mode simulation : aucun fichier ne sera modifié." if simulation else
                        "⚠️ Les fichiers audit d'origine seront modifiés"
                        + (" via Excel (co-édition SharePoint)." if via_excel else " (réécriture directe)."))
        )
        if conflits:
            texte += ("\n\n⚠ Copie de conflit présente — ces communes seront refusées : "
                      + ", ".join(conflits))
        if not messagebox.askyesno("Confirmer la mise à jour", texte, parent=parent):
            return

        self._debut_traitement()
        self._vider_journal_ui()
        self.journaliser(f"🛠️ Mise à jour de {p.nombre(nb_ids)} ID erreur sur {len(groupes)} commune(s)"
                         + (" (simulation)" if simulation else ""), Couleurs.INFO)
        for regle in regles.actives:
            self.journaliser(f"   ⚙️ {regle.resume()}", Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(
            "🔐 Écriture via Excel (co-édition SharePoint) — valeurs vérifiées après enregistrement"
            if via_excel else "📝 Réécriture directe du fichier (openpyxl)", Couleurs.TEXTE_SECONDAIRE)
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        threading.Thread(
            target=self._executer_mise_a_jour,
            args=(groupes, regles, self._racine(), simulation, via_excel, self.var_reauditer.get()),
            daemon=True,
        ).start()

    def _executer_mise_a_jour(self, groupes, regles, racine, simulation, via_excel, reauditer) -> None:
        try:
            bilan = mettre_a_jour(
                groupes, regles, racine, simuler=simulation, via_excel=via_excel,
                fichier_attente=self.fichier_attente,
                on_progression=self._progresser,
                on_etape=lambda c, m: self.journaliser(f"   · {c} — {m}", Couleurs.TEXTE_SECONDAIRE),
                on_resultat=lambda r: self.journaliser(*_ligne_maj(r)),
            )
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

    def _fin_mise_a_jour(self, bilan, reauditer: bool) -> None:
        mode = " (simulation)" if bilan.simulation else ""
        self.journaliser(SEPARATEUR, Couleurs.SEPARATEUR)
        self.journaliser(
            f"🎉 {bilan.nb_communes_ok}/{bilan.nb_communes} fichier(s) traité(s) — "
            f"{p.nombre(bilan.nb_modifications)} modification(s){mode} — {p.duree(bilan.duree)}",
            Couleurs.SUCCES)
        if bilan.echecs:
            self.journaliser(f"⏳ {len(bilan.echecs)} ID non appliqué(s) → liste « ID en attente » "
                             f"({bilan.nb_en_attente} au total, {self.fichier_attente})",
                             Couleurs.AVERTISSEMENT)
        self._fin_traitement()

        def bilan_final():
            details = f"\n\n⚠️ {len(bilan.echecs)} ID non appliqué(s) — voir le journal." if bilan.echecs else ""
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
        racine_ok = os.path.isdir(self._racine())
        nb_com, nb_com_sel = len(self.liste_com.codes), len(self.liste_com.curselection())
        self.lbl_com.config(text=f"{nb_com_sel}/{nb_com} sélectionnée(s)")
        libre = not self.occupe

        def etat(condition: bool) -> str:
            return "normal" if condition and libre else "disabled"

        self.btn_det_com.config(state=etat(racine_ok))
        self.btn_saisie_com.config(state=etat(racine_ok))
        self.btn_audit.config(state=etat(racine_ok and nb_com_sel > 0))
        self.btn_arret.config(state="normal" if self.occupe else "disabled")
        a_rapport = self.rapport is not None
        self.btn_rapport.config(state=etat(a_rapport))
        self.btn_export.config(state=etat(a_rapport))
        self.btn_maj.config(state=etat(a_rapport and bool(self.rapport and self.rapport.nb_vides)))
        if a_rapport:
            self.btn_maj.config(text=f"🛠️  Mettre à jour les {p.nombre(self.rapport.nb_vides)} ID "
                                     "sans Validation finale")

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


# ── Lignes de journal ─────────────────────────────────────────────
def _ligne_audit(c: AuditCommune) -> tuple[str, str]:
    e = p.etat(c)
    if not c.lisible:
        return f"{e.icone} {c.commune} — {c.message}", Couleurs.ERREUR
    texte = (f"{e.icone} {c.commune} — {p.nombre(c.nb_ids)} ID · {p.nombre(c.nb_renseignes)} renseigné(s) · "
             f"{p.nombre(c.nb_vides)} vide(s) · {p.duree(c.duree)}")
    for a in c.avertissements:
        texte += f"\n   ⚠ {a}"
    couleur = Couleurs.AVERTISSEMENT if (c.nb_vides or c.avertissements) else Couleurs.SUCCES
    return texte, couleur


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
