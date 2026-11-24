"""Panneau « Configuration des règles » — constructeur de règles visuel.

Modèle repris du projet ``excel_rule_editor`` : une règle se compose de

    zéro, une ou plusieurs CONDITIONS combinées en ET   +   une ACTION

Les lignes réellement modifiées restent celles dont l'ID erreur a été saisi
dans la fenêtre principale ; les conditions sont un **filtre supplémentaire
facultatif** appliqué à ces lignes.

Le programme est livré avec **16 règles d'audit par défaut**
(``models.regles.REGLES_DEFAUT``). Elles sont modifiables librement et l'état
est conservé d'un démarrage à l'autre dans ``config.json`` ; le bouton
« ↺ Réinitialiser (config par défaut) » restaure la configuration livrée.

Fonctionnalités :
    - désigner la colonne contenant les ID erreur (celle qui sert de filtre) ;
    - ajouter / supprimer / désactiver une règle ;
    - restaurer les 16 règles d'audit par défaut (bouton « Réinitialiser ») ;
    - ajouter / supprimer des conditions : colonne choisie dans la LISTE RÉELLE
      du fichier Excel, opérateur adapté au TYPE de la colonne (numérique →
      = ≠ > < ≥ ≤ ; texte → = ≠ contient… ; date → comparaisons de dates), et
      comparaison à une VALEUR ou à une autre COLONNE ;
    - pour chaque règle, l'ACTION : colonne écrite (celle testée par défaut) et
      valeur inscrite — valeur libre, recopie d'une colonne, ou effacement ;
    - lorsque la colonne écrite est une colonne de dates et que l'action pose
      une valeur libre, la saisie bascule en mode date : gabarit JJ/MM/AAAA,
      bouton « aujourd'hui » et contrôle immédiat de ce qui a été compris ;
    - une option reprend la mise en forme (police, taille, format d'affichage)
      des valeurs déjà présentes dans la colonne.

Les colonnes sont détectées automatiquement dans un fichier audit de la racine
(``<racine>/Dep<XX>/<insee>/audit_*.xlsx``) ; un autre fichier peut être choisi
manuellement.

À l'enregistrement, la configuration reconstruite est renvoyée via ``on_save``.
"""

from __future__ import annotations

from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ..models.regles import (
    Action, COLONNE_ID_DEFAUT, Condition, LIBELLE_SOURCE, OPERATEURS_UNAIRES,
    RegleModification, ReglesConfig, SOURCE_PAR_LIBELLE, SourceValeur, TYPE_DATE,
    TYPE_TEXTE, TypeDroite, operateurs_pour_type,
)
from ..services.colonnes_excel import detecter_colonnes, normaliser
from ..services.valeurs import analyser_date, formater_date
from .theme import Couleurs, Polices

# Rappel de format affiché sous une saisie de date.
GABARIT_DATE = "JJ/MM/AAAA"

_SOURCES = list(LIBELLE_SOURCE.values())        # ["valeur", "colonne", "effacer"]
_TYPES_DROITE = [TypeDroite.VALEUR.value, TypeDroite.COLONNE.value]


class PanneauRegles(tk.Toplevel):
    def __init__(
        self,
        parent,
        config: ReglesConfig,
        on_save,
        fichier_exemple: str = "",
    ):
        super().__init__(parent)
        self.title("⚙️ Configuration des règles de modification")
        self.geometry("1080x760")
        self.configure(bg=Couleurs.FOND)
        self.transient(parent)
        self.on_save = on_save

        self.colonnes = []           # list[Colonne]
        self.noms: list[str] = []    # libellés proposés dans les menus déroulants
        self.info = {}               # référence normalisée -> Colonne
        self.fichier_colonnes = ""

        copie = ReglesConfig.from_dict(config.to_dict())
        self.var_entete = tk.StringVar(value=str(copie.ligne_entete))
        self.var_colonne_id = tk.StringVar(value=copie.colonne_id or COLONNE_ID_DEFAUT)
        self.var_style = tk.BooleanVar(value=copie.harmoniser_style)
        self.regles_ui = [self._regle_vers_ui(r) for r in copie.regles]

        self._construire()
        if fichier_exemple:
            self._charger_colonnes(fichier_exemple, silencieux=True)
        self._afficher()

    # ── Conversion modèle ↔ variables Tk ──────────────────────────
    def _regle_vers_ui(self, r: RegleModification) -> dict:
        return {
            "code": r.code,
            "libelle": tk.StringVar(value=r.libelle),
            "actif": tk.BooleanVar(value=r.actif),
            "exact": tk.BooleanVar(value=r.exact),
            "conditions": [self._condition_vers_ui(c) for c in r.conditions],
            "action": {
                "colonne": tk.StringVar(value=r.action.colonne),
                "source": tk.StringVar(value=LIBELLE_SOURCE[r.action.source]),
                "valeur": tk.StringVar(value=r.action.valeur),
            },
        }

    def _condition_vers_ui(self, c: Condition) -> dict:
        return {
            "colonne": tk.StringVar(value=c.colonne),
            "operateur": tk.StringVar(value=c.operateur),
            "type_droite": tk.StringVar(value=c.type_droite.value),
            "droite": tk.StringVar(value=c.droite),
        }

    def _ui_vers_config(self) -> ReglesConfig:
        regles = []
        for ru in self.regles_ui:
            conditions = []
            for cu in ru["conditions"]:
                colonne = cu["colonne"].get().strip()
                if not colonne:
                    continue
                conditions.append(Condition(
                    colonne=colonne,
                    operateur=cu["operateur"].get().strip() or "=",
                    type_droite=TypeDroite(cu["type_droite"].get().strip() or "valeur"),
                    droite=cu["droite"].get().strip(),
                ))
            au = ru["action"]
            regles.append(RegleModification(
                code=ru["code"],
                libelle=ru["libelle"].get().strip() or ru["code"],
                action=Action(
                    colonne=au["colonne"].get().strip(),
                    source=SOURCE_PAR_LIBELLE.get(au["source"].get(), SourceValeur.VALEUR),
                    valeur=au["valeur"].get(),
                ),
                actif=ru["actif"].get(),
                conditions=conditions,
                exact=ru["exact"].get(),
            ))
        return ReglesConfig(
            regles=regles,
            ligne_entete=_entier(self.var_entete.get(), 1),
            colonne_id=self.var_colonne_id.get().strip() or COLONNE_ID_DEFAUT,
            harmoniser_style=self.var_style.get(),
        )

    # ── Construction de la fenêtre ────────────────────────────────
    def _construire(self) -> None:
        entete = tk.Frame(self, bg=Couleurs.FOND_PANNEAU)
        entete.pack(fill="x")
        tk.Label(
            entete, text="⚙️  Constructeur de règles", font=Polices.TITRE,
            bg=Couleurs.FOND_PANNEAU, fg=Couleurs.TEXTE,
        ).pack(anchor="w", padx=20, pady=(12, 0))
        tk.Label(
            entete,
            text="Une règle = des conditions en ET (facultatives) + une modification. "
                 "Les lignes modifiées sont celles dont l'ID erreur a été saisi ;\n"
                 "les conditions restreignent encore ce choix. Une règle sans valeur "
                 "n'est pas appliquée : vider une cellule se demande par « effacer ».",
            font=Polices.LABEL, bg=Couleurs.FOND_PANNEAU, fg=Couleurs.TEXTE_SECONDAIRE,
            justify="left",
        ).pack(anchor="w", padx=20)

        # Repères de lecture du classeur
        reperes = tk.Frame(entete, bg=Couleurs.FOND_PANNEAU)
        reperes.pack(fill="x", padx=20, pady=(10, 2))
        tk.Label(
            reperes, text="Colonne « ID erreur » :", bg=Couleurs.FOND_PANNEAU,
            fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL,
        ).pack(side="left")
        self.combo_id = ttk.Combobox(
            reperes, textvariable=self.var_colonne_id, values=self.noms,
            font=Polices.LABEL, width=30,
        )
        self.combo_id.pack(side="left", padx=6)
        self.lbl_badge_id = tk.Label(
            reperes, text="", bg=Couleurs.FOND_PANNEAU, fg=Couleurs.TEXTE_TERTIAIRE,
            font=Polices.PETIT,
        )
        self.lbl_badge_id.pack(side="left")
        self.combo_id.bind("<<ComboboxSelected>>", lambda e: self._afficher())

        tk.Entry(
            reperes, textvariable=self.var_entete, width=4, justify="center",
            bg=Couleurs.FOND_JOURNAL, fg=Couleurs.TEXTE, font=Polices.LABEL,
            relief="flat", bd=4, insertbackground=Couleurs.BLANC,
        ).pack(side="right")
        tk.Label(
            reperes, text="Ligne d'en-tête :", bg=Couleurs.FOND_PANNEAU,
            fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL,
        ).pack(side="right", padx=6)

        barre = tk.Frame(entete, bg=Couleurs.FOND_PANNEAU)
        barre.pack(fill="x", padx=20, pady=(6, 12))
        tk.Button(
            barre, text="📄 Colonnes depuis un autre fichier…", command=self._parcourir,
            bg=Couleurs.FOND_PANNEAU2, fg=Couleurs.BLEU_CLAIR, font=Polices.BOUTON_PETIT,
            relief="flat", cursor="hand2", padx=10, pady=4,
        ).pack(side="left")
        self.lbl_colonnes = tk.Label(
            barre, text="aucune colonne détectée — saisie libre possible",
            bg=Couleurs.FOND_PANNEAU, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT,
        )
        self.lbl_colonnes.pack(side="left", padx=10)
        tk.Checkbutton(
            barre, text="Reprendre la mise en forme de la colonne (police, format)",
            variable=self.var_style, bg=Couleurs.FOND_PANNEAU,
            fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT,
            activebackground=Couleurs.FOND_PANNEAU, activeforeground=Couleurs.TEXTE,
            selectcolor=Couleurs.FOND_PANNEAU2, cursor="hand2",
        ).pack(side="right")

        # Zone défilante contenant les cartes de règles
        conteneur = tk.Frame(self, bg=Couleurs.FOND)
        conteneur.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(conteneur, bg=Couleurs.FOND, highlightthickness=0)
        ascenseur = ttk.Scrollbar(conteneur, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=ascenseur.set)
        ascenseur.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.support = tk.Frame(self.canvas, bg=Couleurs.FOND)
        fenetre = self.canvas.create_window((0, 0), window=self.support, anchor="nw")
        self.support.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(fenetre, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._molette)

        # Barre d'actions
        pied = tk.Frame(self, bg=Couleurs.FOND_PANNEAU)
        pied.pack(fill="x")
        tk.Button(
            pied, text="＋ Ajouter une règle", command=self._ajouter_regle,
            bg=Couleurs.FOND_PANNEAU2, fg=Couleurs.VERT, font=Polices.BOUTON_PETIT,
            relief="flat", cursor="hand2", padx=12, pady=8,
        ).pack(side="left", padx=16, pady=10)
        tk.Button(
            pied, text="↺ Réinitialiser (config par défaut)", command=self._reinitialiser,
            bg=Couleurs.FOND_PANNEAU2, fg=Couleurs.AMBRE, font=Polices.BOUTON_PETIT,
            relief="flat", cursor="hand2", padx=12, pady=8,
        ).pack(side="left", pady=10)
        tk.Button(
            pied, text="💾 Enregistrer", command=self._enregistrer,
            bg=Couleurs.VERT, fg="#062017", font=Polices.BOUTON_PETIT,
            relief="flat", cursor="hand2", padx=18, pady=8,
        ).pack(side="right", padx=16, pady=10)
        tk.Button(
            pied, text="Annuler", command=self.destroy,
            bg=Couleurs.FOND_PANNEAU2, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL,
            relief="flat", cursor="hand2", padx=14, pady=8,
        ).pack(side="right", pady=10)

    # ── Rendu ─────────────────────────────────────────────────────
    def _afficher(self) -> None:
        for widget in self.support.winfo_children():
            widget.destroy()

        self.lbl_badge_id.config(text=self._badge(self.var_colonne_id.get()))

        if not self.regles_ui:
            tk.Label(
                self.support, text="Aucune règle. Cliquez sur « ＋ Ajouter une règle ».",
                bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.CORPS,
            ).pack(pady=30)

        for rang, ru in enumerate(self.regles_ui):
            self._afficher_regle(rang, ru)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _afficher_regle(self, rang: int, ru: dict) -> None:
        carte = tk.Frame(
            self.support, bg=Couleurs.FOND_CARTE,
            highlightbackground=Couleurs.BORDURE, highlightthickness=1,
        )
        carte.pack(fill="x", pady=6, padx=12)

        ligne = tk.Frame(carte, bg=Couleurs.FOND_CARTE)
        ligne.pack(fill="x", padx=10, pady=(8, 4))
        tk.Checkbutton(
            ligne, variable=ru["actif"], bg=Couleurs.FOND_CARTE,
            activebackground=Couleurs.FOND_CARTE, selectcolor=Couleurs.FOND_PANNEAU2,
            cursor="hand2",
        ).pack(side="left")
        tk.Label(
            ligne, text=ru["code"], bg=Couleurs.FOND_CARTE, fg=Couleurs.BLEU_CLAIR,
            font=Polices.SOUS_TITRE, width=5, anchor="w",
        ).pack(side="left")
        tk.Entry(
            ligne, textvariable=ru["libelle"], font=Polices.CORPS,
            bg=Couleurs.FOND_JOURNAL, fg=Couleurs.TEXTE, insertbackground=Couleurs.BLANC,
            relief="flat", bd=5,
        ).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(
            ligne, text="🗑", command=lambda i=rang: self._supprimer_regle(i),
            bg=Couleurs.FOND_CARTE, fg=Couleurs.ROUGE, font=Polices.CORPS,
            relief="flat", cursor="hand2",
        ).pack(side="right")
        tk.Checkbutton(
            ligne, text="Aa exact", variable=ru["exact"], bg=Couleurs.FOND_CARTE,
            fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT,
            activebackground=Couleurs.FOND_CARTE, activeforeground=Couleurs.TEXTE,
            selectcolor=Couleurs.FOND_PANNEAU2, cursor="hand2",
        ).pack(side="right", padx=8)

        for index, cu in enumerate(ru["conditions"]):
            self._afficher_condition(carte, rang, index, cu)

        tk.Button(
            carte, text="＋ condition", command=lambda i=rang: self._ajouter_condition(i),
            bg=Couleurs.FOND_CARTE, fg=Couleurs.BLEU_CLAIR, font=Polices.LABEL,
            relief="flat", cursor="hand2",
        ).pack(anchor="w", padx=14, pady=(2, 4))

        self._afficher_action(carte, ru)

    def _afficher_condition(self, carte, rang: int, index: int, cu: dict) -> None:
        ligne = tk.Frame(carte, bg=Couleurs.FOND_CARTE)
        ligne.pack(fill="x", padx=14, pady=2)
        tk.Label(
            ligne, text=("SI" if index == 0 else "ET"), bg=Couleurs.FOND_CARTE,
            fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL_GRAS, width=5, anchor="w",
        ).pack(side="left")

        combo_colonne = ttk.Combobox(
            ligne, textvariable=cu["colonne"], values=self.noms, font=Polices.LABEL, width=26
        )
        combo_colonne.pack(side="left", padx=2)
        combo_colonne.bind("<<ComboboxSelected>>", lambda e: self._afficher())

        type_colonne = self._type_de(cu["colonne"].get())
        operateurs = operateurs_pour_type(type_colonne)
        if cu["operateur"].get() not in operateurs:
            cu["operateur"].set(operateurs[0])
        combo_op = ttk.Combobox(
            ligne, textvariable=cu["operateur"], values=operateurs, state="readonly",
            width=14, font=Polices.LABEL,
        )
        combo_op.pack(side="left", padx=2)
        combo_op.bind("<<ComboboxSelected>>", lambda e: self._afficher())

        if cu["operateur"].get() not in OPERATEURS_UNAIRES:
            combo_type = ttk.Combobox(
                ligne, textvariable=cu["type_droite"], values=_TYPES_DROITE,
                state="readonly", width=8, font=Polices.LABEL,
            )
            combo_type.pack(side="left", padx=2)
            combo_type.bind("<<ComboboxSelected>>", lambda e: self._afficher())

            if cu["type_droite"].get() == TypeDroite.COLONNE.value:
                ttk.Combobox(
                    ligne, textvariable=cu["droite"], values=self.noms,
                    font=Polices.LABEL, width=26,
                ).pack(side="left", padx=2)
            else:
                tk.Entry(
                    ligne, textvariable=cu["droite"], font=Polices.MONO,
                    bg=Couleurs.FOND_JOURNAL, fg=Couleurs.TEXTE,
                    insertbackground=Couleurs.BLANC, relief="flat", bd=5, width=28,
                ).pack(side="left", padx=2)

        tk.Button(
            ligne, text="✕", command=lambda: self._supprimer_condition(rang, index),
            bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL,
            relief="flat", cursor="hand2",
        ).pack(side="right")
        tk.Label(
            ligne, text=self._badge(cu["colonne"].get()), bg=Couleurs.FOND_CARTE,
            fg=Couleurs.TEXTE_TERTIAIRE, font=Polices.PETIT,
        ).pack(side="right", padx=6)

    def _afficher_action(self, carte, ru: dict) -> None:
        """Ligne « ALORS COLONNE ← VALEUR » de la règle."""
        ligne = tk.Frame(carte, bg=Couleurs.FOND_CARTE)
        ligne.pack(fill="x", padx=14, pady=(2, 10))
        tk.Label(
            ligne, text="ALORS", bg=Couleurs.FOND_CARTE, fg=Couleurs.VERT,
            font=Polices.LABEL_GRAS, width=5, anchor="w",
        ).pack(side="left")

        combo_cible = ttk.Combobox(
            ligne, textvariable=ru["action"]["colonne"], values=self.noms,
            font=Polices.LABEL, width=26,
        )
        combo_cible.pack(side="left", padx=2)
        combo_cible.bind("<<ComboboxSelected>>", lambda e: self._afficher())
        tk.Label(
            ligne, text="←", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
            font=Polices.CORPS,
        ).pack(side="left", padx=4)

        combo_source = ttk.Combobox(
            ligne, textvariable=ru["action"]["source"], values=_SOURCES,
            state="readonly", width=8, font=Polices.LABEL,
        )
        combo_source.pack(side="left", padx=2)
        combo_source.bind("<<ComboboxSelected>>", lambda e: self._afficher())

        source = SOURCE_PAR_LIBELLE.get(ru["action"]["source"].get(), SourceValeur.VALEUR)
        cible_ref = ru["action"]["colonne"].get() or self._premiere_colonne_condition(ru)
        est_date = self._type_de(cible_ref) == TYPE_DATE

        tk.Label(
            ligne, text=self._badge(cible_ref), bg=Couleurs.FOND_CARTE,
            fg=Couleurs.TEXTE_TERTIAIRE, font=Polices.PETIT,
        ).pack(side="right", padx=6)

        if source is SourceValeur.VIDE:
            tk.Label(
                ligne, text="🧹 cellule vidée", bg=Couleurs.FOND_CARTE,
                fg=Couleurs.AMBRE, font=Polices.PETIT,
            ).pack(side="left", padx=6)
            if not ru["action"]["colonne"].get().strip() and not ru["conditions"]:
                self._note_colonne_testee(ligne)
            return

        if source is SourceValeur.COLONNE:
            ttk.Combobox(
                ligne, textvariable=ru["action"]["valeur"], values=self.noms,
                font=Polices.LABEL, width=28,
            ).pack(side="left", padx=2)
            if not ru["action"]["colonne"].get().strip() and not ru["conditions"]:
                self._note_colonne_testee(ligne)
            return

        # source == VALEUR
        tk.Label(
            ligne, text="DATE" if est_date else "VALEUR", bg=Couleurs.FOND_CARTE,
            fg=Couleurs.AMBRE if est_date else Couleurs.TEXTE_SECONDAIRE,
            font=Polices.LABEL_GRAS,
        ).pack(side="left", padx=(4, 4))
        saisie = tk.Entry(
            ligne, textvariable=ru["action"]["valeur"], font=Polices.MONO,
            bg=Couleurs.FOND_JOURNAL, fg=Couleurs.TEXTE,
            insertbackground=Couleurs.BLANC, relief="flat", bd=5,
            width=14 if est_date else 30,
        )
        saisie.pack(side="left", padx=2)

        if est_date:
            self._afficher_saisie_date(ligne, ru, saisie)
        elif ru["action"]["valeur"].get() == "":
            tk.Label(
                ligne, text="(sans valeur : règle non appliquée)",
                bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_TERTIAIRE, font=Polices.PETIT,
            ).pack(side="left", padx=6)
        elif not ru["action"]["colonne"].get().strip() and not ru["conditions"]:
            self._note_colonne_testee(ligne)

    def _note_colonne_testee(self, ligne) -> None:
        tk.Label(
            ligne, text="(colonne à préciser)", bg=Couleurs.FOND_CARTE,
            fg=Couleurs.TEXTE_TERTIAIRE, font=Polices.PETIT,
        ).pack(side="left", padx=6)

    def _premiere_colonne_condition(self, ru: dict) -> str:
        for cu in ru["conditions"]:
            if cu["colonne"].get().strip():
                return cu["colonne"].get().strip()
        return ""

    def _afficher_saisie_date(self, ligne, ru: dict, saisie) -> None:
        """Complète une saisie de date : bouton « aujourd'hui » et contrôle vif.

        Le contrôle est immédiat plutôt que différé à l'enregistrement : une
        date mal formée écrite dans un classeur y resterait sous forme de texte,
        invisible à l'œil mais inexploitable par Excel.
        """
        tk.Button(
            ligne, text="📅 aujourd'hui",
            command=lambda: self._dater_aujourdhui(ru),
            bg=Couleurs.FOND_PANNEAU2, fg=Couleurs.BLEU_CLAIR, font=Polices.PETIT,
            relief="flat", cursor="hand2", padx=6,
        ).pack(side="left", padx=4)

        controle = tk.Label(
            ligne, text="", bg=Couleurs.FOND_CARTE, font=Polices.PETIT, anchor="w",
        )
        controle.pack(side="left", padx=6)

        def rafraichir(_event=None) -> None:
            controle.config(**self._controle_date(ru["action"]["valeur"].get()))

        saisie.bind("<KeyRelease>", rafraichir)
        rafraichir()

    def _controle_date(self, texte: str) -> dict:
        """Message et couleur du contrôle de saisie d'une date."""
        if texte.strip() == "":
            return {
                "text": "(sans valeur : règle non appliquée)",
                "fg": Couleurs.TEXTE_TERTIAIRE,
            }
        moment = analyser_date(texte)
        if moment is None:
            return {"text": f"⚠ format attendu {GABARIT_DATE}", "fg": Couleurs.AVERTISSEMENT}
        return {"text": f"✓ {formater_date(moment)}", "fg": Couleurs.VERT}

    def _dater_aujourdhui(self, ru: dict) -> None:
        ru["action"]["valeur"].set(formater_date(datetime.now()))
        self._afficher()

    # ── Actions structurelles ─────────────────────────────────────
    def _nouvelle_condition_ui(self) -> dict:
        return self._condition_vers_ui(Condition("", "=", TypeDroite.VALEUR, ""))

    def _ajouter_regle(self) -> None:
        code = self._ui_vers_config().code_unique("R")
        self.regles_ui.append(self._regle_vers_ui(
            RegleModification(code=code, libelle="Nouvelle règle", action=Action())
        ))
        self._afficher()

    def _supprimer_regle(self, rang: int) -> None:
        if 0 <= rang < len(self.regles_ui):
            self.regles_ui.pop(rang)
            self._afficher()

    def _ajouter_condition(self, rang: int) -> None:
        self.regles_ui[rang]["conditions"].append(self._nouvelle_condition_ui())
        self._afficher()

    def _supprimer_condition(self, rang: int, index: int) -> None:
        self.regles_ui[rang]["conditions"].pop(index)
        self._afficher()

    def _reinitialiser(self) -> None:
        """Restaure la configuration livrée : les règles d'audit par défaut."""
        defaut = ReglesConfig.defaut()
        if not messagebox.askyesno(
            "Réinitialiser (config par défaut)",
            f"Remplacer la configuration actuelle par les {len(defaut.regles)} "
            "règles d'audit par défaut (colonnes, valeurs et opérateurs "
            "d'origine) ?\n\n"
            "Les règles saisies ici seront perdues. Cette action est sans effet "
            "tant qu'elle n'est pas enregistrée.",
            parent=self,
        ):
            return
        self.var_entete.set(str(defaut.ligne_entete))
        self.var_colonne_id.set(defaut.colonne_id)
        self.var_style.set(defaut.harmoniser_style)
        self.regles_ui = [self._regle_vers_ui(r) for r in defaut.regles]
        self._afficher()

    # ── Colonnes ──────────────────────────────────────────────────
    def _parcourir(self) -> None:
        chemin = filedialog.askopenfilename(
            title="Fichier Excel de référence", filetypes=[("Classeur Excel", "*.xlsx")],
            parent=self,
        )
        if chemin:
            self._charger_colonnes(chemin)

    def _charger_colonnes(self, chemin: str, silencieux: bool = False) -> None:
        try:
            colonnes = detecter_colonnes(chemin, ligne_entete=_entier(self.var_entete.get(), 1))
        except Exception as exc:  # noqa: BLE001 — fichier verrouillé, corrompu…
            if not silencieux:
                messagebox.showerror("Colonnes", str(exc), parent=self)
            return

        if not colonnes:
            if not silencieux:
                messagebox.showwarning(
                    "Colonnes", "Aucune colonne détectée dans ce fichier.", parent=self
                )
            return

        self.colonnes = colonnes
        self.noms = [c.nom for c in colonnes]
        self.info = {normaliser(c.nom): c for c in colonnes}
        self.info.update({c.lettre.casefold(): c for c in colonnes})
        self.fichier_colonnes = chemin
        self.combo_id.config(values=self.noms)
        self.lbl_colonnes.config(
            text=f"{len(colonnes)} colonnes détectées — {_raccourcir(chemin)}",
            fg=Couleurs.VERT,
        )
        self._afficher()

    def _colonne(self, reference: str):
        return self.info.get(normaliser(reference))

    def _type_de(self, reference: str) -> str:
        colonne = self._colonne(reference)
        return colonne.type if colonne else TYPE_TEXTE

    def _badge(self, reference: str) -> str:
        colonne = self._colonne(reference)
        if colonne is None:
            return ""
        return f"[{colonne.lettre} · {colonne.type}]"

    # ── Divers ────────────────────────────────────────────────────
    def _molette(self, event) -> None:
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _enregistrer(self) -> None:
        config = self._ui_vers_config()

        if not config.colonne_id.strip():
            messagebox.showerror(
                "Colonne des ID erreur",
                "Indiquez la colonne qui contient les ID erreur : c'est elle qui "
                "sélectionne les lignes à modifier.",
                parent=self,
            )
            return

        sans_colonne = [
            r.code for r in config.regles
            if r.actif and r.renseignee and not r.colonne_cible.strip()
        ]
        if sans_colonne:
            messagebox.showerror(
                "Colonne manquante",
                f"Les règles {', '.join(sans_colonne)} n'indiquent aucune colonne "
                "à modifier (ni action, ni condition).",
                parent=self,
            )
            return

        mauvaises_dates = [
            r.code for r in config.actives
            if r.action.source is SourceValeur.VALEUR
            and self._type_de(r.colonne_cible) == TYPE_DATE
            and analyser_date(r.action.valeur) is None
        ]
        if mauvaises_dates and not messagebox.askyesno(
            "Date non reconnue",
            f"Les règles {', '.join(mauvaises_dates)} visent une colonne de dates "
            f"mais leur valeur n'est pas une date ({GABARIT_DATE}).\n\n"
            "Elle sera écrite comme du texte. Enregistrer quand même ?",
            parent=self,
        ):
            return

        effacements = [
            r.code for r in config.actives if r.action.source is SourceValeur.VIDE
        ]
        if effacements and not messagebox.askyesno(
            "Effacement demandé",
            f"Les règles {', '.join(effacements)} videront le contenu de leur "
            "colonne sur les lignes traitées.\n\nEnregistrer quand même ?",
            parent=self,
        ):
            return

        self.on_save(config)
        self.destroy()

    def destroy(self):  # noqa: D102 — libère le binding global de la molette
        self.canvas.unbind_all("<MouseWheel>")
        super().destroy()


# ── Helpers module ────────────────────────────────────────────────
def _entier(valeur: str, defaut: int) -> int:
    try:
        return max(1, int(str(valeur).strip()))
    except (TypeError, ValueError):
        return defaut


def _raccourcir(chemin: str, longueur: int = 60) -> str:
    return chemin if len(chemin) <= longueur else "…" + chemin[-longueur:]
