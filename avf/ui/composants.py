"""Briques d'interface partagées : thème, cartes, boutons, dialogues de saisie.

Le thème (couleurs, polices) est celui de complete_excel_files
(``sna.ui.theme``) : les deux outils ont la même apparence, et le constructeur
de règles repris tel quel s'y intègre sans détonner.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from sna.ui.theme import Couleurs, Polices

from ..models.colonnes import ColonnesAudit

# Couleurs de statut (validée / non valide), lisibles sur le fond sombre ; elles
# accompagnent toujours une icône et un libellé, jamais seules.
VERT_STATUT = "#0ca30c"
AMBRE_STATUT = "#fab219"
ROUGE_STATUT = "#d03b3b"
PISTE = "#262f42"


def configurer_styles(racine: tk.Misc) -> None:
    """Styles ttk sombres (barre de progression, tableaux, listes déroulantes)."""
    style = ttk.Style(racine)
    style.theme_use("clam")
    style.configure(
        "avf.Horizontal.TProgressbar", troughcolor=Couleurs.FOND_CHAMP,
        background=Couleurs.BLEU, darkcolor=Couleurs.BLEU, lightcolor=Couleurs.BLEU,
        bordercolor=Couleurs.FOND_CHAMP,
    )
    style.configure(
        "avf.Treeview", background=Couleurs.FOND_JOURNAL, fieldbackground=Couleurs.FOND_JOURNAL,
        foreground=Couleurs.TEXTE, rowheight=24, bordercolor=Couleurs.BORDURE,
        font=Polices.LABEL,
    )
    style.map("avf.Treeview", background=[("selected", Couleurs.BLEU)],
              foreground=[("selected", Couleurs.BLANC)])
    style.configure(
        "avf.Treeview.Heading", background=Couleurs.FOND_PANNEAU2, foreground=Couleurs.TEXTE,
        font=Polices.LABEL_GRAS, relief="flat", bordercolor=Couleurs.BORDURE,
    )
    style.map("avf.Treeview.Heading", background=[("active", Couleurs.BORDURE)])
    style.configure(
        "avf.TCombobox", fieldbackground=Couleurs.FOND_JOURNAL, background=Couleurs.FOND_PANNEAU2,
        foreground=Couleurs.TEXTE, arrowcolor=Couleurs.TEXTE, bordercolor=Couleurs.BORDURE,
        padding=4,
    )
    # En lecture seule, clam repeint le champ en gris clair : on impose le fond sombre.
    style.map(
        "avf.TCombobox",
        fieldbackground=[("readonly", Couleurs.FOND_JOURNAL)],
        foreground=[("readonly", Couleurs.TEXTE)],
        selectbackground=[("readonly", Couleurs.FOND_JOURNAL)],
        selectforeground=[("readonly", Couleurs.TEXTE)],
        background=[("active", Couleurs.BORDURE), ("readonly", Couleurs.FOND_PANNEAU2)],
    )
    racine.option_add("*TCombobox*Listbox.background", Couleurs.FOND_JOURNAL)
    racine.option_add("*TCombobox*Listbox.foreground", Couleurs.TEXTE)
    racine.option_add("*TCombobox*Listbox.selectBackground", Couleurs.BLEU)


def carte(parent, titre: str = "", sous_titre: str = "") -> tk.Frame:
    """Cadre bordé, avec un titre facultatif ; renvoie le cadre intérieur."""
    cadre = tk.Frame(parent, bg=Couleurs.FOND_CARTE, highlightbackground=Couleurs.BORDURE,
                     highlightthickness=1)
    if titre:
        entete = tk.Frame(cadre, bg=Couleurs.FOND_CARTE)
        entete.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(entete, text=titre, bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE,
                 font=Polices.SOUS_TITRE).pack(side="left")
        if sous_titre:
            tk.Label(entete, text=sous_titre, bg=Couleurs.FOND_CARTE,
                     fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT).pack(side="left", padx=8)
    interieur = tk.Frame(cadre, bg=Couleurs.FOND_CARTE)
    interieur.pack(fill="both", expand=True, padx=12, pady=10)
    cadre.interieur = interieur  # type: ignore[attr-defined]
    return cadre


def bouton(parent, texte: str, commande, genre: str = "secondaire", **options) -> tk.Button:
    """Bouton plat aux couleurs du thème (``principal``, ``secondaire``, ``succes``, ``alerte``)."""
    fond, texte_couleur, police = {
        "principal": (Couleurs.BLEU, Couleurs.BLANC, Polices.BOUTON_PETIT),
        "succes": (Couleurs.VERT, "#062017", Polices.BOUTON_PETIT),
        "alerte": (Couleurs.FOND_PANNEAU2, Couleurs.AMBRE, Polices.BOUTON_PETIT),
        "secondaire": (Couleurs.FOND_PANNEAU2, Couleurs.BLEU_CLAIR, Polices.BOUTON_PETIT),
        "discret": (Couleurs.FOND_PANNEAU2, Couleurs.TEXTE_SECONDAIRE, Polices.PETIT),
    }[genre]
    parametres = dict(
        text=texte, command=commande, bg=fond, fg=texte_couleur, font=police, relief="flat",
        cursor="hand2", padx=12, pady=5, activebackground=Couleurs.BORDURE,
        activeforeground=Couleurs.BLANC, disabledforeground=Couleurs.TEXTE_TERTIAIRE, bd=0,
    )
    parametres.update(options)
    return tk.Button(parent, **parametres)


def case(parent, texte: str, variable: tk.BooleanVar, couleur: str = Couleurs.TEXTE_SECONDAIRE,
         **options) -> tk.Checkbutton:
    fond = options.pop("bg", Couleurs.FOND_CARTE)
    return tk.Checkbutton(
        parent, text=texte, variable=variable, bg=fond, fg=couleur, font=Polices.PETIT,
        activebackground=fond, activeforeground=couleur, selectcolor=Couleurs.FOND_PANNEAU2,
        cursor="hand2", anchor="w", **options,
    )


def champ(parent, variable: tk.StringVar, **options) -> tk.Entry:
    return tk.Entry(
        parent, textvariable=variable, font=Polices.LABEL, bg=Couleurs.FOND_JOURNAL,
        fg=Couleurs.TEXTE, insertbackground=Couleurs.BLANC, relief="flat", bd=6, **options,
    )


class ZoneDefilante(tk.Frame):
    """Cadre au contenu défilant verticalement (petits écrans).

    Placer les widgets dans :attr:`interieur`. Le contenu prend toute la largeur ;
    en hauteur, il occupe au moins la zone visible (un widget en ``expand=True``
    s'étire donc comme avant sur un grand écran) et défile quand il ne tient pas.
    La barre de défilement n'apparaît que si elle sert. La molette fait défiler
    la page, sauf au-dessus d'une zone qui défile elle-même (journal, listes,
    tableaux) tant que celle-ci peut encore défiler.
    """

    def __init__(self, parent, bg: str = Couleurs.FOND, **options):
        super().__init__(parent, bg=bg, **options)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.barre = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._defilement)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.interieur = tk.Frame(self.canvas, bg=bg)
        self._fenetre = self.canvas.create_window(0, 0, window=self.interieur, anchor="nw")
        self.interieur.bind("<Configure>", lambda e: self._ajuster())
        self.canvas.bind("<Configure>", lambda e: self._ajuster())
        self.bind_all("<MouseWheel>", self._molette, add="+")

    def _ajuster(self) -> None:
        largeur = self.canvas.winfo_width()
        hauteur = max(self.interieur.winfo_reqheight(), self.canvas.winfo_height())
        self.canvas.itemconfigure(self._fenetre, width=largeur, height=hauteur)
        self.canvas.configure(scrollregion=(0, 0, largeur, hauteur))

    def _defilement(self, debut: str, fin: str) -> None:
        self.barre.set(debut, fin)
        utile = (float(debut), float(fin)) != (0.0, 1.0)
        if utile and not self.barre.winfo_ismapped():
            self.barre.pack(side="right", fill="y", before=self.canvas)
        elif not utile and self.barre.winfo_ismapped():
            self.barre.pack_forget()

    def defile(self) -> bool:
        """Vrai si le contenu dépasse la zone visible."""
        return self.canvas.yview() != (0.0, 1.0)

    def _molette(self, event) -> str | None:
        widget = event.widget
        if not isinstance(widget, tk.Misc) or not str(widget).startswith(str(self)):
            return None                       # autre fenêtre (popup du rapport…)
        if not self.defile():
            return None
        if isinstance(widget, (tk.Text, tk.Listbox, ttk.Treeview)):
            debut, fin = widget.yview()
            vers_le_bas = event.delta < 0
            if (vers_le_bas and fin < 1.0) or (not vers_le_bas and debut > 0.0):
                return None                   # la zone défile elle-même
        self.canvas.yview_scroll(int(-event.delta / 120) or (-1 if event.delta > 0 else 1), "units")
        return "break"

    def voir(self, widget) -> None:
        """Fait défiler pour rendre ``widget`` visible."""
        self.update_idletasks()
        haut = widget.winfo_rooty() - self.interieur.winfo_rooty()
        total = max(self.interieur.winfo_height(), 1)
        self.canvas.yview_moveto(max(haut - 20, 0) / total)


class BarreSegmentee(tk.Canvas):
    """Barre fine validé / non valide (même lecture que dans le rapport HTML)."""

    def __init__(self, parent, hauteur: int = 10, **options):
        super().__init__(parent, height=hauteur, highlightthickness=0,
                         bg=options.pop("bg", Couleurs.FOND_CARTE), **options)
        self._valeurs = (0, 0)
        self.bind("<Configure>", lambda e: self._dessiner())

    def definir(self, renseignes: int, vides: int) -> None:
        self._valeurs = (renseignes, vides)
        self._dessiner()

    def _dessiner(self) -> None:
        self.delete("all")
        largeur, hauteur = max(self.winfo_width(), 1), int(self["height"])
        ok, vides = self._valeurs
        total = ok + vides
        if not total:
            self.create_rectangle(0, 0, largeur, hauteur, fill=PISTE, width=0)
            return
        coupe = round(largeur * ok / total)
        if ok:
            self.create_rectangle(0, 0, coupe - (1 if vides else 0), hauteur, fill=VERT_STATUT, width=0)
        if vides:
            self.create_rectangle(coupe + (1 if ok else 0), 0, largeur, hauteur, fill=AMBRE_STATUT, width=0)


# ── Dialogues ─────────────────────────────────────────────────────
class _Modal(tk.Toplevel):
    def __init__(self, parent, titre: str, largeur: int = 520, hauteur: int = 360):
        super().__init__(parent)
        self.title(titre)
        self.configure(bg=Couleurs.FOND)
        self.geometry(f"{largeur}x{hauteur}")
        self.transient(parent)
        self.resultat = None
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda e: self.destroy())

    def montrer(self):
        """Affiche en modal et retourne le résultat (``None`` si annulé)."""
        self.grab_set()
        self.wait_window()
        return self.resultat

    def _pied(self, valider) -> None:
        pied = tk.Frame(self, bg=Couleurs.FOND)
        pied.pack(fill="x", side="bottom", padx=16, pady=12)
        bouton(pied, "Valider", valider, "succes").pack(side="right")
        bouton(pied, "Annuler", self.destroy, "discret").pack(side="right", padx=8)


class DialogueSaisie(_Modal):
    """Saisie libre d'une liste de codes (départements ou communes)."""

    def __init__(self, parent, titre: str, consigne: str, exemple: str, texte: str = ""):
        super().__init__(parent, titre)
        tk.Label(self, text=consigne, bg=Couleurs.FOND, fg=Couleurs.TEXTE, font=Polices.CORPS,
                 justify="left", wraplength=480).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(self, text=f"Exemple : {exemple}", bg=Couleurs.FOND,
                 fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT).pack(anchor="w", padx=16)
        self._pied(self._valider)
        self.zone = tk.Text(self, bg=Couleurs.FOND_CHAMP, fg=Couleurs.BLANC, font=Polices.MONO,
                            relief="flat", insertbackground=Couleurs.BLANC, bd=6, height=8)
        self.zone.pack(fill="both", expand=True, padx=16, pady=(8, 0))
        self.zone.insert("1.0", texte)
        self.zone.focus_set()
        self.bind("<Control-Return>", lambda e: self._valider())

    def _valider(self) -> None:
        self.resultat = self.zone.get("1.0", tk.END)
        self.destroy()


class DialogueColonnes(_Modal):
    """Colonnes contrôlées, valeurs acceptées, colonne des ID erreur et ligne d'en-tête.

    La colonne des ID erreur et la ligne d'en-tête sont celles des règles de
    mise à jour : un seul réglage pour l'audit et la mise à jour.
    """

    def __init__(self, parent, colonnes: ColonnesAudit, colonne_id: str, ligne_entete: int):
        super().__init__(parent, "Colonnes auditées", 580, 500)
        tk.Label(
            self, text="Libellés recherchés dans l'onglet « Audit » (casse et espaces ignorés).",
            bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.LABEL,
        ).pack(anchor="w", padx=16, pady=(14, 8))
        grille = tk.Frame(self, bg=Couleurs.FOND)
        grille.pack(fill="x", padx=16)
        self.vars = {
            "colonne_id": tk.StringVar(value=colonne_id),
            "audit": tk.StringVar(value=colonnes.audit),
            "date_audit": tk.StringVar(value=colonnes.date_audit),
            "auditeur": tk.StringVar(value=colonnes.auditeur),
            "validation": tk.StringVar(value=colonnes.validation),
            "valeurs_audit": tk.StringVar(value=", ".join(colonnes.valeurs_audit)),
            "valeurs_validation": tk.StringVar(value=", ".join(colonnes.valeurs_validation)),
            "ligne_entete": tk.StringVar(value=str(ligne_entete)),
        }
        libelles = [
            ("colonne_id", "ID erreur"), ("audit", "Audit"), ("date_audit", "Date audit"),
            ("auditeur", "Auditeur"), ("validation", "Validation finale"), None,
            ("valeurs_audit", "Audit accepté"), ("valeurs_validation", "Validation finale acceptée"),
            None, ("ligne_entete", "Ligne d'en-tête"),
        ]
        for rang, element in enumerate(libelles):
            if element is None:
                tk.Frame(grille, bg=Couleurs.FOND, height=8).grid(row=rang, column=0)
                continue
            cle, libelle = element
            tk.Label(grille, text=libelle, bg=Couleurs.FOND, fg=Couleurs.TEXTE,
                     font=Polices.LABEL, width=24, anchor="w").grid(row=rang, column=0, pady=3)
            champ(grille, self.vars[cle], width=36).grid(row=rang, column=1, sticky="ew", pady=3)
        grille.columnconfigure(1, weight=1)
        tk.Label(
            self,
            text="Une ligne est validée si : Audit est l'une des valeurs acceptées, Date audit est "
                 "une date (même saisie en texte : 12/03/2026, 2026-03-12, 12 mars 2026…), Auditeur "
                 "est renseigné et Validation finale est l'une des valeurs acceptées. Valeurs "
                 "séparées par des virgules ; casse et espaces ignorés.\n\n"
                 "Les cinq colonnes sont obligatoires : une colonne absente fait échouer l'audit "
                 "de la commune. La colonne ID erreur et la ligne d'en-tête servent aussi à la mise à jour.",
            bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT, justify="left",
            wraplength=540,
        ).pack(anchor="w", padx=16, pady=(10, 0))
        self._pied(self._valider)

    def _valider(self) -> None:
        valeurs = {k: v.get().strip() for k, v in self.vars.items()}
        try:
            ligne = max(int(valeurs.pop("ligne_entete")), 1)
        except ValueError:
            ligne = 1
        colonne_id = valeurs.pop("colonne_id") or "ID erreur"
        self.resultat = (ColonnesAudit.from_dict(valeurs), colonne_id, ligne)
        self.destroy()
