"""Fenêtre « Rapport d'audit » : le rapport HTML, sans quitter l'outil.

    - indicateurs (communes, ID imm, lignes validées, lignes non valides),
      barre globale et motifs de non-validité ;
    - tableau des communes, triable, filtrable par texte et par état ;
    - pour la commune sélectionnée : ses lignes auditées (ID imm, ID erreur, ligne,
      statut, Audit, Date audit, Auditeur, Validation finale, motifs) ;
    - comme dans audit_ppts, le détail ne montre que les lignes non valides ;
      la case « Afficher aussi les validées » y ajoute les autres ;
    - copie en un clic des ID non valides (commune sélectionnée, ou communes
      affichées) ou des lignes sélectionnées, avec le séparateur choisi ;
    - export HTML, Excel ou CSV, et lancement de la mise à jour (commune
      sélectionnée ou toutes) via les callbacks fournis par la fenêtre principale.

La fenêtre ne fait aucun traitement : elle affiche un :class:`RapportAudit`, et
:meth:`rafraichir` la remet à jour après une mise à jour suivie d'un ré-audit.
"""

from __future__ import annotations

from collections.abc import Callable
import tkinter as tk
from tkinter import ttk

from sna.services.identifiants import normaliser_id
from sna.ui.theme import Couleurs, Polices

from ..models.audit import NON_VALIDES, TOUTES, AuditCommune, RapportAudit
from ..views import presentation as p
from .composants import AMBRE_STATUT, VERT_STATUT, BarreSegmentee, bouton, case, champ

_COLONNES = [
    ("dep", "Dép.", 56, "w"),
    ("commune", "Commune", 88, "w"),
    ("ids", "ID", 48, "e"),
    ("lignes", "Lignes", 58, "e"),
    ("ok", "Validées", 70, "e"),
    ("nv", "Non valides", 98, "e"),
    ("taux", "Taux", 60, "e"),
    ("etat", "État", 146, "w"),
]
_DETAIL = [
    ("imm", "ID imm", 118, "w"),
    ("id", "ID erreur", 92, "w"),
    ("ligne", "Ligne", 46, "e"),
    ("statut", "Statut", 112, "w"),
    ("audit", "Audit", 52, "w"),
    ("date", "Date audit", 86, "w"),
    ("auditeur", "Auditeur", 92, "w"),
    ("vf", "Validation finale", 128, "w"),
    ("motifs", "Motifs", 300, "w"),
]
_FILTRES = {
    "Toutes": None,
    "⚠️ À corriger": "a-corriger",
    "✅ Complètes": "complete",
    "❌ En échec": "echec",
}
_SEPARATEUR_IID = "\x1f"   # « ID imm » + séparateur + n° de ligne : iid unique d'une ligne


class FenetreRapport(tk.Toplevel):
    def __init__(
        self,
        parent,
        rapport: RapportAudit,
        on_exporter: Callable[[str], None] | None = None,
        on_mettre_a_jour: Callable[[list[str] | None], None] | None = None,
    ):
        super().__init__(parent)
        self.title("📊 Rapport d'audit — Validation finale")
        self.geometry("1500x820")
        self.minsize(1120, 620)
        self.configure(bg=Couleurs.FOND)
        self.rapport = rapport
        self.on_exporter = on_exporter
        self.on_mettre_a_jour = on_mettre_a_jour

        self.var_recherche = tk.StringVar()
        self.var_filtre = tk.StringVar(value="Toutes")
        self.var_tous = tk.BooleanVar(value=False)   # « Afficher aussi les validées »
        self.var_separateur = tk.StringVar(value="ligne")
        self._tri = ("commune", False)
        self.var_recherche.trace_add("write", lambda *_: self._remplir())

        self._construire()
        self.rafraichir(rapport)

    # ── Construction ──────────────────────────────────────────────
    def _construire(self) -> None:
        entete = tk.Frame(self, bg=Couleurs.FOND_HEADER)
        entete.pack(fill="x")
        tk.Label(entete, text="  📊  Rapport d'audit — Validation finale", font=Polices.TITRE,
                 bg=Couleurs.FOND_HEADER, fg=Couleurs.TEXTE).pack(side="left", pady=10)
        self.lbl_meta = tk.Label(entete, text="", bg=Couleurs.FOND_HEADER,
                                 fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT, justify="right")
        self.lbl_meta.pack(side="right", padx=16)

        self._construire_pied()
        corps = tk.Frame(self, bg=Couleurs.FOND)
        corps.pack(fill="both", expand=True, padx=16, pady=12)

        # Indicateurs
        tuiles = tk.Frame(corps, bg=Couleurs.FOND)
        tuiles.pack(fill="x")
        self.tuiles = {}
        for rang, (cle, libelle, accent) in enumerate([
            ("communes", "Communes auditées", Couleurs.BORDURE),
            ("ids", "ID imm", Couleurs.BORDURE),
            ("valides", "✅ Lignes validées", VERT_STATUT),
            ("non_valides", "⚠️ Lignes non valides", AMBRE_STATUT),
        ]):
            cadre = tk.Frame(tuiles, bg=Couleurs.FOND_CARTE, highlightbackground=Couleurs.BORDURE,
                             highlightthickness=1)
            cadre.grid(row=0, column=rang, sticky="nsew", padx=(0 if rang == 0 else 8, 0))
            tuiles.columnconfigure(rang, weight=1)
            tk.Frame(cadre, bg=accent, width=4).pack(side="left", fill="y")
            dedans = tk.Frame(cadre, bg=Couleurs.FOND_CARTE)
            dedans.pack(side="left", fill="both", expand=True, padx=12, pady=8)
            tk.Label(dedans, text=libelle, bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_SECONDAIRE,
                     font=Polices.LABEL).pack(anchor="w")
            valeur = tk.Label(dedans, text="", bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE,
                              font=("Segoe UI", 20, "bold"))
            valeur.pack(anchor="w")
            detail = tk.Label(dedans, text="", bg=Couleurs.FOND_CARTE,
                              fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT)
            detail.pack(anchor="w")
            self.tuiles[cle] = (valeur, detail)

        self.barre = BarreSegmentee(corps, hauteur=10, bg=Couleurs.FOND)
        self.barre.pack(fill="x", pady=(10, 2))
        legende = tk.Frame(corps, bg=Couleurs.FOND)
        legende.pack(fill="x", pady=(0, 2))
        for couleur, texte in ((VERT_STATUT, "✅ validées"), (AMBRE_STATUT, "⚠️ non valides")):
            tk.Label(legende, text="●", fg=couleur, bg=Couleurs.FOND, font=Polices.LABEL).pack(side="left")
            tk.Label(legende, text=texte, fg=Couleurs.TEXTE_SECONDAIRE, bg=Couleurs.FOND,
                     font=Polices.PETIT).pack(side="left", padx=(2, 14))
        self.lbl_motifs = tk.Label(corps, text="", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                                   font=Polices.PETIT, anchor="w", justify="left")
        self.lbl_motifs.pack(fill="x", pady=(0, 8))

        # Barre d'outils
        outils = tk.Frame(corps, bg=Couleurs.FOND)
        outils.pack(fill="x", pady=(0, 6))
        tk.Label(outils, text="🔎", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                 font=Polices.LABEL).pack(side="left")
        champ(outils, self.var_recherche, width=24).pack(side="left", padx=(4, 10))
        filtre = ttk.Combobox(outils, textvariable=self.var_filtre, values=list(_FILTRES),
                              state="readonly", width=14, style="avf.TCombobox")
        filtre.pack(side="left")
        filtre.bind("<<ComboboxSelected>>", lambda e: self._remplir())
        tk.Label(outils, text="Séparateur :", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                 font=Polices.PETIT).pack(side="left", padx=(16, 4))
        ttk.Combobox(outils, textvariable=self.var_separateur, values=list(p.SEPARATEURS_COPIE),
                     state="readonly", width=12, style="avf.TCombobox").pack(side="left")
        self.btn_copier_affiches = bouton(outils, "📋 Copier les ID non valides affichés",
                                          self._copier_affiches, "principal")
        self.btn_copier_affiches.pack(side="right")

        # Zone principale : tableau des communes | détail
        panneaux = tk.PanedWindow(corps, orient="horizontal", bg=Couleurs.FOND, sashwidth=6,
                                  bd=0, sashrelief="flat")
        panneaux.pack(fill="both", expand=True)

        gauche = tk.Frame(panneaux, bg=Couleurs.FOND)
        self.arbre = ttk.Treeview(gauche, columns=[c[0] for c in _COLONNES], show="headings",
                                  style="avf.Treeview", selectmode="browse")
        for cle, libelle, largeur, ancre in _COLONNES:
            self.arbre.heading(cle, text=libelle, command=lambda c=cle: self._trier(c))
            self.arbre.column(cle, width=largeur, anchor=ancre, stretch=cle == "etat")
        self.arbre.tag_configure("a-corriger", foreground=AMBRE_STATUT)
        self.arbre.tag_configure("complete", foreground=Couleurs.SUCCES)
        self.arbre.tag_configure("echec", foreground=Couleurs.ROUGE)
        self.arbre.tag_configure("vide", foreground=Couleurs.TEXTE_SECONDAIRE)
        defil = ttk.Scrollbar(gauche, orient="vertical", command=self.arbre.yview)
        self.arbre.configure(yscrollcommand=defil.set)
        defil.pack(side="right", fill="y")
        self.arbre.pack(fill="both", expand=True)
        self.arbre.bind("<<TreeviewSelect>>", lambda e: self._afficher_detail())
        self.lbl_aucun = tk.Label(gauche, text="", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                                  font=Polices.PETIT)
        self.lbl_aucun.pack(anchor="w")
        panneaux.add(gauche, minsize=600, width=670)

        droite = tk.Frame(panneaux, bg=Couleurs.FOND_CARTE, highlightbackground=Couleurs.BORDURE,
                          highlightthickness=1)
        self.lbl_commune = tk.Label(droite, text="Sélectionnez une commune", bg=Couleurs.FOND_CARTE,
                                    fg=Couleurs.TEXTE, font=Polices.SOUS_TITRE, anchor="w")
        self.lbl_commune.pack(fill="x", padx=12, pady=(10, 0))
        self.lbl_messages = tk.Label(droite, text="", bg=Couleurs.FOND_CARTE,
                                     fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT,
                                     justify="left", anchor="w", wraplength=560)
        self.lbl_messages.pack(fill="x", padx=12, pady=(2, 6))
        actions = tk.Frame(droite, bg=Couleurs.FOND_CARTE)
        actions.pack(fill="x", padx=12)
        self.btn_copier_commune = bouton(actions, "📋 Copier les ID",
                                         self._copier_commune, "principal", state="disabled")
        self.btn_copier_commune.pack(side="left")
        self.btn_maj_commune = bouton(actions, "🛠️ Mettre à jour la commune",
                                      self._maj_commune, "alerte", state="disabled")
        self.btn_maj_commune.pack(side="left", padx=6)
        self.case_tous = case(actions, "Afficher aussi les validées", self.var_tous,
                              command=self._afficher_detail)
        self.case_tous.pack(side="right")

        cadre_detail = tk.Frame(droite, bg=Couleurs.FOND_CARTE)
        cadre_detail.pack(fill="both", expand=True, padx=12, pady=8)
        self.detail = ttk.Treeview(cadre_detail, columns=[c[0] for c in _DETAIL],
                                   show="headings", style="avf.Treeview", selectmode="extended")
        for cle, libelle, largeur, ancre in _DETAIL:
            self.detail.heading(cle, text=libelle)
            self.detail.column(cle, width=largeur, minwidth=40, anchor=ancre, stretch=cle == "motifs")
        self.detail.tag_configure("validee", foreground=Couleurs.SUCCES)
        self.detail.tag_configure("non-valide", foreground=AMBRE_STATUT)
        defil_x = ttk.Scrollbar(cadre_detail, orient="horizontal", command=self.detail.xview)
        defil_y = ttk.Scrollbar(cadre_detail, orient="vertical", command=self.detail.yview)
        self.detail.configure(xscrollcommand=defil_x.set, yscrollcommand=defil_y.set)
        defil_x.pack(side="bottom", fill="x")
        defil_y.pack(side="right", fill="y")
        self.detail.pack(fill="both", expand=True)
        self.detail.bind("<Control-c>", lambda e: self._copier_selection_detail())
        self.detail.bind("<Double-1>", lambda e: self._copier_selection_detail())
        tk.Label(droite, text="Double-clic ou Ctrl+C : copier les ID des lignes sélectionnées",
                 bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_TERTIAIRE,
                 font=Polices.PETIT).pack(anchor="w", padx=12, pady=(0, 8))
        panneaux.add(droite, minsize=480, stretch="always")

    def _construire_pied(self) -> None:
        pied = tk.Frame(self, bg=Couleurs.FOND_PANNEAU)
        pied.pack(fill="x", side="bottom")
        self.lbl_statut = tk.Label(pied, text="", bg=Couleurs.FOND_PANNEAU, fg=Couleurs.SUCCES,
                                   font=Polices.LABEL)
        self.lbl_statut.pack(side="left", padx=16)
        bouton(pied, "Fermer", self.destroy, "discret").pack(side="right", padx=16, pady=10)
        self.btn_maj_tout = bouton(pied, "🛠️ Mettre à jour tous les ID non valides", self._maj_tout, "alerte")
        self.btn_maj_tout.pack(side="right", pady=10)
        if self.on_exporter:
            for texte, format in (("📝 ID non audités", "txt"), ("📑 CSV", "csv"), ("📊 Excel", "xlsx"),
                                  ("💾 HTML", "html")):
                bouton(pied, texte, lambda f=format: self.on_exporter(f), "secondaire").pack(
                    side="right", padx=(0, 8), pady=10)
            tk.Label(pied, text="Exporter :", bg=Couleurs.FOND_PANNEAU, fg=Couleurs.TEXTE_SECONDAIRE,
                     font=Polices.PETIT).pack(side="right", padx=(0, 6))

    # ── Données ───────────────────────────────────────────────────
    def rafraichir(self, rapport: RapportAudit) -> None:
        """Réaffiche un rapport (nouvel audit, ou ré-audit après mise à jour)."""
        self.rapport = rapport
        r = rapport
        taux_nv = r.nb_lignes_non_valides / r.nb_lignes if r.nb_lignes else None
        valeurs = {
            "communes": (p.nombre(r.nb_communes),
                         f"{len(r.communes_completes)} complète(s) · {len(r.communes_a_corriger)} à "
                         f"corriger · {len(r.en_echec)} en échec"),
            "ids": (p.nombre(r.nb_ids), f"{p.nombre(r.nb_ids_non_valides)} non valide(s) · "
                                        f"{p.nombre(r.nb_ids_valides)} validé(s)"),
            "valides": (p.nombre(r.nb_lignes_valides), f"{p.pourcent(r.taux)} des lignes"),
            "non_valides": (p.nombre(r.nb_lignes_non_valides), f"{p.pourcent(taux_nv)} des lignes"),
        }
        for cle, (valeur, detail) in valeurs.items():
            self.tuiles[cle][0].config(text=valeur)
            self.tuiles[cle][1].config(text=detail)
        self.barre.definir(r.nb_lignes_valides, r.nb_lignes_non_valides)
        motifs = [f"{libelle} : {p.nombre(n)}" for libelle, n in
                  zip(p.libelles_criteres(r.colonnes).values(), r.par_critere().values())]
        self.lbl_motifs.config(text="Lignes non valides par critère non respecté —  " + "   ·   ".join(motifs))
        racine = r.racine if len(r.racine) <= 70 else "…" + r.racine[-69:]
        cible = f" · 🎯 {p.nombre(r.nb_ids_demandes)} ID saisis" if r.par_ids else ""
        self.lbl_meta.config(
            text=f"{racine}\n{r.libelle_source}{cible} · {r.date:%d/%m/%Y %H:%M} · {p.duree(r.duree)}")
        etat_maj = "normal" if (self.on_mettre_a_jour and r.nb_ids_non_valides) else "disabled"
        self.btn_maj_tout.config(state=etat_maj,
                                 text=f"🛠️ Mettre à jour les {p.nombre(r.nb_ids_non_valides)} ID non valides")
        selection = self._commune_selectionnee()
        self._remplir()
        if selection and self.arbre.exists(selection.commune):
            self.arbre.selection_set(selection.commune)
            self.arbre.see(selection.commune)
        self._afficher_detail()

    def categorie(self) -> str:
        """Lignes du détail : non valides, ou toutes si « Afficher aussi les validées » est coché."""
        return TOUTES if self.var_tous.get() else NON_VALIDES

    def _communes_filtrees(self) -> list[AuditCommune]:
        recherche = self.var_recherche.get().strip().lower()
        etat_voulu = _FILTRES.get(self.var_filtre.get())
        resultat = []
        for c in self.rapport.communes:
            if etat_voulu and p.etat(c).cle != etat_voulu:
                continue
            if recherche and recherche not in " ".join(
                [c.commune, f"dep{c.departement}", *(x for l in c.lignes for x in (l.id_imm, l.id_erreur))]
            ).lower():
                continue
            resultat.append(c)
        cle, inverse = self._tri
        return sorted(resultat, key=lambda c: _cle_tri(c, cle), reverse=inverse)

    def _remplir(self) -> None:
        selection = self.arbre.selection()
        self.arbre.delete(*self.arbre.get_children())
        communes = self._communes_filtrees()
        for c in communes:
            e = p.etat(c)
            self.arbre.insert("", "end", iid=c.commune, tags=(e.cle,), values=(
                f"Dep{c.departement}", c.commune, p.nombre(c.nb_ids), p.nombre(c.nb_lignes),
                p.nombre(c.nb_lignes_valides), p.nombre(c.nb_lignes_non_valides),
                p.pourcent(c.taux) if c.lisible else "—",
                f"{e.icone} {e.libelle}" + ("  ⚠" if c.avertissements else ""),
            ))
        # Filtre, tri ou « Afficher » : la commune sélectionnée le reste si elle est encore listée.
        gardees = [iid for iid in selection if self.arbre.exists(iid)]
        if gardees:
            self.arbre.selection_set(gardees)
        nb_ids = sum(c.nb_ids_non_valides for c in communes if c.lisible)
        self.lbl_aucun.config(
            text="Aucune commune ne correspond." if not communes
            else f"{len(communes)} commune(s) affichée(s) · {p.nombre(nb_ids)} ID imm non valide(s)"
        )
        self.btn_copier_affiches.config(
            text=f"📋 Copier les ID non valides affichés ({p.nombre(nb_ids)})",
            state="normal" if nb_ids else "disabled",
        )

    def _trier(self, cle: str) -> None:
        actuelle, inverse = self._tri
        self._tri = (cle, not inverse if actuelle == cle else False)
        for c, libelle, *_ in _COLONNES:
            fleche = (" ▼" if self._tri[1] else " ▲") if c == cle else ""
            self.arbre.heading(c, text=libelle + fleche)
        self._remplir()

    def _commune_selectionnee(self) -> AuditCommune | None:
        selection = self.arbre.selection() if hasattr(self, "arbre") else ()
        return self.rapport.commune(selection[0]) if selection else None

    def _afficher_detail(self) -> None:
        self.detail.delete(*self.detail.get_children())
        c = self._commune_selectionnee()
        if c is None:
            self.lbl_commune.config(text="Sélectionnez une commune")
            self.lbl_messages.config(text="")
            self.btn_copier_commune.config(state="disabled")
            self.btn_maj_commune.config(state="disabled")
            return
        e = p.etat(c)
        self.lbl_commune.config(text=f"{e.icone}  {c.commune} — Dep{c.departement} · {e.libelle}")
        lignes = []
        if c.chemin:
            lignes.append(f"📄 {c.chemin}")
        if not c.lisible:
            lignes.append(f"❌ {c.message}")
        lignes.extend(f"⚠ {a}" for a in c.avertissements)
        if c.lisible:
            lignes.append(
                f"{p.nombre(c.nb_lignes)} ligne(s) · {p.nombre(c.nb_lignes_valides)} validée(s) · "
                f"{p.nombre(c.nb_lignes_non_valides)} non valide(s) — {p.nombre(c.nb_ids)} ID imm, "
                f"dont {p.nombre(c.nb_ids_non_valides)} non valide(s)"
            )
        self.lbl_messages.config(text="\n".join(lignes))
        for l in c.lignes_de(self.categorie()):
            el = p.etat_ligne(l)
            self.detail.insert("", "end", iid=f"{normaliser_id(l.id_imm)}{_SEPARATEUR_IID}{l.ligne}",
                               tags=(el.cle,),
                               values=(l.id_imm, l.id_erreur, l.ligne, f"{el.icone} {el.libelle}", l.audit,
                                       l.date_audit, l.auditeur, l.validation, l.motifs))
        ids = c.ids_non_valides
        self.btn_copier_commune.config(state="normal" if ids else "disabled",
                                       text=f"📋 Copier les ID non valides ({p.nombre(len(ids))})")
        actif = c.lisible and c.nb_ids_non_valides and self.on_mettre_a_jour
        self.btn_maj_commune.config(state="normal" if actif else "disabled")

    # ── Copie ─────────────────────────────────────────────────────
    def _copier(self, ids: list[str]) -> None:
        if not ids:
            return
        separateur = p.SEPARATEURS_COPIE.get(self.var_separateur.get(), "\n")
        self.clipboard_clear()
        self.clipboard_append(p.texte_ids(ids, separateur))
        self.update_idletasks()
        self._statut(f"✓ {p.nombre(len(ids))} ID imm copié(s) dans le presse-papiers")

    def _copier_commune(self) -> None:
        c = self._commune_selectionnee()
        if c is not None:
            self._copier(c.ids_non_valides)

    def _copier_affiches(self) -> None:
        self._copier([i for c in self._communes_filtrees() for i in c.ids_non_valides])

    def _copier_selection_detail(self) -> None:
        ids = [iid.split(_SEPARATEUR_IID)[0] for iid in self.detail.selection()]
        self._copier([i for i in dict.fromkeys(ids) if i])   # ligne sans ID imm : rien à copier

    def _statut(self, texte: str) -> None:
        self.lbl_statut.config(text=texte)
        self.after(3000, lambda: self.lbl_statut.winfo_exists() and self.lbl_statut.config(text=""))

    # ── Mise à jour ───────────────────────────────────────────────
    def _maj_commune(self) -> None:
        c = self._commune_selectionnee()
        if c is not None and self.on_mettre_a_jour:
            self.on_mettre_a_jour([c.commune])

    def _maj_tout(self) -> None:
        if self.on_mettre_a_jour:
            self.on_mettre_a_jour(None)


def _cle_tri(c: AuditCommune, cle: str):
    return {
        "dep": (c.departement, c.commune),
        "commune": c.commune,
        "ids": c.nb_ids,
        "lignes": c.nb_lignes,
        "ok": c.nb_lignes_valides,
        "nv": c.nb_lignes_non_valides,
        "taux": -1 if c.taux is None else c.taux,
        "etat": p.etat(c).cle,
    }[cle]
