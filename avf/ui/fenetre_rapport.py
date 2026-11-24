"""Fenêtre « Rapport d'audit » : le rapport HTML, sans quitter l'outil.

    - indicateurs (communes, ID erreur, renseignés, vides) et barre globale ;
    - tableau des communes, triable, filtrable par texte et par état ;
    - pour la commune sélectionnée : ses ID sans Validation finale, avec
      Audit / Date audit / Auditeur et numéro de ligne ;
    - copie en un clic (commune sélectionnée, ou tout ce qui est affiché),
      avec le séparateur choisi ;
    - export HTML, et lancement de la mise à jour (commune sélectionnée ou
      toutes) via les callbacks fournis par la fenêtre principale.

La fenêtre ne fait aucun traitement : elle affiche un :class:`RapportAudit`, et
:meth:`rafraichir` la remet à jour après une mise à jour suivie d'un ré-audit.
"""

from __future__ import annotations

from collections.abc import Callable
import tkinter as tk
from tkinter import ttk

from sna.ui.theme import Couleurs, Polices

from ..models.audit import AuditCommune, RapportAudit
from ..views import presentation as p
from .composants import AMBRE_STATUT, VERT_STATUT, BarreSegmentee, bouton, champ

_COLONNES = [
    ("dep", "Dép.", 62, "w"),
    ("commune", "Commune", 84, "w"),
    ("ids", "ID erreur", 74, "e"),
    ("ok", "Renseignés", 86, "e"),
    ("vides", "Vides", 56, "e"),
    ("taux", "Taux", 64, "e"),
    ("etat", "État", 150, "w"),
]
_FILTRES = {
    "Toutes": None,
    "⏳ À compléter": "a-completer",
    "✅ Complètes": "complete",
    "❌ En échec": "echec",
}


class FenetreRapport(tk.Toplevel):
    def __init__(
        self,
        parent,
        rapport: RapportAudit,
        on_exporter: Callable[[], None] | None = None,
        on_mettre_a_jour: Callable[[list[str] | None], None] | None = None,
    ):
        super().__init__(parent)
        self.title("📊 Rapport d'audit — Validation finale")
        self.geometry("1180x760")
        self.minsize(960, 600)
        self.configure(bg=Couleurs.FOND)
        self.rapport = rapport
        self.on_exporter = on_exporter
        self.on_mettre_a_jour = on_mettre_a_jour

        self.var_recherche = tk.StringVar()
        self.var_filtre = tk.StringVar(value="Toutes")
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
            ("ids", "ID erreur", Couleurs.BORDURE),
            ("ok", "✅ Validation finale renseignée", VERT_STATUT),
            ("vides", "⏳ Validation finale vide", AMBRE_STATUT),
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
        legende.pack(fill="x", pady=(0, 8))
        for couleur, texte in ((VERT_STATUT, "✅ renseignée"), (AMBRE_STATUT, "⏳ vide")):
            tk.Label(legende, text="●", fg=couleur, bg=Couleurs.FOND, font=Polices.LABEL).pack(side="left")
            tk.Label(legende, text=texte, fg=Couleurs.TEXTE_SECONDAIRE, bg=Couleurs.FOND,
                     font=Polices.PETIT).pack(side="left", padx=(2, 14))

        # Barre d'outils
        outils = tk.Frame(corps, bg=Couleurs.FOND)
        outils.pack(fill="x", pady=(0, 6))
        tk.Label(outils, text="🔎", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                 font=Polices.LABEL).pack(side="left")
        champ(outils, self.var_recherche, width=28).pack(side="left", padx=(4, 10))
        filtre = ttk.Combobox(outils, textvariable=self.var_filtre, values=list(_FILTRES),
                              state="readonly", width=16, style="avf.TCombobox")
        filtre.pack(side="left")
        filtre.bind("<<ComboboxSelected>>", lambda e: self._remplir())
        tk.Label(outils, text="Séparateur :", bg=Couleurs.FOND, fg=Couleurs.TEXTE_SECONDAIRE,
                 font=Polices.PETIT).pack(side="left", padx=(16, 4))
        ttk.Combobox(outils, textvariable=self.var_separateur, values=list(p.SEPARATEURS_COPIE),
                     state="readonly", width=13, style="avf.TCombobox").pack(side="left")
        self.btn_copier_affiches = bouton(outils, "📋 Copier les ID vides affichés",
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
        self.arbre.tag_configure("a-completer", foreground=AMBRE_STATUT)
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
        panneaux.add(gauche, minsize=520, stretch="always")

        droite = tk.Frame(panneaux, bg=Couleurs.FOND_CARTE, highlightbackground=Couleurs.BORDURE,
                          highlightthickness=1)
        self.lbl_commune = tk.Label(droite, text="Sélectionnez une commune", bg=Couleurs.FOND_CARTE,
                                    fg=Couleurs.TEXTE, font=Polices.SOUS_TITRE, anchor="w")
        self.lbl_commune.pack(fill="x", padx=12, pady=(10, 0))
        self.lbl_messages = tk.Label(droite, text="", bg=Couleurs.FOND_CARTE,
                                     fg=Couleurs.TEXTE_SECONDAIRE, font=Polices.PETIT,
                                     justify="left", anchor="w", wraplength=420)
        self.lbl_messages.pack(fill="x", padx=12, pady=(2, 6))
        actions = tk.Frame(droite, bg=Couleurs.FOND_CARTE)
        actions.pack(fill="x", padx=12)
        self.btn_copier_commune = bouton(actions, "📋 Copier les ID",
                                         self._copier_commune, "principal", state="disabled")
        self.btn_copier_commune.pack(side="left")
        self.btn_maj_commune = bouton(actions, "🛠️ Mettre à jour la commune",
                                      self._maj_commune, "alerte", state="disabled")
        self.btn_maj_commune.pack(side="left", padx=6)

        self.detail = ttk.Treeview(droite, columns=("id", "lignes", "audit", "date", "auditeur"),
                                   show="headings", style="avf.Treeview", selectmode="extended")
        for cle, libelle, largeur in (("id", "ID erreur", 110), ("lignes", "Ligne(s)", 60),
                                      ("audit", "Audit", 80), ("date", "Date audit", 90),
                                      ("auditeur", "Auditeur", 110)):
            self.detail.heading(cle, text=libelle)
            self.detail.column(cle, width=largeur, anchor="w")
        self.detail.pack(fill="both", expand=True, padx=12, pady=8)
        self.detail.bind("<Control-c>", lambda e: self._copier_selection_detail())
        self.detail.bind("<Double-1>", lambda e: self._copier_selection_detail())
        tk.Label(droite, text="Double-clic ou Ctrl+C : copier les ID sélectionnés",
                 bg=Couleurs.FOND_CARTE, fg=Couleurs.TEXTE_TERTIAIRE,
                 font=Polices.PETIT).pack(anchor="w", padx=12, pady=(0, 8))
        panneaux.add(droite, minsize=400)

    def _construire_pied(self) -> None:
        pied = tk.Frame(self, bg=Couleurs.FOND_PANNEAU)
        pied.pack(fill="x", side="bottom")
        self.lbl_statut = tk.Label(pied, text="", bg=Couleurs.FOND_PANNEAU, fg=Couleurs.SUCCES,
                                   font=Polices.LABEL)
        self.lbl_statut.pack(side="left", padx=16)
        bouton(pied, "Fermer", self.destroy, "discret").pack(side="right", padx=16, pady=10)
        self.btn_maj_tout = bouton(pied, "🛠️ Mettre à jour tous les ID vides", self._maj_tout, "alerte")
        self.btn_maj_tout.pack(side="right", pady=10)
        if self.on_exporter:
            bouton(pied, "💾 Exporter en HTML", self.on_exporter, "secondaire").pack(
                side="right", padx=8, pady=10)

    # ── Données ───────────────────────────────────────────────────
    def rafraichir(self, rapport: RapportAudit) -> None:
        """Réaffiche un rapport (nouvel audit, ou ré-audit après mise à jour)."""
        self.rapport = rapport
        r = rapport
        taux_vide = r.nb_vides / r.nb_ids if r.nb_ids else None
        valeurs = {
            "communes": (p.nombre(r.nb_communes),
                         f"{len(r.communes_completes)} complète(s) · {len(r.communes_a_traiter)} à "
                         f"compléter · {len(r.en_echec)} en échec"),
            "ids": (p.nombre(r.nb_ids), f"sur {p.nombre(len(r.lues))} fichier(s) lu(s)"),
            "ok": (p.nombre(r.nb_renseignes), f"{p.pourcent(r.taux)} des ID erreur"),
            "vides": (p.nombre(r.nb_vides), f"{p.pourcent(taux_vide)} des ID erreur"),
        }
        for cle, (valeur, detail) in valeurs.items():
            self.tuiles[cle][0].config(text=valeur)
            self.tuiles[cle][1].config(text=detail)
        self.barre.definir(r.nb_renseignes, r.nb_vides)
        racine = r.racine if len(r.racine) <= 70 else "…" + r.racine[-69:]
        self.lbl_meta.config(text=f"{racine}\n{r.date:%d/%m/%Y %H:%M} · {p.duree(r.duree)}")
        etat_maj = "normal" if (self.on_mettre_a_jour and r.nb_vides) else "disabled"
        self.btn_maj_tout.config(state=etat_maj)
        selection = self._commune_selectionnee()
        self._remplir()
        if selection and self.arbre.exists(selection.commune):
            self.arbre.selection_set(selection.commune)
            self.arbre.see(selection.commune)
        self._afficher_detail()

    def _communes_filtrees(self) -> list[AuditCommune]:
        recherche = self.var_recherche.get().strip().lower()
        etat_voulu = _FILTRES.get(self.var_filtre.get())
        resultat = []
        for c in self.rapport.communes:
            if etat_voulu and p.etat(c).cle != etat_voulu:
                continue
            if recherche and recherche not in " ".join(
                [c.commune, f"dep{c.departement}", *c.ids_vides]
            ).lower():
                continue
            resultat.append(c)
        cle, inverse = self._tri
        return sorted(resultat, key=lambda c: _cle_tri(c, cle), reverse=inverse)

    def _remplir(self) -> None:
        self.arbre.delete(*self.arbre.get_children())
        communes = self._communes_filtrees()
        for c in communes:
            e = p.etat(c)
            self.arbre.insert("", "end", iid=c.commune, tags=(e.cle,), values=(
                f"Dep{c.departement}", c.commune, p.nombre(c.nb_ids), p.nombre(c.nb_renseignes),
                p.nombre(c.nb_vides), p.pourcent(c.taux) if c.lisible else "—",
                f"{e.icone} {e.libelle}" + ("  ⚠" if c.avertissements else ""),
            ))
        nb_ids = sum(len(c.ids_vides) for c in communes if c.lisible)
        self.lbl_aucun.config(
            text="Aucune commune ne correspond." if not communes
            else f"{len(communes)} commune(s) affichée(s) · {p.nombre(nb_ids)} ID sans Validation finale"
        )
        self.btn_copier_affiches.config(
            text=f"📋 Copier les ID vides affichés ({p.nombre(nb_ids)})",
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
                f"{p.nombre(c.nb_ids)} ID erreur · {p.nombre(c.nb_renseignes)} renseigné(s) · "
                f"{p.nombre(c.nb_vides)} vide(s) · {p.nombre(c.nb_lignes_lues)} ligne(s) lue(s)"
            )
        self.lbl_messages.config(text="\n".join(lignes))
        for i in c.details_vides:
            ref = i.reference
            self.detail.insert("", "end", iid=i.id_erreur, values=(
                i.id_erreur, ", ".join(str(l.ligne) for l in i.lignes if not l.validee),
                ref.audit, ref.date_audit, ref.auditeur,
            ))
        actif = "normal" if c.lisible and c.nb_vides else "disabled"
        self.btn_copier_commune.config(state=actif, text=f"📋 Copier les ID ({p.nombre(c.nb_vides)})")
        self.btn_maj_commune.config(state=actif if self.on_mettre_a_jour else "disabled")

    # ── Copie ─────────────────────────────────────────────────────
    def _copier(self, ids: list[str]) -> None:
        if not ids:
            return
        separateur = p.SEPARATEURS_COPIE.get(self.var_separateur.get(), "\n")
        self.clipboard_clear()
        self.clipboard_append(p.texte_ids(ids, separateur))
        self.update_idletasks()
        self._statut(f"✓ {p.nombre(len(ids))} ID erreur copié(s) dans le presse-papiers")

    def _copier_commune(self) -> None:
        c = self._commune_selectionnee()
        if c is not None:
            self._copier(c.ids_vides)

    def _copier_affiches(self) -> None:
        self._copier([i for c in self._communes_filtrees() if c.lisible for i in c.ids_vides])

    def _copier_selection_detail(self) -> None:
        self._copier(list(self.detail.selection()))

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
        "ok": c.nb_renseignes,
        "vides": c.nb_vides,
        "taux": -1 if c.taux is None else c.taux,
        "etat": p.etat(c).cle,
    }[cle]
