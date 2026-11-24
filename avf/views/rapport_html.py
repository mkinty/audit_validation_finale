"""Rapport d'audit en HTML autonome (un seul fichier, sans dépendance externe).

Contenu :
    - en-tête : racine, date, durée, colonnes auditées ;
    - indicateurs : communes, ID erreur, Validation finale renseignée / vide ;
    - synthèse par département ;
    - tableau des communes (tri, recherche, filtre par état), chaque ligne se
      dépliant sur ses ID sans Validation finale (avec Audit, Date audit,
      Auditeur et numéro de ligne) ;
    - copie en un clic : les ID vides d'une commune, ou de tout le rapport,
      avec le séparateur choisi (un par ligne par défaut — prêt à coller dans
      complete_excel_files).

Le fichier reste lisible sans JavaScript (tout est rendu côté Python) ; le
script n'ajoute que tri, filtres, dépliage et copie. Toutes les valeurs issues
des classeurs sont échappées.
"""

from __future__ import annotations

from html import escape
import json

from ..models.audit import AuditCommune, RapportAudit
from . import presentation as p


def generer_html(rapport: RapportAudit) -> str:
    """Document HTML complet du rapport."""
    titre = f"Audit Validation finale — {rapport.date:%d/%m/%Y %H:%M}"
    donnees = {
        c.commune: c.ids_vides for c in rapport.communes if c.lisible and c.nb_vides
    }
    return _GABARIT.format(
        titre=escape(titre),
        style=_CSS,
        entete=_entete(rapport),
        indicateurs=_indicateurs(rapport),
        barre_globale=_barre(rapport.nb_renseignes, rapport.nb_vides, grande=True),
        departements=_departements(rapport),
        lignes=_lignes_communes(rapport),
        nb_vides=p.nombre(rapport.nb_vides),
        desactive_tout="" if rapport.nb_vides else "disabled",
        tous_ids=escape(p.texte_ids(rapport.tous_ids_vides())) or "Aucun ID sans Validation finale.",
        echecs=_echecs(rapport),
        colonne_vf=escape(rapport.colonnes.validation),
        colonne_id=escape(rapport.colonne_id),
        donnees=_json_script(donnees),
        script=_JS,
    )


# ── Blocs ─────────────────────────────────────────────────────────
def _entete(r: RapportAudit) -> str:
    c = r.colonnes
    colonnes = ", ".join(
        f"<code>{escape(n)}</code>"
        for n in (r.colonne_id, c.validation, c.audit, c.date_audit, c.auditeur)
    )
    return (
        f'<p class="meta"><span>📂 <strong>Racine</strong> <code>{escape(r.racine)}</code></span>'
        f"<span>🗓 <strong>Généré le</strong> {r.date:%d/%m/%Y à %H:%M}</span>"
        f"<span>⏱ <strong>Durée</strong> {escape(p.duree(r.duree))}</span></p>"
        f'<p class="meta"><span>🔎 <strong>Colonnes lues</strong> (onglet « Audit ») : {colonnes}</span></p>'
    )


def _tuile(libelle: str, valeur: str, detail: str = "", classe: str = "") -> str:
    return (
        f'<div class="tuile {classe}"><div class="tuile-libelle">{libelle}</div>'
        f'<div class="tuile-valeur">{valeur}</div>'
        f'<div class="tuile-detail">{detail}</div></div>'
    )


def _indicateurs(r: RapportAudit) -> str:
    taux_vide = r.nb_vides / r.nb_ids if r.nb_ids else None
    return "".join([
        _tuile(
            "Communes auditées", p.nombre(r.nb_communes),
            f"✅ {p.nombre(len(r.communes_completes))} complète(s) · "
            f"⏳ {p.nombre(len(r.communes_a_traiter))} à compléter · "
            f"❌ {p.nombre(len(r.en_echec))} en échec",
        ),
        _tuile("ID erreur", p.nombre(r.nb_ids), f"sur {p.nombre(len(r.lues))} fichier(s) lu(s)"),
        _tuile(
            "✅ Validation finale renseignée", p.nombre(r.nb_renseignes),
            f"{p.pourcent(r.taux)} des ID erreur", "tuile-ok",
        ),
        _tuile(
            "⏳ Validation finale vide", p.nombre(r.nb_vides),
            f"{p.pourcent(taux_vide)} des ID erreur", "tuile-vide",
        ),
    ])


def _barre(renseignes: int, vides: int, grande: bool = False) -> str:
    """Barre segmentée fine : part renseignée / part vide (titre = infobulle)."""
    total = renseignes + vides
    classe = "barre grande" if grande else "barre"
    if not total:
        return f'<div class="{classe} barre-neutre" role="img" aria-label="Aucun ID erreur"></div>'
    pct_ok = renseignes * 100 / total
    titre = (
        f"{p.nombre(renseignes)} renseigné(s) ({p.pourcent(renseignes / total)}) · "
        f"{p.nombre(vides)} vide(s) ({p.pourcent(vides / total)})"
    )
    segments = []
    if renseignes:
        segments.append(f'<span class="seg seg-ok" style="width:{pct_ok:.3f}%"></span>')
    if vides:
        segments.append(f'<span class="seg seg-vide" style="width:{100 - pct_ok:.3f}%"></span>')
    return (
        f'<div class="{classe}" role="img" aria-label="{titre}" title="{titre}">'
        + "".join(segments) + "</div>"
    )


def _departements(r: RapportAudit) -> str:
    lignes = []
    for dep, s in r.par_departement().items():
        taux = s["renseignes"] / s["ids"] if s["ids"] else None
        lignes.append(
            f"<tr><td><strong>Dep{escape(dep)}</strong></td>"
            f'<td class="num">{p.nombre(s["communes"])}</td>'
            f'<td class="num">{p.nombre(s["ids"])}</td>'
            f'<td class="num">{p.nombre(s["renseignes"])}</td>'
            f'<td class="num">{p.nombre(s["vides"])}</td>'
            f'<td class="num">{p.nombre(s["erreurs"])}</td>'
            f'<td class="taux">{_barre(s["renseignes"], s["vides"])}'
            f'<span class="pct">{p.pourcent(taux)}</span></td></tr>'
        )
    return "".join(lignes)


def _lignes_communes(r: RapportAudit) -> str:
    if not r.communes:
        return '<tr><td colspan="8" class="vide-tableau">Aucune commune auditée.</td></tr>'
    return "".join(_ligne_commune(c, r) for c in r.communes)


def _ligne_commune(c: AuditCommune, r: RapportAudit) -> str:
    e = p.etat(c)
    recherche = " ".join([c.commune, f"dep{c.departement}", *c.ids_vides]).lower()
    taux_tri = -1 if c.taux is None else c.taux
    alerte = f' <span class="alerte" title="{escape(" · ".join(c.avertissements))}">⚠</span>' if c.avertissements else ""
    copier = (
        f'<button type="button" class="btn btn-petit" data-copier="{escape(c.commune)}" '
        f'title="Copier les {c.nb_vides} ID sans Validation finale">📋 Copier ({p.nombre(c.nb_vides)})</button>'
        if c.lisible and c.nb_vides else ""
    )
    pliable = bool(c.details_vides or c.avertissements or not c.lisible)
    ligne = (
        f'<tr class="commune{" pliable" if pliable else ""}" data-etat="{e.cle}" '
        f'data-recherche="{escape(recherche)}" '
        f'data-k-dep="{escape(c.departement)}" data-k-commune="{escape(c.commune)}" '
        f'data-k-ids="{c.nb_ids}" data-k-ok="{c.nb_renseignes}" data-k-vides="{c.nb_vides}" '
        f'data-k-taux="{taux_tri:.6f}" data-k-etat="{e.cle}"'
        f'{" tabindex=\"0\" aria-expanded=\"false\"" if pliable else ""}>'
        f'<td><span class="chevron" aria-hidden="true">{"▸" if pliable else ""}</span>Dep{escape(c.departement)}</td>'
        f'<td><strong>{escape(c.commune)}</strong>{alerte}</td>'
        f'<td class="num">{p.nombre(c.nb_ids)}</td>'
        f'<td class="num">{p.nombre(c.nb_renseignes)}</td>'
        f'<td class="num">{p.nombre(c.nb_vides)}</td>'
        f'<td class="taux">{_barre(c.nb_renseignes, c.nb_vides) if c.lisible else ""}'
        f'<span class="pct">{p.pourcent(c.taux) if c.lisible else "—"}</span></td>'
        f'<td><span class="etat etat-{e.cle}">{e.icone} {e.libelle}</span></td>'
        f"<td>{copier}</td></tr>"
    )
    if not pliable:
        return ligne
    return ligne + f'<tr class="detail" hidden><td colspan="8">{_detail(c, r)}</td></tr>'


def _detail(c: AuditCommune, r: RapportAudit) -> str:
    morceaux = []
    if c.chemin:
        morceaux.append(f'<p class="chemin">📄 <code>{escape(c.chemin)}</code></p>')
    if not c.lisible:
        morceaux.append(f'<p class="message-erreur">❌ {escape(c.message)}</p>')
    for a in c.avertissements:
        morceaux.append(f'<p class="message-alerte">⚠ {escape(a)}</p>')
    if c.details_vides:
        puces = "".join(
            f'<button type="button" class="puce" data-copier-id="{escape(i.id_erreur)}" '
            f'title="Copier cet ID">{escape(i.id_erreur)}</button>'
            for i in c.details_vides
        )
        morceaux.append(
            f'<div class="sous-titre">ID erreur sans {escape(r.colonnes.validation)} '
            f"({p.nombre(c.nb_vides)}) — cliquer sur un ID pour le copier</div>"
            f'<div class="puces">{puces}</div>'
        )
        col = r.colonnes
        lignes = "".join(
            f"<tr><td><code>{escape(i.id_erreur)}</code>"
            f'{" <span class=\"doublon\" title=\"ID présent sur plusieurs lignes\">×" + str(len(i.lignes)) + "</span>" if i.doublon else ""}</td>'
            f'<td class="num">{", ".join(str(l.ligne) for l in i.lignes if not l.validee)}</td>'
            f"<td>{escape(i.reference.audit) or _VIDE}</td>"
            f"<td>{escape(i.reference.date_audit) or _VIDE}</td>"
            f"<td>{escape(i.reference.auditeur) or _VIDE}</td></tr>"
            for i in c.details_vides
        )
        morceaux.append(
            '<table class="table-detail"><thead><tr>'
            f"<th>{escape(r.colonne_id)}</th><th>Ligne(s)</th><th>{escape(col.audit)}</th>"
            f"<th>{escape(col.date_audit)}</th><th>{escape(col.auditeur)}</th>"
            f"</tr></thead><tbody>{lignes}</tbody></table>"
        )
    return "".join(morceaux)


def _echecs(r: RapportAudit) -> str:
    if not r.en_echec:
        return ""
    lignes = "".join(
        f"<li><strong>{escape(c.commune)}</strong> — {p.etat(c).icone} {escape(c.message)}</li>"
        for c in r.en_echec
    )
    return (
        '<section class="carte"><h2>❌ Communes non auditées '
        f"({p.nombre(len(r.en_echec))})</h2><ul class=\"liste-echecs\">{lignes}</ul></section>"
    )


def _json_script(donnees) -> str:
    """JSON sûr dans une balise <script> (aucune fermeture de balise possible)."""
    return (
        json.dumps(donnees, ensure_ascii=False)
        .replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    )


_VIDE = '<span class="muet">—</span>'


# ── Gabarit ───────────────────────────────────────────────────────
_GABARIT = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titre}</title>
<style>{style}</style>
</head>
<body>
<div class="viz-root">
<header class="bandeau">
  <div>
    <h1>Audit de la {colonne_vf}</h1>
    {entete}
  </div>
  <button type="button" class="btn btn-theme" id="theme" title="Basculer clair / sombre">◐ Thème</button>
</header>

<main>
<section class="tuiles">{indicateurs}</section>

<section class="carte">
  <div class="legende">
    <span><i class="pastille seg-ok"></i>✅ {colonne_vf} renseignée</span>
    <span><i class="pastille seg-vide"></i>⏳ {colonne_vf} vide</span>
  </div>
  {barre_globale}
</section>

<section class="carte">
  <h2>Synthèse par département</h2>
  <div class="defile">
  <table class="table">
    <thead><tr><th>Département</th><th class="num">Communes</th><th class="num">ID erreur</th>
    <th class="num">Renseignés</th><th class="num">Vides</th><th class="num">En échec</th><th>Taux renseigné</th></tr></thead>
    <tbody>{departements}</tbody>
  </table>
  </div>
</section>

<section class="carte">
  <div class="barre-outils">
    <h2>Communes</h2>
    <input type="search" id="recherche" placeholder="Commune ou ID erreur…" aria-label="Rechercher">
    <div class="filtres" role="group" aria-label="Filtrer par état">
      <button type="button" class="filtre actif" data-filtre="tous">Toutes</button>
      <button type="button" class="filtre" data-filtre="a-completer">⏳ À compléter</button>
      <button type="button" class="filtre" data-filtre="complete">✅ Complètes</button>
      <button type="button" class="filtre" data-filtre="echec">❌ En échec</button>
    </div>
    <label class="separateur">Séparateur
      <select id="separateur">
        <option value="ligne">un par ligne</option>
        <option value="virgule">virgule</option>
        <option value="point-virgule">point-virgule</option>
        <option value="espace">espace</option>
      </select>
    </label>
    <button type="button" class="btn btn-principal" data-copier-tout="visibles" title="Copie les ID sans Validation finale des communes affichées (filtre et recherche pris en compte)" {desactive_tout}>📋 Copier les ID vides (<span id="nb-visibles">{nb_vides}</span>)</button>
  </div>
  <div class="defile">
  <table class="table" id="communes">
    <thead><tr>
      <th data-tri="dep">Département</th><th data-tri="commune">Commune</th>
      <th data-tri="ids" class="num">ID erreur</th><th data-tri="ok" class="num">Renseignés</th>
      <th data-tri="vides" class="num">Vides</th><th data-tri="taux">Taux renseigné</th>
      <th data-tri="etat">État</th><th>Copie</th>
    </tr></thead>
    <tbody>{lignes}</tbody>
  </table>
  </div>
  <p class="aucun" id="aucun" hidden>Aucune commune ne correspond.</p>
</section>

<section class="carte">
  <div class="barre-outils">
    <h2>Tous les ID erreur sans {colonne_vf} ({nb_vides})</h2>
    <button type="button" class="btn" data-copier-tout="tous" {desactive_tout}>📋 Copier tout</button>
  </div>
  <pre class="bloc-ids" id="tous-ids">{tous_ids}</pre>
</section>

{echecs}
</main>
<footer>Rapport généré par l'outil « Audit Validation finale » — une copie du rapport ne se met pas à jour : relancer l'audit pour un état actualisé.</footer>
</div>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<script id="donnees" type="application/json">{donnees}</script>
<script>{script}</script>
</body>
</html>
"""


_CSS = """
:root{color-scheme:light}
.viz-root{
  --surface-0:#f3f3f1;--surface-1:#fcfcfb;--surface-2:#f0efec;--bordure:#e2e1dc;
  --text-primary:#0b0b0b;--text-secondary:#52514e;--text-muted:#7a7974;
  --accent:#2a78d6;--accent-texte:#ffffff;
  --good:#0ca30c;--warning:#fab219;--critical:#d03b3b;
  --good-fond:#e7f6e7;--warning-fond:#fff4dc;--critical-fond:#fbe9e9;
  --ombre:0 1px 2px rgba(0,0,0,.06),0 2px 8px rgba(0,0,0,.04);
}
@media (prefers-color-scheme:dark){
  :root:where(:not([data-theme="light"])){color-scheme:dark}
  :root:where(:not([data-theme="light"])) .viz-root{
    --surface-0:#111110;--surface-1:#1a1a19;--surface-2:#262624;--bordure:#383835;
    --text-primary:#f0efec;--text-secondary:#c3c2b7;--text-muted:#8f8e86;
    --accent:#3987e5;--good-fond:#14301a;--warning-fond:#3a2e10;--critical-fond:#3a1818;
    --ombre:none;
  }
}
:root[data-theme="dark"]{color-scheme:dark}
:root[data-theme="dark"] .viz-root{
  --surface-0:#111110;--surface-1:#1a1a19;--surface-2:#262624;--bordure:#383835;
  --text-primary:#f0efec;--text-secondary:#c3c2b7;--text-muted:#8f8e86;
  --accent:#3987e5;--good-fond:#14301a;--warning-fond:#3a2e10;--critical-fond:#3a1818;
  --ombre:none;
}
*{box-sizing:border-box}
body{margin:0}
.viz-root{min-height:100vh;background:var(--surface-0);color:var(--text-primary);
  font:14px/1.45 "Segoe UI",system-ui,-apple-system,Roboto,sans-serif}
code,pre{font-family:Consolas,"Cascadia Mono",ui-monospace,monospace;font-size:12.5px}
.bandeau{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;
  padding:24px 32px 16px;background:var(--surface-1);border-bottom:1px solid var(--bordure)}
h1{font-size:22px;margin:0 0 8px;font-weight:600}
h2{font-size:15px;margin:0 0 12px;font-weight:600}
.meta{display:flex;flex-wrap:wrap;gap:6px 22px;margin:2px 0;color:var(--text-secondary);font-size:13px}
.meta code{color:var(--text-primary);word-break:break-all}
main{padding:20px 32px 8px;max-width:1400px}
.tuiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-bottom:12px}
.tuile,.carte{background:var(--surface-1);border:1px solid var(--bordure);border-radius:10px;box-shadow:var(--ombre)}
.tuile{padding:14px 16px}
.tuile-libelle{color:var(--text-secondary);font-size:13px}
.tuile-valeur{font-size:30px;font-weight:600;margin:2px 0;font-variant-numeric:tabular-nums}
.tuile-detail{color:var(--text-muted);font-size:12.5px}
.tuile-ok{border-left:4px solid var(--good)}
.tuile-vide{border-left:4px solid var(--warning)}
.carte{padding:16px 18px;margin-bottom:12px}
.legende{display:flex;gap:18px;color:var(--text-secondary);font-size:13px;margin-bottom:8px}
.pastille{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:-1px}
.barre{display:flex;gap:2px;height:8px;width:100%;min-width:80px;border-radius:999px;overflow:hidden}
.barre.grande{height:12px}
.barre-neutre{background:var(--surface-2)}
.seg{display:block;height:100%;min-width:2px}
.seg-ok{background:var(--good)}
.seg-vide{background:var(--warning)}
.seg:first-child{border-radius:999px 0 0 999px}
.seg:last-child{border-radius:0 999px 999px 0}
.seg:only-child{border-radius:999px}
.defile{overflow-x:auto}
.table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
.table th{text-align:left;font-weight:600;color:var(--text-secondary);font-size:12.5px;
  padding:8px 10px;border-bottom:1px solid var(--bordure);white-space:nowrap;user-select:none}
.table th[data-tri]{cursor:pointer}
.table th[data-tri]:hover{color:var(--text-primary)}
.table th[data-ordre="asc"]::after{content:" ▲";font-size:10px}
.table th[data-ordre="desc"]::after{content:" ▼";font-size:10px}
.table td{padding:8px 10px;border-bottom:1px solid var(--bordure);vertical-align:middle}
.num,.table th.num{text-align:right}
td.taux{min-width:190px}
td.taux .barre{display:inline-flex;width:120px;vertical-align:middle;margin-right:8px}
.pct{color:var(--text-secondary);font-size:12.5px}
tr.commune.pliable{cursor:pointer}
tr.commune:hover td{background:var(--surface-2)}
tr.commune:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}
.chevron{display:inline-block;width:14px;color:var(--text-muted);transition:transform .15s}
tr.ouvert .chevron{transform:rotate(90deg)}
tr.detail td{background:var(--surface-0);padding:12px 16px 14px 34px}
.etat{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;white-space:nowrap;
  border:1px solid var(--bordure);color:var(--text-primary)}
.etat-complete{background:var(--good-fond)}
.etat-a-completer{background:var(--warning-fond)}
.etat-echec{background:var(--critical-fond)}
.alerte{color:var(--critical);cursor:help}
.btn{font:inherit;font-size:13px;padding:6px 12px;border-radius:7px;cursor:pointer;
  border:1px solid var(--bordure);background:var(--surface-1);color:var(--text-primary)}
.btn:hover{background:var(--surface-2)}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn-petit{padding:3px 9px;font-size:12px;white-space:nowrap}
.btn-principal{background:var(--accent);border-color:var(--accent);color:var(--accent-texte)}
.btn-principal:hover{background:var(--accent);filter:brightness(1.08)}
.barre-outils{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:10px}
.barre-outils h2{margin:0;margin-right:auto}
input[type=search],select{font:inherit;font-size:13px;padding:6px 10px;border-radius:7px;
  border:1px solid var(--bordure);background:var(--surface-1);color:var(--text-primary)}
input[type=search]{min-width:260px}
.filtres{display:flex;gap:4px}
.filtre{font:inherit;font-size:12.5px;padding:5px 10px;border-radius:999px;cursor:pointer;
  border:1px solid var(--bordure);background:transparent;color:var(--text-secondary)}
.filtre.actif{background:var(--text-primary);color:var(--surface-1);border-color:var(--text-primary)}
.separateur{color:var(--text-secondary);font-size:12.5px;display:flex;align-items:center;gap:6px}
.sous-titre{font-weight:600;font-size:13px;margin:6px 0}
.puces{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.puce{font:12.5px Consolas,ui-monospace,monospace;padding:3px 9px;border-radius:6px;cursor:pointer;
  border:1px solid var(--bordure);background:var(--surface-1);color:var(--text-primary)}
.puce:hover{border-color:var(--accent)}
.table-detail{border-collapse:collapse;font-size:12.5px;min-width:520px}
.table-detail th,.table-detail td{padding:4px 12px 4px 0;text-align:left;border-bottom:1px solid var(--bordure)}
.table-detail th{color:var(--text-secondary);font-weight:600}
.doublon{font-size:11px;color:var(--text-muted)}
.chemin{margin:0 0 6px;color:var(--text-secondary);word-break:break-all}
.message-erreur,.message-alerte{margin:4px 0}
.muet{color:var(--text-muted)}
.bloc-ids{margin:0;max-height:280px;overflow:auto;padding:12px;border-radius:8px;
  background:var(--surface-0);border:1px solid var(--bordure);white-space:pre-wrap;word-break:break-all}
.liste-echecs{margin:0;padding-left:18px}
.vide-tableau,.aucun{color:var(--text-muted);text-align:center;padding:18px}
footer{color:var(--text-muted);font-size:12px;padding:8px 32px 28px}
.toast{position:fixed;left:50%;bottom:28px;transform:translateX(-50%) translateY(20px);opacity:0;
  background:#0b0b0b;color:#fff;padding:10px 18px;border-radius:8px;font:14px "Segoe UI",sans-serif;
  transition:opacity .2s,transform .2s;pointer-events:none;z-index:10}
.toast.visible{opacity:1;transform:translateX(-50%) translateY(0)}
@media print{.btn,.filtres,input,select,.separateur,.btn-theme{display:none!important}
  tr.detail{display:table-row!important}.viz-root{background:#fff}}
"""

_JS = r"""
(function(){
  const donnees = JSON.parse(document.getElementById('donnees').textContent || '{}');
  const SEP = {'ligne':'\n','virgule':', ','point-virgule':';','espace':' '};
  const sep = () => SEP[document.getElementById('separateur').value] || '\n';
  const toastEl = document.getElementById('toast');
  let minuteur;
  function toast(msg){
    toastEl.textContent = msg; toastEl.classList.add('visible');
    clearTimeout(minuteur); minuteur = setTimeout(() => toastEl.classList.remove('visible'), 1800);
  }
  function copieSecours(texte){
    const zone = document.createElement('textarea');
    zone.value = texte; zone.setAttribute('readonly', '');
    zone.style.position = 'fixed'; zone.style.opacity = '0';
    document.body.appendChild(zone); zone.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch(e) {}
    zone.remove(); return ok;
  }
  async function copier(ids){
    if(!ids.length){ toast('Aucun ID à copier'); return; }
    const texte = ids.join(sep());
    let ok = false;
    try { await navigator.clipboard.writeText(texte); ok = true; }
    catch(e){ ok = copieSecours(texte); }
    toast(ok ? `✓ ${ids.length} ID erreur copié${ids.length > 1 ? 's' : ''}` : 'Copie impossible — sélectionnez le texte manuellement');
  }
  function idsVisibles(){
    const visibles = new Set([...document.querySelectorAll('#communes tr.commune')]
      .filter(tr => !tr.hidden).map(tr => tr.dataset.kCommune));
    return Object.entries(donnees).filter(([c]) => visibles.has(c)).flatMap(([, ids]) => ids);
  }

  document.addEventListener('click', ev => {
    const b = ev.target.closest('[data-copier],[data-copier-id],[data-copier-tout]');
    if(b){
      ev.stopPropagation();
      if(b.hasAttribute('data-copier-tout'))
        copier(b.dataset.copierTout === 'visibles' ? idsVisibles() : Object.values(donnees).flat());
      else if(b.dataset.copierId) copier([b.dataset.copierId]);
      else copier(donnees[b.dataset.copier] || []);
      return;
    }
    const tr = ev.target.closest('tr.commune.pliable');
    if(tr) basculer(tr);
  });
  document.addEventListener('keydown', ev => {
    const tr = ev.target.closest && ev.target.closest('tr.commune.pliable');
    if(tr && (ev.key === 'Enter' || ev.key === ' ')){ ev.preventDefault(); basculer(tr); }
  });
  function basculer(tr){
    const detail = tr.nextElementSibling;
    if(!detail || !detail.classList.contains('detail')) return;
    const ouvert = detail.hidden;
    detail.hidden = !ouvert; tr.classList.toggle('ouvert', ouvert);
    tr.setAttribute('aria-expanded', String(ouvert));
  }

  // Recherche + filtre par état
  let filtre = 'tous';
  const champ = document.getElementById('recherche');
  function appliquer(){
    const q = champ.value.trim().toLowerCase();
    let n = 0;
    document.querySelectorAll('#communes tr.commune').forEach(tr => {
      const okEtat = filtre === 'tous' || tr.dataset.etat === filtre;
      const okTexte = !q || tr.dataset.recherche.includes(q);
      tr.hidden = !(okEtat && okTexte);
      const d = tr.nextElementSibling;
      if(d && d.classList.contains('detail') && tr.hidden){ d.hidden = true; tr.classList.remove('ouvert'); }
      if(!tr.hidden) n++;
    });
    document.getElementById('aucun').hidden = n > 0;
    const compteur = document.getElementById('nb-visibles');
    if(compteur) compteur.textContent = idsVisibles().length.toLocaleString('fr-FR');
  }
  champ.addEventListener('input', appliquer);
  document.querySelectorAll('.filtre').forEach(b => b.addEventListener('click', () => {
    document.querySelectorAll('.filtre').forEach(x => x.classList.remove('actif'));
    b.classList.add('actif'); filtre = b.dataset.filtre; appliquer();
  }));

  // Tri des colonnes (la ligne de détail suit sa commune)
  document.querySelectorAll('#communes th[data-tri]').forEach(th => th.addEventListener('click', () => {
    const cle = th.dataset.tri, asc = th.dataset.ordre !== 'asc';
    document.querySelectorAll('#communes th').forEach(x => delete x.dataset.ordre);
    th.dataset.ordre = asc ? 'asc' : 'desc';
    const corps = document.querySelector('#communes tbody');
    const attr = 'k' + cle.charAt(0).toUpperCase() + cle.slice(1);
    const paires = [...corps.querySelectorAll('tr.commune')].map(tr => {
      const d = tr.nextElementSibling;
      return [tr, d && d.classList.contains('detail') ? d : null];
    });
    paires.sort(([a], [b]) => {
      const x = a.dataset[attr], y = b.dataset[attr];
      const nx = parseFloat(x), ny = parseFloat(y);
      const r = (!isNaN(nx) && !isNaN(ny) && /^-?[\d.]+$/.test(x) && /^-?[\d.]+$/.test(y))
        ? nx - ny : x.localeCompare(y, 'fr', {numeric: true});
      return asc ? r : -r;
    });
    paires.forEach(([tr, d]) => { corps.appendChild(tr); if(d) corps.appendChild(d); });
  }));

  // Thème clair / sombre (le choix du lecteur l'emporte sur celui du système)
  document.getElementById('theme').addEventListener('click', () => {
    const racine = document.documentElement;
    const sombre = racine.dataset.theme
      ? racine.dataset.theme === 'dark'
      : matchMedia('(prefers-color-scheme: dark)').matches;
    racine.dataset.theme = sombre ? 'light' : 'dark';
  });
})();
"""
