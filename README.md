# Audit de la Validation finale — Projet SNA

Application Tkinter qui audite la colonne **« Validation finale »** des fichiers
audit des communes, produit un rapport (fenêtre ou page HTML), puis complète en
masse les ID erreur restés sans validation, avec le **même moteur que
complete_excel_files**.

```
<racine>/Dep<XX>/<commune>/audit_*.xlsx        (XX = 2 premiers caractères du code INSEE)
```

## Déroulé dans l'interface

1. **Dossier racine** : saisi ou choisi (📁), mémorisé.
2. **Communes** (dès que la racine est renseignée) : `🔍 Détecter toutes les
   communes` parcourt tous les `<racine>/Dep<XX>/<commune>` (celles sans fichier
   audit sont listées mais non présélectionnées) ; `✏️ Renseigner…` accepte des
   codes INSEE (`13001 2A004`…), le dossier `Dep<XX>` étant déduit des deux
   premiers caractères. Les codes sans dossier sont signalés et ignorés.
   L'utilisateur ne choisit jamais de département.
3. **Audit** : lecture des colonnes `ID erreur`, `Validation finale`, `Audit`,
   `Date audit`, `Auditeur` de l'onglet **« Audit »**, journal en temps réel,
   arrêt possible. Le rapport s'ouvre en popup à la fin ; `💾 Exporter en HTML`
   l'enregistre et propose de l'ouvrir dans le navigateur.
4. **Mise à jour** : `🛠️ Mettre à jour les N ID sans Validation finale`
   (depuis la fenêtre principale, ou depuis le rapport : toutes les communes ou
   seulement la commune sélectionnée).

La sélection (racine, communes) et les options sont conservées
d'un lancement à l'autre.

## Règle d'audit

| Cas | Compté comme |
|---|---|
| `Validation finale` renseignée | ✅ renseigné |
| `Validation finale` vide **ou composée d'espaces** | ⏳ vide |
| ID présent sur plusieurs lignes | renseigné seulement si **toutes** ses lignes le sont |
| Ligne sans ID erreur | ignorée |

Les ID sont comparés sous forme normalisée (espaces retirés, majuscules), comme
dans complete_excel_files : un ID audité est toujours retrouvé par la mise à jour.

États d'une commune : **Complète**, **À compléter**, **Aucun ID erreur**,
**Fichier introuvable** (pas d'`audit_*.xlsx`), **Erreur de lecture** (onglet
« Audit » absent, colonne `ID erreur` ou `Validation finale` absente, fichier
illisible). Une colonne informative absente (Audit, Date audit, Auditeur) ne
fait qu'un avertissement. Une **copie de conflit** OneDrive/SharePoint
(deuxième `audit_*.xlsx`) est signalée dès l'audit, car la mise à jour la refusera.

Les libellés des colonnes, la ligne d'en-tête et la colonne des ID erreur se
règlent dans `⚙️ Colonnes auditées` (la colonne ID et la ligne d'en-tête sont
partagées avec les règles de mise à jour).

## Rapport

Mêmes informations dans la popup et dans le HTML :

- indicateurs : communes (complètes / à compléter / en échec), ID erreur,
  Validation finale renseignée et vide (avec pourcentages), barre de répartition ;
- synthèse par département (HTML) ;
- tableau des communes : tri par colonne, recherche (commune ou ID erreur),
  filtre par état ;
- détail d'une commune : ses ID sans validation, avec numéro de ligne, Audit,
  Date audit, Auditeur, chemin du fichier et avertissements.

**Copie en un clic** (séparateur au choix : un par ligne, virgule,
point-virgule, espace) :

| Où | Quoi |
|---|---|
| HTML · bouton `📋 Copier (n)` d'une commune | ses ID vides |
| HTML · clic sur un ID | cet ID |
| HTML · `📋 Copier les ID vides (n)` | ID des communes **affichées** (filtre et recherche respectés) |
| HTML · section « Tous les ID… » · `📋 Copier tout` | tous les ID vides du rapport |
| Popup · `📋 Copier les ID (n)` | ID vides de la commune sélectionnée |
| Popup · `📋 Copier les ID vides affichés` | ID des communes affichées |
| Popup · double-clic / Ctrl+C dans le détail | ID sélectionnés |

Le HTML est un fichier unique sans dépendance externe, lisible sans JavaScript,
en thème clair ou sombre (selon le système, bouton `◐ Thème` pour basculer),
imprimable (détails dépliés).

## Mise à jour : le processus de complete_excel_files

La mise à jour appelle **tel quel** le moteur de complete_excel_files
(`sna.services.traitement.traiter_communes`), copié dans `sna/` (voir
[sna/PROVENANCE.md](sna/PROVENANCE.md)) :

- même recherche de fichier, même refus en cas de copie de conflit ;
- même repérage rapide, mêmes règles (conditions en ET + action), même retrait
  du filtre automatique et des lignes/colonnes masquées, même reprise de la
  mise en forme, dates écrites comme de vraies dates ;
- même écriture : openpyxl, ou **via Excel** (co-édition SharePoint, détection du
  classeur verrouillé par un collègue, relecture après enregistrement) ;
- mêmes ID en attente (`ids_en_attente.json`) pour les ID non appliqués.

Seuls les ID **sans Validation finale** du rapport sont transmis.

**Valeurs par défaut** — celles de complete_excel_files :

| Règle | Action |
|---|---|
| R1 | `Audit` ← `OK` |
| R2 | `Date audit` ← date du jour |
| R3 | `Auditeur` ← `Validation en masse` |
| R4 | `Validation finale` ← `OK` |

`⚙️ Réglage des valeurs (règles)` ouvre **le constructeur de règles de
complete_excel_files** : changer les valeurs, ajouter des conditions (ex. ne
poser `Audit` que s'il est vide), désactiver une règle, recopier une colonne,
effacer… `↺ Valeurs par défaut` restaure les 4 règles.

Options :

- **Simulation** : compte les modifications sans rien écrire ;
- **📅 « Date audit » = date du jour** (activée par défaut) : la règle qui écrit
  `Date audit` prend la date du jour de la mise à jour. Dans complete_excel_files,
  la date par défaut est figée au jour où la configuration a été créée ;
- **🔐 Écrire via Excel** : activée si Excel et pywin32 sont présents ;
- **🔁 Ré-auditer après mise à jour** : les communes traitées sont relues et le
  rapport (fenêtre comprise) se met à jour.

Une confirmation récapitule les règles, le nombre d'ID et de communes, le mode
d'écriture et les communes qui seront refusées pour conflit.

## Installation et lancement

```bash
py -3.14 -m venv .venv
.venv\Scripts\python -m pip install openpyxl pywin32 pytest
.venv\Scripts\python main.py
```

(ou `uv sync` puis `uv run main.py`). `pywin32` n'est utile que pour l'écriture via Excel.

## Structure

```
audit_validation_finale/
├── main.py                      # point d'entrée
├── avf/                         # l'outil d'audit
│   ├── chemins.py, config.py    #   config_audit_vf.json (distinct de complete_excel_files)
│   ├── models/                  #   ColonnesAudit · LigneAudit, IdErreurAudit, AuditCommune, RapportAudit
│   ├── services/                #   logique métier, sans Tkinter
│   │   ├── detection.py         #     communes : détection sous la racine, saisie, vérification
│   │   ├── lecture_audit.py     #     lecture XML en flux des colonnes (repli openpyxl)
│   │   ├── audit.py             #     orchestration → RapportAudit
│   │   ├── mise_a_jour.py       #     pont vers le moteur sna (règles, date du jour, session Excel)
│   │   └── export.py            #     écriture du HTML, ouverture navigateur
│   ├── views/                   #   rendu
│   │   ├── presentation.py      #     formats partagés (nombres, %, états)
│   │   └── rapport_html.py      #     rapport HTML autonome
│   └── ui/                      #   Tkinter
│       ├── application.py       #     fenêtre principale
│       ├── fenetre_rapport.py   #     popup du rapport
│       └── composants.py        #     thème, cartes, boutons, dialogues
├── sna/                         # moteur de complete_excel_files (cf. PROVENANCE.md)
└── tests/                       # avf + tests d'origine du moteur (tests/moteur_sna)
```

## Tests

```bash
.venv\Scripts\python -m pytest
```

318 tests : détection, lecture (chemin rapide et repli openpyxl comparés),
audit, rapport HTML (structure, échappement, données de copie), configuration,
mise à jour de bout en bout (valeurs écrites cellule par cellule, simulation,
conflit, règles personnalisées, identité avec le moteur appelé directement),
écriture via une **vraie instance Excel** (ignorée sans Excel), et interface
Tkinter pilotée (détection des communes → audit → popup → copie → mise à jour → ré-audit → export).

## Points d'attention

- Comme dans complete_excel_files, l'écriture openpyxl réécrit tout le
  classeur : testez sur une copie, ou préférez l'écriture via Excel.
- Le rapport HTML est une photographie : relancer l'audit pour un état à jour.
- Les règles par défaut **écrasent** `Audit` et `Auditeur` des ID ciblés, même
  s'ils sont déjà renseignés (comportement de complete_excel_files). Pour les
  préserver, ajouter une condition « `Audit` vide » dans le constructeur de règles.
