# Audit de la Validation finale — Projet SNA

Application Tkinter qui audite, ligne par ligne, les colonnes **Audit**, **Date audit**,
**Auditeur** et **Validation finale** des fichiers audit des communes. Elle produit un
rapport (fenêtre, page HTML, classeur Excel ou CSV), puis corrige en masse les ID erreur
non valides, avec le **même moteur que complete_excel_files**.

```
<racine>/Dep<XX>/<commune>/audit_*.xlsx        (XX = 2 premiers caractères du code INSEE)
```

## Déroulé dans l'interface

1. **Dossier racine** : choisir la source.
   - `📁 Dossier local (parcourir)` : le dossier saisi ou choisi est mémorisé.
   - `🌐 Lien SharePoint (version serveur)` : collez le lien du dossier racine, puis
     cliquez sur `🔗 Se connecter`. L'audit lit alors les fichiers **sur le serveur**
     (voir « Audit depuis SharePoint »).
2. **Communes** (dès que la racine est renseignée) : `🔍 Détecter toutes les
   communes` parcourt tous les `<racine>/Dep<XX>/<commune>` (celles sans fichier
   audit sont listées mais non présélectionnées) ; `✏️ Renseigner…` accepte des
   codes INSEE (`13001 2A004`…), le dossier `Dep<XX>` étant déduit des deux
   premiers caractères. Les codes sans dossier sont signalés et ignorés.
   L'utilisateur ne choisit jamais de département.
   `🎯 ID erreur…` limite l'audit à des **ID erreur** saisis (voir « Audit de
   quelques ID erreur »).
3. **Audit** : contrôle de chaque ligne de l'onglet **« Audit »**, journal en temps réel,
   arrêt possible. Le rapport s'ouvre en popup à la fin. `💾 HTML`, `📊 Excel` et `📑 CSV`
   l'exportent, puis proposent de l'ouvrir dans le navigateur ou dans Excel.
4. **Mise à jour** : `🛠️ Mettre à jour les N ID non valides`, depuis la fenêtre
   principale ou depuis le rapport (toutes les communes, ou seulement la commune
   sélectionnée).

La sélection (racine, communes) et les options sont conservées
d'un lancement à l'autre. Toute la fenêtre **défile** (barre à droite, molette) quand
l'écran est trop petit pour l'afficher en entier. Le journal et la liste des communes
gardent leur propre défilement.

## Audit depuis SharePoint

Une modification faite dans le dossier OneDrive local peut ne jamais arriver sur
SharePoint : conflit de synchronisation, copie « -NOMDUPC », envoi en échec. Auditer
le dossier local montrerait alors un état que personne d'autre ne voit. Avec la source
**Lien SharePoint**, l'audit lit ce que contient réellement le serveur.

**Liens acceptés** : l'adresse du dossier ouvert dans le navigateur (barre d'adresse, y
compris `…/Forms/AllItems.aspx?id=…`), un lien de partage `…/:f:/r/sites/…` ou un
chemin `\\…@SSL\DavWWWRoot\…`. Un lien de partage opaque (`…/:f:/s/…/EaBc…`) ne
contient pas le chemin du dossier : il est refusé, avec la marche à suivre.

**Accès** : `🔗 Se connecter` cherche, dans l'ordre :

| Mode | Comment | Conditions |
|---|---|---|
| **Accès direct (WebDAV)** | Le lien devient un chemin réseau `\\hôte@SSL\DavWWWRoot\…`, comme « Afficher dans l'Explorateur de fichiers » de SharePoint. Détection des communes et lecture des fichiers se font **directement sur le serveur**, exactement comme pour un dossier local. | Service Windows **WebClient** démarré ; session Microsoft 365 ouverte sur le poste (se connecter une fois à SharePoint dans Edge en cochant « Rester connecté »). |
| **Via Excel, bibliothèque synchronisée** (1er repli) | Les communes sont listées depuis la bibliothèque **synchronisée par OneDrive**. Chaque `audit_*.xlsx` est ensuite **ouvert par son adresse SharePoint** dans l'instance Excel dédiée (lecture seule, compte Office), copié dans un dossier temporaire puis lu. Le contenu audité est donc celui du serveur. | Bibliothèque synchronisée sur le poste ; Excel installé. Plus lent : une ouverture Excel par commune. |
| **Via Excel, sans synchronisation** (2e repli) | Le dossier ne peut pas être parcouru : on renseigne les communes (`✏️`) ou les ID erreur (`🎯`), et « Détecter toutes les communes » est désactivé. Le fichier de chaque commune est ouvert à l'adresse `<lien>/Dep<XX>/<commune>/audit_<commune>.xlsx`. Une adresse vide donne « Fichier introuvable ». Les copies de conflit ne sont pas détectables dans ce mode. | Excel installé. |

**Erreur 224 « Accès refusé… sites approuvés »** : Windows n'a pas de session
SharePoint pour l'accès direct. Depuis la fin d'Internet Explorer, Edge ne lui transmet
plus la connexion. L'outil passe alors à un repli et affiche la cause. Pour rétablir
l'accès direct, ouvrez le dossier dans SharePoint avec Edge puis cliquez sur « Afficher
dans l'Explorateur de fichiers ».

**Connexion Microsoft 365** : si Office demande de se connecter au compte professionnel
(fenêtre « Se connecter »), l'outil le détecte en quelques secondes. Une fenêtre Excel
**visible** s'ouvre alors : connectez-vous une fois, puis le traitement reprend, en audit
comme en mise à jour. Office mémorise la connexion pour les fois suivantes. Avant, une
instance Excel invisible restait bloquée sur cette fenêtre jusqu'au délai anti-blocage.

Le mode obtenu s'affiche sous le lien et figure dans le journal et dans tous les
rapports (« Source : SharePoint — … »). En mode Excel, l'adresse SharePoint de chaque
fichier est celle affichée dans le rapport. La connexion se refait automatiquement,
sans fenêtre, au lancement suivant.

**Mise à jour** : en source SharePoint, elle passe **toujours** par Excel (co-édition),
même si « Écrire via Excel » est décoché, et ne réécrit jamais directement un fichier
sur le serveur. Via une bibliothèque synchronisée, Excel est ouvert même si la copie
locale ne contient pas l'ID, puisqu'elle peut différer du serveur. Une mise à jour est
refusée si le rapport a été produit sur une autre source que celle choisie à l'étape 1.

**Limite** : en mode via Excel, la **simulation** compte les modifications sur la copie
locale synchronisée, pas sur le serveur.

## Audit de quelques ID erreur

`🎯 ID erreur…` (étape 2) audite seulement les ID erreur collés, au lieu des communes
entières. On les colle séparés par des espaces, des virgules ou des retours à la ligne :
`13001_2, 13001_15  2a004_1`.

- La commune, et donc le dossier `<racine>/Dep<XX>/<commune>`, est déduite du code INSEE
  en préfixe, comme dans complete_excel_files (`sna.services.identifiants`). Les communes
  concernées sont listées et sélectionnées.
- Seules les lignes de ces ID sont contrôlées, toutes leurs lignes en cas de doublon.
  Les indicateurs, le détail et les exports ne portent que sur elles.
- Un ID **absent de son fichier** est signalé : avertissement sur la commune, ligne
  « ID introuvable » dans les exports CSV et Excel, et liste dans le bilan.
- Sont aussi signalés : un ID dont la commune n'a pas de dossier, et un code INSEE saisi
  sans numéro (`13002`). Ils sont ignorés.
- Le **périmètre** (« 🎯 Audit limité à N ID erreur saisi(s) — K introuvable(s) ») figure
  sous la liste des communes, dans le journal, la fenêtre du rapport, le HTML et la
  synthèse Excel.
- La mise à jour ne touche que les ID saisis non valides. Le ré-audit qui la suit reste
  limité aux mêmes ID.
- Les ID saisis sont mémorisés d'un lancement à l'autre. `✖ Communes entières`,
  `🔍 Détecter toutes les communes` ou `✏️ Renseigner…` reviennent à l'audit des
  communes entières.

Cela fonctionne aussi en source SharePoint.

## Règle d'audit

Chaque **ligne** portant un ID erreur est contrôlée sur quatre critères :

| Colonne | Critère (par défaut) |
|---|---|
| `Audit` | égale à `OK` ou `KO` |
| `Date audit` | renseignée et lisible comme une date, **même saisie en texte** |
| `Auditeur` | renseignée |
| `Validation finale` | égale à `OK` |

Une ligne qui respecte les quatre critères est **validée**. Sinon elle est **non valide**,
avec un motif pour chaque critère non respecté (ex. `Validation finale « Validé » ≠ OK`,
`Date audit « à faire » n'est pas une date`).

- Les valeurs sont comparées sans tenir compte de la casse ni des espaces autour :
  `ok`, ` OK ` valent `OK`. Une cellule faite d'espaces est vide.
- **Dates acceptées** : les vraies dates Excel, et les textes `12/03/2026`, `12-03-2026`,
  `12.03.2026`, `12/03/26`, `2026-03-12` (éventuellement avec une heure), `12 mars 2026`,
  `1er mars 2026`, `12 sept. 2026`, ainsi qu'un numéro de série Excel. Une date
  impossible (`31/02/2026`) est refusée.
- **ID erreur** : un ID peut apparaître sur plusieurs lignes. Il est validé si toutes ses
  lignes le sont, et non valide dès qu'une de ses lignes ne l'est pas. C'est cet ID qui
  est copié et mis à jour (toutes ses lignes).
- Une ligne sans ID erreur est ignorée.

Les ID sont comparés sous forme normalisée (espaces retirés, majuscules), comme
dans complete_excel_files : un ID audité est toujours retrouvé par la mise à jour.

États d'une commune : **Complète** (toutes ses lignes validées), **À corriger**,
**Aucun ID erreur**, **Fichier introuvable** (pas d'`audit_*.xlsx`) et **Erreur de lecture**.
Une erreur de lecture vient d'un onglet « Audit » absent, d'un fichier illisible, ou d'une
colonne contrôlée absente : les cinq colonnes sont obligatoires. Une **copie de conflit**
OneDrive/SharePoint (deuxième `audit_*.xlsx`) est signalée dès l'audit, car la mise à
jour la refusera.

`⚙️ Colonnes auditées` règle les libellés des colonnes, les **valeurs acceptées** pour
Audit et Validation finale (ex. `OK, KO` ; `OK, Validé`), la ligne d'en-tête et la
colonne des ID erreur. La colonne ID et la ligne d'en-tête sont partagées avec les
règles de mise à jour.

## Rapport

Mêmes informations dans la popup et dans le HTML :

- indicateurs : communes (complètes / à corriger / en échec), ID erreur (non valides /
  validés), lignes validées et lignes non valides (avec pourcentages), barre de répartition ;
- **motifs de non-validité** : nombre de lignes par critère non respecté ;
- synthèse par département (HTML) ;
- tableau des communes : ID, lignes, validées, non valides, taux, état ; tri par
  colonne, recherche (commune ou ID erreur), filtre par état ;
- détail d'une commune : une ligne par ligne du fichier, avec ID erreur, n° de ligne,
  statut, Audit, Date audit, Auditeur, **Validation finale** et **motifs**.

**Afficher aussi les validées** : comme dans audit_ppts, le détail d'une commune ne
montre par défaut que ses lignes non valides. Cette case à cocher y ajoute les lignes
validées (✅, en vert). Elle est au-dessus du détail dans la popup, et dans la barre
d'outils des communes dans le HTML. La copie vise toujours les ID non valides, sauf la
copie des lignes sélectionnées dans le détail.

**Copie en un clic** (séparateur au choix : un par ligne, virgule,
point-virgule, espace) :

| Où | Quoi |
|---|---|
| HTML · bouton `📋 Copier (n)` d'une commune | ses ID non valides |
| HTML · clic sur un ID | cet ID |
| HTML · `📋 Copier les ID non valides affichés (n)` | ID non valides des communes **affichées** (filtre et recherche respectés) |
| HTML · section « Tous les ID erreur non valides » · `📋 Copier tout` | tous les ID non valides du rapport |
| Popup · `📋 Copier les ID non valides (n)` | ID non valides de la commune sélectionnée |
| Popup · `📋 Copier les ID non valides affichés (n)` | ID non valides des communes affichées |
| Popup · double-clic / Ctrl+C dans le détail | ID des lignes sélectionnées (une fois chacun) |

Le HTML est un fichier unique sans dépendance externe, lisible sans JavaScript,
en thème clair ou sombre (selon le système, bouton `◐ Thème` pour basculer),
imprimable (détails dépliés, toutes les lignes).

### Exports Excel et CSV

| Format | Contenu |
|---|---|
| `📊 Excel` (.xlsx) | onglets **Synthèse** (indicateurs, critères, motifs, départements), **Communes**, **Lignes** (une par ligne auditée : statut, 4 colonnes, motifs, fichier) et **ID non valides** (un ID par ligne, avec ses lignes fautives et ses motifs). En-têtes figés, filtres automatiques, statuts colorés. |
| `📑 CSV` | une ligne par ligne auditée (plus une par commune non auditée), mêmes colonnes que l'onglet « Lignes ». Séparateur `;` et UTF-8 avec BOM, comme audit_ppts : il s'ouvre directement dans Excel, accents compris. |

Les valeurs sont exportées telles que lues, sans réinterprétation.

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

La copie du moteur intègre les correctifs de complete_excel_files `50c8579`. Elle
y ajoute d'autres corrections, détaillées dans
[sna/PROVENANCE.md](sna/PROVENANCE.md) et à reporter dans complete_excel_files.

**Écriture via Excel sans blocage.** Une boîte de dialogue de fusion SharePoint
qu'aucun utilisateur ne peut valider ne gèle jamais l'outil. Si Excel ne donne
plus signe de vie pendant 90 s (180 s pour ouvrir un classeur), son processus
est tué, jamais celui de l'utilisateur. Une nouvelle instance est relancée et la
commune retentée. Le délai porte sur l'inactivité : une grosse commune qui
travaille normalement n'est jamais interrompue.

**Vérification fiable.** Chaque valeur est relue après l'enregistrement. L'ID
est retrouvé par sa valeur, ce qui résiste à un rechargement du classeur ou à
des lignes insérées par un collègue. Un ID non confirmé rejoint « ID en
attente ».

**Erreurs passagères retentées.** Une déconnexion d'Excel ou un classeur
rechargé pendant la fusion entraîne une nouvelle tentative.

**Doublons.** Toutes les lignes d'un ID en doublon sont mises à jour.

**Threads pour la rapidité.** L'audit lit plusieurs fichiers à la fois. La mise à
jour prépare plusieurs communes en parallèle : recherche du fichier, copies de
conflit et repérage rapide. En écriture openpyxl, elle les écrit aussi en
parallèle. Via Excel, une seule instance écrit les communes une à une pendant
que les suivantes se préparent. Le nombre de communes simultanées vaut 4 par
défaut ; il se règle par la clé `paralleles` de `config_audit_vf.json` (de 1 à
16). Le rapport et les résultats gardent l'ordre des communes.

**Arrêt.** `■ Arrêter` interrompt aussi la mise à jour entre deux communes. Une
commune commencée va à son terme, les suivantes ne sont pas touchées et ne
rejoignent pas « ID en attente ».

Seuls les ID **non valides** du rapport sont transmis (toutes leurs lignes sont réécrites).

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
│   ├── models/                  #   ColonnesAudit (+ valeurs acceptées) · LigneAudit, Anomalie, IdErreurAudit,
│   │                            #   AuditCommune, RapportAudit · SourceAudit (local / SharePoint)
│   ├── services/                #   logique métier, sans Tkinter
│   │   ├── detection.py         #     communes : détection sous la racine, saisie, vérification
│   │   ├── sharepoint.py        #     lien SharePoint → WebDAV ou bibliothèque synchronisée + Excel
│   │   ├── lecture_audit.py     #     lecture XML en flux des colonnes (repli openpyxl)
│   │   ├── controle.py          #     4 critères d'une ligne validée, dates en texte
│   │   ├── audit.py             #     orchestration (threads) → RapportAudit
│   │   ├── mise_a_jour.py       #     pont vers le moteur sna (règles, date du jour, session Excel)
│   │   └── export.py            #     écriture HTML / XLSX / CSV, ouverture
│   ├── views/                   #   rendu
│   │   ├── presentation.py      #     formats partagés (nombres, %, états)
│   │   ├── rapport_html.py      #     rapport HTML autonome
│   │   └── rapport_tableur.py   #     classeur Excel (4 onglets) et CSV
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

482 tests : détection, lecture (chemin rapide et repli openpyxl comparés),
critères de validité (dates en texte, casse, valeurs acceptées), audit par ligne et par
ID, rapport HTML (structure, échappement, données de copie par catégorie), exports
Excel et CSV, audit limité à des ID erreur saisis, source SharePoint (formats de lien, WebDAV, bibliothèque synchronisée,
lecture de la version serveur, connexion lente ou refusée, mise à jour forcée via Excel),
configuration,
mise à jour de bout en bout (valeurs écrites cellule par cellule, simulation,
conflit, règles personnalisées, identité avec le moteur appelé directement),
écriture via une **vraie instance Excel** (ignorée sans Excel) — garde-fou anti-blocage, redémarrage, relecture par ID, appels concurrents —, traitement et audit en parallèle, arrêt, et interface
Tkinter pilotée (détection des communes → audit → popup → « Afficher aussi les validées » →
copie → mise à jour →
ré-audit → exports HTML / Excel / CSV).

## Points d'attention

- Comme dans complete_excel_files, l'écriture openpyxl réécrit tout le
  classeur : testez sur une copie, ou préférez l'écriture via Excel.
- Le rapport HTML est une photographie : relancer l'audit pour un état à jour.
- Les règles par défaut **écrasent** `Audit` et `Auditeur` des ID ciblés, même
  s'ils sont déjà renseignés (comportement de complete_excel_files). Un ID non valide
  pour une autre raison qu'Audit (ex. `Audit = KO` mais `Date audit` vide) verrait
  donc son `KO` remplacé par `OK`. Pour le préserver, ajouter une condition
  « `Audit` vide » à la règle R1 dans le constructeur de règles.
