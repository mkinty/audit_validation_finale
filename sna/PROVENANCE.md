# Provenance du package `sna`

Copie du moteur de l'outil **complete_excel_files** (`../complete_excel_files/sna`,
commit `50c8579`), reprise pour que la mise à jour des ID erreur suive exactement
le même processus : règles (`models/regles.py`, valeurs par défaut comprises),
traitement par commune (`services/traitement.py`), écriture openpyxl
(`services/excel.py`) ou via Excel (`services/excel_com.py`), liste des ID en
attente (`services/attente.py`) et constructeur de règles (`ui/panneau_regles.py`).

Les correctifs de `50c8579` sont repris : la vérification ne dépend plus de
`classeur.Saved`, un fil dédié possède Excel avec un garde-fou anti-blocage,
`AutoSaveOn` et `Interactive` sont désactivés, et le processus tué est attendu
jusqu'à sa disparition.

Ses tests d'origine sont dans `tests/moteur_sna/` (`test_excel_com.py` = celui de `50c8579`).

## Différences avec l'original

`ui/application.py` (fenêtre principale de complete_excel_files) n'est pas repris.
Les autres écarts sont des **correctifs et améliorations, à reporter dans
complete_excel_files**. Chacun est couvert par
`tests/moteur_sna/test_excel_com_robustesse.py` (vrai Excel) ou
`tests/moteur_sna/test_traitement_parallele.py`.

### `services/excel_com.py`

| # | Problème dans `50c8579` | Correction |
|---|---|---|
| 1 | Un ID erreur présent sur plusieurs lignes : seule la **première ligne** est écrite via Excel (`_index_lignes` ne garde qu'une occurrence), alors que l'écriture openpyxl les traite toutes. | `_index_lignes` renvoie toutes les lignes ; chacune est traitée. |
| 2 | Le garde-fou mesure la **durée totale** du classeur (90 s), alors que `_attendre_pret` peut attendre 120 s légitimement : une grosse commune qui travaille normalement est tuée à chaque tentative et n'aboutit jamais. | Délai d'**inactivité** : la tâche donne des signes de vie (`_FilExcel.signe`) à chaque attente, ligne écrite ou relue. L'ouverture a son propre délai (`delai_ouverture_s`, 180 s). |
| 3 | Le délai démarre au **dépôt** de la tâche : avec plusieurs threads, l'attente de son tour compte comme un blocage, et Excel est tué au milieu de l'écriture d'une autre commune. | `appliquer_regles` est protégée par un verrou ; le délai démarre quand la tâche s'exécute. |
| 4 | La relecture après `Save()` réutilise les objets `Range` d'avant : si SharePoint recharge le classeur pendant la fusion, la relecture lève 0x800A01A8 et **toute la commune passe en erreur**, alors que l'enregistrement a réussi. Si un collègue insère des lignes, la mauvaise ligne est relue. | Relecture à partir d'objets neufs, ID retrouvés **par leur valeur** (colonnes par leur en-tête). Si la relecture est impossible, les ID sont « non confirmés » (remis en attente), sans échec de la commune. |
| 5 | La tentative abandonnée utilise `self.xl`, qui désigne alors la **nouvelle** instance appelée depuis le mauvais fil ; ses messages se mêlent au journal de la tentative suivante. | Chaque tâche reçoit son propre fil (`_appliquer_une_fois(fil, …)`) ; les messages d'un fil abandonné sont ignorés. |
| 6 | Un Excel qui démarre au-delà de 30 s reste **orphelin** (invisible). Un redémarrage qui échoue laisse `_fil = None` : toutes les communes suivantes échouent sur `'NoneType' object has no attribute 'executer'`. | Une pilule est déposée pour refermer l'instance tardive. La commune suivante relance Excel si la session n'en a plus. |
| 7 | Erreurs COM transitoires (« déconnecté de ses clients » 0x80010108, serveur RPC indisponible, objet invalidé 0x800A01A8) : la commune échoue directement. | Déconnexion : Excel est remplacé et le classeur retenté. Objet invalidé : le classeur est rouvert et retenté. |
| 8 | PID introuvable par la fenêtre : le processus bloqué ne peut pas être tué. | Repli par différence des processus `EXCEL.EXE` avant/après le démarrage, et signalement au journal sinon. |
| 9 | Lecture et vérification cellule par cellule (2 à 3 appels COM par référence de règle et par ligne). | Lecture **groupée par colonne** (`_CacheColonnes` : un appel par colonne), les écritures étant reportées dans le cache. Même résultat, beaucoup moins d'allers-retours avec Excel. |
| 10 | Une exception levée sur le fil Excel garde dans sa trace les variables locales, dont `classeur`, `feuille` et les cellules. Elle forme un cycle que le ramasse-miettes libère plus tard **depuis n'importe quel thread**, parfois le fil Excel neuf en pleine écriture. Les appels en cours sont alors perturbés : feuille « à moitié initialisée », `AttributeError: <unknown>.Cells`, « Invalid number of arguments ». Le problème est reproductible en enchaînant les tests de `tests/moteur_sna`. | `_liberer_objets_com` vide ces cadres sur le fil Excel lui-même, avant de transmettre l'exception. Le type et le message sont conservés. |
| 11 | *(ajout, audit depuis SharePoint)* | `SessionExcel.copier_version_serveur(adresse, destination)` ouvre un classeur par son URL SharePoint en lecture seule, puis l'enregistre en copie locale (`SaveCopyAs`). Il passe par les mêmes reprises et le même garde-fou que `appliquer_regles` : la boucle de reprises est factorisée dans `_avec_reprises`. |

| 12 | *(ajout)* L'ouverture d'un fichier SharePoint peut afficher la **fenêtre de connexion Microsoft 365** (classe `BasicEmbeddedBrowser`). Dans une instance invisible, l'appel reste bloqué jusqu'au délai anti-blocage, trois fois de suite : 6 minutes pour un fichier, constatées sur un vrai SharePoint. | Le garde-fou détecte cette fenêtre en moins de 2 s et lève `ConnexionRequise`. Avec `SessionExcel(connexion_interactive=True)`, le fichier est ouvert une fois dans un Excel **visible** pour que l'utilisateur se connecte, puis le travail reprend. Sans cette option, un message explique quoi faire. Une adresse `https://` est ouverte en mode « co-édition SharePoint ». |

### `services/onedrive.py`

- *(ajout, audit depuis SharePoint)* `url_sharepoint` convertit aussi un chemin WebDAV
  `\\hôte@SSL\DavWWWRoot\…` en adresse https, pour qu'Excel l'ouvre en co-édition.

### `services/traitement.py`

- `paralleles` (par défaut 1 : séquentiel, identique à l'original) traite
  plusieurs communes à la fois sur un pool de threads.
  - Les résultats gardent l'ordre demandé.
  - Les callbacks de progression et de résultat ne sont jamais appelés en même temps.
  - Via Excel, la préparation (fichier, conflits, repérage rapide) se fait en
    parallèle, tandis que l'écriture reste une à une dans la session unique.
- `doit_arreter` (par défaut `None`) permet d'interrompre un lot entre deux
  communes. Une commune commencée va à son terme.
- `localiser` et `conflits` (par défaut `trouver_fichier_audit` et
  `detecter_fichiers_conflit`) : un fichier peut être désigné par son adresse
  SharePoint quand le dossier ne peut pas être parcouru. Ces appels peuvent passer par
  le réseau : leur échec met la commune en erreur, sans interrompre le lot.
- `prereperage` (par défaut `True` : comportement d'origine). Avec `False`, Excel est
  toujours ouvert, sans vérifier d'abord la présence des ID dans le fichier parcouru.
  C'est nécessaire quand ce fichier est une copie locale qui peut différer du serveur.

### `services/attente.py`

- Un ID « non confirmé » saisi sous une autre forme (`2a004_1` pour `2A004_1`)
  était compté comme **réussi** : la comparaison se fait maintenant sur la forme
  normalisée.

## Resynchroniser

Recopier `../complete_excel_files/sna` (sauf `ui/application.py`), réappliquer
les différences ci-dessus qui n'auraient pas été reportées, puis lancer
`python -m pytest`.
