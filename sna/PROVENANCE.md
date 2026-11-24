# Provenance du package `sna`

Copie du moteur de l'outil **complete_excel_files** (`../complete_excel_files/sna`,
commit `360f11f`), reprise pour que la mise à jour des ID erreur suive exactement
le même processus : règles (`models/regles.py`, valeurs par défaut comprises),
traitement par commune (`services/traitement.py`), écriture openpyxl
(`services/excel.py`) ou via Excel (`services/excel_com.py`), liste des ID en
attente (`services/attente.py`) et constructeur de règles (`ui/panneau_regles.py`).

Ses tests d'origine sont dans `tests/moteur_sna/`.

## Différences avec l'original

- `ui/application.py` (fenêtre principale de complete_excel_files) n'est pas repris.
- **Correctif** dans `services/excel_com.py` : l'écriture via Excel ne traitait que
  la **première ligne** d'un ID erreur présent sur plusieurs lignes
  (`_index_lignes` ne gardait que la première occurrence), alors que l'écriture
  openpyxl — et le README de complete_excel_files — modifient toutes les lignes.
  Ici, toutes les lignes sont traitées. Le même correctif est à reporter dans
  complete_excel_files.

Pour resynchroniser : recopier `../complete_excel_files/sna` (sauf
`ui/application.py`), réappliquer le correctif ci-dessus s'il n'a pas été
reporté, puis lancer `python -m pytest`.
