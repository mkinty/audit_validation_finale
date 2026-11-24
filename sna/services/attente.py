"""ID erreur non appliqués, conservés d'un traitement à l'autre pour être rejoués.

Quatre façons, pour un ID erreur demandé, de ne pas aboutir :

    - le fichier audit de sa commune est introuvable ;
    - son traitement a levé une erreur (colonne absente, onglet absent,
      classeur verrouillé par un autre utilisateur…) ;
    - le fichier a bien été lu, mais cet ID erreur n'y figure pas ;
    - (écriture via Excel uniquement, cf. ``services.excel_com``) la cellule a
      été écrite localement, mais l'enregistrement vers SharePoint n'a pas pu
      être confirmé par relecture.

Plutôt que de laisser l'utilisateur recopier ces identifiants à la main depuis
le journal pour les rejouer plus tard, ils sont conservés dans un petit fichier
JSON (:data:`sna.chemins.IDS_EN_ATTENTE_FILE`) :

    {"13001_1": {"raison": "fichier audit introuvable", "date": "12/03/2026 09:41"}, …}

Chaque nouveau traitement retire de cette liste les ID qui ont cette fois été
trouvés, et y ajoute (ou met à jour) ceux qui ont encore échoué.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from datetime import datetime
import json
import os

from ..models.base import ResultatCommune, Statut


def charger(fichier: str) -> dict[str, dict]:
    """Liste actuelle des ID en attente, ou vide si le fichier est absent/illisible."""
    try:
        with open(fichier, encoding="utf-8") as f:
            donnees = json.load(f)
        return donnees if isinstance(donnees, dict) else {}
    except (OSError, ValueError):
        return {}


def mettre_a_jour(
    fichier: str, echecs: Mapping[str, str], reussis: Iterable[str],
) -> dict[str, dict]:
    """Retire les ID réussis, ajoute ou remplace les échecs (avec leur raison).

    Retourne la liste à jour, telle qu'enregistrée.
    """
    donnees = charger(fichier)
    for identifiant in reussis:
        donnees.pop(identifiant, None)

    date = datetime.now().strftime("%d/%m/%Y %H:%M")
    for identifiant, raison in echecs.items():
        donnees[identifiant] = {"raison": raison, "date": date}

    _ecrire(fichier, donnees)
    return donnees


def vider(fichier: str) -> None:
    """Efface la liste des ID en attente (ne touche à aucun fichier audit)."""
    _ecrire(fichier, {})


def calculer_echecs_et_reussites(
    groupes: Mapping[str, Iterable[str]], resultats: Iterable[ResultatCommune],
) -> tuple[dict[str, str], list[str]]:
    """Répartit les ID erreur d'un traitement entre échecs (avec raison) et réussites.

    Un échec couvre les quatre cas documentés en tête de module. Les autres ID
    sont réussis — trouvés et traités, que la règle ait ou non changé la
    cellule (déjà à jour, par exemple).
    """
    echecs: dict[str, str] = {}
    reussis: list[str] = []
    for resultat in resultats:
        ids = list(groupes.get(resultat.commune, []))
        if resultat.statut is Statut.INTROUVABLE:
            for identifiant in ids:
                echecs[identifiant] = "fichier audit introuvable"
        elif resultat.statut is Statut.ERREUR:
            for identifiant in ids:
                echecs[identifiant] = resultat.message or "erreur inconnue"
        else:
            for identifiant in ids:
                if identifiant in resultat.ids_absents:
                    echecs[identifiant] = "ID erreur introuvable dans le fichier"
                elif identifiant in resultat.ids_non_verifies:
                    echecs[identifiant] = "écriture non confirmée dans SharePoint après enregistrement"
                else:
                    reussis.append(identifiant)
    return echecs, reussis


def _ecrire(fichier: str, donnees: dict) -> None:
    """Écriture atomique (fichier temporaire + remplacement) pour ne jamais
    laisser un fichier à moitié écrit si le programme est interrompu."""
    tmp = fichier + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(donnees, f, indent=2, ensure_ascii=False)
    os.replace(tmp, fichier)
