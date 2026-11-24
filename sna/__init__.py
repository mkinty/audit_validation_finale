"""Projet SNA — Modification des fichiers audit Excel par ID erreur.

L'utilisateur saisit des identifiants ``<insee>_<numéro>`` ; le programme en
déduit les communes concernées, ouvre le fichier audit de chacune, et impose
aux lignes correspondantes les valeurs définies dans les règles.

Le package est volontairement découpé en couches indépendantes :

    sna.chemins           → emplacement du fichier de configuration
    sna.config            → chargement/sauvegarde (dossier racine + règles)
    sna.models            → structures de données (règles, résultat, statut)
    sna.services          → logique métier, sans aucune dépendance à l'interface
    sna.ui                → interface graphique Tkinter (thème, panneau, fenêtre)

Aucun import de l'interface n'est fait ici afin que la couche métier reste
importable et testable sans environnement graphique.
"""

__version__ = "3.0.0"
