"""Audit de la « Validation finale » des fichiers audit SNA.

Arborescence auditée :

    <racine>/Dep<XX>/<commune>/audit_*.xlsx

Le package est découpé en couches indépendantes :

    avf.chemins    → emplacement des fichiers de configuration et de données
    avf.config     → chargement / sauvegarde des réglages de l'outil
    avf.models     → structures de données (colonnes auditées, résultats, rapport)
    avf.services   → logique métier, sans aucune dépendance à l'interface
    avf.views      → rendu du rapport (HTML autonome, résumés texte)
    avf.ui         → interface graphique Tkinter

La mise à jour des ID erreur délègue intégralement au moteur ``sna`` repris de
l'outil ``complete_excel_files`` (mêmes règles, même traitement, même écriture
openpyxl ou via Excel) : cf. ``avf.services.mise_a_jour``.
"""

__version__ = "1.0.0"
