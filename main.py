"""
Point d'entrée — Audit de la Validation finale (Projet SNA)

Audite les fichiers ``<racine>/Dep<XX>/<commune>/audit_*.xlsx`` : nombre d'ID
erreur avec / sans « Validation finale », rapport en popup ou en HTML, puis
mise à jour des ID sans validation avec le moteur de complete_excel_files.

Usage :
    python main.py
"""

from avf.ui.application import Application


def main() -> None:
    Application().mainloop()


if __name__ == "__main__":
    main()
