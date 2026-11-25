"""D'où viennent les fichiers audités : dossier local, ou SharePoint (cf. ``services.sharepoint``)."""

from __future__ import annotations

from dataclasses import dataclass

LOCAL, WEBDAV, EXCEL = "local", "webdav", "excel"

_LIBELLES = {
    LOCAL: "Dossier local",
    WEBDAV: "SharePoint — accès direct (WebDAV)",
    EXCEL: "SharePoint — via Excel (bibliothèque synchronisée)",
}


@dataclass(frozen=True)
class SourceAudit:
    """Source d'un audit.

    ``racine`` est le dossier **parcouru** pour trouver les communes : le dossier
    local, le chemin WebDAV du dossier SharePoint, ou (mode ``excel``) le dossier
    synchronisé qui ne sert qu'à lister — les fichiers, eux, sont lus sur le
    serveur. ``lien`` est l'adresse SharePoint du dossier racine.
    """

    mode: str
    racine: str
    lien: str = ""

    @property
    def sharepoint(self) -> bool:
        return self.mode != LOCAL

    @property
    def libelle(self) -> str:
        return _LIBELLES.get(self.mode, self.mode)

    @property
    def affichage(self) -> str:
        """Ce que l'on montre comme « racine » : le lien SharePoint, sinon le dossier."""
        return self.lien or self.racine
