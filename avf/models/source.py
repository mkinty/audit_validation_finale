"""D'où viennent les fichiers audités : dossier local, ou SharePoint (cf. ``services.sharepoint``)."""

from __future__ import annotations

from dataclasses import dataclass

LOCAL, WEBDAV, EXCEL, URL, API = "local", "webdav", "excel", "url", "api"

_LIBELLES = {
    LOCAL: "Dossier local",
    API: "SharePoint — connexion par le navigateur",
    WEBDAV: "SharePoint — accès direct (WebDAV)",
    EXCEL: "SharePoint — via Excel (bibliothèque synchronisée)",
    URL: "SharePoint — via Excel (sans synchronisation)",
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
    diagnostic: str = ""   # pourquoi l'accès direct (WebDAV) n'a pas été retenu

    @property
    def sharepoint(self) -> bool:
        return self.mode != LOCAL

    @property
    def listable(self) -> bool:
        """Faux en mode ``url`` : le dossier ne peut pas être parcouru, les communes
        (ou ID erreur) doivent être saisies ; leurs fichiers sont ouverts par adresse."""
        return self.mode != URL

    def url_fichier(self, commune: str) -> str:
        """Mode ``url`` : adresse du fichier audit d'une commune
        (``<lien>/Dep<XX>/<commune>/audit_<commune>.xlsx``)."""
        code = commune.strip().upper()
        return f"{self.lien.rstrip('/')}/Dep{code[:2]}/{code}/audit_{code}.xlsx"

    @property
    def libelle(self) -> str:
        return _LIBELLES.get(self.mode, self.mode)

    @property
    def affichage(self) -> str:
        """Ce que l'on montre comme « racine » : le lien SharePoint, sinon le dossier."""
        return self.lien or self.racine
