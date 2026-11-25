"""Correspondance dossier local synchronisé par OneDrive ↔ adresse SharePoint.

OneDrive enregistre chaque bibliothèque synchronisée dans
``HKCU\\Software\\SyncEngines\\Providers\\OneDrive\\<id>`` (valeurs
``MountPoint`` et ``UrlNamespace``). Ouvrir un classeur par son adresse
SharePoint plutôt que par son chemin local garantit qu'Excel rejoint la
session de co-édition au lieu de travailler sur la seule copie locale —
c'est ce qui évite le conflit de fusion lorsque quelqu'un d'autre a le
fichier ouvert (cf. ``services.excel_com``).
"""

from __future__ import annotations

import os
import re
from urllib.parse import quote

_CLE = r"Software\SyncEngines\Providers\OneDrive"


def montages_onedrive() -> list[tuple[str, str]]:
    """``[(dossier local, URL SharePoint), ...]`` des bibliothèques synchronisées.

    Liste vide hors Windows, ou si OneDrive n'est pas configuré.
    """
    try:
        import winreg
    except ImportError:
        return []
    resultat = []
    try:
        racine = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _CLE)
    except OSError:
        return resultat
    with racine:
        i = 0
        while True:
            try:
                sous = winreg.EnumKey(racine, i)
            except OSError:
                break
            i += 1
            try:
                with winreg.OpenKey(racine, sous) as cle:
                    point_montage = winreg.QueryValueEx(cle, "MountPoint")[0]
                    url = winreg.QueryValueEx(cle, "UrlNamespace")[0]
            except OSError:
                continue
            if point_montage and url:
                resultat.append((point_montage, url))
    return resultat


_WEBDAV = re.compile(r"^\\\\(?P<hote>[^\\@]+)@SSL(?:@\d+)?\\DavWWWRoot(?P<chemin>\\.*)?$", re.I)


def url_sharepoint(chemin: str, montages: list[tuple[str, str]] | None = None) -> str | None:
    """Adresse SharePoint d'un fichier situé dans un dossier synchronisé, sinon ``None``.

    Un chemin WebDAV (``\\\\hôte@SSL\\DavWWWRoot\\sites\\…``, accès direct au
    serveur — ajout de la copie vendorisée, cf. PROVENANCE.md) est converti en
    son adresse https, pour qu'Excel l'ouvre en co-édition.
    """
    webdav = _WEBDAV.match(chemin or "")
    if webdav:
        relatif = (webdav.group("chemin") or "").replace("\\", "/")
        return f"https://{webdav.group('hote').lower()}{quote(relatif, safe='/')}"
    montages = montages_onedrive() if montages is None else montages
    absolu = os.path.abspath(chemin)
    p = os.path.normcase(absolu)
    meilleur = None
    for point_montage, url in montages:
        m = os.path.normcase(os.path.abspath(point_montage)).rstrip("\\/")
        if (p == m or p.startswith(m + os.sep)) and (meilleur is None or len(m) > len(meilleur[0])):
            meilleur = (m, url)
    if meilleur is None:
        return None
    relatif = absolu[len(meilleur[0]):].lstrip("\\/").replace("\\", "/")
    return meilleur[1].rstrip("/") + "/" + quote(relatif, safe="/")
