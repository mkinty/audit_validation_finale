"""Accès au dossier SharePoint par son API (REST), avec la session ouverte dans le navigateur.

Après l'identification dans le navigateur (``services.navigateur``), les cookies
de session SharePoint suffisent pour appeler l'API du site, exactement comme le
fait la page SharePoint elle-même :

    GET {site}/_api/web/GetFolderByServerRelativePath(decodedurl='…')/Folders
    GET {site}/_api/web/GetFolderByServerRelativePath(decodedurl='…')/Files
    GET {site}/_api/web/GetFileByServerRelativePath(decodedurl='…')/$value

C'est la version **serveur** des fichiers qui est lue : ni OneDrive, ni WebDAV,
ni Excel. Lecture seule : l'API n'est jamais utilisée pour écrire (la mise à
jour passe par Excel, en co-édition).
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
import re
import urllib.error
import urllib.request
from urllib.parse import quote, unquote, urlsplit

from sna.services.fichiers import EXTENSION, PREFIXE_AUDIT, PREFIXE_DEP

_MOTIF_DATE_NOM = re.compile(r"\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{4}")
_SITE = re.compile(r"^/(sites|teams|personal|portals)/[^/]+", re.I)


class AuthentificationRequise(PermissionError):
    """La session SharePoint est absente, expirée, ou n'ouvre pas ce dossier."""


def site_de(chemin_serveur: str) -> str:
    """Chemin du site qui contient un chemin serveur (``/sites/X/Docs/…`` → ``/sites/X``)."""
    correspondance = _SITE.match(chemin_serveur)
    return correspondance.group(0) if correspondance else ""


def chemin_serveur(url: str) -> str:
    """Adresse https d'un dossier ou d'un fichier → chemin serveur décodé."""
    return unquote(urlsplit(url).path).rstrip("/")


class ClientSharePoint:
    """Appels en lecture à l'API SharePoint, avec les cookies de la session navigateur."""

    def __init__(self, hote: str, cookies: dict[str, str], delai_s: float = 60, base: str | None = None):
        self.hote = hote.lower()
        self.cookies = dict(cookies)
        self.delai_s = delai_s
        self.base = (base or f"https://{self.hote}").rstrip("/")   # remplaçable en test

    # ── Requêtes ──────────────────────────────────────────────────
    def _requete(self, chemin: str, binaire: bool = False) -> bytes:
        requete = urllib.request.Request(f"{self.base}{chemin}", headers={
            "Cookie": "; ".join(f"{k}={v}" for k, v in self.cookies.items()),
            "Accept": "application/octet-stream" if binaire else "application/json;odata=nometadata",
            "User-Agent": "AuditValidationFinale",
        })
        try:
            with urllib.request.urlopen(requete, timeout=self.delai_s) as reponse:
                if urlsplit(reponse.geturl()).hostname != urlsplit(self.base).hostname:
                    raise AuthentificationRequise("session SharePoint absente ou expirée")
                return reponse.read()
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise AuthentificationRequise(
                    "session SharePoint absente ou expirée (ou droits insuffisants sur ce dossier)"
                ) from None
            if exc.code == 404:
                raise FileNotFoundError(f"introuvable sur SharePoint : {chemin}") from None
            raise OSError(f"SharePoint a répondu {exc.code} : {_message_erreur(exc)}") from None

    def _api(self, chemin_serveur: str, ressource: str, suite: str = "") -> str:
        valeur = quote(chemin_serveur.replace("'", "''"), safe="/")
        return f"{site_de(chemin_serveur)}/_api/web/{ressource}(decodedurl='{valeur}'){suite}"

    def _json(self, chemin: str) -> dict:
        return json.loads(self._requete(chemin).decode("utf-8"))

    # ── Dossiers et fichiers ──────────────────────────────────────
    def dossier_existe(self, chemin_serveur: str) -> bool:
        try:
            return bool(self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                              "?$select=Exists")).get("Exists"))
        except FileNotFoundError:
            return False

    def verifier(self, chemin_serveur: str) -> None:
        """Lève si la session n'ouvre pas ce dossier (ou s'il n'existe pas)."""
        if not self.dossier_existe(chemin_serveur):
            raise FileNotFoundError(f"dossier introuvable sur SharePoint : {chemin_serveur}")

    def dossiers(self, chemin_serveur: str) -> list[str]:
        donnees = self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                       "/Folders?$select=Name&$top=5000"))
        return sorted(d["Name"] for d in donnees.get("value", []))

    def fichiers(self, chemin_serveur: str) -> list[dict]:
        """``[{"Name", "TimeLastModified"}]`` des fichiers d'un dossier, triés par nom."""
        donnees = self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                       "/Files?$select=Name,TimeLastModified&$top=5000"))
        return sorted(donnees.get("value", []), key=lambda f: f["Name"])

    def dossiers_et_fichiers(self, chemin_serveur: str) -> dict[str, list[dict]]:
        """Sous-dossiers d'un dossier et leurs fichiers, en **une** requête."""
        donnees = self._json(self._api(
            chemin_serveur, "GetFolderByServerRelativePath",
            "/Folders?$select=Name,Files/Name,Files/TimeLastModified&$expand=Files&$top=5000",
        ))
        return {d["Name"]: sorted(d.get("Files", []), key=lambda f: f["Name"]) for d in donnees.get("value", [])}

    def telecharger(self, chemin_fichier: str, destination: str) -> str:
        contenu = self._requete(self._api(chemin_fichier, "GetFileByServerRelativePath", "/$value"),
                                binaire=True)
        with open(destination, "wb") as f:
            f.write(contenu)
        return destination

    def url(self, chemin_serveur: str) -> str:
        return f"https://{self.hote}{quote(chemin_serveur, safe='/')}"


# ── Règles de complete_excel_files, appliquées aux listes de l'API ─
def fichier_audit(fichiers: list[dict]) -> str | None:
    """Premier ``audit_*.xlsx`` par ordre de nom (même règle que ``trouver_fichier_audit``)."""
    return next((f["Name"] for f in fichiers if _est_audit(f["Name"])), None)


def copies_de_conflit(fichiers: list[dict]) -> list[tuple[str, str]]:
    """``[(nom, date)]`` des ``audit_*.xlsx`` surnuméraires (cf. ``detecter_fichiers_conflit``)."""
    principal = fichier_audit(fichiers)
    resultat = []
    for f in fichiers:
        if _est_audit(f["Name"]) and f["Name"] != principal:
            date = _MOTIF_DATE_NOM.search(f["Name"])
            resultat.append((f["Name"], date.group(0) if date else _date(f.get("TimeLastModified"))))
    return resultat


def lister_communes(client: ClientSharePoint, racine: str, paralleles: int = 8):
    """``[(code, département, a_fichier)]`` de toutes les communes ``<racine>/Dep*/<commune>``.

    Une requête par dossier ``Dep<XX>`` (sous-dossiers et leurs fichiers ensemble),
    en parallèle.
    """
    departements = [d for d in client.dossiers(racine) if d.lower().startswith(PREFIXE_DEP.lower())]
    with ThreadPoolExecutor(max_workers=max(1, min(paralleles, len(departements) or 1))) as pool:
        contenus = list(pool.map(lambda d: client.dossiers_et_fichiers(f"{racine}/{d}"), departements))
    communes = []
    for contenu in contenus:
        for code, fichiers in contenu.items():
            communes.append((code.upper(), code[:2].upper(), fichier_audit(fichiers) is not None))
    return sorted(set(communes))


def _est_audit(nom: str) -> bool:
    return nom.startswith(PREFIXE_AUDIT) and nom.endswith(EXTENSION) and not nom.startswith("~$")


def _date(horodatage: str | None) -> str:
    try:
        return datetime.fromisoformat(str(horodatage).replace("Z", "+00:00")).astimezone().strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return "date inconnue"


def _message_erreur(exc: urllib.error.HTTPError) -> str:
    try:
        donnees = json.loads(exc.read().decode("utf-8"))
        return donnees.get("odata.error", {}).get("message", {}).get("value") or str(donnees)[:200]
    except Exception:  # noqa: BLE001
        return exc.reason or ""
