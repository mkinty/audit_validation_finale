"""Réglages de l'outil, conservés d'un lancement à l'autre (``config_audit_vf.json``).

    source          → « local » (dossier racine) ou « sharepoint » (lien SharePoint)
    racine          → dossier racine des audits (source locale)
    lien_sharepoint → lien SharePoint du dossier racine (source SharePoint)
    communes        → dernière sélection de communes
    ids_cibles      → ID saisis pour un audit restreint (vide : communes entières)
    type_ids        → nature de ces ID : « id_erreur » ou « id_imm »
    colonnes        → libellés des colonnes auditées (Validation finale, Audit…)
    regles          → règles de mise à jour, au format de complete_excel_files
                      (``sna.models.regles.ReglesConfig``) : valeurs par défaut
                      identiques, modifiables dans le même constructeur de règles
    date_du_jour    → la règle « Date audit » prend la date du jour à chaque mise à jour
    ecrire_via_excel→ écriture via Excel (co-édition SharePoint) si disponible
    reauditer       → ré-auditer les communes mises à jour pour rafraîchir le rapport
    dossier_export  → dernier dossier d'export du rapport HTML
    paralleles      → communes lues / préparées en même temps (threads, 1 à 16)

Un fichier absent ou illisible redonne les valeurs par défaut ; une clé
manquante est complétée par sa valeur par défaut.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import os

from sna.models.regles import ReglesConfig

from .chemins import CONFIG_FILE
from .models.colonnes import ColonnesAudit

PARALLELES_DEFAUT = 4
PARALLELES_MAX = 16


@dataclass
class Parametres:
    source: str = "local"
    racine: str = ""
    lien_sharepoint: str = ""
    communes: list[str] = field(default_factory=list)
    ids_cibles: list[str] = field(default_factory=list)
    type_ids: str = "id_erreur"
    colonnes: ColonnesAudit = field(default_factory=ColonnesAudit)
    regles: ReglesConfig = field(default_factory=ReglesConfig.defaut)
    date_du_jour: bool = True
    ecrire_via_excel: bool = True
    reauditer: bool = True
    dossier_export: str = ""
    paralleles: int = PARALLELES_DEFAUT

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "racine": self.racine,
            "lien_sharepoint": self.lien_sharepoint,
            "communes": list(self.communes),
            "ids_cibles": list(self.ids_cibles),
            "type_ids": self.type_ids,
            "colonnes": self.colonnes.to_dict(),
            "regles": self.regles.to_dict(),
            "date_du_jour": self.date_du_jour,
            "ecrire_via_excel": self.ecrire_via_excel,
            "reauditer": self.reauditer,
            "dossier_export": self.dossier_export,
            "paralleles": self.paralleles,
        }

    @staticmethod
    def from_dict(data: dict | None) -> "Parametres":
        data = data if isinstance(data, dict) else {}
        defaut = Parametres()
        parametres = Parametres(
            source="sharepoint" if data.get("source") == "sharepoint" else "local",
            racine=str(data.get("racine") or ""),
            lien_sharepoint=str(data.get("lien_sharepoint") or "").strip(),
            communes=_liste(data.get("communes")),
            ids_cibles=_liste(data.get("ids_cibles")),
            type_ids="id_imm" if data.get("type_ids") == "id_imm" else "id_erreur",
            colonnes=ColonnesAudit.from_dict(data.get("colonnes")),
            regles=_regles(data.get("regles")),
            date_du_jour=bool(data.get("date_du_jour", defaut.date_du_jour)),
            ecrire_via_excel=bool(data.get("ecrire_via_excel", defaut.ecrire_via_excel)),
            reauditer=bool(data.get("reauditer", defaut.reauditer)),
            dossier_export=str(data.get("dossier_export") or ""),
            paralleles=_paralleles(data.get("paralleles", defaut.paralleles)),
        )
        # Identifiant des mises à jour : l'ID imm (unique par ligne), plus l'ID erreur.
        # Ancienne configuration : la colonne ID erreur personnalisée était celle des règles.
        anciennes = data.get("colonnes") if isinstance(data.get("colonnes"), dict) else {}
        ancien_id = parametres.regles.colonne_id
        if "id_erreur" not in anciennes and ancien_id and ancien_id != parametres.colonnes.id_imm:
            parametres.colonnes.id_erreur = ancien_id
        parametres.regles.colonne_id = parametres.colonnes.id_imm
        return parametres


def charger(fichier: str = CONFIG_FILE) -> Parametres:
    """Réglages enregistrés, ou valeurs par défaut."""
    try:
        with open(fichier, encoding="utf-8") as f:
            return Parametres.from_dict(json.load(f))
    except (OSError, ValueError):
        return Parametres()


def enregistrer(parametres: Parametres, fichier: str = CONFIG_FILE) -> bool:
    """Enregistre (écriture atomique). Retourne ``False`` en cas d'échec."""
    tmp = fichier + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(parametres.to_dict(), f, indent=2, ensure_ascii=False)
        os.replace(tmp, fichier)
        return True
    except OSError:
        return False


def _liste(valeur) -> list[str]:
    if not isinstance(valeur, list):
        return []
    return [str(v) for v in valeur if str(v).strip()]


def _paralleles(valeur) -> int:
    try:
        return min(max(int(valeur), 1), PARALLELES_MAX)
    except (TypeError, ValueError):
        return PARALLELES_DEFAUT


def _regles(data) -> ReglesConfig:
    """Mêmes garanties que ``sna.config.charger_regles`` : défauts si absent ou corrompu."""
    if not data:
        return ReglesConfig.defaut()
    try:
        return ReglesConfig.from_dict(data)
    except Exception:  # noqa: BLE001 — données corrompues → règles par défaut
        return ReglesConfig.defaut()
