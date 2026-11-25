"""Connexion par le navigateur + API SharePoint, contre un faux serveur SharePoint local.

Le faux serveur expose, sur l'arborescence de test, les mêmes appels que
SharePoint Online (``GetFolderByServerRelativePath``, ``GetFileByServerRelativePath``)
et refuse toute requête sans le cookie de session ``FedAuth``.
"""

import http.server
import json
import os
import re
import threading
from datetime import datetime, timezone
from urllib.parse import unquote, urlsplit

import pytest

from avf.models.audit import StatutAudit
from avf.models.source import API, SourceAudit
from avf.services import sharepoint as sp
from avf.services import sharepoint_rest as rest
from avf.services.audit import auditer_communes
from avf.services.detection import lister_communes
from avf.services.mise_a_jour import mettre_a_jour
from avf.services.navigateur import NavigateurFerme, cookies_pour
from sna.models.regles import ReglesConfig

HOTE = "contoso.sharepoint.com"
DOSSIER = "/sites/SNA/Shared Documents/Audit"
LIEN = "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Audit"
_APPEL = re.compile(r"^(?P<site>/sites/SNA)?/_api/web/(?P<res>GetFolderByServerRelativePath|"
                    r"GetFileByServerRelativePath)\(decodedurl='(?P<chemin>.*)'\)(?P<suite>/[^?]*)?(?:\?(?P<q>.*))?$")


class _FauxSharePoint(http.server.BaseHTTPRequestHandler):
    racine = ""
    appels: list = []
    versions: dict = {}            # chemin local → numéro de version (ETag)
    verrouilles: set = set()       # chemins locaux verrouillés (423 à l'écriture)
    extraits: dict = {}            # chemins locaux extraits → par qui
    avant_ecriture = None          # rappel(chemin local) juste avant d'accepter une écriture
    ignorer_ecritures = False      # simule un serveur qui « perd » l'écriture

    def _etag(self, local):
        return f'"{{AAAA-BBBB}},{type(self).versions.get(local, 1)}"'

    def log_message(self, *a):
        pass

    def _local(self, chemin):
        if not chemin.startswith(DOSSIER):
            return None
        return os.path.join(self.racine, *[p for p in chemin[len(DOSSIER):].split("/") if p])

    def _envoyer(self, code, corps, type_="application/json", entetes=None):
        donnees = corps if isinstance(corps, bytes) else json.dumps(corps).encode()
        self.send_response(code)
        for k, v in (entetes or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Type", type_)
        self.send_header("Content-Length", str(len(donnees)))
        self.end_headers()
        self.wfile.write(donnees)

    def do_GET(self):
        type(self).appels.append(unquote(self.path))
        if "FedAuth=valide" not in (self.headers.get("Cookie") or ""):
            return self._envoyer(403, {"odata.error": {"message": {"value": "Accès refusé"}}})
        m = _APPEL.match(urlsplit(self.path).path + ("?" + urlsplit(self.path).query if urlsplit(self.path).query else ""))
        if not m or m.group("site") != "/sites/SNA":
            return self._envoyer(400, {"odata.error": {"message": {"value": "requête inattendue"}}})
        chemin = unquote(m.group("chemin")).replace("''", "'")
        local, suite, q = self._local(chemin), m.group("suite") or "", unquote(m.group("q") or "")
        if m.group("res") == "GetFileByServerRelativePath":
            if not (local and os.path.isfile(local)):
                return self._envoyer(404, {"odata.error": {"message": {"value": "fichier introuvable"}}})
            if suite == "/$value":
                return self._envoyer(200, open(local, "rb").read(), "application/octet-stream",
                                     {"ETag": self._etag(local)})
            if suite == "" and "ETag" in q:
                qui = type(self).extraits.get(local)
                return self._envoyer(200, {"ETag": self._etag(local), "CheckOutType": 0 if qui else 2,
                                           "CheckedOutByUser": {"Title": qui} if qui else None})
            return self._envoyer(400, {})
        if not local or not os.path.isdir(local):
            if suite == "" and "Exists" in q:
                return self._envoyer(404, {"odata.error": {"message": {"value": "introuvable"}}})
            return self._envoyer(404, {"odata.error": {"message": {"value": "introuvable"}}})
        if suite == "":
            return self._envoyer(200, {"Exists": True})
        if suite == "/Files":
            return self._envoyer(200, {"value": _fichiers(local)})
        if suite == "/Folders":
            dossiers = [n for n in sorted(os.listdir(local)) if os.path.isdir(os.path.join(local, n))]
            if "$expand=Files" in q:
                return self._envoyer(200, {"value": [
                    {"Name": n, "Files": _fichiers(os.path.join(local, n))} for n in dossiers]})
            return self._envoyer(200, {"value": [{"Name": n} for n in dossiers]})
        return self._envoyer(400, {})

    def do_POST(self):
        cls = type(self)
        cls.appels.append("POST " + unquote(self.path))
        corps = self.rfile.read(int(self.headers.get("Content-Length") or 0))
        if "FedAuth=valide" not in (self.headers.get("Cookie") or ""):
            return self._envoyer(403, {"odata.error": {"message": {"value": "Accès refusé"}}})
        if self.path == "/sites/SNA/_api/contextinfo":
            return self._envoyer(200, {"FormDigestValue": "jeton-valide"})
        m = _APPEL.match(self.path)
        if not m or m.group("res") != "GetFileByServerRelativePath" or m.group("suite") != "/$value" \
                or self.headers.get("X-HTTP-Method") != "PUT":
            return self._envoyer(400, {})
        if self.headers.get("X-RequestDigest") != "jeton-valide":
            return self._envoyer(403, {"odata.error": {"message": {"value": "jeton de formulaire invalide"}}})
        local = self._local(unquote(m.group("chemin")).replace("''", "'"))
        if cls.avant_ecriture:
            cls.avant_ecriture(local)
        if local in cls.verrouilles:
            return self._envoyer(423, {"odata.error": {"message": {
                "value": "The file is locked for shared use by Collègue"}}})
        if self.headers.get("If-Match") not in ("*", self._etag(local)):
            return self._envoyer(412, {"odata.error": {"message": {"value": "version modifiée"}}})
        if not cls.ignorer_ecritures:
            with open(local, "wb") as f:
                f.write(corps)
        cls.versions[local] = cls.versions.get(local, 1) + 1
        return self._envoyer(204, b"")


def _fichiers(dossier):
    return [{"Name": n, "TimeLastModified": datetime.fromtimestamp(
        os.path.getmtime(os.path.join(dossier, n)), timezone.utc).isoformat().replace("+00:00", "Z")}
        for n in sorted(os.listdir(dossier)) if os.path.isfile(os.path.join(dossier, n))]


@pytest.fixture
def serveur(racine):
    gestionnaire = type("Faux", (_FauxSharePoint,), {"racine": racine, "appels": [], "versions": {},
                                                     "verrouilles": set(), "extraits": {}})
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), gestionnaire)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield srv, gestionnaire
    srv.shutdown()


@pytest.fixture
def client(serveur):
    srv, _ = serveur
    return rest.ClientSharePoint(HOTE, {"FedAuth": "valide", "rtFa": "x"}, base=f"http://127.0.0.1:{srv.server_port}")


@pytest.fixture
def source_api(client):
    sp.enregistrer_client(client)
    yield SourceAudit(API, LIEN, LIEN)
    sp._CLIENTS.clear()


# ── Client de l'API ───────────────────────────────────────────────
def test_dossiers_fichiers_et_telechargement(client, racine, tmp_path):
    assert client.dossiers(DOSSIER) == ["Dep13", "Dep2A", "Dep75", "Divers"]
    assert client.dossier_existe(DOSSIER + "/Dep13/13001")
    assert not client.dossier_existe(DOSSIER + "/Dep99/99999")
    assert [f["Name"] for f in client.fichiers(DOSSIER + "/Dep13/13001")] == ["audit_13001.xlsx"]
    copie = client.telecharger(DOSSIER + "/Dep13/13001/audit_13001.xlsx", str(tmp_path / "c.xlsx"))
    assert open(copie, "rb").read() == open(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"), "rb").read()


def test_sans_session_ou_session_expiree(serveur):
    srv, _ = serveur
    anonyme = rest.ClientSharePoint(HOTE, {}, base=f"http://127.0.0.1:{srv.server_port}")
    with pytest.raises(rest.AuthentificationRequise):
        anonyme.verifier(DOSSIER)


def test_introuvable(client):
    with pytest.raises(FileNotFoundError):
        client.telecharger(DOSSIER + "/Dep13/13001/absent.xlsx", "x")
    with pytest.raises(FileNotFoundError):
        client.verifier(DOSSIER + "/nulle-part")


def test_adresses_d_api_echappees():
    c = rest.ClientSharePoint(HOTE, {})
    api = c._api("/sites/SNA/Documents partagés/l'audit #1", "GetFolderByServerRelativePath", "/Files")
    assert api == ("/sites/SNA/_api/web/GetFolderByServerRelativePath(decodedurl="
                   "'/sites/SNA/Documents%20partag%C3%A9s/l%27%27audit%20%231')/Files")
    assert rest.site_de("/teams/Equipe/Docs/x") == "/teams/Equipe"
    assert rest.site_de("/Shared Documents/x") == ""
    assert rest.chemin_serveur(LIEN + "/") == DOSSIER


def test_regles_de_nommage_comme_le_moteur():
    fichiers = [{"Name": "audit_1-PC.xlsx", "TimeLastModified": "2026-11-20T10:00:00Z"},
                {"Name": "audit_1.xlsx"}, {"Name": "notes.txt"}, {"Name": "audit_1 (2026-11-24).xlsx"}]
    fichiers.sort(key=lambda f: f["Name"])
    assert rest.fichier_audit(fichiers) == "audit_1 (2026-11-24).xlsx"   # premier par nom, comme sna
    assert [n for n, _ in rest.copies_de_conflit(fichiers)] == ["audit_1-PC.xlsx", "audit_1.xlsx"]


def test_cookies_envoyes_comme_le_navigateur():
    cookies = [
        {"name": "FedAuth", "value": "a", "domain": "contoso.sharepoint.com", "expires": -1},
        {"name": "rtFa", "value": "b", "domain": ".sharepoint.com", "expires": 4102444800},
        {"name": "autre", "value": "c", "domain": "fabrikam.sharepoint.com", "expires": -1},
        {"name": "perime", "value": "d", "domain": "contoso.sharepoint.com", "expires": 1000},
    ]
    assert cookies_pour(cookies, "contoso.sharepoint.com") == {"FedAuth": "a", "rtFa": "b"}


# ── Détection, audit et mise à jour par l'API ────────────────────
def test_communes_listees_par_l_api_comme_en_local(source_api, racine):
    distant = sp.AccesApi(source_api).lister_communes()
    local = [(c.code, c.departement, c.a_fichier) for c in lister_communes(racine)]
    assert distant == local


def test_audit_par_l_api_identique_au_local(source_api, racine, serveur):
    communes = ["13001", "13002", "13003", "2A004", "75020", "75021", "99999"]
    local = auditer_communes(communes, racine)
    with sp.acces_pour(source_api) as acces:
        distant = auditer_communes(communes, source_api.racine, source=source_api, acces=acces, paralleles=4)
        dossier = acces._dossier
    assert not os.path.exists(dossier)
    sig = lambda r: [(c.commune, c.statut, c.ids_non_valides, c.message.split(" : ")[0]) for c in r.communes]
    assert [s[:3] for s in sig(distant)] == [s[:3] for s in sig(local)]
    c = distant.commune("13001")
    assert c.chemin == LIEN + "/Dep13/13001/audit_13001.xlsx"
    assert any("copie en conflit" in a for a in distant.commune("75021").avertissements)
    assert distant.libelle_source == "SharePoint — connexion par le navigateur"


def test_audit_par_ids_via_l_api(source_api):
    with sp.acces_pour(source_api) as acces:
        r = auditer_communes(["13001"], source_api.racine, source=source_api, acces=acces,
                             ids={"13001": ["13001_2", "13001_404"]})
    assert r.commune("13001").ids_non_valides == ["13001_2"] and r.ids_introuvables() == ["13001_404"]


def test_session_expiree_pendant_l_audit(source_api, serveur):
    sp._CLIENTS[HOTE].cookies = {}
    with sp.acces_pour(source_api) as acces:
        c = auditer_communes(["13001"], source_api.racine, source=source_api, acces=acces).communes[0]
    assert c.statut is StatutAudit.ERREUR and "session SharePoint" in c.message


# ── Connexion par le navigateur (navigateur simulé) ───────────────
class _FauxNavigateur:
    """Simule Edge : la session apparaît après ``connexion_apres`` relevés de cookies."""

    lances: list = []

    def __init__(self, url, visible=True, session_apres=None, ferme_apres=None):
        self.url, self.visible = url, visible
        self.session_apres, self.ferme_apres, self.releves = session_apres, ferme_apres, 0
        type(self).lances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.ferme = True
        return False

    def ouvert(self):
        return self.ferme_apres is None or self.releves < self.ferme_apres

    def cookies(self):
        self.releves += 1
        if self.session_apres is not None and self.releves >= self.session_apres:
            return [{"name": "FedAuth", "value": "valide", "domain": HOTE, "expires": -1}]
        return [{"name": "FedAuth", "value": "anonyme", "domain": HOTE, "expires": -1}]


def _fabrique(silencieux=None, visible=None, ferme_apres=None):
    _FauxNavigateur.lances = []

    def fabriquer(url, visible=True):
        return _FauxNavigateur(url, visible, session_apres=(visible_session if visible else silencieux),
                               ferme_apres=ferme_apres if visible else None)
    visible_session = visible
    return fabriquer


def _client_factory(serveur):
    srv, _ = serveur
    return lambda hote, cookies: rest.ClientSharePoint(hote, cookies, base=f"http://127.0.0.1:{srv.server_port}")


def test_connexion_silencieuse_avec_le_profil(serveur):
    messages = []
    source = sp.connecter_navigateur(LIEN, on_etape=messages.append, fabrique=_fabrique(silencieux=1),
                                     client_factory=_client_factory(serveur))
    assert source == SourceAudit(API, LIEN, LIEN)
    assert [n.visible for n in _FauxNavigateur.lances] == [False]     # aucune fenêtre affichée
    assert sp.client_pour(LIEN).cookies == {"FedAuth": "valide"}
    sp._CLIENTS.clear()


def test_identification_dans_la_fenetre_du_navigateur(serveur, monkeypatch):
    monkeypatch.setattr(sp.time, "sleep", lambda s: None)
    messages = []
    source = sp.connecter_navigateur(LIEN, on_etape=messages.append, delai_silencieux_s=0.05,
                                     fabrique=_fabrique(silencieux=None, visible=3),
                                     client_factory=_client_factory(serveur))
    assert source.mode == API
    assert [n.visible for n in _FauxNavigateur.lances] == [False, True]
    assert _FauxNavigateur.lances[1].url == LIEN                       # ouvre directement le dossier
    assert any("identifiez-vous" in m for m in messages) and any("réussie" in m for m in messages)
    sp._CLIENTS.clear()


def test_sans_session_et_non_interactif(serveur):
    with pytest.raises(rest.AuthentificationRequise):
        sp.connecter_navigateur(LIEN, interactif=False, delai_silencieux_s=0.05,
                                fabrique=_fabrique(), client_factory=_client_factory(serveur))
    assert [n.visible for n in _FauxNavigateur.lances] == [False]


def test_fenetre_fermee_avant_la_connexion(serveur, monkeypatch):
    monkeypatch.setattr(sp.time, "sleep", lambda s: None)
    with pytest.raises(NavigateurFerme):
        sp.connecter_navigateur(LIEN, delai_silencieux_s=0.05, fabrique=_fabrique(ferme_apres=2),
                                client_factory=_client_factory(serveur))


def test_dossier_introuvable_une_fois_connecte(serveur):
    with pytest.raises(sp.AccesImpossible, match="dossier est introuvable"):
        sp.connecter_navigateur(LIEN + "/Inexistant", fabrique=_fabrique(silencieux=1),
                                client_factory=_client_factory(serveur))


def test_delai_de_connexion_depasse(serveur, monkeypatch):
    monkeypatch.setattr(sp.time, "sleep", lambda s: None)
    with pytest.raises(sp.AccesImpossible, match="non terminée"):
        sp.connecter_navigateur(LIEN, delai_silencieux_s=0.05, delai_s=0.2, fabrique=_fabrique(),
                                client_factory=_client_factory(serveur))


def test_oublier_la_connexion(tmp_path, monkeypatch, source_api):
    from avf.services import navigateur
    profil = tmp_path / "profil"
    profil.mkdir()
    monkeypatch.setattr(navigateur, "dossier_profil", lambda: str(profil))
    sp.oublier_connexion()
    assert not profil.exists() and sp._CLIENTS == {}
    with pytest.raises(rest.AuthentificationRequise):
        sp.client_pour(LIEN)


def test_simulation_sur_la_version_serveur_telechargee(source_api, racine):
    """La simulation compte les modifications sur une copie du fichier serveur, sans rien écrire."""
    avant = open(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"), "rb").read()
    bilan = mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), source_api.racine,
                          simuler=True, source=source_api)
    assert bilan.simulation and bilan.nb_communes_ok == 1 and bilan.nb_modifications == 4
    assert open(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"), "rb").read() == avant



# ── Mise à jour par l'API, avec la session du navigateur ──────────
class _PasDExcel:
    def __call__(self):
        raise AssertionError("la mise à jour en connexion navigateur ne doit pas lancer Excel")


def _audit_api(source):
    with sp.acces_pour(source) as acces:
        return auditer_communes(["13001", "2A004", "75021"], source.racine, source=source, acces=acces)


def test_mise_a_jour_par_l_api_sans_excel(source_api, racine, serveur):
    from avf.services.mise_a_jour import groupes_depuis_rapport, preparer_regles
    from datetime import date
    srv, faux = serveur
    regles = preparer_regles(ReglesConfig.defaut(), aujourd_hui=date(2026, 11, 26))
    messages = []
    bilan = mettre_a_jour(groupes_depuis_rapport(_audit_api(source_api)), regles, source_api.racine,
                          via_excel=True, session_factory=_PasDExcel(), source=source_api,
                          on_etape=lambda c, m: messages.append(m))
    assert bilan.via_api and not bilan.via_excel
    par = {r.commune: r for r in bilan.resultats}
    assert par["13001"].nb_modifications > 0 and par["2A004"].nb_modifications > 0
    assert "conflit" in par["75021"].message                      # copie de conflit : refus
    assert any("valeurs vérifiées sur SharePoint" in m for m in messages)
    assert any(a.startswith("POST /sites/SNA/_api/contextinfo") for a in faux.appels)
    apres = _audit_api(source_api)                                 # ré-audit sur le serveur
    assert apres.ids_non_valides_par_commune() == {"75021": ["75021_1"]}
    from openpyxl import load_workbook
    ws = load_workbook(os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx"))["Audit"]
    assert ws["F3"].value == "OK" and ws["D3"].value == datetime(2026, 11, 26)   # vraie date Excel


def test_modification_concurrente_jamais_ecrasee(source_api, racine, serveur):
    """Un collègue enregistre entre la lecture et l'envoi : sa version est relue, rien n'est perdu."""
    from openpyxl import load_workbook
    srv, faux = serveur
    fichier = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    deja = []

    def collegue(local):
        if local == fichier and not deja:
            deja.append(1)
            wb = load_workbook(fichier)
            wb["Audit"]["B2"] = "modifié par un collègue"
            wb.save(fichier)
            faux.versions[fichier] = faux.versions.get(fichier, 1) + 1

    faux.avant_ecriture = staticmethod(collegue)
    messages = []
    bilan = mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), source_api.racine,
                          source=source_api, on_etape=lambda c, m: messages.append(m))
    r = bilan.resultats[0]
    assert r.statut.name == "SUCCES" and r.ids_non_verifies == []
    assert any("enregistré par quelqu'un d'autre" in m for m in messages)
    ws = load_workbook(fichier)["Audit"]
    assert ws["B2"].value == "modifié par un collègue"          # sa modification est conservée
    assert ws["F3"].value == "OK"                               # et la nôtre est appliquée


def test_fichier_verrouille_rien_ecrit(source_api, racine, serveur):
    srv, faux = serveur
    fichier = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    avant = open(fichier, "rb").read()
    faux.verrouilles.add(fichier)
    bilan = mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), source_api.racine, source=source_api)
    assert "verrouill" in bilan.resultats[0].message and "13001_2" in bilan.echecs
    assert open(fichier, "rb").read() == avant


def test_fichier_extrait_rien_ecrit(source_api, racine, serveur):
    srv, faux = serveur
    fichier = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    faux.extraits[fichier] = "Alice Martin"
    bilan = mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), source_api.racine, source=source_api)
    assert bilan.resultats[0].message == "classeur verrouillé par Alice Martin"


def test_ecriture_non_retrouvee_id_en_attente(source_api, serveur):
    srv, faux = serveur
    faux.ignorer_ecritures = True
    bilan = mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), source_api.racine, source=source_api)
    assert bilan.resultats[0].ids_non_verifies == ["13001_2"]
    assert "13001_2" in bilan.echecs


def test_valeurs_deja_en_place_rien_envoye(source_api, serveur):
    srv, faux = serveur
    mettre_a_jour({"13001": ["13001_1"]}, ReglesConfig(regles=[]), source_api.racine, source=source_api)
    assert not any(a.startswith("POST") for a in faux.appels)
