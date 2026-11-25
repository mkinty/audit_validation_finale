"""Audit depuis SharePoint : lien, WebDAV, bibliothèque synchronisée, lecture de la version serveur."""

import os
import shutil
import time

import pytest
from openpyxl import load_workbook

from avf.models.source import EXCEL, LOCAL, URL as URL_MODE, WEBDAV, SourceAudit
from avf.services import sharepoint as sp
from avf.services.audit import auditer_communes
from avf.services.mise_a_jour import (
    EcritureSharePointImpossible, groupes_depuis_rapport, mettre_a_jour,
)
from sna.models.regles import ReglesConfig
from sna.services.excel import appliquer_regles
from sna.services.onedrive import url_sharepoint

HOTE = "contoso.sharepoint.com"
DOSSIER = "/sites/SNA/Shared Documents/Audit SNA"
URL = "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Audit%20SNA"
WEBDAV_RACINE = r"\\contoso.sharepoint.com@SSL\DavWWWRoot\sites\SNA\Shared Documents\Audit SNA"


# ── Lien saisi ────────────────────────────────────────────────────
@pytest.mark.parametrize("lien", [
    URL,
    URL + "/",
    "  " + URL + "  ",
    '"' + URL + '"',
    "https://contoso.sharepoint.com/sites/SNA/Shared Documents/Audit SNA",
    "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Forms/AllItems.aspx"
    "?id=%2Fsites%2FSNA%2FShared%20Documents%2FAudit%20SNA&viewid=1234-abcd",
    "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Forms/AllItems.aspx"
    "?RootFolder=%2Fsites%2FSNA%2FShared%20Documents%2FAudit%20SNA",
    "https://contoso.sharepoint.com/:f:/r/sites/SNA/Shared%20Documents/Audit%20SNA?csf=1&web=1&e=Xy12Ab",
    "https://CONTOSO.sharepoint.com/sites/SNA/Shared%20Documents/Audit%20SNA",
    WEBDAV_RACINE,
])
def test_formats_de_lien_acceptes(lien):
    l = sp.analyser_lien(lien)
    assert (l.hote, l.chemin) == (HOTE, DOSSIER)
    assert l.url == URL
    assert l.chemin_webdav == WEBDAV_RACINE


def test_vue_de_bibliotheque_sans_id_donne_sa_racine():
    l = sp.analyser_lien("https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Forms/AllItems.aspx")
    assert l.chemin == "/sites/SNA/Shared Documents"


def test_onedrive_entreprise():
    l = sp.analyser_lien("https://contoso-my.sharepoint.com/personal/jean_contoso_fr/Documents/Audit")
    assert l.chemin_webdav == r"\\contoso-my.sharepoint.com@SSL\DavWWWRoot\personal\jean_contoso_fr\Documents\Audit"


@pytest.mark.parametrize("lien, message", [
    ("", "aucun lien"),
    (r"C:\Users\moi\OneDrive\Audit", "https://"),
    ("https://contoso.sharepoint.com/:f:/s/SNA/EaBcDeFgHiJ", "lien de partage sans chemin"),
    ("https://contoso.sharepoint.com/sites/SNA/SitePages/Accueil.aspx", "page"),
    ("https://contoso.sharepoint.com/sites", "bibliothèque"),
])
def test_liens_refuses(lien, message):
    with pytest.raises(sp.LienInvalide, match=message):
        sp.analyser_lien(lien)


def test_url_depuis_webdav_et_moteur():
    fichier = WEBDAV_RACINE + r"\Dep13\13001\audit_13001.xlsx"
    attendu = URL + "/Dep13/13001/audit_13001.xlsx"
    assert sp.url_depuis_webdav(fichier) == attendu
    assert url_sharepoint(fichier, montages=[]) == attendu      # la mise à jour ouvre l'URL
    assert sp.url_depuis_webdav(r"C:\dossier\fichier.xlsx") is None


# ── Bibliothèque synchronisée ─────────────────────────────────────
@pytest.mark.parametrize("url_registre", [
    "https://contoso.sharepoint.com/sites/SNA/Shared Documents/",
    "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents",
    "https://CONTOSO.sharepoint.com/sites/sna/shared documents/",
])
def test_dossier_synchronise(url_registre):
    montages = [(r"C:\OD\Autre", "https://contoso.sharepoint.com/sites/Autre/Docs"),
                (r"C:\OD\SNA - Documents", url_registre)]
    local = sp.dossier_synchronise(sp.analyser_lien(URL), montages)
    assert local == os.path.join(r"C:\OD\SNA - Documents", "Audit SNA")


def test_dossier_synchronise_plus_long_prefixe_et_absent():
    montages = [(r"C:\OD\Bibli", "https://contoso.sharepoint.com/sites/SNA/Shared Documents"),
                (r"C:\OD\Raccourci", "https://contoso.sharepoint.com/sites/SNA/Shared Documents/Audit SNA")]
    assert sp.dossier_synchronise(sp.analyser_lien(URL), montages) == r"C:\OD\Raccourci"
    assert sp.dossier_synchronise(sp.analyser_lien(URL), []) is None


# ── Connexion ─────────────────────────────────────────────────────
def test_connexion_webdav_preferee():
    vus = []
    source = sp.connecter(URL, montages=[], est_dossier=lambda c: vus.append(c) or True)
    assert source == SourceAudit(WEBDAV, WEBDAV_RACINE, URL)
    assert vus == [WEBDAV_RACINE]


def test_repli_sur_la_bibliotheque_synchronisee(tmp_path):
    local = tmp_path / "SNA - Documents" / "Audit SNA"
    local.mkdir(parents=True)
    montages = [(str(tmp_path / "SNA - Documents"), "https://contoso.sharepoint.com/sites/SNA/Shared Documents")]
    source = sp.connecter(URL, montages=montages, est_dossier=lambda c: c != WEBDAV_RACINE and os.path.isdir(c))
    assert (source.mode, source.racine, source.lien) == (EXCEL, str(local), URL)
    assert source.sharepoint and source.listable and "bibliothèque synchronisée" in source.libelle


def test_repli_sur_l_ouverture_par_adresse(tmp_path):
    """Ni WebDAV ni synchronisation (cas réel : erreur 224) : fichiers ouverts par adresse via Excel."""
    def refuse(chemin):
        raise OSError(22, "Accès refusé. Avant d'ouvrir des fichiers de cet emplacement, vous devez "
                          "d'abord ajouter le site Web à votre liste de sites approuvés")

    source = sp.connecter(URL, montages=[], est_dossier=refuse, excel_present=lambda: True)
    assert (source.mode, source.racine, source.lien) == (URL_MODE, URL, URL)
    assert not source.listable and "sans synchronisation" in source.libelle
    assert "erreur 224" in source.diagnostic and "Afficher dans l'Explorateur" in source.diagnostic
    assert source.url_fichier("2a004") == URL + "/Dep2A/2A004/audit_2A004.xlsx"


def test_webdav_trop_lent_ne_bloque_pas(tmp_path):
    def lent(chemin):
        if chemin == WEBDAV_RACINE:
            time.sleep(5)
        return os.path.isdir(chemin)

    debut = time.monotonic()
    with pytest.raises(sp.AccesImpossible, match="n'a pas répondu à temps"):
        sp.connecter(URL, delai_s=0.3, montages=[], est_dossier=lent, excel_present=lambda: False)
    assert time.monotonic() - debut < 2


def test_aucun_acces_explique_la_marche_a_suivre():
    def refuse(chemin):
        raise PermissionError(13, "Accès refusé")

    with pytest.raises(sp.AccesImpossible) as erreur:
        sp.connecter(URL, montages=[], est_dossier=refuse, excel_present=lambda: False)
    texte = str(erreur.value)
    assert "Accès refusé" in texte and WEBDAV_RACINE in texte
    assert "Afficher dans l'Explorateur" in texte and "Synchroniser" in texte and "Excel" in texte


def test_cause_reelle_du_refus_webdav_remontee():
    """``os.path.isdir`` taisait la cause : la liste du dossier la donne."""
    with pytest.raises(OSError):
        sp._dossier_accessible(os.path.join(os.getcwd(), "dossier_qui_n_existe_pas"))
    assert sp._cause_webdav("delai").startswith("le serveur n'a pas répondu")
    assert sp._cause_webdav(None) == "accès refusé ou dossier introuvable"


# ── Audit depuis SharePoint ───────────────────────────────────────
def test_audit_webdav_identique_et_source_affichee(racine):
    """En WebDAV, le dossier se parcourt comme un dossier local : même résultat, source tracée."""
    local = auditer_communes(["13001", "13002"], racine)
    source = SourceAudit(WEBDAV, racine, URL)
    distant = auditer_communes(["13001", "13002"], racine, source=source)
    assert [(c.commune, c.ids_non_valides) for c in distant.communes] == \
        [(c.commune, c.ids_non_valides) for c in local.communes]
    assert distant.racine == URL and distant.source is source
    assert distant.libelle_source == "SharePoint — accès direct (WebDAV)"
    assert local.libelle_source == "Dossier local"


class _SessionServeur:
    """Remplace SessionExcel : « télécharge » depuis un dossier qui joue le rôle du serveur."""

    def __init__(self, serveur, local, url_racine):
        self.serveur, self.local, self.url_racine = serveur, local, url_racine
        self.ouvertures, self.ouverte, self.fermee = [], False, False

    def __enter__(self):
        self.ouverte = True
        return self

    def __exit__(self, *exc):
        self.fermee = True
        return False

    def copier_version_serveur(self, url, destination, on_etape=None):
        self.ouvertures.append(url)
        relatif = url[len(self.url_racine):].lstrip("/").replace("%20", " ")
        shutil.copy(os.path.join(self.serveur, *relatif.split("/")), destination)
        return destination

    def appliquer_regles(self, fichier, config, ids, on_etape=None):
        self.ouvertures.append(fichier)
        return appliquer_regles(fichier, config, ids, on_etape=on_etape)


@pytest.fixture
def serveur_divergent(racine, tmp_path):
    """« Serveur » = copie de la racine où 13002 n'est plus valide : la copie locale ne le sait pas."""
    serveur = str(tmp_path / "serveur")
    shutil.copytree(racine, serveur)
    chemin = os.path.join(serveur, "Dep13", "13002", "audit_13002.xlsx")
    wb = load_workbook(chemin)
    wb["Audit"]["F2"] = "KO"          # Validation finale de 13002_1 changée sur le serveur
    wb.save(chemin)
    return serveur


def test_audit_via_excel_lit_la_version_serveur(racine, serveur_divergent):
    source = SourceAudit(EXCEL, racine, URL)
    session = _SessionServeur(serveur_divergent, racine, URL)
    with sp.acces_pour(source, session_factory=lambda: session, montages=[(racine, URL)]) as acces:
        rapport = auditer_communes(["13001", "13002"], racine, source=source, acces=acces)
        dossier = acces._dossier
        assert os.path.isdir(dossier)
    assert session.ouverte and session.fermee
    assert not os.path.exists(dossier)                         # copies temporaires supprimées
    assert session.ouvertures == [URL + "/Dep13/13001/audit_13001.xlsx",
                                  URL + "/Dep13/13002/audit_13002.xlsx"]
    c = rapport.commune("13002")
    assert c.ids_non_valides == ["13002_1"]                    # version serveur auditée
    assert c.chemin == URL + "/Dep13/13002/audit_13002.xlsx"   # adresse SharePoint affichée
    assert auditer_communes(["13002"], racine).commune("13002").complete   # la copie locale, elle, l'est


def test_echec_de_telechargement_isole_la_commune(racine):
    class SessionCassee(_SessionServeur):
        def copier_version_serveur(self, url, destination, on_etape=None):
            raise RuntimeError("fichier verrouillé sur le serveur")

    source = SourceAudit(EXCEL, racine, URL)
    with sp.acces_pour(source, lambda: SessionCassee("", racine, URL), montages=[(racine, URL)]) as acces:
        rapport = auditer_communes(["13001", "13002"], racine, source=source, acces=acces)
    assert all(not c.lisible and "verrouillé" in c.message for c in rapport.communes)


def test_fichier_hors_bibliotheque_synchronisee(racine):
    source = SourceAudit(EXCEL, racine, URL)
    with sp.acces_pour(source, lambda: _SessionServeur("", racine, URL), montages=[]) as acces:
        rapport = auditer_communes(["13001"], racine, source=source, acces=acces)
    assert "adresse SharePoint inconnue" in rapport.communes[0].message


def test_acces_local_par_defaut():
    assert type(sp.acces_pour(SourceAudit(LOCAL, "x"))) is sp.AccesFichiers
    assert type(sp.acces_pour(SourceAudit(WEBDAV, "x"))) is sp.AccesFichiers


# ── Mise à jour d'une source SharePoint ───────────────────────────
def test_mise_a_jour_sharepoint_toujours_via_excel(racine):
    session = _SessionServeur("", racine, URL)
    groupes = groupes_depuis_rapport(auditer_communes(["13001"], racine))
    bilan = mettre_a_jour(groupes, ReglesConfig.defaut(), racine, via_excel=False,
                          session_factory=lambda: session, source=SourceAudit(WEBDAV, racine, URL))
    assert bilan.via_excel and session.ouvertures


def test_bibliotheque_synchronisee_sans_prereperage(racine):
    """La copie locale ne contient pas l'ID, le serveur peut-être : Excel est quand même ouvert."""
    session = _SessionServeur("", racine, URL)
    bilan = mettre_a_jour({"13001": ["13001_999"]}, ReglesConfig.defaut(), racine,
                          session_factory=lambda: session, source=SourceAudit(EXCEL, racine, URL))
    assert session.ouvertures == [os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")]
    assert bilan.echecs == {"13001_999": "ID erreur introuvable dans le fichier"}


def test_sharepoint_sans_excel_refuse(racine, monkeypatch):
    from avf.services import mise_a_jour as module
    monkeypatch.setattr(module, "excel_disponible", lambda: False)
    with pytest.raises(EcritureSharePointImpossible):
        mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), racine,
                      source=SourceAudit(WEBDAV, racine, URL))
    # La simulation, elle, n'écrit rien : autorisée
    assert mettre_a_jour({"13001": ["13001_2"]}, ReglesConfig.defaut(), racine, simuler=True,
                         source=SourceAudit(WEBDAV, racine, URL)).simulation


# ── Vraie instance Excel : copie de la version ouverte ────────────
from sna.services.excel_com import SessionExcel, excel_disponible  # noqa: E402


@pytest.mark.excel
@pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")
def test_copier_version_serveur_excel_reel(racine, tmp_path):
    original = os.path.join(racine, "Dep13", "13001", "audit_13001.xlsx")
    copie = str(tmp_path / "copie.xlsx")
    with SessionExcel() as session:
        assert session.copier_version_serveur(original, copie) == copie
    from avf.models.colonnes import ColonnesAudit
    from avf.services.lecture_audit import lire_audit

    def lu(chemin):   # ce que l'audit lit (une cellule « ␣␣␣ » est vide dans les deux cas)
        return [(l.ligne, l.id_erreur, l.audit, l.date_audit, l.auditeur, l.validation)
                for l in lire_audit(chemin, "ID erreur", ColonnesAudit()).lignes]
    assert lu(copie) == lu(original)


# ── Mode « via Excel, sans synchronisation » (ouverture par adresse) ─
def test_audit_par_adresse_lit_le_serveur(racine, serveur_divergent):
    source = SourceAudit(URL_MODE, URL, URL)
    session = _SessionServeur(serveur_divergent, racine, URL)
    with sp.acces_pour(source, session_factory=lambda: session) as acces:
        rapport = auditer_communes(["13001", "13002"], source.racine, source=source, acces=acces)
    assert session.ouvertures == [URL + "/Dep13/13001/audit_13001.xlsx", URL + "/Dep13/13002/audit_13002.xlsx"]
    assert rapport.commune("13002").ids_non_valides == ["13002_1"]       # version serveur
    assert rapport.commune("13002").chemin == URL + "/Dep13/13002/audit_13002.xlsx"


def test_adresse_absente_sur_le_serveur_donne_fichier_introuvable(racine):
    class SessionSansFichier(_SessionServeur):
        def copier_version_serveur(self, url, destination, on_etape=None):
            raise Exception(-2147352567, "Exception", (0, "Microsoft Excel",
                            "Désolé... Nous ne trouvons pas audit_99999.xlsx.", None, 0, -2146827284), None)

    source = SourceAudit(URL_MODE, URL, URL)
    with sp.acces_pour(source, lambda: SessionSansFichier("", racine, URL)) as acces:
        c = auditer_communes(["99999"], URL, source=source, acces=acces).communes[0]
    from avf.models.audit import StatutAudit
    assert c.statut is StatutAudit.INTROUVABLE and "introuvable sur SharePoint" in c.message


def test_audit_par_ids_en_mode_adresse(racine, serveur_divergent):
    source = SourceAudit(URL_MODE, URL, URL)
    session = _SessionServeur(serveur_divergent, racine, URL)
    with sp.acces_pour(source, lambda: session) as acces:
        r = auditer_communes(["13002"], URL, source=source, acces=acces, ids={"13002": ["13002_1"]})
    assert r.commune("13002").ids_non_valides == ["13002_1"] and r.nb_lignes == 1


def test_mise_a_jour_par_adresse(racine):
    """La mise à jour ouvre le fichier par son adresse, sans repérage local ni copie de conflit."""
    session = _SessionServeur("", racine, URL)
    vus = []

    def appliquer(fichier, config, ids, on_etape=None):
        vus.append(fichier)
        from sna.services.excel import BilanFichier
        return BilanFichier(comptes={"R1": 1}, ids_trouves=set(ids))
    session.appliquer_regles = appliquer
    bilan = mettre_a_jour({"75021": ["75021_1"]}, ReglesConfig.defaut(), URL,
                          session_factory=lambda: session, source=SourceAudit(URL_MODE, URL, URL))
    assert vus == [URL + "/Dep75/75021/audit_75021.xlsx"]
    assert bilan.nb_communes_ok == 1 and bilan.echecs == {}      # pas de refus « copie de conflit »
