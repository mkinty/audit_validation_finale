from html.parser import HTMLParser
import json
import os
import re
import shutil
import subprocess

import pytest

from avf.models.audit import AuditCommune, IdErreurAudit, LigneAudit, RapportAudit, StatutAudit
from avf.services.audit import auditer_communes
from avf.services.export import exporter_html, nom_fichier_defaut
from avf.views import presentation as p
from avf.views.rapport_html import generer_html

from .test_audit import TOUTES


class _Analyse(HTMLParser):
    """Vérifie l'équilibre des balises et collecte les attributs utiles."""

    VIDES = {"meta", "br", "input", "img", "hr", "link"}

    def __init__(self):
        super().__init__()
        self.pile, self.erreurs, self.copier, self.scripts = [], [], [], []
        self._dans_script = False

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if "data-copier" in a:
            self.copier.append(a["data-copier"])
        if tag not in self.VIDES:
            self.pile.append(tag)
        self._dans_script = tag == "script"

    def handle_endtag(self, tag):
        if not self.pile or self.pile[-1] != tag:
            self.erreurs.append(f"</{tag}> inattendu (pile : {self.pile[-3:]})")
        else:
            self.pile.pop()
        self._dans_script = False

    def handle_data(self, data):
        if self._dans_script:
            self.scripts.append(data)


@pytest.fixture
def rapport(racine):
    return auditer_communes(TOUTES, racine)


def _analyser(html):
    a = _Analyse()
    a.feed(html)
    return a


def test_html_bien_forme(rapport):
    html = generer_html(rapport)
    a = _analyser(html)
    assert a.erreurs == [] and a.pile == []
    assert html.startswith("<!DOCTYPE html>")


def test_html_contient_chiffres_et_ids(rapport):
    html = generer_html(rapport)
    assert "Audit de la Validation finale" in html
    for ident in rapport.tous_ids_vides():
        assert ident in html
    assert f">{p.nombre(rapport.nb_vides)}<" in html
    assert "Fichier introuvable" in html and "Erreur de lecture" in html
    assert "copie en conflit" in html


def test_boutons_de_copie_par_commune(rapport):
    a = _analyser(generer_html(rapport))
    assert a.copier == ["13001", "2A004", "75021"]


def test_donnees_json_de_copie(rapport):
    html = generer_html(rapport)
    bloc = re.search(r'<script id="donnees" type="application/json">(.*?)</script>', html, re.S)
    assert json.loads(bloc.group(1)) == rapport.ids_vides_par_commune()


def test_echappement_des_valeurs_du_classeur():
    malveillant = '<img src=x onerror="alert(1)">'
    c = AuditCommune(
        "13001", "13", chemin="C:/a&b/<x>.xlsx",
        ids=[IdErreurAudit("13001_1</script><b>", [LigneAudit(2, "13001_1", audit=malveillant)])],
    )
    html = generer_html(RapportAudit(racine="R<&>", communes=[c]))
    assert malveillant not in html
    assert "&lt;img src=x" in html
    assert "</script><b>" not in html  # le JSON ne peut pas fermer la balise
    a = _analyser(html)
    assert a.erreurs == []


def test_rapport_vide():
    html = generer_html(RapportAudit(racine="r"))
    assert "Aucune commune auditée." in html
    assert "disabled" in html
    assert _analyser(html).erreurs == []


def test_commune_complete_sans_bouton_ni_detail():
    c = AuditCommune("13002", "13", ids=[IdErreurAudit("13002_1", [LigneAudit(2, "13002_1", validation="OK")])])
    html = generer_html(RapportAudit(racine="r", communes=[c]))
    assert 'data-copier="13002"' not in html
    assert 'class="detail"' not in html
    assert "✅ Complète" in html


def test_export_ecrit_le_fichier(rapport, tmp_path):
    cible = tmp_path / "sous" / nom_fichier_defaut(rapport)
    chemin = exporter_html(rapport, str(cible))
    assert os.path.isfile(chemin)
    assert open(chemin, encoding="utf-8").read() == generer_html(rapport)
    assert not os.path.exists(chemin + ".tmp")
    assert re.fullmatch(r"rapport_validation_finale_\d{8}_\d{4}\.html", cible.name)


@pytest.mark.skipif(shutil.which("node") is None, reason="node indisponible")
def test_javascript_syntaxiquement_valide(rapport, tmp_path):
    a = _analyser(generer_html(rapport))
    script = tmp_path / "rapport.js"
    script.write_text(a.scripts[-1], encoding="utf-8")
    resultat = subprocess.run(["node", "--check", str(script)], capture_output=True, text=True)
    assert resultat.returncode == 0, resultat.stderr


# ── presentation ──────────────────────────────────────────────────
@pytest.mark.parametrize("taux, attendu", [
    (None, "—"), (0, "0\u202f%"), (1, "100\u202f%"), (0.875, "87,5\u202f%"), (2 / 3, "66,7\u202f%"),
])
def test_pourcent(taux, attendu):
    assert p.pourcent(taux) == attendu


def test_nombre_et_duree():
    assert p.nombre(1284567) == "1\u202f284\u202f567"
    assert p.duree(2.44) == "2,4 s"
    assert p.duree(125) == "2 min 05 s"


def test_etat():
    assert p.etat(AuditCommune("1", "1", statut=StatutAudit.ERREUR)) is p.ETAT_ERREUR
    assert p.etat(AuditCommune("1", "1", statut=StatutAudit.INTROUVABLE)) is p.ETAT_INTROUVABLE
    assert p.etat(AuditCommune("1", "1")) is p.ETAT_SANS_ID


def test_resume(rapport):
    lignes = p.resume(rapport)
    assert lignes[0].startswith("7 commune(s)")
    assert "6 sans" in lignes[1]


def test_texte_ids():
    assert p.texte_ids(["a", "b"]) == "a\nb"
    assert p.texte_ids(["a", "b"], p.SEPARATEURS_COPIE["virgule"]) == "a, b"
