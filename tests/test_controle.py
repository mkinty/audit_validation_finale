"""Critères de validité d'une ligne (services.controle)."""

import pytest

from avf.models.audit import LigneAudit
from avf.models.colonnes import ColonnesAudit, parser_valeurs
from avf.services.controle import controler_ligne, est_date

VALIDE = dict(audit="OK", date_audit="12/03/2026", auditeur="Alice", validation="OK")


def _criteres(colonnes=None, **valeurs):
    ligne = LigneAudit(2, "13001_1", **{**VALIDE, **valeurs})
    return [a.critere for a in controler_ligne(ligne, colonnes or ColonnesAudit())]


@pytest.mark.parametrize("texte", [
    "12/03/2026", "1/3/2026", "12-03-2026", "12.03.2026", "12/03/26", "2026-03-12", "2026/03/12",
    "12/03/2026 14:30", "2026-03-12T14:30:00", "2026-03-12 14h30", "12 mars 2026", "1er mars 2026",
    "12 sept. 2026", "3 février 2026", "3 fevrier 2026", "12 DÉCEMBRE 2026", "46093", "46093.5",
])
def test_dates_acceptees(texte):
    assert est_date(texte)


@pytest.mark.parametrize("texte", [
    "", "   ", "à faire", "31/02/2026", "13/13/2026", "2026", "12/03", "OK", "12 marsss 2026",
    "123", "999999", "2026-02-30",
])
def test_pas_des_dates(texte):
    assert not est_date(texte)


def test_ligne_complete_validee():
    assert _criteres() == []


@pytest.mark.parametrize("audit", ["OK", "KO", "ok", " ko "])
def test_audit_ok_ou_ko(audit):
    assert _criteres(audit=audit) == []


@pytest.mark.parametrize("audit", ["", "  ", "Fait", "OK KO", "NOK"])
def test_audit_refuse(audit):
    assert _criteres(audit=audit) == ["audit"]


@pytest.mark.parametrize("vf", ["", "   ", "KO", "Validé", "1"])
def test_validation_finale_doit_etre_ok(vf):
    assert _criteres(validation=vf) == ["validation"]


def test_validation_finale_casse_et_espaces():
    assert _criteres(validation=" ok ") == []


def test_auditeur_et_date_obligatoires():
    assert _criteres(auditeur=" ", date_audit="") == ["date_audit", "auditeur"]
    assert _criteres(date_audit="hier") == ["date_audit"]


def test_tous_les_criteres_et_messages():
    ligne = LigneAudit(5, "X", audit="Fait", date_audit="hier", auditeur="", validation="Validé")
    assert [a.message for a in controler_ligne(ligne, ColonnesAudit())] == [
        "Audit « Fait » ≠ OK / KO",
        "Date audit « hier » n'est pas une date",
        "Auditeur vide",
        "Validation finale « Validé » ≠ OK",
    ]


def test_valeurs_acceptees_reglables():
    colonnes = ColonnesAudit(valeurs_validation=("OK", "Validé"), valeurs_audit=("OK",))
    assert _criteres(colonnes, validation="validé") == []
    assert _criteres(colonnes, audit="KO") == ["audit"]


def test_parser_valeurs():
    assert parser_valeurs("OK, ko; OK ,, Validé") == ("OK", "ko", "Validé")
    assert parser_valeurs(["OK", " "]) == ("OK",)
    assert parser_valeurs(None) == ()


def test_colonnes_aller_retour_et_defauts():
    c = ColonnesAudit(valeurs_audit=("OK", "KO", "SO"))
    assert ColonnesAudit.from_dict(c.to_dict()) == c
    assert ColonnesAudit.from_dict({"valeurs_audit": "  ", "valeurs_validation": []}) == ColonnesAudit()
