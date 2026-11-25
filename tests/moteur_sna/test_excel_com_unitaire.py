"""Fonctions de ``excel_com`` testables sans Excel (toujours exécutés)."""

from __future__ import annotations

from datetime import date, datetime
import gc
import weakref

from sna.services.excel_com import (
    XL_OBJET_INVALIDE, _brut_apres_ecriture, _codes_com, _conforme_valeur, _correspondance_colonnes,
    _est_deconnexion, _est_objet_invalide, _liberer_objets_com, _serie,
)


class _FausseComError(Exception):
    """Même forme d'arguments que ``pywintypes.com_error``."""


def test_codes_com_appel_et_excel():
    exc = _FausseComError(-2147352567, "Exception", (0, "Excel", "x", None, 0, XL_OBJET_INVALIDE), None)
    assert _codes_com(exc) == {-2147352567, XL_OBJET_INVALIDE}
    assert _est_objet_invalide(exc) and not _est_deconnexion(exc)
    assert _est_deconnexion(_FausseComError(-2147417848, "déconnecté", None, None))
    assert _codes_com(ValueError("pas COM")) == set()


def test_conforme_valeur():
    assert _conforme_valeur(None, "") and _conforme_valeur("", None)
    assert _conforme_valeur("OK", "OK") and not _conforme_valeur("ok", "OK")
    assert _conforme_valeur("a\nb", "a\r\nb")
    assert _conforme_valeur(_serie(date(2026, 11, 24)), date(2026, 11, 24))
    assert not _conforme_valeur(True, datetime(1899, 12, 31))   # un booléen n'est pas une date
    assert _conforme_valeur(3.0, 3) and not _conforme_valeur("3", 3)


def test_brut_apres_ecriture():
    assert _brut_apres_ecriture("") is None
    assert _brut_apres_ecriture(date(1900, 1, 1)) == 2.0
    assert _brut_apres_ecriture("OK") == "OK"


def test_correspondance_colonnes_apres_insertion():
    avant = {1: "ID erreur", 2: "Audit", 3: "Validation finale"}
    apres = {1: "Commentaire", 2: "ID erreur", 3: "Audit", 4: "Validation  FINALE"}
    assert _correspondance_colonnes(avant, apres) == {1: 2, 2: 3, 3: 4}
    assert _correspondance_colonnes(avant, dict(avant)) == {}


def test_liberer_objets_com_vide_les_cadres_de_la_trace():
    """Les objets locaux (proxies COM) ne survivent pas dans la trace de l'exception."""
    class ObjetCom:
        pass

    def traitement(objet):
        classeur = objet  # noqa: F841 — variable locale gardée par la trace
        try:
            raise KeyError("cause")
        except KeyError as cause:
            raise RuntimeError("échec") from cause

    objet = ObjetCom()
    temoin = weakref.ref(objet)
    try:
        traitement(objet)
    except RuntimeError as exc:
        erreur = exc
    del objet
    assert temoin() is not None          # retenu par la trace
    _liberer_objets_com(erreur)
    gc.collect()
    assert temoin() is None              # libéré, sans attendre le ramasse-miettes d'un autre thread
    assert str(erreur) == "échec" and erreur.__cause__.args == ("cause",)
