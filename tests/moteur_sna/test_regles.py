"""Modèle de règles : conditions + action, sérialisation, valeurs par défaut."""

from sna.models.regles import (
    Action, COLONNE_ID_DEFAUT, Condition, OPERATEURS_DATE, OPERATEURS_NUM,
    OPERATEURS_TEXTE, REGLES_DEFAUT, RegleModification, ReglesConfig,
    SourceValeur, TYPE_DATE, TYPE_NUM, TYPE_TEXTE, TypeDroite,
    operateurs_pour_type,
)


def _regle(colonne="Z", valeur="", *, code="R1", source=SourceValeur.VALEUR,
           conditions=None, actif=True, exact=True):
    return RegleModification(
        code=code, libelle=code,
        action=Action(colonne=colonne, source=source, valeur=valeur),
        conditions=conditions or [], actif=actif, exact=exact,
    )


# ── Configuration livrée ─────────────────────────────────────────
def test_config_par_defaut_livre_les_regles_d_audit():
    """Le programme démarre avec les règles d'audit paramétrées dans REGLES_DEFAUT."""
    config = ReglesConfig.defaut()

    assert [r.code for r in config.regles] == [f"R{i}" for i in range(1, len(REGLES_DEFAUT) + 1)]
    assert len(config.actives) == len(REGLES_DEFAUT)
    assert config.ligne_entete == 1
    assert config.colonne_id == COLONNE_ID_DEFAUT
    # R1 : Audit ← « OK » — même colonne/valeur que dans REGLES_DEFAUT.
    r1 = config.regles[0]
    assert r1.action.colonne == REGLES_DEFAUT[0].action.colonne
    assert r1.action.valeur == REGLES_DEFAUT[0].action.valeur


def test_config_par_defaut_est_une_copie_independante():
    """Modifier le résultat de defaut() ne doit pas altérer REGLES_DEFAUT."""
    nb_avant = len(REGLES_DEFAUT)
    nb_conditions_avant = len(REGLES_DEFAUT[0].conditions)

    config = ReglesConfig.defaut()
    config.regles[0].conditions.append(Condition("X", "vide"))
    config.regles.clear()

    assert len(REGLES_DEFAUT) == nb_avant
    assert len(REGLES_DEFAUT[0].conditions) == nb_conditions_avant


def test_config_par_defaut_serialisable_aller_retour():
    defaut = ReglesConfig.defaut()

    assert ReglesConfig.from_dict(defaut.to_dict()) == defaut


# ── Une règle ────────────────────────────────────────────────────
def test_une_regle_impose_une_valeur_a_une_colonne():
    regle = _regle(colonne="Z", valeur="corrigé")

    assert regle.colonne_cible == "Z"
    assert regle.renseignee is True
    assert regle.actif is True


def test_colonne_cible_par_defaut_est_celle_de_la_premiere_condition():
    regle = RegleModification(
        code="R1", action=Action(),
        conditions=[Condition("Audit - Street View", "=", TypeDroite.VALEUR, "non")],
    )
    assert regle.colonne_cible == "Audit - Street View"


def test_effacement_demande_explicitement():
    """Une valeur vide n'efface rien ; il faut choisir la source « effacer »."""
    assert _regle(colonne="Z", valeur="").renseignee is False
    assert _regle(colonne="Z", source=SourceValeur.VIDE).renseignee is True
    assert _regle(colonne="Z", source=SourceValeur.VIDE).effacer is True


def test_recopie_de_colonne_est_renseignee_si_une_source_est_indiquee():
    assert _regle(colonne="Z", source=SourceValeur.COLONNE, valeur="").renseignee is False
    assert _regle(colonne="Z", source=SourceValeur.COLONNE, valeur="A").renseignee is True


def test_regles_actives_exclut_inactives_sans_colonne_et_sans_valeur():
    config = ReglesConfig(regles=[
        _regle(colonne="Z", valeur="a", code="R1"),
        _regle(colonne="Z", valeur="b", code="R2", actif=False),
        _regle(colonne="  ", valeur="c", code="R3"),
        _regle(colonne="A", valeur="", code="R4"),
        _regle(colonne="B", source=SourceValeur.VIDE, code="R5"),
    ])

    assert [r.code for r in config.actives] == ["R1", "R5"]
    assert [r.code for r in config.en_attente] == ["R4"]


def test_regle_sans_condition_reste_active():
    """Les conditions sont un filtre facultatif : une règle sans condition compte."""
    config = ReglesConfig(regles=[_regle(colonne="Z", valeur="a")])

    assert [r.code for r in config.actives] == ["R1"]


# ── Opérateurs selon le type ─────────────────────────────────────
def test_operateurs_dependent_du_type():
    assert operateurs_pour_type(TYPE_NUM) == OPERATEURS_NUM
    assert operateurs_pour_type(TYPE_TEXTE) == OPERATEURS_TEXTE
    assert operateurs_pour_type(TYPE_DATE) == OPERATEURS_DATE
    assert ">" not in OPERATEURS_TEXTE
    assert "contient" in OPERATEURS_TEXTE
    assert ">" in OPERATEURS_DATE


def test_code_unique_evite_les_doublons():
    config = ReglesConfig(regles=[
        _regle(colonne="Z", valeur="a", code="R1"),
        _regle(colonne="A", valeur="b", code="R2"),
    ])

    assert config.code_unique() == "R3"


# ── Sérialisation ────────────────────────────────────────────────
def test_serialisation_aller_retour():
    avant = ReglesConfig(
        regles=[
            _regle(colonne="Z", valeur="x", code="R1"),
            RegleModification(
                code="R2", libelle="recopie", exact=False,
                conditions=[Condition("A", "≥", TypeDroite.COLONNE, "B")],
                action=Action(colonne="C", source=SourceValeur.COLONNE, valeur="A"),
            ),
            RegleModification(
                code="R3", libelle="nettoyage", actif=False,
                action=Action(colonne="Commentaire", source=SourceValeur.VIDE),
            ),
        ],
        ligne_entete=3,
        colonne_id="Identifiant erreur",
    )

    assert ReglesConfig.from_dict(avant.to_dict()) == avant


def test_effacement_persiste():
    config = ReglesConfig(regles=[_regle(colonne="Z", source=SourceValeur.VIDE)])

    assert ReglesConfig.from_dict(config.to_dict()).regles[0].effacer is True


def test_condition_persiste():
    config = ReglesConfig(regles=[RegleModification(
        code="R1",
        conditions=[Condition("Statut", "contient", TypeDroite.VALEUR, "KO")],
        action=Action(colonne="Audit", valeur="OK"),
    )])

    relu = ReglesConfig.from_dict(config.to_dict()).regles[0]
    assert relu.conditions == [Condition("Statut", "contient", TypeDroite.VALEUR, "KO")]
    assert relu.action == Action(colonne="Audit", source=SourceValeur.VALEUR, valeur="OK")


def test_lecture_ancienne_forme_a_plat():
    """Un config.json «à plat» (colonne / valeur / effacer) reste lisible."""
    plat = {"regles": [
        {"code": "R1", "libelle": "l", "colonne": "Z", "valeur": "v"},
        {"code": "R2", "libelle": "e", "colonne": "A", "effacer": True},
    ]}
    regles = ReglesConfig.from_dict(plat).regles

    assert regles[0].colonne_cible == "Z" and regles[0].valeur == "v"
    assert regles[0].conditions == []
    assert regles[1].effacer is True


def test_option_de_mise_en_forme_persistee():
    config = ReglesConfig(
        regles=[_regle(colonne="Z", valeur="x")], harmoniser_style=False,
    )

    assert ReglesConfig.from_dict(config.to_dict()).harmoniser_style is False


def test_mise_en_forme_reprise_par_defaut():
    assert ReglesConfig.defaut().harmoniser_style is True
    assert ReglesConfig.from_dict({"regles": []}).harmoniser_style is True


# ── Résumé lisible ───────────────────────────────────────────────
def test_resume_sans_condition():
    config = ReglesConfig(regles=[
        _regle(colonne="Z", valeur="corrigé", code="R1"),
        _regle(colonne="A", source=SourceValeur.VIDE, code="R2"),
        _regle(colonne="Audit", valeur="", code="R3"),
    ])

    assert config.resume() == [
        "R1 (actif) : Z ← « corrigé »",
        "R2 (actif) : A ← (effacer)",
        "R3 (actif) : Audit ← (aucune valeur — règle en attente)",
    ]


def test_resume_avec_conditions():
    regle = RegleModification(
        code="R1",
        conditions=[
            Condition("Statut", "=", TypeDroite.VALEUR, "KO"),
            Condition("Audit", "vide"),
        ],
        action=Action(colonne="Audit", valeur="OK"),
    )

    assert regle.resume() == "R1 (actif) : SI Statut = « KO » ET Audit vide ALORS Audit ← « OK »"
