"""Tests du classeur : zéro erreur de formule, contrôles OK, cohérences comptables et de valorisation."""
import pytest


def test_no_formula_errors(wave, fict):
    assert wave.errors() == []
    assert fict.errors() == []


def test_global_control_ok(wave, fict):
    for b in (wave, fict):
        assert b.v("Synthese.ctrl_glob", "B") == "OK"


def test_balance_sheet_balances(wave, fict):
    for b in (wave, fict):
        for c in "BCD":
            assert abs(b.v("Bilan_BFR.ctrl_ta", c)) < 0.5


def test_net_result_ties(fict):
    for c in "BCD":
        assert abs(fict.v("SIG_CAF.ctrl_rn", c)) < 0.5


def test_dcf_scenarios_ordered(wave, fict):
    for b in (wave, fict):
        bear, base, bull = (b.v(f"DCF.{t}.s_ps", "B") for t in ("bear", "base", "bull"))
        assert bear < base < bull


def test_gordon_multiple_consistency(fict):
    """La valeur retenue est la moyenne pondérée des deux méthodes (poids Hypotheses)."""
    ev, evg, evm = (fict.v(f"DCF.base.{k}", "B") for k in ("s_ev", "s_evg", "s_evm"))
    assert min(evg, evm) - 1e-6 <= ev <= max(evg, evm) + 1e-6


def test_lbo_irr_matches_max_price(wave):
    """Au prix maximal pour un TRI cible, le TRI recalculé doit retrouver la cible (relation exacte du modèle)."""
    ws = wave.wb["LBO"]
    moic, irr = ws["B37"].value, ws["B38"].value
    assert moic == pytest.approx((1 + irr) ** 5, rel=1e-6)
    eq_max = ws["B43"].value
    assert ws["B36"].value / eq_max == pytest.approx((1 + ws["B42"].value) ** 5, rel=1e-6)


def test_football_field_has_ranges(wave):
    """Les lignes calculables ont bas <= haut ; une méthode sans données suffisantes affiche n/d (jamais une fausse fourchette)."""
    ws = wave.wb["Synthese"]
    r0, r1 = wave.ix["Synthese.ff_first"][1], wave.ix["Synthese.ff_last"][1]
    numeric = 0
    for r in range(r0, r1 + 1):
        lo, hi = ws[f"B{r}"].value, ws[f"C{r}"].value
        if isinstance(lo, (int, float)) and isinstance(hi, (int, float)):
            numeric += 1
            assert lo <= hi + 1e-9
        else:
            assert lo == "n/d" and hi == "n/d"
    assert numeric >= 3


def test_ifrs16_switch_changes_net_debt(wave):
    assert wave.v("Bilan_BFR.dn16") == pytest.approx(wave.v("Bilan_BFR.dn") - wave.v("Bilan_BFR.leases"))
    assert wave.v("DCF.base.s_nd", "B") == pytest.approx(wave.v("Bilan_BFR.dn16"))
