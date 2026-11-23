"""Validation sur NVIDIA (10-K / communiqués FY2024-FY2026, US GAAP, M$). Voir docs/VALIDATION_NVIDIA.md."""
import pytest

PUBLISHED = [
    ("Revenue FY2026", "SIG_CAF.ca", "D", 215938, 0.5), ("Revenue FY2025", "SIG_CAF.ca", "C", 130497, 0.5), ("Revenue FY2024", "SIG_CAF.ca", "B", 60922, 0.5),
    ("Operating income FY2026", "SIG_CAF.ebit", "D", 130387, 0.5), ("Operating income FY2025", "SIG_CAF.ebit", "C", 81453, 0.5),
    ("Operating income FY2024", "SIG_CAF.ebit", "B", 32972, 0.5),
    ("Net income FY2026", "SIG_CAF.rn", "D", 120067, 0.5), ("Net income FY2025", "SIG_CAF.rn", "C", 72880, 0.5), ("Net income FY2024", "SIG_CAF.rn", "B", 29760, 0.5),
    ("Total assets FY2026", "Bilan_BFR.ta", "D", 206803, 0.5), ("Total assets FY2025", "Bilan_BFR.ta", "C", 111601, 0.5),
    ("Total assets FY2024", "Bilan_BFR.ta", "B", 65728, 0.5),
    ("Shareholders' equity FY2026", "Bilan_BFR.cp", "D", 157293, 0.5),
    ("EBITDA FY2026 = operating income + D&A", "SIG_CAF.ebe", "D", 133230, 0.5),
    ("EBITDA FY2025", "SIG_CAF.ebe", "C", 83317, 0.5),
    ("Net cash (cash + securities - debt) FY2026", "Bilan_BFR.dn16", "D", -54088, 0.5),
]


@pytest.mark.parametrize("label,key,col,published,tol", PUBLISHED, ids=[p[0] for p in PUBLISHED])
def test_against_published(nvda, label, key, col, published, tol):
    assert nvda.v(key, col) == pytest.approx(published, abs=tol)


def test_no_errors_and_controls(nvda):
    assert nvda.errors() == []
    assert nvda.v("Synthese.ctrl_glob", "B") == "OK"


def test_cash_reconciliation_closes_with_full_cash_flow_inputs(nvda):
    """Avec investissements financiers, impôts sur RSU, SBC et gains latents saisis, l'écart FY2026 passe sous 2 % du CA."""
    assert abs(nvda.v("Cashflow.dt_ecart", "D")) < 0.02 * nvda.v("SIG_CAF.ca", "D")


def test_sbc_adjusted_ebitda(nvda):
    assert nvda.v("SIG_CAF.ebe_aj", "D") == pytest.approx(133230 + 6386, abs=0.5)


def test_panel_has_three_dated_peers(nvda):
    """Le panel v1.3 compte au moins 3 pairs exploitables (Broadcom, AMD, TSMC) ; la médiane n'est plus un point unique."""
    ws = nvda.wb["Synthese"]
    vals = [ws[f"B{r}"].value for r in range(1, 120) if ws[f"A{r}"].value == "Panel de comparables : au moins 3 pairs retenus"]
    assert vals and vals[0] == "OK"


def test_transactions_multiples(nvda):
    """5 transactions réelles ; 3 avec EBITDA calculé (Fairchild, Microsemi, Lumentum) : médiane VE/EBITDA ≈ 17,6x ; VE/CA ≈ 5,6x."""
    ws = nvda.wb["Transactions"]
    assert ws["B17"].value == 5
    assert 17.0 < ws["B18"].value < 18.0
    assert 5.0 < ws["B22"].value < 6.0
    assert ws["B21"].value == "n/d"  # seulement 2 primes documentées


def test_valuation_rolled_to_q3_balance(nvda):
    """Bilan du 25/10/2026 : actifs non opérationnels = titres non cotés 65 528 + trésorerie nette générée 11 272."""
    ws = nvda.wb["DCF_10ans"]
    assert ws["B39"].value == -76800
