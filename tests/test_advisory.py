"""Fonctions « conseil / marchés » : DCF 10 ans, DCF inversé, transactions, note de marché."""
import pytest


def test_no_errors(nvda, wave, fict):
    for b in (nvda, wave, fict):
        assert b.errors() == []


def test_fade_defaults_to_flat_margin_without_inputs(wave):
    """Sans cible saisie, la marge de l'année 10 = marge de l'année 5 et la croissance converge vers g."""
    ws = wave.wb["DCF_10ans"]
    assert ws["L15"].value == pytest.approx(ws["G15"].value)
    assert ws["L13"].value == pytest.approx(wave.wb["Hypotheses"][f"B{wave.ix['Hypotheses.g'][1]}"].value)


def test_fade_applies_target_margin_and_growth_converges(nvda):
    ws = nvda.wb["DCF_10ans"]
    assert ws["L15"].value == pytest.approx(0.50)
    g = [ws[f"{c}13"].value for c in "GHIJKL"]
    assert all(a >= b for a, b in zip(g, g[1:]))  # croissance décroissante
    assert g[-1] == pytest.approx(0.03)


def test_longer_horizon_lowers_terminal_weight(nvda):
    """Le DCF 10 ans réduit le poids de la valeur terminale (Gordon) par rapport au DCF 5 ans."""
    base_tv = nvda.v("DCF.base.s_tvpg", "B")
    assert nvda.v("DCF10.tvpg", "B") < base_tv


def test_reverse_dcf_reproduces_price(nvda):
    """À l'intérieur de la grille, le taux implicite doit encadrer le cours : valeur(w-) ≥ cours ≥ valeur(w+)."""
    ws = nvda.wb["DCF_10ans"]
    w = nvda.v("DCF10.implied", "B")
    px = nvda.wb["Saisie"][f"B{nvda.ix['Saisie.cours'][1]}"].value
    rows = [(ws[f"A{r}"].value, ws[f"B{r}"].value) for r in range(57, 98)]
    lo = max(v for a, v in rows if a <= w + 1e-12)
    hi = min(v for a, v in rows if a >= w - 1e-12)
    assert hi - 1e-6 <= px <= lo + 1e-6
    # valeur décroissante avec le taux
    vals = [v for _, v in rows]
    assert all(a > b for a, b in zip(vals, vals[1:]))


def test_target_price_is_weighted_average_of_available_methods(nvda):
    ws = nvda.wb["Reco_Marche"]
    vals = [(ws[f"B{r}"].value, ws[f"C{r}"].value) for r in range(6, 10)
            if isinstance(ws[f"B{r}"].value, (int, float)) and ws[f"G{r}"].value != "Écartée"]
    haircut = nvda.wb["Marche"]["B25"].value or 0  # décote juridique saisie dans Marche
    expected = sum(v * w for v, w in vals) / sum(w for _, w in vals) * (1 - haircut)
    assert nvda.v("Reco.tp", "B") == pytest.approx(expected)


def test_consensus_bridge_is_consistent(nvda):
    ws = nvda.wb["Pont_Consensus"]
    before, haircut, after = ws["B6"].value, ws["B7"].value, ws["B8"].value
    assert after == pytest.approx(before * (1 - haircut))
    assert ws["B33"].value == pytest.approx(1 - ws["B5"].value / before)
    assert ws["D16"].value == pytest.approx(after)


def test_rating_follows_thresholds(nvda):
    up, buy, sell = nvda.v("Reco.up", "B"), nvda.v("Reco.buy", "B"), nvda.v("Reco.sell", "B")
    want = "FAVORABLE" if up >= buy else "DÉFAVORABLE" if up <= sell else "NEUTRE"
    assert nvda.v("Reco.rating", "B") == want


def test_transactions_panel_needs_three_deals(wave):
    """Sans transactions saisies (l'exemple fictif est exclu), le classeur affiche n/d, jamais une fourchette."""
    assert wave.v("Tx.n", "B") == 0
    assert wave.v("Tx.ps", "B") == "n/d"


def test_credit_rating_net_cash_company(nvda):
    assert nvda.v("Reco.nd", "B") < 0
    assert nvda.v("Reco.crd", "B") == "AAA"


def test_football_field_includes_new_methods(nvda):
    ws = nvda.wb["Synthese"]
    labels = [ws[f"A{r}"].value for r in range(nvda.ix["Synthese.ff_first"][1], nvda.ix["Synthese.ff_last"][1] + 1)]
    assert any("DCF 10 ans" in l for l in labels) and any("Transactions" in l for l in labels)


# ---- Sources, Marché, Qualitatif, conviction
def _statuses(book):
    ws = book.wb["Sources"]
    return {ws[f"A{r}"].value: ws[f"H{r}"].value for r in range(5, 20)}


def test_source_statuses_nvidia(nvda):
    st = _statuses(nvda)
    assert st["Taux sans risque"] == "Documentée"
    assert st["Prime de risque du marché"] == "Documentée"  # Damodaran 01/09/2026
    assert st["Bêta retenu"] == "Source ancienne"  # Damodaran janvier 2026 : > 180 jours
    assert st["Croissance perpétuelle"] == "Jugement"
    assert st["Nombre d'actions"] == "Documentée"
    assert st["Marge d'EBITDA, année 5 (Base)"] == "Source ancienne"  # panel FY2023-24 : ancien


def test_documented_share_matches_statuses(nvda):
    st = _statuses(nvda)
    expected = sum(1 for v in st.values() if v == "Documentée") / len(st)
    assert nvda.v("Src.pct", "B") == pytest.approx(expected)


def test_price_age_and_status(nvda):
    assert nvda.v("Mk.age", "B") == 3  # clôture du 20/11/2026 pour une analyse au 23/11/2026
    ws = nvda.wb["Marche"]
    assert ws["B8"].value == "OK"


def test_stale_price_is_flagged(wave):
    """Un cours sans date ne peut pas être présenté comme à jour."""
    ws = wave.wb["Marche"]
    assert ws["B8"].value != "OK"


def test_qualitative_profile_follows_threshold(nvda):
    ws = nvda.wb["Qualitatif"]
    mx, thr = ws["B36"].value, ws["B39"].value
    want = "ÉLEVÉ" if mx >= thr else "MODÉRÉ" if mx >= 0.6 * thr else "FAIBLE"
    assert nvda.v("Ql.prof", "B") == want
    # score = proba x impact
    assert ws["E14"].value == ws["C14"].value * ws["D14"].value


def test_conviction_counts_ok_checks(nvda):
    ws = nvda.wb["Reco_Marche"]
    n_ok = sum(1 for r in range(38, 46) if ws[f"B{r}"].value == "OK")
    assert ws["B48"].value == n_ok
    lvl = "ÉLEVÉE" if n_ok / 8 >= 0.75 else "MOYENNE" if n_ok / 8 >= 0.5 else "FAIBLE"
    assert ws["B49"].value == lvl


def test_no_sources_means_low_conviction_and_no_errors(wave):
    """Une société sans registre de sources ne peut pas afficher une conviction élevée."""
    assert wave.errors() == []
    st = _statuses(wave)
    assert all(v == "Non sourcée" for v in st.values())
    assert wave.wb["Reco_Marche"]["B49"].value in ("FAIBLE", "MOYENNE")
    assert wave.wb["Marche"]["F13"].value == "n/d"  # pas de consensus saisi


def test_outlier_method_is_excluded(nvda):
    ws = nvda.wb["Reco_Marche"]
    assert ws["G9"].value == "Écartée" and ws["D9"].value == 0
    assert ws["G6"].value == "Retenue"


def test_peer_multiple_sensitivity_zero_shift_matches_target(nvda):
    ws = nvda.wb["Pont_Consensus"]
    assert ws["A38"].value == 0
    assert ws["C38"].value == pytest.approx(ws["B8"].value)
    assert ws["C41"].value < ws["C38"].value
