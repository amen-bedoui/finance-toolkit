"""Validation sur LVMH (comptes consolidés 2025, M€). Voir docs/VALIDATION_LVMH.md."""
import pytest

PUBLISHED = [
    ("Ventes 2025", "SIG_CAF.ca", "D", 80807, 0.5),
    ("Ventes 2024", "SIG_CAF.ca", "C", 84683, 0.5),
    ("Ventes 2023", "SIG_CAF.ca", "B", 86153, 0.5),
    ("Résultat opérationnel courant 2025", "SIG_CAF.ebit", "D", 17755, 0.5),
    ("Résultat opérationnel courant 2024", "SIG_CAF.ebit", "C", 19571, 0.5),
    ("Résultat opérationnel courant 2023", "SIG_CAF.ebit", "B", 22802, 0.5),
    ("Résultat net avant minoritaires 2025", "SIG_CAF.rn", "D", 11222, 0.5),
    ("Résultat net avant minoritaires 2024", "SIG_CAF.rn", "C", 12958, 0.5),
    ("Résultat net avant minoritaires 2023", "SIG_CAF.rn", "B", 15952, 0.5),
    ("Total bilan 2025", "Bilan_BFR.ta", "D", 142037, 0.5),
    ("Total bilan 2024", "Bilan_BFR.ta", "C", 149190, 0.5),
    ("Total bilan 2023", "Bilan_BFR.ta", "B", 143694, 0.5),
    ("Capitaux propres 2025 (avec minoritaires)", "Bilan_BFR.cp", "D", 68949, 0.5),
    ("EBITDA 2025 = ROC + dotations (dont droits d'utilisation)", "SIG_CAF.ebe", "D", 25756, 0.5),
    ("Emprunts nets hors dettes locatives 2025 (calcul à partir du bilan)", "Bilan_BFR.dn16", "D", 11549, 0.5),
]


@pytest.mark.parametrize("label,key,col,published,tol", PUBLISHED, ids=[p[0] for p in PUBLISHED])
def test_against_published(lvmh, label, key, col, published, tol):
    assert lvmh.v(key, col) == pytest.approx(published, abs=tol)


def test_controls_and_no_errors(lvmh):
    assert lvmh.errors() == []
    assert lvmh.v("Synthese.ctrl_glob", "B") == "OK"


def test_company_net_debt_gap_is_documented(lvmh):
    """La dette financière nette publiée (6 857) est inférieure aux emprunts nets du bilan : elle déduit aussi des placements financiers courants.
    Le test verrouille l'écart attendu (4,7 Md€) pour qu'une dérive de mapping soit détectée."""
    assert lvmh.v("Bilan_BFR.dn16", "D") - 6857 == pytest.approx(4692, abs=0.5)


def test_dcf_scenarios_ordered(lvmh):
    bear, base, bull = (lvmh.v(f"DCF.{t}.s_ps", "B") for t in ("bear", "base", "bull"))
    assert bear < base < bull
