"""Validation sur une société réelle cotée (Wavestone SA, états consolidés IFRS, M€).
Compare les sorties du classeur aux chiffres publiés dans les communiqués de résultats.
Voir docs/VALIDATION.md pour les écarts de définition expliqués."""
import pytest

# (libellé, clé, colonne, valeur publiée, tolérance absolue)
PUBLISHED = [
    ("CA 2025/26", "SIG_CAF.ca", "D", 954.347, 0.001),
    ("CA 2024/25", "SIG_CAF.ca", "C", 943.666, 0.001),
    ("Résultat net 2025/26", "SIG_CAF.rn", "D", 82.100, 0.001),
    ("Résultat net 2024/25", "SIG_CAF.rn", "C", 75.921, 0.01),
    ("Total bilan 2025/26", "Bilan_BFR.ta", "D", 996.894, 0.001),
    ("Capitaux propres 2025/26", "Bilan_BFR.cp", "D", 702.465, 0.001),
    ("EBITDA = résultat opérationnel courant + dotations nettes 2025/26", "SIG_CAF.ebe", "D", 132.758, 0.001),
    ("Trésorerie nette hors dettes locatives 2025/26 (publiée : 121,4)", "Bilan_BFR.dn16", "D", -121.4, 0.1),
]


@pytest.mark.parametrize("label,key,col,published,tol", PUBLISHED, ids=[p[0] for p in PUBLISHED])
def test_against_published(wave, label, key, col, published, tol):
    assert wave.v(key, col) == pytest.approx(published, abs=tol)


def test_cash_reconciliation_small_gap(wave):
    """Après acquisitions et rachats saisis, l'écart de rapprochement de trésorerie 2025/26 reste < 2 M€."""
    assert abs(wave.v("Cashflow.dt_ecart", "D")) < 2.0
