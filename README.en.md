# finance-toolkit — automated financial analysis and valuation

> **Disclaimer.** Educational analysis tool. Valuation assumptions in the examples are illustrative, peer data comes from free and sometimes dated sources, and the "assessment" in the market note is a mechanical rule. Nothing here is investment advice.

Paste a balance sheet and income statement (3 years) into an Excel template: income-statement balances, cash-flow capacity, working capital, ratios, Altman/Piotroski scores, DuPont, YoY variations, a **3-scenario DCF**, **trading comps / football field**, **peer betas**, **LBO**, **accretion/dilution**, a **10-year DCF with fade**, a **reverse DCF** (return implied by the share price), **precedent transactions** and a **market note** (target price, risk/reward, synthetic credit) compute automatically (live formulas, no macros). A script then builds a PDF report.

```bash
pip install -r requirements.txt
python src/build_workbook.py --blank --out Template.xlsx
python src/build_workbook.py --data data/examples/wavestone_public.json --out Wavestone.xlsx
python src/build_workbook.py --data data/examples/nvidia_public.json --out NVIDIA.xlsx   # US GAAP, hypergrowth
python src/analyse.py Wavestone.xlsx Report.pdf
python src/import_balance.py --template Template.xlsx --out Case.xlsx --files tb_N-2.csv tb_N-1.csv tb_N.csv   # trial balance / FEC (French PCG mapping, editable)
python -m pytest -q
```

Validated against published figures of three listed companies (Wavestone, LVMH, NVIDIA): see `docs/VALIDATION.md`, including definitional differences and known limits. Valuation assumptions are illustrative and peer data comes from free aggregators. Educational tool, not investment advice. MIT licence.
