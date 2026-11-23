"""Import de balance : aller-retour balance synthétique -> Saisie -> mêmes montants que le JSON source."""
import json
import os
import subprocess
import sys

from openpyxl import load_workbook

from conftest import ROOT

KEYS = ["ca", "achats", "ext", "taxes", "perso", "dot", "pfin", "cfin", "impot", "immo_inc", "immo_corp", "immo_fin", "stocks", "clients",
        "autres_cr", "treso", "capital", "reserves", "provisions", "dlt", "dct", "fourn", "fisc", "autres_det"]


def test_roundtrip(tmp_path):
    subprocess.run([sys.executable, os.path.join(ROOT, "tools", "make_sample_balance.py")], check=True, capture_output=True)
    tpl = tmp_path / "t.xlsx"
    subprocess.run([sys.executable, os.path.join(ROOT, "src", "build_workbook.py"), "--blank", "--out", str(tpl)], check=True, capture_output=True)
    out = tmp_path / "o.xlsx"
    files = [os.path.join(ROOT, "data", "samples", f"balance_{y}.csv") for y in ("2023", "2024", "2025")]
    subprocess.run([sys.executable, os.path.join(ROOT, "src", "import_balance.py"), "--template", str(tpl), "--out", str(out), "--files", *files],
                   check=True, capture_output=True)
    wb = load_workbook(out)
    ix = {r[0]: r[2] for r in wb["_Index"].iter_rows(min_row=2, values_only=True) if r[0]}
    ref = json.load(open(os.path.join(ROOT, "data", "examples", "societe_fictive.json"), encoding="utf-8"))["statements"]
    for k in KEYS:
        for i, col in enumerate("BCD"):
            got = wb["Saisie"][f"{col}{ix['Saisie.' + k]}"].value or 0
            assert abs(got - ref[k][i]) < 1e-6, (k, col, got, ref[k][i])
