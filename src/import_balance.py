#!/usr/bin/env python3
"""Importe une balance générale (CSV/XLSX) ou un FEC dans le classeur modèle (onglet Saisie).

Usage :
  python import_balance.py --template Modele_vierge.xlsx --out Mon_dossier.xlsx \
      --files balance_N-2.csv balance_N-1.csv balance_N.csv [--mapping ../data/pcg_mapping.csv]

Un fichier par exercice, du plus ancien au plus récent. Formats reconnus :
  * balance : colonnes Compte / Libellé / Débit / Crédit (ou Solde débiteur / Solde créditeur) ;
  * FEC : colonnes CompteNum, Debit, Credit (séparateur tabulation, point-virgule ou barre verticale).
Le FEC doit contenir les à-nouveaux (journal AN) pour que le bilan soit complet.
Les comptes non reconnus sont listés (jamais ignorés en silence). La ventilation dette > 1 an / < 1 an,
les dettes locatives IFRS 16, les capex, dividendes et acquisitions restent à saisir à la main.
"""
import argparse
import csv
import io
import os
import sys
from collections import defaultdict

from openpyxl import load_workbook

NATURE_RES = ("6", "7")


def num(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(" ", "").replace(" ", "")
    if s in ("", "-"):
        return 0.0
    s = s.replace(",", ".") if s.count(",") == 1 and "." not in s else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_rows(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        ws = load_workbook(path, data_only=True).worksheets[0]
        rows = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        raw = open(path, encoding="utf-8-sig", errors="replace").read()
        first = raw.splitlines()[0]
        delim = max(["\t", ";", "|", ","], key=first.count)
        rows = list(csv.reader(io.StringIO(raw), delimiter=delim))
    return rows


def balance_from_file(path):
    """Retourne {compte: solde débiteur (+) / créditeur (-)}."""
    rows = read_rows(path)
    head = [str(h or "").strip().lower() for h in rows[0]]

    def col(*names):
        for n in names:
            for i, h in enumerate(head):
                if h == n.lower():
                    return i
        return None

    c_acc = col("compteNum", "compte", "numéro de compte", "numero de compte", "n° compte")
    if c_acc is None:
        sys.exit(f"{path} : colonne de compte introuvable (attendu : CompteNum / Compte).")
    c_d, c_c = col("debit", "débit", "solde débiteur", "solde debiteur"), col("credit", "crédit", "solde créditeur", "solde crediteur")
    c_s = col("solde")
    bal = defaultdict(float)
    for r in rows[1:]:
        if not r or c_acc >= len(r) or not str(r[c_acc] or "").strip():
            continue
        acc = "".join(ch for ch in str(r[c_acc]).strip() if ch.isalnum())
        if not acc[:1].isdigit():
            continue
        if c_d is not None and c_c is not None:
            bal[acc] += num(r[c_d]) - num(r[c_c])
        elif c_s is not None:
            bal[acc] += num(r[c_s])
    return bal


def load_mapping(path):
    m = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            if not line.strip() or line.startswith("#") or line.startswith("prefix"):
                continue
            p = line.rstrip("\n").split(";")
            m.append((p[0].strip(), p[1].strip(), p[2].strip(), (p[3].strip() if len(p) > 3 else "")))
    return sorted(m, key=lambda t: -len(t[0]))


def map_year(bal, mapping):
    out, unmapped = defaultdict(float), {}
    has_pl = any(a[0] in NATURE_RES and abs(v) > 1e-9 for a, v in bal.items())
    for acc, v in bal.items():
        if abs(v) < 1e-9 or acc[0] in "89":
            continue
        hit = next((t for t in mapping if acc.startswith(t[0])), None)
        if hit is None:
            unmapped[acc] = v
            continue
        _, key, nat, key_c = hit
        if nat == "D":
            out[key] += v
        elif nat == "C":
            out[key] += -v
        else:  # M
            if v >= 0:
                out[key] += v
            else:
                out[key_c] += -v
    # bilan avant clôture : le résultat est le solde des comptes 6/7
    if has_pl and abs(out.get("res_bilan", 0)) < 1e-9:
        out["res_bilan"] = -sum(v for a, v in bal.items() if a[0] in NATURE_RES)
    return dict(out), unmapped


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--files", nargs="+", required=True, help="1 à 3 fichiers, du plus ancien au plus récent")
    ap.add_argument("--mapping", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "pcg_mapping.csv"))
    ap.add_argument("--unit", type=float, default=1.0, help="diviseur pour changer d'unité (1000 : € -> k€)")
    a = ap.parse_args()
    if not 1 <= len(a.files) <= 3:
        sys.exit("Fournir 1 à 3 fichiers (un par exercice).")
    wb = load_workbook(a.template)
    ix = {r[0]: r[2] for r in wb["_Index"].iter_rows(min_row=2, values_only=True) if r[0]}
    ws = wb["Saisie"]
    mapping = load_mapping(a.mapping)
    cols = ["B", "C", "D"][-len(a.files):]
    report = []
    for col, path in zip(cols, a.files):
        values, unmapped = map_year(balance_from_file(path), mapping)
        for k, v in values.items():
            row = ix.get(f"Saisie.{k}")
            if row is None:
                continue
            ws[f"{col}{row}"].value = round(v / a.unit, 6)
        actif = sum(values.get(k, 0) for k in ["immo_inc", "immo_corp", "immo_fin", "stocks", "clients", "autres_cr", "treso"])
        passif = sum(values.get(k, 0) for k in ["capital", "reserves", "res_bilan", "provisions", "dlt", "dct", "fourn", "fisc", "autres_det"])
        report.append((os.path.basename(path), (actif - passif) / a.unit, unmapped))
    wb.save(a.out)
    print(f"Classeur écrit : {a.out}")
    for name, gap, unm in report:
        flag = "OK" if abs(gap) < 1 else "ÉCART — vérifier comptes non reconnus / à-nouveaux"
        print(f"  {name}: actif - passif = {gap:,.2f}  [{flag}]")
        if unm:
            print("    comptes non reconnus :", ", ".join(f"{k} ({v / a.unit:,.0f})" for k, v in sorted(unm.items())))
    print("À compléter à la main : ventilation LT/CT des dettes, dettes locatives IFRS 16, loyers, capex, dividendes, acquisitions.")
    print("Recalculer ensuite le classeur (Excel l'ouvre et calcule ; sinon LibreOffice).")


if __name__ == "__main__":
    main()
