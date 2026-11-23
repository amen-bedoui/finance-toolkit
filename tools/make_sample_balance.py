#!/usr/bin/env python3
"""Génère des balances générales synthétiques (PCG) à partir de data/examples/societe_fictive.json.
Sert d'exemple pour import_balance.py et de test d'aller-retour. Données 100 % fictives."""
import csv
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "data", "examples", "societe_fictive.json")
OUT = os.path.join(HERE, "..", "data", "samples")
# clé Saisie -> (compte, nature) ; D = solde débiteur, C = créditeur
ACC = {"ca": ("701000", "C"), "achats": ("601000", "D"), "ext": ("622000", "D"), "taxes": ("635000", "D"), "perso": ("641000", "D"),
       "aprod": ("758000", "C"), "acharg": ("658000", "D"), "dot": ("681000", "D"), "pfin": ("761000", "C"), "cfin": ("661000", "D"),
       "pexc": ("771000", "C"), "cexc": ("671000", "D"), "impot": ("695000", "D"), "immo_inc": ("205000", "D"), "immo_corp": ("218000", "D"),
       "immo_fin": ("261000", "D"), "stocks": ("370000", "D"), "clients": ("411000", "D"), "autres_cr": ("409600", "D"), "treso": ("512000", "D"),
       "capital": ("101000", "C"), "reserves": ("106800", "C"), "provisions": ("151000", "C"), "dlt": ("164000", "C"), "dct": ("519000", "C"),
       "fourn": ("401000", "C"), "fisc": ("447000", "C"), "autres_det": ("467000", "C")}


def main():
    d = json.load(open(SRC, encoding="utf-8"))
    os.makedirs(OUT, exist_ok=True)
    for i, y in enumerate(d["years"]):
        rows = []
        for k, (acc, nat) in ACC.items():
            v = d["statements"].get(k, [0, 0, 0])[i]
            if v:
                rows.append((acc, k, v if nat == "D" else 0, v if nat == "C" else 0))
        path = os.path.join(OUT, f"balance_{y.replace('/', '-')}.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Compte", "Libellé", "Débit", "Crédit"])
            for acc, k, dd, cc in rows:
                w.writerow([acc, f"Compte synthétique ({k})", dd, cc])
        print("écrit", path)


if __name__ == "__main__":
    sys.exit(main())
