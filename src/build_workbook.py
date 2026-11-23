#!/usr/bin/env python3
"""Génère le classeur Excel « Outil d'analyse financière » (formules vivantes).

L'utilisateur ne remplit que les onglets Saisie et Hypotheses ; tout le reste
(SIG, CAF, BFR, cash-flows, ratios, scores, variations, DCF 3 scénarios,
sensibilités, synthèse) se recalcule par formules Excel.
"""
import argparse
import datetime as _dt
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

_ap = argparse.ArgumentParser(description="Génère le classeur d'analyse financière (formules vivantes).")
_ap.add_argument("--data", default=None, help="Fichier JSON de données (défaut : société fictive d'exemple)")
_ap.add_argument("--out", default="Outil_Analyse_Financiere.xlsx")
_ap.add_argument("--blank", action="store_true", help="Classeur vierge : efface les données de la société (garde les hypothèses)")
ARGS = _ap.parse_args()
OUT = ARGS.out
_default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "examples", "societe_fictive.json")
with open(ARGS.data or _default, encoding="utf-8") as _f:
    D = json.load(_f)
INFO, EXD, HYD = D["info"], D["statements"], D["hypotheses"]
YEARS3 = D["years"]

FONT = "Arial"
BLUE, GREEN, BLACK = "0000FF", "008000", "000000"
NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0"x";(0.0"x");"-"'
D2 = '0.00;(0.00);"-"'
D1 = '0.0;(0.0);"-"'
INT0 = '0'

fill_hdr = PatternFill("solid", fgColor="1F3864")
fill_sec = PatternFill("solid", fgColor="D9E2F3")
fill_inp = PatternFill("solid", fgColor="FFF2CC")
fill_tot = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
COLS = ["B", "C", "D"]  # N-2, N-1, N


def font(bold=False, color=BLACK, size=10, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


class Sh:
    def __init__(self, ws, widths):
        self.ws, self.n, self.r, self.cur = ws, ws.title, {}, 1
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws.sheet_view.showGridLines = False

    def title(self, text, sub=None):
        c = self.ws.cell(1, 1, text)
        c.font = font(True, "1F3864", 14)
        if sub:
            self.ws.cell(2, 1, sub).font = font(italic=True, color="595959")
        self.cur = 4

    def section(self, text, last="D"):
        r = self.cur
        for col in range(1, ord(last) - 64 + 1):
            self.ws.cell(r, col).fill = fill_sec
        self.ws.cell(r, 1, text).font = font(True, "1F3864")
        self.cur += 1
        return r

    def header(self, labels, start_col=2, formulas=False):
        r = self.cur
        self.ws.cell(r, 1).fill = fill_hdr
        for i, lab in enumerate(labels):
            c = self.ws.cell(r, start_col + i, lab)
            c.font = font(True, "FFFFFF")
            c.fill = fill_hdr
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        self.cur += 1
        return r

    def gap(self, n=1):
        self.cur += n

    def ref(self, k, c):
        return f"{self.n}!{c}{self.r[k]}"

    def loc(self, k, c):
        return f"{c}{self.r[k]}"

    def line(self, key, label, f, fmt=NUM, first=None, bold=False, total=False,
             link=False, cols=COLS, indent=0):
        r = self.cur
        self.r[key] = r
        self.cur += 1
        a = self.ws.cell(r, 1, label)
        a.font = font(bold)
        a.alignment = Alignment(indent=indent)
        for i, c in enumerate(cols):
            p = cols[i - 1] if i > 0 else None
            v = first if (i == 0 and first is not None) else f(c, p)
            cell = self.ws[f"{c}{r}"]
            cell.value = v
            cell.number_format = fmt
            cell.font = font(bold, GREEN if link else BLACK)
            cell.alignment = Alignment(horizontal="right")
            if total:
                cell.fill = fill_tot
        if total:
            a.fill = fill_tot
        return r


wb = Workbook()
wb.remove(wb.active)
ws_lm = wb.create_sheet("LisezMoi")
ws_sa = wb.create_sheet("Saisie")
ws_hy = wb.create_sheet("Hypotheses")
ws_sy = wb.create_sheet("Synthese")
ws_sig = wb.create_sheet("SIG_CAF")
ws_bf = wb.create_sheet("Bilan_BFR")
ws_cf = wb.create_sheet("Cashflow")
ws_ra = wb.create_sheet("Ratios")
ws_sc = wb.create_sheet("Scores")
ws_va = wb.create_sheet("Variations")
ws_dc = wb.create_sheet("DCF")
ws_cp = wb.create_sheet("Comparables")
ws_wp = wb.create_sheet("WACC_Pairs")
ws_lbo = wb.create_sheet("LBO")
ws_ma = wb.create_sheet("MA")
ws_d10 = wb.create_sheet("DCF_10ans")
ws_tx = wb.create_sheet("Transactions")
ws_so = wb.create_sheet("Sources")
ws_mk = wb.create_sheet("Marche")
ws_ql = wb.create_sheet("Qualitatif")
ws_rc = wb.create_sheet("Reco_Marche")
ws_pc = wb.create_sheet("Pont_Consensus")
ws_se = wb.create_sheet("Seuils")
ws_ix = wb.create_sheet("_Index")

# =====================================================================
# SAISIE
# =====================================================================
SA = Sh(ws_sa, {"A": 46, "B": 15, "C": 15, "D": 15, "E": 62})
SA.title("SAISIE — Compte de résultat et bilan",
         "Cellules jaunes/bleues = à remplir. Les charges se saisissent en POSITIF. "
         "Tout le reste du classeur se calcule automatiquement.")

ex_info = [
    ("nom", "Nom de la société", INFO["nom"]),
    ("secteur", "Secteur d'activité", INFO["secteur"]),
    ("devise", "Devise", INFO["devise"]),
    ("unite", "Unité des montants (libellé)", INFO["unite"]),
    ("mult", "Multiplicateur d'unité vers la devise (1 ; 1 000 ; 1 000 000)", INFO["mult"]),
    ("actions", "Nombre d'actions en circulation (optionnel, en unités)", INFO.get("actions")),
    ("cours", "Cours / prix par action actuel (optionnel, en devise)", INFO.get("cours")),
    ("tva", "Taux de TVA appliqué aux créances/dettes TTC (pour DSO/DPO ; 0 si HT)", INFO.get("tva", 0.0)),
    ("profil", "Profil de seuils des ratios (Générique / Services / Industrie)", INFO.get("profil", "Générique")),
]
SA.section("INFORMATIONS GÉNÉRALES", "E")
for k, lab, v in ex_info:
    r = SA.cur
    SA.r[k] = r
    SA.ws.cell(r, 1, lab).font = font()
    c = SA.ws.cell(r, 2, v)
    c.font = font(color=BLUE)
    c.fill = fill_inp
    c.border = box
    if k in ("mult", "actions"):
        c.number_format = '#,##0'
    if k == "cours":
        c.number_format = '#,##0.00'
    if k == "tva":
        c.number_format = '0.0%'
    SA.cur += 1
SA.ws.cell(SA.r["actions"], 5, "Laisser vide si société non cotée : la valeur par action ne sera pas calculée.").font = font(italic=True, color="595959")
SA.gap()

# ligne des exercices
r = SA.cur
SA.r["annee"] = r
SA.ws.cell(r, 1, "Exercice (clôture)").font = font(True)
for c, y in zip(COLS, YEARS3):
    cell = SA.ws[f"{c}{r}"]
    cell.value = y
    cell.font = font(True, BLUE)
    cell.fill = fill_inp
    cell.border = box
    cell.alignment = Alignment(horizontal="center")
SA.ws.cell(r, 5, "N-2, N-1, N : du plus ancien au plus récent. Saisir en texte (ex. 2025).").font = font(italic=True, color="595959")
SA.cur += 2

_KEYS = ["ca", "achats", "ext", "taxes", "perso", "aprod", "acharg", "dot", "pfin", "cfin", "pexc", "cexc", "impot",
         "immo_inc", "immo_corp", "immo_fin", "stocks", "clients", "autres_cr", "treso", "capital", "reserves",
         "res_bilan", "provisions", "dlt", "dct", "fourn", "fisc", "autres_det", "lease_lt", "lease_ct", "loyers",
         "capex", "div", "acq", "rach", "sbc", "autres_inv", "autres_dec", "noncash"]
EX = {k: ([None] * 3 if ARGS.blank else EXD.get(k, [0, 0, 0])) for k in _KEYS}


def inp(key, label, note=""):
    r = SA.cur
    SA.r[key] = r
    SA.cur += 1
    SA.ws.cell(r, 1, label).font = font()
    for c, v in zip(COLS, EX[key]):
        cell = SA.ws[f"{c}{r}"]
        cell.value = v
        cell.font = font(color=BLUE)
        cell.fill = fill_inp
        cell.border = box
        cell.number_format = NUM
    SA.ws.cell(r, 5, note).font = font(italic=True, color="595959")


SA.section("COMPTE DE RÉSULTAT (charges en positif)", "E")
SA.header(["N-2", "N-1", "N"])
SA.ws.cell(SA.cur - 1, 5, "Repère plan comptable général (PCG)").font = font(True, "FFFFFF")
SA.ws.cell(SA.cur - 1, 5).fill = fill_hdr
inp("ca", "Chiffre d'affaires net", "Comptes 70")
inp("achats", "Achats consommés (marchandises, matières, variation de stocks)", "601-607 ± variation de stocks 603 / 71")
inp("ext", "Autres achats et charges externes", "61-62")
inp("taxes", "Impôts, taxes et versements assimilés", "63")
inp("perso", "Charges de personnel", "64")
inp("aprod", "Autres produits d'exploitation (y c. reprises)", "74-75, 781")
inp("acharg", "Autres charges d'exploitation", "65")
inp("dot", "Dotations aux amortissements et provisions d'exploitation", "681")
inp("pfin", "Produits financiers", "76, 786")
inp("cfin", "Charges financières (intérêts, escomptes…)", "66, 686")
inp("pexc", "Produits exceptionnels", "77 (ou éléments non récurrents)")
inp("cexc", "Charges exceptionnelles", "67 (ou éléments non récurrents)")
inp("impot", "Impôt sur les bénéfices", "695, 698 (hors participation salariés : à inclure dans les charges de personnel)")
SA.gap()
SA.section("BILAN — ACTIF (valeurs NETTES)", "E")
SA.header(["N-2", "N-1", "N"])
inp("immo_inc", "Immobilisations incorporelles nettes", "20 - 280/290")
inp("immo_corp", "Immobilisations corporelles nettes", "21-23 - 281/291")
inp("immo_fin", "Immobilisations financières", "26-27")
inp("stocks", "Stocks et en-cours", "31-37 (net de dépréciations)")
inp("clients", "Créances clients et comptes rattachés", "411-416 (net de 491)")
inp("autres_cr", "Autres créances et charges constatées d'avance", "42-46, 48 débiteurs, 486")
inp("treso", "Trésorerie et équivalents (VMP + disponibilités)", "50-53")
r = SA.line("tot_actif", "TOTAL ACTIF",
            lambda c, p: f"=SUM({c}{SA.r['immo_inc']}:{c}{SA.r['treso']})", bold=True, total=True)
SA.gap()
SA.section("BILAN — PASSIF", "E")
SA.header(["N-2", "N-1", "N"])
inp("capital", "Capital social et primes", "10 (y c. primes 104)")
inp("reserves", "Réserves et report à nouveau (+ subventions, écarts)", "11, 13, 14 (hors résultat de l'exercice)")
inp("res_bilan", "Résultat de l'exercice (au bilan)", "12 : doit égaler le résultat net du compte de résultat")
inp("provisions", "Provisions pour risques et charges", "15")
inp("dlt", "Emprunts et dettes financières à plus d'un an", "16-17 (part > 1 an)")
inp("dct", "Dettes financières à moins d'un an (dont concours bancaires)", "519, 512 créditeur, part < 1 an des emprunts")
inp("fourn", "Dettes fournisseurs et comptes rattachés", "40")
inp("fisc", "Dettes fiscales et sociales", "42-44")
inp("autres_det", "Autres dettes et produits constatés d'avance", "45-47, 487 (dont dettes sur immobilisations)")
SA.line("tot_passif", "TOTAL PASSIF",
        lambda c, p: f"=SUM({c}{SA.r['capital']}:{c}{SA.r['autres_det']})", bold=True, total=True)
SA.line("ecart_bilan", "Contrôle : Actif − Passif (doit être 0)",
        lambda c, p: f"={c}{SA.r['tot_actif']}-{c}{SA.r['tot_passif']}")
SA.gap()
SA.section("AUTRES INFORMATIONS (tableau de flux / annexe)", "E")
SA.header(["N-2", "N-1", "N"])
inp("capex", "Investissements de l'exercice (acquisitions d'immobilisations)",
    "Optionnel pour N-1 et N : vide = estimé (Δ immobilisations nettes + dotations). Requis pour N-2.")
inp("div", "Dividendes versés dans l'exercice", "Prélevés sur le résultat de l'exercice précédent")
inp("acq", "Acquisitions de filiales (décaissé net, hors trésorerie acquise)", "Optionnel : améliore le rapprochement de trésorerie")
inp("rach", "Rachats d'actions propres / opérations sur capital (net)", "Optionnel : décaissement positif")
inp("autres_inv", "Autres investissements financiers nets (titres non cotés, licences/actifs acquis, hors capex et acquisitions)", "Optionnel : décaissement net positif ; améliore le rapprochement de trésorerie")
inp("autres_dec", "Autres décaissements nets sur capitaux (impôts sur actions gratuites, remboursements, divers)", "Optionnel : décaissement net positif (ex. impôts payés sur RSU en US GAAP)")
inp("noncash", "Gains latents et autres produits non décaissés inclus dans le résultat", "Optionnel : plus-values latentes sur titres, etc. (positif = produit sans encaissement)")
inp("sbc", "Rémunération en actions (charge non décaissée, déjà incluse dans les charges)", "Optionnel : permet l'EBITDA ajusté et le ratio rémunération en actions / CA (US GAAP : stock-based compensation)")
SA.gap()
SA.section("MÉMO IFRS 16 (dettes locatives déjà incluses dans les dettes financières ci-dessus)", "E")
SA.header(["N-2", "N-1", "N"])
inp("lease_lt", "Dont dettes locatives IFRS 16 à plus d'un an", "Inclure ces dettes dans « Emprunts et dettes financières > 1 an » ; 0 si pas d'IFRS 16")
inp("lease_ct", "Dont dettes locatives IFRS 16 à moins d'un an", "Inclure dans « Dettes financières < 1 an »")
inp("loyers", "Loyers IFRS 16 décaissés (remboursement + intérêts des dettes locatives)", "Tableau de flux : remboursement des dettes locatives + intérêts versés sur dettes locatives")
ws_sa.freeze_panes = "B4"
_dv = DataValidation(type="list", formula1='"Générique,Services,Industrie"', allow_blank=False)
ws_sa.add_data_validation(_dv)
_dv.add(f"B{SA.r['profil']}")

S = lambda k, c: f"Saisie!{c}{SA.r[k]}"
SABS = lambda k: f"Saisie!$B${SA.r[k]}"

# =====================================================================
# HYPOTHESES (lignes fixes)
# =====================================================================
HY = Sh(ws_hy, {"A": 56, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 18, "H": 62})
HY.title("HYPOTHÈSES — Coût du capital, valeur terminale, scénarios DCF",
         "Cellules jaunes/bleues = hypothèses à ajuster. Colonne « Repère » = niveau historique pour calibrer. "
         "Les valeurs fournies sont des hypothèses illustratives, pas des prévisions.")
H = dict(rf=5, erp=6, beta=7, size=8, ke=9, kd=10, tax=11, kd_at=12, wd=13, wacc=14,
         g=17, mult=18, wg=19, mid=20, minor=21, nonop=22, ifrs=23, lease=24, stub=25,
         hy_hdr=28, da=29, capex=30, bfr=31,
         sc_hdr=35, bear_g=36, bear_m=37, base_g=38, base_m=39, bull_g=40, bull_m=41,
         p_hdr=44, p=45, p_chk=46)


def hin(row, label, val, fmt, note="", col="B"):
    HY.ws.cell(row, 1, label).font = font()
    c = HY.ws[f"{col}{row}"]
    c.value = val
    c.font = font(color=BLUE)
    c.fill = fill_inp
    c.border = box
    c.number_format = fmt
    if note:
        HY.ws.cell(row, 8, note).font = font(italic=True, color="595959")


def hcalc(row, label, formula, fmt, bold=False):
    HY.ws.cell(row, 1, label).font = font(bold)
    c = HY.ws[f"B{row}"]
    c.value = formula
    c.font = font(bold)
    c.number_format = fmt
    c.fill = fill_tot


HY.cur = H["rf"] - 1
HY.section("1. COÛT DU CAPITAL (WACC)", "H")
hin(H["rf"], "Taux sans risque (OAT/Bund 10 ans)", HYD["rf"], PCT, "À mettre à jour : taux souverain à 10 ans de la devise du modèle")
hin(H["erp"], "Prime de risque du marché actions", HYD["erp"], PCT, "Repère usuel : 5 % à 6,5 % (Damodaran, Kroll, Fairness finance)")
hin(H["beta"], "Bêta endetté retenu → saisi (col. B) | mode (col. C) | comparables (col. D)", HYD["beta"], D2,
    "Mode : 1 = bêta saisi ; 2 = médiane des comparables ré-endettée (onglet WACC_Pairs) ; 3 = idem ajusté Blume")
hin(H["beta"], "", HYD.get("beta_mode", 1), INT0, col="C")
HY.ws[f"D{H['beta']}"].value = f"=WACC_Pairs!$B$20"
HY.ws[f"D{H['beta']}"].number_format = D2
HY.ws[f"D{H['beta']}"].font = font(color=GREEN)
HY.ws[f"E{H['beta']}"].value = f"=WACC_Pairs!$B$21"
HY.ws[f"E{H['beta']}"].number_format = D2
HY.ws[f"E{H['beta']}"].font = font(color=GREEN)
HY.ws.cell(H["beta"], 1).value = "Bêta : saisi (B) | mode 1-2-3 (C) | comparables (D) | ajusté Blume (E)"
hin(H["size"], "Prime de taille / de risque spécifique", HYD["size"], PCT, "Petites capitalisations ou sociétés non cotées : +1 % à +4 %")
hcalc(H["ke"], "Coût des capitaux propres (CAPM) = rf + β retenu × prime + prime spécifique",
      (f"=B{H['rf']}+IF(ISNUMBER(CHOOSE(C{H['beta']},B{H['beta']},D{H['beta']},E{H['beta']})),"
       f"CHOOSE(C{H['beta']},B{H['beta']},D{H['beta']},E{H['beta']}),B{H['beta']})*B{H['erp']}+B{H['size']}"), PCT)
hin(H["kd"], "Coût de la dette avant impôt", HYD["kd"], PCT, "Taux moyen des financements ou taux de marché pour ce risque de crédit")
hin(H["tax"], "Taux d'impôt normatif", HYD["tax"], PCT, "France : 25 % (taux normal IS). Monaco : 25 %. Allemagne : ≈ 30 %")
hcalc(H["kd_at"], "Coût de la dette après impôt", f"=B{H['kd']}*(1-B{H['tax']})", PCT)
hin(H["wd"], "Poids cible de la dette D/(D+E)", HYD["wd"], PCT, "Structure cible ou moyenne sectorielle (valeur de marché)")
hcalc(H["wacc"], "WACC", f"=(1-B{H['wd']})*B{H['ke']}+B{H['wd']}*B{H['kd_at']}", '0.00%', bold=True)
HY.cur = H["g"] - 1
HY.section("2. VALEUR TERMINALE ET PASSAGE VALEUR D'ENTREPRISE → CAPITAUX PROPRES", "H")
hin(H["g"], "Taux de croissance perpétuelle (g)", HYD["g"], PCT, "Doit rester < croissance nominale long terme du PIB (≈ 1,5 % – 2,5 %)")
hin(H["mult"], "Multiple de sortie VE / EBITDA", HYD["mult"], MULT, "Médiane de comparables ou transactions du secteur (voir onglet Comparables)")
hin(H["wg"], "Pondération Gordon-Shapiro dans la valeur retenue", HYD["wg"], PCT, "Le solde (1 − poids) est affecté à la méthode du multiple de sortie")
hin(H["mid"], "Convention mid-year (1 = oui, 0 = non)", HYD["mid"], INT0, "1 : flux supposés répartis au milieu de l'année (pratique bancaire courante)")
hin(H["minor"], "Intérêts minoritaires, retraites et autres dettes assimilées (à déduire)", HYD["minor"], NUM, "Même unité que les états financiers (engagements de retraite, compléments de prix, minoritaires)")
hin(H["nonop"], "Actifs non opérationnels (à ajouter)", HYD["nonop"], NUM, "Titres de participation non consolidés, immeubles de placement…")
hin(H["ifrs"], "Traitement IFRS 16 : 1 = loyers traités en charge, 0 = dettes locatives dans la dette", HYD.get("ifrs", 0), INT0,
    "1 recommandé pour les activités peu capitalistiques : on déduit les loyers du FCFF et on exclut les dettes locatives de la dette nette")
hin(H["lease"], "Loyers IFRS 16 décaissés (% du CA), utilisés si traitement = 1", HYD.get("lease", 0.0), PCT, "Repère historique en colonne G")
hin(H["stub"], "Part de l'année N+1 déjà écoulée à la date du bilan retenu (0 = valorisation à la clôture N)", HYD.get("stub", 0.0), '0.00',
    "Ex. dernier bilan intermédiaire à 9 mois : 0,75. Saisir alors la dette nette et les actions à cette date (ajustement via les actifs non opérationnels) : la part écoulée du FCFF N+1 est retirée")
HY.cur = H["hy_hdr"] - 2
HY.section("3. HYPOTHÈSES OPÉRATIONNELLES COMMUNES AUX SCÉNARIOS (en % du chiffre d'affaires)", "H")
YEARS = ["N+1", "N+2", "N+3", "N+4", "N+5"]
HCOLS = ["B", "C", "D", "E", "F"]
for i, c in enumerate(HCOLS):
    cell = HY.ws[f"{c}{H['hy_hdr']}"]
    cell.value = YEARS[i]
    cell.font = font(True, "FFFFFF")
    cell.fill = fill_hdr
    cell.alignment = Alignment(horizontal="center")
HY.ws.cell(H["hy_hdr"], 1).fill = fill_hdr
g = HY.ws[f"G{H['hy_hdr']}"]
g.value = "Repère (N)"
g.font = font(True, "FFFFFF")
g.fill = fill_hdr
g.alignment = Alignment(horizontal="center")
hy_rows = [
    ("da", "Dotations aux amortissements (% CA)", HYD["da"]),
    ("capex", "Investissements — capex (% CA)", HYD["capex"]),
    ("bfr", "BFR total (% CA)", HYD["bfr"]),
]
for k, lab, vals in hy_rows:
    HY.ws.cell(H[k], 1, lab).font = font()
    for c, v in zip(HCOLS, vals):
        cell = HY.ws[f"{c}{H[k]}"]
        cell.value = v
        cell.font = font(color=BLUE)
        cell.fill = fill_inp
        cell.border = box
        cell.number_format = PCT
HY.ws.cell(H["capex"], 8, "Converger vers un niveau proche des dotations en année terminale").font = font(italic=True, color="595959")
HY.ws.cell(H["bfr"], 8, "BFR (exploitation + hors exploitation) rapporté au CA ; écart avec N = effet de ΔBFR en N+1").font = font(italic=True, color="595959")

HY.cur = H["sc_hdr"] - 2
HY.section("4. SCÉNARIOS : croissance du CA et marge d'EBITDA", "H")
for i, c in enumerate(HCOLS):
    cell = HY.ws[f"{c}{H['sc_hdr']}"]
    cell.value = YEARS[i]
    cell.font = font(True, "FFFFFF")
    cell.fill = fill_hdr
    cell.alignment = Alignment(horizontal="center")
HY.ws.cell(H["sc_hdr"], 1).fill = fill_hdr
g = HY.ws[f"G{H['sc_hdr']}"]
g.value = "Repère (N)"
g.font = font(True, "FFFFFF")
g.fill = fill_hdr
g.alignment = Alignment(horizontal="center")
_sc = HYD["scenarios"]
sc_rows = [
    ("bear_g", "Bear — croissance du CA", _sc["bear"]["g"], PCT),
    ("bear_m", "Bear — marge d'EBITDA", _sc["bear"]["m"], PCT),
    ("base_g", "Base — croissance du CA", _sc["base"]["g"], PCT),
    ("base_m", "Base — marge d'EBITDA", _sc["base"]["m"], PCT),
    ("bull_g", "Bull — croissance du CA", _sc["bull"]["g"], PCT),
    ("bull_m", "Bull — marge d'EBITDA", _sc["bull"]["m"], PCT),
]
for k, lab, vals, fm in sc_rows:
    HY.ws.cell(H[k], 1, lab).font = font(bold=k.endswith("_g"))
    for c, v in zip(HCOLS, vals):
        cell = HY.ws[f"{c}{H[k]}"]
        cell.value = v
        cell.font = font(color=BLUE)
        cell.fill = fill_inp
        cell.border = box
        cell.number_format = fm
HY.ws.cell(H["base_m"], 8, "Marge d'EBITDA au sens des états (post-IFRS 16 : avant loyers si traitement = 1)").font = font(italic=True, color="595959")
HY.cur = H["p_hdr"] - 1
HY.section("5. PROBABILITÉS DES SCÉNARIOS (pour la valeur pondérée)", "H")
for c, lab in zip(["B", "C", "D"], ["Bear", "Base", "Bull"]):
    cell = HY.ws[f"{c}{H['p_hdr']}"]
    cell.value = lab
    cell.font = font(True, "FFFFFF")
    cell.fill = fill_hdr
    cell.alignment = Alignment(horizontal="center")
HY.ws.cell(H["p_hdr"], 1).fill = fill_hdr
HY.ws.cell(H["p"], 1, "Probabilité").font = font()
for c, v in zip(["B", "C", "D"], HYD["prob"]):
    cell = HY.ws[f"{c}{H['p']}"]
    cell.value = v
    cell.font = font(color=BLUE)
    cell.fill = fill_inp
    cell.border = box
    cell.number_format = PCT
HY.ws.cell(H["p_chk"], 1, "Total (doit faire 100 %)").font = font()
c = HY.ws[f"B{H['p_chk']}"]
c.value = f"=SUM(B{H['p']}:D{H['p']})"
c.number_format = PCT
c.font = font(True)
HYR = lambda k, c="B", ab=True: f"Hypotheses!${c}${H[k]}" if ab else f"Hypotheses!{c}{H[k]}"
ws_hy.freeze_panes = "B4"

# =====================================================================
# SIG_CAF
# =====================================================================
SG = Sh(ws_sig, {"A": 52, "B": 14, "C": 14, "D": 14, "E": 60})
SG.title("SIG, EBITDA, RÉSULTATS ET CAPACITÉ D'AUTOFINANCEMENT",
         "Soldes intermédiaires de gestion calculés à partir de l'onglet Saisie.")
SG.header([f"=Saisie!{c}{SA.r['annee']}" for c in COLS])
sg = SG.ref
L = SG.line
L("ca", "Chiffre d'affaires", lambda c, p: f"={S('ca', c)}", link=True, bold=True)
L("achats", "− Achats consommés", lambda c, p: f"={S('achats', c)}", link=True, indent=1)
L("mb", "Marge brute", lambda c, p: f"={c}{SG.r['ca']}-{c}{SG.r['achats']}", bold=True, total=True)
L("tx_mb", "Taux de marge brute", lambda c, p: f'=IFERROR({c}{SG.r["mb"]}/{c}{SG.r["ca"]},"n/d")', PCT, indent=1)
L("ext", "− Autres achats et charges externes", lambda c, p: f"={S('ext', c)}", link=True, indent=1)
L("va", "Valeur ajoutée", lambda c, p: f"={c}{SG.r['mb']}-{c}{SG.r['ext']}", bold=True, total=True)
L("taxes", "− Impôts et taxes", lambda c, p: f"={S('taxes', c)}", link=True, indent=1)
L("perso", "− Charges de personnel", lambda c, p: f"={S('perso', c)}", link=True, indent=1)
L("aprod", "+ Autres produits d'exploitation", lambda c, p: f"={S('aprod', c)}", link=True, indent=1)
L("acharg", "− Autres charges d'exploitation", lambda c, p: f"={S('acharg', c)}", link=True, indent=1)
L("ebe", "EBE / EBITDA", lambda c, p: (f"={c}{SG.r['va']}-{c}{SG.r['taxes']}-{c}{SG.r['perso']}"
                                       f"+{c}{SG.r['aprod']}-{c}{SG.r['acharg']}"), bold=True, total=True)
L("m_ebe", "Marge d'EBITDA", lambda c, p: f'=IFERROR({c}{SG.r["ebe"]}/{c}{SG.r["ca"]},"n/d")', PCT, indent=1)
L("sbc", "Mémo : rémunération en actions (non décaissée)", lambda c, p: f"={S('sbc', c)}", link=True, indent=1)
L("ebe_aj", "EBITDA ajusté de la rémunération en actions (EBITDA + SBC)", lambda c, p: f"={c}{SG.r['ebe']}+{c}{SG.r['sbc']}", indent=1)
L("loyers", "Mémo : loyers IFRS 16 décaissés", lambda c, p: f"={S('loyers', c)}", link=True, indent=1)
L("ebe_apres_loyers", "EBITDA après loyers IFRS 16 (avant retraitement IFRS 16 : EBITDA − loyers)",
  lambda c, p: f"={c}{SG.r['ebe']}-{c}{SG.r['loyers']}", indent=1)
L("dot", "− Dotations aux amortissements et provisions", lambda c, p: f"={S('dot', c)}", link=True, indent=1)
L("ebit", "Résultat d'exploitation (EBIT)", lambda c, p: f"={c}{SG.r['ebe']}-{c}{SG.r['dot']}", bold=True, total=True)
L("m_ebit", "Marge d'EBIT", lambda c, p: f'=IFERROR({c}{SG.r["ebit"]}/{c}{SG.r["ca"]},"n/d")', PCT, indent=1)
L("pfin", "+ Produits financiers", lambda c, p: f"={S('pfin', c)}", link=True, indent=1)
L("cfin", "− Charges financières", lambda c, p: f"={S('cfin', c)}", link=True, indent=1)
L("rfin", "Résultat financier", lambda c, p: f"={c}{SG.r['pfin']}-{c}{SG.r['cfin']}")
L("rcai", "Résultat courant avant impôt (RCAI)", lambda c, p: f"={c}{SG.r['ebit']}+{c}{SG.r['rfin']}", bold=True, total=True)
L("rexc", "Résultat exceptionnel", lambda c, p: f"={S('pexc', c)}-{S('cexc', c)}")
L("rav", "Résultat avant impôt", lambda c, p: f"={c}{SG.r['rcai']}+{c}{SG.r['rexc']}", bold=True)
L("impot", "− Impôt sur les bénéfices", lambda c, p: f"={S('impot', c)}", link=True, indent=1)
L("rn", "RÉSULTAT NET", lambda c, p: f"={c}{SG.r['rav']}-{c}{SG.r['impot']}", bold=True, total=True)
L("m_rn", "Marge nette", lambda c, p: f'=IFERROR({c}{SG.r["rn"]}/{c}{SG.r["ca"]},"n/d")', PCT, indent=1)
SG.gap()
SG.section("CAPACITÉ D'AUTOFINANCEMENT ET RÉSULTAT NORMATIF", "E")
L("caf", "CAF (résultat net + dotations)", lambda c, p: f"={c}{SG.r['rn']}+{c}{SG.r['dot']}", bold=True)
L("taux_is", "Taux d'impôt effectif (IS / résultat avant impôt)",
  lambda c, p: f'=IFERROR({c}{SG.r["impot"]}/{c}{SG.r["rav"]},"n/d")', PCT)
L("tax_norm", "Taux d'impôt normatif (Hypotheses)", lambda c, p: f"={HYR('tax')}", PCT, link=True)
L("nopat", "NOPAT = EBIT − impôt normatif sur EBIT",
  lambda c, p: f"={c}{SG.r['ebit']}-MAX({c}{SG.r['ebit']},0)*{c}{SG.r['tax_norm']}", bold=True)
L("ctrl_rn", "Contrôle : RN calculé − résultat du bilan (doit être 0)",
  lambda c, p: f"={c}{SG.r['rn']}-{S('res_bilan', c)}")
ws_sig.freeze_panes = "B5"
G = lambda k, c: f"SIG_CAF!{c}{SG.r[k]}"

# =====================================================================
# BILAN_BFR
# =====================================================================
BF = Sh(ws_bf, {"A": 56, "B": 14, "C": 14, "D": 14, "E": 50})
BF.title("BILAN FONCTIONNEL, FRNG, BFR ET ENDETTEMENT",
         "Structure du bilan retraitée à partir de l'onglet Saisie.")
BF.header([f"=Saisie!{c}{SA.r['annee']}" for c in COLS])
B_ = BF.line
BF.section("ACTIF", "E")
B_("immo", "Actif immobilisé net", lambda c, p: f"={S('immo_inc', c)}+{S('immo_corp', c)}+{S('immo_fin', c)}", bold=True)
B_("stocks", "Stocks", lambda c, p: f"={S('stocks', c)}", link=True, indent=1)
B_("clients", "Clients", lambda c, p: f"={S('clients', c)}", link=True, indent=1)
B_("autres_cr", "Autres créances", lambda c, p: f"={S('autres_cr', c)}", link=True, indent=1)
B_("treso", "Trésorerie", lambda c, p: f"={S('treso', c)}", link=True, indent=1)
B_("circ", "Actif circulant (hors immobilisations)", lambda c, p: f"=SUM({c}{BF.r['stocks']}:{c}{BF.r['treso']})", bold=True)
B_("ta", "Total actif", lambda c, p: f"={c}{BF.r['immo']}+{c}{BF.r['circ']}", bold=True, total=True)
BF.gap()
BF.section("PASSIF", "E")
B_("cp", "Capitaux propres", lambda c, p: f"={S('capital', c)}+{S('reserves', c)}+{S('res_bilan', c)}", bold=True)
B_("prov", "Provisions pour risques et charges", lambda c, p: f"={S('provisions', c)}", link=True, indent=1)
B_("dlt", "Dettes financières > 1 an", lambda c, p: f"={S('dlt', c)}", link=True, indent=1)
B_("dct", "Dettes financières < 1 an", lambda c, p: f"={S('dct', c)}", link=True, indent=1)
B_("fourn", "Fournisseurs", lambda c, p: f"={S('fourn', c)}", link=True, indent=1)
B_("fisc", "Dettes fiscales et sociales", lambda c, p: f"={S('fisc', c)}", link=True, indent=1)
B_("autres_det", "Autres dettes", lambda c, p: f"={S('autres_det', c)}", link=True, indent=1)
B_("dettes_ct", "Dettes à court terme (financières < 1 an + exploitation + autres)",
   lambda c, p: f"=SUM({c}{BF.r['dct']}:{c}{BF.r['autres_det']})", bold=True)
B_("tp", "Total passif", lambda c, p: (f"={c}{BF.r['cp']}+{c}{BF.r['prov']}+{c}{BF.r['dlt']}"
                                       f"+{c}{BF.r['dettes_ct']}"), bold=True, total=True)
BF.gap()
BF.section("ÉQUILIBRE FINANCIER", "E")
B_("kperm", "Capitaux permanents (CP + provisions + dettes financières > 1 an)",
   lambda c, p: f"={c}{BF.r['cp']}+{c}{BF.r['prov']}+{c}{BF.r['dlt']}")
B_("frng", "Fonds de roulement net global (FRNG) = capitaux permanents − actif immobilisé",
   lambda c, p: f"={c}{BF.r['kperm']}-{c}{BF.r['immo']}", bold=True, total=True)
B_("bfr_ex", "BFR d'exploitation (stocks + clients − fournisseurs − dettes fisc. et soc.)",
   lambda c, p: f"={c}{BF.r['stocks']}+{c}{BF.r['clients']}-{c}{BF.r['fourn']}-{c}{BF.r['fisc']}")
B_("bfr_hex", "BFR hors exploitation (autres créances − autres dettes)",
   lambda c, p: f"={c}{BF.r['autres_cr']}-{c}{BF.r['autres_det']}")
B_("bfr", "BFR total", lambda c, p: f"={c}{BF.r['bfr_ex']}+{c}{BF.r['bfr_hex']}", bold=True, total=True)
B_("tn", "Trésorerie nette (FRNG − BFR)", lambda c, p: f"={c}{BF.r['frng']}-{c}{BF.r['bfr']}", bold=True)
B_("tn_ck", "Contrôle : trésorerie − dettes financières < 1 an − trésorerie nette (doit être 0)",
   lambda c, p: f"={c}{BF.r['treso']}-{c}{BF.r['dct']}-{c}{BF.r['tn']}")
B_("bfr_j", "BFR en jours de chiffre d'affaires",
   lambda c, p: f'=IFERROR({c}{BF.r["bfr"]}/{G("ca", c)}*365,"n/d")', D1)
BF.gap()
BF.section("ENDETTEMENT ET CAPITAL EMPLOYÉ", "E")
B_("dfin", "Dettes financières brutes", lambda c, p: f"={c}{BF.r['dlt']}+{c}{BF.r['dct']}")
B_("dn", "Dette financière nette (dettes brutes − trésorerie)", lambda c, p: f"={c}{BF.r['dfin']}-{c}{BF.r['treso']}", bold=True, total=True)
B_("leases", "Dont dettes locatives IFRS 16 (mémo)", lambda c, p: f"={S('lease_lt', c)}+{S('lease_ct', c)}", link=True, indent=1)
B_("dn16", "Dette financière nette hors dettes locatives (définition « covenant »)", lambda c, p: f"={c}{BF.r['dn']}-{c}{BF.r['leases']}", bold=True)
B_("ce", "Capital employé économique (capitaux propres + dette nette)", lambda c, p: f"={c}{BF.r['cp']}+{c}{BF.r['dn']}")
B_("ctrl_ta", "Contrôle : total actif − total passif (doit être 0)", lambda c, p: f"={c}{BF.r['ta']}-{c}{BF.r['tp']}")
ws_bf.freeze_panes = "B5"
BR = lambda k, c: f"Bilan_BFR!{c}{BF.r[k]}"

# =====================================================================
# CASHFLOW
# =====================================================================
CF = Sh(ws_cf, {"A": 60, "B": 14, "C": 14, "D": 14, "E": 60})
CF.title("FLUX DE TRÉSORERIE ET FREE CASH-FLOW",
         "Tableau de flux reconstitué à partir des bilans successifs (nécessite au moins deux exercices : N-2 sert de base).")
CF.header([f"=Saisie!{c}{SA.r['annee']}" for c in COLS])
C_ = CF.line
ND = "n/d"
C_("caf", "Capacité d'autofinancement (CAF)", lambda c, p: f"={G('caf', c)}", link=True, bold=True)
C_("d_bfr", "− Variation du BFR (impact trésorerie)", lambda c, p: f"=-({BR('bfr', c)}-{BR('bfr', p)})", first=ND, indent=1)
C_("fte", "Flux de trésorerie d'exploitation (FTE)", lambda c, p: f"={c}{CF.r['caf']}+{c}{CF.r['d_bfr']}", first=ND, bold=True, total=True)
C_("capex", "− Investissements (capex)",
   lambda c, p: (f"=IF(ISNUMBER({S('capex', c)}),{S('capex', c)},"
                 f"({BR('immo', c)}-{BR('immo', p)})+{G('dot', c)})"),
   first=f'=IF(ISNUMBER({S("capex", "B")}),{S("capex", "B")},"n/d")', indent=1)
C_("fcf_op", "Free cash-flow opérationnel (FTE − capex)",
   lambda c, p: f'=IFERROR({c}{CF.r["fte"]}-{c}{CF.r["capex"]},"n/d")', first=ND, bold=True, total=True)
CF.gap()
CF.section("FREE CASH-FLOW TO FIRM / TO EQUITY", "E")
C_("fcff", "FCFF = NOPAT + dotations − capex − ΔBFR",
   lambda c, p: f'=IFERROR({G("nopat", c)}+{G("dot", c)}-{c}{CF.r["capex"]}+{c}{CF.r["d_bfr"]},"n/d")', first=ND, bold=True)
C_("fin_net", "Résultat financier après impôt", lambda c, p: f"={G('rfin', c)}*(1-{G('tax_norm', c)})", indent=1)
C_("d_dette", "Variation des dettes financières brutes", lambda c, p: f"={BR('dfin', c)}-{BR('dfin', p)}", first=ND, indent=1)
C_("fcfe", "FCFE = FCFF + résultat financier après impôt + Δ dettes financières",
   lambda c, p: f'=IFERROR({c}{CF.r["fcff"]}+{c}{CF.r["fin_net"]}+{c}{CF.r["d_dette"]},"n/d")', first=ND, bold=True)
CF.gap()
CF.section("RAPPROCHEMENT AVEC LA VARIATION DE TRÉSORERIE", "E")
C_("div", "− Dividendes versés", lambda c, p: f"={S('div', c)}", link=True, indent=1)
C_("acq", "− Acquisitions de filiales", lambda c, p: f"={S('acq', c)}", link=True, indent=1)
C_("rach", "− Rachats d'actions et opérations sur capital", lambda c, p: f"={S('rach', c)}", link=True, indent=1)
C_("autres_inv", "− Autres investissements financiers nets", lambda c, p: f"={S('autres_inv', c)}", link=True, indent=1)
C_("autres_dec", "− Autres décaissements sur capitaux (impôts sur actions gratuites…)", lambda c, p: f"={S('autres_dec', c)}", link=True, indent=1)
C_("sbc_cf", "+ Rémunération en actions (charge non décaissée)", lambda c, p: f"={S('sbc', c)}", link=True, indent=1)
C_("noncash_cf", "− Gains latents inclus dans le résultat", lambda c, p: f"={S('noncash', c)}", link=True, indent=1)
C_("dt_theo", "Variation de trésorerie théorique (FCF op. + Δ dettes − dividendes − acquisitions − rachats − autres + SBC − gains latents)",
   lambda c, p: (f'=IFERROR({c}{CF.r["fcf_op"]}+{c}{CF.r["d_dette"]}-{c}{CF.r["div"]}'
                 f'-{c}{CF.r["acq"]}-{c}{CF.r["rach"]}-{c}{CF.r["autres_inv"]}-{c}{CF.r["autres_dec"]}'
                 f'+{c}{CF.r["sbc_cf"]}-{c}{CF.r["noncash_cf"]},"n/d")'), first=ND)
C_("dt_reel", "Variation de trésorerie constatée au bilan", lambda c, p: f"={BR('treso', c)}-{BR('treso', p)}", first=ND)
C_("dt_ecart", "Écart (capital, provisions, cessions, éléments hors exploitation)",
   lambda c, p: f'=IFERROR({c}{CF.r["dt_reel"]}-{c}{CF.r["dt_theo"]},"n/d")', first=ND)
CF.gap()
CF.section("QUALITÉ DES FLUX", "E")
C_("conv", "Conversion cash (FCF opérationnel / EBITDA)",
   lambda c, p: f'=IFERROR(IF({G("ebe", c)}<=0,"n/s",{c}{CF.r["fcf_op"]}/{G("ebe", c)}),"n/d")', PCT, first=ND)
C_("m_fcff", "Marge de FCFF (FCFF / CA)",
   lambda c, p: f'=IFERROR({c}{CF.r["fcff"]}/{G("ca", c)},"n/d")', PCT, first=ND)
C_("capex_ca", "Capex / chiffre d'affaires",
   lambda c, p: f'=IFERROR({c}{CF.r["capex"]}/{G("ca", c)},"n/d")', PCT)
C_("capex_dot", "Capex / dotations aux amortissements",
   lambda c, p: f'=IFERROR({c}{CF.r["capex"]}/{G("dot", c)},"n/d")', MULT)
ws_cf.freeze_panes = "B5"
CFR = lambda k, c: f"Cashflow!{c}{CF.r[k]}"

# =====================================================================
# RATIOS
# =====================================================================
RA = Sh(ws_ra, {"A": 46, "B": 11, "C": 11, "D": 11, "E": 8, "F": 15, "G": 52, "H": 26, "I": 10, "J": 10})
RA.title("RATIOS FINANCIERS — rentabilité, croissance, liquidité, solvabilité, activité, cash et marché",
         "Lecture automatique de l'exercice N selon le profil de seuils choisi dans Saisie (table modifiable : onglet Seuils). "
         "Les seuils dépendent du secteur : ils servent de repère, pas de verdict.")
RA.header(["N-2", "N-1", "N", "Tend.", "Lecture N", "Formule", "Repère", "Seuil favorable", "Seuil vigilance"])
for i, c in enumerate(COLS):
    RA.ws[f"{c}{RA.cur - 1}"].value = f"=Saisie!{c}{SA.r['annee']}"

# Seuils de lecture par profil : (favorable, vigilance) pour Services et Industrie ; Générique = valeurs passées en argument.
THR = {
    "m_ebitda": ((0.12, 0.08), (0.12, 0.07)), "m_ebit": ((0.10, 0.05), (0.08, 0.04)), "m_net": ((0.07, 0.03), (0.05, 0.02)),
    "roe": ((0.15, 0.09), (0.10, 0.06)), "roa": ((0.07, 0.03), (0.04, 0.02)), "roce": ((0.15, 0.10), (0.10, 0.07)),
    "roic": ((0.12, 0.08), (0.09, 0.06)), "g_ca": ((0.06, 0.0), (0.03, 0.0)), "g_ebitda": ((0.06, 0.0), (0.03, 0.0)),
    "g_rn": ((0.06, 0.0), (0.03, 0.0)), "liq_g": ((1.2, 1.0), (1.5, 1.2)), "liq_r": ((1.0, 0.8), (0.9, 0.7)),
    "liq_i": ((0.3, 0.15), (0.1, 0.05)), "autonomie": ((0.35, 0.20), (0.35, 0.25)), "gearing": ((0.5, 1.0), (0.8, 1.2)),
    "nd_ebitda": ((1.5, 3.0), (2.5, 4.0)), "nd_ebitda16": ((1.5, 3.0), (2.5, 4.0)), "couv_int": ((6.0, 3.0), (4.0, 2.5)),
    "rembt": ((2.5, 4.5), (4.0, 6.0)), "couv_emplois": ((1.0, 0.9), (1.0, 0.9)), "dso": ((55, 75), (60, 80)),
    "conv": ((0.6, 0.3), (0.4, 0.2)), "m_fcff": ((0.07, 0.02), (0.04, 0.0)),
}
SE = Sh(ws_se, {"A": 54, "B": 9, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 3, "J": 14})
ws_se["A1"].value = "SEUILS DE LECTURE DES RATIOS PAR PROFIL"
ws_se["A1"].font = font(True, "1F3864", 14)
ws_se["A2"].value = "Profil sélectionné (index) :"
ws_se["A2"].font = font(True)
ws_se["J3"].value = "Profils"
ws_se["J3"].font = font(True)
for i_, nm_ in enumerate(["Générique", "Services", "Industrie"]):
    ws_se[f"J{4 + i_}"].value = nm_
ws_se["B2"].value = f"=IFERROR(MATCH(Saisie!B{SA.r['profil']},$J$4:$J$6,0),1)"
ws_se["B2"].font = font(True)
ws_se["A3"].value = "Seuils indicatifs (jugement d'analyste, modifiables). À remplacer par des références sectorielles documentées."
ws_se["A3"].font = font(italic=True, color="595959")
for j_, t_ in enumerate(["Ratio", "Sens", "Générique favorable", "Générique vigilance", "Services favorable", "Services vigilance",
                         "Industrie favorable", "Industrie vigilance"]):
    c_ = ws_se.cell(4, 1 + j_, t_)
    c_.font = font(True, "FFFFFF")
    c_.fill = fill_hdr
    c_.alignment = Alignment(horizontal="center", wrap_text=True)
SE.cur = 5


def ratio(key, label, expr, fmt, direction=None, good=None, warn=None, formula="", repere="", first=None, guard=None):
    """expr(c,p) -> expression string. direction 'hi' / 'lo' / None."""
    r = RA.cur
    RA.r[key] = r
    RA.cur += 1
    RA.ws.cell(r, 1, label).font = font()
    for i, c in enumerate(COLS):
        p = COLS[i - 1] if i else None
        cell = RA.ws[f"{c}{r}"]
        if i == 0 and first is not None:
            cell.value = first
        else:
            cell.value = f'=IFERROR({expr(c, p)},"n/d")'
        cell.number_format = fmt
        cell.font = font()
        cell.alignment = Alignment(horizontal="right")
    t = RA.ws[f"E{r}"]
    t.value = f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),IF(D{r}>C{r}+0.000001,"▲",IF(D{r}<C{r}-0.000001,"▼","►")),"–")'
    t.alignment = Alignment(horizontal="center")
    t.font = font()
    lec = RA.ws[f"F{r}"]
    if direction and good is not None:
        rs = SE.cur
        SE.cur += 1
        serv, ind = THR.get(key, ((good, warn), (good, warn)))
        ws_se.cell(rs, 1, label).font = font()
        ws_se.cell(rs, 2, "≥" if direction == "hi" else "≤").alignment = Alignment(horizontal="center")
        for j_, v_ in enumerate([good, warn, serv[0], serv[1], ind[0], ind[1]]):
            c_ = ws_se.cell(rs, 3 + j_, v_)
            c_.font = font(color=BLUE)
            c_.fill = fill_inp
            c_.number_format = fmt
        RA.ws[f"I{r}"].value = f"=INDEX(Seuils!$C{rs}:$H{rs},2*Seuils!$B$2-1)"
        RA.ws[f"J{r}"].value = f"=INDEX(Seuils!$C{rs}:$H{rs},2*Seuils!$B$2)"
        for cc in ("I", "J"):
            RA.ws[f"{cc}{r}"].font = font(color=GREEN)
            RA.ws[f"{cc}{r}"].number_format = fmt
        if direction == "hi":
            lec.value = f'=IF(ISNUMBER(D{r}),IF(D{r}>=I{r},"Favorable",IF(D{r}>=J{r},"À surveiller","Défavorable")),"n/d")'
        else:
            lec.value = f'=IF(ISNUMBER(D{r}),IF(D{r}<=I{r},"Favorable",IF(D{r}<=J{r},"À surveiller","Défavorable")),"n/d")'
    else:
        lec.value = "—"
    lec.font = font(True)
    lec.alignment = Alignment(horizontal="center")
    RA.ws[f"G{r}"].value = formula
    RA.ws[f"G{r}"].font = font(color="595959", size=9)
    RA.ws[f"H{r}"].value = repere
    RA.ws[f"H{r}"].font = font(color="595959", size=9)


def pos(x):
    return x  # lisibilité


def sec(t):
    RA.section(t, "J")


bfr_g = lambda c: BR("bfr", c)
sec("RENTABILITÉ ET MARGES")
ratio("m_brute", "Taux de marge brute", lambda c, p: G("tx_mb", c), PCT, formula="Marge brute / CA", repere="Selon secteur")
ratio("m_ebitda", "Marge d'EBITDA", lambda c, p: G("m_ebe", c), PCT, "hi", 0.12, 0.06, "EBE / CA", "> 12 % (services)")
ratio("m_ebit", "Marge d'EBIT", lambda c, p: G("m_ebit", c), PCT, "hi", 0.08, 0.03, "EBIT / CA", "> 8 %")
ratio("m_net", "Marge nette", lambda c, p: G("m_rn", c), PCT, "hi", 0.05, 0.02, "Résultat net / CA", "> 5 %")
ratio("roe", "ROE — rentabilité des capitaux propres", lambda c, p: f"{G('rn', c)}/{BR('cp', c)}", PCT, "hi", 0.12, 0.07, "Résultat net / capitaux propres", "> coût des fonds propres")
ratio("roa", "ROA — rentabilité de l'actif", lambda c, p: f"{G('rn', c)}/{BR('ta', c)}", PCT, "hi", 0.05, 0.02, "Résultat net / total actif", "> 5 %")
ratio("roce", "ROCE — rentabilité des capitaux permanents", lambda c, p: f"{G('ebit', c)}/{BR('kperm', c)}", PCT, "hi", 0.12, 0.08, "EBIT / capitaux permanents", "> WACC")
ratio("roic", "ROIC — rentabilité du capital investi après impôt", lambda c, p: f"{G('nopat', c)}/{BR('ce', c)}", PCT, "hi", 0.10, 0.07, "NOPAT / (CP + dette nette)", "> WACC")
sec("CROISSANCE")
ratio("g_ca", "Croissance du chiffre d'affaires", lambda c, p: f"{G('ca', c)}/{G('ca', p)}-1", PCT, "hi", 0.05, 0.0, "CA N / CA N-1 − 1", "> inflation", first="n/d")
ratio("g_ebitda", "Croissance de l'EBITDA", lambda c, p: f"{G('ebe', c)}/{G('ebe', p)}-1", PCT, "hi", 0.05, 0.0, "EBITDA N / EBITDA N-1 − 1", "≥ croissance du CA", first="n/d")
ratio("g_rn", "Croissance du résultat net", lambda c, p: f"{G('rn', c)}/{G('rn', p)}-1", PCT, "hi", 0.05, 0.0, "RN N / RN N-1 − 1", "", first="n/d")
sec("LIQUIDITÉ")
ratio("liq_g", "Liquidité générale", lambda c, p: f"{BR('circ', c)}/{BR('dettes_ct', c)}", D2, "hi", 1.2, 1.0, "Actif circulant / dettes CT", "> 1,2")
ratio("liq_r", "Liquidité réduite", lambda c, p: f"({BR('clients', c)}+{BR('autres_cr', c)}+{BR('treso', c)})/{BR('dettes_ct', c)}", D2, "hi", 1.0, 0.8, "(Créances + trésorerie) / dettes CT", "> 1")
ratio("liq_i", "Liquidité immédiate", lambda c, p: f"{BR('treso', c)}/{BR('dettes_ct', c)}", D2, "hi", 0.2, 0.1, "Trésorerie / dettes CT", "> 0,2")
sec("STRUCTURE FINANCIÈRE ET SOLVABILITÉ")
ratio("autonomie", "Autonomie financière", lambda c, p: f"{BR('cp', c)}/{BR('tp', c)}", PCT, "hi", 0.30, 0.20, "Capitaux propres / total passif", "> 30 %")
ratio("gearing", "Gearing (dette nette / capitaux propres)", lambda c, p: f"{BR('dn', c)}/{BR('cp', c)}", D2, "lo", 0.5, 1.0, "Dette nette / capitaux propres", "< 0,5")
ratio("nd_ebitda", "Dette nette / EBITDA", lambda c, p: f'IF({G("ebe", c)}<=0,"n/s",{BR("dn", c)}/{G("ebe", c)})', MULT, "lo", 2.0, 3.5, "Dette nette / EBITDA", "< 2,0x (3,5x bancaire)")
ratio("sbc_ca", "Rémunération en actions / chiffre d'affaires", lambda c, p: f"{G('sbc', c)}/{G('ca', c)}", PCT, formula="SBC / CA", repere="Charge réelle pour l'actionnaire (dilution) même si non décaissée")
ratio("nd_ebitda16", "Dette nette hors IFRS 16 / EBITDA après loyers", lambda c, p: f'IF({G("ebe_apres_loyers", c)}<=0,"n/s",{BR("dn16", c)}/{G("ebe_apres_loyers", c)})', MULT, "lo", 2.0, 3.5, "(Dette nette − dettes locatives) / (EBITDA − loyers)", "Définition covenant bancaire")
ratio("couv_int", "Couverture des intérêts", lambda c, p: f'IF({G("cfin", c)}<=0,"n/s",{G("ebit", c)}/{G("cfin", c)})', MULT, "hi", 5.0, 2.5, "EBIT / charges financières", "> 5x")
ratio("rembt", "Capacité de remboursement (années de CAF)", lambda c, p: f'IF({G("caf", c)}<=0,"n/s",{BR("dfin", c)}/{G("caf", c)})', D1, "lo", 3.0, 5.0, "Dettes financières brutes / CAF", "< 3 ans")
ratio("couv_emplois", "Couverture des emplois stables", lambda c, p: f"{BR('kperm', c)}/{BR('immo', c)}", D2, "hi", 1.0, 0.9, "Capitaux permanents / actif immobilisé", "> 1")
ratio("levier", "Levier financier (total actif / capitaux propres)", lambda c, p: f"{BR('ta', c)}/{BR('cp', c)}", D2, formula="Total actif / capitaux propres", repere="Composante de DuPont")
sec("ACTIVITÉ ET CYCLE D'EXPLOITATION")
ratio("dso", "Délai clients — DSO (jours)", lambda c, p: f"{BR('clients', c)}/({G('ca', c)}*(1+{SABS('tva')}))*365", D1, "lo", 45, 60, "Clients / (CA × (1 + TVA)) × 365", "Selon profil")
ratio("dpo", "Délai fournisseurs — DPO (jours)", lambda c, p: f"{BR('fourn', c)}/(({G('achats', c)}+{G('ext', c)})*(1+{SABS('tva')}))*365", D1, formula="Fournisseurs / ((achats + charges externes) × (1 + TVA)) × 365", repere="≤ 60 j (LME)")
ratio("dio", "Rotation des stocks — DIO (jours)", lambda c, p: f"{BR('stocks', c)}/{G('achats', c)}*365", D1, formula="Stocks / achats consommés × 365", repere="Selon secteur")
ratio("ccc", "Cycle de conversion de trésorerie (jours)", lambda c, p: f"{c}{RA.r['dso']}+{c}{RA.r['dio']}-{c}{RA.r['dpo']}", D1, formula="DSO + DIO − DPO", repere="Plus bas = mieux")
ratio("bfr_j", "BFR en jours de CA", lambda c, p: BR("bfr_j", c), D1, formula="BFR / CA × 365", repere="Selon secteur")
ratio("rot_actif", "Rotation de l'actif", lambda c, p: f"{G('ca', c)}/{BR('ta', c)}", D2, formula="CA / total actif", repere="Composante de DuPont")
ratio("rot_immo", "Rotation des immobilisations", lambda c, p: f"{G('ca', c)}/{BR('immo', c)}", D2, formula="CA / actif immobilisé", repere="Intensité capitalistique")
sec("CASH-FLOW")
ratio("conv", "Conversion cash (FCF opérationnel / EBITDA)", lambda c, p: CFR("conv", c), PCT, "hi", 0.5, 0.25, "FCF opérationnel / EBITDA", "> 50 %", first="n/d")
ratio("m_fcff", "Marge de FCFF", lambda c, p: CFR("m_fcff", c), PCT, "hi", 0.05, 0.0, "FCFF / CA", "> 5 %", first="n/d")
ratio("capex_ca", "Capex / chiffre d'affaires", lambda c, p: CFR("capex_ca", c), PCT, formula="Capex / CA", repere="Selon intensité capitalistique")
ratio("capex_dot", "Capex / dotations (taux de renouvellement)", lambda c, p: CFR("capex_dot", c), MULT, formula="Capex / dotations", repere="> 1x = expansion")
ratio("payout", "Taux de distribution (dividendes N / résultat N-1)", lambda c, p: f"{S('div', c)}/{G('rn', p)}", PCT, formula="Dividendes versés / résultat de l'exercice précédent", repere="Selon politique de distribution", first="n/d")
RA.gap()
RA.section("MULTIPLES DE MARCHÉ (si cours et nombre d'actions renseignés) — exercice N", "J")
mc = f'IF(AND(Saisie!$B${SA.r["actions"]}>0,Saisie!$B${SA.r["cours"]}>0),Saisie!$B${SA.r["actions"]}*Saisie!$B${SA.r["cours"]}/Saisie!$B${SA.r["mult"]},"n/d")'


def mk(key, label, expr, fmt, formula=""):
    r = RA.cur
    RA.r[key] = r
    RA.cur += 1
    RA.ws.cell(r, 1, label).font = font()
    for c in ("B", "C"):
        RA.ws[f"{c}{r}"].value = "–"
        RA.ws[f"{c}{r}"].alignment = Alignment(horizontal="right")
    cell = RA.ws[f"D{r}"]
    cell.value = f'=IFERROR({expr},"n/d")'
    cell.number_format = fmt
    cell.font = font(True)
    cell.alignment = Alignment(horizontal="right")
    RA.ws[f"G{r}"].value = formula
    RA.ws[f"G{r}"].font = font(color="595959", size=9)


mk("mcap", "Capitalisation boursière", mc, NUM, "Cours × nombre d'actions")
mk("ev_mkt", "Valeur d'entreprise de marché", f"D{RA.r['mcap']}+{BR('dn', 'D')}+{HYR('minor')}-{HYR('nonop')}", NUM, "Capitalisation + dette nette + minoritaires − actifs non opérationnels")
mk("per", "PER (cours / bénéfice)", f"D{RA.r['mcap']}/{G('rn', 'D')}", MULT, "Capitalisation / résultat net")
mk("ev_ebitda", "VE / EBITDA", f"D{RA.r['ev_mkt']}/{G('ebe', 'D')}", MULT, "")
mk("ev_ebit", "VE / EBIT", f"D{RA.r['ev_mkt']}/{G('ebit', 'D')}", MULT, "")
mk("ev_ca", "VE / chiffre d'affaires", f"D{RA.r['ev_mkt']}/{G('ca', 'D')}", D2, "")
mk("pb", "Price-to-book (capitalisation / capitaux propres)", f"D{RA.r['mcap']}/{BR('cp', 'D')}", D2, "")
mk("fcfe_y", "Rendement du FCFE", f"{CFR('fcfe', 'D')}/D{RA.r['mcap']}", PCT, "FCFE / capitalisation")
mk("div_y", "Rendement du dividende", f"{S('div', 'D')}/D{RA.r['mcap']}", PCT, "Dividendes versés N / capitalisation")
ws_ra.freeze_panes = "B5"
last_ra = RA.cur
for col, text, color in (("Favorable", "Favorable", "C6EFCE"), ("À surveiller", "À surveiller", "FFEB9C"), ("Défavorable", "Défavorable", "FFC7CE")):
    ws_ra.conditional_formatting.add(f"F5:F{last_ra}", FormulaRule(formula=[f'$F5="{text}"'], fill=PatternFill("solid", bgColor=color, fgColor=color)))
RAR = lambda k, c: f"Ratios!{c}{RA.r[k]}"

# =====================================================================
# SCORES
# =====================================================================
SC = Sh(ws_sc, {"A": 64, "B": 14, "C": 14, "D": 14, "E": 60})
SC.title("DUPONT, SCORE DE FAILLITE (ALTMAN) ET SCORE DE QUALITÉ (PIOTROSKI)",
         "Trois grilles de lecture utilisées par les investisseurs et les analystes crédit.")
SC.header([f"=Saisie!{c}{SA.r['annee']}" for c in COLS])
Q = SC.line
SC.section("DÉCOMPOSITION DUPONT À 5 FACTEURS DU ROE", "E")
Q("d_tax", "Charge fiscale (résultat net / résultat avant impôt)", lambda c, p: f'=IFERROR({G("rn", c)}/{G("rav", c)},"n/d")', D2)
Q("d_int", "Poids des intérêts et éléments hors exploitation (RAI / EBIT)", lambda c, p: f'=IFERROR({G("rav", c)}/{G("ebit", c)},"n/d")', D2)
Q("d_mebit", "Marge d'EBIT", lambda c, p: f'=IFERROR({G("m_ebit", c)},"n/d")', PCT)
Q("d_rot", "Rotation de l'actif (CA / total actif)", lambda c, p: f'=IFERROR({G("ca", c)}/{BR("ta", c)},"n/d")', D2)
Q("d_lev", "Levier financier (total actif / capitaux propres)", lambda c, p: f'=IFERROR({BR("ta", c)}/{BR("cp", c)},"n/d")', D2)
Q("d_roe", "ROE reconstitué = produit des 5 facteurs",
  lambda c, p: f'=IFERROR({c}{SC.r["d_tax"]}*{c}{SC.r["d_int"]}*{c}{SC.r["d_mebit"]}*{c}{SC.r["d_rot"]}*{c}{SC.r["d_lev"]},"n/d")', PCT, bold=True, total=True)
Q("d_ck", "Contrôle : ROE reconstitué − ROE direct (doit être 0)",
  lambda c, p: f'=IFERROR({c}{SC.r["d_roe"]}-{RAR("roe", c)},"n/d")', '0.0000')
SC.gap()
SC.section("SCORE Z'' D'ALTMAN (sociétés non industrielles / non cotées)", "E")
Q("x1", "X1 = fonds de roulement (actif circulant − dettes CT) / total actif", lambda c, p: f'=IFERROR(({BR("circ", c)}-{BR("dettes_ct", c)})/{BR("ta", c)},"n/d")', D2)
Q("x2", "X2 = résultats mis en réserve (réserves + résultat) / total actif", lambda c, p: f'=IFERROR(({S("reserves", c)}+{S("res_bilan", c)})/{BR("ta", c)},"n/d")', D2)
Q("x3", "X3 = EBIT / total actif", lambda c, p: f'=IFERROR({G("ebit", c)}/{BR("ta", c)},"n/d")', D2)
Q("x4", "X4 = capitaux propres comptables / total des dettes", lambda c, p: f'=IFERROR({BR("cp", c)}/({BR("tp", c)}-{BR("cp", c)}),"n/d")', D2)
Q("z", "Z'' = 6,56·X1 + 3,26·X2 + 6,72·X3 + 1,05·X4",
  lambda c, p: f'=IFERROR(6.56*{c}{SC.r["x1"]}+3.26*{c}{SC.r["x2"]}+6.72*{c}{SC.r["x3"]}+1.05*{c}{SC.r["x4"]},"n/d")', D2, bold=True, total=True)
Q("z_zone", "Zone", lambda c, p: f'=IF(ISNUMBER({c}{SC.r["z"]}),IF({c}{SC.r["z"]}>2.6,"Sûre",IF({c}{SC.r["z"]}>1.1,"Grise","Détresse")),"n/d")')
SC.ws.cell(SC.r["z_zone"], 5, "Seuils : > 2,6 zone sûre ; 1,1 à 2,6 zone grise ; < 1,1 risque de défaillance").font = font(italic=True, color="595959", size=9)
SC.gap()
SC.section("SCORE F DE PIOTROSKI (qualité financière, 0 à 9) — exercices N-1 et N", "E")
PC = ["C", "D"]
pio = [
    ("p1", "1. Rentabilité : ROA > 0", lambda c, p: f"=IF({G('rn', c)}/{BR('ta', c)}>0,1,0)"),
    ("p2", "2. Flux de trésorerie d'exploitation > 0", lambda c, p: f"=IF({CFR('fte', c)}>0,1,0)"),
    ("p3", "3. ROA en hausse", lambda c, p: f"=IF({G('rn', c)}/{BR('ta', c)}>{G('rn', p)}/{BR('ta', p)},1,0)"),
    ("p4", "4. Qualité des résultats : FTE > résultat net", lambda c, p: f"=IF({CFR('fte', c)}>{G('rn', c)},1,0)"),
    ("p5", "5. Levier en baisse (dettes LT / total actif)", lambda c, p: f"=IF({BR('dlt', c)}/{BR('ta', c)}<{BR('dlt', p)}/{BR('ta', p)},1,0)"),
    ("p6", "6. Liquidité générale en hausse", lambda c, p: f"=IF({BR('circ', c)}/{BR('dettes_ct', c)}>{BR('circ', p)}/{BR('dettes_ct', p)},1,0)"),
    ("p7", "7. Pas de dilution (capital et primes non en hausse)", lambda c, p: f"=IF({S('capital', c)}<={S('capital', p)},1,0)"),
    ("p8", "8. Marge brute en hausse", lambda c, p: f"=IF({G('tx_mb', c)}>{G('tx_mb', p)},1,0)"),
    ("p9", "9. Rotation de l'actif en hausse", lambda c, p: f"=IF({G('ca', c)}/{BR('ta', c)}>{G('ca', p)}/{BR('ta', p)},1,0)"),
]
for k, lab, f in pio:
    Q(k, lab, lambda c, p, f=f: f'=IFERROR({f(c, p)[1:]},"n/d")', INT0, first="n/d")
Q("pscore", "Score de Piotroski (total)", lambda c, p: f"=SUM({c}{SC.r['p1']}:{c}{SC.r['p9']})", INT0, first="n/d", bold=True, total=True)
Q("p_lect", "Lecture", lambda c, p: f'=IF({c}{SC.r["pscore"]}>=8,"Solide",IF({c}{SC.r["pscore"]}>=5,"Correct","Fragile"))', first="n/d")
SC.ws.cell(SC.r["p_lect"], 5, "8-9 solide ; 5-7 correct ; 0-4 fragile").font = font(italic=True, color="595959", size=9)
ws_sc.freeze_panes = "B5"
SCR = lambda k, c: f"Scores!{c}{SC.r[k]}"

# =====================================================================
# VARIATIONS
# =====================================================================
VA = Sh(ws_va, {"A": 48, "B": 13, "C": 13, "D": 13, "E": 13, "F": 10, "G": 13, "H": 10, "I": 20})
VA.title("ANALYSE DES VARIATIONS N / N-1 / N-2",
         "Variations en valeur et en %. Alerte si l'évolution dépasse ±20 %.")
VA.header(["N-2", "N-1", "N", "Δ N-1 / N-2", "Δ %", "Δ N / N-1", "Δ %", "Alerte"])
for c in COLS:
    VA.ws[f"{c}{VA.cur - 1}"].value = f"=Saisie!{c}{SA.r['annee']}"
va_items = [
    ("Compte de résultat", None, None),
    ("Chiffre d'affaires", "SIG_CAF", ("ca", SG)),
    ("Achats consommés", "SIG_CAF", ("achats", SG)),
    ("Marge brute", "SIG_CAF", ("mb", SG)),
    ("Charges externes", "SIG_CAF", ("ext", SG)),
    ("Charges de personnel", "SIG_CAF", ("perso", SG)),
    ("EBITDA", "SIG_CAF", ("ebe", SG)),
    ("Dotations", "SIG_CAF", ("dot", SG)),
    ("EBIT", "SIG_CAF", ("ebit", SG)),
    ("Résultat financier", "SIG_CAF", ("rfin", SG)),
    ("Résultat net", "SIG_CAF", ("rn", SG)),
    ("CAF", "SIG_CAF", ("caf", SG)),
    ("Bilan", None, None),
    ("Actif immobilisé", "Bilan_BFR", ("immo", BF)),
    ("Stocks", "Bilan_BFR", ("stocks", BF)),
    ("Clients", "Bilan_BFR", ("clients", BF)),
    ("Trésorerie", "Bilan_BFR", ("treso", BF)),
    ("Total actif", "Bilan_BFR", ("ta", BF)),
    ("Capitaux propres", "Bilan_BFR", ("cp", BF)),
    ("Dettes financières brutes", "Bilan_BFR", ("dfin", BF)),
    ("Dette financière nette", "Bilan_BFR", ("dn", BF)),
    ("Fournisseurs", "Bilan_BFR", ("fourn", BF)),
    ("BFR total", "Bilan_BFR", ("bfr", BF)),
    ("FRNG", "Bilan_BFR", ("frng", BF)),
    ("Flux", None, None),
    ("Capex", "Cashflow", ("capex", CF)),
    ("FCF opérationnel", "Cashflow", ("fcf_op", CF)),
    ("FCFF", "Cashflow", ("fcff", CF)),
]
va_rows = {}
for lab, shn, spec in va_items:
    r = VA.cur
    if spec is None:
        VA.section(lab, "I")
        continue
    k, shobj = spec
    VA.r[lab] = r
    va_rows[lab] = r
    VA.cur += 1
    VA.ws.cell(r, 1, lab).font = font()
    for c in COLS:
        cell = VA.ws[f"{c}{r}"]
        cell.value = f"={shn}!{c}{shobj.r[k]}"
        cell.number_format = NUM
        cell.font = font(color=GREEN)
        cell.alignment = Alignment(horizontal="right")
    for col, fx in (("E", f'=IFERROR(C{r}-B{r},"n/d")'), ("F", f'=IFERROR((C{r}-B{r})/ABS(B{r}),"n/d")'),
                    ("G", f'=IFERROR(D{r}-C{r},"n/d")'), ("H", f'=IFERROR((D{r}-C{r})/ABS(C{r}),"n/d")')):
        cell = VA.ws[f"{col}{r}"]
        cell.value = fx
        cell.number_format = PCT if col in "FH" else NUM
        cell.font = font()
        cell.alignment = Alignment(horizontal="right")
    VA.ws[f"I{r}"].value = f'=IF(ISNUMBER(H{r}),IF(ABS(H{r})>0.2,"Forte variation",""),"")'
    VA.ws[f"I{r}"].font = font(True)
ws_va.conditional_formatting.add(f"I5:I{VA.cur}", FormulaRule(formula=['$I5="Forte variation"'], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
ws_va.freeze_panes = "B5"

# =====================================================================
# DCF
# =====================================================================
DC = Sh(ws_dc, {"A": 58, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 4})
DC.title("VALORISATION DCF — 3 scénarios (Bear / Base / Bull) et tables de sensibilité",
         "Flux de trésorerie libres d'exploitation (FCFF) actualisés au WACC. Hypothèses : onglet Hypotheses.")
DCOLS = ["C", "D", "E", "F", "G"]
WACC = HYR("wacc")
GG = HYR("g")
MID = HYR("mid")
STUB = HYR("stub")
blocks = {}


def dcf_block(top, tag, name, grow_row, marg_row):
    ws = DC.ws
    R = {}
    names = ["title", "hdr", "g", "ca", "m", "ebitda", "da", "ebit", "tax", "nopat", "da2", "capex", "bfr", "dbfr", "lease", "fcff",
             "per", "fac", "pv"]
    for i, n in enumerate(names):
        R[n] = top + i
    summ = ["s_sum", "s_tvg", "s_pvtvg", "s_tvm", "s_pvtvm", "s_evg", "s_evm", "s_ev", "s_nd", "s_adj", "s_eq", "s_ps",
            "s_up", "s_tvpg", "s_tvpm", "s_impm", "s_impg", "s_evebitda", "s_pe"]
    for i, n in enumerate(summ):
        R[n] = top + 20 + i
    for k, v in R.items():
        DC.r[f"{tag}.{k}"] = v
    for col in range(1, 8):
        ws.cell(R["title"], col).fill = fill_hdr
    ws.cell(R["title"], 1, f"SCÉNARIO {name.upper()}").font = font(True, "FFFFFF", 11)
    # en-têtes
    ws.cell(R["hdr"], 1, "").fill = fill_sec
    ws[f"B{R['hdr']}"].value = f"=Saisie!D{SA.r['annee']}"
    for i, c in enumerate(DCOLS):
        ws[f"{c}{R['hdr']}"].value = f'=IFERROR(VALUE(Saisie!$D${SA.r["annee"]})+{i + 1},"N+{i + 1}")'
    for c in ["B"] + DCOLS:
        ws[f"{c}{R['hdr']}"].font = font(True)
        ws[f"{c}{R['hdr']}"].fill = fill_sec
        ws[f"{c}{R['hdr']}"].alignment = Alignment(horizontal="center")
        ws[f"{c}{R['hdr']}"].number_format = "0"
    labels = {
        "g": "Croissance du chiffre d'affaires", "ca": "Chiffre d'affaires", "m": "Marge d'EBITDA", "ebitda": "EBITDA",
        "da": "− Dotations aux amortissements", "ebit": "EBIT", "tax": "− Impôt sur l'EBIT (taux normatif)",
        "nopat": "NOPAT", "da2": "+ Dotations aux amortissements", "capex": "− Investissements (capex)",
        "bfr": "BFR total (fin de période)", "dbfr": "− Variation du BFR", "lease": "− Loyers IFRS 16 (si traitement = 1)", "fcff": "FCFF (flux de trésorerie disponible)",
        "per": "Période d'actualisation (années)", "fac": "Facteur d'actualisation", "pv": "FCFF actualisé",
        "s_sum": "Somme des FCFF actualisés", "s_tvg": "Valeur terminale — Gordon-Shapiro", "s_pvtvg": "Valeur actuelle de la VT (Gordon)",
        "s_tvm": "Valeur terminale — multiple de sortie", "s_pvtvm": "Valeur actuelle de la VT (multiple)",
        "s_evg": "Valeur d'entreprise — Gordon-Shapiro", "s_evm": "Valeur d'entreprise — multiple de sortie",
        "s_ev": "VALEUR D'ENTREPRISE RETENUE (pondérée)", "s_nd": "− Dette financière nette (dernier bilan)",
        "s_adj": "− Minoritaires + actifs non opérationnels", "s_eq": "VALEUR DES CAPITAUX PROPRES",
        "s_ps": "Valeur par action (devise)", "s_up": "Potentiel vs cours actuel",
        "s_tvpg": "Poids de la VT dans la VE (Gordon)", "s_tvpm": "Poids de la VT dans la VE (multiple)",
        "s_impm": "Multiple de sortie implicite (via Gordon)", "s_impg": "Croissance perpétuelle implicite (via multiple)",
        "s_evebitda": "VE / EBITDA N implicite", "s_pe": "PER implicite (capitaux propres / RN N)",
    }
    for k, lab in labels.items():
        bold = k in ("ebitda", "ebit", "nopat", "fcff", "s_ev", "s_eq")
        ws.cell(R[k], 1, lab).font = font(bold)
    f_ = lambda r_, fx, fmt=NUM, bold=False, link=False: (
        ws.__setitem__(r_, fx), None)  # placeholder (non utilisé)

    def put(cell, val, fmt=NUM, bold=False, link=False, fill=None):
        c = ws[cell]
        c.value = val
        c.number_format = fmt
        c.font = font(bold, GREEN if link else BLACK)
        c.alignment = Alignment(horizontal="right")
        if fill:
            c.fill = fill

    # colonne B = dernier exercice réel
    put(f"B{R['ca']}", f"={G('ca', 'D')}", link=True)
    put(f"B{R['m']}", f"={G('m_ebe', 'D')}", PCT, link=True)
    put(f"B{R['ebitda']}", f"={G('ebe', 'D')}", link=True, bold=True)
    put(f"B{R['da']}", f"={G('dot', 'D')}", link=True)
    put(f"B{R['ebit']}", f"={G('ebit', 'D')}", link=True, bold=True)
    put(f"B{R['bfr']}", f"={BR('bfr', 'D')}", link=True)
    put(f"B{R['per']}", f"=-{STUB}", "0.0")
    prev = "B"
    for i, c in enumerate(DCOLS):
        hc = HCOLS[i]
        put(f"{c}{R['g']}", f"=Hypotheses!{hc}{grow_row}", PCT, link=True)
        put(f"{c}{R['ca']}", f"={prev}{R['ca']}*(1+{c}{R['g']})")
        put(f"{c}{R['m']}", f"=Hypotheses!{hc}{marg_row}", PCT, link=True)
        put(f"{c}{R['ebitda']}", f"={c}{R['ca']}*{c}{R['m']}", bold=True)
        put(f"{c}{R['da']}", f"={c}{R['ca']}*Hypotheses!{hc}{H['da']}")
        put(f"{c}{R['ebit']}", f"={c}{R['ebitda']}-{c}{R['da']}", bold=True)
        put(f"{c}{R['tax']}", f"=MAX({c}{R['ebit']},0)*{HYR('tax')}")
        put(f"{c}{R['nopat']}", f"={c}{R['ebit']}-{c}{R['tax']}", bold=True)
        put(f"{c}{R['da2']}", f"={c}{R['da']}")
        put(f"{c}{R['capex']}", f"={c}{R['ca']}*Hypotheses!{hc}{H['capex']}")
        put(f"{c}{R['bfr']}", f"={c}{R['ca']}*Hypotheses!{hc}{H['bfr']}")
        put(f"{c}{R['dbfr']}", f"={c}{R['bfr']}-{prev}{R['bfr']}")
        put(f"{c}{R['lease']}", f"={c}{R['ca']}*Hypotheses!$B${H['lease']}*Hypotheses!$B${H['ifrs']}")
        _sc = f"*(1-{STUB})" if i == 0 else ""
        put(f"{c}{R['fcff']}", f"=({c}{R['nopat']}+{c}{R['da2']}-{c}{R['capex']}-{c}{R['dbfr']}-{c}{R['lease']}){_sc}", bold=True, fill=fill_tot)
        put(f"{c}{R['per']}", (f"=(1-{STUB})*(1-0.5*{MID})+0.5*{MID}" if i == 0 else f"={i + 1}-{STUB}"), "0.00")
        put(f"{c}{R['fac']}", f"=1/(1+{WACC})^({c}{R['per']}-0.5*{MID})", '0.0000')
        put(f"{c}{R['pv']}", f"={c}{R['fcff']}*{c}{R['fac']}")
        prev = c
    ws.cell(R["fcff"], 1).fill = fill_tot
    ws[f"B{R['fcff']}"].fill = fill_tot
    # synthèse
    sv = lambda k: f"B{R[k]}"
    put(sv("s_sum"), f"=SUM(C{R['pv']}:G{R['pv']})")
    put(sv("s_tvg"), f"=G{R['fcff']}*(1+{GG})/({WACC}-{GG})")
    put(sv("s_pvtvg"), f"={sv('s_tvg')}/(1+{WACC})^(5-{STUB}-0.5*{MID})")
    put(sv("s_tvm"), f"=(G{R['ebitda']}-G{R['lease']})*{HYR('mult')}")
    put(sv("s_pvtvm"), f"={sv('s_tvm')}/(1+{WACC})^(5-{STUB})")
    put(sv("s_evg"), f"={sv('s_sum')}+{sv('s_pvtvg')}")
    put(sv("s_evm"), f"={sv('s_sum')}+{sv('s_pvtvm')}")
    put(sv("s_ev"), f"={HYR('wg')}*{sv('s_evg')}+(1-{HYR('wg')})*{sv('s_evm')}", bold=True, fill=fill_tot)
    put(sv("s_nd"), f"=IF({HYR('ifrs')}=1,{BR('dn16', 'D')},{BR('dn', 'D')})", link=True)
    put(sv("s_adj"), f"={HYR('minor')}-{HYR('nonop')}", link=True)
    put(sv("s_eq"), f"={sv('s_ev')}-{sv('s_nd')}-{sv('s_adj')}", bold=True, fill=fill_tot)
    ws.cell(R["s_ev"], 1).fill = fill_tot
    ws.cell(R["s_eq"], 1).fill = fill_tot
    put(sv("s_ps"), f'=IF({SABS("actions")}>0,{sv("s_eq")}*{SABS("mult")}/{SABS("actions")},"n/d")', '#,##0.00', bold=True)
    put(sv("s_up"), f'=IF(AND(ISNUMBER({sv("s_ps")}),{SABS("cours")}>0),{sv("s_ps")}/{SABS("cours")}-1,"n/d")', PCT)
    put(sv("s_tvpg"), f'=IFERROR({sv("s_pvtvg")}/{sv("s_evg")},"n/d")', PCT)
    put(sv("s_tvpm"), f'=IFERROR({sv("s_pvtvm")}/{sv("s_evm")},"n/d")', PCT)
    put(sv("s_impm"), f'=IFERROR({sv("s_tvg")}/(G{R["ebitda"]}-G{R["lease"]}),"n/d")', MULT)
    put(sv("s_impg"), f'=IFERROR(({sv("s_tvm")}*{WACC}-G{R["fcff"]})/({sv("s_tvm")}+G{R["fcff"]}),"n/d")', PCT)
    put(sv("s_evebitda"), f'=IFERROR({sv("s_ev")}/B{R["ebitda"]},"n/d")', MULT)
    put(sv("s_pe"), f'=IFERROR({sv("s_eq")}/{G("rn", "D")},"n/d")', MULT)
    return R


DC.cur = 4
bear = dcf_block(4, "bear", "Bear", H["bear_g"], H["bear_m"])
base = dcf_block(44, "base", "Base", H["base_g"], H["base_m"])
bull = dcf_block(84, "bull", "Bull", H["bull_g"], H["bull_m"])

# ---- sensibilités sur le scénario Base
top = 126
for col in range(1, 8):
    ws_dc.cell(top, col).fill = fill_hdr
ws_dc.cell(top, 1, "SENSIBILITÉS — scénario Base (valeur par action si le nombre d'actions est renseigné, sinon valeur des capitaux propres)").font = font(True, "FFFFFF", 11)
fc = base
FC = f"$C${fc['fcff']}:$G${fc['fcff']}"
PR = f"$C${fc['per']}:$G${fc['per']}"
EB5 = f"($G${fc['ebitda']}-$G${fc['lease']})"
FC5 = f"$G${fc['fcff']}"
ND_ = f"$B${fc['s_nd']}"
ADJ = f"$B${fc['s_adj']}"


def wrap_ps(ev):
    eq = f"({ev}-{ND_}-{ADJ})"
    return f'=IF({SABS("actions")}>0,{eq}*{SABS("mult")}/{SABS("actions")},{eq})'


def sens(row0, title, col_label, col_offsets, col_fmt, col_center, ev_expr):
    ws = ws_dc
    ws.cell(row0, 1, title).font = font(True, "1F3864")
    ws.cell(row0 + 1, 1, f"WACC (lignes) × {col_label} (colonnes)").font = font(italic=True, color="595959")
    hdr = row0 + 1
    for j, off in enumerate(col_offsets):
        cell = ws[f"{DCOLS[j]}{hdr}"]
        cell.value = f"={col_center}+({off})"
        cell.number_format = col_fmt
        cell.font = font(True)
        cell.fill = fill_sec
        cell.alignment = Alignment(horizontal="center")
    for i, off in enumerate([-0.01, -0.005, 0, 0.005, 0.01]):
        rr = row0 + 2 + i
        w = ws[f"B{rr}"]
        w.value = f"={WACC}+({off})"
        w.number_format = '0.0%'
        w.font = font(True)
        w.fill = fill_sec
        for j in range(5):
            c = DCOLS[j]
            cell = ws[f"{c}{rr}"]
            cell.value = wrap_ps(ev_expr(f"$B{rr}", f"{c}${hdr}"))
            cell.number_format = '#,##0.00' if True else NUM
            cell.font = font(bold=(i == 2 and j == 2))
            if i == 2 and j == 2:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
    return hdr


def ev_gordon(w, gg):
    return (f"(SUMPRODUCT({FC}/(1+{w})^({PR}-0.5*{MID}))"
            f"+{FC5}*(1+{gg})/({w}-{gg})/(1+{w})^(5-{STUB}-0.5*{MID}))")


def ev_mult(w, m):
    return (f"(SUMPRODUCT({FC}/(1+{w})^({PR}-0.5*{MID}))"
            f"+{EB5}*{m}/(1+{w})^(5-{STUB}))")


h1 = sens(128, "Table 1 — WACC × croissance perpétuelle (Gordon-Shapiro)", "g", [-0.01, -0.005, 0, 0.005, 0.01], '0.0%', GG, ev_gordon)
h2 = sens(137, "Table 2 — WACC × multiple de sortie VE/EBITDA", "multiple", [-2, -1, 0, 1, 2], MULT, HYR("mult"), ev_mult)
DC.r["sens1_hdr"] = h1
DC.r["sens2_hdr"] = h2
ws_dc.freeze_panes = "B4"

# =====================================================================
# SYNTHESE
# =====================================================================
SY = Sh(ws_sy, {"A": 46, "B": 15, "C": 15, "D": 15, "E": 15, "F": 3})
ws_sy["A1"].value = f'=Saisie!B{SA.r["nom"]}&" — Synthèse financière"'
ws_sy["A1"].font = font(True, "1F3864", 14)
ws_sy["A2"].value = f'="Montants en "&Saisie!B{SA.r["unite"]}&" — exercices "&Saisie!B{SA.r["annee"]}&" à "&Saisie!D{SA.r["annee"]}'
ws_sy["A2"].font = font(italic=True, color="595959")
SY.cur = 4
SY.section("CONTRÔLES DE COHÉRENCE", "E")
ctr = [
    ("Bilan équilibré (actif = passif) sur les 3 exercices",
     f"=IF(MAX(ABS(Saisie!B{SA.r['ecart_bilan']}),ABS(Saisie!C{SA.r['ecart_bilan']}),ABS(Saisie!D{SA.r['ecart_bilan']}))<=1,\"OK\",\"ERREUR\")"),
    ("Résultat du compte de résultat = résultat au bilan",
     f"=IF(MAX(ABS(SIG_CAF!B{SG.r['ctrl_rn']}),ABS(SIG_CAF!C{SG.r['ctrl_rn']}),ABS(SIG_CAF!D{SG.r['ctrl_rn']}))<=1,\"OK\",\"ERREUR\")"),
    ("Trésorerie nette = FRNG − BFR",
     f"=IF(MAX(ABS(Bilan_BFR!B{BF.r['tn_ck']}),ABS(Bilan_BFR!C{BF.r['tn_ck']}),ABS(Bilan_BFR!D{BF.r['tn_ck']}))<=1,\"OK\",\"ERREUR\")"),
    ("Hypothèses DCF : WACC > croissance perpétuelle", f"=IF({WACC}>{GG},\"OK\",\"ERREUR\")"),
    ("Probabilités des scénarios = 100 %", f"=IF(ABS(Hypotheses!B{H['p_chk']}-1)<0.0001,\"OK\",\"ERREUR\")"),
]
ctr_rows = []
for lab, fx in ctr:
    r = SY.cur
    ws_sy.cell(r, 1, lab).font = font()
    c = ws_sy.cell(r, 2, fx)
    c.font = font(True)
    c.alignment = Alignment(horizontal="center")
    ctr_rows.append(r)
    SY.cur += 1
r = SY.cur
ws_sy.cell(r, 1, "Contrôle global").font = font(True)
c = ws_sy.cell(r, 2, f'=IF(COUNTIF(B{ctr_rows[0]}:B{ctr_rows[-1]},"OK")={len(ctr)},"OK","À CORRIGER")')
c.font = font(True)
c.alignment = Alignment(horizontal="center")
SY.r["ctrl_glob"] = r
SY.r["ctrl_first"] = ctr_rows[0]
SY.cur += 2
for rr in ctr_rows + [r]:
    ws_sy.conditional_formatting.add(f"B{rr}", FormulaRule(formula=[f'$B{rr}="OK"'], fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")))
    ws_sy.conditional_formatting.add(f"B{rr}", FormulaRule(formula=[f'$B{rr}<>"OK"'], fill=PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")))

SY.section("INDICATEURS CLÉS", "E")
hdr_r = SY.header(["", "", ""])
for c in COLS:
    ws_sy[f"{c}{hdr_r}"].value = f"=Saisie!{c}{SA.r['annee']}"
ws_sy[f"E{hdr_r}"].value = "Évol. N / N-1"
ws_sy[f"E{hdr_r}"].font = font(True, "FFFFFF")
ws_sy[f"E{hdr_r}"].fill = fill_hdr
ws_sy[f"E{hdr_r}"].alignment = Alignment(horizontal="center")


def kpi(key, label, src, fmt, evol=True, first=None, bold=False):
    r = SY.line(key, label, lambda c, p: f"={src(c)}", fmt, first=first, link=True, bold=bold)
    cell = ws_sy[f"E{r}"]
    if evol == "pts":
        cell.value = f'=IFERROR(D{r}-C{r},"n/d")'
        cell.number_format = '+0.0%;-0.0%;"-"'
    elif evol:
        cell.value = f'=IFERROR(D{r}/C{r}-1,"n/d")'
        cell.number_format = '+0.0%;-0.0%;"-"'
    cell.font = font()
    cell.alignment = Alignment(horizontal="right")


kpi("k_ca", "Chiffre d'affaires", lambda c: G("ca", c), NUM, bold=True)
kpi("k_g", "Croissance du CA", lambda c: RAR("g_ca", c), PCT, evol="pts")
kpi("k_ebitda", "EBITDA", lambda c: G("ebe", c), NUM, bold=True)
kpi("k_mebitda", "Marge d'EBITDA", lambda c: G("m_ebe", c), PCT, evol="pts")
kpi("k_ebit", "EBIT", lambda c: G("ebit", c), NUM)
kpi("k_rn", "Résultat net", lambda c: G("rn", c), NUM, bold=True)
kpi("k_caf", "CAF", lambda c: G("caf", c), NUM)
kpi("k_fcff", "FCFF", lambda c: CFR("fcff", c), NUM, first="n/d")
kpi("k_dn", "Dette financière nette", lambda c: BR("dn", c), NUM)
kpi("k_lev", "Dette nette / EBITDA", lambda c: RAR("nd_ebitda", c), MULT, evol=False)
kpi("k_roe", "ROE", lambda c: RAR("roe", c), PCT, evol="pts")
kpi("k_roic", "ROIC", lambda c: RAR("roic", c), PCT, evol="pts")
kpi("k_liq", "Liquidité générale", lambda c: RAR("liq_g", c), D2, evol=False)
kpi("k_bfrj", "BFR en jours de CA", lambda c: BR("bfr_j", c), D1, evol=False)
kpi("k_z", "Score Z'' d'Altman", lambda c: SCR("z", c), D2, evol=False)
kpi("k_p", "Score de Piotroski (0-9)", lambda c: SCR("pscore", c), INT0, evol=False, first="n/d")
SY.gap()
SY.section("VALORISATION DCF PAR SCÉNARIO", "E")
vh = SY.header(["Bear", "Base", "Bull", "Pondéré"])
val_rows = [("v_ev", "Valeur d'entreprise", "s_ev", NUM), ("v_eq", "Valeur des capitaux propres", "s_eq", NUM),
            ("v_ps", "Valeur par action (devise)", "s_ps", '#,##0.00'), ("v_up", "Potentiel vs cours actuel", "s_up", PCT),
            ("v_tv", "Poids de la VT dans la VE (Gordon)", "s_tvpg", PCT)]
for key, lab, sk, fm in val_rows:
    r = SY.cur
    SY.r[key] = r
    SY.cur += 1
    ws_sy.cell(r, 1, lab).font = font(key in ("v_eq", "v_ps"))
    for col, tag in zip(["B", "C", "D"], ["bear", "base", "bull"]):
        c = ws_sy[f"{col}{r}"]
        c.value = f"=DCF!B{DC.r[f'{tag}.{sk}']}"
        c.number_format = fm
        c.font = font(key in ("v_eq", "v_ps"), GREEN)
        c.alignment = Alignment(horizontal="right")
    e = ws_sy[f"E{r}"]
    if key in ("v_ev", "v_eq"):
        e.value = f"=SUMPRODUCT(B{r}:D{r},Hypotheses!$B${H['p']}:$D${H['p']})"
    elif key == "v_ps":
        e.value = f'=IF(ISNUMBER(B{r}),SUMPRODUCT(B{r}:D{r},Hypotheses!$B${H["p"]}:$D${H["p"]}),"n/d")'
    elif key == "v_up":
        e.value = f'=IF(AND(ISNUMBER(E{SY.r["v_ps"]}),{SABS("cours")}>0),E{SY.r["v_ps"]}/{SABS("cours")}-1,"n/d")'
    else:
        e.value = ""
    e.number_format = fm
    e.font = font(True)
    e.alignment = Alignment(horizontal="right")
SY.gap()
# graphiques
ch = BarChart()
ch.type = "col"
ch.title = "CA, EBITDA et résultat net"
ch.roundedCorners = False
data = Reference(ws_sy, min_col=1, max_col=4, min_row=SY.r["k_ca"], max_row=SY.r["k_ca"])
ch.add_data(Reference(ws_sy, min_col=1, max_col=4, min_row=SY.r["k_ca"]), from_rows=True, titles_from_data=True)
ch.add_data(Reference(ws_sy, min_col=1, max_col=4, min_row=SY.r["k_ebitda"]), from_rows=True, titles_from_data=True)
ch.add_data(Reference(ws_sy, min_col=1, max_col=4, min_row=SY.r["k_rn"]), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws_sy, min_col=2, max_col=4, min_row=hdr_r))
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.legend.overlay = False
ch.height, ch.width = 8, 15
ws_sy.add_chart(ch, "G4")
ch2 = BarChart()
ch2.type = "col"
ch2.title = "Valeur des capitaux propres par scénario"
ch2.roundedCorners = False
ch2.add_data(Reference(ws_sy, min_col=1, max_col=4, min_row=SY.r["v_eq"]), from_rows=True, titles_from_data=True)
ch2.set_categories(Reference(ws_sy, min_col=2, max_col=4, min_row=vh))
ch2.x_axis.delete = False
ch2.y_axis.delete = False
ch2.legend = None
ch2.height, ch2.width = 8, 15
ws_sy.add_chart(ch2, "G21")

# =====================================================================
# COMPARABLES / WACC_PAIRS / LBO / MA
# =====================================================================
PEERS = [] if ARGS.blank else D.get("peers", [])
LBD = D.get("lbo", {})
MAD = D.get("ma", {})
NP = 8  # emplacements de comparables


def cell_(ws, addr, val, fmt=None, kind="calc", bold=False, align="right"):
    c = ws[addr]
    c.value = val
    if fmt:
        c.number_format = fmt
    if kind == "inp":
        c.font, c.fill, c.border = font(bold, BLUE), fill_inp, box
    elif kind == "link":
        c.font = font(bold, GREEN)
    else:
        c.font = font(bold)
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def lab_(ws, r, text, bold=False, col=1, note=False):
    c = ws.cell(r, col, text)
    c.font = font(bold, "595959" if note else BLACK, 9 if note else 10, italic=note)
    return c


def hdr_(ws, r, labels, start=1):
    for i, t in enumerate(labels):
        c = ws.cell(r, start + i, t)
        c.font = font(True, "FFFFFF")
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sec_(ws, r, text, last=8):
    for col in range(1, last + 1):
        ws.cell(r, col).fill = fill_sec
    ws.cell(r, 1, text).font = font(True, "1F3864")


# ------------------------------------------------------------- Comparables
ws = ws_cp
for col, w in zip("ABCDEFGHI", [34, 12, 13, 13, 13, 13, 9, 13, 60]):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws["A1"].value = "COMPARABLES BOURSIERS — multiple VE / EBITDA"
ws["A1"].font = font(True, "1F3864", 14)
ws["A2"].value = ("Données de pairs à saisir ou mettre à jour (unité des états). Les multiples issus d'agrégateurs gratuits sont datés "
                  "différemment et parfois erronés : vérifier la source avant usage. Exclure un pair = mettre 0 en colonne « Inclus ».")
ws["A2"].font = font(italic=True, color="595959")
sec_(ws, 4, "1. PANEL DE PAIRS", 9)
hdr_(ws, 5, ["Société", "Date", "Capitalisation", "Valeur d'entreprise", "EBITDA", "VE/EBITDA publié", "Inclus", "VE/EBITDA retenu", "Source / remarque", "Résultat net (LTM)", "PER retenu"])
ws.column_dimensions["J"].width = 14
ws.column_dimensions["K"].width = 11
CP0 = 6
for i in range(NP):
    r = CP0 + i
    pr = PEERS[i] if i < len(PEERS) else {}
    cell_(ws, f"A{r}", pr.get("nom"), kind="inp", align="left")
    cell_(ws, f"B{r}", pr.get("date"), kind="inp", align="center")
    cell_(ws, f"C{r}", pr.get("mcap"), NUM, "inp")
    cell_(ws, f"D{r}", pr.get("ev"), NUM, "inp")
    cell_(ws, f"E{r}", pr.get("ebitda"), NUM, "inp")
    cell_(ws, f"F{r}", pr.get("ev_ebitda"), MULT, "inp")
    cell_(ws, f"G{r}", (0 if pr.get("exclu") else 1) if pr else 0, INT0, "inp", align="center")
    cell_(ws, f"H{r}", f'=IF(G{r}<>1,"",IF(ISNUMBER(F{r}),F{r},IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),IF(E{r}>0,D{r}/E{r},""),"")))', MULT)
    cell_(ws, f"I{r}", pr.get("source"), kind="inp", align="left")
    cell_(ws, f"J{r}", pr.get("rn"), NUM, "inp")
    cell_(ws, f"K{r}", f'=IF(G{r}<>1,"",IF(AND(ISNUMBER(C{r}),ISNUMBER(J{r})),IF(J{r}>0,C{r}/J{r},""),""))', MULT)
CP1 = CP0 + NP - 1
KR = f"$K${CP0}:$K${CP1}"
HR = f"$H${CP0}:$H${CP1}"
r = CP1 + 2
sec_(ws, r, "2. STATISTIQUES DU PANEL", 9)
stats = [("Nombre de pairs retenus", "COUNT({R})", INT0), ("Minimum", 'IF(COUNT({R})>0,MIN({R}),"n/d")', MULT),
         ("1er quartile", 'IF(COUNT({R})>1,QUARTILE({R},1),"n/d")', MULT), ("Médiane", 'IF(COUNT({R})>0,MEDIAN({R}),"n/d")', MULT),
         ("Moyenne", 'IF(COUNT({R})>0,AVERAGE({R}),"n/d")', MULT), ("3e quartile", 'IF(COUNT({R})>1,QUARTILE({R},3),"n/d")', MULT),
         ("Maximum", 'IF(COUNT({R})>0,MAX({R}),"n/d")', MULT)]
hdr_(ws, r, ["Statistiques du panel", "VE/EBITDA", "PER"])
for c_ in (1, 2, 3):
    ws.cell(r, c_).fill = fill_hdr
CPS = {}
for j, (t, fx, fm) in enumerate(stats):
    rr = r + 1 + j
    lab_(ws, rr, t, bold=(t == "Médiane"))
    cell_(ws, f"B{rr}", "=" + fx.format(R=HR), fm, bold=(t == "Médiane"))
    cell_(ws, f"C{rr}", "=" + fx.format(R=KR), fm, bold=(t == "Médiane"))
    CPS[t] = rr
r = r + len(stats) + 2
sec_(ws, r, "3. VALORISATION DE LA SOCIÉTÉ PAR LES MULTIPLES (EBITDA N publié, dette nette incluant les dettes locatives)", 9)
hdr_(ws, r + 1, ["", "Bas (Q1)", "Médiane", "Haut (Q3)"])
CPV = {}
lines_ = [("mult", "Multiple VE/EBITDA appliqué", None, MULT), ("ebitda", "EBITDA retenu (LTM si saisi en bas de feuille, sinon N)", None, NUM), ("ev", "Valeur d'entreprise", None, NUM),
          ("nd", "− Dette financière nette (avec dettes locatives)", None, NUM), ("adj", "− Minoritaires, retraites + actifs non opérationnels", None, NUM),
          ("eq", "Valeur des capitaux propres", None, NUM), ("ps", "Valeur par action (devise)", None, '#,##0.00')]
for j, (k, t, _, fm) in enumerate(lines_):
    CPV[k] = r + 2 + j
    lab_(ws, CPV[k], t, bold=k in ("eq", "ps"))
src_stat = {"B": CPS["1er quartile"], "C": CPS["Médiane"], "D": CPS["3e quartile"]}
for c, sr in src_stat.items():
    cell_(ws, f"{c}{CPV['mult']}", f"=B{sr}", MULT)
    cell_(ws, f"{c}{CPV['ebitda']}", f"=IF(ISNUMBER($B$46),$B$46,{G('ebe', 'D')})", NUM, "link")
    cell_(ws, f"{c}{CPV['ev']}", f'=IFERROR({c}{CPV["mult"]}*{c}{CPV["ebitda"]},"n/d")', NUM)
    cell_(ws, f"{c}{CPV['nd']}", f"={BR('dn', 'D')}", NUM, "link")
    cell_(ws, f"{c}{CPV['adj']}", f"={HYR('minor')}-{HYR('nonop')}", NUM, "link")
    cell_(ws, f"{c}{CPV['eq']}", f'=IFERROR({c}{CPV["ev"]}-{c}{CPV["nd"]}-{c}{CPV["adj"]},"n/d")', NUM, bold=True)
    cell_(ws, f"{c}{CPV['ps']}", f'=IFERROR(IF({SABS("actions")}>0,{c}{CPV["eq"]}*{SABS("mult")}/{SABS("actions")},"n/d"),"n/d")', '#,##0.00', bold=True)
rp = CPV["ps"] + 2
sec_(ws, rp, "4. VALORISATION PAR LE PER (résultat net N × PER du panel)", 9)
hdr_(ws, rp + 1, ["", "Bas (Q1)", "Médiane", "Haut (Q3)"])
CPP = {}
for j, (k, t, fm) in enumerate([("per", "PER appliqué", MULT), ("rn", "Résultat net retenu (LTM si saisi, sinon N)", NUM), ("eq", "Valeur des capitaux propres", NUM),
                                ("ps", "Valeur par action (devise)", '#,##0.00')]):
    CPP[k] = rp + 2 + j
    lab_(ws, CPP[k], t, bold=k in ("eq", "ps"))
for c, sr in src_stat.items():
    cell_(ws, f"{c}{CPP['per']}", f"=C{sr}", MULT)
    cell_(ws, f"{c}{CPP['rn']}", f"=IF(ISNUMBER($B$47),$B$47,{G('rn', 'D')})", NUM, "link")
    cell_(ws, f"{c}{CPP['eq']}", f'=IFERROR({c}{CPP["per"]}*{c}{CPP["rn"]},"n/d")', NUM, bold=True)
    cell_(ws, f"{c}{CPP['ps']}", f'=IFERROR(IF({SABS("actions")}>0,{c}{CPP["eq"]}*{SABS("mult")}/{SABS("actions")},"n/d"),"n/d")', '#,##0.00', bold=True)
lab_(ws, CPP["ps"] + 1, "Le PER peut être faussé par des éléments non récurrents (plus-values latentes, dépréciations) : comparer à un résultat net normalisé.", note=True)
_ltm = {} if ARGS.blank else D.get("ltm", {})
sec_(ws, 44, "5. BASE LTM (12 DERNIERS MOIS) — optionnelle ; les multiples de marché sont des LTM, ne pas les appliquer à un exercice ancien", 9)
lab_(ws, 45, "Dernier bilan intermédiaire utilisé (date)")
cell_(ws, "B45", (_dt.date.fromisoformat(_ltm["date"]) if _ltm.get("date") else None), "dd/mm/yyyy", "inp")
lab_(ws, 46, "EBITDA LTM (M devise) = EBITDA N + cumul de l'exercice en cours − cumul de l'an passé")
cell_(ws, "B46", _ltm.get("ebitda"), NUM, "inp")
lab_(ws, 47, "Résultat net LTM (M devise)")
cell_(ws, "B47", _ltm.get("rn"), NUM, "inp")
lab_(ws, 48, "Source et calcul :", bold=True)
cell_(ws, "B48", _ltm.get("source"), None, "inp", align="left")
lab_(ws, CPV["ps"] + 1, "Un multiple de comparables est une photographie de marché : il intègre la croissance, la taille et la liquidité des pairs, pas celles de la cible.", note=True)

# --------------------------------------------------------------- WACC_Pairs
ws = ws_wp
for col, w in zip("ABCDEFGHI", [40, 12, 13, 13, 10, 11, 13, 9, 13]):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws["A1"].value = "BÊTA DES COMPARABLES — désendettement / ré-endettement (Hamada)"
ws["A1"].font = font(True, "1F3864", 14)
ws["A2"].value = "β désendetté = β endetté / (1 + (1 − t) × D/E). Médiane du panel ré-endettée à la structure cible (Hypotheses). Bêtas d'agrégateurs : horizon et indice variables."
ws["A2"].font = font(italic=True, color="595959")
sec_(ws, 4, "1. BÊTAS OBSERVÉS", 9)
hdr_(ws, 5, ["Société", "β endetté", "Capitalisation", "Dette nette (VE − cap.)", "Taux d'IS", "D/E", "β désendetté", "Inclus", "β retenu"])
for i in range(NP):
    r = 6 + i
    pr = PEERS[i] if i < len(PEERS) else {}
    cr = CP0 + i
    cell_(ws, f"A{r}", f'=IF(Comparables!A{cr}="","",Comparables!A{cr})', kind="link", align="left")
    cell_(ws, f"B{r}", pr.get("beta"), D2, "inp")
    cell_(ws, f"C{r}", f'=IF(ISNUMBER(Comparables!C{cr}),Comparables!C{cr},"")', NUM, "link")
    cell_(ws, f"D{r}", f'=IF(AND(ISNUMBER(Comparables!C{cr}),ISNUMBER(Comparables!D{cr})),MAX(Comparables!D{cr}-Comparables!C{cr},0),"")', NUM)
    cell_(ws, f"E{r}", pr.get("tax", 0.25) if pr else None, PCT, "inp")
    cell_(ws, f"F{r}", f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),IF(C{r}>0,D{r}/C{r},""),"")', PCT)
    cell_(ws, f"G{r}", f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(F{r}),ISNUMBER(E{r})),B{r}/(1+(1-E{r})*F{r}),"")', D2)
    cell_(ws, f"H{r}", f"=IF(Comparables!G{cr}=1,1,0)", INT0, "link", align="center")
    cell_(ws, f"I{r}", f'=IF(H{r}=1,G{r},"")', D2)
WP1 = 6 + NP - 1
lab_(ws, 15, "Nombre de bêtas retenus"); cell_(ws, "B15", f"=COUNT(I6:I{WP1})", INT0)
lab_(ws, 16, "Médiane des β désendettés"); cell_(ws, "B16", f'=IF(COUNT(I6:I{WP1})>0,MEDIAN(I6:I{WP1}),"n/d")', D2, bold=True)
lab_(ws, 17, "Moyenne des β désendettés"); cell_(ws, "B17", f'=IF(COUNT(I6:I{WP1})>0,AVERAGE(I6:I{WP1}),"n/d")', D2)
lab_(ws, 18, "Structure cible D/E = poids dette / (1 − poids dette)"); cell_(ws, "B18", f"={HYR('wd')}/(1-{HYR('wd')})", PCT, "link")
lab_(ws, 19, "Taux d'impôt normatif"); cell_(ws, "B19", f"={HYR('tax')}", PCT, "link")
lab_(ws, 20, "β ré-endetté (médiane) → mode 2 des Hypotheses", bold=True)
cell_(ws, "B20", '=IFERROR(B16*(1+(1-B19)*B18),"n/d")', D2, bold=True)
ws["B20"].fill = fill_tot
lab_(ws, 21, "β ré-endetté ajusté Blume (0,67 × β + 0,33) → mode 3", bold=True)
cell_(ws, "B21", '=IFERROR(0.67*B20+0.33,"n/d")', D2, bold=True)
ws["B21"].fill = fill_tot
lab_(ws, 23, "La médiane d'un petit panel (3-5 pairs) est fragile ; comparer au bêta saisi et documenter le choix. Le mode 1 conserve le bêta saisi à la main.", note=True)

# --------------------------------------------------------------------- LBO
ws = ws_lbo
for col, w in zip("ABCDEFGH", [58, 13, 13, 13, 13, 13, 13, 4]):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws["A1"].value = "LBO — faisabilité d'une acquisition par effet de levier (scénario Base du DCF)"
ws["A1"].font = font(True, "1F3864", 14)
ws["A2"].value = ("Analyse « ability-to-pay » : quel prix un sponsor peut-il payer pour atteindre un TRI cible ? Hypothèses de financement ci-dessous, "
                  "flux issus du scénario Base. Illustratif : pas de structure de dette réelle (tranches, covenants, OID).")
ws["A2"].font = font(italic=True, color="595959")
sec_(ws, 4, "1. HYPOTHÈSES DE L'OPÉRATION", 7)
lbo_in = [("entry", "Multiple d'entrée VE / EBITDA N", LBD.get("entry_mult", 8.0), MULT), ("fees", "Frais d'opération (% de la VE)", LBD.get("fees", 0.025), PCT),
          ("lev", "Dette totale / EBITDA N", LBD.get("lev", 3.0), MULT), ("rate", "Taux d'intérêt de la dette", LBD.get("rate", 0.065), PCT),
          ("sweep", "Part du cash-flow disponible affectée au remboursement (cash sweep)", LBD.get("sweep", 1.0), PCT),
          ("exit", "Multiple de sortie VE / EBITDA (année N+5)", LBD.get("exit_mult", 8.0), MULT)]
LB = {}
for j, (k, t, v, fm) in enumerate(lbo_in):
    r = 5 + j
    LB[k] = f"$B${r}"
    lab_(ws, r, t)
    cell_(ws, f"B{r}", v, fm, "inp")
lab_(ws, 11, "EBITDA N retenu (après loyers IFRS 16 si traitement = 1)")
cell_(ws, "B11", f"={G('ebe', 'D')}-{HYR('ifrs')}*{G('loyers', 'D')}", NUM, "link")
sec_(ws, 13, "2. SOURCES ET EMPLOIS À L'ENTRÉE", 7)
lab_(ws, 14, "Valeur d'entreprise payée"); cell_(ws, "B14", f"={LB['entry']}*B11", NUM)
lab_(ws, 15, "Frais d'opération"); cell_(ws, "B15", f"={LB['fees']}*B14", NUM)
lab_(ws, 16, "Total des emplois", True); cell_(ws, "B16", "=B14+B15", NUM, bold=True)
lab_(ws, 17, "Dette levée"); cell_(ws, "B17", f"={LB['lev']}*B11", NUM)
lab_(ws, 18, "Fonds propres du sponsor", True); cell_(ws, "B18", "=B16-B17", NUM, bold=True)
lab_(ws, 19, "Fonds propres / total des emplois"); cell_(ws, "B19", '=IFERROR(B18/B16,"n/d")', PCT)
sec_(ws, 21, "3. PROJECTION DE LA DETTE (flux du scénario Base)", 7)
hdr_(ws, 22, ["", "N", "N+1", "N+2", "N+3", "N+4", "N+5"])
lr = dict(ebitda=23, fcff=24, int_=25, cfa=26, open_=27, repay=28, close=29, cash=30, lev=31)
lt = {"ebitda": "EBITDA (après loyers)", "fcff": "FCFF (avant frais financiers)", "int_": "− Intérêts après impôt", "cfa": "Cash-flow disponible pour la dette",
      "open_": "Dette — début de période", "repay": "− Remboursement", "close": "Dette — fin de période", "cash": "Trésorerie cumulée (cash non affecté)",
      "lev": "Dette nette / EBITDA"}
for k, rr in lr.items():
    lab_(ws, rr, lt[k], bold=k in ("close", "cfa"))
cell_(ws, "B23", "=B11", NUM)
cell_(ws, "B29", "=B17", NUM)
cell_(ws, "B30", 0, NUM)
cell_(ws, "B31", '=IFERROR((B29-B30)/B23,"n/d")', MULT)
pv_ = "B"
for i, c in enumerate(DCOLS):
    dr = lambda k: f"DCF!{c}{base[k]}"
    cell_(ws, f"{c}23", f"={dr('ebitda')}-{dr('lease')}", NUM, "link")
    cell_(ws, f"{c}24", f"={dr('fcff')}", NUM, "link")
    cell_(ws, f"{c}27", f"={pv_}29", NUM)
    cell_(ws, f"{c}25", f"={c}27*{LB['rate']}*(1-{HYR('tax')})", NUM)
    cell_(ws, f"{c}26", f"={c}24-{c}25", NUM)
    cell_(ws, f"{c}28", f"=MIN({c}27,MAX({c}26,0)*{LB['sweep']})", NUM)
    cell_(ws, f"{c}29", f"={c}27-{c}28", NUM, bold=True)
    cell_(ws, f"{c}30", f"={pv_}30+{c}26-{c}28", NUM)
    cell_(ws, f"{c}31", f'=IFERROR(({c}29-{c}30)/{c}23,"n/d")', MULT)
    pv_ = c
sec_(ws, 33, "4. RENDEMENT À LA SORTIE (N+5)", 7)
lab_(ws, 34, "Valeur d'entreprise de sortie"); cell_(ws, "B34", f"={LB['exit']}*G23", NUM)
lab_(ws, 35, "− Dette à la sortie + trésorerie"); cell_(ws, "B35", "=-G29+G30", NUM)
lab_(ws, 36, "Valeur des fonds propres à la sortie", True); cell_(ws, "B36", "=B34+B35", NUM, bold=True)
lab_(ws, 37, "Multiple de capital investi (MOIC)", True); cell_(ws, "B37", '=IFERROR(B36/B18,"n/d")', MULT, bold=True)
lab_(ws, 38, "TRI du sponsor (sans dividendes intermédiaires)", True); cell_(ws, "B38", '=IFERROR(IF(B36>0,(B36/B18)^(1/5)-1,-1),"n/d")', PCT, bold=True)
ws["B38"].fill = fill_tot
sec_(ws, 40, "5. PRIX MAXIMUM POUR UN TRI CIBLE (dette inchangée : elle dépend de l'EBITDA, pas du prix)", 7)
hdr_(ws, 41, ["", "TRI 1", "TRI 2", "TRI 3"])
for rr, t in zip(range(42, 48), ["TRI cible", "Fonds propres max. à l'entrée", "Valeur d'entreprise maximale", "Multiple d'entrée maximal VE/EBITDA N",
                                 "Valeur des capitaux propres de la cible", "Prix maximal par action (devise)"]):
    lab_(ws, rr, t, bold=rr in (46, 47))
tg = LBD.get("irr_targets", [0.25, 0.20, 0.15])
for c, tv in zip("BCD", tg):
    cell_(ws, f"{c}42", tv, PCT, "inp")
    cell_(ws, f"{c}43", f"=$B$36/(1+{c}42)^5", NUM)
    cell_(ws, f"{c}44", f"=({c}43+$B$17)/(1+{LB['fees']})", NUM)
    cell_(ws, f"{c}45", f'=IFERROR({c}44/$B$11,"n/d")', MULT)
    cell_(ws, f"{c}46", f"={c}44-DCF!$B${base['s_nd']}-DCF!$B${base['s_adj']}", NUM, bold=True)
    cell_(ws, f"{c}47", f'=IF({SABS("actions")}>0,{c}46*{SABS("mult")}/{SABS("actions")},"n/d")', '#,##0.00', bold=True)
sec_(ws, 49, "6. SENSIBILITÉ DU TRI — multiple d'entrée (lignes) × multiple de sortie (colonnes)", 7)
for j, off in enumerate([-2, -1, 0, 1, 2]):
    cell_(ws, f"{'CDEFG'[j]}50", f"={LB['exit']}+({off})", MULT, bold=True)
    ws[f"{'CDEFG'[j]}50"].fill = fill_sec
for i, off in enumerate([-2, -1, 0, 1, 2]):
    rr = 51 + i
    cell_(ws, f"B{rr}", f"={LB['entry']}+({off})", MULT, bold=True)
    ws[f"B{rr}"].fill = fill_sec
    for j in range(5):
        c = "CDEFG"[j]
        eq0 = f"($B{rr}*$B$11*(1+{LB['fees']})-$B$17)"
        eq5 = f"({c}$50*$G$23-$G$29+$G$30)"
        cell_(ws, f"{c}{rr}", f'=IFERROR(IF(AND({eq0}>0,{eq5}>0),({eq5}/{eq0})^(1/5)-1,"n/s"),"n/s")', PCT, bold=(i == 2 and j == 2))
lab_(ws, 57, "n/s = fonds propres nuls ou négatifs : l'effet de levier dépasse le prix.", note=True)

# ---------------------------------------------------------------------- MA
ws = ws_ma
for col, w in zip("ABCDEF", [62, 14, 14, 14, 14, 4]):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws["A1"].value = "FUSION-ACQUISITION — accrétion / dilution du BPA (la société est la cible)"
ws["A1"].font = font(True, "1F3864", 14)
ws["A2"].value = ("Un acquéreur (fictif ou réel, à saisir) rachète 100 % de la société. Montants dans l'unité des états ; nombre d'actions en unités. "
                  "Modèle simplifié : pas d'effet de timing, de refinancement de la dette de la cible ni de coûts d'intégration.")
ws["A2"].font = font(italic=True, color="595959")
sec_(ws, 4, "1. ACQUÉREUR ET CONDITIONS", 5)
ma_in = [("nom", "Acquéreur", MAD.get("acq_nom", "Acquéreur"), None), ("rn", "Résultat net de l'acquéreur (N)", MAD.get("acq_rn", 0), NUM),
         ("act", "Nombre d'actions de l'acquéreur", MAD.get("acq_actions", 1), '#,##0'), ("cours", "Cours de l'acquéreur (devise)", MAD.get("acq_cours", 1), '#,##0.00'),
         ("prem", "Prime d'offre sur le cours de la cible", MAD.get("premium", 0.3), PCT), ("pc", "Part payée en trésorerie", MAD.get("pct_cash", 0.3), PCT),
         ("pd", "Part financée par dette nouvelle", MAD.get("pct_debt", 0.4), PCT), ("ps", "Part payée en actions", MAD.get("pct_stock", 0.3), PCT),
         ("rd", "Taux de la dette nouvelle", MAD.get("r_debt", 0.055), PCT), ("rc", "Rendement perdu sur la trésorerie utilisée", MAD.get("r_cash", 0.025), PCT),
         ("fees", "Frais de transaction (% du prix, supportés en année 1)", MAD.get("fees", 0.02), PCT),
         ("syn", "Synergies annuelles avant impôt (pleine année)", MAD.get("syn", 0), NUM)]
MA_ = {}
for j, (k, t, v, fm) in enumerate(ma_in):
    r = 5 + j
    MA_[k] = f"$B${r}"
    lab_(ws, r, t)
    cell_(ws, f"B{r}", v, fm, "inp", align="left" if k == "nom" else "right")
r = 5 + len(ma_in)
lab_(ws, r, "Phasage des synergies (N+1, N+2, N+3)")
for c, v in zip("BCD", MAD.get("phase", [0.5, 1, 1])):
    cell_(ws, f"{c}{r}", v, PCT, "inp")
PH = r
lab_(ws, r + 1, "Écart d'acquisition alloué aux actifs incorporels amortissables (% de l'écart)")
cell_(ws, f"B{r + 1}", MAD.get("ppa_pct", 0.2), PCT, "inp")
lab_(ws, r + 2, "Durée d'amortissement des incorporels (années)")
cell_(ws, f"B{r + 2}", MAD.get("ppa_years", 8), INT0, "inp")
PPA_P, PPA_Y = f"$B${r + 1}", f"$B${r + 2}"
r += 4
sec_(ws, r, "2. PRIX ET FINANCEMENT", 5)
cell_(ws, f"B{r + 1}", f'=IF(ABS({MA_["pc"]}+{MA_["pd"]}+{MA_["ps"]}-1)<0.0001,"OK","Total ≠ 100 %")', align="center", bold=True)
lab_(ws, r + 1, "Contrôle : modes de paiement = 100 %")
rows_ = [("px", "Prix d'offre par action de la cible (devise)", f'={SABS("cours")}*(1+{MA_["prem"]})', '#,##0.00'),
         ("eqv", "Prix des capitaux propres de la cible", f'=B{r + 2}*{SABS("actions")}/{SABS("mult")}', NUM),
         ("gw", "Écart d'acquisition brut (prix − capitaux propres comptables)", f'=B{r + 3}-{BR("cp", "D")}', NUM),
         ("cash", "Trésorerie utilisée", f'=B{r + 3}*{MA_["pc"]}', NUM), ("debt", "Dette nouvelle", f'=B{r + 3}*{MA_["pd"]}', NUM),
         ("stock", "Émission d'actions (valeur)", f'=B{r + 3}*{MA_["ps"]}', NUM),
         ("new", "Nombre d'actions nouvelles émises", f'=B{r + 7}*{SABS("mult")}/{MA_["cours"]}', '#,##0')]
MR = {}
for j, (k, t, fx, fm) in enumerate(rows_):
    rr = r + 2 + j
    MR[k] = rr
    lab_(ws, rr, t)
    cell_(ws, f"B{rr}", fx, fm)
r = r + 2 + len(rows_) + 1
sec_(ws, r, "3. BPA PRO FORMA", 5)
hdr_(ws, r + 1, ["", "N+1", "N+2", "N+3"])
mrows = ["Résultat net acquéreur (constant — simplification)", "+ Résultat net de la cible (N, constant)", "− Coût de financement après impôt (dette + cash)",
         "− Amortissement des incorporels après impôt", "− Frais de transaction après impôt (N+1)", "+ Synergies après impôt (phasées)",
         "Résultat net pro forma", "Nombre d'actions pro forma", "BPA pro forma (devise)", "BPA acquéreur seul (devise)", "Accrétion (+) / dilution (−) du BPA",
         "Synergies avant impôt nécessaires pour BPA neutre"]
for j, t in enumerate(mrows):
    lab_(ws, r + 2 + j, t, bold=j in (6, 8, 10))
b0 = r + 2
for j, c in enumerate("BCD"):
    cell_(ws, f"{c}{b0}", f"={MA_['rn']}", NUM)
    cell_(ws, f"{c}{b0 + 1}", f"={G('rn', 'D')}", NUM, "link")
    cell_(ws, f"{c}{b0 + 2}", f"=-($B${MR['debt']}*{MA_['rd']}+$B${MR['cash']}*{MA_['rc']})*(1-{HYR('tax')})", NUM)
    cell_(ws, f"{c}{b0 + 3}", f"=-MAX($B${MR['gw']},0)*{PPA_P}/{PPA_Y}*(1-{HYR('tax')})", NUM)
    cell_(ws, f"{c}{b0 + 4}", f"=-$B${MR['eqv']}*{MA_['fees']}*(1-{HYR('tax')})" if j == 0 else 0, NUM)
    cell_(ws, f"{c}{b0 + 5}", f"={MA_['syn']}*{c}${PH}*(1-{HYR('tax')})", NUM)
    cell_(ws, f"{c}{b0 + 6}", f"=SUM({c}{b0}:{c}{b0 + 5})", NUM, bold=True)
    cell_(ws, f"{c}{b0 + 7}", f"={MA_['act']}+$B${MR['new']}", '#,##0')
    cell_(ws, f"{c}{b0 + 8}", f"={c}{b0 + 6}*{SABS('mult')}/{c}{b0 + 7}", '#,##0.00', bold=True)
    cell_(ws, f"{c}{b0 + 9}", f"={MA_['rn']}*{SABS('mult')}/{MA_['act']}", '#,##0.00')
    cell_(ws, f"{c}{b0 + 10}", f'=IFERROR({c}{b0 + 8}/{c}{b0 + 9}-1,"n/d")', '+0.0%;-0.0%;"-"', bold=True)
    cell_(ws, f"{c}{b0 + 11}", f'=IFERROR(MAX(0,({c}{b0 + 9}*{c}{b0 + 7}/{SABS("mult")}-({c}{b0 + 6}-{c}{b0 + 5}))/(1-{HYR("tax")})),"n/d")', NUM)
MA_R = dict(acc=b0 + 10, bpa=b0 + 8)
lab_(ws, b0 + 13, "Le BPA est un indicateur comptable : une opération accrétive peut détruire de la valeur si le prix excède la valeur intrinsèque (cf. DCF).", note=True)

# =====================================================================
# DCF 10 ans (convergence) + DCF inversé + Transactions comparables
# =====================================================================
FADE = HYD.get("fade", {}) if not ARGS.blank else {}
ws = ws_d10
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 58
for col in "BCDEFGHIJKL":
    ws.column_dimensions[col].width = 11.5
ws["A1"].value = "DCF À 10 ANS — convergence progressive de la croissance et de la marge (scénario Base)"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "Années 1-5 = scénario Base de l'onglet Hypotheses ; années 6-10 : la croissance converge linéairement vers g, la marge vers la cible saisie. "
            "Adapté aux profils d'hypercroissance, où un horizon de 5 ans concentre la valeur dans la valeur terminale.", note=True)
sec_(ws, 4, "HYPOTHÈSES DE LA PHASE DE CONVERGENCE (années 6 à 10)", 12)
lab_(ws, 5, "Marge d'EBITDA cible en année 10 (vide = marge de l'année 5)")
cell_(ws, "B5", FADE.get("marge_lt"), PCT, "inp")
lab_(ws, 6, "Multiple VE/EBITDA de sortie en année 10 (vide = multiple de l'onglet Hypotheses)")
cell_(ws, "B6", FADE.get("mult_lt"), MULT, "inp")
lab_(ws, 7, "Croissance en année 10 = croissance perpétuelle (onglet Hypotheses)")
cell_(ws, "B7", f"={GG}", PCT, "link")
lab_(ws, 8, "Marge d'EBITDA effective en année 10")
cell_(ws, "B8", "=IF(ISNUMBER(B5),B5,G15)", PCT)
lab_(ws, 9, "Multiple de sortie effectif")
cell_(ws, "B9", f"=IF(ISNUMBER(B6),B6,{HYR('mult')})", MULT)

D10 = dict(idx=12, g=13, ca=14, m=15, ebitda=16, da=17, ebit=18, tax=19, nopat=20, capex=21, bfr=22, dbfr=23, lease=24, fcff=25, fac=26, pv=27, per=28)
D10C = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
hdr_(ws, 11, [""] * 12)
ws["B11"].value = f"=Saisie!D{SA.r['annee']}"
for i, c in enumerate(D10C):
    ws[f"{c}11"].value = f'=IFERROR(VALUE(Saisie!$D${SA.r["annee"]})+{i + 1},"N+{i + 1}")'
for k, t in [("idx", "Année"), ("g", "Croissance du chiffre d'affaires"), ("ca", "Chiffre d'affaires"), ("m", "Marge d'EBITDA"),
             ("ebitda", "EBITDA"), ("da", "− Dotations aux amortissements"), ("ebit", "EBIT"), ("tax", "− Impôt sur l'EBIT (taux normatif)"),
             ("nopat", "NOPAT"), ("capex", "− Investissements (capex)"), ("bfr", "BFR total"), ("dbfr", "− Variation du BFR"),
             ("lease", "− Loyers IFRS 16 (si traitement = 1)"), ("fcff", "FCFF (N+1 : part restant à courir si bilan intermédiaire)"), ("fac", "Facteur d'actualisation"), ("pv", "FCFF actualisé"), ("per", "Période d'actualisation (années)")]:
    lab_(ws, D10[k], t, bold=k in ("ebitda", "ebit", "nopat", "fcff"))
cell_(ws, f"B{D10['idx']}", 0, "0")
cell_(ws, f"B{D10['ca']}", f"={G('ca', 'D')}", NUM, "link")
cell_(ws, f"B{D10['m']}", f"={G('m_ebe', 'D')}", PCT, "link")
cell_(ws, f"B{D10['ebitda']}", f"={G('ebe', 'D')}", NUM, "link", bold=True)
cell_(ws, f"B{D10['bfr']}", f"={BR('bfr', 'D')}", NUM, "link")
prev = "B"
for i, c in enumerate(D10C):
    n = i + 1
    hc = HCOLS[min(n, 5) - 1]
    R_ = D10
    cell_(ws, f"{c}{R_['idx']}", n, "0")
    if n <= 5:
        cell_(ws, f"{c}{R_['g']}", f"=Hypotheses!{hc}{H['base_g']}", PCT, "link")
        cell_(ws, f"{c}{R_['m']}", f"=Hypotheses!{hc}{H['base_m']}", PCT, "link")
    else:
        cell_(ws, f"{c}{R_['g']}", f"=$G{R_['g']}+($B$7-$G{R_['g']})*({c}{R_['idx']}-5)/5", PCT)
        cell_(ws, f"{c}{R_['m']}", f"=$G{R_['m']}+($B$8-$G{R_['m']})*({c}{R_['idx']}-5)/5", PCT)
    cell_(ws, f"{c}{R_['ca']}", f"={prev}{R_['ca']}*(1+{c}{R_['g']})", NUM)
    cell_(ws, f"{c}{R_['ebitda']}", f"={c}{R_['ca']}*{c}{R_['m']}", NUM, bold=True)
    cell_(ws, f"{c}{R_['da']}", f"={c}{R_['ca']}*Hypotheses!${hc}${H['da']}", NUM)
    cell_(ws, f"{c}{R_['ebit']}", f"={c}{R_['ebitda']}-{c}{R_['da']}", NUM, bold=True)
    cell_(ws, f"{c}{R_['tax']}", f"=MAX({c}{R_['ebit']},0)*{HYR('tax')}", NUM)
    cell_(ws, f"{c}{R_['nopat']}", f"={c}{R_['ebit']}-{c}{R_['tax']}", NUM, bold=True)
    cell_(ws, f"{c}{R_['capex']}", f"={c}{R_['ca']}*Hypotheses!${hc}${H['capex']}", NUM)
    cell_(ws, f"{c}{R_['bfr']}", f"={c}{R_['ca']}*Hypotheses!${hc}${H['bfr']}", NUM)
    cell_(ws, f"{c}{R_['dbfr']}", f"={c}{R_['bfr']}-{prev}{R_['bfr']}", NUM)
    cell_(ws, f"{c}{R_['lease']}", f"={c}{R_['ca']}*Hypotheses!$B${H['lease']}*Hypotheses!$B${H['ifrs']}", NUM)
    _sc = f"*(1-{STUB})" if i == 0 else ""
    cell_(ws, f"{c}{R_['fcff']}", f"=({c}{R_['nopat']}+{c}{R_['da']}-{c}{R_['capex']}-{c}{R_['dbfr']}-{c}{R_['lease']}){_sc}", NUM, bold=True)
    cell_(ws, f"{c}{R_['per']}", (f"=(1-{STUB})*(1-0.5*{MID})+0.5*{MID}" if i == 0 else f"={n}-{STUB}"), "0.00")
    ws[f"{c}{R_['fcff']}"].fill = fill_tot
    cell_(ws, f"{c}{R_['fac']}", f"=1/(1+{WACC})^({c}{R_['per']}-0.5*{MID})", "0.0000")
    cell_(ws, f"{c}{R_['pv']}", f"={c}{R_['fcff']}*{c}{R_['fac']}", NUM)
    prev = c
ws[f"A{D10['fcff']}"].fill = fill_tot
ws[f"B{D10['fcff']}"].fill = fill_tot

sec_(ws, 29, "VALORISATION", 12)
S10 = dict(sum=30, tvg=31, pvtvg=32, tvm=33, pvtvm=34, evg=35, evm=36, ev=37, nd=38, adj=39, eq=40, ps=41, up=42, tvpg=43, tvpm=44, vs5=45, evebitda=46, cagr=47)
for k, t in [("sum", "Somme des FCFF actualisés (10 ans)"), ("tvg", "Valeur terminale — Gordon-Shapiro"), ("pvtvg", "Valeur actuelle de la VT (Gordon)"),
             ("tvm", "Valeur terminale — multiple de sortie"), ("pvtvm", "Valeur actuelle de la VT (multiple)"),
             ("evg", "Valeur d'entreprise — Gordon-Shapiro"), ("evm", "Valeur d'entreprise — multiple de sortie"),
             ("ev", "VALEUR D'ENTREPRISE RETENUE (pondérée)"), ("nd", "− Dette financière nette"), ("adj", "− Minoritaires + actifs non opérationnels"),
             ("eq", "VALEUR DES CAPITAUX PROPRES"), ("ps", "Valeur par action (devise)"), ("up", "Potentiel vs cours actuel"),
             ("tvpg", "Poids de la VT dans la VE (Gordon)"), ("tvpm", "Poids de la VT dans la VE (multiple)"),
             ("vs5", "Écart vs DCF 5 ans (scénario Base, valeur par action)"), ("evebitda", "VE / EBITDA N implicite"),
             ("cagr", "Croissance annuelle moyenne du CA sur 10 ans")]:
    lab_(ws, S10[k], t, bold=k in ("ev", "eq", "ps"))
sv = lambda k: f"B{S10[k]}"
cell_(ws, sv("sum"), f"=SUM(C{D10['pv']}:L{D10['pv']})", NUM)
cell_(ws, sv("tvg"), f"=L{D10['fcff']}*(1+{GG})/({WACC}-{GG})", NUM)
cell_(ws, sv("pvtvg"), f"={sv('tvg')}/(1+{WACC})^(10-{STUB}-0.5*{MID})", NUM)
cell_(ws, sv("tvm"), f"=(L{D10['ebitda']}-L{D10['lease']})*B9", NUM)
cell_(ws, sv("pvtvm"), f"={sv('tvm')}/(1+{WACC})^(10-{STUB})", NUM)
cell_(ws, sv("evg"), f"={sv('sum')}+{sv('pvtvg')}", NUM)
cell_(ws, sv("evm"), f"={sv('sum')}+{sv('pvtvm')}", NUM)
cell_(ws, sv("ev"), f"={HYR('wg')}*{sv('evg')}+(1-{HYR('wg')})*{sv('evm')}", NUM, bold=True)
cell_(ws, sv("nd"), f"=DCF!B{base['s_nd']}", NUM, "link")
cell_(ws, sv("adj"), f"=DCF!B{base['s_adj']}", NUM, "link")
cell_(ws, sv("eq"), f"={sv('ev')}-{sv('nd')}-{sv('adj')}", NUM, bold=True)
cell_(ws, sv("ps"), f'=IF({SABS("actions")}>0,{sv("eq")}*{SABS("mult")}/{SABS("actions")},"n/d")', '#,##0.00', bold=True)
cell_(ws, sv("up"), f'=IF(AND(ISNUMBER({sv("ps")}),{SABS("cours")}>0),{sv("ps")}/{SABS("cours")}-1,"n/d")', PCT)
cell_(ws, sv("tvpg"), f'=IFERROR({sv("pvtvg")}/{sv("evg")},"n/d")', PCT)
cell_(ws, sv("tvpm"), f'=IFERROR({sv("pvtvm")}/{sv("evm")},"n/d")', PCT)
cell_(ws, sv("vs5"), f'=IFERROR({sv("ps")}/DCF!B{base["s_ps"]}-1,"n/d")', PCT)
cell_(ws, sv("evebitda"), f'=IFERROR({sv("ev")}/B{D10["ebitda"]},"n/d")', MULT)
cell_(ws, sv("cagr"), f'=IFERROR((L{D10["ca"]}/B{D10["ca"]})^(1/10)-1,"n/d")', PCT)

# ---- DCF inversé : taux d'actualisation qui égalise la valeur DCF 10 ans et le cours
sec_(ws, 49, "DCF INVERSÉ — rendement implicite au cours actuel (flux du scénario Base sur 10 ans)", 12)
IMP = dict(px=50, w=51, wacc=52, ecart=53)
GR0, GR1 = 57, 97
lab_(ws, IMP["px"], "Cours actuel")
cell_(ws, f"B{IMP['px']}", f"={SABS('cours')}", '#,##0.00', "link")
lab_(ws, IMP["w"], "Taux d'actualisation implicite (rendement annuel si le scénario Base se réalise)", bold=True)
lab_(ws, IMP["wacc"], "WACC retenu")
cell_(ws, f"B{IMP['wacc']}", f"={WACC}", PCT, "link")
lab_(ws, IMP["ecart"], "Écart : implicite − WACC (négatif = cours supérieur à la valeur Base)")
hdr_(ws, GR0 - 1, ["Taux d'actualisation", "Valeur par action"])
for j in range(GR1 - GR0 + 1):
    r = GR0 + j
    cell_(ws, f"A{r}", f"={GG}+0.01" if j == 0 else f"=A{r - 1}+0.005", PCT, align="left")
    w = f"$A{r}"
    flows = f"SUMPRODUCT($C${D10['fcff']}:$L${D10['fcff']}/(1+{w})^($C${D10['per']}:$L${D10['per']}-0.5*{MID}))"
    evg = f"{flows}+$L${D10['fcff']}*(1+{GG})/({w}-{GG})/(1+{w})^(10-{STUB}-0.5*{MID})"
    evm = f"{flows}+($L${D10['ebitda']}-$L${D10['lease']})*$B$9/(1+{w})^(10-{STUB})"
    cell_(ws, f"B{r}", f'=IFERROR((({HYR("wg")}*({evg})+(1-{HYR("wg")})*({evm}))-$B${S10["nd"]}-$B${S10["adj"]})*{SABS("mult")}/{SABS("actions")},"n/d")', '#,##0.00')
VR = f"$B${GR0}:$B${GR1}"
mt = f"MATCH($B${IMP['px']},{VR},-1)"
cell_(ws, f"B{IMP['w']}", (f'=IFERROR(INDEX($A${GR0}:$A${GR1},{mt})+(INDEX({VR},{mt})-$B${IMP["px"]})'
                           f'/(INDEX({VR},{mt})-INDEX({VR},{mt}+1))*0.005,"n/d")'), PCT, bold=True)
cell_(ws, f"B{IMP['ecart']}", f'=IFERROR(B{IMP["w"]}-B{IMP["wacc"]},"n/d")', PCT)
lab_(ws, GR0 + 42, "Lecture : si le scénario Base se réalise et que l'on achète au cours actuel, le rendement annuel attendu est le taux implicite. "
                   "S'il est inférieur au WACC, le marché anticipe mieux que le Base.", note=True)

# ---- Transactions comparables
ws = ws_tx
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDEFGHIJ", [44, 10, 13, 13, 13, 12, 10, 12, 10, 12]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "TRANSACTIONS COMPARABLES — valeur de contrôle (M&A)"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "À renseigner à partir de bases de transactions (Mergermarket, Capital IQ, communiqués, documents AMF/SEC). Les multiples de transaction incluent une prime de contrôle : "
            "ils ne s'appliquent pas à une valeur minoritaire cotée. La ligne 5 est un exemple fictif (non retenu).", note=True)
hdr_(ws, 4, ["Cible / acquéreur", "Date", "VE (M devise)", "CA (dernier exercice/LTM)", "EBITDA", "Prime payée", "Retenir (1/0)", "VE/EBITDA", "VE/CA", "Prime retenue"])
_tx = [] if ARGS.blank else D.get("transactions", [])
if _tx:
    ex_rows = [(t["nom"], t.get("date"), t.get("ev"), t.get("ca"), t.get("ebitda"), t.get("prime"), t.get("retenir", 1), t.get("source")) for t in _tx[:10]]
    ex_rows += [(None,) * 8] * (10 - len(ex_rows))
else:
    ex_rows = [("Exemple fictif — cible A / acquéreur B", "2024", 1000, 800, 125, 0.30, 0, "Exemple fictif")] + [(None,) * 8] * 9
ws.column_dimensions["K"].width = 70
cell_(ws, "K4", "Source et base de calcul", None, bold=True)
for i, row in enumerate(ex_rows):
    r = 5 + i
    for c, v, f_ in zip("ABCDEFGK", row, [None, "@", NUM, NUM, NUM, PCT, "0", None]):
        cell_(ws, f"{c}{r}", v, f_, "inp", align="left" if c in "AK" else "right")
    cell_(ws, f"H{r}", f'=IFERROR(IF(AND(G{r}=1,E{r}>0,C{r}>0),C{r}/E{r},""),"")', MULT)
    cell_(ws, f"I{r}", f'=IFERROR(IF(AND(G{r}=1,D{r}>0,C{r}>0),C{r}/D{r},""),"")', MULT)
    cell_(ws, f"J{r}", f'=IFERROR(IF(AND(G{r}=1,ISNUMBER(F{r})),F{r},""),"")', PCT)
sec_(ws, 16, "STATISTIQUES (transactions retenues)", 10)
TXR = dict(n=17, med=18, q1=19, q3=20, prem=21, evs=22, hdr=24, ev=25, eq=26, ps=27, vs=28)
lab_(ws, TXR["n"], "Nombre de transactions retenues")
cell_(ws, f"B{TXR['n']}", "=COUNTIF(G5:G14,1)", "0")
lab_(ws, TXR["n"], "Nombre de transactions retenues (au moins 3 données exploitables requises par statistique)")
for k, t, fx, f_ in [("med", "VE/EBITDA — médiane", 'MEDIAN(H5:H14)', MULT), ("q1", "VE/EBITDA — 1er quartile", 'QUARTILE(H5:H14,1)', MULT),
                     ("q3", "VE/EBITDA — 3e quartile", 'QUARTILE(H5:H14,3)', MULT), ("prem", "Prime de contrôle — médiane", 'MEDIAN(J5:J14)', PCT),
                     ("evs", "VE/CA — médiane", 'MEDIAN(I5:I14)', MULT)]:
    lab_(ws, TXR[k], t)
    _rg = {"med": "H", "q1": "H", "q3": "H", "prem": "J", "evs": "I"}[k]
    cell_(ws, f"B{TXR[k]}", f'=IFERROR(IF(COUNT({_rg}5:{_rg}14)>=3,{fx},"n/d"),"n/d")', f_)
sec_(ws, 23, "VALEUR DE CONTRÔLE IMPLICITE (appliquée à l'EBITDA N)", 10)
hdr_(ws, TXR["hdr"], ["", "Q1", "Médiane", "Q3"])
lab_(ws, TXR["ev"], "Valeur d'entreprise implicite")
lab_(ws, TXR["eq"], "Valeur des capitaux propres")
lab_(ws, TXR["ps"], "Valeur par action (contrôle)", bold=True)
lab_(ws, TXR["vs"], "Écart vs cours actuel (prime implicite)")
for c, k in zip("BCD", ["q1", "med", "q3"]):
    cell_(ws, f"{c}{TXR['ev']}", f'=IFERROR($B${TXR[k]}*{G("ebe", "D")},"n/d")', NUM)
    cell_(ws, f"{c}{TXR['eq']}", f'=IFERROR({c}{TXR["ev"]}-DCF!$B${base["s_nd"]}-DCF!$B${base["s_adj"]},"n/d")', NUM)
    cell_(ws, f"{c}{TXR['ps']}", f'=IFERROR(IF({SABS("actions")}>0,{c}{TXR["eq"]}*{SABS("mult")}/{SABS("actions")},"n/d"),"n/d")', '#,##0.00', bold=True)
    cell_(ws, f"{c}{TXR['vs']}", f'=IFERROR({c}{TXR["ps"]}/{SABS("cours")}-1,"n/d")', PCT)
lab_(ws, TXR["vs"] + 2, "Fourchette de contrôle : ne pas la comparer à un cours minoritaire sans retrancher la prime ; un panel de moins de 3 transactions n'est pas exploitable.", note=True)

# ---------------------------------------------------------------- Football field (Synthèse)
SY.section("FOOTBALL FIELD — valeur par action selon les méthodes (devise)", "E")
fh = SY.header(["Bas", "Haut", "Amplitude"])
ff = [("DCF — scénarios Bear à Bull", f"=MIN(B{SY.r['v_ps']}:D{SY.r['v_ps']})", f"=MAX(B{SY.r['v_ps']}:D{SY.r['v_ps']})"),
      ("Comparables VE/EBITDA — Q1 à Q3", f"=Comparables!B{CPV['ps']}", f"=Comparables!D{CPV['ps']}"),
      ("Comparables PER — Q1 à Q3", f"=Comparables!B{CPP['ps']}", f"=Comparables!D{CPP['ps']}"),
      ("DCF 10 ans (convergence), valeur centrale ± 10 %", f"=DCF_10ans!B{S10['ps']}*0.9", f"=DCF_10ans!B{S10['ps']}*1.1"),
      ("Transactions comparables — Q1 à Q3 (valeur de contrôle)", f"=Transactions!B{TXR['ps']}", f"=Transactions!D{TXR['ps']}"),
      ("LBO — prix max. selon TRI cibles (onglet LBO)", "=LBO!B47", "=LBO!D47"),
      ("Cours actuel", f"={SABS('cours')}", f"={SABS('cours')}")]
SY.r["ff_first"] = SY.cur
for t, lo, hi in ff:
    r = SY.cur
    SY.cur += 1
    ws_sy.cell(r, 1, t).font = font()
    for c, v in zip("BC", [lo, hi]):
        cell_(ws_sy, f"{c}{r}", f'=IFERROR(IF(ISNUMBER({v[1:]}),{v[1:]},"n/d"),"n/d")', '#,##0.00', "link")
    cell_(ws_sy, f"D{r}", f'=IFERROR(C{r}-B{r},"n/d")', '#,##0.00')
SY.r["ff_last"] = SY.cur - 1
SY.gap()
fch = BarChart()
fch.type = "bar"
fch.grouping = "stacked"
fch.overlap = 100
fch.title = "Football field — valeur par action"
fch.roundedCorners = False
fch.add_data(Reference(ws_sy, min_col=2, min_row=fh, max_row=SY.r["ff_last"]), titles_from_data=True)
fch.add_data(Reference(ws_sy, min_col=4, min_row=fh, max_row=SY.r["ff_last"]), titles_from_data=True)
fch.set_categories(Reference(ws_sy, min_col=1, min_row=SY.r["ff_first"], max_row=SY.r["ff_last"]))
fch.series[0].graphicalProperties.noFill = True
fch.series[0].graphicalProperties.line.noFill = True
fch.series[1].graphicalProperties.solidFill = "2F5597"
fch.x_axis.delete = False
fch.y_axis.delete = False
fch.legend = None
fch.height, fch.width = 8, 15
ws_sy.add_chart(fch, "G38")

SY.section("CE QUE LE MARCHÉ ANTICIPE — multiples implicites au cours actuel", "E")
mk = [("m_cap", "Capitalisation boursière (cours × actions)", f'=IFERROR({SABS("cours")}*{SABS("actions")}/{SABS("mult")},"n/d")', NUM),
      ("m_nd", "+ Dette financière nette (avec dettes locatives) + minoritaires − actifs non opérationnels", f"={BR('dn', 'D')}+{HYR('minor')}-{HYR('nonop')}", NUM),
      ("m_ev", "Valeur d'entreprise de marché", None, NUM),
      ("m_evs", "VE / chiffre d'affaires N", None, MULT), ("m_eve", "VE / EBITDA N", None, MULT),
      ("m_evea", "VE / EBITDA ajusté SBC N", None, MULT), ("m_per", "PER (capitalisation / résultat net N)", None, MULT),
      ("m_fcf", "Rendement du FCFF (FCFF N / VE)", None, PCT),
      ("m_dcf", "Valeur DCF pondérée / cours (prime ou décote du DCF)", None, PCT)]
MKR = {}
for k, t, fx, fm in mk:
    r = SY.cur
    SY.cur += 1
    MKR[k] = r
    ws_sy.cell(r, 1, t).font = font(k in ("m_ev",))
    if fx:
        cell_(ws_sy, f"B{r}", fx, fm, "link")
cell_(ws_sy, f"B{MKR['m_ev']}", f'=IFERROR(B{MKR["m_cap"]}+B{MKR["m_nd"]},"n/d")', NUM, bold=True)
cell_(ws_sy, f"B{MKR['m_evs']}", f'=IFERROR(B{MKR["m_ev"]}/{G("ca", "D")},"n/d")', MULT)
cell_(ws_sy, f"B{MKR['m_eve']}", f'=IFERROR(B{MKR["m_ev"]}/{G("ebe", "D")},"n/d")', MULT)
cell_(ws_sy, f"B{MKR['m_evea']}", f'=IFERROR(B{MKR["m_ev"]}/{G("ebe_aj", "D")},"n/d")', MULT)
cell_(ws_sy, f"B{MKR['m_per']}", f'=IFERROR(B{MKR["m_cap"]}/{G("rn", "D")},"n/d")', MULT)
cell_(ws_sy, f"B{MKR['m_fcf']}", f'=IFERROR({CFR("fcff", "D")}/B{MKR["m_ev"]},"n/d")', PCT)
cell_(ws_sy, f"B{MKR['m_dcf']}", f'=IFERROR(E{SY.r["v_ps"]}/{SABS("cours")}-1,"n/d")', PCT)
lab_(ws_sy, SY.cur, "Lecture : un multiple élevé suppose une croissance durablement forte ; comparer au DCF et à la dispersion des pairs.", note=True)
SY.gap(2)

# contrôles supplémentaires (hors total existant : informatifs)
SY.section("CONTRÔLES ET ALERTES COMPLÉMENTAIRES", "E")
extra = [("Panel de comparables : au moins 3 pairs retenus", f'=IF(Comparables!B{CPS["Nombre de pairs retenus"]}>=3,"OK","< 3 pairs")'),
         ("Bêta : mode 1-2-3 valide et bêta comparables disponible si mode ≠ 1", f'=IF(AND(OR(Hypotheses!C{H["beta"]}=1,Hypotheses!C{H["beta"]}=2,Hypotheses!C{H["beta"]}=3),IF(Hypotheses!C{H["beta"]}=1,TRUE,ISNUMBER(WACC_Pairs!B20))),"OK","À CORRIGER")'),
         ("MA : modes de paiement = 100 %", f"=MA!B{MR['px'] - 1}"),
         ("Rapprochement de trésorerie N (écart résiduel < 2 % du CA)", f'=IFERROR(IF(ABS({CFR("dt_ecart", "D")})<=0.02*{G("ca", "D")},"OK","Écart à expliquer"),"n/d")')]
for t, fx in extra:
    r = SY.cur
    SY.cur += 1
    ws_sy.cell(r, 1, t).font = font()
    cell_(ws_sy, f"B{r}", fx, align="center", bold=True)
    ws_sy.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'$B{r}="OK"'], fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")))
    ws_sy.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'$B{r}<>"OK"'], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))


# =====================================================================
# Repères historiques dans Hypotheses
# =====================================================================
bench = [
    (H["wd"], f"={BR('dfin', 'D')}/({BR('dfin', 'D')}+{BR('cp', 'D')})", PCT),
    (H["da"], f"={G('dot', 'D')}/{G('ca', 'D')}", PCT),
    (H["capex"], f"={CFR('capex_ca', 'D')}", PCT),
    (H["bfr"], f"={BR('bfr', 'D')}/{G('ca', 'D')}", PCT),
    (H["bear_g"], f"={RAR('g_ca', 'D')}", PCT), (H["base_g"], f"={RAR('g_ca', 'D')}", PCT), (H["bull_g"], f"={RAR('g_ca', 'D')}", PCT),
    (H["bear_m"], f"={G('m_ebe', 'D')}", PCT), (H["base_m"], f"={G('m_ebe', 'D')}", PCT), (H["bull_m"], f"={G('m_ebe', 'D')}", PCT),
]
for row, fx, fm in bench:
    c = ws_hy[f"G{row}"]
    c.value = f'=IFERROR({fx[1:]},"n/d")'
    c.number_format = fm
    c.font = font(color=GREEN)
    c.alignment = Alignment(horizontal="center")
ws_hy["G13"].comment = Comment("Structure actuelle en valeurs comptables : dette brute / (dette brute + capitaux propres).", "Outil")

# =====================================================================
# LISEZMOI
# =====================================================================
LM = Sh(ws_lm, {"A": 28, "B": 110})
LM.title("OUTIL D'ANALYSE FINANCIÈRE — mode d'emploi",
         "À partir du bilan et du compte de résultat : SIG, CAF, BFR, flux, ratios, scores, variations, DCF, rapport PDF.")
rows = [
    ("COMMENT L'UTILISER", None),
    ("1. Saisie", "Onglet Saisie : remplacer les chiffres d'exemple (cellules jaunes, police bleue) par ceux de la société sur 3 exercices. Charges en positif, valeurs nettes au bilan."),
    ("2. Hypotheses", "Onglet Hypotheses : régler le WACC, la croissance perpétuelle, le multiple de sortie et les 3 scénarios (Bear / Base / Bull). Les « repères » en colonne G donnent le niveau historique."),
    ("3. Contrôles", "Onglet Synthese : tous les contrôles doivent être « OK » (bilan équilibré, résultat cohérent, trésorerie nette…). Sinon corriger la saisie avant de lire les résultats."),
    ("4. Lire les résultats", "SIG_CAF, Bilan_BFR, Cashflow, Ratios, Scores, Variations, DCF se recalculent seuls (formules Excel, aucune macro)."),
    ("5. Rapport PDF", "Enregistrer le classeur puis lancer :  python analyse.py Outil_Analyse_Financiere.xlsx   →  génère le rapport PDF. Import automatique d'une balance ou d'un FEC : python import_balance.py (voir README)."),
    ("CODE COULEUR", None),
    ("Bleu sur fond jaune", "Donnée ou hypothèse à saisir."),
    ("Noir", "Formule de calcul."),
    ("Vert", "Lien direct vers un autre onglet."),
    ("CONVENTIONS ET LIMITES", None),
    ("CAF", "Calculée comme résultat net + dotations (méthode additive simplifiée ; ne retraite pas les cessions d'actifs ni les reprises)."),
    ("Capex", "Si non saisi : estimé par la variation des immobilisations nettes + dotations (surestime si cessions, sous-estime si dépréciations d'actifs)."),
    ("FCFF / NOPAT", "NOPAT = EBIT − impôt normatif ; FCFF = NOPAT + dotations − capex − variation du BFR. Le taux d'impôt normatif est celui des hypothèses."),
    ("DCF", "WACC par le CAPM ; valeur terminale par Gordon-Shapiro et par multiple de sortie, pondérées. La valeur dépend fortement des hypothèses : consulter les tables de sensibilité."),
    ("Ratios", "Seuils de lecture selon le profil choisi dans Saisie (Générique / Services / Industrie), modifiables dans l'onglet Seuils : ce sont des jugements d'analyste, pas des normes. DSO et DPO corrigés de la TVA saisie."),
    ("IFRS 16", "Hypotheses : traitement 1 = loyers en charge (FCFF net des loyers, dette nette hors dettes locatives, EBITDA après loyers) ; 0 = dettes locatives dans la dette. Saisir les dettes locatives et loyers décaissés dans le mémo de Saisie."),
    ("Comparables / WACC_Pairs", "Panel de pairs (8 lignes) : multiples VE/EBITDA, football field, bêtas désendettés/ré-endettés (Hamada). Qualité des données à vérifier : agrégateurs gratuits, dates hétérogènes."),
    ("LBO", "Capacité de paiement d'un sponsor (TRI cible → prix maximal) et sensibilité du TRI. Structure de dette simplifiée."),
    ("Sources", "Registre des hypothèses : source, date, confiance, justification ; statut « documentée / jugement / non sourcée / ancienne »."),
    ("Marche", "Cours daté, consensus (CA, EBITDA, cours cible) comparé au modèle. À renseigner depuis une source professionnelle."),
    ("Qualitatif", "Forces concurrentielles, registre des risques (probabilité × impact), catalyseurs ; profil de risque qualitatif."),
    ("Reco_Marche", "Note de marché : cours cible (moyenne pondérée des méthodes), appréciation indicative, profil risque/rendement, rendement implicite du cours, crédit synthétique."),
    ("DCF_10ans", "DCF à 10 ans : la croissance et la marge convergent progressivement vers les valeurs de long terme ; DCF inversé (rendement implicite au cours actuel)."),
    ("Transactions", "Transactions comparables (M&A) : multiples et prime de contrôle ; à renseigner. Un panel de moins de 3 transactions n'est pas exploité."),
    ("MA", "Accrétion / dilution du BPA si un acquéreur rachète la société (paiement mixte, synergies phasées, amortissement des incorporels)."),
    ("Scores", "Altman Z'' (sociétés non industrielles) et Piotroski F : outils de screening, pas de prévision certaine."),
    ("Avertissement", "Outil d'analyse et de formation. Ne constitue pas un conseil en investissement ni une certification des comptes."),
]
LM.cur = 4
for a, b in rows:
    r = LM.cur
    if b is None:
        LM.section(a, "B")
        continue
    ca_ = ws_lm.cell(r, 1, a)
    ca_.font = font(True)
    ca_.alignment = Alignment(vertical="top")
    cb = ws_lm.cell(r, 2, b)
    cb.font = font()
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    ws_lm.row_dimensions[r].height = 30 if len(b) > 110 else 16
    LM.cur += 1
ws_lm["A9"].fill = fill_inp  # placeholder (neutralisé ci-dessous)
ws_lm["A9"].fill = PatternFill(fill_type=None)

# =====================================================================
# Sources (registre des hypothèses), Marche (données datées + consensus), Qualitatif (Porter, risques, catalyseurs)
# =====================================================================
SRC = {} if ARGS.blank else D.get("sources", {})
MKD = {} if ARGS.blank else D.get("marche", {})
QLD = {} if ARGS.blank else D.get("qualitatif", {})
_d = lambda s: _dt.date.fromisoformat(s) if s else None
DATEF = "dd/mm/yyyy"

# ---- Marche
ws = ws_mk
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDEFG", [52, 15, 15, 15, 15, 13, 13]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "DONNÉES DE MARCHÉ DATÉES ET CONSENSUS"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "Toute donnée de marché doit être datée et sourcée. Le consensus (analystes) se saisit depuis une source professionnelle (Bloomberg, FactSet, LSEG, CapIQ) ou à défaut une source publique datée.", note=True)
sec_(ws, 3, "COURS ET DATES", 7)
MKR2 = dict(ref=4, px=5, pxd=6, age=7, st=8, src=9, c_hdr=12, c_ca=13, c_ebitda=14, c_tp=17, c_n=18, c_mod=19, c_gap=20)
lab_(ws, MKR2["ref"], "Date de référence de l'analyse")
cell_(ws, f"B{MKR2['ref']}", _d(MKD.get("date_ref")), DATEF, "inp")
lab_(ws, MKR2["px"], "Cours actuel (onglet Saisie)")
cell_(ws, f"B{MKR2['px']}", f"={SABS('cours')}", '#,##0.00', "link")
lab_(ws, MKR2["pxd"], "Date du cours")
cell_(ws, f"B{MKR2['pxd']}", _d(MKD.get("cours_date")), DATEF, "inp")
lab_(ws, MKR2["age"], "Ancienneté du cours (jours)")
cell_(ws, f"B{MKR2['age']}", f'=IF(AND(ISNUMBER(B{MKR2["ref"]}),ISNUMBER(B{MKR2["pxd"]})),B{MKR2["ref"]}-B{MKR2["pxd"]},"n/d")', "0")
lab_(ws, MKR2["st"], "Statut (cours de moins de 5 jours)")
cell_(ws, f"B{MKR2['st']}", f'=IF(ISNUMBER(B{MKR2["age"]}),IF(B{MKR2["age"]}<=5,"OK","Périmé"),"n/d")', None, bold=True, align="center")
lab_(ws, MKR2["src"], "Source du cours")
cell_(ws, f"B{MKR2['src']}", MKD.get("cours_source"), None, "inp", align="left")
sec_(ws, 11, "MODÈLE (SCÉNARIO BASE) CONTRE CONSENSUS", 7)
hdr_(ws, MKR2["c_hdr"], ["Indicateur", "Consensus N+1", "Consensus N+2", "Modèle N+1", "Modèle N+2", "Écart N+1", "Écart N+2"])
for k, t, key in [("c_ca", "Chiffre d'affaires", "ca"), ("c_ebitda", "EBITDA", "ebitda")]:
    r = MKR2[k]
    lab_(ws, r, t)
    cons = MKD.get("consensus", {}).get(key, [None, None])
    for c, v in zip("BC", cons):
        cell_(ws, f"{c}{r}", v, NUM, "inp")
    cell_(ws, f"D{r}", f"=DCF!C{base[key]}", NUM, "link")
    cell_(ws, f"E{r}", f"=DCF!D{base[key]}", NUM, "link")
    cell_(ws, f"F{r}", f'=IFERROR(IF(ISNUMBER(B{r}),D{r}/B{r}-1,"n/d"),"n/d")', '+0.0%;-0.0%;"-"')
    cell_(ws, f"G{r}", f'=IFERROR(IF(ISNUMBER(C{r}),E{r}/C{r}-1,"n/d"),"n/d")', '+0.0%;-0.0%;"-"')
lab_(ws, 15, "Un écart supérieur à 10 % sur N+1 doit être expliqué (hypothèse différenciée du marché ou donnée à corriger).", note=True)
sec_(ws, 16, "COURS CIBLE DES ANALYSTES", 7)
lab_(ws, MKR2["c_tp"], "Cours cible moyen du consensus (devise)")
cell_(ws, f"B{MKR2['c_tp']}", MKD.get("consensus_tp"), '#,##0.00', "inp")
lab_(ws, MKR2["c_n"], "Nombre d'analystes")
cell_(ws, f"B{MKR2['c_n']}", MKD.get("consensus_n"), "0", "inp")
lab_(ws, MKR2["c_mod"], "Cours cible du modèle (onglet Reco_Marche)")
cell_(ws, f"B{MKR2['c_mod']}", "=Reco_Marche!B10", '#,##0.00', "link")
lab_(ws, MKR2["c_gap"], "Écart modèle / consensus")
cell_(ws, f"B{MKR2['c_gap']}", f'=IFERROR(IF(ISNUMBER(B{MKR2["c_tp"]}),B{MKR2["c_mod"]}/B{MKR2["c_tp"]}-1,"n/d"),"n/d")', '+0.0%;-0.0%;"-"')
sec_(ws, 24, "DÉCOTE JURIDIQUE / RÉGLEMENTAIRE (risque non modélisé dans les flux)", 7)
lab_(ws, 25, "Décote appliquée au cours cible (% de la valeur)")
cell_(ws, "B25", MKD.get("decote", 0), PCT, "inp")
lab_(ws, 26, "Justification et source de la décote")
cell_(ws, "B26", MKD.get("decote_source", "Aucune décote appliquée (0 %)"), None, "inp", align="left")
lab_(ws, 27, "La décote est un jugement, non calibré statistiquement : tester 0 % à 30 % (onglet Pont_Consensus). Elle ne remplace pas un scénario de litige chiffré.", note=True)
lab_(ws, 22, "Source et date du consensus :", bold=True)
cell_(ws, "B22", MKD.get("consensus_source"), None, "inp", align="left")

# ---- Sources
ws = ws_so
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDEFGH", [46, 13, 46, 13, 12, 52, 11, 15]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "REGISTRE DES HYPOTHÈSES — source, date, confiance, justification"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "Une hypothèse est « documentée » si sa source est citée (hors « Jugement… »), datée et de moins de 180 jours à la date de référence (onglet Marche). "
            "Une hypothèse laissée au jugement doit le dire : c'est la transparence, pas un défaut.", note=True)
hdr_(ws, 4, ["Hypothèse", "Valeur retenue", "Source (ou « Jugement de l'auteur »)", "Date de la source", "Confiance", "Justification", "Âge (jours)", "Statut"])
srcrows = [("rf", "Taux sans risque", f"={HYR('rf')}", PCT), ("erp", "Prime de risque du marché", f"={HYR('erp')}", PCT),
           ("beta", "Bêta retenu", f"=Hypotheses!B{H['beta']}", "0.00"), ("kd", "Coût de la dette avant impôt", f"={HYR('kd')}", PCT),
           ("tax", "Taux d'impôt normatif", f"={HYR('tax')}", PCT), ("g", "Croissance perpétuelle", f"={GG}", PCT),
           ("mult", "Multiple de sortie (5 ans)", f"={HYR('mult')}", MULT),
           ("g1", "Croissance du CA, année 1 (Base)", f"=Hypotheses!B{H['base_g']}", PCT),
           ("m5", "Marge d'EBITDA, année 5 (Base)", f"=Hypotheses!F{H['base_m']}", PCT),
           ("mlt", "Marge d'EBITDA cible, année 10", "=DCF_10ans!B8", PCT), ("mult10", "Multiple de sortie, année 10", "=DCF_10ans!B9", MULT),
           ("cours", "Cours actuel", f"={SABS('cours')}", '#,##0.00'), ("actions", "Nombre d'actions", f"={SABS('actions')}", '#,##0'),
           ("nonop", "Actifs non opérationnels", f"={HYR('nonop')}", NUM), ("wd", "Poids de la dette cible", f"={HYR('wd')}", PCT)]
SO0 = 5
SO1 = SO0 + len(srcrows) - 1
for i, (k, t, fx, f_) in enumerate(srcrows):
    r = SO0 + i
    sd = SRC.get(k, {})
    lab_(ws, r, t)
    cell_(ws, f"B{r}", fx, f_, "link")
    cell_(ws, f"C{r}", sd.get("source"), None, "inp", align="left")
    cell_(ws, f"D{r}", _d(sd.get("date")), DATEF, "inp")
    cell_(ws, f"E{r}", sd.get("conf"), None, "inp", align="center")
    cell_(ws, f"F{r}", sd.get("note"), None, "inp", align="left")
    cell_(ws, f"G{r}", f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(Marche!$B${MKR2["ref"]})),Marche!$B${MKR2["ref"]}-D{r},"n/d")', "0")
    cell_(ws, f"H{r}", (f'=IF(LEN(C{r})=0,"Non sourcée",IF(LEFT(C{r},7)="Jugemen","Jugement",IF(NOT(ISNUMBER(D{r})),"Non datée",'
                        f'IF(NOT(ISNUMBER(G{r})),"Date réf. manquante",IF(G{r}<=180,"Documentée","Source ancienne")))))'), None, align="center")
lab_(ws, SO1 + 2, "Hypothèses documentées (part du total)", bold=True)
SOP = SO1 + 2
cell_(ws, f"B{SOP}", f'=COUNTIF(H{SO0}:H{SO1},"Documentée")/ROWS(H{SO0}:H{SO1})', PCT, bold=True)
lab_(ws, SOP + 1, "Hypothèses laissées au jugement de l'auteur")
cell_(ws, f"B{SOP + 1}", f'=COUNTIF(H{SO0}:H{SO1},"Jugement")/ROWS(H{SO0}:H{SO1})', PCT)
lab_(ws, SOP + 2, "Hypothèses sans source")
cell_(ws, f"B{SOP + 2}", f'=(COUNTIF(H{SO0}:H{SO1},"Non sourcée")+COUNTIF(H{SO0}:H{SO1},"Non datée")+COUNTIF(H{SO0}:H{SO1},"Source ancienne")+COUNTIF(H{SO0}:H{SO1},"Date réf. manquante"))/ROWS(H{SO0}:H{SO1})', PCT)
for stt, col_ in [("Documentée", "C6EFCE"), ("Jugement", "FFEB9C"), ("Non sourcée", "FFC7CE"), ("Non datée", "FFC7CE"), ("Source ancienne", "FFC7CE"), ("Date réf. manquante", "FFC7CE")]:
    ws.conditional_formatting.add(f"H{SO0}:H{SO1}", FormulaRule(formula=[f'$H{SO0}="{stt}"'], fill=PatternFill("solid", bgColor=col_, fgColor=col_)))

# ---- Qualitatif
ws = ws_ql
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDEFG", [46, 18, 12, 12, 10, 56, 40]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "ANALYSE QUALITATIVE — concurrence, risques, catalyseurs"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "Les notes sont des appréciations de l'analyste : chacune doit citer une source (rapport annuel, facteurs de risque, presse, données sectorielles). Échelle 1 à 5.", note=True)
sec_(ws, 4, "1. FORCES CONCURRENTIELLES (1 = favorable à la société, 5 = défavorable)", 7)
hdr_(ws, 5, ["Force", "Note (1-5)", "", "", "", "Commentaire", "Source"])
porter = QLD.get("porter", [])
for i, nm in enumerate(["Rivalité entre concurrents", "Menace des nouveaux entrants", "Pouvoir des fournisseurs", "Pouvoir des clients", "Menace des substituts"]):
    r = 6 + i
    p = porter[i] if i < len(porter) else {}
    lab_(ws, r, nm)
    cell_(ws, f"B{r}", p.get("note"), "0", "inp")
    cell_(ws, f"F{r}", p.get("commentaire"), None, "inp", align="left")
    cell_(ws, f"G{r}", p.get("source"), None, "inp", align="left")
sec_(ws, 12, "2. REGISTRE DES RISQUES (score = probabilité × impact, de 1 à 25)", 7)
hdr_(ws, 13, ["Risque", "Catégorie", "Proba. (1-5)", "Impact (1-5)", "Score", "Atténuation / commentaire", "Source"])
QR0, QR1 = 14, 23
risks = QLD.get("risques", [])
for i in range(QR1 - QR0 + 1):
    r = QR0 + i
    k = risks[i] if i < len(risks) else {}
    cell_(ws, f"A{r}", k.get("risque"), None, "inp", align="left")
    cell_(ws, f"B{r}", k.get("categorie"), None, "inp", align="left")
    cell_(ws, f"C{r}", k.get("proba"), "0", "inp")
    cell_(ws, f"D{r}", k.get("impact"), "0", "inp")
    cell_(ws, f"E{r}", f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*D{r},"")', "0", bold=True)
    cell_(ws, f"F{r}", k.get("note"), None, "inp", align="left")
    cell_(ws, f"G{r}", k.get("source"), None, "inp", align="left")
ws.conditional_formatting.add(f"E{QR0}:E{QR1}", FormulaRule(formula=[f'AND(ISNUMBER(E{QR0}),E{QR0}>=15)'], fill=PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")))
ws.conditional_formatting.add(f"E{QR0}:E{QR1}", FormulaRule(formula=[f'AND(ISNUMBER(E{QR0}),E{QR0}>=8,E{QR0}<15)'], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
sec_(ws, 25, "3. CATALYSEURS", 7)
hdr_(ws, 26, ["Événement", "Date prévue", "Sens (+ / −)", "", "", "Impact attendu", "Source"])
cats = QLD.get("catalyseurs", [])
for i in range(6):
    r = 27 + i
    k = cats[i] if i < len(cats) else {}
    cell_(ws, f"A{r}", k.get("evenement"), None, "inp", align="left")
    cell_(ws, f"B{r}", _d(k.get("date")), DATEF, "inp")
    cell_(ws, f"C{r}", k.get("sens"), None, "inp", align="center")
    cell_(ws, f"F{r}", k.get("impact"), None, "inp", align="left")
    cell_(ws, f"G{r}", k.get("source"), None, "inp", align="left")
sec_(ws, 34, "4. SYNTHÈSE QUALITATIVE", 7)
QS = dict(n=35, mx=36, avg=37, porter=38, thr=39, prof=40)
lab_(ws, QS["n"], "Nombre de risques renseignés")
cell_(ws, f"B{QS['n']}", f"=COUNT(E{QR0}:E{QR1})", "0")
lab_(ws, QS["mx"], "Score de risque maximal")
cell_(ws, f"B{QS['mx']}", f'=IF(B{QS["n"]}>0,MAX(E{QR0}:E{QR1}),"n/d")', "0")
lab_(ws, QS["avg"], "Score de risque moyen")
cell_(ws, f"B{QS['avg']}", f'=IF(B{QS["n"]}>0,AVERAGE(E{QR0}:E{QR1}),"n/d")', "0.0")
lab_(ws, QS["porter"], "Intensité concurrentielle moyenne (1 = favorable, 5 = défavorable)")
cell_(ws, f"B{QS['porter']}", '=IF(COUNT(B6:B10)>0,AVERAGE(B6:B10),"n/d")', "0.0")
lab_(ws, QS["thr"], "Seuil « risque élevé » : score maximal supérieur ou égal à")
cell_(ws, f"B{QS['thr']}", 16, "0", "inp")
lab_(ws, QS["prof"], "Profil de risque qualitatif", bold=True)
cell_(ws, f"B{QS['prof']}", f'=IF(B{QS["n"]}=0,"n/d",IF(B{QS["mx"]}>=B{QS["thr"]},"ÉLEVÉ",IF(B{QS["mx"]}>=0.6*B{QS["thr"]},"MODÉRÉ","FAIBLE")))', None, bold=True, align="center")

# =====================================================================
# Reco_Marche : cours cible, recommandation indicative, risque/rendement, DCF inversé, crédit
# =====================================================================
ws = ws_rc
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDE", [62, 16, 14, 14, 14]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "NOTE DE MARCHÉ — cours cible, risque / rendement, anticipations du marché, crédit"
ws["A1"].font = font(True, "1F3864", 14)
lab_(ws, 2, "Sorties indicatives fondées sur des règles mécaniques et sur les hypothèses saisies : elles ne constituent pas un conseil en investissement.", note=True)
sec_(ws, 4, "1. COURS CIBLE — moyenne pondérée des méthodes disponibles", 5)
meth = [("DCF 5 ans — pondéré par les probabilités des scénarios", f"=Synthese!E{SY.r['v_ps']}", 0.35),
        ("DCF 10 ans — convergence (scénario Base)", f"=DCF_10ans!B{S10['ps']}", 0.30),
        ("Comparables VE/EBITDA — médiane", f"=Comparables!C{CPV['ps']}", 0.20),
        ("Comparables PER — médiane", f"=Comparables!C{CPP['ps']}", 0.15)]
for i, (t, fx, wt) in enumerate(meth):
    r = 6 + i
    lab_(ws, r, t)
    cell_(ws, f"B{r}", f'=IFERROR(IF(ISNUMBER({fx[1:]}),{fx[1:]},"n/d"),"n/d")', '#,##0.00', "link")
    cell_(ws, f"C{r}", wt, PCT, "inp")
    cell_(ws, f"D{r}", f'=IF(AND(ISNUMBER(B{r}),G{r}<>"Écartée"),C{r},0)', PCT)
    cell_(ws, f"E{r}", f"=IF(ISNUMBER(B{r}),B{r}*D{r},0)", '#,##0.00')
    cell_(ws, f"F{r}", f'=IFERROR(IF(ISNUMBER(B{r}),B{r}/MEDIAN($B$6:$B$9)-1,"n/d"),"n/d")', '+0%;-0%;"-"')
    cell_(ws, f"G{r}", f'=IF(ISNUMBER(F{r}),IF(ABS(F{r})>$B$17,"Écartée","Retenue"),"n/d")', None, align="center")
hdr_(ws, 5, ["Méthode", "Valeur / action", "Poids saisi", "Poids effectif", "Contribution", "Écart médiane", "Statut"])
for col, w_ in zip("FG", [14, 12]):
    ws.column_dimensions[col].width = w_
lab_(ws, 17, "Seuil d'écart à la médiane des méthodes au-delà duquel une valeur est écartée de la moyenne")
cell_(ws, "B17", 0.40, PCT, "inp")
_fb = BarChart(); _fb.type = "bar"; _fb.title = "Valeur par action selon la méthode ($)"
_fb.add_data(Reference(ws, min_col=2, min_row=5, max_row=9), titles_from_data=True)
_fb.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
_fb.legend = None; _fb.roundedCorners = False; _fb.height = 7; _fb.width = 16
_fb.x_axis.delete = False; _fb.y_axis.delete = False
ws.add_chart(_fb, "I4")
RC = dict(tp=10, px=11, up=12, buy=13, sell=14, rating=15, nm=16, bear=19, base=20, bull=21, wtd=22, er=23, skew=24,
          impl=27, wacc=28, gap=29, read=30, nd=33, icr=34, crd=35, chk=38)
lab_(ws, RC["tp"], "COURS CIBLE (valeur par action)", bold=True)
cell_(ws, f"B{RC['tp']}", '=IFERROR(IF(SUM(D6:D9)>0,SUM(E6:E9)/SUM(D6:D9)*(1-Marche!$B$25),"n/d"),"n/d")', '#,##0.00', bold=True)
lab_(ws, RC["px"], "Cours actuel")
cell_(ws, f"B{RC['px']}", f"={SABS('cours')}", '#,##0.00', "link")
lab_(ws, RC["up"], "Potentiel vs cours actuel", bold=True)
cell_(ws, f"B{RC['up']}", f'=IFERROR(B{RC["tp"]}/B{RC["px"]}-1,"n/d")', '+0.0%;-0.0%;"-"', bold=True)
lab_(ws, RC["buy"], "Seuil « favorable » : potentiel supérieur ou égal à")
cell_(ws, f"B{RC['buy']}", 0.15, PCT, "inp")
lab_(ws, RC["sell"], "Seuil « défavorable » : potentiel inférieur ou égal à")
cell_(ws, f"B{RC['sell']}", -0.10, PCT, "inp")
lab_(ws, RC["rating"], "Appréciation indicative (règle mécanique sur le potentiel)", bold=True)
cell_(ws, f"B{RC['rating']}", (f'=IF(ISNUMBER(B{RC["up"]}),IF(B{RC["up"]}>=B{RC["buy"]},"FAVORABLE",IF(B{RC["up"]}<=B{RC["sell"]},"DÉFAVORABLE","NEUTRE")),"n/d")'), None, bold=True, align="center")
lab_(ws, RC["nm"], "Nombre de méthodes disponibles (sur 4)")
cell_(ws, f"B{RC['nm']}", "=COUNT(B6:B9)", "0")
sec_(ws, 18, "2. PROFIL RISQUE / RENDEMENT (scénarios DCF 5 ans)", 5)
for k, t, col in [("bear", "Valeur par action — Bear", "B"), ("base", "Valeur par action — Base", "C"), ("bull", "Valeur par action — Bull", "D")]:
    lab_(ws, RC[k], t)
    cell_(ws, f"B{RC[k]}", f"=Synthese!{col}{SY.r['v_ps']}", '#,##0.00', "link")
lab_(ws, RC["wtd"], "Valeur pondérée par les probabilités")
cell_(ws, f"B{RC['wtd']}", f"=Synthese!E{SY.r['v_ps']}", '#,##0.00', "link")
lab_(ws, RC["er"], "Rendement espéré (valeur pondérée / cours − 1)")
cell_(ws, f"B{RC['er']}", f'=IFERROR(B{RC["wtd"]}/B{RC["px"]}-1,"n/d")', '+0.0%;-0.0%;"-"')
lab_(ws, RC["skew"], "Ratio gain potentiel (Bull) / perte potentielle (Bear) vs cours")
cell_(ws, f"B{RC['skew']}", f'=IFERROR(IF(B{RC["px"]}<=B{RC["bear"]},"Cours ≤ Bear",(B{RC["bull"]}-B{RC["px"]})/(B{RC["px"]}-B{RC["bear"]})),"n/d")', '0.00"x"')
sec_(ws, 26, "3. CE QUE LE MARCHÉ ANTICIPE (DCF inversé, flux Base sur 10 ans)", 5)
lab_(ws, RC["impl"], "Rendement annuel implicite si le scénario Base se réalise")
cell_(ws, f"B{RC['impl']}", f"=DCF_10ans!B{IMP['w']}", PCT, "link")
lab_(ws, RC["wacc"], "WACC retenu")
cell_(ws, f"B{RC['wacc']}", f"={WACC}", PCT, "link")
lab_(ws, RC["gap"], "Écart (implicite − WACC)")
cell_(ws, f"B{RC['gap']}", f'=IFERROR(B{RC["impl"]}-B{RC["wacc"]},"n/d")', '+0.0%;-0.0%;"-"')
lab_(ws, RC["read"], "Lecture")
cell_(ws, f"B{RC['read']}", ('=IF(ISNUMBER(B29),IF(B29<-0.01,"Le cours intègre mieux que le Base : peu de marge de sécurité",'
                              'IF(B29>0.01,"Le cours rémunère au-delà du WACC si le Base se réalise","Cours cohérent avec le Base")),"n/d")'), None, align="left")
sec_(ws, 32, "4. QUALITÉ DE CRÉDIT (indicatif)", 5)
lab_(ws, RC["nd"], "Dette nette / EBITDA N (négatif = trésorerie nette)")
cell_(ws, f"B{RC['nd']}", f'=IFERROR({BR("dn", "D")}/{G("ebe", "D")},"n/d")', '0.0"x";-0.0"x"')
lab_(ws, RC["icr"], "Couverture des intérêts (EBIT / charges financières)")
cell_(ws, f"B{RC['icr']}", f'=IFERROR(IF(Saisie!D{SA.r["cfin"]}>0,{G("ebit", "D")}/Saisie!D{SA.r["cfin"]},"n/d"),"n/d")', '0.0"x"')
lab_(ws, RC["crd"], "Notation synthétique indicative (table de couverture ci-dessous)")
cell_(ws, f"B{RC['crd']}", f'=IFERROR(INDEX($B$71:$B$85,MATCH(B{RC["icr"]},$A$71:$A$85,1)),"n/d")', None, align="center")
sec_(ws, 37, "5. CONTRÔLES DE FIABILITÉ", 5)
chk = [("Au moins 3 méthodes de valeur disponibles", f'=IF(B{RC["nm"]}>=3,"OK","Insuffisant")'),
       ("Poids de la valeur terminale (DCF 10 ans, Gordon) ≤ 75 %", f'=IFERROR(IF(DCF_10ans!B{S10["tvpg"]}<=0.75,"OK","Élevé : forte sensibilité à g et au WACC"),"n/d")'),
       ("Dispersion des scénarios (Bull − Bear) / Base ≤ 100 %", f'=IFERROR(IF((B{RC["bull"]}-B{RC["bear"]})/B{RC["base"]}<=1,"OK","Très large : conviction faible"),"n/d")'),
       ("Panel de comparables : au moins 3 pairs", f'=IF(Comparables!B{CPS["Nombre de pairs retenus"]}>=3,"OK","< 3 pairs")'),
       ("Hypothèses documentées (source citée, datée, < 180 j) ≥ 60 %", f'=IFERROR(IF(Sources!B{SOP}>=0.6,"OK","Insuffisant : "&TEXT(Sources!B{SOP},"0%")),"n/d")'),
       ("Cours daté et de moins de 5 jours", f"=Marche!B{MKR2['st']}"),
       ("Profil de risque qualitatif non élevé", f'=IF(Qualitatif!B{QS["prof"]}="n/d","n/d",IF(Qualitatif!B{QS["prof"]}="ÉLEVÉ","Risque élevé","OK"))'),
       ("Cours cible dans la fourchette du football field (min − max)", f'=IFERROR(IF(AND(B{RC["tp"]}>=MIN(Synthese!B{SY.r["ff_first"]}:B{SY.r["ff_last"] - 1}),B{RC["tp"]}<=MAX(Synthese!C{SY.r["ff_first"]}:C{SY.r["ff_last"] - 1})),"OK","Hors fourchette"),"n/d")')]
for i, (t, fx) in enumerate(chk):
    r = RC["chk"] + i
    lab_(ws, r, t)
    cell_(ws, f"B{r}", fx, None, bold=True, align="center")
    ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'$B{r}="OK"'], fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")))
    ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'$B{r}<>"OK"'], fill=PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")))
hdr_(ws, 70, ["Couverture des intérêts ≥", "Notation indicative"])
for i, (th, rt) in enumerate([(-1000, "D"), (0.2, "C"), (0.65, "CC"), (0.8, "CCC"), (1.25, "B-"), (1.5, "B"), (1.75, "B+"), (2, "BB"), (2.25, "BB+"),
                              (2.5, "BBB"), (3, "A-"), (4.25, "A"), (5.5, "A+"), (6.5, "AA"), (8.5, "AAA")]):
    cell_(ws, f"A{71 + i}", th, '0.00"x"', "inp", align="left")
    cell_(ws, f"B{71 + i}", rt, None, "inp", align="center")
lab_(ws, 87, "Table de correspondance pédagogique (grandes entreprises non financières, type Damodaran) : à remplacer par la notation des agences. "
             "Sans charge financière, la notation est « n/d ».", note=True)
NCHK = len(chk)
sec_(ws, RC["chk"] + NCHK + 1, "6. CONVICTION (nombre de contrôles réussis)", 5)
RC["conv_n"], RC["conv_lvl"] = RC["chk"] + NCHK + 2, RC["chk"] + NCHK + 3
lab_(ws, RC["conv_n"], f"Contrôles réussis (sur {NCHK})")
cell_(ws, f"B{RC['conv_n']}", f'=COUNTIF(B{RC["chk"]}:B{RC["chk"] + NCHK - 1},"OK")', "0")
lab_(ws, RC["conv_lvl"], "Niveau de conviction indicatif (≥ 75 % élevée, ≥ 50 % moyenne, sinon faible)", bold=True)
cell_(ws, f"B{RC['conv_lvl']}", f'=IF(B{RC["conv_n"]}/{NCHK}>=0.75,"ÉLEVÉE",IF(B{RC["conv_n"]}/{NCHK}>=0.5,"MOYENNE","FAIBLE"))', None, bold=True, align="center")

# ---- Pont_Consensus
ws = ws_pc
ws.sheet_view.showGridLines = False
for col, w_ in zip("ABCDEFG", [58, 18, 18, 18, 14, 14, 14]):
    ws.column_dimensions[col].width = w_
ws["A1"].value = "PONT MODÈLE ↔ CONSENSUS ↔ COURS — ce que chaque prix implique"
ws["A1"].font = font(True, "1F3864", 14)
ws["A2"].value = "Un analyste doit pouvoir expliquer pourquoi sa cible diffère du consensus et du cours. Tout est calculé à partir des autres onglets."
ws["A2"].font = font(italic=True, color="595959")
sec_(ws, 4, "1. ÉCART DE COURS CIBLE", 7)
_wt = "SUM(Reco_Marche!E6:E9)/SUM(Reco_Marche!D6:D9)"  # méthodes retenues uniquement (D = 0 si écartée)
rows_ = [("Cours actuel", f"={SABS('cours')}", '#,##0.00'),
         ("Cours cible du modèle avant décote", f'=IFERROR({_wt},"n/d")', '#,##0.00'),
         ("Décote juridique / réglementaire (onglet Marche)", "=Marche!B25", PCT),
         ("Cours cible du modèle après décote", "=Reco_Marche!B10", '#,##0.00'),
         ("Cours cible moyen du consensus", "=Marche!B17", '#,##0.00'),
         ("Écart modèle / consensus (après décote)", '=IFERROR(B8/B9-1,"n/d")', '+0.0%;-0.0%;"-"'),
         ("Écart modèle / cours actuel (après décote)", '=IFERROR(B8/B5-1,"n/d")', '+0.0%;-0.0%;"-"'),
         ("Écart consensus / cours actuel", '=IFERROR(B9/B5-1,"n/d")', '+0.0%;-0.0%;"-"')]
for j, (t, fx, f_) in enumerate(rows_):
    lab_(ws, 5 + j, t)
    cell_(ws, f"B{5 + j}", fx, f_, "link" if j in (0, 3, 4) else None)
sec_(ws, 14, "2. CE QUE CHAQUE PRIX IMPLIQUE (même bilan, mêmes actions)", 7)
hdr_(ws, 15, ["", "Cours actuel", "Cible consensus", "Cible modèle"])
PX_ = {"B": "$B$5", "C": "$B$9", "D": "$B$8"}
lab_(ws, 16, "Prix par action (devise)")
lab_(ws, 17, "Capitalisation (M devise)")
lab_(ws, 18, "Valeur d'entreprise (M devise)")
lab_(ws, 19, "VE / EBITDA N+1 du modèle")
lab_(ws, 20, "VE / chiffre d'affaires N+1 du modèle")
lab_(ws, 21, "PER sur BPA N+1 du consensus")
for c, ref in PX_.items():
    cell_(ws, f"{c}16", f"={ref}", '#,##0.00')
    cell_(ws, f"{c}17", f'=IFERROR({c}16*{SABS("actions")}/{SABS("mult")},"n/d")', NUM)
    cell_(ws, f"{c}18", f'=IFERROR({c}17+{BR("dn", "D")}+{HYR("minor")}-{HYR("nonop")},"n/d")', NUM)
    cell_(ws, f"{c}19", f'=IFERROR({c}18/Marche!$D$14,"n/d")', MULT)
    cell_(ws, f"{c}20", f'=IFERROR({c}18/Marche!$D$13,"n/d")', MULT)
    cell_(ws, f"{c}21", f'=IFERROR(IF(ISNUMBER($B$22),{c}16/$B$22,"n/d"),"n/d")', MULT)
lab_(ws, 22, "BPA N+1 du consensus (devise, saisie)")
cell_(ws, "B22", MKD.get("bpa_consensus"), '#,##0.00', "inp")
lab_(ws, 23, "Source du BPA", None)
cell_(ws, "B23", MKD.get("bpa_source"), None, "inp", align="left")
sec_(ws, 25, "3. SENSIBILITÉ DE LA CIBLE À LA DÉCOTE JURIDIQUE", 7)
hdr_(ws, 26, ["Décote", "Cible du modèle", "vs cours actuel", "vs consensus"])
for j, h_ in enumerate([0, 0.10, 0.20, 0.30, 0.40]):
    r = 27 + j
    cell_(ws, f"A{r}", h_, PCT, "inp", align="left")
    cell_(ws, f"B{r}", f'=IFERROR($B$6*(1-A{r}),"n/d")', '#,##0.00')
    cell_(ws, f"C{r}", f'=IFERROR(B{r}/$B$5-1,"n/d")', '+0.0%;-0.0%;"-"')
    cell_(ws, f"D{r}", f'=IFERROR(B{r}/$B$9-1,"n/d")', '+0.0%;-0.0%;"-"')
lab_(ws, 33, "Décote qui ramène la cible du modèle au cours actuel :", bold=True)
cell_(ws, "B33", '=IFERROR(1-B5/B6,"n/d")', PCT, bold=True)
lab_(ws, 34, "Lecture :", bold=True)
cell_(ws, "B34", '=IFERROR("Le marché paie "&TEXT(B19,"0.0")&"x l\'EBITDA N+1 du modèle contre "&TEXT(D19,"0.0")&"x à la cible ; pour justifier le cours actuel il faudrait une décote de "&TEXT(B33,"0%")&" sur la valeur.","n/d")', None, align="left")
ws.freeze_panes = "B4"
sec_(ws, 36, "4. SENSIBILITÉ AU REPLI DES MULTIPLES DES PAIRS (multiples datés du relevé, secteur en baisse depuis)", 7)
hdr_(ws, 37, ["Variation du multiple VE/EBITDA des pairs", "Valeur comparables / action", "Cible du modèle", "vs cours actuel"])
for j, sh in enumerate([0, -0.10, -0.20, -0.30]):
    r = 38 + j
    cell_(ws, f"A{r}", sh, '+0%;-0%;0%', "inp", align="left")
    cell_(ws, f"B{r}", f'=IFERROR((Comparables!$C${CPV["ev"]}*(1+A{r})-Comparables!$C${CPV["nd"]}-Comparables!$C${CPV["adj"]})*{SABS("mult")}/{SABS("actions")},"n/d")', '#,##0.00')
    cell_(ws, f"C{r}", f'=IFERROR((SUM(Reco_Marche!$E$6:$E$9)-Reco_Marche!$E$8+Reco_Marche!$D$8*B{r})/SUM(Reco_Marche!$D$6:$D$9)*(1-Marche!$B$25),"n/d")', '#,##0.00')
    cell_(ws, f"D{r}", f'=IFERROR(C{r}/$B$5-1,"n/d")', '+0.0%;-0.0%;"-"')
lab_(ws, 42, "Les multiples du panel sont ceux du relevé (11/11 pour la plupart) ; les cours des pairs ont depuis varié. Cette table borne l'effet sans prétendre le mesurer.", note=True)
ws.freeze_panes = "B4"
wb.move_sheet("Reco_Marche", offset=-(wb.worksheets.index(ws_rc) - wb.worksheets.index(ws_sy) - 1))
ws_ix_extra = [("DCF10.ps", "DCF_10ans", S10["ps"]), ("DCF10.ev", "DCF_10ans", S10["ev"]), ("DCF10.tvpg", "DCF_10ans", S10["tvpg"]),
               ("DCF10.cagr", "DCF_10ans", S10["cagr"]), ("DCF10.vs5", "DCF_10ans", S10["vs5"]), ("DCF10.implied", "DCF_10ans", IMP["w"]),
               ("Pc.gap", "Pont_Consensus", 10), ("Tx.n", "Transactions", TXR["n"]), ("Src.pct", "Sources", SOP), ("Mk.age", "Marche", MKR2["age"]), ("Ql.prof", "Qualitatif", QS["prof"]), ("Ql.mx", "Qualitatif", QS["mx"]), ("Tx.ps", "Transactions", TXR["ps"])] + [(f"Reco.{k}", "Reco_Marche", v) for k, v in RC.items()]

# =====================================================================
# _Index (pour analyse.py)
# =====================================================================
ws_ix.append(["clé", "feuille", "ligne"])
for sh in (SA, SG, BF, CF, RA, SC, VA, DC, SY):
    for k, v in sh.r.items():
        ws_ix.append([f"{sh.n}.{k}", sh.n, v])
for k, v in H.items():
    ws_ix.append([f"Hypotheses.{k}", "Hypotheses", v])
for k, v in CPS.items():
    ws_ix.append([f"Comparables.stat.{k}", "Comparables", v])
for k, v in CPV.items():
    ws_ix.append([f"Comparables.val.{k}", "Comparables", v])
for k, v in CPP.items():
    ws_ix.append([f"Comparables.per.{k}", "Comparables", v])
for k, v in MKR.items():
    ws_ix.append([f"Synthese.{k}", "Synthese", v])
for k_, sh_, r_ in ws_ix_extra:
    ws_ix.append([k_, sh_, r_])
ws_ix.append(["Comparables.first", "Comparables", CP0])
ws_ix.append(["Comparables.last", "Comparables", CP1])
ws_ix.append(["MA.acc", "MA", MA_R["acc"]])
ws_ix.append(["MA.bpa", "MA", MA_R["bpa"]])
ws_ix.append(["MA.px", "MA", MR["px"]])
ws_ix.sheet_state = "hidden"

# police uniforme + mise en page impression
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = wb.worksheets.index(ws_sy)
wb.save(OUT)
print("OK", OUT)
