#!/usr/bin/env python3
"""Génère le rapport PDF d'analyse financière à partir du classeur Excel.

Usage :  python analyse.py Outil_Analyse_Financiere.xlsx [Rapport.pdf]

Le classeur doit avoir été enregistré par Excel (valeurs calculées en cache).
Si ce n'est pas le cas et que LibreOffice est installé, il est recalculé automatiquement.
Dépendances : openpyxl, matplotlib, reportlab.
"""
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

NAVY = colors.HexColor("#1F3864")
GREY = colors.HexColor("#595959")
LIGHT = colors.HexColor("#D9E2F3")
GREEN, AMBER, RED = "#2E7D32", "#ED9B00", "#C62828"

# ------------------------------------------------------------------ données
src = sys.argv[1] if len(sys.argv) > 1 else "Outil_Analyse_Financiere.xlsx"
out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0].replace("Outil_", "Rapport_") + ".pdf"


def load(path):
    wb = load_workbook(path, data_only=True)
    if wb["Synthese"]["B5"].value is None:  # pas de valeurs en cache
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            sys.exit("Ouvrez et enregistrez le classeur dans Excel avant de lancer ce script.")
        tmp = tempfile.mkdtemp()
        subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, path],
                       check=True, capture_output=True)
        wb = load_workbook(os.path.join(tmp, os.path.basename(path)), data_only=True)
    return wb


wb = load(src)
IDX = {r[0]: (r[1], r[2]) for r in wb["_Index"].iter_rows(min_row=2, values_only=True) if r[0]}


def V(key, col="D"):
    sh, row = IDX[key]
    return wb[sh][f"{col}{row}"].value


def num(x):
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else None


NA = "n/d"


def fcell(c):
    """Formate la valeur d'une cellule Excel selon son format d'affichage."""
    v, nf = c.value, c.number_format or ""
    if num(v) is None:
        return str(v) if v is not None else ""
    if "%" in nf:
        return fp(v)
    if "x" in nf:
        return fx(v)
    if nf.startswith("0.0;"):
        return fd(v, 1)
    if nf.startswith("#,##0"):
        return fn(v)
    return fd(v)


def fn(x, d=0):
    x = num(x)
    if x is None:
        return NA
    s = f"{abs(x):,.{d}f}".replace(",", " ").replace(".", ",")
    return f"({s})" if x < 0 else s


def fp(x, d=1):
    x = num(x)
    return NA if x is None else f"{x * 100:.{d}f}".replace(".", ",") + " %"


def fx(x, d=1):
    x = num(x)
    return NA if x is None else f"{x:.{d}f}".replace(".", ",") + "x"


def fd(x, d=2):
    x = num(x)
    return NA if x is None else f"{x:.{d}f}".replace(".", ",")


name = wb["Saisie"][f"B{IDX['Saisie.nom'][1]}"].value or "Société"
secteur = wb["Saisie"][f"B{IDX['Saisie.secteur'][1]}"].value or ""
unit = wb["Saisie"][f"B{IDX['Saisie.unite'][1]}"].value or ""
devise = wb["Saisie"][f"B{IDX['Saisie.devise'][1]}"].value or ""
cours = num(wb["Saisie"][f"B{IDX['Saisie.cours'][1]}"].value)
years = [str(V("Saisie.annee", c)) for c in "BCD"]
COLS = "BCD"

# ------------------------------------------------------------------ polices et styles
fdir = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
pdfmetrics.registerFont(TTFont("DV", os.path.join(fdir, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DVB", os.path.join(fdir, "DejaVuSans-Bold.ttf")))
pdfmetrics.registerFont(TTFont("DVI", os.path.join(fdir, "DejaVuSans-Oblique.ttf")))
pdfmetrics.registerFontFamily("DV", normal="DV", bold="DVB", italic="DVI", boldItalic="DVB")
plt.rcParams["font.family"] = "DejaVu Sans"

ss = getSampleStyleSheet()
BODY = ParagraphStyle("body", parent=ss["Normal"], fontName="DV", fontSize=9, leading=13, alignment=TA_JUSTIFY, spaceAfter=4)
SMALL = ParagraphStyle("small", parent=BODY, fontSize=7.5, leading=10, textColor=GREY, alignment=0)
H1 = ParagraphStyle("h1", parent=ss["Heading1"], fontName="DVB", fontSize=15, textColor=NAVY, spaceBefore=4, spaceAfter=8)
H2 = ParagraphStyle("h2", parent=ss["Heading2"], fontName="DVB", fontSize=10.5, textColor=NAVY, spaceBefore=8, spaceAfter=4)
BUL = ParagraphStyle("bul", parent=BODY, leftIndent=12, bulletIndent=2, alignment=0)
CELL = ParagraphStyle("cell", parent=BODY, fontSize=7.8, leading=10, alignment=0, spaceAfter=0)

tmpdir = tempfile.mkdtemp()


def bullets(items):
    return [Paragraph(t, BUL, bulletText="•") for t in items]


def table(data, widths, header=True, align_right_from=1, zebra=True, extra=None, font=7.8):
    rows = [[Paragraph(str(c), CELL) if (isinstance(c, str) and len(c) > 28) else c for c in r] for r in data]
    t = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    st = [("FONT", (0, 0), (-1, -1), "DV", font), ("FONT", (0, 0), (-1, 0), "DVB", font),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("BACKGROUND", (0, 0), (-1, 0), NAVY),
          ("ALIGN", (align_right_from, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
          ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    if zebra:
        for i in range(2, len(data), 2):
            st.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F5F7FB")))
    if extra:
        st += extra
    t.setStyle(TableStyle(st))
    return t


def savefig(fig, nm):
    p = os.path.join(tmpdir, nm)
    fig.savefig(p, dpi=170, bbox_inches="tight")
    plt.close(fig)
    return p


def style_ax(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=8)
    ax.grid(axis="y", alpha=0.25)


# ------------------------------------------------------------------ valeurs clés
ca = [V("SIG_CAF.ca", c) for c in COLS]
ebe = [V("SIG_CAF.ebe", c) for c in COLS]
ebit = [V("SIG_CAF.ebit", c) for c in COLS]
rn = [V("SIG_CAF.rn", c) for c in COLS]
caf = [V("SIG_CAF.caf", c) for c in COLS]
m_ebe = [V("SIG_CAF.m_ebe", c) for c in COLS]
m_rn = [V("SIG_CAF.m_rn", c) for c in COLS]
fcff = [V("Cashflow.fcff", c) for c in COLS]
dn = [V("Bilan_BFR.dn", c) for c in COLS]
cp = [V("Bilan_BFR.cp", c) for c in COLS]
R = lambda k, c="D": V(f"Ratios.{k}", c)
Z = [V("Scores.z", c) for c in COLS]
ZZ = [V("Scores.z_zone", c) for c in COLS]
PS = [V("Scores.pscore", c) for c in COLS]
ctrl = V("Synthese.ctrl_glob", "B")

# ------------------------------------------------------------------ graphiques
# 1. Performance
fig, ax = plt.subplots(1, 2, figsize=(9, 3.1))
w = 0.26
xs = range(3)
ax[0].bar([x - w for x in xs], ca, w, label="CA", color="#1F3864")
ax[0].bar(xs, ebe, w, label="EBITDA", color="#4F81BD")
ax[0].bar([x + w for x in xs], rn, w, label="Résultat net", color="#9DC3E6")
ax[0].set_xticks(list(xs))
ax[0].set_xticklabels(years)
ax[0].set_title(f"CA, EBITDA et résultat net ({unit})", fontsize=9, loc="left")
ax[0].legend(fontsize=7, frameon=False)
style_ax(ax[0])
for lab, ser, col in (("Marge d'EBITDA", m_ebe, "#4F81BD"), ("Marge nette", m_rn, "#9DC3E6"),
                      ("Marge d'EBIT", [V("SIG_CAF.m_ebit", c) for c in COLS], "#1F3864")):
    ax[1].plot(years, [100 * v if num(v) is not None else None for v in ser], marker="o", label=lab, color=col)
ax[1].set_title("Marges (% du CA)", fontsize=9, loc="left")
ax[1].legend(fontsize=7, frameon=False)
style_ax(ax[1])
img_perf = savefig(fig, "perf.png")

# 2. Structure du bilan
fig, ax = plt.subplots(1, 2, figsize=(9, 3.1))
parts_a = [("Actif immobilisé", [V("Bilan_BFR.immo", c) for c in COLS], "#1F3864"),
           ("Stocks", [V("Bilan_BFR.stocks", c) for c in COLS], "#4F81BD"),
           ("Clients + autres créances", [V("Bilan_BFR.clients", c) + V("Bilan_BFR.autres_cr", c) for c in COLS], "#9DC3E6"),
           ("Trésorerie", [V("Bilan_BFR.treso", c) for c in COLS], "#C5E0B4")]
bot = [0, 0, 0]
for lab, vals, col in parts_a:
    ax[0].bar(years, vals, bottom=bot, label=lab, color=col, width=0.55)
    bot = [b + v for b, v in zip(bot, vals)]
ax[0].set_title(f"Emplois — actif ({unit})", fontsize=9, loc="left")
ax[0].legend(fontsize=6.5, frameon=False, loc="upper left")
style_ax(ax[0])
parts_p = [("Capitaux propres", cp, "#1F3864"),
           ("Dettes financières", [V("Bilan_BFR.dfin", c) for c in COLS], "#C00000"),
           ("Provisions", [V("Bilan_BFR.prov", c) for c in COLS], "#F4B183"),
           ("Dettes d'exploitation et autres", [V("Bilan_BFR.fourn", c) + V("Bilan_BFR.fisc", c) + V("Bilan_BFR.autres_det", c) for c in COLS], "#BFBFBF")]
bot = [0, 0, 0]
for lab, vals, col in parts_p:
    ax[1].bar(years, vals, bottom=bot, label=lab, color=col, width=0.55)
    bot = [b + v for b, v in zip(bot, vals)]
ax[1].set_title(f"Ressources — passif ({unit})", fontsize=9, loc="left")
ax[1].legend(fontsize=6.5, frameon=False, loc="upper left")
style_ax(ax[1])
img_bilan = savefig(fig, "bilan.png")

# 3. Cascade de trésorerie (N)
fig, ax = plt.subplots(figsize=(9, 3.1))
steps = [("CAF", V("Cashflow.caf")), ("Δ BFR", V("Cashflow.d_bfr")), ("Capex", -V("Cashflow.capex")),
         ("Δ dettes", V("Cashflow.d_dette")), ("Dividendes", -V("Cashflow.div")), ("Autres", V("Cashflow.dt_ecart"))]
cum, xs_, lbl = 0, [], []
for i, (lab, v) in enumerate(steps):
    v = num(v) or 0
    ax.bar(i, v, bottom=cum if v >= 0 else cum + v, color=GREEN if v >= 0 else RED, width=0.55)
    ax.text(i, (cum + v if v >= 0 else cum) + 0.01 * max(abs(s[1] or 0) for s in steps), fn(v), ha="center", va="bottom", fontsize=7)
    cum += v
    lbl.append(lab)
ax.bar(len(steps), cum, color="#1F3864", width=0.55)
ax.text(len(steps), cum, fn(cum), ha="center", va="bottom", fontsize=7)
ax.set_xticks(range(len(steps) + 1))
ax.set_xticklabels(lbl + ["Δ trésorerie"], fontsize=8)
ax.set_title(f"Des résultats à la variation de trésorerie — exercice {years[2]} ({unit})", fontsize=9, loc="left")
style_ax(ax)
img_cash = savefig(fig, "cash.png")

# 4. Valorisation
tags = ["bear", "base", "bull"]
eq = [V(f"DCF.{t}.s_eq", "B") for t in tags]
ps = [V(f"DCF.{t}.s_ps", "B") for t in tags]
use_ps = all(num(p) is not None for p in ps)
vals_plot = ps if use_ps else eq
sens1 = [[wb["DCF"].cell(IDX["DCF.sens1_hdr"][1] + 1 + i, 3 + j).value for j in range(5)] for i in range(5)]
sens2 = [[wb["DCF"].cell(IDX["DCF.sens2_hdr"][1] + 1 + i, 3 + j).value for j in range(5)] for i in range(5)]
allv = [v for r_ in sens1 + sens2 for v in r_ if num(v) is not None]
fig, ax = plt.subplots(figsize=(9, 2.8))
rows_ff = [("Scénarios DCF (Bear → Bull)", min(vals_plot), max(vals_plot), vals_plot[1]),
           ("Sensibilité WACC × g (Base)", min(v for r_ in sens1 for v in r_), max(v for r_ in sens1 for v in r_), sens1[2][2]),
           ("Sensibilité WACC × multiple (Base)", min(v for r_ in sens2 for v in r_), max(v for r_ in sens2 for v in r_), sens2[2][2])]
for i, (lab, lo, hi, mid) in enumerate(rows_ff[::-1]):
    ax.barh(i, hi - lo, left=lo, color="#9DC3E6", height=0.45)
    ax.plot([mid, mid], [i - 0.28, i + 0.28], color="#1F3864", lw=2)
    ax.text(lo, i + 0.32, fd(lo, 1), fontsize=7, ha="center")
    ax.text(hi, i + 0.32, fd(hi, 1), fontsize=7, ha="center")
ax.set_yticks(range(3))
ax.set_yticklabels([r_[0] for r_ in rows_ff[::-1]], fontsize=8)
if cours:
    ax.axvline(cours, color=RED, ls="--", lw=1.2)
    ax.text(cours, 2.62, f"Cours {fd(cours, 2)}", color=RED, fontsize=7, ha="center")
ax.set_xlabel(f"Valeur par action ({devise})" if use_ps else f"Valeur des capitaux propres ({unit})", fontsize=8)
ax.set_title("« Football field » : fourchettes de valeur", fontsize=9, loc="left")
style_ax(ax)
ax.grid(axis="x", alpha=0.25)
ax.grid(axis="y", alpha=0)
img_ff = savefig(fig, "ff.png")

# ------------------------------------------------------------------ commentaires automatiques
def evol(a, b):
    a, b = num(a), num(b)
    return None if a in (None, 0) or b is None else b / a - 1


def trend_word(x, hi="hausse", lo="baisse"):
    return hi if x >= 0 else lo


g_ca = evol(ca[1], ca[2])
cagr = (ca[2] / ca[0]) ** 0.5 - 1 if ca[0] else None
perf_txt = [
    f"Le chiffre d'affaires atteint <b>{fn(ca[2])} {unit}</b> en {years[2]}, en {trend_word(g_ca)} de <b>{fp(abs(g_ca))}</b> "
    f"sur un an et de {fp(cagr)} par an en moyenne depuis {years[0]}.",
    f"L'EBITDA s'établit à <b>{fn(ebe[2])} {unit}</b> (marge {fp(m_ebe[2])}), contre {fp(m_ebe[0])} en {years[0]} : "
    f"{'amélioration' if m_ebe[2] >= m_ebe[0] else 'dégradation'} de "
    f"{abs(m_ebe[2] - m_ebe[0]) * 100:.1f} point(s) de marge".replace(".", ",") +
    (" ; sur le dernier exercice l'EBITDA progresse plus vite que le CA (lecture à confirmer : mix, éléments non récurrents)." if (evol(ebe[1], ebe[2]) or 0) > (g_ca or 0) else "."),
    f"Le résultat net ressort à <b>{fn(rn[2])} {unit}</b> (marge nette {fp(m_rn[2])}) ; la CAF s'élève à {fn(caf[2])} {unit}.",
    f"La rentabilité est de <b>{fp(R('roe'))}</b> sur capitaux propres (ROE) et de <b>{fp(R('roic'))}</b> sur capital investi après impôt (ROIC), "
    f"à comparer au WACC retenu de {fp(wb['Hypotheses'][f'B{IDX['Hypotheses.wacc'][1]}'].value, 2)}.",
]
roic, wacc = num(R("roic")), num(wb["Hypotheses"][f"B{IDX['Hypotheses.wacc'][1]}"].value)
if roic is not None and wacc is not None:
    perf_txt.append("Le ROIC dépasse le WACC : la croissance <b>crée de la valeur</b>." if roic > wacc
                    else "Le ROIC est inférieur au WACC : la croissance <b>détruit de la valeur</b> au rythme actuel.")

bs_txt = [
    f"La dette financière nette est de <b>{fn(dn[2])} {unit}</b>, soit <b>{fx(R('nd_ebitda'))}</b> l'EBITDA et "
    f"un gearing de {fd(R('gearing'))} ; les intérêts sont couverts {fx(R('couv_int'))} par l'EBIT.",
    f"L'autonomie financière est de {fp(R('autonomie'))} ; la liquidité générale de {fd(R('liq_g'))} et la liquidité immédiate de {fd(R('liq_i'))}.",
    f"Le BFR représente <b>{fd(V('Bilan_BFR.bfr_j'), 1)} jours</b> de chiffre d'affaires (FRNG : {fn(V('Bilan_BFR.frng'))} {unit}, "
    f"trésorerie nette : {fn(V('Bilan_BFR.tn'))} {unit}). Le cycle de conversion de trésorerie est de {fd(R('ccc'), 0)} jours "
    f"(clients {fd(R('dso'), 0)} j, stocks {fd(R('dio'), 0)} j, fournisseurs {fd(R('dpo'), 0)} j).",
]
conv = num(R("conv"))
cash_txt = [
    f"Le <b>FCFF</b> de {years[2]} est de <b>{fn(fcff[2])} {unit}</b> (FCF opérationnel : {fn(V('Cashflow.fcf_op'))} {unit}, conversion cash {fp(conv)} de l'EBITDA).",
    f"Les investissements atteignent {fn(V('Cashflow.capex'))} {unit} ({fp(R('capex_ca'))} du CA, {fx(R('capex_dot'))} les dotations) et la variation du BFR absorbe {fn(-V('Cashflow.d_bfr'))} {unit}.",
]
if conv is not None and conv < 0.25:
    cash_txt.append("La <b>faible conversion en cash</b> traduit une phase d'investissement et de croissance du BFR : "
                    "à surveiller si elle se prolonge, car les dividendes et le désendettement dépendent de ce flux.")
elif conv is not None and conv >= 0.5:
    cash_txt.append("La conversion en cash est solide : l'activité finance ses investissements et sa distribution.")
ecart = num(V("Cashflow.dt_ecart"))
if ecart is not None:
    cash_txt.append(f"Le rapprochement avec la variation de trésorerie du bilan laisse un écart de {fn(ecart)} {unit} "
                    "(capital, provisions, cessions, éléments hors exploitation).")

# points d'attention
alerts, strengths = [], []
ra = wb["Ratios"]
for k, (sh, row) in IDX.items():
    if sh == "Ratios" and k.startswith("Ratios."):
        lec = ra[f"F{row}"].value
        lab = ra[f"A{row}"].value
        val = fcell(ra[f"D{row}"])
        if lec == "Défavorable":
            alerts.append(f"{lab} : {val} (lecture défavorable selon les seuils indicatifs)")
        elif lec == "À surveiller":
            alerts.append(f"{lab} : {val} (à surveiller)")
        elif lec == "Favorable" and len(strengths) < 6 and k.split(".")[1] in ("m_ebitda", "roic", "nd_ebitda", "couv_int", "liq_g", "autonomie", "g_ca"):
            strengths.append(f"{lab} : {val}")
va = wb["Variations"]
big = [r_ for r_ in va.iter_rows(min_row=5, values_only=True) if r_[8] == "Forte variation"]
big.sort(key=lambda r_: -abs(r_[7]))
for row in big[:6]:
    alerts.append(f"{row[0]} : {fp(row[7])} sur un an (variation supérieure à 20 %)")
if len(big) > 6:
    alerts.append(f"{len(big) - 6} autre(s) poste(s) varient de plus de 20 % (voir l'onglet Variations).")
if num(Z[2]) is not None and ZZ[2] != "Sûre":
    alerts.append(f"Score Z'' d'Altman de {fd(Z[2])} : zone « {ZZ[2]} »")

# valorisation
wts = [wb["Hypotheses"].cell(IDX["Hypotheses.p"][1], c).value for c in (2, 3, 4)]
pond_eq = wb["Synthese"][f"E{IDX['Synthese.v_eq'][1]}"].value
pond_ps = wb["Synthese"][f"E{IDX['Synthese.v_ps'][1]}"].value
base_ev, base_tvp = V("DCF.base.s_ev", "B"), V("DCF.base.s_tvpg", "B")
val_txt = [
    f"Dans le scénario <b>Base</b>, la valeur d'entreprise ressort à <b>{fn(base_ev)} {unit}</b> et la valeur des capitaux propres à "
    f"<b>{fn(eq[1])} {unit}</b>" + (f", soit <b>{fd(ps[1])} {devise} par action</b>" if use_ps else "") + ".",
    f"Les scénarios Bear et Bull encadrent la valeur entre {fn(eq[0])} et {fn(eq[2])} {unit} ; la valeur pondérée par les probabilités "
    f"({', '.join(fp(w_, 0) for w_ in wts)}) est de <b>{fn(pond_eq)} {unit}</b>" + (f" ({fd(pond_ps)} {devise} par action)." if use_ps else "."),
    f"La valeur terminale pèse <b>{fp(base_tvp, 0)}</b> de la valeur d'entreprise : l'évaluation est très sensible au WACC, à g et au multiple de sortie "
    f"(voir les tables de sensibilité).",
    f"Multiples implicites (Base) : VE/EBITDA N de {fx(V('DCF.base.s_evebitda', 'B'))}, PER de {fx(V('DCF.base.s_pe', 'B'))} ; "
    f"multiple de sortie implicite via Gordon {fx(V('DCF.base.s_impm', 'B'))}, croissance implicite via le multiple {fp(V('DCF.base.s_impg', 'B'))}.",
]
if cours and use_ps:
    up = ps[1] / cours - 1
    val_txt.append(f"Par rapport au cours de {fd(cours)} {devise}, le scénario Base fait ressortir un potentiel de <b>{fp(up)}</b> "
                   f"({'sous-évaluation' if up > 0 else 'surévaluation'} apparente selon ces hypothèses).")

# ------------------------------------------------------------------ PDF
def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("DV", 7)
    canvas.setFillColor(GREY)
    canvas.drawString(1.8 * cm, 1.1 * cm, f"{name} — Rapport d'analyse financière — généré le {date.today().strftime('%d/%m/%Y')}")
    canvas.drawRightString(A4[0] - 1.8 * cm, 1.1 * cm, f"Page {doc.page}")
    canvas.restoreState()


doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm, topMargin=1.6 * cm, bottomMargin=1.8 * cm,
                        title=f"Analyse financière — {name}", author="Outil d'analyse financière")
W = A4[0] - 3.6 * cm
S = []

# --- Page 1 : couverture et synthèse
S.append(Paragraph(f"Rapport d'analyse financière", ParagraphStyle("t", parent=H1, fontSize=22, spaceAfter=2)))
S.append(Paragraph(f"<b>{name}</b>" + (f" — {secteur}" if secteur else ""), ParagraphStyle("t2", parent=BODY, fontSize=12, textColor=NAVY, alignment=0)))
S.append(Paragraph(f"Exercices {years[0]} à {years[2]} — montants en {unit}", SMALL))
S.append(Spacer(1, 6))
if ctrl != "OK":
    S.append(Paragraph("<font color='#C62828'><b>Attention : certains contrôles de cohérence sont en erreur (onglet Synthese). "
                       "Corriger la saisie avant d'utiliser ce rapport.</b></font>", BODY))
S.append(Paragraph("Synthèse", H2))
head = [f"Ce rapport analyse {name} sur trois exercices : performance, structure financière, flux de trésorerie, "
        f"scores de risque et valorisation par actualisation des flux (DCF)."]
S.append(Paragraph(" ".join(head), BODY))
kp = [["Indicateur"] + years + ["Évol. N/N-1"]]
rowsk = [("Chiffre d'affaires", ca, "n"), ("Croissance du CA", [V("Ratios.g_ca", c) for c in COLS], "p"),
         ("EBITDA", ebe, "n"), ("Marge d'EBITDA", m_ebe, "p"), ("EBIT", ebit, "n"), ("Résultat net", rn, "n"),
         ("CAF", caf, "n"), ("FCFF", fcff, "n"), ("Dette financière nette", dn, "n"),
         ("Dette nette / EBITDA", [V("Ratios.nd_ebitda", c) for c in COLS], "x"),
         ("ROE", [V("Ratios.roe", c) for c in COLS], "p"), ("ROIC", [V("Ratios.roic", c) for c in COLS], "p"),
         ("Liquidité générale", [V("Ratios.liq_g", c) for c in COLS], "d"),
         ("Score Z'' d'Altman", Z, "d"), ("Score de Piotroski (0-9)", PS, "i")]
for lab, ser, kind in rowsk:
    f_ = {"n": fn, "p": fp, "x": fx, "d": fd, "i": lambda v: NA if num(v) is None else str(int(v))}[kind]
    ev = evol(ser[1], ser[2]) if kind == "n" else None
    kp.append([lab] + [f_(v) for v in ser] + [fp(ev) if ev is not None else ""])
S.append(table(kp, [W * 0.38, W * 0.17, W * 0.17, W * 0.17, W * 0.11]))
S.append(Spacer(1, 6))
S.append(Paragraph("Messages clés", H2))
msgs = [perf_txt[0], bs_txt[0], cash_txt[0], val_txt[0]]
S += bullets(msgs)
S.append(Paragraph("Valeur d'entreprise selon le scénario", H2))
vt = [["", "Bear", "Base", "Bull", "Pondéré"],
      [f"Valeur d'entreprise ({unit})"] + [fn(V(f'DCF.{t}.s_ev', 'B')) for t in tags] + [fn(wb['Synthese'][f"E{IDX['Synthese.v_ev'][1]}"].value)],
      [f"Capitaux propres ({unit})"] + [fn(v) for v in eq] + [fn(pond_eq)]]
if use_ps:
    vt.append([f"Valeur par action ({devise})"] + [fd(v) for v in ps] + [fd(pond_ps)])
S.append(table(vt, [W * 0.36] + [W * 0.16] * 4))
S.append(PageBreak())

# --- Page 2 : performance
S.append(Paragraph("1. Performance et rentabilité", H1))
S.append(Image(img_perf, width=W, height=W * 0.345))
S += bullets(perf_txt)
S.append(Paragraph("Soldes intermédiaires de gestion", H2))
sig_rows = [("Chiffre d'affaires", "ca"), ("Marge brute", "mb"), ("Valeur ajoutée", "va"), ("EBE / EBITDA", "ebe"), ("Résultat d'exploitation (EBIT)", "ebit"),
            ("Résultat financier", "rfin"), ("Résultat avant impôt", "rav"), ("Résultat net", "rn"), ("CAF", "caf"), ("NOPAT", "nopat")]
sg = [["Solde"] + years + ["Δ N/N-1"]]
for lab, k in sig_rows:
    ser = [V(f"SIG_CAF.{k}", c) for c in COLS]
    ev = evol(ser[1], ser[2])
    sg.append([lab] + [fn(v) for v in ser] + [fp(ev) if ev is not None else ""])
S.append(table(sg, [W * 0.38, W * 0.17, W * 0.17, W * 0.17, W * 0.11]))
S.append(Paragraph("Décomposition DuPont du ROE", H2))
dp = [["Facteur"] + years]
for lab, k, f_ in (("Charge fiscale (RN / RAI)", "d_tax", fd), ("Poids des intérêts (RAI / EBIT)", "d_int", fd), ("Marge d'EBIT", "d_mebit", fp),
                   ("Rotation de l'actif", "d_rot", fd), ("Levier financier", "d_lev", fd), ("ROE", "d_roe", fp)):
    dp.append([lab] + [f_(V(f"Scores.{k}", c)) for c in COLS])
S.append(table(dp, [W * 0.46, W * 0.18, W * 0.18, W * 0.18]))
d_m, d_r, d_l = [V("Scores.d_mebit", c) for c in "BD"], [V("Scores.d_rot", c) for c in "BD"], [V("Scores.d_lev", c) for c in "BD"]
drivers = {"la marge d'exploitation": (d_m[1] / d_m[0] - 1) if d_m[0] else 0, "la rotation de l'actif": (d_r[1] / d_r[0] - 1) if d_r[0] else 0,
           "le levier financier": (d_l[1] / d_l[0] - 1) if d_l[0] else 0}
main = max(drivers, key=lambda k: abs(drivers[k]))
S.append(Paragraph(f"Sur la période, l'évolution du ROE s'explique principalement par <b>{main}</b> ({drivers[main] * 100:+.0f} % entre {years[0]} et {years[2]}).", BODY))
S.append(PageBreak())

# --- Page 3 : structure
S.append(Paragraph("2. Structure financière et liquidité", H1))
S.append(Image(img_bilan, width=W, height=W * 0.345))
S += bullets(bs_txt)
S.append(Paragraph("Équilibre financier", H2))
eqt = [["Poste"] + years]
for lab, k in (("Capitaux permanents", "kperm"), ("Actif immobilisé", "immo"), ("FRNG", "frng"), ("BFR d'exploitation", "bfr_ex"),
               ("BFR hors exploitation", "bfr_hex"), ("BFR total", "bfr"), ("Trésorerie nette", "tn"), ("Dette financière nette", "dn")):
    eqt.append([lab] + [fn(V(f"Bilan_BFR.{k}", c)) for c in COLS])
S.append(table(eqt, [W * 0.46, W * 0.18, W * 0.18, W * 0.18]))
S.append(PageBreak())

# --- Page 4 : cash
S.append(Paragraph("3. Flux de trésorerie", H1))
S.append(Image(img_cash, width=W, height=W * 0.345))
S += bullets(cash_txt)
cft = [["Flux"] + years]
for lab, k in (("CAF", "caf"), ("Variation du BFR", "d_bfr"), ("Flux de trésorerie d'exploitation", "fte"), ("Capex", "capex"),
               ("FCF opérationnel", "fcf_op"), ("FCFF", "fcff"), ("FCFE", "fcfe"), ("Dividendes versés", "div"), ("Variation de trésorerie constatée", "dt_reel")):
    cft.append([lab] + [fn(V(f"Cashflow.{k}", c)) for c in COLS])
S.append(table(cft, [W * 0.46, W * 0.18, W * 0.18, W * 0.18]))
S.append(PageBreak())

# --- Page 5-6 : ratios
S.append(Paragraph("4. Tableau complet des ratios", H1))
S.append(Paragraph("Lecture de l'exercice N selon des seuils indicatifs (onglet Ratios) ; ils varient selon le secteur et ne remplacent pas une comparaison à des pairs.", SMALL))
rt = [["Ratio"] + years + ["Lecture"]]
style_extra = []
r_i = 1
for cells in ra.iter_rows(min_row=5, max_row=IDX["Ratios.mcap"][1] - 3):
    row = [c.value for c in cells]
    if row[0] is None:
        continue
    if row[1] is None and row[2] is None and row[3] is None:
        rt.append([Paragraph(f"<b>{row[0]}</b>", CELL), "", "", "", ""])
        style_extra += [("BACKGROUND", (0, r_i), (-1, r_i), LIGHT), ("FONT", (0, r_i), (-1, r_i), "DVB", 7.8)]
    else:
        vals = [fcell(c) for c in cells[1:4]]
        lec = row[5] if row[5] and row[5] != "—" else ""
        rt.append([row[0]] + vals + [lec])
        col = {"Favorable": GREEN, "À surveiller": AMBER, "Défavorable": RED}.get(lec)
        if col:
            style_extra.append(("TEXTCOLOR", (4, r_i), (4, r_i), colors.HexColor(col)))
            style_extra.append(("FONT", (4, r_i), (4, r_i), "DVB", 7.8))
    r_i += 1
S.append(table(rt, [W * 0.44, W * 0.14, W * 0.14, W * 0.14, W * 0.14], extra=style_extra, zebra=False, font=7.4))
S.append(PageBreak())
if cours and num(R("per")) is not None:
    S.append(Paragraph("Multiples de marché (cours renseigné)", H2))
    mk = [["Indicateur", "Valeur"]]
    for lab, k, f_ in (("Capitalisation boursière", "mcap", fn), ("Valeur d'entreprise de marché", "ev_mkt", fn), ("PER", "per", fx), ("VE / EBITDA", "ev_ebitda", fx),
                       ("VE / EBIT", "ev_ebit", fx), ("VE / CA", "ev_ca", fd), ("Price-to-book", "pb", fd), ("Rendement du FCFE", "fcfe_y", fp), ("Rendement du dividende", "div_y", fp)):
        mk.append([lab, f_(R(k))])
    S.append(table(mk, [W * 0.5, W * 0.2]))

# --- Scores
S.append(Paragraph("5. Scores de risque et de qualité", H1))
S.append(Paragraph("Score Z'' d'Altman", H2))
zt = [["Composante"] + years]
for lab, k in (("X1 Fonds de roulement / actif", "x1"), ("X2 Résultats mis en réserve / actif", "x2"), ("X3 EBIT / actif", "x3"), ("X4 Capitaux propres / dettes", "x4"), ("Z''", "z")):
    zt.append([lab] + [fd(V(f"Scores.{k}", c)) for c in COLS])
zt.append(["Zone"] + [str(v) for v in ZZ])
S.append(table(zt, [W * 0.46, W * 0.18, W * 0.18, W * 0.18]))
S.append(Paragraph(f"Z'' = {fd(Z[2])} en {years[2]} : zone <b>{ZZ[2]}</b> (seuils : > 2,6 sûre ; 1,1-2,6 grise ; < 1,1 détresse). "
                   "Modèle conçu pour les sociétés non industrielles et non cotées ; c'est un outil de dépistage, pas une prévision.", BODY))
S.append(Paragraph("Score F de Piotroski", H2))
pt = [["Critère", years[1], years[2]]]
for k in ("p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9"):
    sh, row = IDX[f"Scores.{k}"]
    pt.append([wb["Scores"][f"A{row}"].value] + [("Oui" if V(f"Scores.{k}", c) == 1 else "Non") if num(V(f"Scores.{k}", c)) is not None else NA for c in "CD"])
pt.append(["Score total (0-9)"] + [str(int(v)) if num(v) is not None else NA for v in PS[1:]])
S.append(table(pt, [W * 0.62, W * 0.19, W * 0.19]))
S.append(Paragraph(f"Score de {PS[2] if num(PS[2]) is None else int(PS[2])}/9 en {years[2]} ({V('Scores.p_lect')}). 8-9 : solide ; 5-7 : correct ; 0-4 : fragile.", BODY))
S.append(PageBreak())

# --- DCF
S.append(Paragraph("6. Valorisation par actualisation des flux (DCF)", H1))
hy = wb["Hypotheses"]
hrow = lambda k: IDX[f"Hypotheses.{k}"][1]
hp = [["Hypothèse", "Valeur"],
      ["Coût des capitaux propres (CAPM)", fp(hy[f"B{hrow('ke')}"].value, 2)],
      ["Coût de la dette après impôt", fp(hy[f"B{hrow('kd_at')}"].value, 2)],
      ["Poids de la dette cible", fp(hy[f"B{hrow('wd')}"].value, 0)],
      ["WACC", fp(hy[f"B{hrow('wacc')}"].value, 2)],
      ["Croissance perpétuelle g", fp(hy[f"B{hrow('g')}"].value, 1)],
      ["Multiple de sortie VE/EBITDA", fx(hy[f"B{hrow('mult')}"].value)],
      ["Pondération Gordon / multiple", f"{fp(hy[f'B{hrow('wg')}'].value, 0)} / {fp(1 - hy[f'B{hrow('wg')}'].value, 0)}"],
      ["Convention mid-year", "Oui" if hy[f"B{hrow('mid')}"].value == 1 else "Non"]]
S.append(table(hp, [W * 0.55, W * 0.2]))
S.append(Spacer(1, 4))
bp = [["Scénario Base"] + [str(V("DCF.base.hdr", c)) if False else "" for c in ""]]
dcf = wb["DCF"]
rb = lambda k: IDX[f"DCF.base.{k}"][1]
hdr_years = [dcf.cell(rb("hdr"), c).value for c in range(2, 8)]
bp = [[f"Projection Base ({unit})"] + [str(int(y)) if isinstance(y, (int, float)) else str(y) for y in hdr_years]]
for lab, k, f_ in (("Chiffre d'affaires", "ca", fn), ("EBITDA", "ebitda", fn), ("EBIT", "ebit", fn), ("NOPAT", "nopat", fn), ("Capex", "capex", fn),
                   ("Variation du BFR", "dbfr", fn), ("FCFF", "fcff", fn), ("FCFF actualisé", "pv", fn)):
    bp.append([lab] + [f_(dcf.cell(rb(k), c).value) for c in range(2, 8)])
S.append(table(bp, [W * 0.28] + [W * 0.12] * 6))
S.append(Spacer(1, 6))
S.append(Image(img_ff, width=W, height=W * 0.31))
S += bullets(val_txt)


def heat(m, rows_lab, cols_lab, title, fmt):
    data = [[title] + cols_lab]
    for lab, r_ in zip(rows_lab, m):
        data.append([lab] + [fd(v, 2) for v in r_])
    flat = [v for r_ in m for v in r_ if num(v) is not None]
    lo, hi = min(flat), max(flat)
    st = []
    for i, r_ in enumerate(m):
        for j, v in enumerate(r_):
            t = (v - lo) / (hi - lo) if hi > lo else 0.5
            col = colors.Color(1 - 0.35 * t, 1 - 0.1 * t + 0.0, 1 - 0.35 * (1 - t) * 0 - 0.55 * t * 0.5)
            st.append(("BACKGROUND", (j + 1, i + 1), (j + 1, i + 1), colors.Color(0.95 - 0.25 * t, 0.97 - 0.05 * t, 0.95 - 0.2 * t)))
    st.append(("BOX", (3, 3), (3, 3), 1.2, NAVY))
    return table(data, [W * 0.22] + [W * 0.13] * 5, extra=st, zebra=False)


w_ax = [dcf.cell(IDX["DCF.sens1_hdr"][1] + 1 + i, 2).value for i in range(5)]
g_ax = [dcf.cell(IDX["DCF.sens1_hdr"][1], 3 + j).value for j in range(5)]
m_ax = [dcf.cell(IDX["DCF.sens2_hdr"][1], 3 + j).value for j in range(5)]
unit_s = f"{devise} par action" if use_ps else f"{unit}, capitaux propres"
S.append(PageBreak())
S.append(Paragraph("Tables de sensibilité du scénario Base", H1))
S.append(Paragraph(f"Sensibilité (Base, Gordon-Shapiro) — {unit_s}", H2))
S.append(heat(sens1, [fp(x, 1) for x in w_ax], [fp(x, 1) for x in g_ax], "WACC \\ g", None))
S.append(Paragraph(f"Sensibilité (Base, multiple de sortie) — {unit_s}", H2))
S.append(heat(sens2, [fp(x, 1) for x in w_ax], [fx(x) for x in m_ax], "WACC \\ multiple", None))
S.append(PageBreak())

# --- Positionnement de valeur : football field, comparables, bêtas, LBO, M&A
def cellv(sh, addr):
    return wb[sh][addr].value


S.append(Paragraph("7. Positionnement de valeur : méthodes croisées", H1))
ffr0, ffr1 = IDX["Synthese.ff_first"][1], IDX["Synthese.ff_last"][1]
ff_rows = []
for r_ in range(ffr0, ffr1 + 1):
    lab_, lo_, hi_ = wb["Synthese"][f"A{r_}"].value, num(wb["Synthese"][f"B{r_}"].value), num(wb["Synthese"][f"C{r_}"].value)
    if lo_ is not None and hi_ is not None:
        ff_rows.append((lab_, lo_, hi_))
if ff_rows:
    fig, ax = plt.subplots(figsize=(7.2, 2.6))
    for i_, (lab_, lo_, hi_) in enumerate(ff_rows[:-1][::-1] if len(ff_rows) > 1 else ff_rows):
        ax.barh(i_, max(hi_ - lo_, 0.0001), left=lo_, color="#2F5597", height=0.45)
        ax.text(lo_, i_ + 0.32, fd(lo_), ha="center", fontsize=7)
        ax.text(hi_, i_ + 0.32, fd(hi_), ha="center", fontsize=7)
    ax.set_yticks(range(len(ff_rows) - 1))
    ax.set_yticklabels([r_[0] for r_ in ff_rows[:-1][::-1]], fontsize=7.5)
    if cours:
        ax.axvline(cours, color="#C62828", ls="--", lw=1)
        ax.text(cours, len(ff_rows) - 1.1, f"Cours {fd(cours)}", color="#C62828", fontsize=7.5, ha="center")
    ax.set_xlabel(f"Valeur par action ({devise})", fontsize=8)
    style_ax(ax)
    S.append(Image(savefig(fig, "football.png"), width=16.5 * cm, height=6 * cm))
S.append(Paragraph("Chaque méthode répond à une question différente : le DCF valorise les flux prévus, les comparables mesurent le prix auquel le marché "
                   "valorise des pairs, le LBO donne la capacité de paiement d'un acquéreur financier. Un écart entre méthodes est une information, pas une anomalie.", BODY))
cs_ = IDX["Comparables.stat.Médiane"][1]
q1_, q3_ = IDX["Comparables.stat.1er quartile"][1], IDX["Comparables.stat.3e quartile"][1]
n_ = IDX["Comparables.stat.Nombre de pairs retenus"][1]
S.append(Paragraph("Comparables boursiers", H2))
S.append(Paragraph(f"Panel de <b>{cellv('Comparables', f'B{n_}')}</b> pair(s) : multiple VE/EBITDA médian {fx(cellv('Comparables', f'B{cs_}'))} "
                   f"(Q1 {fx(cellv('Comparables', f'B{q1_}'))}, Q3 {fx(cellv('Comparables', f'B{q3_}'))}). "
                   "Les multiples d'agrégateurs gratuits ont des dates et définitions hétérogènes : à recouper avant tout usage décisionnel.", BODY))
per_med_ = cellv("Comparables", f"C{cs_}")
S.append(Paragraph(f"PER médian du panel : {fx(per_med_)} (Q1 {fx(cellv('Comparables', f'C{q1_}'))}, Q3 {fx(cellv('Comparables', f'C{q3_}'))}). "
                   "Un panel de moins de 3 pairs sur un multiple donné ne permet pas de conclure : le modèle l'affiche comme alerte.", BODY))
if "Synthese.m_ev" in IDX:
    mk_ = lambda k: cellv("Synthese", f"B{IDX['Synthese.' + k][1]}")
    S.append(Paragraph("Ce que le marché anticipe", H2))
    S.append(Paragraph(f"Au cours actuel, la valeur d'entreprise de marché est de {fn(mk_('m_ev'))} {unit}, soit {fx(mk_('m_eve'))} l'EBITDA N "
                       f"({fx(mk_('m_evea'))} l'EBITDA ajusté de la rémunération en actions), {fx(mk_('m_evs'))} le chiffre d'affaires et un PER de {fx(mk_('m_per'))} ; "
                       f"le rendement du FCFF sur la valeur d'entreprise est de {fp(mk_('m_fcf'))}. Valeur DCF pondérée par rapport au cours : {fp(mk_('m_dcf'))}.", BODY))
S.append(Paragraph("Coût du capital et bêta", H2))
S.append(Paragraph(f"Bêta désendetté médian des pairs : {fd(cellv('WACC_Pairs', 'B16'))} ; ré-endetté à la structure cible : {fd(cellv('WACC_Pairs', 'B20'))} "
                   f"(ajusté Blume : {fd(cellv('WACC_Pairs', 'B21'))}). Un panel de quelques valeurs donne une médiane fragile : le choix du mode est à documenter.", BODY))
S.append(Paragraph("LBO : capacité de paiement d'un sponsor", H2))
S.append(Paragraph(f"Aux hypothèses de l'onglet LBO (entrée {fx(cellv('LBO', 'B5'))}, dette {fx(cellv('LBO', 'B7'))} d'EBITDA, sortie {fx(cellv('LBO', 'B10'))}), "
                   f"le TRI du sponsor ressort à <b>{fp(cellv('LBO', 'B38'))}</b> (MOIC {fx(cellv('LBO', 'B37'))}). "
                   f"Prix maximal par action pour un TRI de {fp(cellv('LBO', 'B42'), 0)} / {fp(cellv('LBO', 'C42'), 0)} / {fp(cellv('LBO', 'D42'), 0)} : "
                   f"{fd(cellv('LBO', 'B47'))} / {fd(cellv('LBO', 'C47'))} / {fd(cellv('LBO', 'D47'))} {devise}.", BODY))
ma_a, ma_b = IDX["MA.acc"][1], IDX["MA.px"][1]
S.append(Paragraph("Accrétion / dilution (acquéreur illustratif)", H2))
S.append(Paragraph(f"Avec une offre à {fd(cellv('MA', f'B{ma_b}'))} {devise} par action, l'opération ferait varier le BPA de l'acquéreur de "
                   f"{fp(cellv('MA', f'B{ma_a}'))} (N+1), {fp(cellv('MA', f'C{ma_a}'))} (N+2), {fp(cellv('MA', f'D{ma_a}'))} (N+3). "
                   "Acquéreur, financement et synergies sont des hypothèses d'école : ils ne reflètent aucune opération réelle.", BODY))
S.append(PageBreak())

# --- Note de marché : DCF 10 ans, DCF inversé, cours cible, crédit
if "Reco.tp" in IDX:
    rc = lambda k: cellv("Reco_Marche", f"B{IDX['Reco.' + k][1]}")
    d10 = lambda k: cellv("DCF_10ans", f"B{IDX['DCF10.' + k][1]}")
    S.append(Paragraph("8. Note de marché : horizon long, anticipations et cours cible", H1))
    S.append(Paragraph("DCF à 10 ans (convergence)", H2))
    S.append(Paragraph(f"En prolongeant le scénario Base jusqu'en année 10, avec une croissance qui converge vers g et une marge qui converge vers la cible saisie, "
                       f"la valeur par action ressort à <b>{fd(d10('ps'))} {devise}</b> ({fp(d10('vs5'))} vs le DCF à 5 ans). "
                       f"Le poids de la valeur terminale (Gordon) passe à {fp(d10('tvpg'), 0)} de la valeur d'entreprise ; la croissance annuelle moyenne du chiffre d'affaires sur 10 ans est de {fp(d10('cagr'))}.", BODY))
    S.append(Paragraph("Ce que le cours anticipe (DCF inversé)", H2))
    S.append(Paragraph(f"Si le scénario Base se réalise et que l'on achète au cours actuel, le rendement annuel implicite est de <b>{fp(d10('implied'))}</b>, "
                       f"contre un WACC retenu de {fp(rc('wacc'))} (écart {fp(rc('gap'))}). Lecture : {rc('read')}.", BODY))
    S.append(Paragraph("Cours cible et profil risque / rendement", H2))
    sk_ = rc("skew")
    S.append(Paragraph(f"Moyenne pondérée de {rc('nm')} méthode(s) disponible(s) (DCF 5 ans, DCF 10 ans, comparables VE/EBITDA, comparables PER) : "
                       f"<b>{fd(rc('tp'))} {devise}</b>, soit un potentiel de {fp(rc('up'))} vs le cours. Appréciation indicative par règle mécanique (seuils {fp(rc('buy'), 0)} / {fp(rc('sell'), 0)}) : "
                       f"<b>{rc('rating')}</b>. Valeur pondérée par les probabilités : {fd(rc('wtd'))} {devise} (rendement espéré {fp(rc('er'))}) ; "
                       f"ratio gain potentiel / perte potentielle vs le cours : {fx(sk_, 2) if isinstance(sk_, (int, float)) else sk_}.", BODY))
    S.append(Paragraph("Qualité de crédit (indicatif)", H2))
    icr_ = rc("icr")
    S.append(Paragraph(f"Dette nette / EBITDA : {fx(rc('nd'))} (négatif = trésorerie nette). Couverture des intérêts : {fx(icr_) if isinstance(icr_, (int, float)) else icr_}. "
                       f"Notation synthétique indicative : {rc('crd')} (table pédagogique, à ne pas confondre avec la notation d'une agence).", BODY))
    S.append(Paragraph("Qualité des hypothèses, risques et conviction", H2))
    so_ = lambda k: cellv("Sources", f"B{IDX['Src.' + k][1]}")
    ql_ = lambda k: cellv("Qualitatif", f"B{IDX['Ql.' + k][1]}")
    age_ = cellv("Marche", f"B{IDX['Mk.age'][1]}")
    S.append(Paragraph(f"Hypothèses documentées (source citée, datée, de moins de 180 jours) : <b>{fp(so_('pct'), 0)}</b> ; le reste relève du jugement de l'analyste ou d'une source ancienne ou non datée (onglet Sources). "
                       f"Ancienneté du cours : {age_} jour(s). Profil de risque qualitatif : <b>{ql_('prof')}</b> (score maximal {ql_('mx')} sur 25). "
                       f"Niveau de conviction indicatif de la note : <b>{rc('conv_lvl')}</b> ({rc('conv_n')} contrôles réussis sur 8).", BODY))
    S.append(Paragraph("Ces sorties reposent sur des hypothèses illustratives et des règles mécaniques ; elles ne constituent pas un conseil en investissement.", BODY))
    S.append(PageBreak())

# --- Points d'attention
S.append(Paragraph("9. Points d'attention et points forts", H1))
S.append(Paragraph("Points forts", H2))
S += bullets(strengths or ["Aucun point fort signalé automatiquement."])
S.append(Paragraph("Points d'attention", H2))
S += bullets(alerts or ["Aucun point d'attention signalé automatiquement."])
S.append(Paragraph("Méthodologie et limites", H2))
S += bullets([
    "Les calculs proviennent du classeur Excel (formules visibles) alimenté par les états financiers saisis ; ce rapport en est la mise en forme.",
    "CAF = résultat net + dotations. Capex estimé par Δ immobilisations nettes + dotations lorsqu'il n'est pas saisi. NOPAT = EBIT − impôt normatif.",
    "DCF : WACC par le CAPM ; valeur terminale par Gordon-Shapiro et par multiple de sortie, pondérées. Les hypothèses sont celles de l'onglet Hypotheses.",
    "Les seuils de lecture des ratios dépendent du profil choisi (Générique / Services / Industrie) : ce sont des jugements d'analyste, à comparer aux pairs.",
    "IFRS 16 : selon le traitement choisi, les loyers sont déduits du FCFF et la dette nette exclut les dettes locatives (ou l'inverse).",
    "Les données de comparables proviennent d'agrégateurs gratuits aux dates hétérogènes ; les hypothèses de valorisation sont illustratives.",
    "Document d'analyse et de formation : il ne constitue ni un conseil en investissement, ni une certification des comptes.",
])
doc.build(S, onFirstPage=footer, onLaterPages=footer)
print("Rapport généré :", out)
