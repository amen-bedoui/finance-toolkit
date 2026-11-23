"""Panel de pairs à partir du jeu Hugging Face DenyTranDFW/SEC_10K_FSNoNDS_Zip (états SEC extraits en parquet).
Usage : python tools/build_peer_panel.py <dossier_parquet> <sortie.pkl>
Les données sont des états financiers (pas de cours) : multiples de valorisation à compléter avec des cours datés."""
import pandas as pd, glob, re, os, collections
import sys
os.chdir(sys.argv[1])
REV=["Revenues","RevenueFromContractWithCustomerExcludingAssessedTax","SalesRevenueNet","RevenueFromContractWithCustomerIncludingAssessedTax"]
TAGS={"rev":REV,"opi":["OperatingIncomeLoss"],"ni":["NetIncomeLoss","ProfitLoss"],"da":["DepreciationDepletionAndAmortization","DepreciationAndAmortization","DepreciationAmortizationAndAccretionNet","DepreciationAmortizationAndOther"],
 "cfo":["NetCashProvidedByUsedInOperatingActivities"],"capex":["PaymentsToAcquirePropertyPlantAndEquipment","PaymentsToAcquireProductiveAssets"],
 "assets":["Assets"],"eq":["StockholdersEquity","StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
 "cash":["CashAndCashEquivalentsAtCarryingValue"],"sti":["MarketableSecuritiesCurrent","AvailableForSaleSecuritiesDebtSecuritiesCurrent","ShortTermInvestments"],
 "ltd":["LongTermDebtNoncurrent","LongTermDebt","LongTermDebtAndCapitalLeaseObligations"],"std":["DebtCurrent","LongTermDebtCurrent"],"sbc":["ShareBasedCompensation","AllocatedShareBasedCompensationExpense"],
 "rd":["ResearchAndDevelopmentExpense"]}
files=collections.defaultdict(list)
for f in glob.glob("10K_*.parquet"):
    m=re.match(r"10K_(\w+)_(\d{10})_(\d+)_(\d{8})\.parquet",f)
    files[m[2]].append((m[1],m[3],f))
rows=[]
for cik,lst in files.items():
    # latest filing = max adsh-year/seq by period file date
    byfil=collections.defaultdict(dict)
    for t,adsh,f in lst: byfil[adsh][t]=f
    best=None
    for adsh,d in byfil.items():
        if len(d)<3: continue
        mx=0;name=None
        for t,f in d.items():
            df=pd.read_parquet(f,columns=["name","ddate"]); mx=max(mx,int(df.ddate.max())); name=df.name.iloc[0]
        if best is None or mx>best[0]: best=(mx,adsh,d,name)
    if not best: continue
    mx,adsh,d,name=best
    df=pd.concat([pd.read_parquet(f) for f in d.values()])
    df["v"]=pd.to_numeric(df.value.astype(str).str.replace(",",""),errors="coerce")
    ds=sorted(df.ddate.unique())
    def get(keys,dd,stmt=None):
        for k in keys:
            s=df[(df.tag==k)&(df.ddate==dd)]
            if stmt: s=s[s.stmt==stmt]
            if len(s): return float(s.v.iloc[0])
        return None
    # fiscal-year-end dates present in IS with ~12m: use ddate list from IS
    isd=sorted(df[df.stmt=="IS"].ddate.unique())
    if len(isd)<2: continue
    last,prev=isd[-1],isd[-2]
    r={"cik":cik.lstrip("0"),"name":name,"last":last,"adsh":adsh}
    for k,tg in TAGS.items():
        r[k]=get(tg,last); r[k+"_p"]=get(tg,prev)
    if r["da"] is None:
        c=df[(df.stmt=="CF")&(df.ddate==last)&df.tag.str.contains("Depreciation|Amortization",regex=True)&~df.tag.str.contains("Debt|Intangible|Lease|Premium|Financing|Deferred|Right",regex=True)]
        r["da"]=float(c.v.iloc[0]) if len(c) else None
        c=df[(df.stmt=="CF")&(df.ddate==prev)&df.tag.str.contains("Depreciation|Amortization",regex=True)&~df.tag.str.contains("Debt|Intangible|Lease|Premium|Financing|Deferred|Right",regex=True)]
        r["da_p"]=float(c.v.iloc[0]) if len(c) else None
    for k in ("capex","capex_p"):
        if r[k] is not None: r[k]=abs(r[k])
    rows.append(r)
P=pd.DataFrame(rows); P.to_pickle(sys.argv[2]); print(len(P))
