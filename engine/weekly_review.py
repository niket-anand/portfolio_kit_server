#!/usr/bin/env python3
# ==============================================================================
#  WEEKLY PORTFOLIO REVIEW  —  CONVICTION-6  (one-shot builder)
#  Drop in this week's files, set the 5 paths below, run.  Produces the full
#  6-sheet workbook: Portfolio + Factors · Equity Master · Industry ·
#  Cup & Handle · Recommendations · TradingView Watchlist.
#
#  RUN:  python3 weekly_review.py
#
#  WEEKLY INPUTS (set in CONFIG):
#   1. CURRENT_FILES   this week's broker exports (ICICI .xls TSV + Kite .csv)
#   2. LASTWEEK        last week's broker exports (list)  OR a prior review .xlsx
#   3. SCREENER_HELD   Screener.in CSV covering held names   (scoring universe)
#   4. WATCHLIST_CSV   TradingView watchlist  (cols: Group,Symbol[,Company,Last])
#   5. SCREENER_WATCH  Screener.in CSV covering watchlist names (optional; unions
#                      with #3 so watch-candidates get scored & recommended)
#
#  Notes: broker files auto-detected by content. Holdings merged by ISIN across
#  accounts. Scoring delegated to the maintained skill scorer so the framework
#  stays canonical. Not investment advice.
# ==============================================================================
import pandas as pd, numpy as np, re, os, sys, json, subprocess, shutil

# =============================== CONFIG =======================================
# You only edit the TWO date lines below each week. Everything else is auto-read
# from the inputs/ folder — just drop your files in and run.
import os, glob
ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
HERE = os.path.dirname(ENGINE_DIR)   # kit root
os.chdir(HERE)                      # work + outputs land next to this script
IN   = os.path.join(HERE, "inputs")

# ---- EDIT THESE EACH WEEK (or set them in the web app) ----
DATE_CUR = os.environ.get("WR_DATE_CUR", "4 July 2026");  LBL_CUR = os.environ.get("WR_LBL_CUR", "4-Jul")     # this week
DATE_PREV= os.environ.get("WR_DATE_PREV","30 June");      LBL_PREV= os.environ.get("WR_LBL_PREV","30-Jun")    # last week
# -----------------------------------------------------------

def _g(*pats):
    out=[]
    for p in pats: out += sorted(glob.glob(os.path.join(IN,p)))
    return out
def _first(*names):
    for n in names:
        p=os.path.join(IN,n)
        if os.path.exists(p): return p
    return None

_current  = _g("current/*.xls","current/*.csv","current/*.xlsx")
_lw_xlsx  = _g("lastweek/*.xlsx")
_lastweek = _lw_xlsx[0] if _lw_xlsx else _g("lastweek/*.xls","lastweek/*.csv")

CFG = dict(
    CURRENT_FILES  = _current,                         # inputs/current/*  (broker exports)
    LASTWEEK       = _lastweek,                         # inputs/lastweek/* (prior .xlsx OR broker files)
    SCREENER_HELD  = _first("screener_holdings.csv"),   # inputs/screener_holdings.csv
    SCREENER_WATCH = _first("screener_watchlist.csv"),  # inputs/screener_watchlist.csv (optional; scores watch names)
    WATCHLIST_CSV  = _first("watchlist.csv"),           # inputs/watchlist.csv  (Group,Symbol[,Company,Last])
    DATE_CUR=DATE_CUR, DATE_PREV=DATE_PREV, LBL_CUR=LBL_CUR, LBL_PREV=LBL_PREV,
    OUT            = os.path.join(HERE,"outputs","Weekly Portfolio Review.xlsx"),
    SYMBOLS_CSV    = os.path.join(HERE,"data","symbols.csv"),
    SCORER         = os.path.join(ENGINE_DIR,"conviction6.py"),
)
os.makedirs(os.path.join(HERE,"outputs"), exist_ok=True)

# =============================== HELPERS ======================================
def num(x):
    if pd.isna(x): return np.nan
    s=str(x).strip().replace(",","").replace(" ","")
    if s in ("","-","nan"): return np.nan
    neg=s.startswith("(") and s.endswith(")"); s=s.strip("()").replace("%","")
    try: v=float(s)
    except: return np.nan
    return -v if neg else v

def KEY(nse, comp=""):
    b=str(nse).strip().upper()
    if b in ("","NAN","NONE"): b=str(comp).strip().upper()
    b=re.sub(r"-(BE|BZ|BL|SM|P\d+|N\d+)$","",b)
    return b.replace(" ","").replace(".","").replace("&","")

# ---- NSE resolution (ISIN -> NSE) : symbols.csv + Screener NSE + manual -------
_sym = pd.read_csv(CFG["SYMBOLS_CSV"]) if os.path.exists(CFG["SYMBOLS_CSV"]) else pd.DataFrame(columns=["NSE Symbol","ISIN"])
S2I = dict(zip(_sym["NSE Symbol"], _sym["ISIN"]))
I2S = dict(zip(_sym["ISIN"], _sym["NSE Symbol"]))
MAN_S2I = {"IDEAFORGE":"INE0LZP01017","GRWRHITECH":"INE291A01017","MSTCLTD":"INE255X01014","HSCL":"INE019C01026",
 "HFCL":"INE548A01028","MAHABANK":"INE457A01014","CGPOWER":"INE067A01029","RPEL":"INE912T01018",
 "NETWEB":"INE0NT901020","OPTIEMUS":"INE738Y01010","HDFCGOLD":"INF179KC1981"}
MAN_I2S = {v:k for k,v in MAN_S2I.items()}
MAN_I2S.update({"INE089C01029":"STLTECH","INE00LO01017":"CRAFTSMAN","INE600L01024":"LALPATHLAB",
 "INE881D01027":"OFSS","INE961O01016":"RAINBOW","INF109K012R6":"NIFTYIETF","INF769K01JP9":"GOLDETF"})
def isin2nse(isin, fb=""): return I2S.get(isin) or MAN_I2S.get(isin) or fb

def classify(name, nse=""):
    t=(str(name)+" "+str(nse)).upper()
    if any(w in t for w in ["GOLD","SILVER"]): return "Gold/Commodities"
    if "NIFTY" in t or "ETF" in t: return "ETF (Index)"
    return "Equity"

def nice(n):
    n=str(n).strip()
    return n.title() if n.isupper() else n

# =============================== HOLDINGS =====================================
def detect_and_load(path, holder_hint=None):
    """Return list of holding dicts. Auto-detect ICICI-TSV vs Kite-CSV."""
    with open(path, "r", errors="ignore") as f: head=f.readline()
    holder = holder_hint or os.path.basename(path)
    if "Stock Symbol" in head and "\t" in head:          # ICICI tab-separated
        d=pd.read_csv(path, sep="\t"); d.columns=[c.strip() for c in d.columns]
        d=d.dropna(subset=["Stock Symbol"]); d=d[d["Stock Symbol"].astype(str).str.strip()!=""]
        out=[]
        for _,r in d.iterrows():
            isin=str(r["ISIN Code"]).strip()
            out.append(dict(ISIN=isin, Name=str(r["Company Name"]).strip(), NSE=isin2nse(isin),
                Qty=num(r["Qty"]), Avg=num(r["Average Cost Price"]), CMP=num(r["Current Market Price"]),
                Inv=num(r.get("Value At Cost")), Cur=num(r["Value At Market Price"]),
                PnL=num(r["Unrealized Profit/Loss"]), acct=holder))
        return out
    else:                                                # Kite CSV
        d=pd.read_csv(path); d.columns=[c.strip() for c in d.columns]
        d=d[[c for c in d.columns if c and not c.startswith("Unnamed")]]
        d=d.dropna(subset=["Instrument"]); d=d[d["Instrument"].astype(str).str.strip()!=""]
        out=[]
        for _,r in d.iterrows():
            raw=str(r["Instrument"]).strip(); base=re.sub(r"-(BE|BZ|BL|P\d+|N\d+)$","",raw)
            isin=MAN_S2I.get(base) or S2I.get(base) or f"SYM:{base}"
            out.append(dict(ISIN=isin, Name=nice(base), NSE=base,
                Qty=num(r["Qty."]), Avg=num(r["Avg. cost"]), CMP=num(r["LTP"]),
                Inv=num(r.get("Invested")), Cur=num(r["Cur. val"]), PnL=num(r["P&L"]), acct=holder))
        return out

def account_label(path):
    b=os.path.basename(path).lower()
    # heuristic labels; adjust to your filenames if needed
    return None

def consolidate(files):
    """Merge all broker files by ISIN -> one row per true position."""
    rows=[]
    for p in files:
        rows += detect_and_load(p, holder_hint=os.path.basename(p))
    df=pd.DataFrame(rows)
    df["Class"]=[classify(r.Name,r.NSE) for r in df.itertuples()]
    def agg(g):
        qty=g["Qty"].sum(); cur=g["Cur"].sum(); pnl=g["PnL"].sum()
        cost=(g["Qty"]*g["Avg"]).sum(); inv=g["Inv"].sum() if g["Inv"].notna().any() else cost
        avg=cost/qty if qty else np.nan
        cmp=g["CMP"].dropna().iloc[0] if g["CMP"].notna().any() else np.nan
        accts=list(dict.fromkeys(g["acct"]))
        name=max(g["Name"].tolist(), key=len); nse=next((x for x in g["NSE"] if x), "")
        return pd.Series(dict(Company=name, NSE=nse, Acct="+".join(accts), nAcct=len(accts),
            Class=g["Class"].iloc[0], Qty=qty, Avg=avg, CMP=cmp, Cur=cur, PnL=pnl, Cost=(inv if inv else cost)))
    m=df.groupby("ISIN", as_index=False).apply(agg, include_groups=False)
    m["PnLpct"]=np.where(m["Cost"]>0, m["PnL"]/m["Cost"]*100, np.nan)
    m["key"]=m.apply(lambda r:KEY(r["NSE"],r["Company"]), axis=1)
    return m

# =============================== BASELINE =====================================
def load_baseline(spec):
    """spec = prior review .xlsx (reads Equity Master + Portfolio+Factors) OR list of broker files."""
    if isinstance(spec, (list,tuple)):
        b=consolidate(list(spec))
        return pd.DataFrame(dict(key=b["key"], Company=b["Company"], NSE=b["NSE"],
                Class=b["Class"], JunVal=b["Cur"], JunQty=b["Qty"], JunReco=""))
    # else: prior workbook
    xl=pd.ExcelFile(spec); SKIP=re.compile(r"^(sub-?total|grand|=|\+|full book|exited|total\b)", re.I)
    em=pd.read_excel(xl,"Equity Master",header=4)
    ren={}
    # pick exactly ONE baseline value col (prefer LBL_PREV match, else last "*Val") and one Qty col
    valcols=[c for c in em.columns if str(c).strip().endswith("Val")]
    qtycols=[c for c in em.columns if str(c).strip().endswith("Qty")]
    bval=next((c for c in valcols if CFG["LBL_PREV"] in str(c)), valcols[-1] if valcols else None)
    bqty=next((c for c in qtycols if CFG["LBL_PREV"] in str(c)), qtycols[-1] if qtycols else None)
    if bval: ren[bval]="JunVal"
    if bqty: ren[bqty]="JunQty"
    em=em.rename(columns=ren); em=em.rename(columns={"Recommendation":"JunReco"})
    em=em.dropna(subset=["Company"]); em=em[~em["Company"].astype(str).str.match(SKIP)]
    em["JunVal"]=pd.to_numeric(em["JunVal"],errors="coerce"); em=em.dropna(subset=["JunVal"]); em["Class"]="Equity"
    parts=[em[["Company","NSE","JunQty","JunVal","JunReco","Class"]]]
    # gold/etf from Portfolio + Factors
    try:
        pf=pd.read_excel(xl,"Portfolio + Factors",header=4)
        vcol=[c for c in pf.columns if str(c).endswith("Value")][-1]
        pf=pf.rename(columns={vcol:"JunVal"}); pf["JunVal"]=pd.to_numeric(pf["JunVal"],errors="coerce")
        G2N={"HDFC GOLD ETF":"HDFCGOLD","MIRAE ASSET GOLD ETF":"GOLDETF","ICICI PRUDENTIAL NIFTY 50 ETF":"NIFTYIETF"}
        sec=None; rr=[]
        for _,r in pf.iterrows():
            c=str(r["Company"]).strip()
            if c.upper().startswith("GOLD /"): sec="Gold/Commodities"; continue
            if c.upper().startswith("ETF"): sec="ETF (Index)"; continue
            if c.upper().startswith("EQUIT"): sec="Q"; continue
            if SKIP.match(c) or c=="nan" or pd.isna(r["JunVal"]) or sec in (None,"Q"): continue
            rr.append(dict(Company=c, NSE=G2N.get(c.upper(),c.upper()),
                JunQty=r.get("Jun Qty"), JunVal=r["JunVal"], JunReco="", Class=sec))
        if rr: parts.append(pd.DataFrame(rr))
    except Exception: pass
    base=pd.concat(parts, ignore_index=True)
    base["key"]=base.apply(lambda r:KEY(r["NSE"],r["Company"]), axis=1)
    base=base.groupby("key",as_index=False).agg(Company=("Company","first"),NSE=("NSE","first"),
        Class=("Class","first"),JunVal=("JunVal","sum"),JunQty=("JunQty","sum"),JunReco=("JunReco","first"))
    return base

# =============================== SCORING ======================================
def run_scorer(screener_paths, held_nse_codes):
    """Union screener CSVs -> score via canonical skill scorer -> scored DataFrame."""
    dfs=[]
    for p in screener_paths:
        if p and os.path.exists(p):
            d=pd.read_csv(p); d.columns=[c.strip() for c in d.columns]; dfs.append(d)
    if not dfs: raise SystemExit("No Screener CSV found.")
    scr=pd.concat(dfs, ignore_index=True)
    if "ISIN Code" in scr: scr=scr.drop_duplicates("ISIN Code")
    scr.to_csv("_screener_combined.csv", index=False)
    shutil.copy("_screener_combined.csv","screener.csv")   # build code reads screener.csv
    open("held.txt","w").write("\n".join(held_nse_codes))
    env=dict(os.environ, HELD_FILE=os.path.abspath("held.txt"))
    subprocess.run([sys.executable, CFG["SCORER"], "screener.csv", "scored.pkl"], check=True, env=env)
    return pd.read_pickle("scored.pkl")

# =============================== MASTER =======================================
def build_master(cur, base, sc, scr):
    for df,c in [(sc,"key")]:
        if "key" not in df: df["key"]=df.apply(lambda r:KEY(r.get("NSE"),r.get("Name")),axis=1)
    scr["key"]=scr.apply(lambda r:KEY(r.get("NSE Code"),r.get("Name")),axis=1)
    sc["key"]=sc.apply(lambda r:KEY(r.get("NSE"),r["Name"]),axis=1)
    bI=base.set_index("key"); sI=sc.set_index("key"); rI=scr.set_index("key")
    # secondary screener index by ISIN so abbreviated broker symbols still match
    rISIN=scr.dropna(subset=["ISIN Code"]).drop_duplicates("ISIN Code").set_index("ISIN Code") if "ISIN Code" in scr.columns else None
    DII_COL="Change in DII holding" if "Change in DII holding" in scr.columns else "Change in DII holding 3Years"
    def gv(idx,k,c):
        try:
            v=idx.loc[k,c]; return v if (not isinstance(v,pd.Series) and pd.notna(v)) else (v.iloc[0] if isinstance(v,pd.Series) else np.nan)
        except: return np.nan
    def gv2(k,isin,c):
        """screener value by key, falling back to ISIN match (handles ICICI symbol abbreviations)"""
        v=gv(rI,k,c)
        if (v is None or (isinstance(v,float) and pd.isna(v))) and rISIN is not None and isinstance(isin,str) and isin in rISIN.index:
            v=gv(rISIN,isin,c)
        return v
    def dfh_of(k,isin):
        d=gv2(k,isin,"Down from 52w high")
        if pd.notna(d): return -abs(float(d))          # negative = below 52w high
        hp=gv2(k,isin,"High price"); cp=gv2(k,isin,"Current Price")
        if pd.notna(hp) and pd.notna(cp) and float(hp)!=0: return (float(cp)-float(hp))/float(hp)*100
        return np.nan
    rows=[]
    for _,r in cur.iterrows():
        k=r["key"]
        junv=gv(bI,k,"JunVal"); junq=gv(bI,k,"JunQty")
        S=(k in sI.index)
        disp = gv(sI,k,"Name") if S else (gv(rI,k,"Name") if k in rI.index else nice(r["Company"]))
        if not isinstance(disp,str) or not disp: disp=nice(r["Company"])
        rows.append(dict(key=k, Company=disp, NSE=r["NSE"], Acct=r["Acct"], Class=r["Class"],
            JulQty=r["Qty"], JunQty=junq, Avg=r["Avg"], CMP=r["CMP"], JulVal=r["Cur"], JunVal=junv,
            PnL=r["PnL"], PnLpct=r["PnLpct"], PortPct=np.nan,
            RSI=gv2(k,r.get("ISIN"),"RSI"), Dfh=dfh_of(k,r.get("ISIN")), Wk=gv2(k,r.get("ISIN"),"Return over 1week"), M1=gv2(k,r.get("ISIN"),"Return over 1month"),
            QSales=gv2(k,r.get("ISIN"),"YOY Quarterly sales growth"), QProfit=gv2(k,r.get("ISIN"),"YOY Quarterly profit growth"),
            S3=gv2(k,r.get("ISIN"),"Sales growth 3Years"), S5=gv2(k,r.get("ISIN"),"Sales growth 5Years"), OPM=gv2(k,r.get("ISIN"),"OPM"),
            PE=gv2(k,r.get("ISIN"),"Price to Earning"), ROCE=gv2(k,r.get("ISIN"),"Return on capital employed"), Pledge=gv2(k,r.get("ISIN"),"Pledged percentage"),
            Industry=gv2(k,r.get("ISIN"),"Industry"),
            FIIchg=gv2(k,r.get("ISIN"),"Change in FII holding"), DIIchg=gv2(k,r.get("ISIN"),DII_COL),
            Fund=gv(sI,k,"Fund"), FG=gv(sI,k,"FGrade"), Tech=gv(sI,k,"Tech"), TG=gv(sI,k,"TGrade"), Grand=gv(sI,k,"Grand"),
            Verdict=gv(sI,k,"Verdict"), Conv=gv(sI,k,"Conv"), Tier=gv(sI,k,"Tier"), Wt=gv(sI,k,"Weight"), PEG=gv(sI,k,"PEG"),
            MA=gv(sI,k,"MA Position"), Mom=gv(sI,k,"Mom"), Reco=gv(sI,k,"Reco"), Stage1=gv(sI,k,"Stage1")))
    M=pd.DataFrame(rows)
    M["MA"]=M["MA"].apply(lambda x: x if isinstance(x,str) else "")
    book=M["JulVal"].sum(); M["PortPct"]=M["JulVal"]/book*100
    M["Dqty"]=M["JulQty"]-M["JunQty"]; M["Dval"]=M["JulVal"]-M["JunVal"].fillna(0)
    M["PctChg"]=np.where(M["JunVal"]>0,(M["JulVal"]-M["JunVal"])/M["JunVal"]*100,np.nan)
    def tsig(r):
        return "Technical Buy" if (pd.notna(r["Dfh"]) and abs(r["Dfh"])<=2 and pd.notna(r["Wk"]) and r["Wk"]>0) else ""
    M["TechSig"]=M.apply(tsig,axis=1)
    return M, book

# =============================== CUP SCREEN ===================================
def cup_screen(M, scr):
    scr["key"]=scr.apply(lambda r:KEY(r.get("NSE Code"),r.get("Name")),axis=1); sI=scr.set_index("key")
    def gg(k,c):
        try: v=sI.loc[k,c]; return float(v) if (not isinstance(v,pd.Series) and pd.notna(v)) else np.nan
        except: return np.nan
    rows=[]
    for _,m in M[M.Class=="Equity"].iterrows():
        k=m["key"]; p=gg(k,"Current Price"); d50=gg(k,"DMA 50"); d200=gg(k,"DMA 200")
        dfh=gg(k,"Down from 52w high"); ufl=gg(k,"Up from 52w low"); rsi=gg(k,"RSI"); wk=gg(k,"Return over 1week"); m1=gg(k,"Return over 1month")
        if np.isnan(p) or np.isnan(dfh) or np.isnan(ufl): continue
        above=(pd.notna(d50) and pd.notna(d200) and p>d50 and p>d200)
        handle=(pd.notna(wk) and -6<=wk<=2); healthy=(pd.notna(rsi) and 50<=rsi<=70)
        tier=None
        if abs(dfh)<=6 and ufl>=40 and above and healthy and handle: tier="Breakout-ready"
        elif abs(dfh)<=12 and ufl>=30 and above and (pd.notna(rsi) and 48<=rsi<=72): tier="Cup forming"
        if not tier: continue
        rows.append(dict(Company=m["Company"],NSE=m["NSE"],Tier=tier,Price=p,DFH=dfh,UFL=ufl,RSI=rsi,
            MA=m["MA"],Wk=wk,M1=m1,Conv=m["Conv"],Grand=m["Grand"],TG=m["TG"],Verdict=m["Verdict"],Reco=m["Reco"]))
    cup=pd.DataFrame(rows)
    if len(cup):
        cup["to"]=cup["Tier"].map({"Breakout-ready":0,"Cup forming":1}); cup=cup.sort_values(["to","DFH"]).drop(columns="to")
    return cup

# =============================== WATCHLIST ====================================
def load_watchlist(path, cur, sc):
    if not path or not os.path.exists(path): return None
    wl=pd.read_csv(path); wl.columns=[c.strip() for c in wl.columns]
    if "Company" not in wl: wl["Company"]=wl["Symbol"]
    if "Last" not in wl: wl["Last"]=np.nan
    wl["key"]=wl["Symbol"].map(lambda s:KEY(s))
    sc["key"]=sc.apply(lambda r:KEY(r.get("NSE"),r["Name"]),axis=1)
    heldk=set(cur["key"]); sck=set(sc["key"])
    wl["Held"]=wl["key"].isin(heldk); wl["Scored"]=wl["key"].isin(sck)
    return wl

# =============================== MAIN (prep) ==================================
def main_prep():
    cur=consolidate(CFG["CURRENT_FILES"])
    base=load_baseline(CFG["LASTWEEK"])
    held_codes=[c for c in cur[cur.Class=="Equity"]["NSE"].tolist() if c]
    screeners=[CFG["SCREENER_HELD"], CFG["SCREENER_WATCH"]]
    sc=run_scorer(screeners, held_codes)
    scr=pd.read_csv("screener.csv"); scr.columns=[c.strip() for c in scr.columns]
    _dq=("Change in DII holding" in scr.columns)
    M, book=build_master(cur, base, sc, scr)
    cup=cup_screen(M, scr)
    wl=load_watchlist(CFG["WATCHLIST_CSV"], cur, sc)
    # write intermediates consumed by the build stage
    EXITED_JUN=base[~base["key"].isin(set(cur["key"]))]["JunVal"].sum()
    FULL_JUN=base["JunVal"].sum(); HELD_JUN=M["JunVal"].sum()
    cur.to_pickle("consN.pkl"); base.to_pickle("baseline.pkl"); M.to_pickle("master.pkl")
    cup.to_pickle("cup.pkl")
    if wl is not None: wl.to_pickle("wl2.pkl")
    json.dump(dict(book=float(book),EXITED_JUN=float(EXITED_JUN),FULL_JUN=float(FULL_JUN),jun_held=float(HELD_JUN),
        DATE_CUR=CFG["DATE_CUR"],DATE_PREV=CFG["DATE_PREV"],LBL_CUR=CFG["LBL_CUR"],LBL_PREV=CFG["LBL_PREV"],
        HAS_WL=(wl is not None), OUT=CFG["OUT"],
        FII_HDR="FII Δ Qtr %", DII_HDR=("DII Δ Qtr %" if _dq else "DII Δ 3Y %")), open("meta.json","w"))
    print(f"[prep] positions {len(M)} | book Rs{book/1e5:.1f}L | scored {(M.Verdict.notna()).sum()} | "
          f"cup {len(cup)} | watchlist {'yes' if wl is not None else 'no'}")

if __name__=="__main__":
    main_prep()
    # ---- build stage ----
    exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'weekly_build.py')).read())
