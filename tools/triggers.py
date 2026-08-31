#!/usr/bin/env python3
"""Ownership-triggers infographic (HTML+PDF) from a Screener held export with
FII/DII holding + quarterly change columns.
Usage: python tools/triggers.py inputs/screener_holdings.csv "24 Aug 2026" "<FII_net>" "<DII_net>" """
import sys, pandas as pd
CSV=sys.argv[1]; FLOWDATE=sys.argv[2] if len(sys.argv)>2 else "latest close"
FII_DAY=sys.argv[3] if len(sys.argv)>3 else "n/a"; DII_DAY=sys.argv[4] if len(sys.argv)>4 else "n/a"
d=pd.read_csv(CSV); d.columns=[c.strip() for c in d.columns]; cols={c.lower():c for c in d.columns}
def col(*n):
    for x in n:
        if x.lower() in cols: return cols[x.lower()]
fh,fc,dh,dc=col("FII holding"),col("Change in FII holding"),col("DII holding"),col("Change in DII holding")
for k in (fh,fc,dh,dc): d[k]=pd.to_numeric(d[k],errors="coerce")
d=d.dropna(subset=[fc,dc]); d["Name"]=d["Name"].astype(str)
pos=d[(d[fc]>0)&(d[dc]>0)].assign(m=d[fc]+d[dc]).sort_values("m",ascending=False).head(8)
neg=d[(d[fc]<0)&(d[dc]<0)].assign(m=d[fc]+d[dc]).sort_values("m").head(6)
wat=d[((d[fc]>0)&(d[dc]<0))|((d[fc]<0)&(d[dc]>0))].assign(m=d[fc].abs()+d[dc].abs()).sort_values("m",ascending=False).head(6)
def imp(f,x):
    m=abs(f)+abs(x); return "High" if m>=3 else "Medium" if m>=1 else "Low"
def ic(i): return {"High":"#1a7f37","Medium":"#c47f17","Low":"#8a8a8a"}[i]
def arrow(c): return "▲" if c>0 else "▼"
def ac(c): return "#1a7f37" if c>0 else "#c02626"
def cell(r):
    f=r[fc]; x=r[dc]
    return (f'<div><b>FII {r[fh]:.2f}%</b> <span style="color:{ac(f)}">{arrow(f)} ({f:+.2f})</span></div>'
            f'<div><b>DII {r[dh]:.2f}%</b> <span style="color:{ac(x)}">{arrow(x)} ({x:+.2f})</span></div>'
            f'<div style="color:#777;font-size:10px">{r[fh]-f:.2f}→{r[fh]:.2f} · {r[dh]-x:.2f}→{r[dh]:.2f}</div>')
def sig(r,kind):
    f=r[fc]; x=r[dc]
    if kind=="pos": return f"Institutional accumulation — FII {f:+.2f}pp, DII {x:+.2f}pp."
    if kind=="neg": return f"Institutional distribution — FII {f:+.2f}pp, DII {x:+.2f}pp."
    return f"Divergent — FII {f:+.2f}pp vs DII {x:+.2f}pp (ownership rotating)."
def rows(df,kind):
    o=""
    for _,r in df.iterrows():
        i=imp(r[fc],r[dc])
        o+=(f'<tr><td class="stk">{r["Name"]}</td><td class="own">{cell(r)}</td>'
            f'<td class="why">{sig(r,kind)}</td><td class="imp" style="color:{ic(i)}"><b>{i}</b></td>'
            f'<td class="src">Screener.in<br><span style="color:#999">Jun-2026 qtr</span></td></tr>')
    return o
H=f"""<!DOCTYPE html><html><head><meta charset="utf-8"><style>
*{{box-sizing:border-box;margin:0;padding:0}}body{{font-family:'Segoe UI',Arial,sans-serif;background:#eef1f4;padding:18px}}
.page{{max-width:960px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,.12)}}
.hdr{{background:linear-gradient(135deg,#0f2148,#1c3a6e);color:#fff;text-align:center;padding:18px}}.hdr h1{{font-size:25px}}.hdr p{{font-size:13px;opacity:.9;margin-top:4px}}
.bar{{background:#eaf0f7;border-bottom:1px solid #d5deea;padding:8px 16px;font-size:12px;color:#33415c;display:flex;justify-content:space-between}}
.sec{{color:#fff;font-weight:700;font-size:15px;padding:9px 16px}}.green{{background:#1a7f37}}.red{{background:#b3261e}}.amber{{background:#c47f17}}
table{{width:100%;border-collapse:collapse;font-size:12px}}th{{background:#f2f5f9;color:#33415c;text-align:left;padding:7px 10px;font-size:11px;border-bottom:2px solid #d5deea}}
td{{padding:8px 10px;border-bottom:1px solid #eceff3;vertical-align:top}}.stk{{font-weight:700;color:#0f2148;width:15%}}.own{{width:27%;font-size:11px}}.why{{width:38%;color:#333}}.imp{{width:8%;text-align:center}}.src{{width:12%;font-size:10px}}
.flows{{display:flex;gap:12px;padding:14px 16px;background:#f7f9fc;border-top:2px solid #d5deea}}.flowbox{{flex:1;border:1px solid #d5deea;border-radius:6px;padding:10px;text-align:center;background:#fff}}.flowbox .v{{font-size:19px;font-weight:800}}
.note{{padding:12px 16px;font-size:11px;color:#555;line-height:1.5}}.disc{{background:#0f2148;color:#cdd7e6;font-size:10.5px;text-align:center;padding:10px 16px;line-height:1.5}}</style></head><body><div class="page">
<div class="hdr"><h1>PORTFOLIO — INSTITUTIONAL OWNERSHIP TRIGGERS</h1><p>FII / DII shareholding shifts across holdings · Jun 2026 vs Mar 2026</p></div>
<div class="bar"><span>Source: Screener.in shareholding · NSE provisional flows</span><span>Flows as of {FLOWDATE}</span></div>
<div class="sec green">● INSTITUTIONAL ACCUMULATION — both FII &amp; DII increased</div>
<table><tr><th>Stock</th><th>FII / DII Change (QoQ)</th><th>What it signals</th><th>Importance</th><th>Source</th></tr>{rows(pos,'pos')}</table>
<div class="sec red">● INSTITUTIONAL DISTRIBUTION — both FII &amp; DII reduced</div>
<table><tr><th>Stock</th><th>FII / DII Change (QoQ)</th><th>What it signals</th><th>Importance</th><th>Source</th></tr>{rows(neg,'neg')}</table>
<div class="sec amber">● WATCH — divergent flows</div>
<table><tr><th>Stock</th><th>FII / DII Change (QoQ)</th><th>What it signals</th><th>Importance</th><th>Source</th></tr>{rows(wat,'wat')}</table>
<div class="flows"><div class="flowbox"><div style="font-size:11px;color:#555">FII/FPI (Net)</div><div class="v" style="color:{ac(1 if not FII_DAY.startswith('-') else -1)}">₹{FII_DAY} Cr</div></div>
<div class="flowbox"><div style="font-size:11px;color:#555">DII (Net)</div><div class="v" style="color:{ac(1 if not DII_DAY.startswith('-') else -1)}">₹{DII_DAY} Cr</div></div>
<div class="flowbox"><div style="font-size:11px;color:#555">As of</div><div style="font-size:13px;margin-top:6px">{FLOWDATE}</div></div></div>
<div class="note"><b>How to read:</b> triggers are the change in FII &amp; DII shareholding (Mar→Jun 2026 quarter). Both-up=accumulation, both-down=distribution, divergent=ownership rotating. Not a live news scan. ▲ up ▼ down · pp=percentage points.</div>
<div class="disc">Informational only, not investment advice. Ownership data is quarterly (shareholding pattern). Do your own research.</div></div></body></html>"""
out=CSV.rsplit("/",1)[0].replace("inputs","outputs") if "inputs" in CSV else "outputs"
import os; os.makedirs("outputs",exist_ok=True)
open("outputs/Portfolio Triggers Update.html","w",encoding="utf-8").write(H)
print("triggers -> outputs/Portfolio Triggers Update.html")
