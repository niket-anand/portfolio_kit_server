import pandas as pd, numpy as np
import sys
INPUT=sys.argv[1] if len(sys.argv)>1 else 'screener.csv'
OUTPUT=sys.argv[2] if len(sys.argv)>2 else 'scored.pkl'
df = pd.read_csv(INPUT)
df.columns=[c.strip() for c in df.columns]
G=lambda r,c: (r[c] if c in r and pd.notna(r[c]) else np.nan)

FINCO=['Bank','Financ','NBFC','Insurance','Capital Market','Holding']
def is_finco(ig,ind):
    s=f"{ig} {ind}".lower()
    return any(k.lower() in s for k in FINCO)

def stage1(r):
    reasons=[]
    pl=G(r,'Pledged percentage')
    if pd.notna(pl) and pl>5: reasons.append(f'Pledge {pl:.1f}%>5')
    debt=G(r,'Debt'); sales=G(r,'Sales')
    if not is_finco(r['Industry Group'],r['Industry']) and pd.notna(debt) and pd.notna(sales) and sales>0:
        if debt/sales>3.5: reasons.append(f'Debt/Sales {debt/sales:.1f}>3.5')
    opm=G(r,'OPM'); opm5=G(r,'OPM 5Year')
    if pd.notna(opm) and pd.notna(opm5) and opm<5 and opm5<8: reasons.append('OPM destruction')
    g3=G(r,'Sales growth 3Years'); g5=G(r,'Sales growth 5Years')
    if pd.notna(g3) and pd.notna(g5) and g3<0 and g5<0: reasons.append('Persistent decline')
    return reasons

def band(v,bands,default=0):
    if pd.isna(v): return default
    for t,p in bands:
        if v>=t: return p
    return default

def L1(r):
    g5=G(r,'Sales growth 5Years'); g3=G(r,'Sales growth 3Years'); ttm=G(r,'YOY Quarterly sales growth')
    sc=lambda v: band(v,[(20,4),(15,3.5),(10,3),(6,2),(3,1)], -1 if (pd.notna(v) and v<0) else 0)
    s5=sc(g5); s3=sc(g3)
    sttm=band(ttm,[(25,3),(15,2.5),(8,2),(3,1)], -1 if (pd.notna(ttm) and ttm<0) else 0)
    # consistency
    cons=0
    if pd.notna(g5) and pd.notna(g3) and pd.notna(ttm):
        if g5>10 and g3>10 and ttm>10: cons=2
        elif g5>8 and g3>8 and ttm>8: cons=1
    # acceleration
    acc=0
    if pd.notna(g5) and pd.notna(g3) and pd.notna(ttm):
        if ttm>g3>g5: acc=2
        elif ttm>g3: acc=1
        elif ttm<g3-15 and g3<g5-15: acc=-1
    return round(s5+s3+sttm+cons+acc,2)

def L2(r):
    pg=G(r,'YOY Quarterly profit growth'); sg=G(r,'YOY Quarterly sales growth')
    pgs=band(pg,[(25,3),(15,2.5),(8,2),(0,1)], -1 if (pd.notna(pg) and pg<-25) else 0)
    # operating leverage (spec: Profit>Sales*1.5 & >20% ->2 ; [reconstructed mild step])
    ol=0
    if pd.notna(pg) and pd.notna(sg):
        if pg>sg*1.5 and pg>20: ol=2
        elif pg>sg and pg>0: ol=1
    opm=G(r,'OPM'); ml=band(opm,[(25,2),(18,1.5),(12,1)],0)
    opm5=G(r,'OPM 5Year'); mt=0
    if pd.notna(opm) and pd.notna(opm5):
        d=opm-opm5
        mt=1 if d>=3 else (0.5 if d>=1 else (-0.5 if d<=-3 else 0))
    return round(pgs+ol+ml+mt,2)

def L3(r):
    roce=G(r,'Return on capital employed'); roce3=G(r,'Average return on capital employed 3Years'); roe=G(r,'Return on equity')
    # ROCE level /3  [band reconstructed; >60 capped at 2]
    rl=band(roce,[(30,3),(22,2.5),(17,2),(12,1.25),(8,0.75)],0)
    if pd.notna(roce) and roce>60: rl=min(rl,2)
    # ROCE 3Y durability /2 [>60 capped 1.25]
    rd=band(roce3,[(25,2),(18,1.5),(12,1),(8,0.5)],0)
    if pd.notna(roce3) and roce3>60: rd=min(rd,1.25)
    # ROE/ROCE quality /2
    q=0
    if pd.notna(roe) and pd.notna(roce) and roce!=0:
        ratio=roe/roce
        q=2 if 0.7<=ratio<=1.3 else (1 if 0.5<=ratio<=1.6 else 0.5)
    return round(rl+rd+q,2)

REPUTED=['tata','bajaj','3m india','siemens']
def L4(r):
    s=2.0
    pg=G(r,'YOY Quarterly profit growth'); g3=G(r,'Sales growth 3Years'); roce=G(r,'Return on capital employed')
    debt=G(r,'Debt'); sales=G(r,'Sales'); opm=G(r,'OPM'); opm5=G(r,'OPM 5Year'); fcf=G(r,'Free cash flow last year')
    ds=debt/sales if (pd.notna(debt) and pd.notna(sales) and sales>0) else np.nan
    # FCF quality score
    f=0
    f+= 2 if (pd.notna(pg) and pg>20) else 0
    f+= 2 if (pd.notna(g3) and g3>15) else 0
    f+= 2 if (pd.notna(roce) and roce>20) else 0
    f+= 1 if (pd.notna(ds) and ds<0.5) else 0
    f+= 1 if (pd.notna(opm) and pd.notna(opm5) and opm>opm5) else 0
    if pd.notna(fcf) and fcf>0:
        s+= 1.5 if f>=6 else (1.0 if f>=3 else 0.5)
    else:
        s+= 0.5 if f>=6 else (-0.5 if f>=3 else -1.5)
    # Debt quality score
    d=0
    d+= 2 if (pd.notna(roce) and roce>15) else 0
    d+= 1 if (pd.notna(pg) and pg>0) else 0
    d+= 1 if (pd.notna(ds) and ds<1.0) else 0
    d+= 2 if (pd.notna(roce) and roce>18 and pd.notna(pg) and pg>0) else 0
    if pd.notna(ds) and ds<0.1: s+=1.0
    elif d>=5: s+=0.5
    elif pd.notna(ds) and ds>2: s-=1.5
    elif d>=3: s-=0.5
    else: s-=0.5
    nm=str(r['Name']).lower()
    if any(k in nm for k in REPUTED): s+=0.5
    return round(max(0,min(5,s)),2)

MEGA1=['electrical','renewable','defen','solar','power','aerospace']
MEGA05=['pharma','auto','fintech','biotech']
def L5(r):
    roce3=G(r,'Average return on capital employed 3Years'); g5=G(r,'Sales growth 5Years')
    opm5=G(r,'OPM 5Year'); mcap=G(r,'Market Capitalization')
    rd=band(roce3,[(25,2),(18,1.5),(12,1),(8,0.5)],0)
    if pd.notna(roce3) and roce3>60: rd=min(rd,1.25)
    melt=-0.75 if (pd.notna(g5) and pd.notna(roce3) and g5<3 and roce3>60) else 0
    pw=band(opm5,[(25,1.5),(15,1),(8,0.5)],0)
    sc=0.5 if (pd.notna(mcap) and mcap>200000) else (0.25 if (pd.notna(mcap) and mcap>50000) else 0)
    s=f"{r['Industry Group']} {r['Industry']}".lower()
    mt=1 if any(k in s for k in MEGA1) else (0.5 if any(k in s for k in MEGA05) else 0)
    return round(max(0,min(5,rd+melt+pw+sc+mt)),2)

def L6(r):
    p=G(r,'Current Price'); d50=G(r,'DMA 50'); d200=G(r,'DMA 200'); rsi=G(r,'RSI')
    r1m=G(r,'Return over 1month'); dfh=G(r,'Down from 52w high')
    # P1 MA position
    if pd.notna(p) and pd.notna(d50) and pd.notna(d200):
        if p>d50 and p>d200 and d50>d200: p1=2
        elif p>d50 and p>d200: p1=1.5
        elif p>d50 or p>d200: p1=1
        elif d50<d200 and pd.notna(r1m) and r1m>0: p1=0.25
        else: p1=0.5
    else: p1=0.5
    # P2 momentum
    if pd.notna(rsi):
        if rsi>70: p2=0.5
        elif 50<=rsi<=70 and pd.notna(r1m) and r1m>0: p2=2
        elif 45<=rsi<70: p2=1.5
        elif rsi<40: p2=0.5
        else: p2=1
    else: p2=1
    # P3 52w position (Down from 52w high)
    if pd.notna(dfh):
        ad=abs(dfh)
        if 15<=ad<=30: p3=2
        elif ad<15: p3=1
        elif 30<ad<=45: p3=1
        else: p3=0.75
    else: p3=1
    # P4 trend quality
    if pd.notna(r1m) and pd.notna(p) and pd.notna(d50) and pd.notna(d200) and p>d50 and p>d200:
        p4=2 if r1m>15 else (1.5 if r1m>8 else 1)
    elif pd.notna(r1m) and r1m>0: p4=1
    else: p4=0.5
    return round(p1+p2+p3+p4,2)

def fgrade(f): return 'A+' if f>=35 else 'A' if f>=30 else 'B+' if f>=25 else 'B' if f>=20 else 'C' if f>=15 else 'D'
def tgrade(t): return 'Strong' if t>=6.5 else 'Healthy' if t>=5 else 'Mixed' if t>=3.5 else 'Weak' if t>=2 else 'Bearish'
def verdict(fg,tg):
    strong = tg in ('Strong','Healthy')
    if fg in ('A+','A'): return 'GREEN ZONE' if strong else 'VALUE OPPORTUNITY'
    if fg in ('B+','B'): return 'MOMENTUM PLAY' if strong else 'AVOID'
    return 'EXIT'

def conv(r,grand,l1,f):
    pe=G(r,'Price to Earning')
    growths=[G(r,'Sales growth 3Years'),G(r,'Sales growth 5Years'),G(r,'YOY Quarterly sales growth'),G(r,'YOY Quarterly profit growth')]
    growths=[g for g in growths if pd.notna(g) and g>0]
    mg=max(growths) if growths else np.nan
    peg=pe/mg if (pd.notna(pe) and pd.notna(mg) and mg>0) else np.nan
    val=band(peg,[],0)
    if pd.isna(peg): val=9
    else: val=25 if peg<0.8 else 22 if peg<1.2 else 18 if peg<1.8 else 14 if peg<2.5 else 9 if peg<3.5 else 5 if peg<5 else 2
    quality=grand/48*30
    gq=l1/15*20
    mc=G(r,'Market Capitalization')
    liq=15 if (pd.notna(mc) and mc>=500000) else 14 if mc>=200000 else 13 if mc>=100000 else 11 if mc>=50000 else 9 if mc>=20000 else 7 if mc>=10000 else 5 if mc>=5000 else 3
    dfh=abs(G(r,'Down from 52w high')) if pd.notna(G(r,'Down from 52w high')) else np.nan
    if pd.isna(dfh): mos=6
    else: mos=10 if 15<=dfh<=30 else 8 if 30<dfh<=45 else 7 if 8<=dfh<15 else 4 if dfh<8 else 6
    return round(quality+val+gq+liq+mos,1), (round(peg,2) if pd.notna(peg) else None)

def tier(c): return 'Tier 1' if c>=85 else 'Tier 2' if c>=70 else 'Tier 3' if c>=55 else 'Tier 4' if c>=40 else 'Skip'
def tierwt(c): return '8-10%' if c>=85 else '5-7%' if c>=70 else '3-5%' if c>=55 else '1-3%' if c>=40 else '0%'

rows=[]
for _,r in df.iterrows():
    rej=stage1(r)
    rec={'Name':r['Name'],'NSE':r.get('NSE Code'),'ISIN':r.get('ISIN Code')}
    if rej:
        rec.update({'Verdict':'EXIT (Stage 1)','Stage1':'; '.join(rej),'Fund':np.nan,'Tech':np.nan,'Grand':np.nan,'Conv':np.nan})
        rows.append(rec); continue
    l1,l2,l3,l4,l5,l6=L1(r),L2(r),L3(r),L4(r),L5(r),L6(r)
    fund=round(l1+l2+l3+l4+l5,2); grand=round(fund+l6,2)
    fg,tg=fgrade(fund),tgrade(l6); v=verdict(fg,tg)
    cs,peg=conv(r,grand,l1,fund)
    rec.update({'L1':l1,'L2':l2,'L3':l3,'L4':l4,'L5':l5,'L6':l6,'Fund':fund,'FGrade':fg,
                'Tech':l6,'TGrade':tg,'Grand':grand,'Verdict':v,'Conv':cs,'PEG':peg,
                'Tier':tier(cs),'Weight':tierwt(cs),'Stage1':''})
    rows.append(rec)
out=pd.DataFrame(rows)

# ---------- MOMENTUM (0-100) ----------
def gv(r,c):
    v=r[c] if (c in r and pd.notna(r[c])) else np.nan
    try: return float(v)
    except: return np.nan
srcI=df.set_index('Name')
def momentum(nm):
    if nm not in srcI.index: return np.nan
    r=srcI.loc[nm]
    p=gv(r,'Current Price');d50=gv(r,'DMA 50');d200=gv(r,'DMA 200');rsi=gv(r,'RSI')
    r1w=gv(r,'Return over 1week');r1m=gv(r,'Return over 1month');dfh=gv(r,'Down from 52w high')
    s=0
    if pd.notna(p) and pd.notna(d50) and pd.notna(d200):
        s+= 30 if (p>d50 and p>d200 and d50>d200) else 22 if (p>d50 and p>d200) else 12 if (p>d50 or p>d200) else 0
    if pd.notna(rsi): s+= 20 if 55<=rsi<=70 else 14 if (50<=rsi<55 or 70<rsi<=75) else 8 if 45<=rsi<50 else 6 if rsi>75 else 0
    if pd.notna(r1m): s+= 25 if r1m>15 else 20 if r1m>8 else 14 if r1m>3 else 8 if r1m>0 else 2
    if pd.notna(r1w): s+= 10 if r1w>5 else 7 if r1w>2 else 4 if r1w>0 else 0
    if pd.notna(dfh):
        ad=abs(dfh); s+= 15 if ad<10 else 12 if ad<20 else 8 if ad<35 else 3
    return round(s,1)
def ma_position(nm):
    if nm not in srcI.index: return ''
    r=srcI.loc[nm]; p=gv(r,'Current Price');d50=gv(r,'DMA 50');d200=gv(r,'DMA 200')
    if pd.notna(p) and pd.notna(d50) and pd.notna(d200):
        if p>d50 and p>d200: return '>50 & >200 DMA'
        if p<d50 and p<d200: return '<both DMA'
    return 'mixed DMA' if pd.notna(d50) else ''
out['Mom']=out['Name'].map(momentum)
out['MA Position']=out['Name'].map(ma_position)

# ---------- RECOMMENDATION (verdict x momentum x held) ----------
# held: optional newline/comma list of NSE codes via env HELD or file held.txt
import os
held=set()
hp=os.environ.get('HELD_FILE','held.txt')
if os.path.exists(hp):
    held={x.strip().upper() for x in open(hp).read().replace(',','\n').split() if x.strip()}
def is_held(nse): return str(nse).strip().upper() in held if held else False
def reco(v,m,h):
    if v=='EXIT (Stage 1)': return 'EXIT' if h else 'AVOID'
    if v=='EXIT': return 'TRIM/EXIT' if h else 'AVOID'
    if v=='AVOID': return 'TRIM' if h else 'AVOID'
    if v=='GREEN ZONE': return ('ADD' if h else 'BUY') if (pd.notna(m) and m>=70) else ('HOLD/ADD' if h else 'BUY') if (pd.notna(m) and m>=45) else ('HOLD' if h else 'ACCUMULATE (weak mom)')
    if v=='MOMENTUM PLAY': return ('HOLD (ride)' if h else 'BUY (momentum)') if (pd.notna(m) and m>=70) else ('HOLD' if h else 'WATCH') if (pd.notna(m) and m>=45) else ('TRIM (fading)' if h else 'WATCH')
    if v=='VALUE OPPORTUNITY': return 'HOLD (value)' if h else 'WATCH (value, weak mom)'
    return 'REVIEW'
out['Held']=out['NSE'].map(is_held)
out['Reco']=out.apply(lambda r: reco(r['Verdict'],r['Mom'],r['Held']),axis=1)

out.to_pickle(OUTPUT)
out.to_csv(OUTPUT.rsplit('.',1)[0]+'.csv',index=False)
print(f"Scored {len(out)} names | Stage-1 EXIT {(out['Verdict']=='EXIT (Stage 1)').sum()} | GREEN ZONE {(out['Verdict']=='GREEN ZONE').sum()}")
print("Output:",OUTPUT,"and",OUTPUT.rsplit('.',1)[0]+'.csv')
