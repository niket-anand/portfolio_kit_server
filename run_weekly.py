#!/usr/bin/env python3
"""
ONE-COMMAND WEEKLY PORTFOLIO REVIEW  —  runs the whole pipeline, no Claude needed.

    python run_weekly.py --cur "22 August 2026" --prev "25 July"

What it does, using whatever you've dropped into inputs/:
  1. merge_for_screener  -> outputs/Holdings for Screener.csv   (from inputs/current/*)
  2. weekly_review       -> 4-sheet workbook (needs inputs/current, inputs/lastweek, inputs/screener_holdings.csv;
                            watchlist.csv + screener_watchlist.csv optional -> TradingView sheet)
  3. finalize            -> layout: technicals-before-fundamentals, inline Industry subtotals,
                            full-date Qty headers, LibreOffice-normalized (opens clean in Excel)
  4. pnl_grid            -> adds P&L × Weekly-Trend grid + classification sheets
  5. trade_journal       -> outputs/Trade Journal.xlsx     (if inputs/trades/* present)
  6. triggers            -> outputs/Portfolio Triggers Update.html (from FII/DII in the screener)

Dates default to today / last week if not given.  LBL are auto ("22-Aug","25-Jul").
"""
import os, sys, subprocess, argparse, datetime, glob
HERE=os.path.dirname(os.path.abspath(__file__)); os.chdir(HERE)
ENG="engine"; TOOL="tools"
def L(d):  # "22 August 2026" -> "22-Aug"
    for f in ("%d %B %Y","%d %b %Y","%d %B","%d %b"):
        try: return datetime.datetime.strptime(d,f).strftime("%d-%b")
        except: pass
    return d
ap=argparse.ArgumentParser()
ap.add_argument("--cur",default=datetime.date.today().strftime("%d %B %Y"))
ap.add_argument("--prev",default=(datetime.date.today()-datetime.timedelta(days=7)).strftime("%d %B %Y"))
ap.add_argument("--fii",default="n/a"); ap.add_argument("--dii",default="n/a"); ap.add_argument("--flowdate",default="latest close")
a=ap.parse_args()
LBL_CUR=L(a.cur); LBL_PREV=L(a.prev)
env=dict(os.environ, WR_DATE_CUR=a.cur, WR_LBL_CUR=LBL_CUR, WR_DATE_PREV=a.prev, WR_LBL_PREV=LBL_PREV)
def run(cmd,**kw): print("  $",*cmd); return subprocess.run(cmd,check=True,**kw)
os.makedirs("outputs",exist_ok=True)
REVIEW=f"outputs/Weekly Portfolio Review {LBL_CUR}.xlsx"

print("\n[1/6] merge holdings for Screener")
try: run([sys.executable,f"{TOOL}/merge_for_screener.py"],env=env)
except Exception as e: print("   (skip merge:",e,")")

print("\n[2/6] weekly review")
run([sys.executable,f"{ENG}/weekly_review.py"],env=env)
raw="outputs/Weekly Portfolio Review.xlsx"

print("\n[3/6] finalize layout")
run([sys.executable,f"{ENG}/finalize.py",raw,REVIEW,LBL_PREV,LBL_CUR])

print("\n[4/6] P&L × Trend grid")
run([sys.executable,f"{TOOL}/pnl_grid.py",REVIEW])

print("\n[5/6] trade journal")
if glob.glob("inputs/trades/*.csv"):
    try: run([sys.executable,f"{TOOL}/trade_journal.py"])
    except Exception as e: print("   (skip journal:",e,")")
else: print("   (no inputs/trades/*.csv — skipped)")

print("\n[6/6] ownership triggers")
sh="inputs/screener_holdings.csv"
if os.path.exists(sh):
    try: run([sys.executable,f"{TOOL}/triggers.py",sh,a.flowdate,a.fii,a.dii])
    except Exception as e: print("   (skip triggers:",e,")")

print("\nDONE. See outputs/ :"); [print("   -",os.path.basename(f)) for f in sorted(glob.glob("outputs/*")) if os.path.isfile(f)]
