# Weekly Portfolio Kit — run it yourself, no Claude needed

This kit reproduces the entire weekly workflow we built: merge holdings for Screener,
the full CONVICTION-6 weekly review (6 sheets, your exact layout), the trade journal,
the P&L × trend grid, and the institutional-ownership triggers infographic — from one
command.

## What you get each run (in `outputs/`)
1. **Holdings for Screener.csv** — all accounts merged by ISIN, ready to import into Screener.
2. **Weekly Portfolio Review <date>.xlsx** — 6 sheets: Portfolio + Factors, Equity Master,
   Industry (inline subtotals), TradingView Watchlist, P&L × Trend Grid, Grid — Classification.
   Technicals-before-fundamentals, full-date Qty headers, no merged cells, Excel-clean.
3. **Trade Journal.xlsx** — from broker trade files (if provided).
4. **Portfolio Triggers Update.html** — FII/DII ownership triggers from the Screener export.

## One-time setup
```
pip install -r requirements.txt
```
Also install **LibreOffice** (free) — used to make Excel files open cleanly and to render the
triggers PDF. macOS: `brew install --cask libreoffice` · Ubuntu: `sudo apt install libreoffice`
· Windows: download from libreoffice.org. (If LibreOffice isn't present the workbook is still
produced; it just skips the final normalize step.)

## Each week: drop your files in, then run one command
Put this week's exports here:
```
inputs/current/        this week's broker files  (ICICI .xls + Kite/Zerodha holdings .csv)
inputs/lastweek/       last week's broker files  (for week-over-week deltas)
inputs/screener_holdings.csv     Screener export of your HELD names (RSI + FII/DII columns on)
inputs/watchlist.csv             OPTIONAL — your watchlist as  Group,Symbol,Company,Last
inputs/screener_watchlist.csv    OPTIONAL — Screener export covering watchlist names
inputs/trades/                   OPTIONAL — ICICI orderBook_*.csv + Zerodha tradebook-*.csv
```
Then:
```
python run_weekly.py --cur "22 August 2026" --prev "25 July"
```
(Dates default to today / 7 days ago if omitted. Add daily flows for the triggers panel:
`--fii "+1,181.66" --dii "+2,493.41" --flowdate "24 Aug 2026 close"`.)

## Web app — Python runs in the background (recommended, no command line)
```
pip install flask
python webapp.py
```
Then open **http://127.0.0.1:8000** in your browser. You:
1. Upload this week's files (pick the target folder in the dropdown),
2. Set the dates (and optional FII/DII flows),
3. Click **Run**.
The pipeline runs in a **background thread** on your machine; the page shows live progress
and then download links for everything in `outputs/`. Nothing leaves your computer — the
server is local (127.0.0.1).

(There's also a Streamlit version — `streamlit run app.py` — if you prefer it.)

## What is fully automated vs. what still needs you
**Fully automated (this kit):** the merge, the whole review with all layout rules, week-over-week
reconciliation, the journal, the grid, and the FII/DII ownership triggers. The ISIN-based join
means ICICI's abbreviated tickers (e.g. TDPSYS) now match Screener automatically, and
% from 52-week-high is computed from the price when Screener omits the column.

**Still needs a human (by design):**
- **TradingView watchlist from *screenshots*** — a picture can't be parsed reliably. **Fix: export
  your TradingView watchlist as CSV** (right-click list → Export) and save it as
  `inputs/watchlist.csv` with columns `Group,Symbol,Company,Last`. That removes the only manual
  transcription step and makes the watchlist fully automated too.
- **News-flow triggers** (order wins, results, USFDA, capex) — these need live news research and
  judgement; the kit does the *ownership* triggers from data, not the news ones.

## Automate the schedule (optional)
- **macOS/Linux (cron):** run every Saturday 9am →
  `0 9 * * 6 cd /path/to/portfolio-kit && /usr/bin/python3 run_weekly.py >> outputs/run.log 2>&1`
- **Windows (Task Scheduler):** New Task → Trigger: Weekly Sat 9am → Action: `python`,
  arguments `run_weekly.py`, Start-in `C:\path\to\portfolio-kit`.
You still drop the week's exports into `inputs/` first; the schedule just runs the build.

## Folder map
```
run_weekly.py         one-command orchestrator
app.py                optional Streamlit UI
engine/               weekly_review.py, weekly_build.py, conviction6.py (scorer),
                      finalize.py, make_filterable.py  (layout + scoring)
tools/                merge_for_screener.py, trade_journal.py, pnl_grid.py, triggers.py
data/symbols.csv      NSE symbol / ISIN reference
inputs/  outputs/     your files in, deliverables out
```

## Tuning
- CONVICTION-6 thresholds live in `engine/conviction6.py`.
- Layout (colours, column order, sheet set) in `engine/weekly_build.py` + `engine/finalize.py`.
- ICICI DP-Id → person mapping for the journal: edit `HOLDER_MAP` in `tools/trade_journal.py`.

_Not investment advice. A rules-based synthesis of your own data._
