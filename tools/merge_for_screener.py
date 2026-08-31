#!/usr/bin/env python3
# ==============================================================================
#  MERGE HOLDINGS -> SCREENER IMPORT LIST
#  Reads your broker files from inputs/current/ , merges by ISIN across accounts,
#  and writes ONE ISIN-keyed CSV you can import into Screener.in to create a
#  watchlist (Screener > Watchlist > import from demat statement). Once the
#  watchlist exists, export it from Screener and use that as screener_holdings.csv
#  for the full weekly review.
#
#  RUN:  python merge_for_screener.py
#  Output: outputs/Holdings for Screener.csv
# ==============================================================================
import os, sys, glob, importlib.util
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE); os.chdir(ROOT)   # kit root (inputs/ outputs/ live here)

# reuse the loaders from weekly_review.py (import does not run its pipeline)
spec = importlib.util.spec_from_file_location("wr", os.path.join(ROOT, "engine", "weekly_review.py"))
wr = importlib.util.module_from_spec(spec); spec.loader.exec_module(wr)

files = sorted(glob.glob("inputs/current/*.xls") + glob.glob("inputs/current/*.csv") + glob.glob("inputs/current/*.xlsx"))
if not files:
    sys.exit("No broker files found in inputs/current/  — put this week's ICICI/Kite exports there.")

rows = []
for p in files:
    for r in wr.detect_and_load(p, holder_hint=os.path.basename(p)):
        rows.append(dict(r))
d = pd.DataFrame(rows)
d["Class"] = [wr.classify(r.Name, r.NSE) for r in d.itertuples()]

def agg(g):
    return pd.Series(dict(Name=max(g["Name"], key=len),
                          Symbol=next((x for x in g["NSE"] if x), ""),
                          Quantity=int(g["Qty"].sum()),
                          Class=g["Class"].iloc[0]))
m = d.groupby("ISIN", as_index=False).apply(agg, include_groups=False)

# Screener carries listed equities only — drop ETFs / gold / silver
eq = m[(m["Class"] == "Equity") & m["ISIN"].str.match(r"^IN[EF]", na=False)]
excl = m[m["Class"] != "Equity"]["Name"].tolist()

os.makedirs("outputs", exist_ok=True)
out = eq[["ISIN", "Symbol", "Name", "Quantity"]].sort_values("Name")
dest = os.path.join("outputs", "Holdings for Screener.csv")
out.to_csv(dest, index=False)
print(f"Wrote {len(out)} equity holdings -> {dest}")
print("Excluded (not on Screener):", ", ".join(excl) if excl else "none")
