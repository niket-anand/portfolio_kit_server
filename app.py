#!/usr/bin/env python3
"""Optional web UI for the weekly kit.  Run:  streamlit run app.py"""
import streamlit as st, os, glob, subprocess, sys, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
st.set_page_config(page_title="Weekly Portfolio Kit", layout="centered")
st.title("Weekly Portfolio Review — Kit")
st.caption("Drop this week's files into the inputs/ folders, set the dates, and run.")

def status(path, label):
    n = len(glob.glob(path))
    st.write(("✅" if n else "⬜") + f" {label}: {n} file(s)")

st.subheader("1 · Inputs")
status("inputs/current/*", "Current broker files (inputs/current/)")
status("inputs/lastweek/*", "Last-week broker files (inputs/lastweek/)")
status("inputs/screener_holdings.csv", "Screener held export")
status("inputs/watchlist.csv", "Watchlist CSV (optional)")
status("inputs/trades/*.csv", "Trade files (optional)")

st.info("Upload files below to add them to the inputs folders.")
tgt = st.selectbox("Add uploaded files to:", ["inputs/current", "inputs/lastweek", "inputs/trades", "inputs (screener_holdings.csv / watchlist.csv)"])
ups = st.file_uploader("Upload", accept_multiple_files=True)

if ups:
    for u in ups:
        d = "inputs" if tgt.startswith("inputs (") else tgt
        os.makedirs(d, exist_ok=True)
        open(os.path.join(d, u.name), "wb").write(u.getbuffer())
    st.success(f"Saved {len(ups)} file(s). Reload to refresh status.")

st.subheader("2 · Dates & flows")
c1, c2 = st.columns(2)
cur = c1.text_input("This week", "%s" % datetime.date.today().strftime("%d %B %Y"))
prev = c2.text_input("Last week", (datetime.date.today() - datetime.timedelta(days=7)).strftime("%d %B %Y"))

c3, c4, c5 = st.columns(3)
fii = c3.text_input("FII net (Cr)", "n/a")
dii = c4.text_input("DII net (Cr)", "n/a")
fd = c5.text_input("Flow date", "latest close")

st.subheader("3 · Run")
if st.button("▶ Run weekly review", type="primary"):
    with st.spinner("Building… (~1–2 min)"):
        r = subprocess.run([sys.executable, "run_weekly.py", "--cur", cur, "--prev", prev, "--fii", fii, "--dii", dii, "--flowdate", fd], capture_output=True, text=True)
    st.code(r.stdout[-3000:] or r.stderr[-3000:])
    st.success("Done" if r.returncode == 0 else "Finished with errors — see log above")

st.subheader("4 · Downloads")
for f in sorted(glob.glob("outputs/*")):
    if os.path.isfile(f):
        st.download_button("⬇ " + os.path.basename(f), open(f, "rb").read(), file_name=os.path.basename(f))
