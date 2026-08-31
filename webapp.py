#!/usr/bin/env python3
"""
Local WEB APP for the weekly kit — Python pipeline runs in a background thread.

    pip install flask
    python webapp.py
    # open http://127.0.0.1:8000 in your browser

Upload this week's files, set the dates, click Run. The review builds server-side
(background thread running run_weekly.py); the page polls for live progress and then
shows download links for everything in outputs/.
"""
import os, sys, glob, threading, subprocess, queue, datetime, mimetypes, secrets
from functools import wraps
from flask import Flask, request, jsonify, send_from_directory, Response

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
IN = "inputs"
OUT = "outputs"
for d in ["inputs/current", "inputs/lastweek", "inputs/trades", OUT]:
    os.makedirs(d, exist_ok=True)
app = Flask(__name__)

# --- auth: single-user HTTP Basic Auth, required on a server -------------
WEBAPP_USER = os.environ.get("WEBAPP_USER")
WEBAPP_PASS = os.environ.get("WEBAPP_PASS")
if not WEBAPP_USER or not WEBAPP_PASS:
    sys.exit("Set WEBAPP_USER and WEBAPP_PASS env vars before starting (see .env / docker-compose.yml). Refusing to start without auth on a server.")

def check_auth(u, p):
    return secrets.compare_digest(u, WEBAPP_USER) and secrets.compare_digest(p, WEBAPP_PASS)

@app.before_request
def _guard():
    if request.endpoint and request.endpoint != "static":
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return Response("Auth required", 401, {"WWW-Authenticate": 'Basic realm="Weekly Portfolio Kit"'})

JOB = {"running": False, "log": [], "done": False, "rc": None, "started": None}
LOCK = threading.Lock()

FOLDERS = {
    "current": "inputs/current",
    "lastweek": "inputs/lastweek",
    "trades": "inputs/trades",
    "screener_holdings": "inputs",
    "watchlist": "inputs",
    "screener_watchlist": "inputs",
}
SINGLE = {
    "screener_holdings": "screener_holdings.csv",
    "watchlist": "watchlist.csv",
    "screener_watchlist": "screener_watchlist.csv"
}

def input_status():
    def n(p):
        return len([f for f in glob.glob(p) if os.path.basename(f) != ".gitkeep"])
    return {
        "current (ICICI + Kite this week)": n("inputs/current/*"),
        "lastweek (for weekly returns)": n("inputs/lastweek/*"),
        "screener_holdings.csv": 1 if os.path.exists("inputs/screener_holdings.csv") else 0,
        "watchlist.csv (buy list)": 1 if os.path.exists("inputs/watchlist.csv") else 0,
        "trades (optional journal)": n("inputs/trades/*.csv"),
    }

def worker(args):
    with LOCK:
        JOB.update(running=True, log=[], done=False, rc=None, started=datetime.datetime.now().strftime("%H:%M:%S"))
    try:
        p = subprocess.Popen([sys.executable, "run_weekly.py", *args], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
        for line in p.stdout:
            with LOCK:
                JOB["log"].append(line.rstrip())
        p.wait()
        with LOCK:
            JOB.update(running=False, done=True, rc=p.returncode)
    except Exception as e:
        with LOCK:
            JOB["log"].append("ERROR: " + str(e))
            JOB.update(running=False, done=True, rc=1)

@app.route("/")
def index():
    return Response(PAGE, mimetype="text/html")

@app.route("/inputs")
def inputs():
    return jsonify(input_status())

@app.route("/upload", methods=["POST"])
def upload():
    tgt = request.form.get("target", "current")
    folder = FOLDERS.get(tgt, "inputs/current")
    files = request.files.getlist("files")
    saved = []
    for f in files:
        name = SINGLE[tgt] if tgt in SINGLE else os.path.basename(f.filename)
        f.save(os.path.join(folder, name))
        saved.append(name)
    return jsonify(saved=saved, status=input_status())

@app.route("/run", methods=["POST"])
def run():
    with LOCK:
        if JOB["running"]:
            return jsonify(error="already running"), 409
    d = request.get_json(force=True, silent=True) or {}
    args = [
        "--cur", d.get("cur", datetime.date.today().strftime("%d %B %Y")),
        "--prev", d.get("prev", (datetime.date.today() - datetime.timedelta(days=7)).strftime("%d %B %Y")),
        "--fii", d.get("fii", "n/a"),
        "--dii", d.get("dii", "n/a"),
        "--flowdate", d.get("flowdate", "latest close")
    ]
    threading.Thread(target=worker, args=(args,), daemon=True).start()
    return jsonify(ok=True)

@app.route("/status")
def status():
    with LOCK:
        j = dict(JOB)
    outs = [os.path.basename(f) for f in sorted(glob.glob(OUT + "/*")) if os.path.isfile(f) and os.path.basename(f) != ".gitkeep"]
    return jsonify(running=j["running"], done=j["done"], rc=j["rc"], started=j["started"], log="\n".join(j["log"]), outputs=outs)

@app.route("/download/<path:name>")
def download(name):
    return send_from_directory(OUT, name, as_attachment=True)

PAGE = """<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Weekly Portfolio Kit</title>
<style>
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:760px;margin:24px auto;padding:0 16px;color:#1a1a2e;background:#f4f6f9}
h1{font-size:22px;color:#0f2148}
h2{font-size:15px;color:#33415c;margin:22px 0 8px}
.card{background:#fff;border:1px solid #dbe2ea;border-radius:10px;padding:16px;margin-bottom:14px}
.st{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #eef1f5;font-size:14px}
.ok{color:#1a7f37;font-weight:600}
.no{color:#b3261e}
label{font-size:13px;color:#555;display:block;margin:8px 0 3px}
input,select{padding:7px 9px;border:1px solid #cdd6e0;border-radius:6px;font-size:14px;width:100%}
.row{display:flex;gap:10px;flex-wrap:wrap}
.row>div{flex:1;min-width:130px}
button{background:#1c3a6e;color:#fff;border:0;border-radius:7px;padding:11px 18px;font-size:15px;cursor:pointer;margin-top:10px}
button:disabled{background:#9aa6b8;cursor:not-allowed}
#log{background:#0f2148;color:#cfe0f5;font-family:ui-monospace,Menlo,monospace;font-size:12px;padding:12px;border-radius:8px;white-space:pre-wrap;max-height:260px;overflow:auto;margin-top:10px;display:none}
a.dl{display:block;padding:9px 12px;background:#eaf0f7;border:1px solid #d5deea;border-radius:7px;margin:6px 0;text-decoration:none;color:#0f2148;font-size:14px}
.hint{font-size:12px;color:#777;margin-top:4px}
</style>
</head>
<body>
<h1>Weekly Portfolio Review — Kit</h1>
<div class="card"><h2>1 · Inputs</h2><div id="status">loading…</div></div>
<div class="card"><h2>2 · Add files</h2>
<div class="row">
<div><label>Add to</label><select id="target">
<option value="current">current — this week's ICICI + Kite</option>
<option value="lastweek">lastweek — for weekly returns</option>
<option value="screener_holdings">screener_holdings.csv — held export</option>
<option value="watchlist">watchlist.csv — TradingView buy list</option>
<option value="screener_watchlist">screener_watchlist.csv — optional</option>
<option value="trades">trades — order/tradebooks (optional)</option>
</select></div>
<div><label>Files</label><input type="file" id="files" multiple></div>
</div>
<button onclick="up()">Upload</button><div class="hint" id="uphint"></div>
</div>
<div class="card"><h2>3 · Dates &amp; flows</h2>
<div class="row">
<div><label>This week</label><input id="cur" placeholder="22 August 2026"></div>
<div><label>Last week</label><input id="prev" placeholder="25 July"></div>
</div>
<div class="row">
<div><label>FII net (Cr)</label><input id="fii" placeholder="+1,181.66"></div>
<div><label>DII net (Cr)</label><input id="dii" placeholder="+2,493.41"></div>
<div><label>Flow date</label><input id="flowdate" placeholder="24 Aug 2026 close"></div>
</div>
</div>
<div class="card"><h2>4 · Run</h2>
<button id="runbtn" onclick="run()">▶ Run weekly review</button>
<pre id="log"></pre>
</div>
<div class="card" id="dlcard" style="display:none"><h2>5 · Downloads</h2><div id="downloads"></div></div>
<script>
async function refresh(){
  try {
    let res = await fetch('/inputs');
    let s = await res.json();
    document.getElementById('status').innerHTML = Object.entries(s).map(function(item){
      var k = item[0], v = item[1];
      return '<div class="st"><span>' + k + '</span><span class="' + (v ? 'ok' : 'no') + '">' + (v ? '✓ ' + v : '—') + '</span></div>';
    }).join('');
  } catch(e) {
    document.getElementById('status').innerHTML = '<span class="no">Error loading inputs: ' + e + '</span>';
  }
}
async function up(){
  let fd = new FormData();
  fd.append('target', document.getElementById('target').value);
  let fileInput = document.getElementById('files');
  for(let f of fileInput.files) fd.append('files', f);
  let uphint = document.getElementById('uphint');
  if(!fileInput.files.length){ uphint.textContent = 'pick file(s) first'; return; }
  uphint.textContent = 'uploading…';
  let res = await fetch('/upload', {method: 'POST', body: fd});
  let r = await res.json();
  uphint.textContent = 'added: ' + (r.saved ? r.saved.join(', ') : 'done');
  fileInput.value = '';
  refresh();
}
let timer;
async function run(){
  let runbtn = document.getElementById('runbtn');
  let log = document.getElementById('log');
  let dlcard = document.getElementById('dlcard');
  runbtn.disabled = true;
  log.style.display = 'block';
  log.textContent = 'starting…';
  dlcard.style.display = 'none';
  let body = {
    cur: document.getElementById('cur').value,
    prev: document.getElementById('prev').value,
    fii: document.getElementById('fii').value || 'n/a',
    dii: document.getElementById('dii').value || 'n/a',
    flowdate: document.getElementById('flowdate').value || 'latest close'
  };
  let r = await fetch('/run', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)});
  if(r.status == 409){ log.textContent = 'already running…'; }
  timer = setInterval(poll, 1000);
}
async function poll(){
  let runbtn = document.getElementById('runbtn');
  let log = document.getElementById('log');
  let dlcard = document.getElementById('dlcard');
  let downloads = document.getElementById('downloads');
  let res = await fetch('/status');
  let s = await res.json();
  log.textContent = s.log || 'working…';
  log.scrollTop = log.scrollHeight;
  if(s.done){
    clearInterval(timer);
    runbtn.disabled = false;
    log.textContent += (s.rc === 0 ? '\\n\\n✔ done' : ('\\n\\n⚠ finished with errors (rc=' + s.rc + ')'));
    if(s.outputs && s.outputs.length){
      dlcard.style.display = 'block';
      downloads.innerHTML = s.outputs.map(function(o){
        return '<a class="dl" href="/download/' + encodeURIComponent(o) + '">⬇ ' + o + '</a>';
      }).join('');
    }
  }
}
refresh();
</script>
</body>
</html>"""

if __name__ == "__main__":
    print("Weekly Portfolio Kit (DEV) — open http://127.0.0.1:8000")
    app.run(host="127.0.0.1", port=8000, debug=False)
