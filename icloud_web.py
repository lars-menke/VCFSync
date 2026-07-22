"""
iCloud Kontakte Sync - Web-Oberfläche
=====================================
Eine schlichte lokale Web-App (Flask), die Export und Import per Knopfdruck
bedienbar macht. Nutzt dieselbe, bewährte Kernlogik wie das CLI-Skript
(icloud_contacts.py) - hier gibt es keine doppelte Sync-Logik.

Start:
    pip install -r requirements.txt
    python icloud_web.py
    -> Browser öffnen: http://127.0.0.1:5000

Zugangsdaten kommen wie beim CLI aus der .env-Datei oder aus
Umgebungsvariablen (ICLOUD_USER / ICLOUD_PASS).

Hinweis: Die App hat bewusst keine Anmeldung und ist nur für die lokale,
persönliche Nutzung gedacht (localhost bzw. privater Codespace-Port).
"""

import os
import re
import threading
from pathlib import Path
from types import SimpleNamespace

import requests
from flask import (Flask, Response, jsonify, render_template_string, request,
                   send_file)
from requests.auth import HTTPBasicAuth

import icloud_contacts as ic

app = Flask(__name__)

DATA_DIR = Path("web_data")
DATA_DIR.mkdir(exist_ok=True)
EXPORT_VCF = DATA_DIR / "export.vcf"
EXPORT_XLSX = DATA_DIR / "kontakte.xlsx"
IMPORT_VCF = DATA_DIR / "import.vcf"

# Gemeinsamer Job-Status (nur eine Operation gleichzeitig - reicht für einen Nutzer)
JOB = {"kind": None, "running": False, "done": 0, "total": 0,
       "message": "", "result": None, "error": None}
JOB_LOCK = threading.Lock()


def _job_reset(kind):
    with JOB_LOCK:
        JOB.update({"kind": kind, "running": True, "done": 0, "total": 0,
                    "message": "Starte ...", "result": None, "error": None})


def _job_progress(done, total, message):
    with JOB_LOCK:
        JOB["done"], JOB["total"], JOB["message"] = done, total, message


def _job_finish(result=None, error=None):
    with JOB_LOCK:
        JOB["running"] = False
        JOB["result"] = result
        JOB["error"] = error


def build_session():
    """Baut eine authentifizierte Session + ermittelt Adressbücher.
    Wirft eine aussagekräftige Exception, wenn Zugangsdaten fehlen."""
    ic.load_env()
    user = os.environ.get("ICLOUD_USER")
    pw = os.environ.get("ICLOUD_PASS")
    if not user or not pw:
        raise RuntimeError(
            "Zugangsdaten fehlen. ICLOUD_USER und ICLOUD_PASS in der "
            ".env-Datei setzen (oder als Umgebungsvariablen).")
    session = requests.Session()
    session.auth = HTTPBasicAuth(user, pw)
    session.headers.update({"User-Agent": "iCloudContactsScript/1.0"})
    principal = ic.get_principal_url(session, user)
    books = ic.get_all_addressbook_urls(session, principal)
    return session, books, books[0][0]


def _uid_name_map(vcards):
    """UID -> Anzeigename, um in der Ergebnisliste echte Namen zu zeigen."""
    out = {}
    for v in vcards:
        uid = ic.extract_uid(v)
        m = re.search(r"^FN:(.+)$", "\n".join(ic.unfold_vcard(v)), re.MULTILINE)
        if uid:
            out[uid] = m.group(1).strip() if m else uid
    return out


# ---------------------------------------------------------------------------
# Hintergrund-Aufgaben
# ---------------------------------------------------------------------------

def _run_export():
    try:
        session, books, _ = build_session()
        _job_progress(0, 0, "Lese Kontaktliste aus iCloud ...")

        def progress(i, total):
            _job_progress(i, total, f"Lade Kontakt {i}/{total} ...")

        result = ic.export_contacts(session, books, embed=True, progress=progress)
        EXPORT_VCF.write_text(result["text"], encoding="utf-8")
        _job_finish(result={
            "loaded": result["loaded"],
            "photos": result["photos"],
            "warnings": result["warnings"],
        })
    except Exception as e:  # noqa: BLE001 - Fehler soll in der UI landen
        _job_finish(error=str(e))


def _run_import(dry_run):
    try:
        if not IMPORT_VCF.exists():
            raise RuntimeError("Keine Import-Datei vorhanden. Bitte erst hochladen.")
        vcards = ic.split_vcards(IMPORT_VCF.read_text(encoding="utf-8"))
        names = _uid_name_map(vcards)

        session, books, primary = build_session()
        _job_progress(0, len(vcards), "Lese vorhandene Kontakte aus iCloud (UID-Abgleich) ...")

        def progress(i, total, message):
            _job_progress(i, total, f"Verarbeite {i}/{total} ...")

        result = ic.import_vcards(session, books, primary, vcards,
                                  dry_run=dry_run, progress=progress)

        # "Neu"- und Fehlerzeilen mit echten Namen anreichern (die interessieren)
        def enrich(line):
            m = re.search(r":\s*([0-9a-fA-F-]{6,})", line)
            if m:
                for uid, name in names.items():
                    if uid.startswith(m.group(1)) or m.group(1) in uid:
                        return f"{line}  ({name})"
            return line

        highlights = [enrich(l) for l in result["log"]
                      if any(k in l for k in ("NEU ANLEGEN", "NEU ANGELEGT", "LÖSCHEN",
                                              "GELÖSCHT", "FEHLER", "Hinweis"))]

        _job_finish(result={
            "dry_run": dry_run,
            "updated": result["updated"],
            "created": result["created"],
            "deleted": result["deleted"],
            "errors": result["errors"],
            "total": result["total"],
            "highlights": highlights,
        })
    except Exception as e:  # noqa: BLE001
        _job_finish(error=str(e))


def _start(kind, target, *args):
    with JOB_LOCK:
        if JOB["running"]:
            return False
    _job_reset(kind)
    threading.Thread(target=target, args=args, daemon=True).start()
    return True


# ---------------------------------------------------------------------------
# Routen
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    ic.load_env()
    creds_ok = bool(os.environ.get("ICLOUD_USER") and os.environ.get("ICLOUD_PASS"))
    user = os.environ.get("ICLOUD_USER", "")
    return render_template_string(PAGE, creds_ok=creds_ok, user=user)


@app.route("/api/status")
def api_status():
    with JOB_LOCK:
        return jsonify(dict(JOB))


@app.route("/api/export", methods=["POST"])
def api_export():
    if not _start("export", _run_export):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


@app.route("/api/download/vcf")
def api_download_vcf():
    if not EXPORT_VCF.exists():
        return "Noch kein Export vorhanden.", 404
    return send_file(EXPORT_VCF, as_attachment=True, download_name="icloud_kontakte.vcf")


@app.route("/api/download/xlsx")
def api_download_xlsx():
    if not EXPORT_VCF.exists():
        return "Noch kein Export vorhanden.", 404
    ic.cmd_to_excel(SimpleNamespace(input=str(EXPORT_VCF), output=str(EXPORT_XLSX)))
    return send_file(EXPORT_XLSX, as_attachment=True, download_name="kontakte.xlsx")


@app.route("/api/upload", methods=["POST"])
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "Keine Datei ausgewählt."}), 400
    name = f.filename.lower()

    try:
        if name.endswith(".xlsx"):
            up = DATA_DIR / "upload.xlsx"
            f.save(up)
            ic.cmd_from_excel(SimpleNamespace(input=str(up), output=str(IMPORT_VCF)))
        elif name.endswith(".vcf"):
            data = f.read().decode("utf-8", errors="replace")
            IMPORT_VCF.write_text(data, encoding="utf-8")
        else:
            return jsonify({"error": "Bitte eine .vcf- oder .xlsx-Datei hochladen."}), 400
    except SystemExit as e:  # cmd_from_excel bricht bei fehlenden Spalten ab
        return jsonify({"error": f"Datei konnte nicht gelesen werden (Code {e.code}). "
                                 "Stammt die Excel-Datei aus 'Als Excel herunterladen'?"}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": str(e)}), 400

    count = IMPORT_VCF.read_text(encoding="utf-8").count("BEGIN:VCARD")
    return jsonify({"count": count})


@app.route("/api/import", methods=["POST"])
def api_import():
    dry_run = bool(request.json and request.json.get("dry_run"))
    if not _start("import", _run_import, dry_run):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


PAGE = """
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>iCloud Kontakte Sync</title>
<style>
  :root { color-scheme: light dark; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
         margin: 0; padding: 1.5rem; max-width: 760px; margin-inline: auto;
         line-height: 1.5; }
  h1 { font-size: 1.5rem; margin: 0 0 .25rem; }
  .sub { color: #888; margin: 0 0 1.5rem; font-size: .9rem; }
  .card { border: 1px solid rgba(128,128,128,.3); border-radius: 12px;
          padding: 1.25rem; margin-bottom: 1.25rem; }
  .card h2 { font-size: 1.1rem; margin: 0 0 .75rem; }
  button { font: inherit; padding: .6rem 1.1rem; border-radius: 8px;
           border: none; background: #2563eb; color: #fff; cursor: pointer;
           margin: .25rem .25rem .25rem 0; }
  button.secondary { background: rgba(128,128,128,.2); color: inherit; }
  button:disabled { opacity: .5; cursor: not-allowed; }
  input[type=file] { margin: .5rem 0; display: block; }
  .warn { background: rgba(220,80,20,.12); border: 1px solid rgba(220,80,20,.4);
          padding: .75rem 1rem; border-radius: 8px; margin-bottom: 1.25rem; }
  progress { width: 100%; height: 1.1rem; }
  #status { margin-top: 1rem; }
  .muted { color: #888; font-size: .85rem; }
  .hl { font-family: ui-monospace, Menlo, Consolas, monospace; font-size: .8rem;
        background: rgba(128,128,128,.12); padding: .5rem .75rem; border-radius: 8px;
        max-height: 220px; overflow: auto; white-space: pre-wrap; }
  .ok { color: #16a34a; } .bad { color: #dc2626; } .big { font-size: 1.6rem; font-weight: 700; }
  .counts { display: flex; gap: 1.5rem; margin: .5rem 0; }
  .counts div { text-align: center; }
</style>
</head>
<body>
  <h1>iCloud Kontakte Sync</h1>
  <p class="sub">Angemeldet als <strong>{{ user or "—" }}</strong></p>

  {% if not creds_ok %}
  <div class="warn">
    <strong>Zugangsdaten fehlen.</strong> Bitte <code>ICLOUD_USER</code> und
    <code>ICLOUD_PASS</code> in der <code>.env</code>-Datei setzen und die Seite neu laden.
  </div>
  {% endif %}

  <div class="card">
    <h2>1 · Export</h2>
    <p class="muted">Holt alle Kontakte inkl. Fotos aus iCloud.</p>
    <button id="btnExport" onclick="startExport()">Kontakte aus iCloud exportieren</button>
    <div id="exportDone" style="display:none">
      <a href="/api/download/vcf"><button class="secondary">VCF herunterladen</button></a>
      <a href="/api/download/xlsx"><button class="secondary">Als Excel herunterladen</button></a>
    </div>
  </div>

  <div class="card">
    <h2>2 · Import</h2>
    <p class="muted">Bearbeitete Datei hochladen (VCF oder die Excel-Liste),
       erst einen Testlauf machen, dann wirklich importieren.</p>
    <input type="file" id="fileInput" accept=".vcf,.xlsx">
    <button onclick="upload()">Datei hochladen</button>
    <div id="uploadInfo" class="muted"></div>
    <div id="importButtons" style="display:none; margin-top:.75rem">
      <button onclick="startImport(true)">Testlauf (nichts wird geschrieben)</button>
      <button id="btnReal" onclick="startImport(false)" disabled>Wirklich importieren</button>
    </div>
  </div>

  <div id="status"></div>

<script>
let polling = null;

function setBusy(b){
  document.querySelectorAll('button').forEach(x => {
    if (x.id !== 'btnReal') x.disabled = b;
  });
}

async function startExport(){
  document.getElementById('exportDone').style.display = 'none';
  const r = await fetch('/api/export', {method:'POST'});
  if(r.ok){ setBusy(true); poll(); }
}

async function upload(){
  const f = document.getElementById('fileInput').files[0];
  if(!f){ alert('Bitte erst eine Datei auswählen.'); return; }
  const fd = new FormData(); fd.append('file', f);
  document.getElementById('uploadInfo').textContent = 'Lade hoch ...';
  const r = await fetch('/api/upload', {method:'POST', body: fd});
  const j = await r.json();
  if(!r.ok){ document.getElementById('uploadInfo').textContent = 'Fehler: ' + j.error; return; }
  document.getElementById('uploadInfo').textContent = j.count + ' Kontakte in der Datei gefunden.';
  document.getElementById('importButtons').style.display = 'block';
  document.getElementById('btnReal').disabled = true;
}

async function startImport(dry){
  const r = await fetch('/api/import', {method:'POST', headers:{'Content-Type':'application/json'},
                        body: JSON.stringify({dry_run: dry})});
  if(r.ok){ setBusy(true); poll(); }
}

function poll(){
  if(polling) clearInterval(polling);
  polling = setInterval(refresh, 800);
  refresh();
}

async function refresh(){
  const r = await fetch('/api/status');
  const j = await r.json();
  const s = document.getElementById('status');
  let pct = j.total ? Math.round(j.done/j.total*100) : 0;
  if(j.running){
    s.innerHTML = '<div class="card"><strong>' +
      (j.kind === 'export' ? 'Export läuft …' : 'Import läuft …') + '</strong>' +
      '<p class="muted">'+ (j.message||'') +'</p>' +
      '<progress max="100" value="'+pct+'"></progress></div>';
    return;
  }
  clearInterval(polling); polling = null; setBusy(false);
  if(j.error){
    s.innerHTML = '<div class="card"><strong class="bad">Fehler:</strong> '+ j.error +'</div>';
    return;
  }
  if(!j.result){ s.innerHTML=''; return; }
  if(j.kind === 'export'){
    document.getElementById('exportDone').style.display = 'block';
    let w = j.result.warnings && j.result.warnings.length
      ? '<div class="hl">'+ j.result.warnings.join('\\n') +'</div>' : '';
    s.innerHTML = '<div class="card"><strong class="ok">Export fertig.</strong>' +
      '<p>'+ j.result.loaded +' Kontakte geladen, '+ j.result.photos +' Fotos eingebettet.</p>' +
      'Jetzt oben herunterladen.'+ w +'</div>';
  } else {
    const res = j.result;
    document.getElementById('btnReal').disabled = res.dry_run ? false : true;
    let hl = res.highlights && res.highlights.length
      ? '<div class="hl">'+ res.highlights.join('\\n') +'</div>' : '';
    let head = res.dry_run
      ? '<strong>Testlauf-Ergebnis (nichts geschrieben):</strong>'
      : '<strong class="ok">Import abgeschlossen.</strong>';
    let delWarn = (res.dry_run && res.deleted)
      ? '<p class="bad">Achtung: '+res.deleted+' Kontakt(e) würden gelöscht (siehe unten).</p>' : '';
    let hint = res.dry_run
      ? '<p class="muted">Sieht das gut aus? Dann auf „Wirklich importieren“.</p>' : '';
    s.innerHTML = '<div class="card">'+ head +
      '<div class="counts">' +
      '<div><div class="big">'+res.updated+'</div>aktualisiert</div>' +
      '<div><div class="big">'+res.created+'</div>neu</div>' +
      '<div><div class="big '+(res.deleted?'bad':'')+'">'+res.deleted+'</div>gelöscht</div>' +
      '<div><div class="big '+(res.errors?'bad':'')+'">'+res.errors+'</div>Fehler</div>' +
      '</div>'+ delWarn + hint + hl +'</div>';
  }
}

refresh();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "5000"))
    print(f"iCloud Kontakte Sync - Web-Oberfläche")
    print(f"Öffne im Browser: http://{host}:{port}")
    app.run(host=host, port=port, debug=False)
