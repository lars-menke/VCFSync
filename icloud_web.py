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


def _highlights_of(log, names):
    """Neu-/Lösch-/Fehlerzeilen mit echten Namen anreichern (die interessieren)."""
    def enrich(line):
        m = re.search(r":\s*([0-9a-fA-F-]{6,})", line)
        if m:
            for uid, name in names.items():
                if uid.startswith(m.group(1)) or m.group(1) in uid:
                    return f"{line}  ({name})"
        return line

    return [enrich(l) for l in log
            if any(k in l for k in ("NEU ANLEGEN", "NEU ANGELEGT", "LÖSCHEN",
                                    "GELÖSCHT", "FEHLER", "Hinweis"))]


def _run_import(dry_run, sync_target):
    try:
        if not IMPORT_VCF.exists():
            raise RuntimeError("Keine Import-Datei vorhanden. Bitte erst hochladen.")
        vcards = ic.split_vcards(IMPORT_VCF.read_text(encoding="utf-8"))
        names = _uid_name_map(vcards)
        total_steps = len(vcards) * (2 if sync_target == "both" else 1)
        base_result = {"dry_run": dry_run, "icloud": None, "google": None}

        if sync_target in ("icloud", "both"):
            session, books, primary = build_session()
            _job_progress(0, total_steps, "Lese vorhandene Kontakte aus iCloud (UID-Abgleich) ...")

            def progress_ic(i, total, message):
                _job_progress(i, total_steps, f"iCloud: {i}/{total} ...")

            result = ic.import_vcards(session, books, primary, vcards,
                                      dry_run=dry_run, progress=progress_ic)
            changes = [f"{c['name']}: {', '.join(c['fields'])}" for c in result["changed_list"]]
            base_result["icloud"] = {
                "updated": result["updated"], "created": result["created"],
                "deleted": result["deleted"], "changed": result["changed"],
                "changes": changes, "errors": result["errors"],
                "highlights": _highlights_of(result["log"], names),
            }

        if sync_target in ("google", "both"):
            import google_contacts as gc
            offset = len(vcards) if sync_target == "both" else 0
            _job_progress(offset, total_steps, "Verbinde mit Google Contacts ...")
            service = gc.build_google_service(interactive=False)

            def progress_g(i, total, message):
                _job_progress(offset + i, total_steps, f"Google: {i}/{total} ...")

            gresult = gc.push_contacts_to_google(service, vcards, dry_run=dry_run,
                                                 progress=progress_g)
            base_result["google"] = {
                "updated": gresult["updated"], "created": gresult["created"],
                "deleted": gresult["deleted"], "errors": gresult["errors"],
                "unchanged": gresult.get("unchanged", 0),
                "highlights": _highlights_of(gresult["log"], names),
            }

        _job_finish(result=base_result)
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
    sync_target = (request.json or {}).get("target", "icloud")
    if sync_target not in ("icloud", "google", "both"):
        return jsonify({"error": "Ungültiges Ziel."}), 400
    if not _start("import", _run_import, dry_run, sync_target):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


@app.route("/api/google/status")
def api_google_status():
    """Ob eine gültige Google-Anmeldung vorliegt - für den UI-Hinweis, ob
    'python icloud_contacts.py google-auth' erst noch nötig ist."""
    import google_contacts as gc
    try:
        creds = gc.get_google_credentials(interactive=False)
        return jsonify({"connected": bool(creds and creds.valid)})
    except Exception:  # noqa: BLE001 - fehlende Anmeldung/Pakete zaehlt als "nicht verbunden"
        return jsonify({"connected": False})


PAGE = """
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>iCloud Kontakte Sync</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {
    color-scheme: light dark;
    --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    --font-mono: 'JetBrains Mono', ui-monospace, Menlo, Consolas, monospace;

    --color-primary: #2563eb;
    --color-primary-hover: #1d4ed8;
    --color-primary-soft: #eff6ff;
    --color-accent: #ea580c;
    --color-accent-hover: #c2410c;
    --color-accent-soft: #fff7ed;
    --color-success: #16a34a;
    --color-success-soft: #f0fdf4;
    --color-danger: #dc2626;
    --color-danger-soft: #fef2f2;
    --color-warning: #d97706;
    --color-warning-soft: #fffbeb;

    --color-bg: #f8fafc;
    --color-surface: #ffffff;
    --color-surface-2: #f1f5f9;
    --color-border: #e2e8f0;
    --color-text: #0f172a;
    --color-text-muted: #64748b;
    --color-on-primary: #ffffff;
    --color-on-accent: #ffffff;

    --radius: 14px;
    --radius-sm: 9px;
    --shadow-card: 0 1px 2px rgba(15,23,42,.04), 0 1px 12px rgba(15,23,42,.05);
    --shadow-pop: 0 8px 24px rgba(15,23,42,.10);
    --ease: cubic-bezier(.16,1,.3,1);
  }

  @media (prefers-color-scheme: dark) {
    :root {
      --color-primary: #3b82f6;
      --color-primary-hover: #60a5fa;
      --color-primary-soft: rgba(59,130,246,.12);
      --color-accent: #fb923c;
      --color-accent-hover: #fdba74;
      --color-accent-soft: rgba(251,146,60,.12);
      --color-success: #4ade80;
      --color-success-soft: rgba(74,222,128,.10);
      --color-danger: #f87171;
      --color-danger-soft: rgba(248,113,113,.10);
      --color-warning: #fbbf24;
      --color-warning-soft: rgba(251,191,36,.10);

      --color-bg: #0b1220;
      --color-surface: #131b2c;
      --color-surface-2: #1a2337;
      --color-border: #253048;
      --color-text: #e6ebf5;
      --color-text-muted: #8b96ac;
      --color-on-primary: #0b1220;
      --color-on-accent: #0b1220;
      --shadow-card: 0 1px 2px rgba(0,0,0,.3), 0 1px 16px rgba(0,0,0,.25);
      --shadow-pop: 0 12px 32px rgba(0,0,0,.4);
    }
  }

  * { box-sizing: border-box; }
  html { -webkit-text-size-adjust: 100%; }
  body {
    font-family: var(--font-sans);
    background: var(--color-bg);
    color: var(--color-text);
    margin: 0;
    padding: 2rem 1.25rem 4rem;
    line-height: 1.55;
    font-size: 15px;
    -webkit-font-smoothing: antialiased;
  }
  .page { max-width: 720px; margin-inline: auto; }

  .topbar { display: flex; align-items: center; gap: .9rem; margin-bottom: 1.75rem; }
  .brand-mark {
    width: 44px; height: 44px; flex: none; border-radius: 12px;
    background: linear-gradient(145deg, var(--color-primary), var(--color-accent));
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-card);
  }
  .brand-mark svg { width: 24px; height: 24px; color: #fff; }
  h1 { font-size: 1.3rem; font-weight: 700; margin: 0; letter-spacing: -.01em; }
  .sub { color: var(--color-text-muted); margin: .15rem 0 0; font-size: .875rem; }
  .sub strong { color: var(--color-text); font-weight: 600; }

  .card {
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    box-shadow: var(--shadow-card);
    padding: 1.5rem;
    margin-bottom: 1.25rem;
    animation: rise .35s var(--ease) both;
  }
  @keyframes rise { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }

  .card-head { display: flex; align-items: flex-start; gap: .75rem; margin-bottom: 1.1rem; }
  .card-icon {
    width: 34px; height: 34px; flex: none; border-radius: 10px;
    background: var(--color-primary-soft); color: var(--color-primary);
    display: flex; align-items: center; justify-content: center;
  }
  .card-icon svg { width: 18px; height: 18px; }
  .card h2 { font-size: 1rem; font-weight: 600; margin: .1rem 0 .2rem; }
  .card .muted.desc { margin: 0; }

  .muted { color: var(--color-text-muted); font-size: .85rem; }
  code {
    font-family: var(--font-mono); font-size: .82em;
    background: var(--color-surface-2); padding: .1em .4em; border-radius: 5px;
  }

  .banner {
    display: flex; gap: .75rem; align-items: flex-start;
    background: var(--color-warning-soft); border: 1px solid var(--color-warning);
    color: var(--color-text); padding: 1rem 1.1rem; border-radius: var(--radius-sm);
    margin-bottom: 1.25rem; font-size: .875rem;
  }
  .banner svg { width: 20px; height: 20px; flex: none; color: var(--color-warning); margin-top: .05rem; }

  .btn {
    font: inherit; font-weight: 600; font-size: .875rem;
    display: inline-flex; align-items: center; justify-content: center; gap: .5rem;
    min-height: 42px; padding: 0 1.1rem; border-radius: var(--radius-sm);
    border: 1px solid transparent; cursor: pointer; text-decoration: none;
    transition: transform .12s var(--ease), background-color .15s var(--ease),
                border-color .15s var(--ease), box-shadow .15s var(--ease), opacity .15s var(--ease);
    margin: .2rem .4rem .2rem 0; user-select: none; -webkit-tap-highlight-color: transparent;
  }
  .btn svg { width: 17px; height: 17px; flex: none; }
  .btn:active { transform: scale(.97); }
  .btn:focus-visible { outline: 2px solid var(--color-primary); outline-offset: 2px; }
  .btn:disabled { opacity: .45; cursor: not-allowed; transform: none; }

  .btn-primary { background: var(--color-primary); color: var(--color-on-primary); }
  .btn-primary:hover:not(:disabled) { background: var(--color-primary-hover); box-shadow: 0 4px 14px rgba(37,99,235,.25); }

  .btn-accent { background: var(--color-accent); color: var(--color-on-accent); }
  .btn-accent:hover:not(:disabled) { background: var(--color-accent-hover); box-shadow: 0 4px 14px rgba(234,88,12,.25); }

  .btn-secondary { background: var(--color-surface-2); color: var(--color-text); border-color: var(--color-border); }
  .btn-secondary:hover:not(:disabled) { background: var(--color-border); }

  .download-row, .button-row { margin-top: .9rem; }
  .hidden { display: none !important; }

  .target-group {
    display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;
    border: 1px solid var(--color-border); border-radius: var(--radius-sm);
    padding: .6rem .9rem .6rem 0; margin: 0 0 .5rem;
  }
  .target-group legend { padding: 0 .6rem; font-size: .78rem; text-transform: uppercase; letter-spacing: .04em; }
  .target-option { display: inline-flex; align-items: center; gap: .4rem; font-size: .85rem; cursor: pointer; }
  .target-option input { accent-color: var(--color-primary); width: 16px; height: 16px; cursor: pointer; }
  #googleHint { margin: -.2rem 0 .6rem; }
  #googleHint code { display: inline-block; margin-top: .2rem; }

  .upload-row { display: flex; align-items: center; gap: .75rem; flex-wrap: wrap; margin-bottom: .6rem; }
  input[type=file] { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0,0,0,0); }
  .file-label {
    display: inline-flex; align-items: center; gap: .5rem; min-height: 42px; padding: 0 1.1rem;
    border-radius: var(--radius-sm); border: 1px dashed var(--color-border);
    background: var(--color-surface-2); color: var(--color-text); cursor: pointer;
    font-size: .875rem; font-weight: 600; transition: border-color .15s var(--ease), background-color .15s var(--ease);
  }
  .file-label:hover { border-color: var(--color-primary); background: var(--color-primary-soft); }
  .file-label svg { width: 17px; height: 17px; }
  .file-label:has(+ input:focus-visible) { outline: 2px solid var(--color-primary); outline-offset: 2px; }
  #fileName { font-size: .85rem; color: var(--color-text-muted); }

  #uploadInfo:not(:empty) { margin: .6rem 0 0; font-size: .85rem; }

  progress { width: 100%; height: 10px; border: none; border-radius: 999px; overflow: hidden;
             margin-top: .75rem; -webkit-appearance: none; appearance: none; }
  progress::-webkit-progress-bar { background: var(--color-surface-2); border-radius: 999px; }
  progress::-webkit-progress-value { background: var(--color-primary); border-radius: 999px; transition: width .3s var(--ease); }
  progress::-moz-progress-bar { background: var(--color-primary); border-radius: 999px; }
  progress.indeterminate { background: var(--color-surface-2); position: relative; }
  progress.indeterminate::-webkit-progress-value { background: var(--color-primary); }
  .progress-track { position: relative; }
  .progress-track.indeterminate progress { opacity: 0; }
  .progress-track.indeterminate::after {
    content: ""; position: absolute; left: 0; top: .75rem; height: 10px; width: 40%;
    border-radius: 999px; background: var(--color-primary);
    animation: indeterminate 1.1s ease-in-out infinite;
  }
  @keyframes indeterminate {
    0% { transform: translateX(-100%); } 100% { transform: translateX(250%); }
  }

  .spinner {
    width: 15px; height: 15px; border-radius: 50%;
    border: 2px solid rgba(255,255,255,.4); border-top-color: currentColor;
    animation: spin .7s linear infinite; flex: none;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  #status { margin-top: 1.25rem; }
  #status .card { margin-bottom: 0; }

  .status-head { display: flex; align-items: center; gap: .6rem; margin-bottom: .15rem; }
  .status-head svg { width: 19px; height: 19px; flex: none; }
  .status-title { font-weight: 700; font-size: .95rem; }
  .ok { color: var(--color-success); } .bad { color: var(--color-danger); } .warn-text { color: var(--color-warning); }

  .stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: .6rem; margin: 1rem 0; }
  .stat-tile {
    background: var(--color-surface-2); border: 1px solid var(--color-border);
    border-radius: var(--radius-sm); padding: .75rem .5rem; text-align: center;
  }
  .stat-tile.stat-bad { background: var(--color-danger-soft); border-color: var(--color-danger); }
  .stat-value { font-size: 1.5rem; font-weight: 700; line-height: 1.1; font-variant-numeric: tabular-nums; }
  .stat-label { font-size: .72rem; color: var(--color-text-muted); margin-top: .25rem;
                text-transform: uppercase; letter-spacing: .04em; }

  .callout { display: flex; gap: .6rem; align-items: flex-start; padding: .75rem .9rem;
             border-radius: var(--radius-sm); font-size: .85rem; margin: .75rem 0; }
  .callout svg { width: 18px; height: 18px; flex: none; margin-top: .05rem; }
  .callout-danger { background: var(--color-danger-soft); color: var(--color-danger); }
  .callout-info { background: var(--color-primary-soft); color: var(--color-primary); }

  .hl {
    font-family: var(--font-mono); font-size: .78rem; line-height: 1.65;
    background: var(--color-surface-2); border: 1px solid var(--color-border);
    padding: .7rem .85rem; border-radius: var(--radius-sm);
    max-height: 220px; overflow: auto; white-space: pre-wrap; word-break: break-word;
  }
  .hl .l-err { color: var(--color-danger); }
  .hl .l-warn { color: var(--color-warning); }
  .hl .l-new { color: var(--color-primary); }
  .hl-title { margin: .9rem 0 .35rem; font-size: .8rem; font-weight: 600; color: var(--color-text-muted); }

  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation-duration: .001ms !important; transition-duration: .001ms !important; }
  }
  @media (max-width: 480px) {
    .stat-grid { grid-template-columns: repeat(2, 1fr); }
    body { padding: 1.5rem 1rem 3rem; }
  }
</style>
</head>
<body>
<div class="page">

  <div class="topbar">
    <div class="brand-mark">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M4 4v5h5M20 20v-5h-5"/>
        <path d="M20 9a8 8 0 0 0-14.93-3M4 15a8 8 0 0 0 14.93 3"/>
      </svg>
    </div>
    <div>
      <h1>iCloud Kontakte Sync</h1>
      <p class="sub">Angemeldet als <strong>{{ user or "—" }}</strong></p>
    </div>
  </div>

  {% if not creds_ok %}
  <div class="banner">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 9v4m0 4h.01M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/>
    </svg>
    <div><strong>Zugangsdaten fehlen.</strong> Bitte <code>ICLOUD_USER</code> und
    <code>ICLOUD_PASS</code> in der <code>.env</code>-Datei setzen und die Seite neu laden.</div>
  </div>
  {% endif %}

  <div class="card">
    <div class="card-head">
      <div class="card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 3v12m0 0-4-4m4 4 4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
      </div>
      <div>
        <h2>1 · Export</h2>
        <p class="muted desc">Holt alle Kontakte inkl. Fotos aus iCloud.</p>
      </div>
    </div>
    <button id="btnExport" class="btn btn-primary" onclick="startExport()">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M4 4v5h5M20 20v-5h-5"/><path d="M20 9a8 8 0 0 0-14.93-3M4 15a8 8 0 0 0 14.93 3"/>
      </svg>
      Kontakte aus iCloud exportieren
    </button>
    <div id="exportDone" class="download-row hidden">
      <a href="/api/download/vcf" class="btn btn-secondary">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 3v12m0 0-4-4m4 4 4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
        VCF herunterladen
      </a>
      <a href="/api/download/xlsx" class="btn btn-secondary">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 14h18M9 4v16M15 4v16"/>
        </svg>
        Als Excel herunterladen
      </a>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <div class="card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 21V9m0 0-4 4m4-4 4 4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
      </div>
      <div>
        <h2>2 · Import</h2>
        <p class="muted desc">Bearbeitete Datei hochladen (VCF oder die Excel-Liste),
           erst einen Testlauf machen, dann wirklich importieren.</p>
      </div>
    </div>

    <div class="upload-row">
      <label class="file-label" for="fileInput">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 21V9m0 0-4 4m4-4 4 4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
        Datei auswählen
      </label>
      <input type="file" id="fileInput" accept=".vcf,.xlsx" onchange="onFileChosen()">
      <span id="fileName" class="muted">Keine Datei ausgewählt</span>
    </div>
    <button class="btn btn-secondary" onclick="upload()">Hochladen</button>
    <p id="uploadInfo" class="muted"></p>

    <div id="importButtons" class="button-row hidden">
      <fieldset class="target-group">
        <legend class="muted">Ziel</legend>
        <label class="target-option">
          <input type="radio" name="syncTarget" value="icloud" checked> iCloud
        </label>
        <label class="target-option">
          <input type="radio" name="syncTarget" value="google" id="targetGoogle"> Google
        </label>
        <label class="target-option">
          <input type="radio" name="syncTarget" value="both" id="targetBoth"> Beide
        </label>
      </fieldset>
      <p id="googleHint" class="muted hidden">
        Google ist noch nicht verbunden. Einmalig im Terminal:
        <code>python icloud_contacts.py google-auth</code>
      </p>
      <button class="btn btn-secondary" onclick="startImport(true)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>
        </svg>
        Testlauf (nichts wird geschrieben)
      </button>
      <button id="btnReal" class="btn btn-accent" onclick="startImport(false)" disabled>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 21V9m0 0-4 4m4-4 4 4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
        Wirklich importieren
      </button>
    </div>
  </div>

  <div id="status" aria-live="polite"></div>

</div>

<script>
let polling = null;

const ICONS = {
  ok: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>',
  bad: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>',
  warn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4m0 4h.01M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/></svg>',
  spin: '<span class="spinner"></span>'
};

function esc(s){
  return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
}

function highlightLog(text){
  return text.split('\\n').map(line => {
    let cls = '';
    if (/FEHLER/.test(line)) cls = 'l-err';
    else if (/GEL(Ö|OE)SCHT|W(Ü|UE)RDE L(Ö|OE)SCHEN/.test(line)) cls = 'l-warn';
    else if (/NEU ANGELEGT|NEU ANLEGEN/.test(line)) cls = 'l-new';
    const safe = esc(line);
    return cls ? '<span class="'+cls+'">'+safe+'</span>' : safe;
  }).join('\\n');
}

function onFileChosen(){
  const f = document.getElementById('fileInput').files[0];
  document.getElementById('fileName').textContent = f ? f.name : 'Keine Datei ausgewählt';
}

function setBusy(b){
  document.querySelectorAll('.btn').forEach(x => {
    if (x.id !== 'btnReal') x.disabled = b;
  });
}

async function startExport(){
  document.getElementById('exportDone').classList.add('hidden');
  const r = await fetch('/api/export', {method:'POST'});
  if(r.ok){ setBusy(true); poll(); }
}

async function upload(){
  const f = document.getElementById('fileInput').files[0];
  if(!f){ alert('Bitte erst eine Datei auswählen.'); return; }
  const fd = new FormData(); fd.append('file', f);
  document.getElementById('uploadInfo').textContent = 'Lade hoch …';
  const r = await fetch('/api/upload', {method:'POST', body: fd});
  const j = await r.json();
  if(!r.ok){ document.getElementById('uploadInfo').innerHTML = '<span class="bad">Fehler: '+esc(j.error)+'</span>'; return; }
  document.getElementById('uploadInfo').textContent = j.count + ' Kontakte in der Datei gefunden.';
  document.getElementById('importButtons').classList.remove('hidden');
  document.getElementById('btnReal').disabled = true;
}

function selectedTarget(){
  const el = document.querySelector('input[name=syncTarget]:checked');
  return el ? el.value : 'icloud';
}

async function startImport(dry){
  const r = await fetch('/api/import', {method:'POST', headers:{'Content-Type':'application/json'},
                        body: JSON.stringify({dry_run: dry, target: selectedTarget()})});
  if(r.ok){ setBusy(true); poll(); }
}

async function checkGoogleStatus(){
  try {
    const r = await fetch('/api/google/status');
    const j = await r.json();
    if(!j.connected){
      document.getElementById('targetGoogle').disabled = true;
      document.getElementById('targetBoth').disabled = true;
      document.getElementById('googleHint').classList.remove('hidden');
    }
  } catch(e){ /* Status ist rein informativ - bei Fehler einfach nichts anzeigen */ }
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
  let indeterminate = !j.total;

  if(j.running){
    s.innerHTML = '<div class="card">' +
      '<div class="status-head">' + ICONS.spin +
      '<span class="status-title">'+ (j.kind === 'export' ? 'Export läuft …' : 'Import läuft …') +'</span></div>' +
      '<p class="muted">'+ esc(j.message||'') +'</p>' +
      '<div class="progress-track'+(indeterminate?' indeterminate':'')+'">' +
      '<progress max="100" value="'+pct+'"></progress></div></div>';
    return;
  }
  clearInterval(polling); polling = null; setBusy(false);

  if(j.error){
    s.innerHTML = '<div class="card"><div class="status-head bad">'+ ICONS.bad +
      '<span class="status-title">Fehler</span></div><p>'+ esc(j.error) +'</p></div>';
    return;
  }
  if(!j.result){ s.innerHTML=''; return; }

  if(j.kind === 'export'){
    document.getElementById('exportDone').classList.remove('hidden');
    let w = j.result.warnings && j.result.warnings.length
      ? '<p class="hl-title">Warnungen</p><div class="hl">'+ highlightLog(j.result.warnings.join('\\n')) +'</div>' : '';
    s.innerHTML = '<div class="card"><div class="status-head ok">'+ ICONS.ok +
      '<span class="status-title">Export fertig</span></div>' +
      '<p class="muted" style="margin:.35rem 0 0">'+ j.result.loaded +' Kontakte geladen, '+ j.result.photos +
      ' Fotos eingebettet. Jetzt oben herunterladen.</p>'+ w +'</div>';
  } else {
    const res = j.result;
    const anyDeleted = (res.icloud && res.icloud.deleted) || (res.google && res.google.deleted);
    document.getElementById('btnReal').disabled = res.dry_run ? false : true;

    let icon = res.dry_run ? ICONS.warn : ICONS.ok;
    let head = res.dry_run ? 'Testlauf-Ergebnis (nichts geschrieben)' : 'Import abgeschlossen';
    let delWarn = (res.dry_run && anyDeleted)
      ? '<div class="callout callout-danger">'+ ICONS.warn +
        '<div>Achtung: es würden Kontakte gelöscht (siehe unten).</div></div>' : '';
    let hint = res.dry_run
      ? '<div class="callout callout-info">'+ ICONS.ok +
        '<div>Sieht das gut aus? Dann auf „Wirklich importieren“ klicken.</div></div>' : '';

    let body = '<div class="card">' +
      '<div class="status-head '+(res.dry_run?'warn-text':'ok')+'">'+ icon +
      '<span class="status-title">'+ head +'</span></div>' + delWarn + hint;

    if(res.icloud) body += renderTargetResult('iCloud', res.icloud, true);
    if(res.google) body += renderTargetResult('Google', res.google, false);

    s.innerHTML = body + '</div>';
  }
}

function renderTargetResult(label, res, showChanged){
  let hl = res.highlights && res.highlights.length
    ? '<p class="hl-title">'+label+' – Details</p><div class="hl">'+ highlightLog(res.highlights.join('\\n')) +'</div>' : '';
  let changeList = (res.changes && res.changes.length)
    ? '<p class="hl-title">'+label+' – geänderte Kontakte</p><div class="hl">'+ esc(res.changes.join('\\n')) +'</div>' : '';
  let firstTile = showChanged
    ? '<div class="stat-tile"><div class="stat-value">'+res.changed+'</div><div class="stat-label">Geändert</div></div>'
    : '<div class="stat-tile"><div class="stat-value">'+res.updated+'</div><div class="stat-label">Aktualisiert</div></div>';
  let summary = showChanged
    ? res.updated+' vorhandene Kontakte geprüft, davon '+res.changed+' inhaltlich geändert.'
    : res.updated+' aktualisiert, '+(res.unchanged||0)+' bereits aktuell (übersprungen).';
  return '<p class="hl-title" style="margin-top:1.1rem">'+label+'</p>' +
    '<div class="stat-grid">' + firstTile +
    '<div class="stat-tile"><div class="stat-value">'+res.created+'</div><div class="stat-label">Neu</div></div>' +
    '<div class="stat-tile'+(res.deleted?' stat-bad':'')+'"><div class="stat-value">'+res.deleted+'</div><div class="stat-label">Gelöscht</div></div>' +
    '<div class="stat-tile'+(res.errors?' stat-bad':'')+'"><div class="stat-value">'+res.errors+'</div><div class="stat-label">Fehler</div></div>' +
    '</div><p class="muted">'+summary+'</p>' + changeList + hl;
}

checkGoogleStatus();
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
