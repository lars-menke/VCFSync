"""
iCloud Kontakte Sync - Web-Oberfläche
=====================================
Eine schlichte lokale Web-App (Flask), die Export und Import per Knopfdruck
bedienbar macht. Nutzt dieselbe, bewährte Kernlogik wie das CLI-Skript
(icloud_contacts.py) - hier gibt es keine doppelte Sync-Logik.

Start:
    pip install -r requirements.txt
    python icloud_web.py
    -> Browser öffnen: http://127.0.0.1:8000

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
import mail_cleanup as mc

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
                "unchanged": result["unchanged"],
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


# ---------------------------------------------------------------------------
# E-Mail aufräumen
# ---------------------------------------------------------------------------

MAIL_XLSX = DATA_DIR / "mail_aufraeumen.xlsx"


def _mail_error(e):
    return jsonify({"error": str(e)}), 400


def _run_mail_scan(selection, min_age):
    try:
        settings = {"min_age_days": min_age} if min_age is not None else None
        scan = mc.scan_accounts(
            selection, settings,
            progress=lambda message: _job_progress(0, 0, message))
        _job_finish(result={"kind": "scan", **mc.scan_summary(scan)})
    except Exception as e:  # noqa: BLE001
        _job_finish(error=str(e))


def _run_mail_clean(dry_run):
    try:
        scan = mc.load_scan()
        result = mc.clean(scan, dry_run=dry_run,
                          progress=lambda message: _job_progress(0, 0, message))
        _job_finish(result={"kind": "clean", **result})
    except Exception as e:  # noqa: BLE001
        _job_finish(error=str(e))


def _run_mail_search(selection, criteria):
    try:
        result = mc.search_mails(
            selection, criteria,
            progress=lambda message: _job_progress(0, 0, message))
        _job_finish(result={"kind": "search", **result})
    except Exception as e:  # noqa: BLE001
        _job_finish(error=str(e))


@app.route("/mail")
def mail_index():
    return render_template_string(MAIL_PAGE)


@app.route("/mail/suche")
def mail_search_index():
    return render_template_string(MAIL_SEARCH_PAGE)


@app.route("/api/mail/search", methods=["POST"])
def api_mail_search():
    """Postfach durchsuchen - über ALLE Ordner, auch die das Aufräumen
    bewusst ausschließt (Papierkorb, Spam). Läuft im Hintergrund wie scan/
    clean, weil ein ganzes Konto ohne Ordnerauswahl dauern kann."""
    data = request.json or {}
    criteria = {k: (data.get(k) or "").strip()
               for k in ("sender", "subject", "since", "before", "body_text")}
    if not any(criteria.values()):
        return jsonify({"error": "Bitte mindestens ein Kriterium angeben."}), 400

    accounts = mc.load_accounts()
    if not accounts:
        return jsonify({"error": "Noch kein Konto eingerichtet."}), 400
    wanted = data.get("account")
    names = [wanted] if wanted else [a["name"] for a in accounts]
    selection = [{"account": name} for name in names]

    if not _start("mailsearch", _run_mail_search, selection, criteria):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


@app.route("/api/mail/accounts", methods=["GET", "POST", "DELETE"])
def api_mail_accounts():
    """Konten verwalten. Passwörter werden nie zurückgeliefert."""
    try:
        if request.method == "GET":
            return jsonify({
                "accounts": [{"name": a["name"], "user": a["user"], "host": a["host"],
                              "port": a.get("port", 993)} for a in mc.load_accounts()],
                "presets": mc.PRESETS,
            })
        if request.method == "DELETE":
            name = (request.json or {}).get("name", "")
            if not name:
                return jsonify({"error": "Kein Konto angegeben."}), 400
            mc.remove_account(name)
            return jsonify({"ok": True})

        data = request.json or {}
        missing = [f for f in ("name", "host", "user", "password") if not data.get(f)]
        if missing:
            return jsonify({"error": f"Es fehlt: {', '.join(missing)}."}), 400
        mc.add_account(data["name"], data["host"], data["user"], data["password"],
                       data.get("port", 993))
        try:
            info = mc.test_account(mc.account_by_name(data["name"]))
        except Exception as e:  # noqa: BLE001 - Konto bleibt gespeichert, Fehler nur melden
            return jsonify({"ok": True, "tested": False, "message": str(e)})
        return jsonify({"ok": True, "tested": True, "folders": info["folders"],
                        "trash": info["trash"]})
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


@app.route("/api/mail/folders")
def api_mail_folders():
    name = request.args.get("account", "")
    try:
        return jsonify({"folders": mc.list_account_folders(mc.account_by_name(name))})
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


@app.route("/api/mail/diagnose")
def api_mail_diagnose():
    """Postfach durchleuchten - rein lesend, deshalb ohne Hintergrundaufgabe."""
    name = request.args.get("account", "")
    try:
        return jsonify(mc.diagnose_account(mc.account_by_name(name)))
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


@app.route("/api/mail/scan", methods=["DELETE"])
def api_mail_scan_clear():
    """Gespeicherte Liste verwerfen. Rührt das Postfach nicht an - es wird nur
    der Stand weggeworfen, mit dem das Werkzeug arbeitet."""
    return jsonify({"cleared": mc.clear_scan()})


@app.route("/api/mail/scan", methods=["POST"])
def api_mail_scan():
    data = request.json or {}
    selection = data.get("selection") or []
    if not selection or not any(s.get("folders") for s in selection):
        return jsonify({"error": "Bitte mindestens einen Ordner auswählen."}), 400
    min_age = data.get("min_age")
    if min_age is not None:
        try:
            min_age = max(0, int(min_age))
        except (TypeError, ValueError):
            return jsonify({"error": "Mindestalter muss eine Zahl sein."}), 400
    if not _start("mailscan", _run_mail_scan, selection, min_age):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


@app.route("/api/mail/result")
def api_mail_result():
    """Scan-Ergebnis für die Anzeige: nach Absender gruppiert, damit man
    ganze Newsletter-Serien auf einen Blick abhaken kann."""
    try:
        scan = mc.load_scan()
    except RuntimeError as e:
        return _mail_error(e)

    groups = {}
    for acc, folder, mail in mc.iter_mails(scan):
        # Nie vorgeschlagene Mails gehören nicht in die Auswahlliste - es sei
        # denn, sie wurden in der Excel-Liste von Hand angehakt. Sonst wären
        # sie hier unsichtbar und würden beim nächsten Klick stillschweigend
        # wieder abgewählt (die Web-Auswahl überschreibt den ganzen Stand).
        if mail.get("recommendation") == "keep" and not mail.get("delete"):
            continue
        sender = mail.get("from") or "(ohne Absender)"
        group = groups.setdefault(sender, {
            "sender": sender, "name": mail.get("from_name", ""), "mails": [],
            "selected": 0, "bytes": 0})
        group["mails"].append({
            "key": mc.mail_key(acc["account"], folder["folder"], mail["uid"]),
            "account": acc["account"], "folder": folder["folder"],
            "subject": mail.get("subject", ""), "date": (mail.get("date") or "")[:10],
            "size": mail.get("size", 0), "score": mail.get("score", 0),
            "recommendation": mail.get("recommendation"),
            "reasons": mail.get("reasons", []), "delete": bool(mail.get("delete")),
            "attachment": bool(mail.get("has_attachment")),
            "attachments": mail.get("attachments", []),
            "mail_type": mail.get("mail_type"),
        })
        group["selected"] += 1 if mail.get("delete") else 0
        group["bytes"] += mail.get("size", 0)

    ordered = sorted(groups.values(), key=lambda g: (-len(g["mails"]), g["sender"]))
    for group in ordered:
        group["mails"].sort(key=lambda m: m["date"])
        group["attachments"] = sum(1 for m in group["mails"] if m["attachment"])
    return jsonify({"groups": ordered, "summary": mc.scan_summary(scan),
                    "created": scan.get("created"),
                    "executed": scan.get("executed")})


@app.route("/api/mail/select", methods=["POST"])
def api_mail_select():
    keys = (request.json or {}).get("keys")
    if not isinstance(keys, list):
        return jsonify({"error": "Ungültige Auswahl."}), 400
    try:
        return jsonify({"selected": mc.apply_selection(mc.load_scan(), keys)})
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


@app.route("/api/mail/download/xlsx")
def api_mail_download_xlsx():
    try:
        mc.scan_to_excel(mc.load_scan(), MAIL_XLSX)
    except RuntimeError as e:
        return str(e), 404
    return send_file(MAIL_XLSX, as_attachment=True, download_name="mail_aufraeumen.xlsx")


@app.route("/api/mail/upload", methods=["POST"])
def api_mail_upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Bitte die bearbeitete .xlsx-Datei hochladen."}), 400
    up = DATA_DIR / "mail_upload.xlsx"
    f.save(up)
    try:
        return jsonify(mc.scan_from_excel(mc.load_scan(), up))
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


@app.route("/api/mail/clean", methods=["POST"])
def api_mail_clean():
    dry_run = bool((request.json or {}).get("dry_run", True))
    if not _start("mailclean", _run_mail_clean, dry_run):
        return jsonify({"error": "Es läuft bereits eine Aufgabe."}), 409
    return jsonify({"started": True})


@app.route("/api/mail/learned")
def api_mail_learned():
    return jsonify({"rows": mc.learned_summary()})


@app.route("/api/mail/rules", methods=["GET", "POST", "DELETE"])
def api_mail_rules():
    """Von Hand gepflegte Regeln. Nach einer Änderung muss neu bewertet werden -
    das meldet die Antwort über 'rescan', damit die Oberfläche darauf hinweisen
    kann, statt stillschweigend veraltete Vorschläge stehen zu lassen."""
    try:
        if request.method == "GET":
            return jsonify({"rules": mc.load_rules(), "lists": list(mc.RULE_LISTS)})

        data = request.json or {}
        list_name, value = data.get("list", ""), data.get("value", "")
        if not value:
            return jsonify({"error": "Kein Wert angegeben."}), 400

        if request.method == "DELETE":
            for name in mc.RULE_LISTS:
                mc.remove_rule(name, value)
        else:
            if list_name not in mc.RULE_LISTS:
                return jsonify({"error": "Unbekannte Regelliste."}), 400
            mc.add_rule(list_name, value)
        return jsonify({"rules": mc.load_rules(), "rescan": True})
    except Exception as e:  # noqa: BLE001
        return _mail_error(e)


BASE_CSS = """
<style>
  :root {
    color-scheme: light dark;
    --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    --font-mono: 'JetBrains Mono', ui-monospace, Menlo, Consolas, monospace;

    /* Petrol/Rost statt Standard-SaaS-Blau/Orange - siehe UI/UX-Review
       (.impeccable/critique). Halbtöne (-hover/-soft) davon abgeleitet,
       Kontrastwerte gegen die tatsächlichen Hintergründe nachgemessen. */
    --color-primary: #0f6b63;        /* Petrol - 6.4:1 mit weißer Button-Schrift */
    --color-primary-hover: #0b4a45;
    --color-primary-soft: #e3efec;
    --color-accent: #b8560e;         /* Rost - 4.8:1 mit weißer Button-Schrift */
    --color-accent-hover: #8f4009;
    --color-accent-soft: #fbede0;
    --color-success: #16a34a;
    --color-success-soft: #f0fdf4;
    --color-danger: #dc2626;
    --color-danger-soft: #fef2f2;
    --color-danger-text: #b91c1c;    /* eigener Ton für Text auf -soft: 4.41:1 reichte nicht für WCAG AA */
    --color-warning: #d97706;
    --color-warning-soft: #fffbeb;
    --color-warning-text: #92400e;   /* dito - 3.07:1 auf -soft war zu wenig */

    --color-bg: #f6f7f5;             /* Papier - kühles Off-White statt neutralem Grau */
    --color-surface: #ffffff;
    --color-surface-2: #edf1ef;
    --color-border: #dce3e0;
    --color-text: #1b2430;           /* Tinte statt reinem Schwarz */
    --color-text-muted: #55666a;     /* ≈6:1 auf --color-surface-2/--color-bg */
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
      --color-primary: #4fbbae;
      --color-primary-hover: #6fcfc3;
      --color-primary-soft: #1d3936;
      --color-accent: #e08a48;
      --color-accent-hover: #eda36c;
      --color-accent-soft: #3a2a1b;
      --color-success: #4ade80;
      --color-success-soft: rgba(74,222,128,.10);
      --color-danger: #f87171;
      --color-danger-soft: rgba(248,113,113,.10);
      --color-danger-text: #fecaca;
      --color-warning: #fbbf24;
      --color-warning-soft: rgba(251,191,36,.10);
      --color-warning-text: #fde68a;

      --color-bg: #161c22;
      --color-surface: #1e262d;
      --color-surface-2: #222a31;
      --color-border: #2c363d;
      --color-text: #eaedec;
      --color-text-muted: #93a3a2;
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
    font-size: 16px;
    -webkit-font-smoothing: antialiased;
  }
  .page { max-width: 720px; margin-inline: auto; }

  .topbar { display: flex; align-items: center; gap: 1rem; margin-bottom: 1.9rem; }
  .brand-mark {
    width: 48px; height: 48px; flex: none; border-radius: 13px;
    background: linear-gradient(145deg, var(--color-primary), var(--color-accent));
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-card);
  }
  .brand-mark svg { width: 26px; height: 26px; color: #fff; }
  h1 { font-size: 2rem; font-weight: 700; margin: 0; letter-spacing: -.015em; line-height: 1.15; }
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

  /* --- Ablauf-Rail: verbundene, zustandsbehaftete Schritte statt loser,
     gleichwertiger Karten. Reihenfolge trägt hier echte Bedeutung (Konto vor
     Scan vor Prüfen vor Aufräumen), deshalb eine Linie mit Zustandspunkten
     statt sechs optisch identischer Karten. */
  .rail { position: relative; padding-left: 1.9rem; margin-bottom: 1.25rem; }
  .rail::before {
    content: ""; position: absolute; left: 8px; top: 10px; bottom: 10px; width: 1.5px;
    background: var(--color-border);
  }
  .rail-step {
    position: relative; background: var(--color-surface); border: 1px solid var(--color-border);
    border-radius: var(--radius); box-shadow: var(--shadow-card);
    padding: 1.25rem 1.4rem; margin-bottom: .9rem;
    animation: rise .35s var(--ease) both;
  }
  .rail-step::before {
    content: ""; position: absolute; left: -1.9rem; top: 1.4rem; width: 17px; height: 17px;
    border-radius: 50%; background: var(--color-surface); border: 2px solid var(--color-border);
    transition: background-color .2s var(--ease), border-color .2s var(--ease), box-shadow .2s var(--ease);
  }
  .rail-step.done::before { background: var(--color-primary); border-color: var(--color-primary); }
  .rail-step.active::before { border-color: var(--color-primary); box-shadow: 0 0 0 4px var(--color-primary-soft); }
  .rail-head { display: flex; align-items: flex-start; gap: .75rem; cursor: pointer; user-select: none; }
  .rail-step.done .rail-title { color: var(--color-text-muted); font-weight: 500; }
  .rail-summary { font-size: .85rem; color: var(--color-text-muted); margin: .15rem 0 0; min-height: 1.3em; }
  .rail-fold {
    margin-left: auto; flex: none; width: 20px; height: 20px; align-self: center;
    color: var(--color-text-muted); transition: transform .2s var(--ease);
  }
  .rail-step.collapsed .rail-fold { transform: rotate(-90deg); }
  .rail-step.collapsed .rail-body { display: none; }
  .rail-step .rail-body { margin-top: 1.15rem; }
  @media (max-width: 480px) {
    .rail { padding-left: 1.5rem; }
    .rail::before { left: 6px; }
    .rail-step::before { left: -1.5rem; }
  }

  .card-head { display: flex; align-items: flex-start; gap: .75rem; margin-bottom: 1.1rem; }
  .card-head.foldable { cursor: pointer; user-select: none; }
  .card-fold {
    margin-left: auto; flex: none; width: 22px; height: 22px; align-self: center;
    color: var(--color-text-muted); transition: transform .2s var(--ease);
  }
  .card.collapsed .card-head { margin-bottom: 0; }
  .card.collapsed .card-fold { transform: rotate(-90deg); }
  .card.collapsed > *:not(.card-head) { display: none; }
  .card-icon {
    width: 34px; height: 34px; flex: none; border-radius: 10px;
    background: var(--color-primary-soft); color: var(--color-primary);
    display: flex; align-items: center; justify-content: center;
  }
  .card-icon svg { width: 18px; height: 18px; }
  .card h2, .rail-title { font-size: 1.25rem; font-weight: 600; margin: .1rem 0 .2rem; letter-spacing: -.005em; }
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
    min-height: 42px; padding: .5rem 1.1rem; border-radius: var(--radius-sm);
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
  .btn-primary:hover:not(:disabled) { background: var(--color-primary-hover); box-shadow: 0 4px 14px rgba(15,107,99,.25); }

  .btn-accent { background: var(--color-accent); color: var(--color-on-accent); }
  .btn-accent:hover:not(:disabled) { background: var(--color-accent-hover); box-shadow: 0 4px 14px rgba(184,86,14,.25); }

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
    display: inline-flex; align-items: center; gap: .5rem; min-height: 42px; padding: .5rem 1.1rem;
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
  progress::-webkit-progress-value { background: var(--color-primary); border-radius: 999px; }
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
  .stat-value { font-family: var(--font-mono); font-size: 1.4rem; font-weight: 600; line-height: 1.1; font-variant-numeric: tabular-nums; }
  .stat-label { font-size: .72rem; color: var(--color-text-muted); margin-top: .25rem;
                text-transform: uppercase; letter-spacing: .04em; }

  .callout { display: flex; gap: .6rem; align-items: flex-start; padding: .75rem .9rem;
             border-radius: var(--radius-sm); font-size: .85rem; margin: .75rem 0; }
  .callout svg { width: 18px; height: 18px; flex: none; margin-top: .05rem; }
  .callout-danger { background: var(--color-danger-soft); color: var(--color-danger-text); }
  .callout-info { background: var(--color-primary-soft); color: var(--color-primary); }
  .callout-warn { background: var(--color-warning-soft); color: var(--color-warning-text); }

  .diag { margin: .5rem 0 1rem; }
  .diag-table-wrap { overflow-x: auto; margin: 0 -1px; }
  .diag-table { width: 100%; border-collapse: collapse; font-size: .85rem; }
  .diag-table th, .diag-table td { padding: .35rem .5rem; text-align: left;
                                   border-bottom: 1px solid var(--color-border); }
  .diag-table th { color: var(--color-text-muted); font-weight: 600; }
  .diag-table .num { text-align: right; font-family: var(--font-mono); font-variant-numeric: tabular-nums; }
  .diag-table .diag-hit td { background: var(--color-warning-soft); }

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
    /* Marke+Titel und die drei Nav-Links passen nicht mehr nebeneinander,
       sobald "Suchen" als dritter Link dazukam - umbrechen statt die Seite
       ueberbreit werden zu lassen. */
    .topbar { flex-wrap: wrap; }
    .nav { margin-left: 0; width: 100%; }
    .nav a { flex: 1 1 auto; justify-content: center; }
  }

  /* --- Navigation zwischen Kontakten und E-Mail --- */
  .nav { display: flex; gap: .4rem; margin-left: auto; }
  .nav a {
    display: inline-flex; align-items: center; gap: .4rem; text-decoration: none;
    padding: .45rem .85rem; border-radius: 8px; font-size: .88rem; font-weight: 500;
    color: var(--color-text-muted); border: 1px solid transparent;
  }
  .nav a:hover { background: var(--color-surface-alt); color: var(--color-text); }
  .nav a.active {
    background: var(--color-primary-soft); color: var(--color-primary);
    border-color: var(--color-primary); font-weight: 600;
  }
  .nav a svg { width: 16px; height: 16px; }

  /* --- E-Mail: Konten, Ordner, Absendergruppen --- */
  .acc-row {
    display: flex; align-items: center; gap: .6rem; padding: .6rem .8rem;
    border: 1px solid var(--color-border); border-radius: 10px; margin-bottom: .5rem;
    background: var(--color-surface-alt); flex-wrap: wrap;
  }
  .acc-row .acc-name { font-weight: 600; }
  .acc-row .acc-detail { color: var(--color-text-muted); font-size: .85rem; }
  .acc-row .acc-actions { margin-left: auto; display: flex; gap: .4rem; }
  .form-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: .6rem; margin: .8rem 0;
  }
  .form-grid label { display: flex; flex-direction: column; gap: .25rem; font-size: .85rem;
    color: var(--color-text-muted); font-weight: 500; }
  .form-grid input, .form-grid select {
    padding: .5rem .65rem; border-radius: 8px; font: inherit; font-size: .9rem;
    border: 1px solid var(--color-border); background: var(--color-surface);
    color: var(--color-text);
  }
  .form-grid input:focus, .form-grid select:focus {
    outline: 2px solid var(--color-primary); outline-offset: 1px; border-color: transparent;
  }
  .folder-list {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr));
    gap: .3rem; margin: .7rem 0; max-height: 15rem; overflow-y: auto;
    padding: .6rem; border: 1px solid var(--color-border); border-radius: 10px;
  }
  .folder-list label { display: flex; align-items: center; gap: .45rem; font-size: .88rem; }
  .group {
    border: 1px solid var(--color-border); border-radius: 10px; margin-bottom: .6rem;
    overflow: hidden;
  }
  .group-head {
    display: flex; align-items: center; gap: .6rem; padding: .6rem .8rem;
    background: var(--color-surface-alt); cursor: pointer; flex-wrap: wrap;
  }
  .group-head .g-sender { font-weight: 600; font-family: var(--font-mono); font-size: .85rem; }
  .group-head .g-meta { color: var(--color-text-muted); font-size: .82rem; margin-left: auto; }
  .group-head .g-rules { display: flex; gap: .25rem; }
  .group-head .g-rules .btn { padding: .18rem .5rem; font-size: .75rem; }
  .group-mails { display: none; padding: .3rem .8rem .6rem; }
  .group.open .group-mails { display: block; }
  .mail-row {
    display: flex; align-items: flex-start; gap: .55rem; padding: .4rem 0;
    border-top: 1px solid var(--color-border); font-size: .87rem;
  }
  .mail-row:first-child { border-top: none; }
  .mail-row .m-body { flex: 1; min-width: 0; }
  .mail-row .m-subject { display: block; overflow: hidden; text-overflow: ellipsis;
    white-space: nowrap; }
  .mail-row .m-why { color: var(--color-text-muted); font-size: .78rem; }
  .mail-row .m-date { color: var(--color-text-muted); font-size: .78rem;
    font-family: var(--font-mono); white-space: nowrap; }
  .pill {
    display: inline-block; padding: .05rem .4rem; border-radius: 999px;
    font-size: .72rem; font-weight: 600;
  }
  .pill-delete { background: #fee2e2; color: var(--color-danger-text); }
  .pill-unsure { background: #fef3c7; color: var(--color-warning-text); }
  .pill-keep { background: var(--color-primary-soft); color: var(--color-primary); }
  .pill-type { background: var(--color-surface-2); color: var(--color-text-muted);
               border: 1px solid var(--color-border); }
  @media (prefers-color-scheme: dark) {
    .pill-delete { background: #7f1d1d; }
    .pill-unsure { background: #78350f; }
  }
  :root[data-theme="dark"] .pill-delete { background: #7f1d1d; color: #fecaca; }
  :root[data-theme="dark"] .pill-unsure { background: #78350f; color: #fde68a; }
  :root[data-theme="light"] .pill-delete { background: #fee2e2; color: #b91c1c; }
  :root[data-theme="light"] .pill-unsure { background: #fef3c7; color: #92400e; }
</style>
"""

# Kopfbereich einmal definiert und von beiden Seiten (Kontakte, E-Mail) genutzt -
# so bleibt das Aussehen automatisch einheitlich, wenn am Stil etwas geaendert wird.
_HEAD_BEFORE_TITLE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>"""
_HEAD_AFTER_TITLE = """</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
"""


def _head(title):
    return _HEAD_BEFORE_TITLE + title + _HEAD_AFTER_TITLE + BASE_CSS + "\n</head>\n<body>\n"


_NAV_LINKS = (
    ("/", "Kontakte",
     '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/>'),
    ("/mail", "E-Mail",
     '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>'),
    ("/mail/suche", "Suchen",
     '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.35-4.35"/>'),
)


def _nav(active_href):
    """Dieselbe Navigation auf allen drei Seiten (Kontakte, E-Mail, Suchen) -
    einmal definiert, damit sie beim Ändern nicht auseinanderlaufen."""
    links = "\n      ".join(
        '<a href="{href}"{cls}>\n'
        '        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
        'stroke-linecap="round" stroke-linejoin="round">{icon}</svg>\n'
        '        {label}\n      </a>'.format(
            href=href, label=label, icon=icon,
            cls=' class="active"' if href == active_href else "")
        for href, label, icon in _NAV_LINKS)
    return f'<nav class="nav">\n      {links}\n    </nav>'


PAGE = _head("iCloud Kontakte Sync") + """
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
    __NAV__
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
  if(!dry){
    if(!confirm('Wirklich jetzt schreiben? Dabei können Kontakte angelegt, geändert oder gelöscht werden.')) return;
    document.getElementById('btnReal').disabled = true;
  }
  const r = await fetch('/api/import', {method:'POST', headers:{'Content-Type':'application/json'},
                        body: JSON.stringify({dry_run: dry, target: selectedTarget()})});
  if(r.ok){ setBusy(true); poll(); }
  else if(!dry){ document.getElementById('btnReal').disabled = false; }
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
      ? '<h3 class="hl-title">Warnungen</h3><div class="hl">'+ highlightLog(j.result.warnings.join('\\n')) +'</div>' : '';
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
    ? '<h3 class="hl-title">'+label+' – Details</h3><div class="hl">'+ highlightLog(res.highlights.join('\\n')) +'</div>' : '';
  let changeList = (res.changes && res.changes.length)
    ? '<h3 class="hl-title">'+label+' – geänderte Kontakte</h3><div class="hl">'+ esc(res.changes.join('\\n')) +'</div>' : '';
  let firstTile = showChanged
    ? '<div class="stat-tile"><div class="stat-value">'+res.changed+'</div><div class="stat-label">Geändert</div></div>'
    : '<div class="stat-tile"><div class="stat-value">'+res.updated+'</div><div class="stat-label">Aktualisiert</div></div>';
  let summary = showChanged
    ? res.updated+' aktualisiert (davon '+res.changed+' inhaltlich geändert), '+(res.unchanged||0)+' bereits aktuell (übersprungen).'
    : res.updated+' aktualisiert, '+(res.unchanged||0)+' bereits aktuell (übersprungen).';
  return '<h3 class="hl-title" style="margin-top:1.1rem">'+label+'</h3>' +
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
""".replace("__NAV__", _nav("/"))


MAIL_PAGE = _head("E-Mail aufräumen") + """
<div class="page">

  <div class="topbar">
    <div class="brand-mark">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>
      </svg>
    </div>
    <div>
      <h1>E-Mail aufräumen</h1>
      <p class="sub">Vorschläge prüfen, dann in den Papierkorb verschieben</p>
    </div>
    __NAV__
  </div>

  <div class="callout callout-info">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>
    </svg>
    <div>Nichts wird endgültig gelöscht: Mails wandern in den <strong>Papierkorb</strong>
    des jeweiligen Kontos und lassen sich dort zurückholen.</div>
  </div>

  <div class="rail">

    <div class="rail-step active" id="railStep1">
      <div class="rail-head" onclick="toggleRailStep(1)" role="button" tabindex="0"
           aria-expanded="true" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleRailStep(1);}">
        <div class="card-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>
          </svg>
        </div>
        <div>
          <h2 class="rail-title">Postfächer</h2>
          <p class="rail-summary" id="railSum1">IMAP-Konten mit app-spezifischem Passwort. Die Zugangsdaten
             bleiben lokal in <code>mail_accounts.json</code>.</p>
        </div>
        <svg class="rail-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
      </div>
      <div class="rail-body">
        <div id="accountList"></div>
        <button class="btn btn-secondary" onclick="toggleAccountForm()">Konto hinzufügen</button>
        <div id="accountForm" class="hidden">
          <div class="form-grid">
            <label>Anzeigename<input id="accName" placeholder="iCloud"></label>
            <label>Anbieter
              <select id="accPreset" onchange="applyPreset()">
                <option value="">eigener Server</option>
              </select>
            </label>
            <label>IMAP-Server<input id="accHost" placeholder="imap.mail.me.com"></label>
            <label>Port<input id="accPort" value="993"></label>
            <label>Benutzername<input id="accUser" placeholder="name@icloud.com"></label>
            <label>Passwort<input id="accPass" type="password" placeholder="app-spezifisch"></label>
          </div>
          <button class="btn btn-primary" onclick="addAccount()">Speichern und testen</button>
          <p id="accInfo" class="muted"></p>
        </div>
      </div>
    </div>

    <div class="rail-step" id="railStep2">
      <div class="rail-head" onclick="toggleRailStep(2)" role="button" tabindex="0"
           aria-expanded="true" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleRailStep(2);}">
        <div class="card-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>
          </svg>
        </div>
        <div>
          <h2 class="rail-title">Durchsuchen</h2>
          <p class="rail-summary" id="railSum2">Ordner auswählen und bewerten lassen. Dabei wird
             ausschließlich gelesen - es wird nichts verändert.</p>
        </div>
        <svg class="rail-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
      </div>
      <div class="rail-body">
        <div id="folderArea" class="muted">Erst ein Konto anlegen.</div>
        <div class="form-grid" style="max-width:240px">
          <label>Mindestalter in Tagen
            <input id="minAge" type="number" min="0" value="30">
          </label>
        </div>
        <button id="btnScan" class="btn btn-primary" onclick="startScan()">
          Ausgewählte Ordner durchsuchen
        </button>
      </div>
    </div>

    <div class="rail-step" id="railStep3">
      <div class="rail-head" onclick="toggleRailStep(3)" role="button" tabindex="0"
           aria-expanded="true" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleRailStep(3);}">
        <div class="card-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <path d="M9 11l3 3 8-8"/><path d="M20 12v7a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h9"/>
          </svg>
        </div>
        <div>
          <h2 class="rail-title">Prüfen</h2>
          <p class="rail-summary" id="railSum3">Vorgeschlagene Mails abhaken - hier im Browser oder
             bequemer in Excel. Vorangehakt ist nur, was deutlich weg kann.</p>
        </div>
        <svg class="rail-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
      </div>
      <div class="rail-body">
        <div class="download-row">
          <a href="/api/mail/download/xlsx" class="btn btn-secondary">Als Excel herunterladen</a>
          <label class="file-label" for="mailFile">Bearbeitete Excel hochladen</label>
          <input type="file" id="mailFile" accept=".xlsx" onchange="uploadExcel()">
          <button class="btn btn-secondary" onclick="discardScan()">Liste verwerfen</button>
        </div>
        <p id="uploadInfo" class="muted"></p>
        <fieldset class="target-group" id="sortBox" style="display:none">
          <legend class="muted">Sortierung</legend>
          <label class="target-option">
            <input type="radio" name="sortMode" value="count" checked onchange="renderGroups()"> nach Absender
          </label>
          <label class="target-option">
            <input type="radio" name="sortMode" value="size" onchange="renderGroups()"> nach Größe
          </label>
          <label class="target-option">
            <input type="radio" name="sortMode" value="attachment" onchange="renderGroups()"> nur mit Anhang
          </label>
        </fieldset>
        <div id="resultArea"><p class="muted">Noch kein Scan vorhanden.</p></div>
      </div>
    </div>

    <div class="rail-step" id="railStep4">
      <div class="rail-head" onclick="toggleRailStep(4)" role="button" tabindex="0"
           aria-expanded="true" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleRailStep(4);}">
        <div class="card-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
          </svg>
        </div>
        <div>
          <h2 class="rail-title">Aufräumen</h2>
          <p class="rail-summary" id="railSum4">Erst der Testlauf, dann wirklich verschieben.</p>
        </div>
        <svg class="rail-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
      </div>
      <div class="rail-body">
        <div class="button-row">
          <button class="btn btn-secondary" onclick="startClean(true)">
            Testlauf (nichts wird verschoben)
          </button>
          <button id="btnRealClean" class="btn btn-accent" onclick="startClean(false)" disabled>
            Wirklich in den Papierkorb
          </button>
        </div>
      </div>
    </div>

  </div>

  <div id="status" aria-live="polite"></div>

  <div class="card collapsed" id="rulesCard">
    <div class="card-head foldable" role="button" tabindex="0" aria-expanded="false" onclick="toggleCard(this)" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleCard(this);}">
      <div class="card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M4 6h16M4 12h16M4 18h10"/><circle cx="18" cy="18" r="3"/>
        </svg>
      </div>
      <div>
        <h2>Eigene Regeln</h2>
        <p class="muted desc">Feste Anweisungen, die über dem Gelernten stehen.
           Am bequemsten legst du sie oben direkt an einem Absender an.</p>
      </div>
      <svg class="card-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
    </div>
    <div class="form-grid">
      <label>Liste
        <select id="ruleList">
          <option value="never_delete">nie löschen (Absender)</option>
          <option value="always_delete">immer löschen (Absender)</option>
          <option value="subject_never">nie löschen (Betreff enthält)</option>
          <option value="subject_delete">eher löschen (Betreff enthält)</option>
        </select>
      </label>
      <label>Wert<input id="ruleValue" placeholder="news@shop.de oder *@werbung.de"></label>
    </div>
    <button class="btn btn-secondary" onclick="addRuleFromForm()">Regel hinzufügen</button>
    <div id="rulesArea"></div>
  </div>

  <div class="card collapsed" id="learnedCard">
    <div class="card-head foldable" role="button" tabindex="0" aria-expanded="false" onclick="toggleCard(this)" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleCard(this);}">
      <div class="card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 3v3m0 12v3M3 12h3m12 0h3M5.6 5.6l2.1 2.1m8.6 8.6 2.1 2.1m0-12.8-2.1 2.1m-8.6 8.6-2.1 2.1"/>
        </svg>
      </div>
      <div>
        <h2>Gelernt</h2>
        <p class="muted desc">Was das Tool aus deinen bisherigen Entscheidungen
           mitgenommen hat. Wird nach jedem echten Aufräumen ergänzt.</p>
      </div>
      <svg class="card-fold" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
    </div>
    <div id="learnedArea"><p class="muted">Noch nichts gelernt.</p></div>
  </div>

</div>

<script>
let polling = null;
let selected = new Set();
let presets = {};

const I_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>';
const I_BAD = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>';
const I_WARN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4m0 4h.01M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/></svg>';

function esc(s){
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}
function kb(bytes){
  if(bytes >= 1048576) return (bytes/1048576).toFixed(1) + ' MB';
  return Math.round(bytes/1024) + ' KB';
}
async function jget(url){
  const r = await fetch(url);
  return await r.json();
}
async function jpost(url, body, method){
  const r = await fetch(url, {method: method || 'POST',
    headers: {'Content-Type':'application/json'}, body: JSON.stringify(body || {})});
  return {ok: r.ok, data: await r.json()};
}

/* ---------- Konten ---------- */
async function loadAccounts(){
  const j = await jget('/api/mail/accounts');
  presets = j.presets || {};
  const sel = document.getElementById('accPreset');
  if(sel.options.length <= 1){
    Object.keys(presets).forEach(name => {
      const o = document.createElement('option');
      o.value = name; o.textContent = name;
      sel.appendChild(o);
    });
  }
  const list = document.getElementById('accountList');
  if(!j.accounts.length){
    list.innerHTML = '<p class="muted">Noch kein Konto eingerichtet.</p>';
  } else {
    list.innerHTML = j.accounts.map(a =>
      '<div class="acc-row"><span class="acc-name">' + esc(a.name) + '</span>' +
      '<span class="acc-detail">' + esc(a.user) + ' · ' + esc(a.host) + ':' + a.port + '</span>' +
      '<span class="acc-actions">' +
      '<button class="btn btn-secondary" onclick="diagnose(\\'' + esc(a.name) + '\\')">Postfach prüfen</button> ' +
      '<button class="btn btn-secondary" onclick="removeAccount(\\'' + esc(a.name) + '\\')">Entfernen</button>' +
      '</span></div>' +
      '<div class="diag hidden" id="diag-' + esc(a.name) + '"></div>').join('');
  }
  renderFolderArea(j.accounts);
  if(j.accounts.length){
    setRailState(1, 'done', j.accounts.length === 1 ? '1 Konto verbunden'
      : j.accounts.length + ' Konten verbunden');
  } else {
    setRailState(1, 'active', '');
  }
  refreshRailChain();
}

/* ---------- Postfach prüfen ----------
   Beantwortet "warum sehe ich im Mailprogramm keine Löschungen?" - liest nur. */
async function diagnose(name){
  const box = document.getElementById('diag-' + name);
  box.classList.remove('hidden');
  box.innerHTML = '<p class="muted">Prüfe Postfach ...</p>';
  const j = await jget('/api/mail/diagnose?account=' + encodeURIComponent(name));
  if(j.error){ box.innerHTML = '<span class="l-err">' + esc(j.error) + '</span>'; return; }

  let head = '';
  if(j.method === 'copy_flag'){
    head = '<div class="callout callout-warn">' + I_WARN + '<div><b>Dieser Server ' +
      'kann Mails nicht selbst verschieben.</b> Beim Aufräumen werden sie in den ' +
      'Papierkorb kopiert und im Ordner nur als gelöscht markiert — sie bleiben ' +
      'liegen. Viele Mailprogramme zeigen solche Mails ganz normal an. Genau so ' +
      'sieht es aus, als wäre nichts passiert.</div></div>';
  } else {
    head = '<div class="callout callout-info">' + I_OK + '<div>Der Server kann Mails ' +
      'selbst verschieben (' + esc(j.method_text) + ').</div></div>';
  }
  if(!j.trash){
    head += '<div class="callout callout-warn">' + I_WARN + '<div>Es wurde <b>kein ' +
      'Papierkorb gefunden</b> — ohne Ziel kann nichts verschoben werden.</div></div>';
  } else if(j.trash_via === 'name'){
    head += '<div class="callout callout-warn">' + I_WARN + '<div>Der Papierkorb ' +
      '„' + esc(j.trash) + '" wurde <b>über den Namen geraten</b>, der Server nennt ' +
      'ihn nicht selbst. Bitte im Mailprogramm nachsehen, ob das der richtige ' +
      'Ordner ist.</div></div>';
  }

  box.innerHTML = head +
    '<h3 class="hl-title">Ordner</h3><div class="diag-table-wrap"><table class="diag-table"><thead><tr>' +
    '<th>Ordner</th><th class="num">Mails</th><th class="num">nur markiert</th></tr></thead><tbody>' +
    j.folders.map(f =>
      '<tr' + (f.deleted ? ' class="diag-hit"' : '') + '><td>' + esc(f.name) +
      (f.is_trash ? ' <span class="muted">(Papierkorb)</span>' : '') + '</td>' +
      '<td class="num">' + (f.error ? '?' : f.total) + '</td>' +
      '<td class="num">' + (f.error ? esc(f.error) : f.deleted) + '</td></tr>').join('') +
    '</tbody></table></div>' +
    '<p class="muted">„Nur markiert" heißt: als gelöscht gekennzeichnet, aber noch ' +
    'im Ordner. Stehen hier Zahlen, sind das genau die Mails, die im Mailprogramm ' +
    'noch auftauchen.</p>';
}
function toggleAccountForm(){
  document.getElementById('accountForm').classList.toggle('hidden');
}
function applyPreset(){
  const p = presets[document.getElementById('accPreset').value];
  if(p){
    document.getElementById('accHost').value = p.host;
    document.getElementById('accPort').value = p.port;
  }
}
async function addAccount(){
  const info = document.getElementById('accInfo');
  info.textContent = 'Verbinde ...';
  const res = await jpost('/api/mail/accounts', {
    name: document.getElementById('accName').value.trim(),
    host: document.getElementById('accHost').value.trim(),
    port: document.getElementById('accPort').value.trim() || 993,
    user: document.getElementById('accUser').value.trim(),
    password: document.getElementById('accPass').value
  });
  if(!res.ok){ info.textContent = 'Fehler: ' + res.data.error; return; }
  if(res.data.tested){
    info.textContent = 'Verbunden: ' + res.data.folders + ' Ordner, Papierkorb: ' +
      (res.data.trash || 'nicht gefunden!');
  } else {
    info.textContent = 'Gespeichert, aber Verbindung fehlgeschlagen: ' + res.data.message;
  }
  document.getElementById('accPass').value = '';
  loadAccounts();
}
async function removeAccount(name){
  if(!confirm('Konto "' + name + '" entfernen?')) return;
  await jpost('/api/mail/accounts', {name: name}, 'DELETE');
  loadAccounts();
}

/* ---------- Ordner ---------- */
function renderFolderArea(accounts){
  const area = document.getElementById('folderArea');
  if(!accounts.length){ area.innerHTML = '<p class="muted">Erst ein Konto anlegen.</p>'; return; }
  area.innerHTML = accounts.map(a =>
    '<div data-account="' + esc(a.name) + '">' +
    '<h3 class="hl-title">' + esc(a.name) +
    ' <button class="btn btn-secondary" onclick="loadFolders(\\'' + esc(a.name) + '\\')">Ordner laden</button>' +
    ' <button class="btn btn-secondary" onclick="setFolders(\\'' + esc(a.name) + '\\', true)">Alle</button>' +
    ' <button class="btn btn-secondary" onclick="setFolders(\\'' + esc(a.name) + '\\', false)">Keine</button></h3>' +
    '<div class="folder-list hidden" id="fl-' + esc(a.name) + '"></div></div>').join('');
}
async function loadFolders(name){
  const box = document.getElementById('fl-' + name);
  box.classList.remove('hidden');
  box.innerHTML = '<span class="muted">Lade ...</span>';
  const j = await jget('/api/mail/folders?account=' + encodeURIComponent(name));
  if(j.error){ box.innerHTML = '<span class="l-err">' + esc(j.error) + '</span>'; return; }
  box.innerHTML = j.folders.map(f =>
    '<label><input type="checkbox" class="folder-cb" data-account="' + esc(name) + '"' +
    ' value="' + esc(f.name) + '"' + (f.skip ? '' : ' checked') + '> ' +
    esc(f.name) + (f.skip ? ' <span class="muted">(übersprungen)</span>' : '') + '</label>').join('');
}
function setFolders(name, checked){
  const box = document.getElementById('fl-' + name);
  if(!box || box.classList.contains('hidden')){ loadFolders(name).then(() => setFolders(name, checked)); return; }
  box.querySelectorAll('.folder-cb').forEach(cb => { cb.checked = checked; });
}

/* ---------- Zuklappbare Karten (Eigene Regeln, Gelernt) ---------- */
function toggleCard(head){
  const collapsed = head.closest('.card').classList.toggle('collapsed');
  head.setAttribute('aria-expanded', String(!collapsed));
}

/* ---------- Ablauf-Rail: Zustand je Schritt ----------
   Nur Schritte mit einem klaren Ja/Nein-Abschlusssignal (Konto vorhanden,
   Scan geladen) klappen automatisch ein - Prüfen und Aufräumen bleiben immer
   offen, weil es dort kein "fertig" gibt, nur "gerade dran". Ein Nutzer, der
   einen erledigten Schritt von Hand wieder aufklappt, bekommt keinen
   automatischen Rückfall - erst ein erneuter Zustandswechsel überschreibt das. */
function toggleRailStep(n){
  const el = document.getElementById('railStep' + n);
  const collapsed = el.classList.toggle('collapsed');
  el.dataset.userOpened = collapsed ? '' : '1';
  el.querySelector('.rail-head').setAttribute('aria-expanded', String(!collapsed));
}
function setRailState(n, state, summary){
  const el = document.getElementById('railStep' + n);
  el.classList.remove('done', 'active');
  if(state) el.classList.add(state);
  if(summary !== undefined) document.getElementById('railSum' + n).textContent = summary;
  if(state === 'done' && !el.dataset.userOpened){
    el.classList.add('collapsed');
    el.querySelector('.rail-head').setAttribute('aria-expanded', 'false');
  } else if(state !== 'done'){
    el.classList.remove('collapsed');
    el.querySelector('.rail-head').setAttribute('aria-expanded', 'true');
  }
}
function refreshRailChain(){
  const step1Done = document.getElementById('railStep1').classList.contains('done');
  const step2Done = document.getElementById('railStep2').classList.contains('done');
  const reached = step1Done && step2Done;
  setRailState(3, reached ? 'active' : '');
  setRailState(4, reached ? 'active' : '');
}

/* ---------- Scan ---------- */
async function startScan(){
  const byAccount = {};
  document.querySelectorAll('.folder-cb:checked').forEach(cb => {
    const a = cb.dataset.account;
    if(!byAccount[a]) byAccount[a] = [];
    byAccount[a].push(cb.value);
  });
  const selection = Object.keys(byAccount).map(a => ({account: a, folders: byAccount[a]}));
  if(!selection.length){ alert('Bitte zuerst Ordner auswählen (Ordner laden, dann anhaken).'); return; }
  const res = await jpost('/api/mail/scan', {
    selection: selection,
    min_age: document.getElementById('minAge').value
  });
  if(!res.ok){ alert(res.data.error); return; }
  poll();
}

/* ---------- Ergebnis ---------- */
let groupData = [];
let summaryData = null;
let scanInfo = {};       // Wann eingelesen, wann zuletzt aufgeräumt

/* Die Liste ist eine Momentaufnahme des Postfachs, kein Live-Blick hinein.
   Passt sie nicht mehr, wirft man sie weg und liest neu ein. */
async function discardScan(){
  if(!confirm('Gespeicherte Liste verwerfen? Im Postfach ändert sich dadurch ' +
              'nichts — du musst danach neu durchsuchen.')) return;
  await jpost('/api/mail/scan', {}, 'DELETE');
  groupData = []; summaryData = null; selected = new Set(); scanInfo = {};
  document.getElementById('sortBox').style.display = 'none';
  document.getElementById('uploadInfo').textContent = '';
  document.getElementById('resultArea').innerHTML =
    '<p class="muted">Liste verworfen. Oben unter „Durchsuchen" neu einlesen.</p>';
  document.getElementById('btnRealClean').disabled = true;
  const step1Done = document.getElementById('railStep1').classList.contains('done');
  setRailState(2, step1Done ? 'active' : '', '');
  refreshRailChain();
}

async function loadResult(){
  const j = await jget('/api/mail/result');
  const area = document.getElementById('resultArea');
  if(j.error){
    area.innerHTML = '<p class="muted">' + esc(j.error) + '</p>';
    document.getElementById('sortBox').style.display = 'none';
    groupData = []; summaryData = null;
    const step1Done = document.getElementById('railStep1').classList.contains('done');
    setRailState(2, step1Done ? 'active' : '', '');
    refreshRailChain();
    return;
  }
  groupData = j.groups;
  summaryData = j.summary;
  scanInfo = {created: j.created, executed: j.executed};
  selected = new Set();
  groupData.forEach(g => g.mails.forEach(m => { if(m.delete) selected.add(m.key); }));
  document.getElementById('sortBox').style.display = groupData.length ? '' : 'none';
  renderGroups();
  setRailState(2, 'done', summaryData.total + ' Mails geprüft, ' + summaryData.delete + ' vorgeschlagen');
  refreshRailChain();
}

function sortMode(){
  const el = document.querySelector('input[name="sortMode"]:checked');
  return el ? el.value : 'count';
}

function selectedBytes(){
  let total = 0;
  groupData.forEach(g => g.mails.forEach(m => { if(selected.has(m.key)) total += m.size; }));
  return total;
}

function stamp(iso){
  if(!iso) return '';
  const d = new Date(iso);
  return isNaN(d) ? '' : d.toLocaleString('de-DE',
    {day: '2-digit', month: '2-digit', year: 'numeric',
     hour: '2-digit', minute: '2-digit'});
}

function renderGroups(){
  const area = document.getElementById('resultArea');
  if(!summaryData) return;
  const s = summaryData;

  // Nach einem echten Lauf hat sich das Postfach geändert - die Liste zeigt
  // dann zwangsläufig einen alten Stand. Das muss dranstehen, sonst hält man
  // sie für den Blick ins Postfach, der sie nicht ist.
  let html = '';
  if(scanInfo.executed){
    html += '<div class="callout callout-warn">' + I_WARN + '<div>Diese Liste ist ' +
      'vom ' + esc(stamp(scanInfo.created)) + ' und wurde nach dem Aufräumen ' +
      '(' + esc(stamp(scanInfo.executed)) + ') <b>nicht neu eingelesen</b> — sie ' +
      'kann vom Postfach abweichen. Oben unter „Durchsuchen" neu einlesen ' +
      'oder die Liste verwerfen.</div></div>';
  } else if(scanInfo.created){
    html += '<p class="muted">Stand: ' + esc(stamp(scanInfo.created)) + '</p>';
  }

  html += '<div class="stat-grid">' +
    '<div class="stat-tile"><div class="stat-value">' + s.total + '</div><div class="stat-label">Geprüft</div></div>' +
    '<div class="stat-tile"><div class="stat-value">' + s.delete + '</div><div class="stat-label">Vorgeschlagen</div></div>' +
    '<div class="stat-tile"><div class="stat-value" id="selCount">' + selected.size + '</div><div class="stat-label">Ausgewählt</div></div>' +
    '<div class="stat-tile"><div class="stat-value" id="selBytes">' + kb(selectedBytes()) + '</div><div class="stat-label">Wird frei</div></div>' +
    '</div>';

  let groups = groupData.slice();
  const mode = sortMode();
  if(mode === 'size'){
    groups.sort((a, b) => b.bytes - a.bytes);
  } else if(mode === 'attachment'){
    groups = groups.filter(g => g.attachments > 0);
    groups.sort((a, b) => b.bytes - a.bytes);
  } else {
    groups.sort((a, b) => b.mails.length - a.mails.length || a.sender.localeCompare(b.sender));
  }

  if(!groups.length){
    area.innerHTML = html + '<p class="muted">' +
      (mode === 'attachment' ? 'Keine Vorschläge mit Anhang.' : 'Nichts zum Aufräumen gefunden.') +
      '</p>';
    updateCleanButton();
    return;
  }

  html += groups.map((g) => {
    const gi = groupData.indexOf(g);
    const mails = (mode === 'attachment') ? g.mails.filter(m => m.attachment) : g.mails;
    return '<div class="group" id="g' + gi + '">' +
      '<div class="group-head" onclick="toggleGroup(' + gi + ', event)">' +
        '<input type="checkbox" onclick="toggleAll(' + gi + ', this, event)"' +
          (mails.every(m => selected.has(m.key)) ? ' checked' : '') + '>' +
        '<span class="g-sender">' + esc(g.sender) + '</span>' +
        (g.name ? '<span class="muted">' + esc(g.name) + '</span>' : '') +
        '<span class="g-meta">' + mails.length + (mails.length === 1 ? ' Mail · ' : ' Mails · ') +
          kb(g.bytes) + (g.attachments ? ' · ' + g.attachments + '× Anhang' : '') + '</span>' +
        '<span class="g-rules">' +
          '<button class="btn btn-secondary" title="Diesen Absender künftig immer vorschlagen"' +
            ' onclick="ruleForSender(\\'' + esc(g.sender) + '\\', \\'always_delete\\', event)">immer</button>' +
          '<button class="btn btn-secondary" title="Diesen Absender künftig nie vorschlagen"' +
            ' onclick="ruleForSender(\\'' + esc(g.sender) + '\\', \\'never_delete\\', event)">nie</button>' +
        '</span>' +
      '</div>' +
      '<div class="group-mails">' + mails.map(m =>
        '<label class="mail-row">' +
          '<input type="checkbox" class="mail-cb" data-group="' + gi + '" value="' + esc(m.key) + '"' +
            (selected.has(m.key) ? ' checked' : '') + ' onchange="onMailToggle(this)">' +
          '<span class="m-body">' +
            '<span class="m-subject">' +
              (m.attachment ? '<span title="' + esc(m.attachments.join(', ')) + '">📎 </span>' : '') +
              esc(m.subject || '(kein Betreff)') + '</span>' +
            '<span class="m-why"><span class="pill pill-' + m.recommendation + '">' +
              (m.recommendation === 'delete' ? 'löschen'
                : (m.recommendation === 'unsure' ? 'unklar' : 'selbst gewählt')) + '</span> ' +
              (m.mail_type ? '<span class="pill pill-type">' + esc(MAIL_TYPE_LABELS[m.mail_type] || m.mail_type) + '</span> ' : '') +
              esc(m.reasons.join(', ')) + ' · ' + esc(m.folder) +
              (m.attachments.length ? ' · ' + esc(m.attachments.join(', ')) : '') + '</span>' +
          '</span>' +
          '<span class="m-date">' + esc(m.date) + '<br>' + kb(m.size) + '</span>' +
        '</label>').join('') +
      '</div>' +
    '</div>';
  }).join('');

  area.innerHTML = html;
  updateCleanButton();
}
function toggleGroup(gi, ev){
  if(ev.target.tagName === 'INPUT') return;
  document.getElementById('g' + gi).classList.toggle('open');
}
function toggleAll(gi, cb, ev){
  ev.stopPropagation();
  document.querySelectorAll('.mail-cb[data-group="' + gi + '"]').forEach(m => {
    m.checked = cb.checked;
    if(cb.checked) selected.add(m.value); else selected.delete(m.value);
  });
  refreshCount();
}
function onMailToggle(cb){
  if(cb.checked) selected.add(cb.value); else selected.delete(cb.value);
  refreshCount();
}
function refreshCount(){
  const el = document.getElementById('selCount');
  if(el) el.textContent = selected.size;
  const b = document.getElementById('selBytes');
  if(b) b.textContent = kb(selectedBytes());
  updateCleanButton();
}
function updateCleanButton(){
  document.getElementById('btnRealClean').disabled = true;
}

/* ---------- Excel ---------- */
async function uploadExcel(){
  const f = document.getElementById('mailFile').files[0];
  const info = document.getElementById('uploadInfo');
  if(!f) return;
  const fd = new FormData();
  fd.append('file', f);
  info.textContent = 'Lese Datei ...';
  const r = await fetch('/api/mail/upload', {method: 'POST', body: fd});
  const j = await r.json();
  if(!r.ok){ info.textContent = 'Fehler: ' + j.error; return; }
  info.textContent = j.gelesen + ' Zeilen gelesen, ' + j['geändert'] +
    ' Häkchen geändert, ' + j['ausgewählt'] + ' Mails ausgewählt.';
  loadResult();
}

/* ---------- Aufräumen ---------- */
async function startClean(dryRun){
  const res = await jpost('/api/mail/select', {keys: Array.from(selected)});
  if(!res.ok){ alert(res.data.error); return; }
  if(!dryRun && !confirm('Wirklich ' + selected.size +
      ' Mails in den Papierkorb verschieben?')) return;
  const started = await jpost('/api/mail/clean', {dry_run: dryRun});
  if(!started.ok){ alert(started.data.error); return; }
  poll();
}

/* ---------- Status ---------- */
function poll(){
  if(polling) clearInterval(polling);
  polling = setInterval(refresh, 700);
  refresh();
}
async function refresh(){
  const j = await jget('/api/status');
  const s = document.getElementById('status');
  const busy = j.running && (j.kind === 'mailscan' || j.kind === 'mailclean');

  document.getElementById('btnScan').disabled = j.running;
  if(busy){
    s.innerHTML = '<div class="card"><div class="status-head">' +
      '<span class="spinner"></span><span class="status-title">' +
      (j.kind === 'mailscan' ? 'Durchsuche Postfächer ...' : 'Räume auf ...') +
      '</span></div><p class="muted">' + esc(j.message) + '</p></div>';
    return;
  }
  if(polling){ clearInterval(polling); polling = null; }
  if(j.error){
    s.innerHTML = '<div class="card"><div class="status-head bad">' + I_BAD +
      '<span class="status-title">Fehler</span></div><p>' + esc(j.error) + '</p></div>';
    return;
  }
  if(!j.result || !j.result.kind){ s.innerHTML = ''; return; }

  if(j.result.kind === 'scan'){
    s.innerHTML = '<div class="card"><div class="status-head ok">' + I_OK +
      '<span class="status-title">Durchsuchen fertig</span></div>' +
      '<p class="muted">' + j.result.total + ' Mails geprüft, ' + j.result.delete +
      ' vorgeschlagen. Bitte unten prüfen.</p></div>';
    loadResult();
  } else {
    const r = j.result;
    const flagged = r.flagged || 0, failed = r.failed || 0;
    // "Aufgeräumt" nur, wenn wirklich etwas wegkam und nichts liegen blieb.
    const clean_ok = r.moved > 0 && !failed && !flagged;
    const head = r.dry_run ? 'Testlauf-Ergebnis (nichts verschoben)'
               : clean_ok ? 'Aufgeräumt'
               : 'Durchgelaufen — bitte nachlesen';
    const trash = Object.entries(r.trash || {})
      .map(([acc, t]) => acc + ' → ' + (t || 'kein Papierkorb!')).join(', ');

    s.innerHTML = '<div class="card"><div class="status-head ' +
      (r.dry_run || !clean_ok ? 'warn-text' : 'ok') + '">' +
      (r.dry_run || !clean_ok ? I_WARN : I_OK) +
      '<span class="status-title">' + head + '</span></div>' +
      (trash ? '<p class="muted">Papierkorb: ' + esc(trash) + '</p>' : '') +
      '<div class="stat-grid">' +
      '<div class="stat-tile"><div class="stat-value">' + r.moved + '</div><div class="stat-label">' +
        (r.dry_run ? 'Würden weg' : 'Nachweislich weg') + '</div></div>' +
      (flagged ? '<div class="stat-tile stat-bad"><div class="stat-value">' + flagged +
        '</div><div class="stat-label">Nur markiert</div></div>' : '') +
      (failed ? '<div class="stat-tile stat-bad"><div class="stat-value">' + failed +
        '</div><div class="stat-label">Noch da</div></div>' : '') +
      '<div class="stat-tile' + (r.skipped ? ' stat-bad' : '') + '"><div class="stat-value">' +
        r.skipped + '</div><div class="stat-label">Übersprungen</div></div>' +
      '<div class="stat-tile' + (r.errors ? ' stat-bad' : '') + '"><div class="stat-value">' +
        r.errors + '</div><div class="stat-label">Fehler</div></div>' +
      '<div class="stat-tile"><div class="stat-value">' + r.total + '</div><div class="stat-label">Ausgewählt</div></div>' +
      '</div>' +
      (flagged ? '<div class="callout callout-warn">' + I_WARN + '<div><b>Dein Server ' +
        'kann Mails nicht selbst verschieben.</b> Die ' + flagged + ' Mails liegen ' +
        'jetzt als Kopie im Papierkorb, im Ursprungsordner sind sie nur als gelöscht ' +
        'markiert — deshalb zeigen viele Mailprogramme sie weiter an. Im Mailprogramm ' +
        'hilft „Ordner aufräumen" bzw. „Gelöschte endgültig entfernen".</div></div>' : '') +
      (failed ? '<div class="callout callout-danger">' + I_BAD + '<div><b>' + failed +
        ' Mails sind noch da.</b> Der Server hat den Befehl mit Erfolg beantwortet, ' +
        'beim Nachsehen lagen die Mails aber unverändert im Ordner. Über „Postfach ' +
        'prüfen" oben lässt sich nachsehen, woran es liegt.</div></div>' : '') +
      (r.dry_run
        ? '<div class="callout callout-info">' + I_OK +
          '<div>Sieht das gut aus? Dann auf „Wirklich in den Papierkorb" klicken.</div></div>'
        : '') +
      (r.log.length ? '<h3 class="hl-title">Protokoll</h3><div class="hl">' +
        esc(r.log.join('\\n')) + '</div>' : '') +
      '</div>';
    document.getElementById('btnRealClean').disabled = !r.dry_run || !r.moved;
    if(!r.dry_run){ loadResult(); loadLearned(); }
  }
}

const MAIL_TYPE_LABELS = {
  invoice: '📄 Rechnung/Bestellung',
  official: '🏦 Bank/Versicherung/Behörde',
};

/* ---------- Eigene Regeln ---------- */
const RULE_LABELS = {
  never_delete: 'nie löschen (Absender)',
  always_delete: 'immer löschen (Absender)',
  subject_never: 'nie löschen (Betreff enthält)',
  subject_delete: 'eher löschen (Betreff enthält)'
};

async function loadRules(){
  const j = await jget('/api/mail/rules');
  renderRules(j.rules);
}
function renderRules(rules){
  const area = document.getElementById('rulesArea');
  const names = Object.keys(RULE_LABELS).filter(n => (rules[n] || []).length);
  if(!names.length){
    area.innerHTML = '<p class="muted">Noch keine eigenen Regeln.</p>';
    return;
  }
  area.innerHTML = names.map(n =>
    '<h3 class="hl-title">' + RULE_LABELS[n] + '</h3>' +
    rules[n].map(v =>
      '<div class="acc-row"><span class="acc-name">' + esc(v) + '</span>' +
      '<span class="acc-actions"><button class="btn btn-secondary" onclick="removeRule(\\'' +
      esc(v).replace(/'/g, "\\\\'") + '\\')">Entfernen</button></span></div>').join('')
  ).join('');
}
async function addRuleFromForm(){
  const value = document.getElementById('ruleValue').value.trim();
  if(!value){ alert('Bitte einen Wert eingeben.'); return; }
  await saveRule(document.getElementById('ruleList').value, value);
  document.getElementById('ruleValue').value = '';
}
async function saveRule(list, value){
  const res = await jpost('/api/mail/rules', {list: list, value: value});
  if(!res.ok){ alert(res.data.error); return; }
  renderRules(res.data.rules);
  noteRescan();
}
async function removeRule(value){
  const res = await jpost('/api/mail/rules', {value: value}, 'DELETE');
  if(!res.ok){ alert(res.data.error); return; }
  renderRules(res.data.rules);
  noteRescan();
}
function noteRescan(){
  const info = document.getElementById('uploadInfo');
  info.textContent = 'Regel geändert — bitte noch einmal durchsuchen, damit sie wirkt.';
}
function ruleForSender(sender, list, ev){
  ev.stopPropagation();
  if(list === 'always_delete'){
    if(!confirm(sender + ' wird ab jetzt bei jedem Scan automatisch zum Löschen vorgeschlagen, ' +
        'bis du die Regel unter "Eigene Regeln" wieder entfernst. Fortfahren?')) return;
    const card = document.getElementById('rulesCard');
    card.classList.remove('collapsed');
    card.querySelector('.card-head').setAttribute('aria-expanded', 'true');
  }
  saveRule(list, sender);
}

/* ---------- Gelernt ---------- */
async function loadLearned(){
  const j = await jget('/api/mail/learned');
  const area = document.getElementById('learnedArea');
  if(!j.rows || !j.rows.length){
    area.innerHTML = '<p class="muted">Noch nichts gelernt - das passiert nach dem ersten echten Aufräumen.</p>';
    return;
  }
  area.innerHTML = '<div class="hl">' + j.rows.map(r =>
    r.sender + '  —  gelöscht: ' + r.deleted + ', behalten: ' + r.kept + '  →  ' + r.tendency
  ).map(esc).join('<br>') + '</div>';
}

loadAccounts();
loadResult();
loadRules();
loadLearned();
refresh();
</script>
</body>
</html>
""".replace("__NAV__", _nav("/mail"))


MAIL_SEARCH_PAGE = _head("E-Mail suchen") + """
<div class="page">

  <div class="topbar">
    <div class="brand-mark">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="11" cy="11" r="7"/><path d="m21 21-4.35-4.35"/>
      </svg>
    </div>
    <div>
      <h1>E-Mail suchen</h1>
      <p class="sub">Eine bestimmte Mail wiederfinden - über das ganze Postfach</p>
    </div>
    __NAV__
  </div>

  <div class="callout callout-info">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>
    </svg>
    <div>Durchsucht <strong>alle Ordner</strong> eines Kontos, auch Papierkorb,
    Spam und Entwürfe - anders als beim Aufräumen, wo diese bewusst
    ausgeschlossen sind. Rein lesend, es wird nichts verändert.</div>
  </div>

  <div class="card">
    <div class="card-head">
      <div class="card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="11" cy="11" r="7"/><path d="m21 21-4.35-4.35"/>
        </svg>
      </div>
      <div>
        <h2>Suchkriterien</h2>
        <p class="muted desc">Mindestens eines ausfüllen. Alle Angaben sind
           einfache Teiltextsuchen, keine Groß-/Kleinschreibung.</p>
      </div>
    </div>
    <div class="form-grid">
      <label>Konto
        <select id="sAccount"><option value="">Alle Konten</option></select>
      </label>
      <label>Von enthält<input id="sSender" placeholder="z.B. amazon oder chef@firma.de"></label>
      <label>Betreff enthält<input id="sSubject" placeholder="z.B. Rechnung"></label>
      <label>Seit<input id="sSince" type="date"></label>
      <label>Bis<input id="sBefore" type="date"></label>
      <label>Auch im Mailtext suchen
        <input id="sBody" placeholder="z.B. Bestellnummer 4711">
      </label>
    </div>
    <p class="muted" style="font-size:.82rem">Textsuche liest die Mails selbst
       gegen (weiterhin nur lesend) und ist deshalb langsamer - je Ordner auf
       die ersten 500 Mails begrenzt.</p>
    <button id="btnSearch" class="btn btn-accent" onclick="startSearch()">Suchen</button>
  </div>

  <div id="status" aria-live="polite"></div>

  <div id="resultArea"></div>

</div>
<script>
function esc(s){
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}
function kb(bytes){
  if(bytes >= 1048576) return (bytes/1048576).toFixed(1) + ' MB';
  return Math.round(bytes/1024) + ' KB';
}
async function jget(url){
  const r = await fetch(url);
  return await r.json();
}
async function jpost(url, body, method){
  const r = await fetch(url, {method: method || 'POST',
    headers: {'Content-Type':'application/json'}, body: JSON.stringify(body || {})});
  return {ok: r.ok, data: await r.json()};
}
const I_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>';
const I_BAD = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>';
const I_WARN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4m0 4h.01M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/></svg>';

async function loadAccounts(){
  const j = await jget('/api/mail/accounts');
  const sel = document.getElementById('sAccount');
  j.accounts.forEach(a => {
    const o = document.createElement('option');
    o.value = a.name; o.textContent = a.name;
    sel.appendChild(o);
  });
}

let polling = null;
function poll(){
  if(polling) clearInterval(polling);
  polling = setInterval(refresh, 700);
  refresh();
}

async function startSearch(){
  const criteria = {
    account: document.getElementById('sAccount').value,
    sender: document.getElementById('sSender').value,
    subject: document.getElementById('sSubject').value,
    since: document.getElementById('sSince').value,
    before: document.getElementById('sBefore').value,
    body_text: document.getElementById('sBody').value,
  };
  const res = await jpost('/api/mail/search', criteria);
  if(!res.ok){ alert(res.data.error); return; }
  document.getElementById('resultArea').innerHTML = '';
  poll();
}

async function refresh(){
  const j = await jget('/api/status');
  const s = document.getElementById('status');
  const busy = j.running && j.kind === 'mailsearch';

  document.getElementById('btnSearch').disabled = j.running;
  if(busy){
    s.innerHTML = '<div class="card"><div class="status-head">' +
      '<span class="spinner"></span><span class="status-title">Durchsuche Postfach ...</span>' +
      '</div><p class="muted">' + esc(j.message) + '</p></div>';
    return;
  }
  if(polling){ clearInterval(polling); polling = null; }
  if(j.error){
    s.innerHTML = '<div class="card"><div class="status-head bad">' + I_BAD +
      '<span class="status-title">Fehler</span></div><p>' + esc(j.error) + '</p></div>';
    return;
  }
  if(!j.result || j.result.kind !== 'search'){ return; }
  renderResult(j.result);
}

function renderResult(r){
  const s = document.getElementById('status');
  const hits = r.hits || [];
  s.innerHTML = '<div class="card"><div class="status-head ' + (hits.length ? 'ok' : 'warn-text') + '">' +
    (hits.length ? I_OK : I_WARN) +
    '<span class="status-title">' + hits.length + ' Treffer in ' + r.folders_searched + ' Ordnern</span></div>' +
    (r.truncated ? '<p class="muted">Liste wurde begrenzt - bitte die Kriterien enger fassen, ' +
      'wenn die gesuchte Mail nicht dabei ist.</p>' : '') + '</div>';

  const area = document.getElementById('resultArea');
  if(!hits.length){ area.innerHTML = ''; return; }

  area.innerHTML = '<div class="card"><div class="diag-table-wrap"><table class="diag-table"><thead><tr>' +
    '<th>Datum</th><th>Konto</th><th>Ordner</th><th>Von</th><th>Betreff</th>' +
    '<th class="num">Größe</th><th>Kennzeichen</th></tr></thead><tbody>' +
    hits.map(h => {
      const flags = (h.flags || []).map(f => f.replace('\\\\', ''));
      const deleted = flags.some(f => f.toLowerCase() === 'deleted');
      return '<tr' + (deleted ? ' class="diag-hit"' : '') + '>' +
        '<td>' + esc((h.date || '').slice(0, 10)) + '</td>' +
        '<td>' + esc(h.account) + '</td>' +
        '<td>' + esc(h.folder) + (deleted ? ' <span class="muted">(als gelöscht markiert)</span>' : '') + '</td>' +
        '<td>' + esc(h.from_name || h.from) + '</td>' +
        '<td>' + esc(h.subject || '(kein Betreff)') + '</td>' +
        '<td class="num">' + kb(h.size) + '</td>' +
        '<td class="muted">' + esc(flags.join(', ')) + '</td></tr>';
    }).join('') +
    '</tbody></table></div></div>';
}

loadAccounts();
</script>
</body>
</html>
""".replace("__NAV__", _nav("/mail/suche"))


if __name__ == "__main__":
    # 0.0.0.0 statt 127.0.0.1: in Codespaces/Dev-Containern ist der Server auf
    # 127.0.0.1 nur innerhalb des Containers erreichbar - die Portweiterleitung
    # kommt dann nicht durch, obwohl der Prozess sauber läuft. 0.0.0.0 ist
    # GitHubs eigene Empfehlung für Codespaces und schadet lokal nicht.
    host = os.environ.get("HOST", "0.0.0.0")
    # 8000 statt Flasks üblichem 5000: in diesem Codespace blieb Port 5000
    # in der Weiterleitung hängen (vermutlich wegen "onAutoForward":
    # "openBrowser" im Devcontainer - der Browser öffnete sich, bevor die
    # Weiterleitung/Anmeldung stand). 8000 hat sich als zuverlässig erwiesen.
    port = int(os.environ.get("PORT", "8000"))
    print(f"iCloud Kontakte Sync - Web-Oberfläche")
    print(f"Öffne im Browser: http://127.0.0.1:{port}")
    print(f"E-Mail aufräumen:  http://127.0.0.1:{port}/mail")
    app.run(host=host, port=port, debug=False)
