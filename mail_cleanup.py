"""
E-Mail aufräumen - Bewertung, Excel-Rundlauf und Ausführung
============================================================
Durchsucht ausgewählte IMAP-Postfächer und -Ordner, bewertet jede Mail nach
festen Regeln und schlägt begründet vor, was in den Papierkorb kann. Der
Ablauf ist bewusst derselbe wie beim Kontakte-Sync:

    scan  ->  to-excel  ->  (in Excel oder im Browser prüfen)  ->  from-excel
          ->  clean (Testlauf)  ->  clean --execute

Grundsätze:
  * **Nichts wird endgültig gelöscht.** Mails wandern in den Papierkorb des
    Kontos und können dort zurückgeholt werden.
  * **Der Testlauf ist der Standard.** Ein echter Lauf braucht --execute.
  * **Nur Kopfzeilen werden gelesen**, keine Mail-Texte. Alles bleibt lokal,
    es wird nichts an einen Dienst gesendet.
  * **Das Tool lernt mit**: Nach jedem echten Lauf merkt es sich, welche
    Absender du gelöscht und welche du bewusst behalten hast, und bewertet
    beim nächsten Mal entsprechend.

Zugangsdaten stehen in mail_accounts.json (nicht im Git). Für iCloud und
Gmail wird ein app-spezifisches Passwort benötigt, nicht das normale.
"""

import argparse
import json
import sys
from datetime import datetime, timezone
from getpass import getpass
from pathlib import Path

import mail_imap as mi
from icloud_contacts import _cell_is_delete

ACCOUNTS_FILE = Path("mail_accounts.json")
DECISIONS_FILE = Path("mail_decisions.json")
SCAN_FILE = Path("web_data") / "mail_scan.json"

# Bekannte Anbieter - erspart das Nachschlagen der Serveradressen.
PRESETS = {
    "icloud": {"host": "imap.mail.me.com", "port": 993},
    "gmail": {"host": "imap.gmail.com", "port": 993},
    "gmx": {"host": "imap.gmx.net", "port": 993},
    "web.de": {"host": "imap.web.de", "port": 993},
    "outlook": {"host": "outlook.office365.com", "port": 993},
}

DEFAULT_SETTINGS = {
    "min_age_days": 30,     # jünger wird nie vorgeschlagen
    "delete_at": 60,        # ab dieser Punktzahl vorangehakt
    "unsure_at": 35,        # darunter gilt "behalten"
}

# Obergrenze pro Lauf. Verhindert, dass ein Versehen (falscher Ordner, zu
# scharfe Einstellung) auf einen Schlag das halbe Postfach räumt.
MAX_PER_RUN = 2000

# Absender, die praktisch nie eine Antwort erwarten.
AUTOMATED_PREFIXES = (
    "noreply", "no-reply", "no_reply", "donotreply", "do-not-reply",
    "notifications", "notification", "newsletter", "mailer-daemon",
    "bounce", "bounces", "automail", "auto-confirm", "postmaster",
)


# ---------------------------------------------------------------------------
# Konten
# ---------------------------------------------------------------------------

def load_accounts():
    if not ACCOUNTS_FILE.exists():
        return []
    try:
        data = json.loads(ACCOUNTS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise RuntimeError(f"{ACCOUNTS_FILE} ist beschädigt: {e}") from e
    return data.get("accounts", [])


def save_accounts(accounts):
    ACCOUNTS_FILE.write_text(
        json.dumps({"accounts": accounts}, indent=2, ensure_ascii=False),
        encoding="utf-8")


def account_by_name(name):
    for acc in load_accounts():
        if acc["name"] == name:
            return acc
    raise RuntimeError(f"Kein Konto namens '{name}'. Vorhanden: "
                       f"{', '.join(a['name'] for a in load_accounts()) or '(keins)'}")


def add_account(name, host, user, password, port=993):
    accounts = [a for a in load_accounts() if a["name"] != name]
    accounts.append({"name": name, "host": host, "port": int(port),
                     "user": user, "password": password})
    save_accounts(accounts)
    return accounts


def remove_account(name):
    accounts = [a for a in load_accounts() if a["name"] != name]
    save_accounts(accounts)
    return accounts


def connect_account(account):
    return mi.connect(account["host"], account["user"], account["password"],
                      account.get("port", 993))


def test_account(account):
    """Verbindung prüfen und Ordner zählen - für den 'Testen'-Knopf."""
    conn = connect_account(account)
    try:
        folders = mi.list_folders(conn)
        return {"ok": True, "folders": len(folders),
                "trash": mi.find_trash(folders)}
    finally:
        mi.close(conn)


def list_account_folders(account):
    """Ordner eines Kontos mit Kennzeichen, ob sie zum Scannen taugen."""
    conn = connect_account(account)
    try:
        out = []
        for name, flags in mi.list_folders(conn):
            if "\\noselect" in [f.lower() for f in flags]:
                continue
            out.append({"name": name, "skip": mi.is_skipped_folder(name, flags)})
        return out
    finally:
        mi.close(conn)


# ---------------------------------------------------------------------------
# Lernspeicher
# ---------------------------------------------------------------------------

def load_decisions():
    if not DECISIONS_FILE.exists():
        return {"senders": {}, "domains": {}}
    try:
        data = json.loads(DECISIONS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"senders": {}, "domains": {}}
    data.setdefault("senders", {})
    data.setdefault("domains", {})
    return data


def save_decisions(decisions):
    DECISIONS_FILE.write_text(json.dumps(decisions, indent=2, ensure_ascii=False),
                              encoding="utf-8")


def record_decisions(mails, decisions=None):
    """Merkt sich die bestätigten Entscheidungen - in BEIDE Richtungen.

    Abgehakte Mails zählen als "gelöscht", bewusst nicht abgehakte Kandidaten
    als "behalten". Ohne die Behalten-Seite würde das Tool nur lernen, wen es
    wegwerfen soll, aber nie, wen es künftig in Ruhe lassen muss.
    """
    decisions = decisions if decisions is not None else load_decisions()
    today = datetime.now(timezone.utc).date().isoformat()

    for mail in mails:
        sender = (mail.get("from") or "").strip().lower()
        if not sender:
            continue
        domain = mail.get("domain") or sender.rpartition("@")[2]
        field = "deleted" if mail.get("delete") else "kept"
        for store, key in ((decisions["senders"], sender), (decisions["domains"], domain)):
            if not key:
                continue
            entry = store.setdefault(key, {"deleted": 0, "kept": 0, "last": ""})
            entry[field] = entry.get(field, 0) + 1
            entry["last"] = today
    return decisions


def learned_summary(decisions=None, limit=40):
    """Was das Tool bisher gelernt hat - für Anzeige und Nachvollziehbarkeit."""
    decisions = decisions if decisions is not None else load_decisions()
    rows = []
    for sender, e in decisions.get("senders", {}).items():
        deleted, kept = e.get("deleted", 0), e.get("kept", 0)
        rows.append({"sender": sender, "deleted": deleted, "kept": kept,
                     "last": e.get("last", ""),
                     "tendency": "löschen" if deleted > kept else
                                 ("behalten" if kept > deleted else "gemischt")})
    rows.sort(key=lambda r: -(r["deleted"] + r["kept"]))
    return rows[:limit]


# ---------------------------------------------------------------------------
# Bewertung
# ---------------------------------------------------------------------------

def _age_days(mail, now=None):
    raw = mail.get("date")
    if not raw:
        return None
    try:
        dt = datetime.fromisoformat(raw)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    now = now or datetime.now(timezone.utc)
    return max(0, (now - dt).days)


def _is_bulk(mail):
    if mail.get("list_unsubscribe") or mail.get("list_id"):
        return True
    if mail.get("precedence") in ("bulk", "list", "junk"):
        return True
    auto = mail.get("auto_submitted", "")
    return bool(auto) and auto != "no"


def _is_automated_sender(mail):
    local = (mail.get("from") or "").partition("@")[0].lower()
    return any(local.startswith(p) for p in AUTOMATED_PREFIXES)


def _learned_hint(mail, decisions):
    """Was frühere Entscheidungen über diesen Absender sagen.

    Gibt (punkte, begründung, urteil) zurück. urteil ist None, "keep" oder
    "delete" - bei eindeutiger Vorgeschichte entscheidet das Gelernte direkt,
    statt nur Punkte beizusteuern. Wer dreimal dasselbe weggeworfen und nie
    etwas davon behalten hat, hat seine Meinung deutlich genug gesagt; ein
    Punktesystem würde dieses klare Signal nur verwässern.
    """
    sender = (mail.get("from") or "").strip().lower()
    entry = decisions.get("senders", {}).get(sender)
    if entry:
        deleted, kept = entry.get("deleted", 0), entry.get("kept", 0)
        if kept >= 2 and deleted == 0:
            return 0, f"von dir bisher immer behalten ({kept}×)", "keep"
        if deleted >= 3 and kept == 0:
            return 0, f"von dir bisher immer gelöscht ({deleted}×)", "delete"
        if deleted >= 1 and kept == 0:
            return 15, f"von dir schon gelöscht ({deleted}×)", None
        return 0, "", None

    # Nur wenn die genaue Adresse unbekannt ist, zählt die Domain - und
    # schwächer, weil hinter einer Domain auch persönliche Post stecken kann.
    domain = mail.get("domain") or sender.rpartition("@")[2]
    entry = decisions.get("domains", {}).get(domain)
    if entry:
        deleted, kept = entry.get("deleted", 0), entry.get("kept", 0)
        if deleted >= 5 and kept == 0:
            return 15, f"Absender {domain} bisher immer gelöscht ({deleted}×)", None
    return 0, "", None


def score_mail(mail, decisions=None, settings=None, now=None):
    """Bewertet eine Mail: (punkte, empfehlung, begründungen).

    empfehlung ist "delete", "unsure" oder "keep". Nur "delete" wird in Excel
    und Web-Oberfläche vorangehakt - alles andere muss der Nutzer selbst
    ankreuzen.
    """
    decisions = decisions if decisions is not None else load_decisions()
    cfg = {**DEFAULT_SETTINGS, **(settings or {})}
    flags = [f.lower() for f in mail.get("flags", [])]
    reasons = []

    # --- Harte Schutzregeln: nie vorschlagen, egal wie die Punkte stehen ---
    if "\\flagged" in flags:
        return 0, "keep", ["markiert - wird nie vorgeschlagen"]
    if "\\answered" in flags:
        return 0, "keep", ["du hast geantwortet - wird nie vorgeschlagen"]
    if "\\draft" in flags:
        return 0, "keep", ["Entwurf - wird nie vorgeschlagen"]

    age = _age_days(mail, now)
    if age is None:
        return 0, "keep", ["kein lesbares Datum - sicherheitshalber behalten"]
    if age < cfg["min_age_days"]:
        return 0, "keep", [f"erst {age} Tage alt - wird nie vorgeschlagen"]

    learned_points, learned_reason, learned_verdict = _learned_hint(mail, decisions)
    if learned_verdict == "keep":
        return 0, "keep", [learned_reason + " - wird nie vorgeschlagen"]
    if learned_verdict == "delete":
        return cfg["delete_at"], "delete", [learned_reason]

    # --- Punktevergabe ---
    score = 0

    for threshold, points, label in ((1825, 30, "über 5 Jahre alt"),
                                     (1095, 25, "über 3 Jahre alt"),
                                     (730, 20, "über 2 Jahre alt"),
                                     (365, 15, "über 1 Jahr alt"),
                                     (180, 10, "über ein halbes Jahr alt"),
                                     (90, 5, "über 3 Monate alt")):
        if age >= threshold:
            score += points
            reasons.append(label)
            break

    if _is_bulk(mail):
        score += 35
        reasons.append("Newsletter/Massenmail")
    if _is_automated_sender(mail):
        score += 20
        reasons.append("automatischer Absender")

    if "\\seen" in flags:
        score += 10
        reasons.append("gelesen")
    else:
        score += 5
        reasons.append("nie gelesen")

    size = mail.get("size", 0)
    if size >= 5_000_000:
        score += 10
        reasons.append(f"sehr groß ({size // 1_000_000} MB)")
    elif size >= 1_000_000:
        score += 5
        reasons.append(f"groß ({size // 1_000_000} MB)")

    if learned_points:
        score += learned_points
        reasons.append(learned_reason)

    score = max(0, min(100, score))
    if score >= cfg["delete_at"]:
        recommendation = "delete"
    elif score >= cfg["unsure_at"]:
        recommendation = "unsure"
    else:
        recommendation = "keep"
    return score, recommendation, reasons


# ---------------------------------------------------------------------------
# Scannen
# ---------------------------------------------------------------------------

RECOMMENDATION_DE = {"delete": "löschen", "unsure": "unklar", "keep": "behalten"}


def scan_accounts(selection, settings=None, progress=None):
    """Durchsucht die gewählten Konten/Ordner - streng nur lesend.

    selection: [{"account": "iCloud", "folders": ["INBOX", "Archiv"]}, ...]
    Gibt die Scan-Struktur zurück (und speichert sie über save_scan()).
    """
    cfg = {**DEFAULT_SETTINGS, **(settings or {})}
    decisions = load_decisions()
    scan = {"created": datetime.now(timezone.utc).isoformat(),
            "settings": cfg, "accounts": []}

    for entry in selection:
        account = account_by_name(entry["account"])
        conn = connect_account(account)
        try:
            folders = mi.list_folders(conn)
            trash = mi.find_trash(folders)
            acc_out = {"account": account["name"], "trash": trash, "folders": []}

            for folder in entry.get("folders", []):
                if progress:
                    progress(f"{account['name']}: lese '{folder}' ...")

                def on_block(done, total, _f=folder, _a=account["name"]):
                    if progress:
                        progress(f"{_a}: '{_f}' {done}/{total} Mails gelesen ...")

                mails, uidvalidity = mi.scan_folder(conn, folder, progress=on_block)
                for mail in mails:
                    score, recommendation, reasons = score_mail(mail, decisions, cfg)
                    mail["score"] = score
                    mail["recommendation"] = recommendation
                    mail["reasons"] = reasons
                    mail["delete"] = recommendation == "delete"
                acc_out["folders"].append(
                    {"folder": folder, "uidvalidity": uidvalidity, "mails": mails})
            scan["accounts"].append(acc_out)
        finally:
            mi.close(conn)

    save_scan(scan)
    return scan


def save_scan(scan):
    SCAN_FILE.parent.mkdir(exist_ok=True)
    SCAN_FILE.write_text(json.dumps(scan, ensure_ascii=False), encoding="utf-8")


def load_scan():
    if not SCAN_FILE.exists():
        raise RuntimeError("Es liegt kein Scan vor. Bitte zuerst 'scan' ausführen.")
    return json.loads(SCAN_FILE.read_text(encoding="utf-8"))


def iter_mails(scan):
    """Alle Mails eines Scans als (konto, ordner_eintrag, mail)."""
    for acc in scan.get("accounts", []):
        for folder in acc.get("folders", []):
            for mail in folder.get("mails", []):
                yield acc, folder, mail


def mail_key(account_name, folder_name, uid):
    return f"{account_name}|{folder_name}|{uid}"


def scan_summary(scan):
    counts = {"total": 0, "delete": 0, "unsure": 0, "keep": 0, "bytes": 0}
    for _acc, _folder, mail in iter_mails(scan):
        counts["total"] += 1
        counts[mail.get("recommendation", "keep")] += 1
        if mail.get("delete"):
            counts["bytes"] += mail.get("size", 0)
    counts["selected"] = sum(1 for _a, _f, m in iter_mails(scan) if m.get("delete"))
    return counts


# ---------------------------------------------------------------------------
# Excel-Rundlauf
# ---------------------------------------------------------------------------

EXCEL_COLUMNS = [
    "Konto", "Ordner", "UID", "Löschen", "Empfehlung", "Punkte", "Begründung",
    "Datum", "Alter (Tage)", "Von", "Absender", "Domain", "Betreff",
    "Größe (KB)", "Gelesen", "Markiert", "Beantwortet",
]
_EXCEL_WIDTHS = [16, 22, 10, 9, 12, 8, 46, 12, 12, 26, 34, 22, 52, 12, 9, 9, 12]


def scan_to_excel(scan, path):
    """Scan-Ergebnis als Excel-Liste - Aufbau wie bei den Kontakten, damit der
    Ablauf vertraut bleibt (Kopfzeile fixiert, Filter an, Löschen-Spalte vorn)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Aufräumen"
    ws.append(EXCEL_COLUMNS)
    fill = PatternFill("solid", fgColor="305496")
    font = Font(bold=True, color="FFFFFF")
    for col in range(1, len(EXCEL_COLUMNS) + 1):
        cell = ws.cell(1, col)
        cell.fill, cell.font = fill, font
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    rows = 0
    for acc, folder, mail in iter_mails(scan):
        flags = [f.lower() for f in mail.get("flags", [])]
        age = _age_days(mail)
        ws.append([
            acc["account"], folder["folder"], mail["uid"],
            "x" if mail.get("delete") else "",
            RECOMMENDATION_DE.get(mail.get("recommendation"), ""),
            mail.get("score", 0), ", ".join(mail.get("reasons", [])),
            (mail.get("date") or "")[:10], age if age is not None else "",
            mail.get("from_name", ""), mail.get("from", ""), mail.get("domain", ""),
            mail.get("subject", ""), round(mail.get("size", 0) / 1024),
            "ja" if "\\seen" in flags else "",
            "ja" if "\\flagged" in flags else "",
            "ja" if "\\answered" in flags else "",
        ])
        rows += 1

    ws.freeze_panes = "E2"   # Konto/Ordner/UID/Löschen bleiben sichtbar
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EXCEL_COLUMNS))}{rows + 1}"
    for col, width in enumerate(_EXCEL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.column_dimensions["C"].hidden = True   # UID nicht versehentlich ändern
    wb.save(path)
    return rows


def scan_from_excel(scan, path):
    """Liest die bearbeitete Excel-Liste zurück und überträgt die Häkchen der
    Spalte 'Löschen' in den Scan. Nutzt dieselbe Auswertung wie bei den
    Kontakten (_cell_is_delete), damit 'x', 'ja', '1' ... gleich wirken."""
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "").strip() for c in next(rows, [])]
    for needed in ("Konto", "Ordner", "UID", "Löschen"):
        if needed not in header:
            raise RuntimeError(
                f"Spalte '{needed}' fehlt. Stammt die Datei aus "
                "'Als Excel herunterladen' bzw. 'to-excel'?")
    idx = {name: header.index(name) for name in ("Konto", "Ordner", "UID", "Löschen")}

    wanted = {}
    for row in rows:
        if row is None or all(c is None for c in row):
            continue
        try:
            uid = int(row[idx["UID"]])
        except (TypeError, ValueError):
            continue
        key = mail_key(str(row[idx["Konto"]]), str(row[idx["Ordner"]]), uid)
        wanted[key] = _cell_is_delete(row[idx["Löschen"]])
    wb.close()

    changed = 0
    for acc, folder, mail in iter_mails(scan):
        key = mail_key(acc["account"], folder["folder"], mail["uid"])
        if key in wanted and bool(mail.get("delete")) != wanted[key]:
            mail["delete"] = wanted[key]
            changed += 1
    save_scan(scan)
    return {"gelesen": len(wanted), "geändert": changed,
            "ausgewählt": sum(1 for _a, _f, m in iter_mails(scan) if m.get("delete"))}


def apply_selection(scan, keys):
    """Auswahl aus der Web-Oberfläche übernehmen (Liste von mail_key-Strings)."""
    wanted = set(keys)
    for acc, folder, mail in iter_mails(scan):
        mail["delete"] = mail_key(acc["account"], folder["folder"], mail["uid"]) in wanted
    save_scan(scan)
    return sum(1 for _a, _f, m in iter_mails(scan) if m.get("delete"))


# ---------------------------------------------------------------------------
# Ausführen
# ---------------------------------------------------------------------------

def clean(scan, dry_run=True, force=False, progress=None):
    """Verschiebt die angehakten Mails in den Papierkorb des jeweiligen Kontos.

    Testlauf ist der Standard. Beim echten Lauf werden anschließend die
    Entscheidungen gelernt - und zwar nur über Mails, die überhaupt zur
    Auswahl standen (Empfehlung 'löschen' oder 'unklar'). Über Mails, die nie
    vorgeschlagen wurden, hat der Nutzer auch nichts entschieden; sie als
    'behalten' zu zählen würde den Lernspeicher mit Zufallsdaten fluten.
    """
    result = {"moved": 0, "skipped": 0, "errors": 0, "total": 0, "log": [],
              "dry_run": dry_run}

    selected = [(acc, folder, mail) for acc, folder, mail in iter_mails(scan)
                if mail.get("delete")]
    result["total"] = len(selected)
    if not selected:
        result["log"].append("Nichts ausgewählt - es gibt nichts zu tun.")
        return result

    if len(selected) > MAX_PER_RUN and not force:
        raise RuntimeError(
            f"{len(selected)} Mails ausgewählt - das ist mehr als die "
            f"Sicherheitsgrenze von {MAX_PER_RUN} pro Lauf. Bitte die Auswahl "
            "prüfen (stimmen Ordner und Einstellungen?) oder den Lauf mit "
            "--force bzw. in kleineren Portionen wiederholen.")

    for acc in scan.get("accounts", []):
        by_folder = {}
        for folder in acc.get("folders", []):
            items = [{"uid": m["uid"], "message_id": m.get("message_id", "")}
                     for m in folder.get("mails", []) if m.get("delete")]
            if items:
                by_folder[folder["folder"]] = (items, folder.get("uidvalidity"))
        if not by_folder:
            continue

        trash = acc.get("trash")
        if not trash:
            count = sum(len(v[0]) for v in by_folder.values())
            result["errors"] += count
            result["log"].append(
                f"FEHLER: Für Konto '{acc['account']}' wurde kein Papierkorb "
                "gefunden - nichts verschoben.")
            continue

        try:
            account = account_by_name(acc["account"])
            conn = connect_account(account)
        except Exception as e:  # noqa: BLE001
            count = sum(len(v[0]) for v in by_folder.values())
            result["errors"] += count
            result["log"].append(f"FEHLER: Konto '{acc['account']}' nicht erreichbar: {e}")
            continue

        try:
            for folder_name, (items, uidvalidity) in by_folder.items():
                if progress:
                    progress(f"{acc['account']}: '{folder_name}' - {len(items)} Mails ...")
                part = mi.move_to_trash(conn, folder_name, items, trash,
                                        expected_uidvalidity=uidvalidity,
                                        dry_run=dry_run)
                for key in ("moved", "skipped", "errors"):
                    result[key] += part[key]
                result["log"].extend(
                    f"{acc['account']} / {folder_name}: {line}" for line in part["log"])
        finally:
            mi.close(conn)

    if not dry_run:
        reviewed = [m for _a, _f, m in iter_mails(scan)
                    if m.get("recommendation") in ("delete", "unsure")]
        save_decisions(record_decisions(reviewed))
        result["log"].append(
            f"Gelernt aus {len(reviewed)} geprüften Vorschlägen "
            "(gelöscht wie behalten).")
        scan["executed"] = datetime.now(timezone.utc).isoformat()
        save_scan(scan)
    return result


# ---------------------------------------------------------------------------
# Kommandozeile
# ---------------------------------------------------------------------------

def _print_accounts():
    accounts = load_accounts()
    if not accounts:
        print("Noch keine Konten eingerichtet. Anlegen mit:")
        print("  python mail_cleanup.py konten --add")
        return
    print(f"{len(accounts)} Konto/Konten:")
    for acc in accounts:
        print(f"  - {acc['name']}: {acc['user']} auf {acc['host']}:{acc.get('port', 993)}")


def cmd_konten(args):
    if args.add:
        print("Neues Konto anlegen (Abbruch mit Strg+C)\n")
        name = input("Anzeigename (z.B. iCloud): ").strip()
        print(f"Anbieter-Kürzel für die Voreinstellung: {', '.join(PRESETS)}")
        provider = input("Anbieter (oder leer für eigenen Server): ").strip().lower()
        preset = PRESETS.get(provider, {})
        host = input(f"IMAP-Server [{preset.get('host', '')}]: ").strip() or preset.get("host", "")
        port = input(f"Port [{preset.get('port', 993)}]: ").strip() or preset.get("port", 993)
        user = input("Benutzername (meist die E-Mail-Adresse): ").strip()
        print("Bei iCloud und Gmail wird ein APP-SPEZIFISCHES Passwort gebraucht,")
        print("nicht das normale Kennwort.")
        password = getpass("Passwort: ")
        if not (name and host and user and password):
            print("Abgebrochen - es fehlten Angaben.")
            sys.exit(1)
        add_account(name, host, user, password, port)
        print(f"\nKonto '{name}' gespeichert in {ACCOUNTS_FILE}.")
        try:
            info = test_account(account_by_name(name))
            print(f"Verbindung ok - {info['folders']} Ordner, "
                  f"Papierkorb: {info['trash'] or 'nicht gefunden!'}")
        except Exception as e:  # noqa: BLE001
            print(f"ACHTUNG: Verbindung fehlgeschlagen: {e}")
        return

    if args.remove:
        remove_account(args.remove)
        print(f"Konto '{args.remove}' entfernt.")
        return

    if args.test:
        for acc in load_accounts():
            try:
                info = test_account(acc)
                print(f"  {acc['name']}: ok, {info['folders']} Ordner, "
                      f"Papierkorb: {info['trash'] or 'NICHT GEFUNDEN'}")
            except Exception as e:  # noqa: BLE001
                print(f"  {acc['name']}: FEHLER - {e}")
        return

    if args.folders:
        acc = account_by_name(args.folders)
        for entry in list_account_folders(acc):
            mark = "  (wird übersprungen)" if entry["skip"] else ""
            print(f"  {entry['name']}{mark}")
        return

    _print_accounts()


def cmd_scan(args):
    accounts = load_accounts()
    if not accounts:
        print("Noch keine Konten eingerichtet: python mail_cleanup.py konten --add")
        sys.exit(1)

    names = args.account or [a["name"] for a in accounts]
    selection = []
    for name in names:
        acc = account_by_name(name)
        if args.folder:
            folders = list(args.folder)
        else:
            folders = [f["name"] for f in list_account_folders(acc) if not f["skip"]]
            print(f"{name}: {len(folders)} Ordner werden gelesen "
                  f"(Papierkorb, Entwürfe und Spam bleiben außen vor).")
        selection.append({"account": name, "folders": folders})

    settings = {"min_age_days": args.min_age} if args.min_age is not None else None
    scan = scan_accounts(selection, settings, progress=lambda m: print(f"  {m}", flush=True))

    counts = scan_summary(scan)
    print(f"\n{counts['total']} Mails geprüft: {counts['delete']} zum Löschen "
          f"vorgeschlagen, {counts['unsure']} unklar, {counts['keep']} behalten.")
    print(f"Vorgeschlagen wären {counts['bytes'] / 1_048_576:.1f} MB.")
    print(f"\nGespeichert: {SCAN_FILE}")
    print("Weiter mit:  python mail_cleanup.py to-excel")


def cmd_to_excel(args):
    rows = scan_to_excel(load_scan(), args.output)
    print(f"{rows} Mails geschrieben.")
    print(f"Gespeichert: {args.output}")
    print("Jetzt in Excel die Spalte 'Löschen' prüfen, dann:")
    print(f"  python mail_cleanup.py from-excel --input {args.output}")


def cmd_from_excel(args):
    info = scan_from_excel(load_scan(), args.input)
    print(f"{info['gelesen']} Zeilen gelesen, {info['geändert']} Häkchen geändert.")
    print(f"{info['ausgewählt']} Mails sind jetzt zum Löschen ausgewählt.")
    print("Weiter mit:  python mail_cleanup.py clean        (Testlauf)")


def cmd_clean(args):
    scan = load_scan()
    dry_run = not args.execute
    if dry_run:
        print("TESTLAUF: Es wird nichts verschoben. Echter Lauf mit --execute.\n")
    result = clean(scan, dry_run=dry_run, force=args.force,
                   progress=lambda m: print(f"  {m}", flush=True))
    for line in result["log"]:
        print(f"  {line}")
    verb = "würden verschoben" if dry_run else "verschoben"
    print(f"\n{result['moved']} Mails {verb}, {result['skipped']} übersprungen, "
          f"{result['errors']} Fehler (von {result['total']} ausgewählten).")
    if dry_run and result["moved"]:
        print("\nSieht das gut aus? Dann:")
        print("  python mail_cleanup.py clean --execute")


def cmd_gelernt(args):
    if args.reset:
        if DECISIONS_FILE.exists():
            DECISIONS_FILE.unlink()
        print("Lernspeicher zurückgesetzt.")
        return
    rows = learned_summary()
    if not rows:
        print("Noch nichts gelernt - das passiert nach dem ersten echten Aufräumen.")
        return
    print(f"{'Absender':44} {'gelöscht':>9} {'behalten':>9}  Tendenz")
    print("-" * 78)
    for row in rows:
        print(f"{row['sender'][:44]:44} {row['deleted']:>9} {row['kept']:>9}  {row['tendency']}")


def main():
    parser = argparse.ArgumentParser(
        description="E-Mail-Postfächer aufräumen: bewerten, prüfen, in den Papierkorb verschieben")
    sub = parser.add_subparsers(dest="cmd", required=True)

    konten = sub.add_parser("konten", help="IMAP-Konten anzeigen, anlegen, prüfen")
    konten.add_argument("--add", action="store_true", help="Neues Konto anlegen (fragt nach)")
    konten.add_argument("--remove", metavar="NAME", help="Konto entfernen")
    konten.add_argument("--test", action="store_true", help="Alle Konten auf Erreichbarkeit prüfen")
    konten.add_argument("--folders", metavar="NAME", help="Ordner eines Kontos auflisten")

    scan = sub.add_parser("scan", help="Postfächer durchsuchen und bewerten (nur lesend)")
    scan.add_argument("--account", action="append",
                      help="Konto (mehrfach angebbar; Standard: alle)")
    scan.add_argument("--folder", action="append",
                      help="Ordner (mehrfach angebbar; Standard: alle sinnvollen)")
    scan.add_argument("--min-age", type=int, metavar="TAGE",
                      help=f"Mails jünger als N Tage nie vorschlagen "
                           f"(Standard: {DEFAULT_SETTINGS['min_age_days']})")

    toexcel = sub.add_parser("to-excel", help="Scan-Ergebnis als Excel-Liste ausgeben")
    toexcel.add_argument("--output", default="mail_aufraeumen.xlsx")

    fromexcel = sub.add_parser("from-excel", help="Bearbeitete Excel-Liste einlesen")
    fromexcel.add_argument("--input", required=True)

    clean_p = sub.add_parser("clean", help="Ausgewählte Mails in den Papierkorb verschieben")
    clean_p.add_argument("--execute", action="store_true",
                         help="Wirklich verschieben (ohne diese Angabe nur Testlauf)")
    clean_p.add_argument("--force", action="store_true",
                         help=f"Sicherheitsgrenze von {MAX_PER_RUN} Mails pro Lauf aufheben")

    gelernt = sub.add_parser("gelernt", help="Zeigen, was aus deinen Entscheidungen gelernt wurde")
    gelernt.add_argument("--reset", action="store_true", help="Lernspeicher leeren")

    args = parser.parse_args()
    handlers = {"konten": cmd_konten, "scan": cmd_scan, "to-excel": cmd_to_excel,
                "from-excel": cmd_from_excel, "clean": cmd_clean, "gelernt": cmd_gelernt}
    try:
        handlers[args.cmd](args)
    except RuntimeError as e:
        print(f"\nFehler: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nAbgebrochen.")
        sys.exit(1)


if __name__ == "__main__":
    main()
