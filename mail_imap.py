"""
IMAP-Grundlagen für die E-Mail-Aufräumaktion
=============================================
Reine Transportschicht: Verbindung, Ordner, Header lesen, Mails in den
Papierkorb verschieben. Enthält bewusst KEINE Bewertungslogik - die steckt in
mail_cleanup.py, damit sie ohne Postfach testbar bleibt.

Zwei Dinge, die man bei IMAP leicht falsch macht und die hier deshalb
ausdrücklich behandelt werden:

1. **Ordnernamen** sind in "modifiziertem UTF-7" kodiert (RFC 3501 5.1.3).
   Ohne Umwandlung bricht jeder Ordner mit Umlauten - also z.B. "Entwürfe"
   oder "Gelöschte Objekte". Python bringt dafür keinen Codec mit.

2. **UIDs statt Sequenznummern.** Sequenznummern verschieben sich, sobald eine
   Mail verschwindet - wer damit löscht, erwischt irgendwann die falsche Mail.
   Hier wird durchgängig mit UID-Befehlen gearbeitet, und vor dem Verschieben
   werden UIDVALIDITY und Message-ID gegen den Scan-Stand geprüft.

Gelesen werden beim Durchsuchen selbst ausschließlich Kopfzeilen
(BODY.PEEK[HEADER.FIELDS ...]) - das reicht für die Bewertung, ist deutlich
schneller und lässt den Gelesen-Status unangetastet (PEEK). fetch_body_snippets()
ist die einzige Ausnahme: ein kurzer Textausschnitt für die Handvoll Mails, bei
denen mail_cleanup.py aus den Kopfzeilen keinen Mail-Typ (Rechnung, Bank/
Behörde) erkennen konnte, ebenfalls per PEEK und ohne Nebenwirkungen.
"""

import base64
import email.utils
import imaplib
import re
import time as _time
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.parser import BytesParser

# imaplib begrenzt Antwortzeilen auf 10.000 Zeichen. Ordnerlisten großer
# Postfächer sprengen das - großzügig anheben.
imaplib._MAXLINE = 10_000_000

FETCH_BATCH = 200          # UIDs pro FETCH-Befehl
DEFAULT_TIMEOUT = 60       # Sekunden pro IMAP-Operation

# Kopfzeilen, die für die Bewertung gebraucht werden. Bewusst knapp gehalten.
HEADER_FIELDS = ("FROM SUBJECT DATE MESSAGE-ID LIST-UNSUBSCRIBE LIST-ID "
                 "PRECEDENCE AUTO-SUBMITTED")

# Ordner, die nie gescannt werden - dort aufzuräumen ergibt keinen Sinn bzw.
# wäre gefährlich (Papierkorb ist ja schon das Ziel).
SKIP_FOLDER_FLAGS = {"\\trash", "\\drafts", "\\junk", "\\all"}
SKIP_FOLDER_NAMES = {
    "trash", "papierkorb", "deleted messages", "gelöschte objekte",
    "gelöschte elemente", "[gmail]/papierkorb", "[gmail]/trash",
    "drafts", "entwürfe", "[gmail]/entwürfe", "[gmail]/drafts",
    "junk", "spam", "junk e-mail", "[gmail]/spam",
    "[gmail]/alle nachrichten", "[gmail]/all mail",
}

# Papierkorb-Erkennung, wenn der Server kein \Trash-Sonderflag liefert.
TRASH_NAMES = [
    "trash", "papierkorb", "deleted messages", "gelöschte objekte",
    "gelöschte elemente", "[gmail]/papierkorb", "[gmail]/trash", "deleted items",
]


# ---------------------------------------------------------------------------
# Modifiziertes UTF-7 (RFC 3501 5.1.3) für Ordnernamen
# ---------------------------------------------------------------------------

def _b64_chunk(text):
    """Nicht-ASCII-Abschnitt in modifiziertes BASE64 (',' statt '/', ohne '=')."""
    raw = base64.b64encode(text.encode("utf-16-be")).decode("ascii")
    return raw.rstrip("=").replace("/", ",")


def utf7_encode(name):
    """Klartext-Ordnername -> IMAP-Darstellung ("Entwürfe" -> "Entw&APw-rfe")."""
    out, buf = [], []
    for ch in name:
        if 0x20 <= ord(ch) <= 0x7E:
            if buf:
                out.append("&" + _b64_chunk("".join(buf)) + "-")
                buf = []
            out.append("&-" if ch == "&" else ch)
        else:
            buf.append(ch)
    if buf:
        out.append("&" + _b64_chunk("".join(buf)) + "-")
    return "".join(out)


def utf7_decode(name):
    """IMAP-Darstellung -> Klartext-Ordnername."""
    if isinstance(name, bytes):
        name = name.decode("ascii", errors="replace")
    out, i, n = [], 0, len(name)
    while i < n:
        if name[i] != "&":
            out.append(name[i])
            i += 1
            continue
        end = name.find("-", i)
        if end == -1:
            end = n
        chunk = name[i + 1:end]
        if chunk == "":
            out.append("&")  # "&-" steht für ein einzelnes &
        else:
            b64 = chunk.replace(",", "/")
            b64 += "=" * (-len(b64) % 4)
            try:
                out.append(base64.b64decode(b64).decode("utf-16-be"))
            except Exception:  # noqa: BLE001 - kaputte Kodierung unverändert lassen
                out.append(name[i:end + 1])
        i = end + 1
    return "".join(out)


def _quote(name):
    """Ordnername für einen IMAP-Befehl: kodieren und in Anführungszeichen."""
    encoded = utf7_encode(name).replace("\\", "\\\\").replace('"', '\\"')
    return f'"{encoded}"'


# ---------------------------------------------------------------------------
# Verbindung
# ---------------------------------------------------------------------------

def connect(host, user, password, port=993, timeout=DEFAULT_TIMEOUT):
    """Baut eine IMAP-SSL-Verbindung auf und meldet an.

    Wirft RuntimeError mit einer verständlichen Meldung - imaplib-Fehler sind
    für sich genommen wenig hilfreich ("command SELECT illegal in state AUTH").
    """
    try:
        conn = imaplib.IMAP4_SSL(host, int(port), timeout=timeout)
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"Verbindung zu {host}:{port} fehlgeschlagen: {e}") from e
    try:
        conn.login(user, password)
    except Exception as e:  # noqa: BLE001
        try:
            conn.logout()
        except Exception:  # noqa: BLE001
            pass
        raise RuntimeError(
            f"Anmeldung als {user} fehlgeschlagen: {e}. Bei iCloud und Gmail muss "
            "ein App-spezifisches Passwort verwendet werden, nicht das normale."
        ) from e
    return conn


def close(conn):
    """Verbindung beenden, ohne dass ein Fehler beim Abmelden alles umwirft."""
    try:
        conn.logout()
    except Exception:  # noqa: BLE001
        pass


def has_capability(conn, name):
    """imaplib liefert capabilities als str-Tupel; ein Fake im Test darf auch
    bytes liefern - beides zulassen."""
    wanted = name.upper()
    for cap in getattr(conn, "capabilities", ()) or ():
        if isinstance(cap, bytes):
            cap = cap.decode("ascii", errors="replace")
        if cap.upper() == wanted:
            return True
    return False


# ---------------------------------------------------------------------------
# Ordner
# ---------------------------------------------------------------------------

_LIST_RE = re.compile(rb'^\((?P<flags>[^)]*)\)\s+(?P<delim>"(?:[^"\\]|\\.)*"|NIL)\s+(?P<name>.+)$')


def _list_line_to_folder(item):
    """Eine Zeile der LIST-Antwort in (name, flags) zerlegen.

    Ordnernamen kommen je nach Server als Literal (dann liefert imaplib ein
    Tupel), in Anführungszeichen oder nackt - alle drei Fälle behandeln.
    """
    if isinstance(item, tuple):
        head, raw_name = item[0], item[1]
        m = re.match(rb"^\(([^)]*)\)", head)
        flags = m.group(1).decode(errors="replace").split() if m else []
        return utf7_decode(raw_name), flags

    m = _LIST_RE.match(item)
    if not m:
        return None, []
    flags = m.group("flags").decode(errors="replace").split()
    raw = m.group("name").decode("ascii", errors="replace").strip()
    if raw.startswith('"') and raw.endswith('"') and len(raw) >= 2:
        raw = raw[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    return utf7_decode(raw), flags


def list_folders(conn):
    """Alle Ordner als Liste von (name, flags). Namen sind Klartext."""
    typ, data = conn.list()
    if typ != "OK":
        raise RuntimeError(f"Ordner konnten nicht gelesen werden: {typ}")
    folders = []
    for item in data or []:
        if not item:
            continue
        name, flags = _list_line_to_folder(item)
        if name:
            folders.append((name, flags))
    return folders


def is_skipped_folder(name, flags):
    """True für Ordner, die nicht gescannt werden sollen (Papierkorb, Entwürfe,
    Spam, Gmails 'Alle Nachrichten')."""
    if any(f.lower() in SKIP_FOLDER_FLAGS for f in flags):
        return True
    return name.strip().lower() in SKIP_FOLDER_NAMES


def find_trash_detail(folders):
    """(Papierkorb, Erkennungsweg): bevorzugt über das \\Trash-Sonderflag
    (RFC 6154), sonst über bekannte Namen.

    Der Erkennungsweg ist "flag", "name" oder None. Er gehört mit ausgegeben:
    ein über den Namen geratener Papierkorb kann der falsche Ordner sein, und
    dann landen die Mails woanders, als der Nutzer erwartet.
    """
    for name, flags in folders:
        if any(f.lower() == "\\trash" for f in flags):
            return name, "flag"
    lookup = {name.strip().lower(): name for name, _ in folders}
    for candidate in TRASH_NAMES:
        if candidate in lookup:
            return lookup[candidate], "name"
    return None, None


def find_trash(folders):
    """Papierkorb-Ordner bestimmen. None, wenn nichts passt."""
    return find_trash_detail(folders)[0]


# ---------------------------------------------------------------------------
# Lesen
# ---------------------------------------------------------------------------

_UID_RE = re.compile(rb"UID\s+(\d+)")
_SIZE_RE = re.compile(rb"RFC822\.SIZE\s+(\d+)")
_FLAGS_RE = re.compile(rb"FLAGS\s+\(([^)]*)\)")
_INTERNALDATE_RE = re.compile(rb'INTERNALDATE\s+"([^"]+)"')

# Anhänge aus der BODYSTRUCTURE. Die Struktur ist beliebig tief verschachtelt;
# sie vollständig zu zerlegen wäre aufwendig und fehleranfällig. Für "hat die
# Mail einen Anhang und wie heißt er" reichen zwei gezielte Suchen:
#   ("attachment" ("filename" "rechnung.pdf"))
# Bewusst NICHT auf den "name"-Parameter des Content-Type geschaut - den tragen
# auch eingebettete Bilder in HTML-Newslettern, die niemand als Anhang versteht.
_ATTACHMENT_RE = re.compile(rb'\(\s*"attachment"', re.IGNORECASE)
_FILENAME_RE = re.compile(rb'"filename"\s+"((?:[^"\\]|\\.)*)"', re.IGNORECASE)
MAX_ATTACHMENT_NAMES = 5


def _header_str(value):
    """msg.get(...) in jedem Fall als str zurückgeben.

    Enthält eine Kopfzeile rohe Bytes, die sich nicht als UTF-8 lesen lassen
    (kaputt kodierte alte oder fremde Mails, ohne korrektes RFC-2047-Encoding),
    liefert Message.get() dafür laut email-Paket (Compat32.header_fetch_parse)
    ein email.header.Header-Objekt statt str. Jedes direkte .strip() darauf
    bricht dann mit AttributeError ab - str() macht zuverlässig Text daraus.
    """
    return "" if value is None else str(value)


def _decode_header_value(raw):
    """MIME-kodierte Kopfzeile (=?utf-8?B?...?=) in lesbaren Text."""
    if not raw:
        return ""
    try:
        return str(make_header(decode_header(raw))).strip()
    except Exception:  # noqa: BLE001 - kaputte Kodierungen kommen real vor
        return str(raw).strip()


def _parse_date(msg, internaldate):
    """Versanddatum als UTC-datetime. Date-Kopfzeile bevorzugt, sonst das vom
    Server vergebene INTERNALDATE (fehlende/kaputte Date-Header sind häufig)."""
    raw = _header_str(msg.get("Date"))
    if raw:
        try:
            dt = email.utils.parsedate_to_datetime(raw)
            if dt is not None:
                return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except Exception:  # noqa: BLE001
            pass
    if internaldate:
        try:
            parsed = imaplib.Internaldate2tuple(b'INTERNALDATE "' + internaldate + b'"')
            if parsed:
                return datetime.fromtimestamp(_time.mktime(parsed), tz=timezone.utc)
        except Exception:  # noqa: BLE001
            pass
    return None


def _balanced_group(data, start):
    """Die geklammerte Gruppe ab data[start] == '(' einschließlich Klammern."""
    depth = 0
    for i in range(start, len(data)):
        char = data[i:i + 1]
        if char == b"(":
            depth += 1
        elif char == b")":
            depth -= 1
            if depth == 0:
                return data[start:i + 1]
    return data[start:]


def _parse_attachments(prefix):
    """(hat_anhang, [dateinamen]) aus dem BODYSTRUCTURE-Teil der Antwort.

    Dateinamen werden ausdrücklich nur aus der Parameterliste der
    "attachment"-Disposition gelesen. Würde man einfach jedes "filename" im
    Datenstrom nehmen, zählte auch das eingebettete Logo eines HTML-Newsletters
    als Anhang - das trägt eine "inline"-Disposition und meint etwas anderes.
    """
    names = []
    has_attachment = False
    for match in _ATTACHMENT_RE.finditer(prefix):
        has_attachment = True
        rest = prefix[match.end():]
        opening = rest.find(b"(")
        # Direkt hinter der Disposition steht entweder die Parameterliste oder
        # NIL. Ein weit entferntes "(" gehört schon zu einem anderen Teil.
        if opening == -1 or opening > 3:
            continue
        found = _FILENAME_RE.search(_balanced_group(rest, opening))
        if not found:
            continue
        raw = found.group(1).replace(b'\\"', b'"').replace(b"\\\\", b"\\")
        name = _decode_header_value(raw.decode("utf-8", errors="replace"))
        if name and name not in names:
            names.append(name)
        if len(names) >= MAX_ATTACHMENT_NAMES:
            break
    return has_attachment, names


def _parse_fetch_item(prefix, header_bytes):
    """Ein FETCH-Ergebnis in ein Mail-Dict umwandeln."""
    uid_m = _UID_RE.search(prefix)
    if not uid_m:
        return None

    flags_m = _FLAGS_RE.search(prefix)
    flags = flags_m.group(1).decode(errors="replace").split() if flags_m else []
    size_m = _SIZE_RE.search(prefix)
    date_m = _INTERNALDATE_RE.search(prefix)

    msg = BytesParser().parsebytes(header_bytes or b"")
    from_raw = _header_str(msg.get("From"))
    display, address = email.utils.parseaddr(from_raw)
    address = (address or "").strip().lower()

    sent = _parse_date(msg, date_m.group(1) if date_m else None)
    has_attachment, attachments = _parse_attachments(prefix)
    return {
        "uid": int(uid_m.group(1)),
        "flags": flags,
        "size": int(size_m.group(1)) if size_m else 0,
        "has_attachment": has_attachment,
        "attachments": attachments,
        "date": sent.isoformat() if sent else "",
        "from_name": _decode_header_value(display),
        "from": address,
        "domain": address.rpartition("@")[2],
        "subject": _decode_header_value(msg.get("Subject", "")),
        "message_id": _header_str(msg.get("Message-ID")).strip(),
        "list_unsubscribe": _header_str(msg.get("List-Unsubscribe")).strip(),
        "list_id": _header_str(msg.get("List-Id")).strip(),
        "precedence": _header_str(msg.get("Precedence")).strip().lower(),
        "auto_submitted": _header_str(msg.get("Auto-Submitted")).strip().lower(),
    }


def _parse_fetch_response(data):
    """imaplib-FETCH-Antwort in Mail-Dicts. Literale kommen als Tupel
    (Vorspann, Nutzdaten), reine Statuszeilen als bytes - beides berücksichtigen."""
    mails = []
    for item in data or []:
        if isinstance(item, tuple) and len(item) >= 2:
            parsed = _parse_fetch_item(item[0], item[1])
            if parsed:
                mails.append(parsed)
    return mails


def folder_stats(conn, folder):
    """(Anzahl Mails, davon als gelöscht markiert) - rein lesend.

    Die zweite Zahl ist der Fingerabdruck des Rückfallwegs: Mails, die kopiert
    und nur markiert wurden, stapeln sich hier an.
    """
    typ, _ = conn.select(_quote(folder), readonly=True)
    if typ != "OK":
        raise RuntimeError(f"Ordner '{folder}' konnte nicht geöffnet werden.")
    counts = []
    for criteria in (("ALL",), ("DELETED",)):
        typ, data = conn.uid("SEARCH", None, *criteria)
        counts.append(len((data[0] or b"").split()) if typ == "OK" else None)
    return counts[0], counts[1]


def get_uidvalidity(conn):
    """UIDVALIDITY des gerade gewählten Ordners. Ändert sich der Wert, sind alle
    früher notierten UIDs wertlos - dann darf nichts mehr gelöscht werden."""
    _typ, data = conn.response("UIDVALIDITY")
    if data and data[0]:
        try:
            return int(data[0])
        except (TypeError, ValueError):
            pass
    return None


def scan_folder(conn, folder, progress=None):
    """Liest die Kopfzeilen aller Mails eines Ordners - streng nur lesend.

    Gibt (mails, uidvalidity) zurück. progress(gelesen, gesamt) wird nach jedem
    Block aufgerufen.
    """
    typ, _ = conn.select(_quote(folder), readonly=True)
    if typ != "OK":
        raise RuntimeError(f"Ordner '{folder}' konnte nicht geöffnet werden.")
    uidvalidity = get_uidvalidity(conn)

    typ, data = conn.uid("SEARCH", None, "ALL")
    if typ != "OK":
        raise RuntimeError(f"Suche in '{folder}' fehlgeschlagen.")
    uids = (data[0] or b"").split()
    total = len(uids)
    if not total:
        return [], uidvalidity

    # BODYSTRUCTURE steht bewusst VOR dem BODY.PEEK-Teil: dann landet es im
    # Vorspann der Antwort und nicht hinter dem Literal, wo es schwerer
    # zu greifen wäre.
    query = ("(FLAGS RFC822.SIZE INTERNALDATE BODYSTRUCTURE "
             f"BODY.PEEK[HEADER.FIELDS ({HEADER_FIELDS})])")
    mails = []
    for start in range(0, total, FETCH_BATCH):
        block = uids[start:start + FETCH_BATCH]
        uid_set = ",".join(u.decode() for u in block)
        typ, data = conn.uid("FETCH", uid_set, query)
        if typ != "OK":
            raise RuntimeError(f"Kopfzeilen aus '{folder}' konnten nicht gelesen werden.")
        mails.extend(_parse_fetch_response(data))
        if progress:
            progress(min(start + FETCH_BATCH, total), total)
    return mails, uidvalidity


def fetch_message_ids(conn, uids):
    """Message-IDs zu bestimmten UIDs im aktuell gewählten Ordner: {uid: message_id}.
    Dient der Sicherheitsprüfung unmittelbar vor dem Verschieben."""
    found = {}
    uids = list(uids)
    for start in range(0, len(uids), FETCH_BATCH):
        block = uids[start:start + FETCH_BATCH]
        uid_set = ",".join(str(u) for u in block)
        typ, data = conn.uid("FETCH", uid_set, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
        if typ != "OK":
            raise RuntimeError("Message-IDs konnten nicht gelesen werden.")
        for item in data or []:
            if isinstance(item, tuple) and len(item) >= 2:
                uid_m = _UID_RE.search(item[0])
                if not uid_m:
                    continue
                msg = BytesParser().parsebytes(item[1] or b"")
                found[int(uid_m.group(1))] = _header_str(msg.get("Message-ID")).strip()
    return found


def fetch_body_snippets(conn, folder, uids, max_bytes=3000):
    """Ersten Teil des Mailtexts zu bestimmten UIDs lesen - rein lesend
    (BODY.PEEK, markiert nichts als gelesen), nur so viele Bytes wie für eine
    Stichwortsuche nötig.

    Bewusst kein MIME-Parser: quoted-printable/base64-kodierte Anteile werden
    nicht aufgelöst, das reicht für eine einfache Stichwortsuche in
    Klartext-Anteilen. Rein Base64-kodierte oder verschlüsselte Mails liefern
    entsprechend keine Treffer - eine bewusste Grenze, kein Fehler.

    Aufrufer muss dafür sorgen, dass der Ordner zur Übergabe passt; hier wird
    er selbst noch einmal schreibgeschützt geöffnet, unabhängig vom Zustand
    der Verbindung.
    """
    uids = list(uids)
    if not uids:
        return {}
    typ, _ = conn.select(_quote(folder), readonly=True)
    if typ != "OK":
        raise RuntimeError(f"Ordner '{folder}' konnte nicht geöffnet werden.")

    found = {}
    query = f"(BODY.PEEK[TEXT]<0.{int(max_bytes)}>)"
    for start in range(0, len(uids), FETCH_BATCH):
        block = uids[start:start + FETCH_BATCH]
        uid_set = ",".join(str(u) for u in block)
        typ, data = conn.uid("FETCH", uid_set, query)
        if typ != "OK":
            raise RuntimeError(f"Textausschnitt aus '{folder}' konnte nicht gelesen werden.")
        for item in data or []:
            if isinstance(item, tuple) and len(item) >= 2:
                uid_m = _UID_RE.search(item[0])
                if not uid_m:
                    continue
                found[int(uid_m.group(1))] = (item[1] or b"").decode("utf-8", errors="replace")
    return found


# ---------------------------------------------------------------------------
# Verschieben
# ---------------------------------------------------------------------------

def uids_still_present(conn, uids):
    """Welche der UIDs im aktuell gewählten Ordner noch liegen.

    Grundlage der Nachprüfung: ein "OK" des Servers heißt nur, dass er den
    Befehl angenommen hat - nicht, dass die Mail auch wirklich weg ist.
    """
    uids = list(uids)
    if not uids:
        return set()
    still = set()
    for start in range(0, len(uids), FETCH_BATCH):
        block = uids[start:start + FETCH_BATCH]
        uid_set = ",".join(str(u) for u in block)
        typ, data = conn.uid("SEARCH", None, "UID", uid_set)
        if typ != "OK":
            # Nicht nachprüfbar: lieber als "noch da" behandeln, damit kein
            # Erfolg gemeldet wird, den niemand geprüft hat.
            still.update(block)
            continue
        for raw in (data[0] or b"").split():
            try:
                still.add(int(raw))
            except ValueError:
                pass
    return still


def move_to_trash(conn, folder, items, trash_folder, expected_uidvalidity=None,
                  dry_run=True, progress=None):
    """Verschiebt Mails eines Ordners in den Papierkorb.

    items: [{"uid": int, "message_id": str}, ...] aus dem Scan.

    Vor dem Verschieben wird geprüft, ob der Ordner noch derselbe ist
    (UIDVALIDITY) und ob hinter jeder UID noch dieselbe Mail steckt
    (Message-ID). Weicht etwas ab, wird die betroffene Mail übersprungen statt
    auf gut Glück gelöscht - eine falsch gelöschte Mail wäre nicht zu heilen.

    NACH dem Verschieben wird nachgesehen, was tatsächlich verschwunden ist.
    Nur das zählt als "moved". Denn ein Server, der auf MOVE mit OK antwortet,
    hat damit noch nichts bewegt - und der Rückfallweg (kopieren + markieren)
    lässt die Mail bewusst liegen.

    Gibt {"moved", "moved_uids", "flagged", "failed", "skipped", "errors",
    "method", "log"} zurück:
      moved      - nachweislich nicht mehr im Ursprungsordner
      moved_uids - welche das waren; nur diese dürfen aus dem Scan fallen
      flagged    - kopiert und als gelöscht markiert, liegt aber noch da
      failed     - Server meldete Erfolg, die Mail ist trotzdem noch da

    Bei dry_run wird der Ordner nur lesend geöffnet, es kann also gar nichts
    geschrieben werden.
    """
    result = {"moved": 0, "moved_uids": [], "flagged": 0, "failed": 0,
              "skipped": 0, "errors": 0, "method": None, "log": []}

    def note(message):
        result["log"].append(message)

    if not items:
        return result

    typ, _ = conn.select(_quote(folder), readonly=dry_run)
    if typ != "OK":
        result["errors"] += len(items)
        note(f"FEHLER: Ordner '{folder}' konnte nicht geöffnet werden - übersprungen.")
        return result

    current_uidvalidity = get_uidvalidity(conn)
    if expected_uidvalidity is not None and current_uidvalidity != expected_uidvalidity:
        result["skipped"] += len(items)
        note(f"ÜBERSPRUNGEN: Ordner '{folder}' hat sich seit dem Scan geändert "
             f"(UIDVALIDITY {expected_uidvalidity} -> {current_uidvalidity}). "
             "Bitte neu scannen.")
        return result

    # Sicherheitsprüfung: steckt hinter jeder UID noch dieselbe Mail?
    wanted = {int(it["uid"]): (it.get("message_id") or "").strip() for it in items}
    try:
        actual = fetch_message_ids(conn, wanted.keys())
    except Exception as e:  # noqa: BLE001
        result["errors"] += len(items)
        note(f"FEHLER: Prüfung in '{folder}' fehlgeschlagen ({e}) - nichts verschoben.")
        return result

    verified = []
    for uid, expected_id in wanted.items():
        found_id = actual.get(uid)
        if found_id is None:
            result["skipped"] += 1
            note(f"ÜBERSPRUNGEN: Mail {uid} in '{folder}' ist nicht mehr da.")
        elif expected_id and found_id and expected_id != found_id:
            result["skipped"] += 1
            note(f"ÜBERSPRUNGEN: Hinter Nummer {uid} in '{folder}' steckt inzwischen "
                 "eine andere Mail.")
        else:
            verified.append(uid)

    if not verified:
        return result

    method = planned_method(conn)
    result["method"] = method

    if dry_run:
        result["moved"] = len(verified)
        note(f"(Testlauf) WÜRDE {len(verified)} Mails aus '{folder}' in "
             f"'{trash_folder}' {METHOD_DE[method]}.")
        if method == "copy_flag":
            note("(Testlauf) ACHTUNG: Dieser Server kann Mails nicht selbst "
                 "verschieben. Die Mails blieben im Ordner liegen und wären nur "
                 "als gelöscht markiert.")
        return result

    # Vor dem Kopieren merken, wie voll der Papierkorb schon ist. Ein "OK" des
    # Servers auf UID COPY heißt laut Protokoll, dass die Mail dort jetzt liegt -
    # real antworten manche Server aber OK, ohne wirklich etwas abzulegen. Genau
    # dieselbe Vorsicht wie bei MOVE, nur auf den Zielordner angewandt.
    trash_count_before = None
    if method == "copy_flag":
        try:
            trash_count_before, _ = folder_stats(conn, trash_folder)
        except Exception:  # noqa: BLE001
            trash_count_before = None
        # folder_stats() hat trash_folder selektiert - zurück zum Quellordner,
        # sonst laufen STORE/EXPUNGE gleich gegen den falschen Ordner.
        typ, _ = conn.select(_quote(folder), readonly=False)
        if typ != "OK":
            result["errors"] += len(verified)
            note(f"FEHLER: Ordner '{folder}' konnte nach der Papierkorb-Prüfung "
                 "nicht wieder geöffnet werden.")
            return result

    target = _quote(trash_folder)
    attempted = []
    for start in range(0, len(verified), FETCH_BATCH):
        block = verified[start:start + FETCH_BATCH]
        uid_set = ",".join(str(u) for u in block)
        try:
            ok = _move_block(conn, uid_set, target, method)
        except Exception as e:  # noqa: BLE001
            result["errors"] += len(block)
            note(f"FEHLER beim Verschieben aus '{folder}': {e}")
            continue
        if ok:
            attempted.extend(block)
        else:
            result["errors"] += len(block)
        if progress:
            progress(min(start + FETCH_BATCH, len(verified)), len(verified))

    if not attempted:
        return result

    # Nachsehen statt glauben. Was noch im Ordner liegt, ist nicht weg.
    try:
        still_there = uids_still_present(conn, attempted)
    except Exception as e:  # noqa: BLE001
        result["failed"] += len(attempted)
        note(f"FEHLER: Nachprüfung in '{folder}' fehlgeschlagen ({e}). Ob die "
             f"{len(attempted)} Mails wirklich weg sind, ist damit ungeklärt.")
        return result

    gone = [u for u in attempted if u not in still_there]
    result["moved_uids"].extend(gone)
    result["moved"] += len(gone)

    if method == "copy_flag" and still_there:
        # Nachsehen statt glauben, Teil 2: nicht nur "liegt die Mail noch im
        # Quellordner", sondern auch "ist im Papierkorb wirklich etwas Neues
        # angekommen". Ein Server, der COPY faelschlich mit OK bestaetigt,
        # wuerde sonst als Erfolg durchgehen, obwohl nichts kopiert wurde.
        copy_confirmed = True
        if trash_count_before is not None:
            try:
                trash_count_after, _ = folder_stats(conn, trash_folder)
                if trash_count_after - trash_count_before < len(still_there):
                    copy_confirmed = False
            except Exception:  # noqa: BLE001
                pass  # nicht nachpruefbar - im Zweifel wie gemeldet behandeln

        if copy_confirmed:
            # Hier ist Liegenbleiben der Normalfall, kein Fehler.
            result["flagged"] += len(still_there)
            note(f"ACHTUNG: Server kann weder MOVE noch UID EXPUNGE. {len(still_there)} "
                 f"Mails wurden nach '{trash_folder}' KOPIERT und in '{folder}' nur als "
                 "gelöscht markiert - sie liegen dort weiter und viele Mailprogramme "
                 "zeigen sie ganz normal an.")
        else:
            result["failed"] += len(still_there)
            note(f"FEHLER: Server hat das Kopieren von {len(still_there)} Mails nach "
                 f"'{trash_folder}' mit OK bestätigt, dort ist aber nichts Neues "
                 "angekommen. Vermutlich wurde nichts kopiert - die Mails sind "
                 f"unverändert in '{folder}'. Bitte im Postfach von Hand nachsehen.")
    elif still_there:
        result["failed"] += len(still_there)
        note(f"FEHLER: {len(still_there)} Mails liegen trotz erfolgreicher "
             f"Server-Antwort noch in '{folder}'. Es wurde nichts gelöscht.")
    return result


METHOD_DE = {
    "move": "verschieben",
    "copy_expunge": "kopieren und im Ordner entfernen",
    "copy_flag": "kopieren und nur als gelöscht markieren",
}


def planned_method(conn):
    """Welcher der drei Wege bei diesem Server benutzt wird - vorab bestimmbar,
    damit schon der Testlauf sagen kann, was passieren wird."""
    if has_capability(conn, "MOVE"):
        return "move"
    if has_capability(conn, "UIDPLUS"):
        return "copy_expunge"
    return "copy_flag"


def _move_block(conn, uid_set, target, method):
    """Ein Block Mails ins Ziel. Drei Wege, absteigend nach Sauberkeit.

    Gibt zurück, ob der Server die Befehle angenommen hat - NICHT, ob die Mails
    weg sind. Das stellt erst die Nachprüfung in move_to_trash() fest.
    """
    # 1. MOVE - der Normalfall bei iCloud, Gmail und allen aktuellen Servern.
    if method == "move":
        typ, _ = conn.uid("MOVE", uid_set, target)
        return typ == "OK"

    # 2. Kein MOVE: kopieren, als gelöscht markieren, gezielt entfernen.
    typ, _ = conn.uid("COPY", uid_set, target)
    if typ != "OK":
        return False
    typ, _ = conn.uid("STORE", uid_set, "+FLAGS", "(\\Deleted)")
    if typ != "OK":
        return False

    if method == "copy_expunge":
        conn.uid("EXPUNGE", uid_set)
        return True

    # 3. Weder MOVE noch UIDPLUS: NICHT expunge aufrufen. Ein blankes EXPUNGE
    #    würde jede als gelöscht markierte Mail im Ordner endgültig entfernen -
    #    auch solche, die der Nutzer selbst mal markiert hat. Die Kopie liegt
    #    im Papierkorb, das Original ist markiert; der Mail-Client räumt auf.
    return True
