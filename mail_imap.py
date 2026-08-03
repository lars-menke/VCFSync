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

Gelesen werden ausschließlich Kopfzeilen (BODY.PEEK[HEADER.FIELDS ...]),
niemals Mail-Texte. Das reicht für die Bewertung, ist deutlich schneller und
lässt den Gelesen-Status unangetastet (PEEK).
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


def find_trash(folders):
    """Papierkorb-Ordner bestimmen: bevorzugt über das \\Trash-Sonderflag
    (RFC 6154), sonst über bekannte Namen. None, wenn nichts passt."""
    for name, flags in folders:
        if any(f.lower() == "\\trash" for f in flags):
            return name
    lookup = {name.strip().lower(): name for name, _ in folders}
    for candidate in TRASH_NAMES:
        if candidate in lookup:
            return lookup[candidate]
    return None


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
    raw = msg.get("Date")
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
    from_raw = msg.get("From", "")
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
        "message_id": (msg.get("Message-ID") or "").strip(),
        "list_unsubscribe": (msg.get("List-Unsubscribe") or "").strip(),
        "list_id": (msg.get("List-Id") or "").strip(),
        "precedence": (msg.get("Precedence") or "").strip().lower(),
        "auto_submitted": (msg.get("Auto-Submitted") or "").strip().lower(),
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
                found[int(uid_m.group(1))] = (msg.get("Message-ID") or "").strip()
    return found


# ---------------------------------------------------------------------------
# Verschieben
# ---------------------------------------------------------------------------

def move_to_trash(conn, folder, items, trash_folder, expected_uidvalidity=None,
                  dry_run=True, progress=None):
    """Verschiebt Mails eines Ordners in den Papierkorb.

    items: [{"uid": int, "message_id": str}, ...] aus dem Scan.

    Vor dem Verschieben wird geprüft, ob der Ordner noch derselbe ist
    (UIDVALIDITY) und ob hinter jeder UID noch dieselbe Mail steckt
    (Message-ID). Weicht etwas ab, wird die betroffene Mail übersprungen statt
    auf gut Glück gelöscht - eine falsch gelöschte Mail wäre nicht zu heilen.

    Gibt {"moved", "skipped", "errors", "log"} zurück. Bei dry_run wird der
    Ordner nur lesend geöffnet, es kann also gar nichts geschrieben werden.
    """
    result = {"moved": 0, "skipped": 0, "errors": 0, "log": []}

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

    if dry_run:
        result["moved"] = len(verified)
        note(f"(Testlauf) WÜRDE {len(verified)} Mails aus '{folder}' in "
             f"'{trash_folder}' verschieben.")
        return result

    target = _quote(trash_folder)
    for start in range(0, len(verified), FETCH_BATCH):
        block = verified[start:start + FETCH_BATCH]
        uid_set = ",".join(str(u) for u in block)
        try:
            moved = _move_block(conn, uid_set, target, note)
        except Exception as e:  # noqa: BLE001
            result["errors"] += len(block)
            note(f"FEHLER beim Verschieben aus '{folder}': {e}")
            continue
        if moved:
            result["moved"] += len(block)
        else:
            result["errors"] += len(block)
        if progress:
            progress(min(start + FETCH_BATCH, len(verified)), len(verified))
    return result


def _move_block(conn, uid_set, target, note):
    """Ein Block Mails ins Ziel. Drei Wege, absteigend nach Sauberkeit."""
    # 1. MOVE - der Normalfall bei iCloud, Gmail und allen aktuellen Servern.
    if has_capability(conn, "MOVE"):
        typ, _ = conn.uid("MOVE", uid_set, target)
        return typ == "OK"

    # 2. Kein MOVE: kopieren, als gelöscht markieren, gezielt entfernen.
    typ, _ = conn.uid("COPY", uid_set, target)
    if typ != "OK":
        return False
    typ, _ = conn.uid("STORE", uid_set, "+FLAGS", "(\\Deleted)")
    if typ != "OK":
        return False

    if has_capability(conn, "UIDPLUS"):
        conn.uid("EXPUNGE", uid_set)
        return True

    # 3. Weder MOVE noch UIDPLUS: NICHT expunge aufrufen. Ein blankes EXPUNGE
    #    würde jede als gelöscht markierte Mail im Ordner endgültig entfernen -
    #    auch solche, die der Nutzer selbst mal markiert hat. Die Kopie liegt
    #    im Papierkorb, das Original ist markiert; der Mail-Client räumt auf.
    note("Hinweis: Server kann weder MOVE noch UID EXPUNGE - Mails wurden in den "
         "Papierkorb kopiert und im Quellordner als gelöscht markiert.")
    return True
