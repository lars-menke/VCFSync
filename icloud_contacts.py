"""
iCloud Kontakte - CardDAV Export & Import
==========================================
Verwendung:
  Export:  python icloud_contacts.py export --output meine_kontakte.vcf
  Import:  python icloud_contacts.py import --input bearbeitete_kontakte.vcf

Zugangsdaten werden beim ersten Start abgefragt und in .env gespeichert,
oder als Umgebungsvariablen gesetzt:
  ICLOUD_USER=lars@icloud.com
  ICLOUD_PASS=xxxx-xxxx-xxxx-xxxx   (app-spezifisches Passwort)
"""

import argparse
import base64
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from getpass import getpass
from pathlib import Path
from urllib.parse import urljoin

import requests
from requests.auth import HTTPBasicAuth

# ---------------------------------------------------------------------------
# Konstanten
# ---------------------------------------------------------------------------
CARDDAV_BASE = "https://contacts.icloud.com"
ENV_FILE = Path(".env")
NS = {
    "d": "DAV:",
    "card": "urn:ietf:params:xml:ns:carddav",
    "cs": "http://calendarserver.org/ns/",
}

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def load_env():
    """Lädt Zugangsdaten aus .env-Datei (KEY=VALUE Zeilen)."""
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def get_credentials():
    """Gibt (user, password) zurück, fragt interaktiv falls nötig."""
    load_env()
    user = os.environ.get("ICLOUD_USER") or input("Apple-ID (E-Mail): ").strip()
    pw   = os.environ.get("ICLOUD_PASS") or getpass("App-spezifisches Passwort: ")

    if not os.environ.get("ICLOUD_USER"):
        save = input("Zugangsdaten in .env speichern? (j/n): ").strip().lower()
        if save == "j":
            ENV_FILE.write_text(f"ICLOUD_USER={user}\nICLOUD_PASS={pw}\n")
            print(".env gespeichert.")
    return user, pw


def get_principal_url(session, user):
    """
    Ermittelt die Principal-URL des Nutzers via PROPFIND auf /.
    iCloud leitet dabei auf /[dsnumber]/ weiter.
    """
    body = """<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:">
  <d:prop>
    <d:current-user-principal/>
  </d:prop>
</d:propfind>"""

    r = session.request(
        "PROPFIND",
        f"{CARDDAV_BASE}/",
        data=body.encode(),
        headers={"Depth": "0", "Content-Type": "application/xml"},
    )
    r.raise_for_status()
    root = ET.fromstring(r.text)
    href = root.find(".//d:current-user-principal/d:href", NS)
    if href is None:
        raise RuntimeError("Konnte Principal-URL nicht ermitteln.")
    # iCloud liefert hier teils absolute URLs (eigener Shard-Server),
    # teils relative Pfade — urljoin behandelt beide Fälle korrekt.
    return urljoin(r.url, href.text)


def get_all_addressbook_urls(session, principal_url):
    """
    Ermittelt ALLE Adressbuch-Collections des Nutzers.
    Gibt Liste von (url, displayname) zurück.
    """
    body = """<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:" xmlns:card="urn:ietf:params:xml:ns:carddav">
  <d:prop>
    <card:addressbook-home-set/>
  </d:prop>
</d:propfind>"""

    r = session.request(
        "PROPFIND",
        principal_url,
        data=body.encode(),
        headers={"Depth": "0", "Content-Type": "application/xml"},
    )
    r.raise_for_status()
    root = ET.fromstring(r.text)
    href = root.find(".//card:addressbook-home-set/d:href", NS)
    if href is None:
        raise RuntimeError("Konnte Adressbuch-Home-Set nicht ermitteln.")

    ab_home = urljoin(r.url, href.text)

    # Listet verfügbare Adressbücher auf
    r2 = session.request(
        "PROPFIND",
        ab_home,
        data="""<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:">
  <d:prop><d:displayname/><d:resourcetype/></d:prop>
</d:propfind>""".encode(),
        headers={"Depth": "1", "Content-Type": "application/xml"},
    )
    r2.raise_for_status()
    root2 = ET.fromstring(r2.text)

    books = []
    for resp in root2.findall("d:response", NS):
        rt = resp.find(".//d:resourcetype", NS)
        if rt is not None and rt.find("card:addressbook", NS) is not None:
            h = resp.find("d:href", NS)
            dn = resp.find(".//d:displayname", NS)
            if h is not None:
                name = dn.text if (dn is not None and dn.text) else "(ohne Namen)"
                books.append((urljoin(ab_home, h.text), name))

    if not books:
        raise RuntimeError("Kein Adressbuch gefunden.")
    return books


def get_addressbook_url(session, principal_url):
    """Ermittelt die (erste) Adressbuch-URL – für Import/Löschen."""
    return get_all_addressbook_urls(session, principal_url)[0][0]


def _parse_contact_hrefs(xml_text, addressbook_url):
    """Extrahiert {url: etag} aller Kontakt-Ressourcen aus einer Multistatus-Antwort."""
    found = {}
    root = ET.fromstring(xml_text)
    ab = addressbook_url.rstrip("/")
    for resp in root.findall("d:response", NS):
        rt = resp.find(".//d:resourcetype/d:collection", NS)
        if rt is not None:
            continue  # Überspringe Sammlungen (das Adressbuch selbst)
        href = resp.find("d:href", NS)
        if href is None or not href.text:
            continue
        etag = resp.find(".//d:getetag", NS)
        # Absolute URL aufbauen (iCloud-hrefs können relativ oder absolut sein)
        url = urljoin(addressbook_url, href.text)
        if url.rstrip("/") == ab:
            continue  # das Adressbuch selbst
        found[url] = etag.text if etag is not None else ""
    return found


PROPFIND_HREFS_BODY = """<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:">
  <d:prop>
    <d:getetag/>
    <d:resourcetype/>
  </d:prop>
</d:propfind>"""

ADDRESSBOOK_QUERY_BODY = """<?xml version="1.0" encoding="utf-8"?>
<card:addressbook-query xmlns:d="DAV:" xmlns:card="urn:ietf:params:xml:ns:carddav">
  <d:prop>
    <d:getetag/>
  </d:prop>
  <card:filter>
    <card:prop-filter name="UID"/>
  </card:filter>
</card:addressbook-query>"""

SYNC_COLLECTION_BODY = """<?xml version="1.0" encoding="utf-8"?>
<d:sync-collection xmlns:d="DAV:">
  <d:sync-token></d:sync-token>
  <d:sync-level>1</d:sync-level>
  <d:prop>
    <d:getetag/>
  </d:prop>
</d:sync-collection>"""


def _enumerate(session, addressbook_url, method, body):
    """Führt PROPFIND/REPORT aus und liefert {url: etag}; bei Fehler ({}, Fehlertext)."""
    try:
        r = session.request(
            method,
            addressbook_url,
            data=body.encode(),
            headers={"Depth": "1", "Content-Type": "application/xml"},
        )
        if r.status_code in (200, 207):
            return _parse_contact_hrefs(r.text, addressbook_url), None
        return {}, f"HTTP {r.status_code}"
    except Exception as e:
        return {}, str(e)


def fetch_all_contact_hrefs(session, addressbook_url, verbose=True):
    """
    Gibt Liste von (href, etag) aller Kontakte im Adressbuch zurück.

    iClouds PROPFIND listet in manchen Fällen einzelne Kontakte nicht mit auf.
    Deshalb werden drei Enumerationsverfahren kombiniert und ihre Ergebnisse
    vereinigt (Union kann nur mehr Kontakte finden, nie weniger):
      1. PROPFIND (Depth:1)
      2. addressbook-query REPORT (alle Karten mit UID)
      3. sync-collection REPORT (vollständige Sync-Auflistung, RFC 6578)
    Mit verbose=True wird pro Verfahren die Trefferzahl ausgegeben (Diagnose).
    """
    pf, pf_err = _enumerate(session, addressbook_url, "PROPFIND", PROPFIND_HREFS_BODY)
    aq, aq_err = _enumerate(session, addressbook_url, "REPORT", ADDRESSBOOK_QUERY_BODY)
    sc, sc_err = _enumerate(session, addressbook_url, "REPORT", SYNC_COLLECTION_BODY)

    merged = {}
    for d in (pf, aq, sc):
        for url, etag in d.items():
            merged.setdefault(url, etag)

    if verbose:
        def fmt(name, d, err):
            return f"{name}={len(d)}" + (f" (Fehler: {err})" if err else "")
        print("    Enumeration: " + " | ".join([
            fmt("PROPFIND", pf, pf_err),
            fmt("addressbook-query", aq, aq_err),
            fmt("sync-collection", sc, sc_err),
        ]) + f" | Union={len(merged)}")

    if not merged and pf_err:
        raise RuntimeError(f"Konnte Kontakte nicht auflisten (PROPFIND: {pf_err}).")
    return list(merged.items())


def fetch_vcard(session, url):
    """Lädt eine einzelne vCard per GET. Erwartet eine absolute URL."""
    r = session.get(url, headers={"Accept": "text/vcard"})
    r.raise_for_status()
    return r.text


def extract_uid(vcard_text):
    """Extrahiert den UID-Wert aus einer vCard."""
    m = re.search(r"^UID[^:]*:(.+)$", vcard_text, re.MULTILINE)
    return m.group(1).strip() if m else None


# ---------------------------------------------------------------------------
# Foto-Einbettung (URI -> Base64)
# ---------------------------------------------------------------------------

def unfold_vcard(text):
    """Hebt vCard-Zeilenfaltung auf (Fortsetzungszeilen beginnen mit Space/Tab)."""
    out = []
    for line in text.splitlines():
        if line[:1] in (" ", "\t") and out:
            out[-1] += line[1:]
        else:
            out.append(line)
    return out


def fold_line(line, limit=73):
    """Faltet eine lange vCard-Zeile (Folding nach RFC 6350: max ~75 Oktette)."""
    if len(line) <= limit:
        return line
    parts = [line[:limit]]
    rest = line[limit:]
    while rest:
        parts.append(" " + rest[:limit - 1])
        rest = rest[limit - 1:]
    return "\r\n".join(parts)


def _photo_subtype(content_type, data):
    """Ermittelt den Bild-Subtyp (JPEG/PNG/...) aus Content-Type oder Magic-Bytes."""
    if content_type:
        ct = content_type.lower()
        if "jpeg" in ct or "jpg" in ct:
            return "JPEG"
        if "png" in ct:
            return "PNG"
        if "gif" in ct:
            return "GIF"
        if "webp" in ct:
            return "WEBP"
    if data[:3] == b"\xff\xd8\xff":
        return "JPEG"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "PNG"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "GIF"
    return "JPEG"  # iCloud liefert in aller Regel JPEG


def embed_photos(session, vcard_text):
    """
    Ersetzt PHOTO-Verweise (PHOTO;...;VALUE=uri:https://...) durch eingebettete
    Base64-Daten (PHOTO;ENCODING=b;TYPE=...). Gibt (vcard_text, anzahl) zurück.
    Bei Fehlern bleibt der ursprüngliche Verweis erhalten.
    """
    lines = unfold_vcard(vcard_text)
    embedded = 0
    for idx, line in enumerate(lines):
        if not line.upper().startswith("PHOTO"):
            continue
        head, sep, value = line.partition(":")
        if not sep or not value.lower().startswith(("http://", "https://")):
            continue  # bereits eingebettet oder kein URL-Verweis

        # iCloud faltet bei "geteilten Fotos" die Property X-SHARED-PHOTO-
        # DISPLAY-PREF direkt an die PHOTO-URL an. Saubere Karten-URL bis zum
        # Hash extrahieren, angehängten Property-Müll abschneiden.
        m = re.match(r"https?://\S*?/ck/card/[0-9a-fA-F]+", value.strip())
        photo_url = m.group(0) if m else value.strip().split("X-SHARED-PHOTO-DISPLAY-PREF")[0]

        data = resp = None
        for attempt in range(3):  # kleiner Retry gegen transiente Aussetzer
            try:
                resp = session.get(photo_url, timeout=30)
                resp.raise_for_status()
                data = resp.content
                break
            except Exception as e:
                if attempt == 2:
                    print(f"\n  WARNUNG: Foto konnte nicht geladen werden ({e})")
                else:
                    time.sleep(0.5 * (attempt + 1))
        if not data:
            continue

        subtype = _photo_subtype(resp.headers.get("Content-Type", ""), data)
        b64 = base64.b64encode(data).decode("ascii")

        # Parameter aus dem Kopf übernehmen, VALUE entfernen, ENCODING/TYPE setzen
        params = [p for p in head.split(";")[1:] if not p.upper().startswith("VALUE=")]
        params = [p for p in params if not p.upper().startswith("ENCODING=")
                  and not p.upper().startswith("TYPE=")]
        new_head = ";".join(["PHOTO", *params, "ENCODING=b", f"TYPE={subtype}"])
        lines[idx] = f"{new_head}:{b64}"
        embedded += 1

    # Alle Zeilen RFC-6350-konform falten (Unfold/Rejoin hätte sonst lange
    # iCloud-Felder wie X-ADDRESSING-GRAMMAR ungefaltet gelassen).
    return "\r\n".join(fold_line(l) for l in lines), embedded


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def cmd_export(args, session, addressbook_urls):
    output = Path(args.output)
    # Über ALLE Adressbücher hinweg sammeln und per URL vereinigen.
    if isinstance(addressbook_urls, str):
        addressbook_urls = [(addressbook_urls, "")]
    print(f"Durchsuche {len(addressbook_urls)} Adressbuch/Adressbücher ...")
    href_map = {}
    for ab_url, ab_name in addressbook_urls:
        print(f"  Adressbuch '{ab_name}':")
        for href, etag in fetch_all_contact_hrefs(session, ab_url):
            href_map.setdefault(href, etag)
    hrefs = list(href_map.items())
    total = len(hrefs)
    print(f"{total} Kontakte gefunden (über alle Adressbücher).")

    embed = not getattr(args, "skip_photos", False)
    photo_count = 0
    vcards = []
    for i, (href, etag) in enumerate(hrefs, 1):
        print(f"\r  Lade {i}/{total} ...", end="", flush=True)
        try:
            vcard = fetch_vcard(session, href)
            if embed:
                # iCloud liefert Fotos als URL-Verweis statt eingebettet -
                # hier werden sie nachgeladen und als Base64 eingebettet.
                vcard, n = embed_photos(session, vcard)
                photo_count += n
            vcards.append(vcard.strip())
        except Exception as e:
            print(f"\n  WARNUNG: {href} konnte nicht geladen werden: {e}")
        time.sleep(0.05)  # sanftes Rate-Limiting

    print(f"\n{len(vcards)} Kontakte geladen.")
    if embed:
        print(f"{photo_count} Fotos als Base64 eingebettet.")

    combined = "\r\n".join(vcards) + "\r\n"
    output.write_text(combined, encoding="utf-8")
    print(f"Gespeichert: {output} ({output.stat().st_size / 1024:.1f} KB)")


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def split_vcards(text):
    """Zerlegt eine VCF-Datei in einzelne vCard-Strings."""
    cards = []
    current = []
    for line in text.splitlines():
        current.append(line)
        if line.strip().upper() == "END:VCARD":
            cards.append("\r\n".join(current))
            current = []
    return cards


def fetch_existing_uids(session, addressbook_url):
    """
    Gibt Dict {uid: href} aller vorhandenen Kontakte zurück.
    Lädt dazu jede vCard und liest die UID.
    """
    print("Lese vorhandene Kontakte aus iCloud (für UID-Abgleich) ...")
    hrefs = fetch_all_contact_hrefs(session, addressbook_url)
    uid_map = {}
    for i, (href, _) in enumerate(hrefs, 1):
        print(f"\r  {i}/{len(hrefs)} ...", end="", flush=True)
        try:
            vcard = fetch_vcard(session, href)
            uid = extract_uid(vcard)
            if uid:
                uid_map[uid] = href
        except Exception as e:
            print(f"\n  WARNUNG: {href}: {e}")
        time.sleep(0.05)
    print(f"\n{len(uid_map)} UIDs eingelesen.")
    return uid_map


def put_vcard(session, url, vcard_text, etag=None):
    """Schreibt eine vCard per PUT (Update oder Neuanlage)."""
    headers = {
        "Content-Type": "text/vcard; charset=utf-8",
    }
    if etag:
        headers["If-Match"] = etag  # Optimistic Locking bei Updates
    r = session.put(url, data=vcard_text.encode("utf-8"), headers=headers)
    return r


def cmd_import(args, session, addressbook_url):
    input_file = Path(args.input)
    if not input_file.exists():
        print(f"Datei nicht gefunden: {input_file}")
        sys.exit(1)

    vcards = split_vcards(input_file.read_text(encoding="utf-8"))
    print(f"{len(vcards)} Kontakte in {input_file} gefunden.")

    # Vorhandene UIDs aus iCloud laden
    uid_map = fetch_existing_uids(session, addressbook_url)

    dry_run = getattr(args, "dry_run", False)

    updated = 0
    created = 0
    errors  = 0

    for i, vcard in enumerate(vcards, 1):
        uid = extract_uid(vcard)
        if not uid:
            print(f"  [{i}] KEIN UID - übersprungen")
            errors += 1
            continue

        if uid in uid_map:
            # Update: vorhandener Kontakt (uid_map enthält bereits absolute URLs)
            url = uid_map[uid]
            if dry_run:
                print(f"  [{i}] (dry-run) WÜRDE AKTUALISIEREN: {uid[:20]}...")
                updated += 1
                continue
            r = put_vcard(session, url, vcard)
            if r.status_code in (200, 201, 204):
                print(f"  [{i}] AKTUALISIERT: {uid[:20]}...")
                updated += 1
            else:
                print(f"  [{i}] FEHLER Update {r.status_code}: {uid[:20]}")
                errors += 1
        else:
            # Neuanlage: neue .vcf-Ressource im Adressbuch
            safe_uid = re.sub(r"[^a-zA-Z0-9\-]", "", uid)
            url = addressbook_url.rstrip("/") + f"/{safe_uid}.vcf"
            if dry_run:
                print(f"  [{i}] (dry-run) WÜRDE NEU ANLEGEN: {uid[:20]}...")
                created += 1
                continue
            r = put_vcard(session, url, vcard)
            if r.status_code in (200, 201, 204):
                print(f"  [{i}] NEU ANGELEGT: {uid[:20]}...")
                created += 1
            else:
                print(f"  [{i}] FEHLER Neuanlage {r.status_code}: {uid[:20]}")
                errors += 1

        time.sleep(0.1)  # sanftes Rate-Limiting

    print(f"\nFertig: {updated} aktualisiert, {created} neu, {errors} Fehler.")


# ---------------------------------------------------------------------------
# Löschen
# ---------------------------------------------------------------------------

def delete_contact(session, url):
    """Löscht eine Kontakt-Ressource per HTTP DELETE."""
    return session.delete(url)


def cmd_delete(args, session, addressbook_url):
    """Löscht Kontakte anhand ihrer UID (aus --uid und/oder --input VCF)."""
    uids = []
    if getattr(args, "input", None):
        input_file = Path(args.input)
        if not input_file.exists():
            print(f"Datei nicht gefunden: {input_file}")
            sys.exit(1)
        for vcard in split_vcards(input_file.read_text(encoding="utf-8")):
            u = extract_uid(vcard)
            if u:
                uids.append(u)
    for u in getattr(args, "uid", None) or []:
        uids.append(u.strip())
    uids = list(dict.fromkeys(uids))  # Reihenfolge erhalten, Duplikate raus

    if not uids:
        print("Keine UIDs angegeben. Nutze --input <vcf> und/oder --uid <UID>.")
        sys.exit(1)
    print(f"{len(uids)} zu löschende Kontakt-UID(s).")

    uid_map = fetch_existing_uids(session, addressbook_url)
    dry_run = getattr(args, "dry_run", False)

    deleted = missing = errors = 0
    for i, uid in enumerate(uids, 1):
        if uid not in uid_map:
            print(f"  [{i}] NICHT GEFUNDEN (evtl. schon gelöscht): {uid[:20]}...")
            missing += 1
            continue
        url = uid_map[uid]
        if dry_run:
            print(f"  [{i}] (dry-run) WÜRDE LÖSCHEN: {uid[:20]}...")
            deleted += 1
            continue
        r = delete_contact(session, url)
        if r.status_code in (200, 204):
            print(f"  [{i}] GELÖSCHT: {uid[:20]}...")
            deleted += 1
        else:
            print(f"  [{i}] FEHLER Löschen {r.status_code}: {uid[:20]}")
            errors += 1
        time.sleep(0.1)  # sanftes Rate-Limiting

    print(f"\nFertig: {deleted} gelöscht, {missing} nicht gefunden, {errors} Fehler.")


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="iCloud Kontakte via CardDAV exportieren und importieren"
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    exp = sub.add_parser("export", help="Kontakte aus iCloud exportieren")
    exp.add_argument("--output", default="icloud_kontakte.vcf",
                     help="Zieldatei (Standard: icloud_kontakte.vcf)")
    exp.add_argument("--skip-photos", action="store_true",
                     help="Fotos nicht nachladen/einbetten (nur URL-Verweise, schneller)")

    imp = sub.add_parser("import", help="Bearbeitete VCF in iCloud zurückspielen")
    imp.add_argument("--input", required=True,
                     help="Zu importierende VCF-Datei")
    imp.add_argument("--dry-run", action="store_true",
                     help="Nur simulieren, nichts schreiben")

    dele = sub.add_parser("delete", help="Kontakte aus iCloud löschen (per UID oder VCF)")
    dele.add_argument("--input",
                      help="VCF-Datei; alle darin enthaltenen UIDs werden gelöscht")
    dele.add_argument("--uid", action="append",
                      help="Einzelne UID zum Löschen (mehrfach angebbar)")
    dele.add_argument("--dry-run", action="store_true",
                      help="Nur simulieren, nichts löschen")

    args = parser.parse_args()

    # Session aufbauen
    user, pw = get_credentials()
    session = requests.Session()
    session.auth = HTTPBasicAuth(user, pw)
    session.headers.update({"User-Agent": "iCloudContactsScript/1.0"})

    print("Verbinde mit iCloud ...")
    try:
        principal_url  = get_principal_url(session, user)
        books = get_all_addressbook_urls(session, principal_url)
        print(f"{len(books)} Adressbuch/Adressbücher gefunden:")
        for url, name in books:
            print(f"  - {name}: {url}")
    except Exception as e:
        print(f"Verbindungsfehler: {e}")
        print("Hinweis: Stelle sicher, dass du ein app-spezifisches Passwort verwendest")
        print("         (appleid.apple.com > Sicherheit > App-spezifische Passwörter)")
        sys.exit(1)

    primary_url = books[0][0]  # erstes Adressbuch für Import/Neuanlage/Löschen
    if args.cmd == "export":
        cmd_export(args, session, books)
    elif args.cmd == "import":
        if getattr(args, "dry_run", False):
            print("DRY-RUN: Es werden keine Daten geschrieben.")
        cmd_import(args, session, primary_url)
    elif args.cmd == "delete":
        if getattr(args, "dry_run", False):
            print("DRY-RUN: Es werden keine Daten gelöscht.")
        cmd_delete(args, session, primary_url)


if __name__ == "__main__":
    main()
