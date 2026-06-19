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


def get_addressbook_url(session, principal_url):
    """
    Ermittelt die Adressbuch-URL via PROPFIND auf der Principal-URL.
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
        raise RuntimeError("Konnte Adressbuch-URL nicht ermitteln.")

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

    for resp in root2.findall("d:response", NS):
        rt = resp.find(".//d:resourcetype", NS)
        if rt is not None and rt.find("card:addressbook", NS) is not None:
            h = resp.find("d:href", NS)
            if h is not None:
                return urljoin(ab_home, h.text)

    raise RuntimeError("Kein Adressbuch gefunden.")


def fetch_all_contact_hrefs(session, addressbook_url):
    """
    Gibt Liste von (href, etag) aller Kontakte im Adressbuch zurück.
    """
    body = """<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:">
  <d:prop>
    <d:getetag/>
    <d:resourcetype/>
  </d:prop>
</d:propfind>"""

    r = session.request(
        "PROPFIND",
        addressbook_url,
        data=body.encode(),
        headers={"Depth": "1", "Content-Type": "application/xml"},
    )
    r.raise_for_status()
    root = ET.fromstring(r.text)

    contacts = []
    for resp in root.findall("d:response", NS):
        rt = resp.find(".//d:resourcetype/d:collection", NS)
        if rt is not None:
            continue  # Überspringe Sammlungen (das Adressbuch selbst)
        href = resp.find("d:href", NS)
        etag = resp.find(".//d:getetag", NS)
        if href is not None:
            # Absolute URL aufbauen (iCloud-hrefs können relativ oder absolut sein)
            url = urljoin(addressbook_url, href.text)
            contacts.append((url, etag.text if etag is not None else ""))
    return contacts


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
# Export
# ---------------------------------------------------------------------------

def cmd_export(args, session, addressbook_url):
    output = Path(args.output)
    print(f"Lade Kontaktliste aus {addressbook_url} ...")
    hrefs = fetch_all_contact_hrefs(session, addressbook_url)
    total = len(hrefs)
    print(f"{total} Kontakte gefunden.")

    vcards = []
    for i, (href, etag) in enumerate(hrefs, 1):
        print(f"\r  Lade {i}/{total} ...", end="", flush=True)
        try:
            vcard = fetch_vcard(session, href)
            vcards.append(vcard.strip())
        except Exception as e:
            print(f"\n  WARNUNG: {href} konnte nicht geladen werden: {e}")
        time.sleep(0.05)  # sanftes Rate-Limiting

    print(f"\n{len(vcards)} Kontakte geladen.")

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

    imp = sub.add_parser("import", help="Bearbeitete VCF in iCloud zurückspielen")
    imp.add_argument("--input", required=True,
                     help="Zu importierende VCF-Datei")
    imp.add_argument("--dry-run", action="store_true",
                     help="Nur simulieren, nichts schreiben")

    args = parser.parse_args()

    # Session aufbauen
    user, pw = get_credentials()
    session = requests.Session()
    session.auth = HTTPBasicAuth(user, pw)
    session.headers.update({"User-Agent": "iCloudContactsScript/1.0"})

    print("Verbinde mit iCloud ...")
    try:
        principal_url  = get_principal_url(session, user)
        addressbook_url = get_addressbook_url(session, principal_url)
        print(f"Adressbuch: {addressbook_url}")
    except Exception as e:
        print(f"Verbindungsfehler: {e}")
        print("Hinweis: Stelle sicher, dass du ein app-spezifisches Passwort verwendest")
        print("         (appleid.apple.com > Sicherheit > App-spezifische Passwörter)")
        sys.exit(1)

    if args.cmd == "export":
        cmd_export(args, session, addressbook_url)
    elif args.cmd == "import":
        if getattr(args, "dry_run", False):
            print("DRY-RUN: Es werden keine Daten geschrieben.")
        cmd_import(args, session, addressbook_url)


if __name__ == "__main__":
    main()
