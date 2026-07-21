"""
Diagnose-Werkzeug: zeigt die aktuelle iCloud-UID einer Karte anhand eines
Namens-Suchbegriffs (sucht im gesamten vCard-Text, Groß-/Kleinschreibung egal).

Aufruf (im gleichen Ordner wie icloud_contacts.py):
    python3 find_contact.py "Kramer"
"""
import sys

import requests
from requests.auth import HTTPBasicAuth

from icloud_contacts import (
    get_credentials,
    get_principal_url,
    get_all_addressbook_urls,
    fetch_all_contact_hrefs,
    fetch_vcard,
    extract_uid,
)

search = sys.argv[1] if len(sys.argv) > 1 else "Kramer"

user, pw = get_credentials()
session = requests.Session()
session.auth = HTTPBasicAuth(user, pw)
session.headers.update({"User-Agent": "iCloudContactsScript/1.0"})

print("Verbinde mit iCloud ...")
principal_url = get_principal_url(session, user)
books = get_all_addressbook_urls(session, principal_url)
print(f"{len(books)} Adressbuch/Adressbücher gefunden.")

found = 0
for ab_url, ab_name in books:
    hrefs = fetch_all_contact_hrefs(session, ab_url, verbose=False)
    print(f"Durchsuche '{ab_name}' ({len(hrefs)} Kontakte) ...")
    for i, (href, etag) in enumerate(hrefs, 1):
        print(f"\r  {i}/{len(hrefs)} ...", end="", flush=True)
        try:
            vcard = fetch_vcard(session, href)
        except Exception:
            continue
        if search.lower() not in vcard.lower():
            continue
        uid = extract_uid(vcard)
        fn_line = next((l for l in vcard.splitlines() if l.upper().startswith("FN:")), "?")
        print(f"\n--- Treffer in Adressbuch '{ab_name}' ---")
        print("href:", href)
        print("FN:  ", fn_line)
        print("UID: ", uid)
        found += 1
    print()

print(f"\n{found} Treffer für '{search}'.")
