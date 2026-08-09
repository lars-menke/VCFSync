"""
Google Contacts Sync (People API) - optionales zweites Ziel
=============================================================
iCloud bleibt die Quelle der Wahrheit. Dieses Modul schreibt dieselbe
(bereits bearbeitete) VCF zusätzlich nach Google Contacts - es liest nie
aus Google zurück und spiegelt nichts nach iCloud. Genutzt über
`icloud_contacts.py import --target google` bzw. `--target both`.

Voraussetzung (einmalig, nur vom Nutzer selbst machbar):
  1. https://console.cloud.google.com -> neues Projekt
  2. "People API" aktivieren
  3. OAuth-Consent-Screen einrichten (External reicht, dich selbst als
     Testnutzer eintragen)
  4. "Credentials" -> "Create Credentials" -> OAuth-Client-ID vom Typ
     "Desktop app" -> JSON-Datei herunterladen
  5. Pfad zur JSON-Datei in der .env-Datei eintragen:
       GOOGLE_CLIENT_SECRET=pfad/zur/heruntergeladenen_datei.json
  6. Einmalig anmelden: python icloud_contacts.py google-auth

Matching-Mechanismus: Google-Kontakte kennen keine iCloud-UIDs. Die UID
wird daher in Googles eigenem "userDefined"-Feld gespeichert
(key=vcfsync_uid) und beim nächsten Sync darüber wiedergefunden - sonst
gäbe es bei jedem Lauf Duplikate (genau das Problem, das uns beim
UID-Abgleich mit iCloud selbst lange beschäftigt hat).

Absichtlich einfach gehalten: sequenzielle Einzel-Requests (kein Batch-
API), Pause zwischen den Zugriffen. Googles Standard-Quota liegt bei nur
90 "kritischen" Lese- bzw. Schreibzugriffen pro Minute und Nutzer - bei
mehreren hundert Kontakten (v.a. beim ersten vollen Sync) wird das leicht
erreicht. Dagegen zwei Maßnahmen: die Pause zwischen Zugriffen ist so
bemessen, dass die Quota gar nicht erst ausgeschöpft wird, und ein 429
("Quota exceeded") wird automatisch mit wachsender Wartezeit wiederholt
statt den Kontakt sofort als Fehler zu zählen. Reicht das bei sehr großen
Beständen nicht, in der Google Cloud Console unter "APIs & Services" ->
"People API" -> "Quotas" ein höheres Limit anfragen, oder den Sync in
Teilen laufen lassen.
"""

import base64
import os
import re
import time
from pathlib import Path

from icloud_contacts import (_parse_label_pairs, _split_address, extract_uid,
                             load_env, split_vcards, unfold_vcard,
                             vcard_to_fields, wants_delete)

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/contacts"]
TOKEN_FILE = Path(".google_token.json")
UID_FIELD_KEY = "vcfsync_uid"

PERSON_FIELDS = ("names,nicknames,organizations,biographies,birthdays,"
                 "phoneNumbers,emailAddresses,addresses,urls,userDefined")
UPDATE_FIELDS = PERSON_FIELDS  # dieselben Felder werden geschrieben wie gelesen
# Fürs Lesen zusätzlich "photos" mitnehmen, um zu erkennen, ob ein Kontakt
# schon ein echtes (nicht generiertes) Foto hat - separat von UPDATE_FIELDS,
# da Fotos über updateContactPhoto() statt updateContact() geschrieben werden.
LIST_PERSON_FIELDS = PERSON_FIELDS + ",photos"

# Googles Standard-Quota: 90 kritische Lese-/Schreibzugriffe pro Minute und
# Nutzer. Mit dieser Pause zwischen den Zugriffen (60s / 80 statt / 90 als
# Sicherheitsmarge) wird die Quota im Normalfall gar nicht erst erreicht.
WRITE_DELAY = 60 / 80
MAX_RETRIES = 6
RETRY_BASE_DELAY = 5  # Sekunden, verdoppelt sich bei jedem weiteren Versuch

_GOOGLE_TEL_TYPE = {"mobil": "mobile", "privat": "home", "arbeit": "work"}
_GOOGLE_EMAIL_TYPE = {"privat": "home", "arbeit": "work"}
_GOOGLE_ADR_TYPE = {"privat": "home", "arbeit": "work"}


# ---------------------------------------------------------------------------
# OAuth-Anmeldung
# ---------------------------------------------------------------------------

def _client_secret_path():
    load_env()
    path = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
    if not path:
        raise RuntimeError(
            "GOOGLE_CLIENT_SECRET fehlt. Pfad zur heruntergeladenen OAuth-"
            "Client-JSON (Google Cloud Console) in der .env-Datei setzen, z.B.\n"
            "  GOOGLE_CLIENT_SECRET=client_secret_1234.json")
    p = Path(path)
    if not p.exists():
        raise RuntimeError(f"GOOGLE_CLIENT_SECRET zeigt auf eine nicht vorhandene Datei: {path}")
    return p


def get_google_credentials(interactive=True):
    """Lädt gespeicherte Google-Zugangsdaten, erneuert sie bei Bedarf, oder
    startet (interactive=True) den einmaligen Browser-Login. Google-
    Abhängigkeiten werden erst hier importiert, damit das Modul auch ohne
    installierte Google-Pakete geladen werden kann (nur wer den Google-Sync
    tatsächlich nutzt, braucht sie)."""
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import Flow
    except ImportError as e:
        raise RuntimeError(
            "Google-Abhängigkeiten fehlen. Einmalig installieren:\n"
            "  pip install -r requirements-google.txt"
        ) from e

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), GOOGLE_SCOPES)

    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
        return creds

    if not interactive:
        raise RuntimeError(
            "Keine gültige Google-Anmeldung vorhanden. Einmalig ausführen:\n"
            "  python icloud_contacts.py google-auth")

    # Bewusst kein run_local_server(): in Cloud-Umgebungen wie GitHub
    # Codespaces landet Googles Rückleitung auf "localhost" im Browser
    # des Nutzers, nicht im Container, in dem dieses Skript läuft -
    # "Verbindung verweigert". Der manuelle Copy-Paste-Weg funktioniert
    # überall (Codespaces, a-Shell, lokal) und braucht keinen laufenden
    # lokalen Webserver.
    #
    # oauthlib besteht sonst auf https für die Rückleitungs-URL - für den
    # Loopback-Sonderfall (http://localhost) ist das laut OAuth-Standard
    # ausdrücklich erlaubt; run_local_server() setzt dasselbe Flag intern,
    # das übernehmen wir hier manuell.
    os.environ.setdefault("OAUTHLIB_INSECURE_TRANSPORT", "1")

    flow = Flow.from_client_secrets_file(
        str(_client_secret_path()), scopes=GOOGLE_SCOPES, redirect_uri="http://localhost")
    auth_url, _ = flow.authorization_url(access_type="offline", prompt="consent")

    print("1. Diese Adresse im Browser öffnen und bei Google anmelden:\n")
    print(f"   {auth_url}\n")
    print("2. Google leitet danach auf eine 'localhost'-Adresse um, die im Browser")
    print("   einen Fehler zeigt ('Seite nicht erreichbar') - das ist in Codespaces")
    print("   normal und kein Problem.")
    print("3. Die komplette Adresse aus der Adressleiste kopieren (beginnt mit")
    print("   'http://localhost/?state=...') und hier einfügen:\n")
    redirected_url = input("   > ").strip()

    try:
        flow.fetch_token(authorization_response=redirected_url)
    except Exception as e:
        raise RuntimeError(
            f"Anmeldung fehlgeschlagen ({e}). War die eingefügte Adresse vollständig "
            "und aktuell (Codes gelten nur wenige Minuten)? Einfach 'google-auth' "
            "erneut ausführen."
        ) from e

    creds = flow.credentials
    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return creds


def build_google_service(interactive=True):
    try:
        from googleapiclient.discovery import build
    except ImportError as e:
        raise RuntimeError(
            "Google-Abhängigkeiten fehlen. Einmalig installieren:\n"
            "  pip install -r requirements-google.txt"
        ) from e
    creds = get_google_credentials(interactive=interactive)
    return build("people", "v1", credentials=creds)


# ---------------------------------------------------------------------------
# vCard -> Google-Person-Feldabbildung
# ---------------------------------------------------------------------------

def _google_type(label, mapping):
    """Bekannte Labels (mobil/privat/Arbeit) auf Googles Standardtypen
    abbilden; alles andere wird 1:1 als eigenes Google-Label übernommen
    (Google akzeptiert dafür beliebigen Text als 'type')."""
    lab = label.strip().lower()
    if lab in mapping:
        return mapping[lab]
    if lab in ("sonstige", "other", ""):
        return "other"
    return label.strip()


def _parse_bday(value):
    """'YYYY-MM-DD' oder '--MM-DD' (ohne Jahr) in Googles Birthday-Date-Dict."""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", value.strip())
    if m:
        y, mo, da = (int(x) for x in m.groups())
        return {"year": y, "month": mo, "day": da}
    m = re.match(r"^--(\d{2})-(\d{2})$", value.strip())
    if m:
        mo, da = (int(x) for x in m.groups())
        return {"month": mo, "day": da}
    return None


def _typed_values(items, mapping):
    """['label: wert', ...] (bereits von vcard_to_fields formatiert) in
    [{'value':..., 'type':...}, ...] für phoneNumbers/emailAddresses/urls."""
    return [{"value": v, "type": _google_type(l, mapping)}
            for l, v in _parse_label_pairs(" | ".join(items))]


def _typed_addresses(items):
    out = []
    for label, val in _parse_label_pairs(" | ".join(items)):
        street, city, plz, country = _split_address(val)
        addr = {"type": _google_type(label, _GOOGLE_ADR_TYPE)}
        if street:
            addr["streetAddress"] = street
        if city:
            addr["city"] = city
        if plz:
            addr["postalCode"] = plz
        if country:
            addr["country"] = country
        out.append(addr)
    return out


def vcard_to_google_person(vcard_text):
    """Baut den Google-People-API-Request-Body aus einer vCard. Die UID
    wird immer im userDefined-Feld mitgeschrieben - darüber findet der
    nächste Sync den Kontakt wieder."""
    d = vcard_to_fields(unfold_vcard(vcard_text))
    body = {}

    if d["Vorname"] or d["Nachname"] or d["Praefix"] or d["Suffix"]:
        body["names"] = [{
            "givenName": d["Vorname"], "familyName": d["Nachname"],
            "honorificPrefix": d["Praefix"], "honorificSuffix": d["Suffix"],
        }]
    if d["Nick"]:
        body["nicknames"] = [{"value": d["Nick"]}]

    if d["ORG"] or d["Abt"] or d["Titel"]:
        org = {"name": d["ORG"], "department": d["Abt"]}
        if d["Titel"]:
            org["title"] = d["Titel"]
        body["organizations"] = [org]

    note = d["NOTE"]
    if d["CAT"]:
        # Google People kennt keine freien Kategorien wie iCloud CATEGORIES;
        # statt die Information zu verlieren, wird sie an die Notiz angehängt.
        note = (note + "\n" if note else "") + f"Kategorien: {d['CAT']}"
    if note:
        body["biographies"] = [{"value": note, "contentType": "TEXT_PLAIN"}]

    if d["BDAY"]:
        bday = _parse_bday(d["BDAY"])
        if bday:
            body["birthdays"] = [{"date": bday}]

    if d["TEL"]:
        body["phoneNumbers"] = _typed_values(d["TEL"], _GOOGLE_TEL_TYPE)
    if d["EMAIL"]:
        body["emailAddresses"] = _typed_values(d["EMAIL"], _GOOGLE_EMAIL_TYPE)
    if d["URL"]:
        body["urls"] = _typed_values(d["URL"], {})
    if d["ADR"]:
        body["addresses"] = _typed_addresses(d["ADR"])

    body["userDefined"] = [{"key": UID_FIELD_KEY, "value": d["UID"]}]
    return body


def _photo_b64_of(vcard_text):
    return vcard_to_fields(unfold_vcard(vcard_text))["PHOTO_B64"]


# ---------------------------------------------------------------------------
# Bestehende Google-Kontakte lesen (für UID-Abgleich)
# ---------------------------------------------------------------------------

def _extract_vcfsync_uid(person):
    for ud in person.get("userDefined", []):
        if ud.get("key") == UID_FIELD_KEY:
            return ud.get("value")
    return None


def _execute_with_retry(request_factory, max_retries=MAX_RETRIES, base_delay=RETRY_BASE_DELAY):
    """Führt einen Google-API-Request aus und wiederholt ihn bei 429 (Quota
    überschritten) oder vorübergehenden 5xx-Fehlern mit wachsender Wartezeit
    (5s, 10s, 20s, ...). request_factory muss bei jedem Aufruf ein FRISCHES
    Request-Objekt liefern (google-api-python-client-Requests sind nur einmal
    ausführbar) - deshalb ein Callable statt eines fertigen Requests."""
    from googleapiclient.errors import HttpError
    delay = base_delay
    for attempt in range(max_retries):
        try:
            return request_factory().execute()
        except HttpError as e:
            status = e.resp.status if e.resp is not None else None
            retryable = status == 429 or (status is not None and 500 <= status < 600)
            if not retryable or attempt == max_retries - 1:
                raise
            time.sleep(delay)
            delay = min(delay * 2, 60)


def fetch_existing_google_uids(service):
    """Gibt {icloud_uid: person} über alle eigenen Google-Kontakte zurück,
    die schon einmal von hier aus angelegt wurden (haben das userDefined-
    Feld vcfsync_uid). Kontakte ohne dieses Feld (z.B. von Hand angelegte)
    werden ignoriert - für die gibt es keinen UID-Bezug. Der volle
    Person-Datensatz (nicht nur resourceName/etag) wird behalten, damit
    push_contacts_to_google() prüfen kann, ob sich inhaltlich überhaupt
    etwas geändert hat, statt bei jedem Lauf alle Kontakte neu zu schreiben."""
    uid_map = {}
    page_token = None
    while True:
        resp = _execute_with_retry(lambda pt=page_token: service.people().connections().list(
            resourceName="people/me", pageSize=200,
            personFields=LIST_PERSON_FIELDS, pageToken=pt))
        for person in resp.get("connections", []):
            uid = _extract_vcfsync_uid(person)
            if uid:
                uid_map[uid] = person
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return uid_map


# ---------------------------------------------------------------------------
# Änderungserkennung: nur wirklich abweichende Kontakte schreiben
# ---------------------------------------------------------------------------

def _comparable_from_person(person):
    """Baut aus einem von Google gelesenen Person-Objekt dieselbe Teilmenge
    an Feldern, die vcard_to_google_person() beim Schreiben erzeugt (ohne
    Googles Metadaten wie formattedType/metadata/source). So lässt sich der
    aktuelle Google-Stand direkt mit dem neuen Body vergleichen."""
    out = {}

    names = person.get("names")
    if names:
        n = names[0]
        val = {"givenName": n.get("givenName", ""), "familyName": n.get("familyName", ""),
               "honorificPrefix": n.get("honorificPrefix", ""), "honorificSuffix": n.get("honorificSuffix", "")}
        if any(val.values()):
            out["names"] = [val]

    nicknames = person.get("nicknames")
    if nicknames and nicknames[0].get("value"):
        out["nicknames"] = [{"value": nicknames[0]["value"]}]

    orgs = person.get("organizations")
    if orgs:
        o = orgs[0]
        org = {"name": o.get("name", ""), "department": o.get("department", "")}
        if o.get("title"):
            org["title"] = o["title"]
        if org["name"] or org["department"] or org.get("title"):
            out["organizations"] = [org]

    bios = person.get("biographies")
    if bios and bios[0].get("value"):
        out["biographies"] = [{"value": bios[0]["value"],
                               "contentType": bios[0].get("contentType", "TEXT_PLAIN")}]

    for b in person.get("birthdays", []):
        d = b.get("date")
        if d:
            out["birthdays"] = [{"date": {k: d[k] for k in ("year", "month", "day") if k in d}}]
            break

    if person.get("phoneNumbers"):
        out["phoneNumbers"] = [{"value": t.get("value", ""), "type": t.get("type", "other")}
                               for t in person["phoneNumbers"]]
    if person.get("emailAddresses"):
        out["emailAddresses"] = [{"value": e.get("value", ""), "type": e.get("type", "other")}
                                 for e in person["emailAddresses"]]
    if person.get("urls"):
        out["urls"] = [{"value": u.get("value", ""), "type": u.get("type", "other")}
                       for u in person["urls"]]
    if person.get("addresses"):
        out_addrs = []
        for a in person["addresses"]:
            addr = {"type": a.get("type", "other")}
            for key in ("streetAddress", "city", "postalCode", "country"):
                if a.get(key):
                    addr[key] = a[key]
            out_addrs.append(addr)
        out["addresses"] = out_addrs

    out["userDefined"] = [ud for ud in person.get("userDefined", []) if ud.get("key") == UID_FIELD_KEY]
    return out


# ---------------------------------------------------------------------------
# Google-Person -> dasselbe Feld-Dict-Format wie vcard_to_fields() (iCloud) -
# für contact_check.py, damit dieselbe Prüfung für beide Quellen gilt.
# ---------------------------------------------------------------------------

# Wie _GOOGLE_TEL_TYPE etc. umgekehrt, aber mit iCloud-Groß-/Kleinschreibung
# ("Arbeit", nicht "arbeit") - dieselbe Konvention wie bei _TEL_LABELS in
# icloud_contacts.py, damit ein Label unabhängig von der Quelle gleich aussieht.
_GOOGLE_TEL_LABEL = {"mobile": "mobil", "home": "privat", "work": "Arbeit"}
_GOOGLE_EMAIL_LABEL = {"home": "privat", "work": "Arbeit"}
_GOOGLE_ADR_LABEL = {"home": "privat", "work": "Arbeit"}


def _google_label(typ, reverse_map):
    """Kehrt _google_type() um: Googles Typ zurück in unser Label (mobil/
    privat/Arbeit), sonst wie beim Schreiben 1:1 übernommen bzw. 'Sonstige'."""
    t = (typ or "").strip().lower()
    if t in reverse_map:
        return reverse_map[t]
    if t in ("", "other"):
        return "Sonstige"
    return typ.strip()


def _bday_from_google(person):
    for b in person.get("birthdays", []):
        d = b.get("date")
        if not d:
            continue
        if "year" in d:
            return f"{d['year']:04d}-{d.get('month', 1):02d}-{d.get('day', 1):02d}"
        return f"--{d.get('month', 1):02d}-{d.get('day', 1):02d}"
    return ""


def person_to_fields(person):
    """Google-Person-Objekt (aus people.connections.list) in dasselbe Feld-
    Dict-Format wie vcard_to_fields() - damit contact_check.py unverändert
    auf beide Quellen anwendbar ist."""
    n = (person.get("names") or [{}])[0]
    o = (person.get("organizations") or [{}])[0]
    bio = (person.get("biographies") or [{}])[0]

    tel = [f"{_google_label(t.get('type'), _GOOGLE_TEL_LABEL)}: {t.get('value', '').strip()}"
           for t in person.get("phoneNumbers", []) if (t.get("value") or "").strip()]
    email = [f"{_google_label(e.get('type'), _GOOGLE_EMAIL_LABEL)}: {e.get('value', '').strip()}"
             for e in person.get("emailAddresses", []) if (e.get("value") or "").strip()]
    url = [f"{_google_label(u.get('type'), {})}: {u.get('value', '').strip()}"
           for u in person.get("urls", []) if (u.get("value") or "").strip()]
    adr = []
    for a in person.get("addresses", []):
        pieces = [a.get("streetAddress", ""), a.get("postalCode", ""), a.get("city", ""),
                 "", a.get("country", "")]
        flat = " ".join(x.strip() for x in pieces if x and x.strip())
        if flat:
            adr.append(f"{_google_label(a.get('type'), _GOOGLE_ADR_LABEL)}: {flat}")

    return {
        "UID": _extract_vcfsync_uid(person) or person.get("resourceName", ""),
        "FN": " ".join(x for x in (n.get("givenName", ""), n.get("familyName", "")) if x).strip(),
        "Nachname": n.get("familyName", "") or "", "Vorname": n.get("givenName", "") or "",
        "Zusatz": "", "Praefix": n.get("honorificPrefix", "") or "",
        "Suffix": n.get("honorificSuffix", "") or "",
        "Nick": (person.get("nicknames") or [{}])[0].get("value", "") or "",
        "ORG": o.get("name", "") or "", "Abt": o.get("department", "") or "",
        "Titel": o.get("title", "") or "",
        "BDAY": _bday_from_google(person),
        "NOTE": bio.get("value", "") or "", "CAT": "",
        "TEL": tel, "EMAIL": email, "ADR": adr, "URL": url, "PHOTO_B64": "",
    }


def fetch_contacts_fields(progress=None, interactive=True):
    """Alle eigenen Google-Kontakte lesen, roh in Feld-Dicts umgewandelt - für
    die Plausibilitätsprüfung (contact_check.py). Baut sich seine eigene
    Verbindung auf, unabhängig vom iCloud-/vCard-Sync in diesem Modul.

    interactive=False (z.B. aus der Web-Oberfläche): wirft RuntimeError statt
    interaktiv nach einer Anmeldung zu fragen, wenn noch kein Token vorliegt.

    Gibt (felder_liste, warnungen) zurück - warnungen ist hier bewusst immer
    leer (ein einzelner kaputter Kontakt kann die Liste nicht "verlieren",
    anders als beim iCloud-Einzelabruf je Kontakt); Verbindungsfehler werden
    stattdessen als Exception nach oben gereicht.
    """
    service = build_google_service(interactive=interactive)
    rows = []
    page_token = None
    while True:
        resp = _execute_with_retry(lambda pt=page_token: service.people().connections().list(
            resourceName="people/me", pageSize=200, personFields=PERSON_FIELDS, pageToken=pt))
        for person in resp.get("connections", []):
            rows.append(person_to_fields(person))
            if progress:
                progress(len(rows), len(rows))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return rows, []


def _normalize_for_compare(body):
    """Wandelt einen Person-Body in eine ordnungsunabhängige, vergleichbare
    Form um (Listen von dicts werden sortiert), damit z.B. eine von Google
    anders sortierte Telefonliste keine unnötige Aktualisierung auslöst."""
    def norm(value):
        if isinstance(value, dict):
            return tuple(sorted((k, norm(v)) for k, v in value.items()))
        if isinstance(value, list):
            return tuple(sorted(norm(v) for v in value))
        return value
    return norm({k: v for k, v in body.items() if k != "etag"})


def _person_unchanged(new_body, existing_person):
    return _normalize_for_compare(new_body) == _normalize_for_compare(_comparable_from_person(existing_person))


def _has_real_photo(person):
    """True, wenn Google für den Kontakt schon ein hochgeladenes (nicht
    automatisch generiertes) Foto hat."""
    return any(not p.get("default") for p in person.get("photos", []))


# ---------------------------------------------------------------------------
# Schreiben
# ---------------------------------------------------------------------------

def _upload_photo(service, resource_name, vcard_text):
    b64 = _photo_b64_of(vcard_text)
    if not b64:
        return
    try:
        _execute_with_retry(lambda: service.people().updateContactPhoto(
            resourceName=resource_name, body={"photoBytes": b64}))
    except Exception:
        pass  # Foto ist nice-to-have - soll den restlichen Sync nicht scheitern lassen


def push_contacts_to_google(service, vcards, dry_run=False, progress=None):
    """Kernlogik des Google-Syncs, im selben Stil wie import_vcards() für
    iCloud: legt neue Kontakte an, aktualisiert vorhandene (per UID-Abgleich
    über das userDefined-Feld) und löscht zum Löschen markierte. Vorhandene
    Kontakte, die inhaltlich (und im Foto) unverändert sind, werden nicht
    geschrieben, sondern nur gezählt (unchanged) - erspart bei großen
    Beständen fast alle Schreibzugriffe, wenn sich nur wenige Kontakte
    tatsächlich geändert haben. Gibt ein Dict zurück: updated, created,
    deleted, errors, unchanged, total, log."""
    existing = fetch_existing_google_uids(service)

    updated = created = deleted = errors = unchanged = 0
    log = []
    total = len(vcards)

    def emit(i, message):
        log.append(message)
        if progress:
            progress(i, total, message)

    for i, vcard in enumerate(vcards, 1):
        uid = extract_uid(vcard)
        if not uid:
            emit(i, f"[{i}] KEIN UID - übersprungen (Google)")
            errors += 1
            continue

        if wants_delete(vcard):
            if uid not in existing:
                emit(i, f"[{i}] LÖSCHEN übersprungen (nicht in Google): {uid[:20]}...")
                continue
            resource_name = existing[uid]["resourceName"]
            if dry_run:
                emit(i, f"[{i}] (dry-run) WÜRDE LÖSCHEN (Google): {uid[:20]}...")
                deleted += 1
                continue
            try:
                _execute_with_retry(lambda: service.people().deleteContact(resourceName=resource_name))
                emit(i, f"[{i}] GELÖSCHT (Google): {uid[:20]}...")
                deleted += 1
            except Exception as e:  # noqa: BLE001
                emit(i, f"[{i}] FEHLER Löschen (Google) {e}: {uid[:20]}")
                errors += 1
            if not dry_run:
                time.sleep(WRITE_DELAY)
            continue

        body = vcard_to_google_person(vcard)

        if uid in existing:
            existing_person = existing[uid]
            resource_name = existing_person["resourceName"]
            etag = existing_person.get("etag")

            fields_unchanged = _person_unchanged(body, existing_person)
            photo_b64 = _photo_b64_of(vcard)
            needs_photo_upload = bool(photo_b64) and not _has_real_photo(existing_person)

            if fields_unchanged and not needs_photo_upload:
                emit(i, f"[{i}] UNVERÄNDERT (Google, übersprungen): {uid[:20]}...")
                unchanged += 1
                continue

            if dry_run:
                emit(i, f"[{i}] (dry-run) WÜRDE AKTUALISIEREN (Google): {uid[:20]}...")
                updated += 1
                continue
            try:
                if not fields_unchanged:
                    body["etag"] = etag
                    person = _execute_with_retry(lambda: service.people().updateContact(
                        resourceName=resource_name, updatePersonFields=UPDATE_FIELDS,
                        body=body))
                    resource_name = person.get("resourceName", resource_name)
                if needs_photo_upload:
                    _upload_photo(service, resource_name, vcard)
                emit(i, f"[{i}] AKTUALISIERT (Google): {uid[:20]}...")
                updated += 1
            except Exception as e:  # noqa: BLE001
                emit(i, f"[{i}] FEHLER Update (Google) {e}: {uid[:20]}")
                errors += 1
        else:
            if dry_run:
                emit(i, f"[{i}] (dry-run) WÜRDE NEU ANLEGEN (Google): {uid[:20]}...")
                created += 1
                continue
            try:
                person = _execute_with_retry(lambda: service.people().createContact(body=body))
                _upload_photo(service, person["resourceName"], vcard)
                emit(i, f"[{i}] NEU ANGELEGT (Google): {uid[:20]}...")
                created += 1
            except Exception as e:  # noqa: BLE001
                emit(i, f"[{i}] FEHLER Neuanlage (Google) {e}: {uid[:20]}")
                errors += 1

        if not dry_run:
            time.sleep(WRITE_DELAY)  # unter Googles 90/min-Quota bleiben

    return {"updated": updated, "created": created, "deleted": deleted,
            "errors": errors, "unchanged": unchanged, "total": total, "log": log}
