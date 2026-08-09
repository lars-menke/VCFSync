"""Plausibilitätsprüfung für Kontakte - findet bekannte Inkonsistenzen, ohne
irgendetwas zu verändern.

Arbeitet ausschließlich auf dem Feld-Dict, das vcard_to_fields() (iCloud) bzw.
der entsprechende Google-Mapper erzeugen - dieselbe Prüfung gilt also für
beide Quellen, ohne Sonderfälle. Reine Funktionen, kein Netzwerk-/Dateizugriff.

Geprüft wird:
  - Telefonnummer sieht nach einer deutschen Handynummer aus (015x/016x/017x),
    ist aber nicht als "mobil" gelabelt
  - Telefon-Label uneindeutig ("Sonstige")
  - Telefonnummer/E-Mail strukturell unplausibel (keine Ziffern/kein @ usw.)
  - Telefonnummer/E-Mail kommt am selben Kontakt doppelt vor
  - Adresse ohne erkennbare 5-stellige PLZ
  - Anzeigename vorhanden, aber Vor-/Nachname beide leer
"""
import re

_PHONE_STRIP_RE = re.compile(r"[\s\-/().]")
_EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")

STATUS_LABELS_DE = {"ok": "OK", "warn": "Hinweis", "error": "Fehler"}
# Reihenfolge = Dringlichkeit; wird von overall_status() zum Vergleichen genutzt.
_SEVERITY_RANK = {"ok": 0, "warn": 1, "error": 2}


def _normalize_phone(value):
    """Nur Ziffern und ein evtl. führendes '+' - Trennzeichen raus."""
    return _PHONE_STRIP_RE.sub("", value or "")


def _looks_like_mobile(value):
    """Deutsche Mobilfunk-Vorwahl (015x/016x/017x), auch mit Landesvorwahl
    (+49.../0049...) statt der führenden 0."""
    v = _normalize_phone(value)
    if v.startswith("+49"):
        v = "0" + v[3:]
    elif v.startswith("0049"):
        v = "0" + v[4:]
    return bool(re.match(r"^0(15|16|17)\d", v))


def _looks_like_phone(value):
    """Grobe Plausibilität: nur Ziffern (+ optional führendes +), 5-15 Stellen.
    Keine Formatprüfung im Detail - nur ein Sicherheitsnetz gegen offensichtlich
    kaputte Werte (Buchstaben, Textreste, leere Nummern)."""
    v = _normalize_phone(value)
    if not v:
        return False
    digits = v[1:] if v.startswith("+") else v
    return digits.isdigit() and 5 <= len(digits) <= 15


def _looks_like_email(value):
    return bool(_EMAIL_RE.match((value or "").strip()))


def check_contact(fields):
    """Prüft einen Kontakt (Feld-Dict) gegen die bekannten Inkonsistenzen.

    Gibt eine Liste von Findings zurück, jedes als
    {"severity": "warn"|"error", "field": "<Excel-Spaltenname>", "message": "..."}.
    Leere Liste heißt: keine Auffälligkeit gefunden.
    """
    findings = []

    seen_numbers = set()
    for entry in fields.get("TEL", []):
        label, _, value = entry.partition(":")
        label, value = label.strip(), value.strip()
        if not value:
            continue
        norm = _normalize_phone(value)
        if norm in seen_numbers:
            findings.append({"severity": "warn", "field": "Telefone",
                             "message": f"Nummer „{value}“ kommt doppelt vor."})
        seen_numbers.add(norm)

        if not _looks_like_phone(value):
            findings.append({"severity": "error", "field": "Telefone",
                             "message": f"„{value}“ sieht nicht wie eine gültige Telefonnummer aus."})
            continue  # Label einer erkennbar kaputten Nummer bringt nichts

        if label.lower() == "sonstige":
            findings.append({"severity": "warn", "field": "Telefone",
                             "message": f"Label von „{value}“ ist uneindeutig („Sonstige“)."})
        elif _looks_like_mobile(value) and label.lower() != "mobil":
            findings.append({"severity": "error", "field": "Telefone",
                             "message": f"„{value}“ sieht nach einer Handynummer aus, ist aber "
                                        f"als „{label}“ gelabelt."})

    seen_mails = set()
    for entry in fields.get("EMAIL", []):
        _label, _, value = entry.partition(":")
        value = value.strip()
        if not value:
            continue
        key = value.lower()
        if key in seen_mails:
            findings.append({"severity": "warn", "field": "E-Mails",
                             "message": f"E-Mail „{value}“ kommt doppelt vor."})
        seen_mails.add(key)
        if not _looks_like_email(value):
            findings.append({"severity": "error", "field": "E-Mails",
                             "message": f"„{value}“ sieht nicht wie eine gültige E-Mail-Adresse aus."})

    for entry in fields.get("ADR", []):
        _label, _, value = entry.partition(":")
        value = value.strip()
        if value and not re.search(r"\b\d{5}\b", value):
            findings.append({"severity": "warn", "field": "Adressen",
                             "message": f"Adresse „{value}“ enthält keine erkennbare 5-stellige PLZ."})

    if (fields.get("FN") or "").strip() and not (
            (fields.get("Nachname") or "").strip() or (fields.get("Vorname") or "").strip()):
        findings.append({"severity": "warn", "field": "Name",
                         "message": "Anzeigename vorhanden, aber Vor- und Nachname beide leer."})

    return findings


def overall_status(findings):
    """"error", wenn mindestens ein Finding rot ist, sonst "warn" bei
    mindestens einem gelben, sonst "ok"."""
    if not findings:
        return "ok"
    return "error" if any(f["severity"] == "error" for f in findings) else "warn"


def notes_text(findings):
    """Alle Fundstellen als ein Text für die Hinweise-Spalte."""
    return " | ".join(f["message"] for f in findings)
