# iCloud Kontakte Sync

Exportiert und importiert iCloud-Kontakte via CardDAV. Der Export lässt sich
manuell per GitHub Actions Workflow im Browser starten und stellt die
exportierte VCF-Datei als downloadbares Artifact bereit. Der Import spielt eine
bearbeitete VCF zurück nach iCloud — vorhandene Kontakte werden anhand der UID
aktualisiert (keine Duplikate), neue Kontakte werden angelegt.

Für den bequemen Alltag gibt es zusätzlich eine **Web-Oberfläche** mit Buttons
(siehe unten), die Export und Import ohne Kommandozeile erledigt.

## Web-Oberfläche (empfohlen)

Statt die Befehle einzeln einzutippen, lässt sich alles per Knopfdruck über
eine kleine lokale Web-App bedienen:

```bash
pip install -r requirements.txt
python icloud_web.py
```

Dann im Browser `http://127.0.0.1:5000` öffnen. Die Oberfläche bietet:

- **Export** — ein Klick holt alle Kontakte inkl. Fotos aus iCloud, mit
  Fortschrittsbalken; danach als **VCF** oder direkt als **Excel** herunterladen.
- **Import** — bearbeitete Datei (VCF oder die Excel-Liste) hochladen, erst
  einen **Testlauf** machen, dann mit „Wirklich importieren“ übernehmen. Der
  Testlauf zeigt, was neu angelegt / gelöscht würde und **welche Kontakte sich
  inhaltlich wirklich ändern** (mit Namen und betroffenen Feldern) — nicht nur,
  wie viele überschrieben werden.
- **Löschen** — in der Excel-Liste ein `x` in die Spalte `Löschen` setzen; der
  Import entfernt diese Kontakte dann aus iCloud. Der Testlauf zeigt vorher
  deutlich an, wie viele gelöscht würden.

Die App nutzt dieselbe Kernlogik wie das CLI und dieselben Zugangsdaten aus der
`.env`-Datei bzw. den Umgebungsvariablen.

> **Wo läuft das?**
> - **GitHub Codespaces** (siehe unten): nach `python icloud_web.py` blendet
>   VS Code eine Meldung „Im Browser öffnen“ für Port 5000 ein — anklicken.
> - **a-Shell auf dem iPhone**: App starten, dann in Safari auf demselben
>   Gerät `http://127.0.0.1:5000` öffnen.
> - **Lokal am PC**: einfach den Browser auf `http://127.0.0.1:5000`.
>
> Die App hat bewusst **keine Anmeldung** und ist nur für die lokale,
> persönliche Nutzung gedacht. Den Port nicht öffentlich freigeben.

## Einrichtung

1. **Repository als privat anlegen** — Kontaktdaten sind sensibel.
2. **GitHub Secrets setzen:**
   - `Settings` → `Secrets and variables` → `Actions` → `New repository secret`
   - `ICLOUD_USER` = Apple-ID E-Mail
   - `ICLOUD_PASS` = App-spezifisches Passwort
     (erstellen auf [appleid.apple.com](https://appleid.apple.com) →
     Anmeldung und Sicherheit → App-spezifische Passwörter)

   > Das iCloud-Hauptpasswort funktioniert **nicht** — es muss ein
   > app-spezifisches Passwort sein.

## Export ausführen

1. GitHub → Tab `Actions` → `iCloud Kontakte Export`
2. Klick auf `Run workflow` → `Run workflow`
3. Nach ~2–3 Minuten: Workflow-Run anklicken → unter `Artifacts` die VCF
   herunterladen
4. Das Artifact ist **7 Tage** verfügbar, danach wird es automatisch gelöscht
   (aus Datenschutzgründen)

## Import ausführen

> **Wichtig:** Das Eingabefeld von GitHub Actions ist **einzeilig** — beim
> Einfügen einer mehrzeiligen VCF wirft der Browser alle Zeilenumbrüche weg und
> es wird nichts importiert. Deshalb wird die VCF **Base64-kodiert** (eine
> einzige Zeile) eingefügt; der Workflow dekodiert sie automatisch.

1. VCF Base64-kodieren (eine Zeile):
   - Linux: `base64 -w0 meine_kontakte.vcf`
   - macOS: `base64 -i meine_kontakte.vcf`
2. Den Base64-String kopieren
3. GitHub → Tab `Actions` → `iCloud Kontakte Import`
4. `Run workflow` → Base64-String in das Feld einfügen → `Run workflow`
5. Bei großen Dateien (Eingabe-Limit ~65.000 Zeichen) oder vielen Fotos: Skript
   **lokal** ausführen (siehe unten) — für den Vollbestand der empfohlene Weg.

   > Hinweis: roher VCF-Text wird weiterhin akzeptiert (z. B. eine einzelne,
   > einzeilige Karte). Sobald Zeilenumbrüche im Spiel sind, ist Base64 über die
   > Weboberfläche der zuverlässige Weg.

## Nutzung ohne lokale Python-Installation (GitHub Codespaces)

Falls kein Python zur Verfügung steht (z. B. Firmen-PC ohne Adminrechte) oder
a-Shell auf dem iPhone zu umständlich ist: Repository auf GitHub öffnen →
`Code` → Tab `Codespaces` → `Create codespace on main`. Startet eine
fertig eingerichtete Cloud-Umgebung mit Python und allen Abhängigkeiten
(`.devcontainer/devcontainer.json`) — Terminal öffnet sich automatisch im
Browser, alle Befehle aus diesem README funktionieren dort 1:1. Dateien
lassen sich per Rechtsklick im Datei-Explorer hoch-/runterladen.

## Lokale Nutzung

```bash
pip install -r requirements.txt

# Export
python icloud_contacts.py export --output meine_kontakte.vcf

# Import
python icloud_contacts.py import --input bearbeitete_kontakte.vcf

# Import simulieren (nichts wird geschrieben)
python icloud_contacts.py import --input bearbeitete_kontakte.vcf --dry-run

# Kontakte löschen (per UID oder anhand einer VCF mit den zu löschenden Karten)
python icloud_contacts.py delete --uid <UID> [--uid <UID> ...]
python icloud_contacts.py delete --input zu_loeschen.vcf
python icloud_contacts.py delete --input zu_loeschen.vcf --dry-run
```

> Hinweis: `import` kann Kontakte nur anlegen/aktualisieren, nicht entfernen.
> Zum vollständigen Löschen von Kontakten dient `delete` (per HTTP DELETE).
> Immer erst mit `--dry-run` prüfen — Löschen ist nicht umkehrbar.

## Standard-Workflow: Kontakte in Excel bearbeiten

Für größere Aufräumaktionen (Telefonlabels vereinheitlichen, Adressen
korrigieren, Dubletten bereinigen ...) lohnt sich der Umweg über Excel statt
die VCF von Hand zu editieren. Der komplette Zyklus:

```bash
# 1. Aktuellen Bestand aus iCloud holen
python icloud_contacts.py export --output export.vcf

# 2. In eine Excel-Liste wandeln
python icloud_contacts.py to-excel --input export.vcf --output kontakte.xlsx

# 3. kontakte.xlsx in Excel bearbeiten (siehe Spaltenformat unten)

# 4. Bearbeitete Liste zurück in eine VCF wandeln
python icloud_contacts.py from-excel --input kontakte.xlsx --output bearbeitet.vcf

# 5. Prüfen, was sich ändern würde (kein Duplikat-Risiko eingehen!)
python icloud_contacts.py import --input bearbeitet.vcf --dry-run

# 6. Wirklich importieren
python icloud_contacts.py import --input bearbeitet.vcf
```

**Spaltenformat in der Excel-Liste:**

- Einfache Felder (Anzeigename, Nachname, Vorname, Organisation, Titel,
  Geburtstag, Notiz, Kategorien ...): normaler Zellwert.
- Telefone/E-Mails/Adressen/URLs: mehrere Einträge in einer Zelle, getrennt
  durch ` | `, jeweils im Format `Label: Wert`, z. B.
  `mobil: +49171234567 | Arbeit: +494412345`. Als Label gelten `mobil`,
  `privat` und `Arbeit` mit fester Bedeutung — alles andere (z. B.
  `Sonstige` oder ein eigener Text) wird als eigenes Apple-Label übernommen.
- `UID`-Spalte **nicht verändern** (ist standardmäßig ausgeblendet) — sie
  entscheidet beim Import, ob ein bestehender Kontakt aktualisiert oder ein
  neuer angelegt wird. Eine leere UID-Zelle erzeugt einen neuen Kontakt.
- **Spalte `Löschen`**: ein `x` (oder `ja`) in dieser Zelle löscht den Kontakt
  beim Import aus iCloud. Nur genau markierte Kontakte werden entfernt — eine
  Zeile einfach zu löschen genügt **nicht** (dann bleibt der Kontakt in iCloud).
  Der Testlauf zeigt vorab, was gelöscht würde. Alternativ zum gezielten
  Entfernen weiterhin `delete --uid <UID>` möglich.
- Foto liegt Base64-kodiert in den Spalten `Foto_Base64_1`, `_2`, ... (auf
  mehrere Spalten aufgeteilt wegen Excels Zellenlimit von 32.767 Zeichen) —
  normalerweise nicht von Hand bearbeiten.
- Gruppen-Karten (z. B. "Familie", "Archiv") tauchen in der Excel-Liste
  nicht auf, da sie auf dem iPhone ohnehin nicht einsehbar/bearbeitbar sind.

> Adressen werden beim Zurückwandeln nur bestmöglich in Straße/PLZ/Ort
> zerlegt (Excel enthält nur Fließtext pro Adresse) — die PLZ wird per
> Mustererkennung herausgezogen, der Rest landet in der Straßen-Komponente.
> Das reicht für iCloud/iPhone völlig aus.

> Vor dem Import werden die vorhandenen iCloud-Kontakte gelesen und
> feldweise mit den neuen Daten verglichen — nur Kontakte, die neu sind,
> gelöscht werden sollen oder sich inhaltlich tatsächlich unterscheiden,
> lösen einen Schreibzugriff aus. Inhaltlich unveränderte Kontakte werden
> übersprungen und als „bereits aktuell" gezählt statt unnötig neu
> geschrieben zu werden.

Zugangsdaten werden beim ersten Start abgefragt und optional in einer
`.env`-Datei gespeichert, oder als Umgebungsvariablen gesetzt:

```
ICLOUD_USER=lars@icloud.com
ICLOUD_PASS=xxxx-xxxx-xxxx-xxxx
```

## Google Contacts Sync (optional)

Dieselbe (bearbeitete) VCF lässt sich zusätzlich zu iCloud auch nach Google
Contacts schreiben. **iCloud bleibt dabei die Quelle der Wahrheit** — der
Sync geht nur in eine Richtung (iCloud/Excel → Google), es wird nie aus
Google gelesen und zurückgespiegelt. Google-Kontakte, die von hier aus
angelegt wurden, werden über ein eigenes Kennzeichen wiedererkannt (Googles
`userDefined`-Feld `vcfsync_uid`) — dadurch entstehen bei wiederholtem Sync
keine Duplikate, genau wie bei iCloud per UID.

### Einmalige Einrichtung (nur du kannst das machen)

1. [console.cloud.google.com](https://console.cloud.google.com) → neues
   Projekt anlegen
2. „APIs & Services" → „Library" → **„People API"** suchen → aktivieren
3. „APIs & Services" → „OAuth consent screen" → Typ **„External"** →
   App-Name/E-Mail ausfüllen, dich selbst als Testnutzer eintragen
4. „Credentials" → „Create Credentials" → **„OAuth client ID"** →
   Anwendungstyp **„Desktop app"** → JSON-Datei herunterladen
5. Google-Abhängigkeiten installieren:
   ```bash
   pip install -r requirements-google.txt
   ```
6. Pfad zur heruntergeladenen JSON-Datei in der `.env` eintragen:
   ```
   GOOGLE_CLIENT_SECRET=client_secret_1234.json
   ```
7. Einmalig anmelden:
   ```bash
   python icloud_contacts.py google-auth
   ```
   Gibt einen Link aus, der im Browser geöffnet und dort bei Google bestätigt
   werden muss. Danach leitet Google auf eine `localhost`-Adresse um, die im
   Browser einen Fehler zeigt („Seite nicht erreichbar") — das ist **normal**
   in Codespaces und kein Problem: einfach die komplette Adresse aus der
   Adressleiste kopieren und im Terminal einfügen, wenn danach gefragt wird.
   Speichert die Anmeldung in `.google_token.json` (lokal, nie committet).
   Läuft der Token irgendwann ab, reicht ein erneutes `google-auth`.

### Nutzung

**Kommandozeile** — `--target` steuert, wohin importiert wird:

```bash
python icloud_contacts.py import --input bearbeitet.vcf --target icloud   # Standard
python icloud_contacts.py import --input bearbeitet.vcf --target google
python icloud_contacts.py import --input bearbeitet.vcf --target both
python icloud_contacts.py import --input bearbeitet.vcf --target both --dry-run
```

**Web-Oberfläche** — im Import-Bereich erscheint nach dem Hochladen eine
Ziel-Auswahl (iCloud / Google / Beide). Ist Google noch nicht verbunden,
zeigt die Seite einen Hinweis auf `google-auth` an; die Optionen sind bis
dahin deaktiviert. Aus Sicherheitsgründen läuft die OAuth-Anmeldung selbst
**nur über die Kommandozeile** (nicht per Klick im Browser) — die Web-App
liest lediglich den bereits gespeicherten Token.

### Einschränkungen

- Google People API kennt keine freien `CATEGORIES` wie iCloud — Kategorien
  werden an die Notiz angehängt statt verloren zu gehen.
- Adressen werden wie beim Excel-Import bestmöglich in Straße/Ort/PLZ/Land
  zerlegt (siehe oben).
- Kein Batch-API, einzelne Requests mit Pause dazwischen (siehe
  „Rate-Limiting" unten).

### Rate-Limiting (Googles 90/Minute-Quota)

Googles Standard-Kontingent für die People API liegt bei nur 90 „kritischen"
Lese- bzw. Schreibzugriffen pro Minute und Nutzer — bei mehreren hundert
Kontakten (v.a. beim allerersten vollen Sync, bevor die Änderungserkennung
greift) ist das schnell erreicht. Dagegen zwei Maßnahmen:

- Zwischen den Schreibzugriffen wird automatisch pausiert (Tempo knapp unter
  der Quota), statt sie so schnell wie möglich abzufeuern.
- Antwortet Google trotzdem mit „429 Quota exceeded" (oder einem
  vorübergehenden 5xx-Fehler), wird der Zugriff automatisch mit wachsender
  Wartezeit wiederholt (5s, 10s, 20s, ... bis zu 6 Versuche), statt den
  Kontakt sofort als Fehler zu zählen.

Bei sehr großen Beständen kann das trotzdem an die Quota stoßen — in der
Google Cloud Console unter „APIs & Services" → „People API" → „Quotas" ggf.
ein höheres Limit anfragen, oder den Sync in mehreren Durchgängen laufen
lassen (dank Änderungserkennung wiederholt ein erneuter Lauf nur die noch
fehlenden/fehlgeschlagenen Kontakte).

### Änderungserkennung (nur wirklich geänderte Kontakte werden geschrieben)

Genau wie beim iCloud-Import (siehe oben) werden vor jedem Google-Sync die
vorhandenen Google-Kontakte gelesen und mit den neuen Daten verglichen. Nur
Kontakte, die neu sind, gelöscht werden sollen oder sich inhaltlich (oder im
Foto) tatsächlich unterscheiden, lösen einen Schreibzugriff
(`createContact`/`updateContact`/`deleteContact`) aus — inhaltlich
unveränderte Kontakte werden übersprungen und als „unverändert" gezählt,
nicht als „aktualisiert". Das spart bei großen Beständen fast alle
Schreibzugriffe, wenn seit dem letzten Sync nur wenige Kontakte geändert
wurden, und schont Googles Schreib-Quota entsprechend.

## Datenschutzhinweis

- Repository unbedingt **privat** halten
- Zugangsdaten niemals direkt in Code oder Workflow-Dateien schreiben —
  ausschließlich über GitHub Secrets bzw. die lokale `.env`
- Exportierte Artifacts enthalten alle Kontaktdaten inkl. Fotos
- Artifacts werden nach 7 Tagen automatisch gelöscht
- Die Dateien `.env` und `*.vcf` sind via `.gitignore` vom Commit ausgeschlossen
- Die Google-OAuth-Client-JSON (`client_secret*.json`) und der gespeicherte
  Google-Token (`.google_token.json`) sind ebenfalls via `.gitignore`
  ausgeschlossen — niemals committen
