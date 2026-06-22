# iCloud Kontakte Sync

Exportiert und importiert iCloud-Kontakte via CardDAV. Der Export lässt sich
manuell per GitHub Actions Workflow im Browser starten und stellt die
exportierte VCF-Datei als downloadbares Artifact bereit. Der Import spielt eine
bearbeitete VCF zurück nach iCloud — vorhandene Kontakte werden anhand der UID
aktualisiert (keine Duplikate), neue Kontakte werden angelegt.

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

## Lokale Nutzung

```bash
pip install requests

# Export
python icloud_contacts.py export --output meine_kontakte.vcf

# Import
python icloud_contacts.py import --input bearbeitete_kontakte.vcf

# Import simulieren (nichts wird geschrieben)
python icloud_contacts.py import --input bearbeitete_kontakte.vcf --dry-run
```

Zugangsdaten werden beim ersten Start abgefragt und optional in einer
`.env`-Datei gespeichert, oder als Umgebungsvariablen gesetzt:

```
ICLOUD_USER=lars@icloud.com
ICLOUD_PASS=xxxx-xxxx-xxxx-xxxx
```

## Datenschutzhinweis

- Repository unbedingt **privat** halten
- Zugangsdaten niemals direkt in Code oder Workflow-Dateien schreiben —
  ausschließlich über GitHub Secrets bzw. die lokale `.env`
- Exportierte Artifacts enthalten alle Kontaktdaten inkl. Fotos
- Artifacts werden nach 7 Tagen automatisch gelöscht
- Die Dateien `.env` und `*.vcf` sind via `.gitignore` vom Commit ausgeschlossen
