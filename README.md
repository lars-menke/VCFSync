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

1. VCF-Inhalt in die Zwischenablage kopieren
2. GitHub → Tab `Actions` → `iCloud Kontakte Import`
3. Klick auf `Run workflow` → VCF-Inhalt in das Textfeld einfügen →
   `Run workflow`
4. Bei großen Dateien (>65.000 Zeichen): Skript **lokal** ausführen (siehe unten)

   > GitHub Actions unterstützt keine Datei-Uploads als Workflow-Input, daher
   > wird der VCF-Inhalt als Text in eine Textarea eingefügt. Diese ist auf
   > **65.536 Zeichen** limitiert. Bei vielen Kontakten mit Base64-Fotos wird
   > die Datei schnell größer als dieses Limit — dann lokal importieren.

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
