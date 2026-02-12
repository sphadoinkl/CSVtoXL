# Changelog - CSVtoXL

## [1.3.1] - 2026-02-12

### Behoben

- **Auto-Open Fix**: Problem behoben, bei dem die `config.json` manchmal ignoriert wurde und die Excel-Datei nicht automatisch geöffnet werden konnte.
- **Logging erweitert**: Detaillierte Ablauf-Protokollierung in `CSVtoXL.log`.

## [1.3] - 2026-02-12

### Hinzugefügt

- **Metadaten**: Die `.exe` enthält nun professionelle Datei-Informationen (Version 1.3, Produktname, Lizenz).
- **Fehler-Protokoll**: Bei Abstürzen oder Fehlern wird automatisch eine `CSVtoXL.log` erstellt, um die Fehlersuche zu erleichtern.

## [1.2] - 2026-02-12

### Gelöscht / Geändert

- **SMTP entfernt**: Der automatische E-Mail-Versand im Hintergrund wurde durch Öffnen des E-Mail-Clients ersetzt.

### Hinzugefügt

- **E-Mail-Client Integration**: Neue Option `open_email_client` öffnet den lokalen Mail-Client (z.B. Outlook) mit der Datei im Anhang.

## [1.1] - 2026-02-12

### Hinzugefügt

- Option zum direkten Öffnen der Excel-Datei nach Konvertierung (`auto_open_file`).

## [1.0] - 2026-02-12

### Hinzugefügt

- Erster offizieller Release als **CSVtoXL**.
- Premium Design (Dunkelblau) als Standard.
- Freeze Panes (oberste Zeile fixiert).
- Auto-Spaltenbreite.
- Automatische Erkennung von CSV-Trennzeichen.
- Header-Reinigung.
- Konfigurations-Datei `config.json` mit Auto-Synchronisierung neuer Keys.
- Portabler Build ohne Konsole mit Icon.
- GitHub Integration und Release-Ordner Struktur.
