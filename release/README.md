# CSVtoXL (Portable)

Ein leistungsstarkes, portables Tool zur automatischen Konvertierung von CSV-Dateien in professionell formatierte Excel-Dateien (.xlsx).

[![GitHub](https://img.shields.io/badge/GitHub-CSVtoXL-blue?logo=github)](https://github.com/sphadoinkl/CSVtoXL)

## ✨ Features

- **Portable `.exe`**: Keine Installation erforderlich.
- **Smarte Erkennung**: Erkennt automatisch das Trennzeichen (Semikolon `;` oder Komma `,`).
- **Professionelles Design**: Erzeugt eine "Intelligente Tabelle" in dezentem Blau.
- **Auto-Styling**: Automatische Anpassung der Spaltenbreiten.
- **Kein Konsolen-Fenster**: Saubere Ausführung im Hintergrund.
- **Quick-Access**: Öffnet nach der Konvertierung automatisch den Ordner und markiert die neue Datei.

## 🚀 Benutzung

### 1. Drag & Drop (Empfohlen)

Ziehe einfach eine `.csv`-Datei mit der Maus auf die `CSVtoXL.exe`. Die konvertierte Datei erscheint sofort im selben Ordner.

### 2. "Senden an" Menü (Profi-Tipp)

Für noch schnelleren Zugriff kannst du das Tool in dein Rechtsklick-Menü einbinden:

1. Drücke `Win + R`, gib `shell:sendto` ein und bestätige mit Enter.
2. Erstelle dort eine Verknüpfung zur `CSVtoXL.exe`.
3. Jetzt kannst du jede CSV-Datei mit **Rechtsklick -> Senden an -> CSVtoXL** konvertieren.

### 3. Doppelklick / Verknüpfung

Du kannst auch eine Verknüpfung auf dem Desktop erstellen und Dateien darauf ziehen.

## 🛠 Integration & Automatisierung

Dieser Abschnitt ist für IT-Spezialisten gedacht. Er bedeutet, dass man das Tool auch über andere Programme oder Skripte (wie PowerShell) aufrufen kann.

Das Tool nimmt den Pfad zur CSV-Datei als "Argument" (Parameter) entgegen:
`CSVtoXL.exe "C:\Pfad\zur\datei.csv"`

Dies ist technisch gesehen genau das, was passiert, wenn du eine Datei per **Drag & Drop** auf die `.exe` ziehst.

## ⚙️ Einstellungen (config.json)

Beim ersten Start erstellt das Tool automatisch eine `config.json` im selben Ordner. Du kannst sie mit jedem Texteditor öffnen und anpassen:

- **`design`**: Wähle dein Lieblings-Design (siehe unten).
- **`auto_open_explorer`**: `true` (Öffnet nach der Konvertierung den Windows-Explorer und markiert die neue Datei sofort) oder `false`.
- **`header_cleaning`**: `true` (bereinigt Leerzeichen in Überschriften) oder `false`.
- **`freeze_top_row`**: `true` (fixiert Kopfzeile beim Scrollen, an) oder `false`.
- **`auto_open_file`**: `true` (öffnet die Excel-Datei sofort nach Erstellung) oder `false` (Standard).
- **`send_email`**: `true` (versendet die Datei automatisch als E-Mail) oder `false` (Standard).
- **`email_smtp_server`**: Der SMTP-Server deines E-Mail-Anbieters (z.B. `smtp.gmail.com`).
- **`email_smtp_port`**: Der Port (meist `587` für TLS).
- **`email_sender`**: Deine E-Mail-Adresse.
- **`email_password`**: Dein Passwort (bei Gmail ein "App-Passwort" verwenden!).
- **`email_recipient`**: An wen die Datei gesendet werden soll.
- **`email_subject`**: Betreff der E-Mail (Platzhalter `{filename}` möglich).
- **`output_directory`**: Gib einen festen Pfad an (z.B. `"C:\\Exporte"`) oder lass es leer für den Quellordner.

### Verfügbare Designs

- `Blau (Premium)` - Dunkler Header (wie im Screenshot, Standard)
- `Blau (Standard)` - Klassisches Office-Blau
- `Hellgrau` - Dezent und minimalistisch
- `Dunkelblau` - Kräftige Farben
- `Gruen` - Excel-Style
- `Orange` - Auffällig
- `Kein Design` - Nur die nackten Daten

> [!TIP]
> Du kannst jeden offiziellen Excel-Tabellenstyle verwenden (z.B. `TableStyleMedium 5`). Eine Übersicht aller Designs findest du in der [XlsxWriter Dokumentation](https://xlsxwriter.readthedocs.io/working_with_tables.html#table-styles).

---

## 📄 Lizenz

Dieses Projekt ist unter der **MIT-Lizenz** lizenziert. Weitere Details findest du in der [LICENSE](https://github.com/sphadoinkl/CSVtoXL/blob/main/LICENSE) Datei auf GitHub.

---
Erstellt für effiziente Workflows.
