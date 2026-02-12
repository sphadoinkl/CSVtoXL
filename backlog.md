# Backlog - CSVtoXL Ideen-Sammlung

Hier werden zukünftige Features und Verbesserungsvorschläge gesammelt, um das Tool weiter zu professionalisieren.

## 🚀 Funktionale Erweiterungen

- [ ] **Batch-Verarbeitung**: Mehrere CSV-Dateien gleichzeitig per Drag & Drop konvertieren.
- [ ] **Zusammenführen (Merging)**: Mehrere CSVs in eine einzige Excel-Datei mit mehreren Arbeitsblättern (Tabs) kombinieren.
- [ ] **Datums-Erkennung**: Automatisches Formatieren von Spalten, die Datumsangaben enthalten, als echtes Excel-Datum.
- [ ] **Formeln unterstützen**: Erkennung von Feldern, die wie Formeln aussehen, und deren Konvertierung in aktive Excel-Formeln.

## 🎨 Optik & User Experience

- [ ] **Minimalistisches GUI**: Ein kleines Fenster für Einstellungen (statt `config.json` manuell zu editieren).
- [ ] **Benutzerdefinierte Styles**: Unterstützung für eigene Excel-Tabellen-Styles über die Config.
- [ ] **Tray-Icon**: Option, das Tool im System-Tray laufen zu lassen für noch schnelleren Zugriff.
- [ ] **Fortschrittsbalken**: Anzeige des Status bei sehr großen CSV-Dateien (> 500MB).

## 🛠️ Technische Verbesserungen

- [ ] **Auto-Update**: Prüfung beim Start, ob eine neue Version auf GitHub verfügbar ist.
- [ ] **Plugin-System**: Unterstützung für Python-Skripte, die Daten vor der Konvertierung manipulieren (z.B. Zeilen filtern).
- [ ] **SQLite Support**: CSV-Daten zusätzlich in eine temporäre SQLite-Datenbank laden für SQL-Abfragen vor dem Export.

## 🖇️ Integration

- [ ] **Rechter-Mausklick-Menü**: "Konvertieren mit CSVtoXL" direkt in das Windows-Kontextmenü für CSV-Dateien einbauen.
- [ ] **Cloud-Upload**: Direkter Upload der fertigen Datei nach SharePoint oder OneDrive.
