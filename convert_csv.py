import pandas as pd
import sys
import os
import subprocess
import csv
import ctypes
import json

# Mapping für "sprechende" Style-Namen
STYLE_MAPPING = {
    "Blau (Premium)": "TableStyleMedium 2",
    "Blau (Standard)": "TableStyleLight 9",
    "Hellgrau": "TableStyleLight 8",
    "Dunkelblau": "TableStyleMedium 2",
    "Gruen": "TableStyleLight 10",
    "Orange": "TableStyleLight 12",
    "Kein Design": None
}

DEFAULT_CONFIG = {
    "design": "Blau (Premium)",
    "auto_open_explorer": True,
    "header_cleaning": True,
    "freeze_top_row": True,
    "output_directory": ""
}

def get_config():
    config_path = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), 'config.json')
    
    # 1. Falls Datei gar nicht existiert: Erstellen
    if not os.path.exists(config_path):
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CONFIG, f, indent=4)
        except:
            pass
        return DEFAULT_CONFIG
    
    # 2. Bestehende Datei laden
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            user_config = json.load(f)
        
        # Prüfen ob Keys fehlen (Synchronisierung)
        updated = False
        for key, value in DEFAULT_CONFIG.items():
            if key not in user_config:
                user_config[key] = value
                updated = True
        
        if updated:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(user_config, f, indent=4)
        
        return user_config
    except:
        return DEFAULT_CONFIG

def show_error(message):
    """Zeigt eine Windows-Fehlermeldung an."""
    ctypes.windll.user32.MessageBoxW(0, message, "Fehler bei CSV-Konvertierung", 0x10)

def detect_delimiter(file_path, encoding):
    try:
        with open(file_path, 'r', encoding=encoding) as f:
            sample = f.read(2048)
            if not sample:
                return ';'
            dialect = csv.Sniffer().sniff(sample, delimiters=',;')
            return dialect.delimiter
    except Exception:
        return ';'

def clean_headers(df):
    """Reinigt Header: Leerzeichen weg, Zeilenumbrüche weg, seltsame Zeichen zu Leerzeichen."""
    new_columns = []
    for col in df.columns:
        c = str(col).strip()
        c = c.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        new_columns.append(c)
    df.columns = new_columns
    return df

def main():
    if len(sys.argv) < 2:
        return

    config = get_config()
    input_file = sys.argv[1]
    
    if not input_file.lower().endswith('.csv'):
        return

    if not os.path.exists(input_file):
        show_error(f"Datei nicht gefunden: {input_file}")
        return

    # Zielverzeichnis bestimmen
    if config["output_directory"] and os.path.isdir(config["output_directory"]):
        output_dir = config["output_directory"]
    else:
        output_dir = os.path.dirname(os.path.abspath(input_file))
        
    output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file))[0] + '.xlsx')
    writer = None

    try:
        # 1. Daten laden mit Encoding-Fallback
        encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'latin1']
        df = None
        last_error = ""

        for enc in encodings:
            try:
                delim = detect_delimiter(input_file, enc)
                df = pd.read_csv(input_file, sep=delim, encoding=enc)
                break
            except Exception as e:
                last_error = str(e)
                continue

        if df is None:
            raise Exception(f"Konnte CSV nicht lesen. Letzter Fehler: {last_error}")

        # Header Reinigung
        if config.get("header_cleaning", True):
            df = clean_headers(df)

        # 2. Excel-Datei mit Styling erstellen
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Daten-Export')
        workbook  = writer.book
        worksheet = writer.sheets['Daten-Export']

        # 3. Design anwenden
        design_name = config.get("design", "Blau (Premium)")
        excel_style = STYLE_MAPPING.get(design_name, "TableStyleMedium 2")

        (max_row, max_col) = df.shape
        if max_row > 0 and max_col > 0:
            if excel_style:
                columns = [{'header': str(col)} for col in df.columns]
                worksheet.add_table(0, 0, max_row, max_col - 1, {
                    'columns': columns,
                    'style': excel_style,
                })

            # 4. Spaltenbreite automatisch anpassen (Auto-Width)
            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max() if not df[col].empty else 0,
                    len(str(col))
                ) + 4
                worksheet.set_column(i, i, min(max_len, 60))

            # 5. Kopfzeile fixieren (Freeze Panes)
            if config.get("freeze_top_row", True):
                worksheet.freeze_panes(1, 0)

        writer.close()
        writer = None

        # 6. AUTOMATIK: Ordner öffnen
        if config.get("auto_open_explorer", True):
            subprocess.run(['explorer', '/select,', output_file])

    except Exception as e:
        show_error(f"Ein Fehler ist aufgetreten:\n{str(e)}")
        if writer:
            try: writer.close()
            except: pass
        if os.path.exists(output_file) and os.path.getsize(output_file) == 0:
            try: os.remove(output_file)
            except: pass

if __name__ == "__main__":
    main()