import pandas as pd
import sys
import os
import subprocess
import csv
import ctypes
import json
import logging
import traceback
import time

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
    "auto_open_file": False,
    "open_email_client": False,
    "output_directory": ""
}

def setup_logging():
    """Konfiguriert das Logging in eine Datei CSVtoXL.log."""
    try:
        log_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
        log_path = os.path.join(log_dir, 'CSVtoXL.log')
        logging.basicConfig(
            filename=log_path,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
        logging.info("--- Programmstart (v1.3.1) ---")
    except:
        pass

def get_config():
    base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
    config_path = os.path.join(base_dir, 'config.json')
    logging.info(f"Suche Konfig unter: {config_path}")
    
    # 1. Falls Datei gar nicht existiert: Erstellen
    if not os.path.exists(config_path):
        try:
            logging.info("Konfigurationsdatei fehlt, erstelle Standard-Datei.")
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CONFIG, f, indent=4)
        except Exception as e:
            logging.error(f"Konnte Standard-Konfig nicht erstellen: {e}")
        return DEFAULT_CONFIG
    
    # 2. Bestehende Datei laden
    try:
        # utf-8-sig ignoriert ein potenzielles Windows-BOM
        with open(config_path, 'r', encoding='utf-8-sig') as f:
            user_config = json.load(f)
        
        logging.info(f"Konfiguration erfolgreich geladen: {user_config}")
        
        # Synchronisierung
        updated = False
        obsolete_keys = ["send_email", "email_smtp_server", "email_smtp_port", "email_sender", "email_password", "email_recipient", "email_subject"]
        for ok in obsolete_keys:
            if ok in user_config:
                del user_config[ok]
                updated = True

        for key, value in DEFAULT_CONFIG.items():
            if key not in user_config:
                user_config[key] = value
                updated = True
        
        if updated:
            try:
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(user_config, f, indent=4)
                logging.info("Konfiguration mit Standardwerten synchronisiert.")
            except:
                pass
        
        return user_config
    except Exception as e:
        logging.error(f"Fehler beim Laden der Konfiguration (nutze Defaults): {str(e)}")
        return DEFAULT_CONFIG

def show_error(message, title="Fehler"):
    """Zeigt eine Windows-Fehlermeldung an und loggt diese."""
    logging.error(f"{title}: {message}")
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)

def trigger_default_email(file_path):
    """Öffnet den Standard-E-Mail Client mit der Datei als Anhang."""
    if not os.path.exists(file_path):
        return
    
    logging.info(f"Trigger E-Mail Client für: {file_path}")
    abs_path = os.path.abspath(file_path)
    ps_cmd = f"$o = New-Object -ComObject Shell.Application; $o.NameSpace('{os.path.dirname(abs_path)}').ParseName('{os.path.basename(abs_path)}').InvokeVerb('email')"
    
    try:
        subprocess.run(["powershell", "-Command", ps_cmd], check=True, capture_output=True)
    except Exception as e:
        show_error(f"E-Mail Client konnte nicht geöffnet werden:\n{str(e)}")

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
    new_columns = []
    for col in df.columns:
        c = str(col).strip()
        c = c.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        new_columns.append(c)
    df.columns = new_columns
    return df

def main():
    setup_logging()
    if len(sys.argv) < 2:
        logging.info("Keine Eingabedatei übergeben.")
        return

    config = get_config()
    input_file = sys.argv[1]
    logging.info(f"Eingabedatei: {input_file}")
    
    if not input_file.lower().endswith('.csv'):
        logging.error(f"Ungültiger Dateityp: {input_file} (nur .csv erlaubt)")
        return

    if not os.path.exists(input_file):
        show_error(f"Datei nicht gefunden: {input_file}")
        return

    # Zielverzeichnis bestimmen
    if config.get("output_directory") and os.path.isdir(config["output_directory"]):
        output_dir = config["output_directory"]
    else:
        output_dir = os.path.dirname(os.path.abspath(input_file))
        
    output_basename = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(output_dir, output_basename + '.xlsx')
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
                logging.info(f"CSV erfolgreich gelesen mit Encoding {enc} (Trennzeichen: '{delim}')")
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

            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max() if not df[col].empty else 0,
                    len(str(col))
                ) + 4
                worksheet.set_column(i, i, min(max_len, 60))

            if config.get("freeze_top_row", True):
                worksheet.freeze_panes(1, 0)

        # Writer IMMER schließen, um Datei freizugeben
        writer.close()
        writer = None
        logging.info(f"Excel-Datei erfolgreich erstellt: {output_file}")

        # 6. E-Mail Client öffnen
        if config.get("open_email_client", False):
            trigger_default_email(output_file)

        # 7. AUTOMATIK: Ordner öffnen
        if config.get("auto_open_explorer", True):
            logging.info("Öffne Windows Explorer...")
            subprocess.run(['explorer', '/select,', output_file])
        else:
            logging.info("Auto-Explorer deaktiviert.")

        # 8. AUTOMATIK: Datei direkt öffnen
        if config.get("auto_open_file", False):
            logging.info("Öffne Excel-Datei...")
            # Kurz warten, damit Windows die Datei registriert
            time.sleep(0.5)
            os.startfile(output_file)
        else:
            logging.info("Auto-File-Open deaktiviert.")

    except Exception as e:
        msg = f"Ein Fehler ist aufgetreten:\n{str(e)}\n\nDetails wurden in CSVtoXL.log gespeichert."
        logging.error(traceback.format_exc())
        show_error(msg)
        if writer:
            try: writer.close()
            except: pass
        if os.path.exists(output_file) and os.path.getsize(output_file) == 0:
            try: os.remove(output_file)
            except: pass
    
    logging.info("--- Programm beendet ---")

if __name__ == "__main__":
    main()