import pandas as pd
import os
import re
from functools import lru_cache

DATA_DIR = "excel_data/data"

# Globaler Cache für Excel-Daten
_excel_cache = {}
_files_loaded = set()

def clear_excel_cache():
    """Leert den Excel-Cache"""
    global _excel_cache, _files_loaded
    _excel_cache.clear()
    _files_loaded.clear()
    print("🧹 Excel-Cache geleert")

@lru_cache(maxsize=32)
def get_sector_excel_files(gics_sectors_tuple):
    """
    Cached Version: Filtert Excel-Dateien basierend auf GICS Sektoren
    """
    gics_sectors = list(gics_sectors_tuple) if gics_sectors_tuple else None

    if not gics_sectors:
        # Wenn keine Sektoren angegeben, alle Dateien zurückgeben
        return tuple([os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR)
                if f.endswith(".xlsx") and not f.startswith("~$")])

    # Normalisiere Sektoren für Vergleich
    normalized_sectors = [sector.strip().lower() for sector in gics_sectors]

    # Erweiterte Mapping von Sektor-Namen zu Datei-Patterns
    sector_patterns = {
        'consumer': [r'^Consumer_.*\.xlsx$', r'^Basic.*Consumer.*\.xlsx$'],
        'materials': [r'^Materials_.*\.xlsx$'],
        'health': [r'^Health.*\.xlsx$'],
        'it': [r'^IT.*\.xlsx$', r'^.*Technology.*\.xlsx$'],
        'technology': [r'^IT.*\.xlsx$', r'^.*Technology.*\.xlsx$'],
        'utilities': [r'^Utilities.*\.xlsx$'],
        'housing': [r'^Housing.*\.xlsx$']
    }

    filtered_files = []
    available_files = [f for f in os.listdir(DATA_DIR) if f.endswith(".xlsx") and not f.startswith("~$")]

    print(f"🎯 Gesuchte GICS Sektoren: {gics_sectors}")

    for sector in normalized_sectors:
        if sector in sector_patterns:
            patterns = sector_patterns[sector]
            for pattern in patterns:
                matching_files = [f for f in available_files if re.match(pattern, f, re.IGNORECASE)]
                for file in matching_files:
                    full_path = os.path.join(DATA_DIR, file)
                    if full_path not in filtered_files:
                        filtered_files.append(full_path)
        else:
            # Fallback: Suche nach Dateien die den Sektor-Namen enthalten
            fallback_files = [f for f in available_files if sector in f.lower()]
            for file in fallback_files:
                full_path = os.path.join(DATA_DIR, file)
                if full_path not in filtered_files:
                    filtered_files.append(full_path)

    if not filtered_files:
        print("⚠️ Keine passenden Excel-Dateien gefunden, verwende alle verfügbaren")
        filtered_files = [os.path.join(DATA_DIR, f) for f in available_files]

    print(f"📊 {len(filtered_files)} Excel-Dateien für Sektoren {gics_sectors}")
    return tuple(filtered_files)

def load_excel_files_once(file_paths):
    """
    Lädt alle Excel-Dateien einmalig in den Cache
    """
    global _excel_cache, _files_loaded

    newly_loaded = 0
    for file_path in file_paths:
        if file_path in _files_loaded:
            continue

        file = os.path.basename(file_path)
        print(f"📁 Lade Datei in Cache: {file}")

        try:
            xls = pd.ExcelFile(file_path)
            _excel_cache[file_path] = {}

            for sheet_name in xls.sheet_names:
                try:
                    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    _excel_cache[file_path][sheet_name] = df_raw
                except Exception as e:
                    print(f"❌ Fehler beim Lesen von Sheet {sheet_name}: {e}")
                    continue

            _files_loaded.add(file_path)
            newly_loaded += 1

        except Exception as e:
            print(f"❌ Fehler beim Öffnen von {file}: {e}")
            continue

    if newly_loaded > 0:
        print(f"✅ {newly_loaded} neue Dateien in Cache geladen")

def fetch_excel_kennzahlen_by_ric_filtered(ric: str, fields: list, gics_sectors=None) -> dict:
    """
    VEREINFACHT: Suche Kennzahlen direkt über RIC mit einfachem Header-Vergleich
    """
    result = {}

    # Konvertiere zu Tuple für Caching
    gics_sectors_tuple = tuple(gics_sectors) if gics_sectors else None

    # Hole gefilterte Excel-Dateien (cached)
    excel_files = get_sector_excel_files(gics_sectors_tuple)

    # Lade alle Dateien einmalig in Cache
    load_excel_files_once(excel_files)

    for file_path in excel_files:
        try:
            # Lese Excel-Datei und teste verschiedene Sheets
            xls = pd.ExcelFile(file_path)

            for sheet_name in xls.sheet_names:
                # Priorisiere bestimmte Sheet-Namen
                if not any(keyword in sheet_name.lower() for keyword in ['equity', 'key', 'figures', 'data']):
                    continue

                try:
                    # KORRIGIERT: Finde die richtige Header-Zeile dynamisch
                    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                    # Suche nach Header-Zeile mit RIC
                    header_row = None
                    for i in range(min(10, len(df_raw))):
                        row = df_raw.iloc[i]
                        # Prüfe jede Zelle in der Zeile
                        for cell in row.values:
                            if pd.notna(cell) and str(cell).strip().upper() == "RIC":
                                header_row = i
                                break
                        if header_row is not None:
                            break

                    if header_row is None:
                        continue

                    # Lese mit korrektem Header
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

                    # Prüfe ob RIC-Spalte vorhanden
                    if "RIC" not in df.columns:
                        continue

                    # Suche nach der Zeile mit dem spezifischen RIC
                    matching_rows = df[df['RIC'].astype(str).str.upper().str.strip() == ric.upper().strip()]

                    if matching_rows.empty:
                        continue

                    # Verwende die erste passende Zeile
                    matched_row = matching_rows.iloc[0]
                    print(f"✅ RIC {ric} gefunden in {os.path.basename(file_path)} -> {sheet_name}")

                    # Sammle alle gewünschten Felder aus dieser Zeile
                    for field in fields:
                        if field in result:
                            continue  # Bereits gefunden

                        # KORRIGIERT: Bereinige Feldname von Zeilenumbrüchen
                        clean_field = field.replace('\n', ' ').replace('\r', ' ').strip()

                        # Suche nach exakter Übereinstimmung oder bereinigter Version
                        search_fields = [field, clean_field]

                        found = False
                        for search_field in search_fields:
                            if search_field in df.columns:
                                value = matched_row[search_field]

                                # Verbesserte Überprüfung: Stelle sicher, dass der Wert nicht ein RIC-Code ist
                                if pd.notna(value) and str(value).strip() != "":
                                    str_value = str(value).strip().upper()

                                    # Prüfe auf Fehlermeldungen ZUERST
                                    error_messages = [
                                        "THE RECORD COULD NOT BE FOUND",
                                        "ERROR CODE: 0",
                                        "NO DATA AVAILABLE",
                                        "DATA NOT AVAILABLE",
                                        "N/A",
                                        "#N/A",
                                        "#ERROR",
                                        "NULL"
                                    ]

                                    # Wenn eine Fehlermeldung enthalten ist, setze leeren Wert
                                    is_error_message = any(error_msg in str_value for error_msg in error_messages)

                                    if is_error_message:
                                        print(f"⚠️ Fehlermeldung '{value}' erkannt, setze leeren Wert")
                                        result[field] = ""  # Leerer String statt überspringen
                                        found = True
                                        break

                                    # Prüfe ob der Wert wie ein RIC aussieht (enthält typische RIC-Muster)
                                    is_ric_like = (
                                        # Exakter Match mit dem gesuchten RIC
                                        str_value == ric.upper().strip() or
                                        # Andere typische RIC-Muster (Buchstaben + Punkt + Buchstabe)
                                        bool(re.match(r'^[A-Z]{1,6}\.[A-Z]{1,3}$', str_value)) or
                                        # Nur Buchstaben ohne Punkt (wie "APD", "CLN")
                                        (bool(re.match(r'^[A-Z]{1,6}$', str_value)) and len(str_value) <= 6)
                                    )

                                    if not is_ric_like:
                                        result[field] = value
                                        print(f"✅ Gefunden: {search_field} = {value}")
                                        found = True
                                        break
                                    else:
                                        print(f"⚠️ Wert '{value}' sieht wie ein RIC aus, überspringe")

                        # Wenn nicht gefunden, erweiterte Suche
                        if not found:
                            # Suche nach ähnlichen Spalten (ohne Zeilenumbrüche)
                            for col in df.columns:
                                col_clean = str(col).replace('\n', ' ').replace('\r', ' ').strip()
                                if col_clean.lower() == clean_field.lower():
                                    value = matched_row[col]
                                    if pd.notna(value) and str(value).strip() != "":
                                        str_value = str(value).strip().upper()

                                        # Prüfe auf Fehlermeldungen ZUERST
                                        error_messages = [
                                            "THE RECORD COULD NOT BE FOUND",
                                            "ERROR CODE: 0",
                                            "NO DATA AVAILABLE",
                                            "DATA NOT AVAILABLE",
                                            "N/A",
                                            "#N/A",
                                            "#ERROR",
                                            "NULL"
                                        ]

                                        # Wenn eine Fehlermeldung enthalten ist, setze leeren Wert
                                        is_error_message = any(error_msg in str_value for error_msg in error_messages)

                                        if is_error_message:
                                            print(f"⚠️ Fehlermeldung '{value}' erkannt, setze leeren Wert")
                                            result[field] = ""  # Leerer String statt überspringen
                                            break

                                        # Gleiche RIC-Überprüfung wie oben
                                        is_ric_like = (
                                            str_value == ric.upper().strip() or
                                            bool(re.match(r'^[A-Z]{1,6}\.[A-Z]{1,3}$', str_value)) or
                                            (bool(re.match(r'^[A-Z]{1,6}$', str_value)) and len(str_value) <= 6)
                                        )

                                        if not is_ric_like:
                                            result[field] = value
                                            print(f"✅ Gefunden (ähnlich): {col} = {value}")
                                            break
                                        else:
                                            print(f"⚠️ Wert '{value}' sieht wie ein RIC aus, überspringe")

                    # Wenn Kennzahlen gefunden wurden, breche Sheet-Schleife ab
                    if result:
                        break

                except Exception as e:
                    continue

        except Exception as e:
            # Debug: Zeige welche Datei Probleme macht
            print(f"❌ Fehler in {file_path}: {e}")
            continue

    return result


def fetch_excel_kennzahlen(name: str, gruppe: str, fields: list) -> dict:
    result = {}
    print(f"🔍 Suche nach Kennzahlen für: {name}")
    print(f"📋 Gewünschte Felder: {fields}")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue
        print(f"📁 Durchsuche Datei: {file}")
        path = os.path.join(DATA_DIR, file)
        xls = pd.ExcelFile(path)

        for sheet_name in xls.sheet_names:
            print(f"📄 Sheet: {sheet_name}")

            # Lese erst ohne Header, um dynamisch zu suchen
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            # Suche nach Header-Zeile mit einer Namensspalte oder RIC
            header_row = None
            name_column = None

            for i in range(min(10, len(df_raw))):
                row = df_raw.iloc[i]
                row_str = row.astype(str).str.lower().str.strip()

                # Prüfe auf typische Namensspalten oder RIC
                if any(col in ["holding", "universe", "ric"] for col in row_str.values):
                    header_row = i
                    # Bestimme die Namensspalte (bevorzuge Holding/Universe über RIC)
                    for j, col_name in enumerate(row_str.values):
                        if col_name in ["holding", "universe"]:
                            name_column = j
                            break
                    # Falls keine Holding/Universe gefunden, nutze RIC als Fallback
                    if name_column is None:
                        for j, col_name in enumerate(row_str.values):
                            if col_name == "ric":
                                name_column = j
                                break
                    if name_column is not None:
                        break

            if header_row is None or name_column is None:
                print(f"⚠️ Keine passende Header-Zeile in {sheet_name}")
                continue

            # Lese mit dem gefundenen Header
            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

            # Korrigiere Spaltennamen mit Informationen aus vorherigen Zeilen
            if header_row > 0:
                new_columns = []
                for col_idx, orig_col in enumerate(df.columns):
                    # Prüfe die Zeilen oberhalb des Headers für bessere Spaltennamen
                    better_name = None
                    for row_above in range(header_row):
                        if col_idx < len(df_raw.columns):
                            cell_value = df_raw.iloc[row_above, col_idx]
                            if pd.notna(cell_value) and str(cell_value).strip() != "":
                                cell_str = str(cell_value).strip()
                                # Prüfe auf wichtige Kennzahlen-Namen
                                cell_upper = cell_str.upper()
                                if any(keyword in cell_upper for keyword in ["ISIN", "FLOAT", "FREE", "MARKET", "CURRENCY", "P/E", "P/B", "ROE", "ROA", "EBIT", "EBITDA"]):
                                    better_name = cell_str
                                    break

                    if better_name and str(orig_col).startswith("Unnamed"):
                        new_columns.append(better_name)
                        print(f"🔧 Spalte korrigiert: '{orig_col}' → '{better_name}'")
                    else:
                        new_columns.append(str(orig_col).strip())

                df.columns = new_columns

            # Bestimme den echten Spaltennamen für die Namenssuche
            name_col_name = df.columns[name_column]
            print(f"🎯 Verwende Namensspalte: {name_col_name}")

            # Suche nach der Zeile mit dem passenden Namen (auch RIC als Fallback)
            matching_rows = []

            # Zuerst versuche Name-Match
            for idx, row in df.iterrows():
                cell_value = row[name_col_name]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).lower().strip()
                    if cell_str == name.lower().strip():
                        matching_rows.append(idx)
                        break  # Nehme den ersten gefundenen Match

            # Falls kein Name-Match, versuche RIC-Match falls verfügbar
            if not matching_rows and "RIC" in df.columns:
                print(f"🔄 Kein Name-Match, versuche RIC-Match...")
                for idx, row in df.iterrows():
                    ric_value = row.get("RIC")
                    if pd.notna(ric_value) and str(ric_value).strip() != "":
                        # Extrahiere RIC aus dem ursprünglichen Namen falls vorhanden
                        # (manchmal ist der RIC Teil des Namens oder in der Gruppe)
                        print(f"🔍 Prüfe RIC: {ric_value}")
                        matching_rows.append(idx)
                        break

            if not matching_rows:
                print(f"⚠️ Name '{name}' nicht in {sheet_name} gefunden")
                continue

            # Verwende die erste passende Zeile
            match_idx = matching_rows[0]
            matched_row = df.iloc[match_idx]
            print(f"✅ Name gefunden in Zeile {match_idx + header_row + 1}")

            # Sammle alle verfügbaren Felder aus dieser Zeile
            for field in fields:
                if field in result:
                    continue  # Bereits gefunden

                value = None
                # Direkte Suche nach Feldname
                if field in df.columns:
                    value = matched_row[field]
                    if pd.notna(value) and str(value).strip() != "":
                        result[field] = value
                        print(f"✅ {field}: {value}")
                        continue

                # Fuzzy-Suche nach ähnlichen Spaltennamen
                for col in df.columns:
                    col_clean = str(col).strip()
                    field_clean = field.strip()

                    # Case-insensitive Vergleich
                    if col_clean.lower() == field_clean.lower():
                        value = matched_row[col]
                        if pd.notna(value) and str(value).strip() != "":
                            result[field] = value
                            print(f"✅ {field} (als {col}): {value}")
                            break

                    # Prüfe, ob das Feld im Spaltennamen enthalten ist
                    elif field_clean.lower() in col_clean.lower() or col_clean.lower() in field_clean.lower():
                        value = matched_row[col]
                        if pd.notna(value) and str(value).strip() != "":
                            result[field] = value
                            print(f"✅ {field} (ähnlich: {col}): {value}")
                            break

    print(f"📊 Gesammelte Kennzahlen: {list(result.keys())}")
    return result
def resolve_ric_by_name(name: str) -> str:
    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue
        path = os.path.join(DATA_DIR, file)
        df = pd.read_excel(path)
        name_column = None
        for col in df.columns:
            if str(col).strip().lower() in ["holding", "universe"]:
                name_column = col
                break
        if not name_column:
            continue
        if "RIC" not in df.columns:
            continue
        match = df[df[name_column].astype(str).str.lower() == name.lower()]
        if not match.empty:
            return match["RIC"].iloc[0]
    return ""

def resolve_name_by_ric(ric: str) -> str:
    print("📁 Starte RIC-Suche in Excel-Dateien...")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        print(f"🔍 Öffne Datei: {file}")
        path = os.path.join(DATA_DIR, file)
        xls = pd.ExcelFile(path)
        for sheet_name in xls.sheet_names:
            # Dynamisch nach Header mit "RIC" suchen (bis Zeile 10)
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            header_row = None
            for i in range(min(10, len(df_raw))):
                row = df_raw.iloc[i]
                # Neue Erkennung: prüfe, ob einer der Werte exakt "RIC" ist (Großschreibung, ohne Leerzeichen)
                if row.astype(str).str.upper().str.strip().isin(["RIC"]).any():
                    header_row = i
                    break
            if header_row is None:
                print(f"⚠️ Keine Kopfzeile mit 'RIC' in Sheet {sheet_name}")
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
            print(f"📄 [Sheet: {sheet_name}] Spalten:", df.columns.tolist())

            if "RIC" not in df.columns:
                print("⚠️ RIC-Spalte fehlt")
                continue

            # Clean RIC-Spalte
            df["RIC_clean"] = df["RIC"].astype(str).str.upper().str.strip()
            ric_clean = ric.upper().strip()
            print("📃 Enthaltene RICs:", df["RIC_clean"].dropna().unique())

            match = df[df["RIC_clean"] == ric_clean]
            if not match.empty:
                name_column = None
                for col in df.columns:
                    if str(col).strip().lower() in ["holding", "universe"]:
                        name_column = col
                        break
                if name_column:
                    print(f"✅ Treffer: {ric_clean} → {match[name_column].iloc[0]}")
                    return match[name_column].iloc[0]
                else:
                    print(f"⚠️ Kein Name-Spaltenmatch in {file}")
    print(f"❌ Kein Treffer für RIC '{ric}' gefunden.")
    return ""

def fetch_excel_kennzahlen_by_ric(ric: str, fields: list) -> dict:
    """
    Wrapper-Funktion für Rückwärtskompatibilität:
    Suche Kennzahlen direkt über RIC ohne GICS Sector-Filter
    """
    return fetch_excel_kennzahlen_by_ric_filtered(ric, fields, gics_sectors=None)
