import pandas as pd
import os

DATA_DIR = "data/excel_data"

def fetch_excel_kennzahlen_by_ric(ric: str, fields: list) -> dict:
    """Neue Funktion: Suche Kennzahlen direkt über RIC"""
    result = {}
    print(f"🔍 Suche nach Kennzahlen für RIC: {ric}")
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

            # Suche nach Header-Zeile mit RIC
            header_row = None
            for i in range(min(10, len(df_raw))):
                row = df_raw.iloc[i]
                row_str = row.astype(str).str.lower().str.strip()
                if "ric" in row_str.values:
                    header_row = i
                    break

            if header_row is None:
                print(f"⚠️ Keine RIC-Header-Zeile in {sheet_name}")
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

            if "RIC" not in df.columns:
                print(f"⚠️ Keine RIC-Spalte in {sheet_name}")
                continue

            # Suche nach der Zeile mit dem spezifischen RIC
            matching_rows = []
            for idx, row in df.iterrows():
                ric_value = row.get("RIC")
                if pd.notna(ric_value) and str(ric_value).upper().strip() == ric.upper().strip():
                    matching_rows.append(idx)

            if not matching_rows:
                print(f"⚠️ RIC '{ric}' nicht in {sheet_name} gefunden")
                continue

            # Verwende die erste passende Zeile
            match_idx = matching_rows[0]
            matched_row = df.iloc[match_idx]
            print(f"✅ RIC {ric} gefunden in Zeile {match_idx + header_row + 1}")

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

    print(f"📊 Gesammelte Kennzahlen für {ric}: {list(result.keys())}")
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