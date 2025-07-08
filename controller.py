import os
import pandas as pd

DATA_DIR = "data/excel_data"

def process_companies():
    """Hauptfunktion: Liest user_input.xlsx und erstellt output.xlsx mit Daten aus Excel-Dateien - NUR RIC-BASIERT"""
    print("🚀 STARTE RIC-BASIERTE VERARBEITUNG...")

    # 1. Lese user_input.xlsx
    try:
        df_input = pd.read_excel("data/user_input.xlsx")
        first_row = df_input.iloc[0]

        input_ric = str(first_row.get("RIC", "")).strip()
        is_focus = str(first_row.get("Focus?", "")).strip().lower() == "ja"

        # Sammle gewünschte Kennzahlen
        excel_fields = df_input["Kennzahlen aus Excel"].dropna().astype(str).str.strip().tolist()

        print(f"📋 Input: RIC='{input_ric}', Focus={is_focus}")
        print(f"📋 Gewünschte Kennzahlen: {excel_fields}")

    except Exception as e:
        print(f"❌ Fehler beim Lesen von user_input.xlsx: {e}")
        return []

    # 2. Finde das Start-Unternehmen nur über RIC
    if not input_ric or input_ric.lower() in ["", "nan", "none"]:
        print("❌ Kein RIC im Input gefunden!")
        return []

    print(f"🔍 Suche nach RIC: '{input_ric}'")
    start_company = find_company_by_ric(input_ric)

    if not start_company:
        print("❌ Start-Unternehmen nicht gefunden!")
        return []

    print(f"✅ Start-Unternehmen gefunden: {start_company['Name']} ({start_company['RIC']})")
    print(f"   Sub-Industry: {start_company.get('Sub-Industry', 'N/A')}")
    print(f"   Focus: {start_company.get('Focus', 'N/A')}")

    # 3. Finde alle Unternehmen der gleichen Gruppe
    if is_focus and start_company.get('Focus'):
        print(f"🎯 Focus-Modus: Suche nach Focus-Gruppe '{start_company['Focus']}'")
        peer_companies = find_companies_by_focus(start_company['Focus'])
    else:
        print(f"🎯 Sub-Industry-Modus: Suche nach Sub-Industry '{start_company.get('Sub-Industry')}'")
        peer_companies = find_companies_by_sub_industry(start_company.get('Sub-Industry'))

    print(f"📊 {len(peer_companies)} Unternehmen der gleichen Gruppe gefunden")

    # 4. Sammle Kennzahlen für alle Unternehmen
    results = []
    for i, company in enumerate(peer_companies, 1):
        print(f"\n🏢 Verarbeite {i}/{len(peer_companies)}: {company['Name']} ({company['RIC']})")

        # Sammle Excel-Kennzahlen
        kennzahlen = get_kennzahlen_for_company(company['RIC'], excel_fields)

        # Erstelle Ergebnis
        result = {
            "Name": company['Name'],
            "RIC": company['RIC'],
            "Sub-Industry": company.get('Sub-Industry', ''),
            "Focus": company.get('Focus', '')
        }
        result.update(kennzahlen)
        results.append(result)

        print(f"✅ {len(kennzahlen)} Kennzahlen gesammelt")

    # 5. Speichere Output
    if results:
        output_path = "data/output.xlsx"
        df_output = pd.DataFrame(results)
        df_output.to_excel(output_path, index=False)

        print(f"\n✅ OUTPUT GESPEICHERT: {output_path}")
        print(f"📊 {len(results)} Unternehmen mit {len(df_output.columns)} Spalten")

        # Zeige Übersicht
        print(f"\n📋 ERGEBNIS-ÜBERSICHT:")
        for i, result in enumerate(results, 1):
            isin = result.get('ISIN', 'N/A')
            print(f"{i}. {result['Name']} ({result['RIC']}) - ISIN: {isin}")

    return results


def find_company_by_ric(ric):
    """Finde Unternehmen anhand des RIC - ROBUSTE VERSION"""
    print(f"🔍 ROBUSTE RIC-Suche: '{ric}'")

    # Teste nur die wichtigsten Excel-Dateien zuerst
    priority_files = [
        "Consumer_Equity_Keyfigures.xlsx",
        "Consumer_Financial_Stability.xlsx",
        "Consumer_Growth_Rates.xlsx",
        "Consumer_Revenue_Profitability_CF.xlsx",
        "Consumer_Working_Capital.xlsx"
    ]

    for file in priority_files:
        file_path = os.path.join(DATA_DIR, file)
        if not os.path.exists(file_path):
            continue

        print(f"📁 Teste prioritäre Datei: {file}")

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                print(f"  📄 Teste Sheet: {sheet_name}")

                # Teste verschiedene Header-Positionen
                for header_row in [0, 1, 2]:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, nrows=200)

                        if "RIC" not in df.columns:
                            continue

                        # Suche nach dem spezifischen RIC
                        matches = df[df["RIC"].astype(str).str.upper().str.strip() == ric.upper().strip()]

                        if not matches.empty:
                            row = matches.iloc[0]

                            # Finde Name-Spalte
                            name_col = None
                            for col in df.columns:
                                col_upper = str(col).upper().strip()
                                if col_upper in ["UNIVERSE", "HOLDING"]:
                                    name_col = col
                                    break

                            company = {
                                "Name": str(row[name_col]).strip() if name_col else f"Company_{ric}",
                                "RIC": str(row["RIC"]).strip(),
                                "Sub-Industry": str(row.get("Sub-Industry", "")).strip(),
                                "Focus": str(row.get("Focus", "")).strip()
                            }
                            print(f"✅ GEFUNDEN in {file}, Sheet {sheet_name}: {company['Name']} ({company['RIC']})")
                            return company

                    except Exception as e:
                        print(f"    ❌ Header {header_row+1}: {str(e)[:50]}")
                        continue

        except Exception as e:
            print(f"  ❌ Datei-Fehler: {str(e)[:50]}")
            continue

    print(f"❌ RIC '{ric}' nicht in prioritären Dateien gefunden")
    return None


def find_companies_by_focus(focus):
    """Finde alle Unternehmen der gleichen Focus-Gruppe"""
    companies = []

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        path = os.path.join(DATA_DIR, file)

        try:
            xls = pd.ExcelFile(path)
            for sheet_name in xls.sheet_names:
                for header_row in [0, 1, 2]:
                    try:
                        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)

                        # Prüfe ob alle nötigen Spalten vorhanden sind
                        required_cols = ["RIC", "Focus"]
                        name_col = None
                        for col in df.columns:
                            if str(col).upper().strip() in ["UNIVERSE", "HOLDING"]:
                                name_col = col
                                break

                        if not all(col in df.columns for col in required_cols) or not name_col:
                            continue

                        # Finde alle mit gleichem Focus
                        matches = df[df["Focus"].astype(str).str.strip().str.lower() == focus.lower().strip()]

                        for _, row in matches.iterrows():
                            if pd.notna(row[name_col]) and pd.notna(row["RIC"]):
                                company = {
                                    "Name": str(row[name_col]).strip(),
                                    "RIC": str(row["RIC"]).strip(),
                                    "Sub-Industry": str(row.get("Sub-Industry", "")).strip(),
                                    "Focus": str(row["Focus"]).strip()
                                }

                                # Vermeide Duplikate
                                if not any(c["RIC"] == company["RIC"] for c in companies):
                                    companies.append(company)

                    except Exception:
                        continue

        except Exception:
            continue

    return companies


def find_companies_by_sub_industry(sub_industry):
    """Finde alle Unternehmen der gleichen Sub-Industry"""
    companies = []

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        path = os.path.join(DATA_DIR, file)

        try:
            xls = pd.ExcelFile(path)
            for sheet_name in xls.sheet_names:
                for header_row in [0, 1, 2]:
                    try:
                        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)

                        # Prüfe ob alle nötigen Spalten vorhanden sind
                        required_cols = ["RIC", "Sub-Industry"]
                        name_col = None
                        for col in df.columns:
                            if str(col).upper().strip() in ["UNIVERSE", "HOLDING"]:
                                name_col = col
                                break

                        if not all(col in df.columns for col in required_cols) or not name_col:
                            continue

                        # Finde alle mit gleicher Sub-Industry
                        matches = df[df["Sub-Industry"].astype(str).str.strip().str.lower() == sub_industry.lower().strip()]

                        for _, row in matches.iterrows():
                            if pd.notna(row[name_col]) and pd.notna(row["RIC"]):
                                company = {
                                    "Name": str(row[name_col]).strip(),
                                    "RIC": str(row["RIC"]).strip(),
                                    "Sub-Industry": str(row["Sub-Industry"]).strip(),
                                    "Focus": str(row.get("Focus", "")).strip()
                                }

                                # Vermeide Duplikate
                                if not any(c["RIC"] == company["RIC"] for c in companies):
                                    companies.append(company)

                    except Exception:
                        continue

        except Exception:
            continue

    return companies


def get_kennzahlen_for_company(ric, fields):
    """Sammle alle gewünschten Kennzahlen für ein Unternehmen basierend auf RIC"""
    kennzahlen = {}

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        path = os.path.join(DATA_DIR, file)

        try:
            xls = pd.ExcelFile(path)
            for sheet_name in xls.sheet_names:
                # Suche RIC in verschiedenen Header-Positionen
                for header_row in [0, 1, 2]:
                    try:
                        df_raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
                        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)

                        # Korrigiere Spaltennamen falls nötig (für ISIN etc.)
                        if header_row > 0:
                            new_columns = []
                            for col_idx, orig_col in enumerate(df.columns):
                                better_name = None
                                for row_above in range(header_row):
                                    if col_idx < len(df_raw.columns):
                                        cell_value = df_raw.iloc[row_above, col_idx]
                                        if pd.notna(cell_value) and str(cell_value).strip():
                                            cell_str = str(cell_value).strip()
                                            if any(kw in cell_str.upper() for kw in ["ISIN", "MARKET", "FLOAT", "CURRENCY"]):
                                                better_name = cell_str
                                                break

                                if better_name and str(orig_col).startswith("Unnamed"):
                                    new_columns.append(better_name)
                                else:
                                    new_columns.append(str(orig_col))
                            df.columns = new_columns

                        if "RIC" not in df.columns:
                            continue

                        # Finde die Zeile mit dem RIC
                        ric_matches = df[df["RIC"].astype(str).str.upper().str.strip() == ric.upper().strip()]

                        if not ric_matches.empty:
                            row = ric_matches.iloc[0]

                            # Sammle alle gewünschten Felder
                            for field in fields:
                                if field in kennzahlen:  # Bereits gefunden
                                    continue

                                # Direkte Suche
                                if field in df.columns:
                                    value = row[field]
                                    if pd.notna(value) and str(value).strip():
                                        kennzahlen[field] = value
                                        continue

                                # Fuzzy-Suche nach ähnlichen Spaltennamen
                                for col in df.columns:
                                    if field.lower() in str(col).lower() or str(col).lower() in field.lower():
                                        value = row[col]
                                        if pd.notna(value) and str(value).strip():
                                            kennzahlen[field] = value
                                            break

                    except Exception:
                        continue

        except Exception:
            continue

    return kennzahlen
