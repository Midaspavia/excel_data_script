import os
import pandas as pd
from excel_kennzahlen import fetch_excel_kennzahlen_by_ric
import glob

DATA_DIR = "excel_data/data"

def cleanup_temp_files():
    """Bereinigt temporäre Excel-Dateien (~$*.xlsx) nach der Ausführung"""
    print("🧹 BEREINIGE TEMPORÄRE DATEIEN...")

    # Suche in allen relevanten Verzeichnissen
    directories = ["excel_data/", "excel_data/data/", "."]

    deleted_count = 0
    for directory in directories:
        if os.path.exists(directory):
            temp_files = glob.glob(os.path.join(directory, "~$*.xlsx"))
            for temp_file in temp_files:
                try:
                    os.remove(temp_file)
                    print(f"🗑️  Gelöscht: {temp_file}")
                    deleted_count += 1
                except Exception as e:
                    print(f"⚠️  Fehler beim Löschen von {temp_file}: {e}")

    if deleted_count > 0:
        print(f"✅ {deleted_count} temporäre Dateien bereinigt")
    else:
        print("✅ Keine temporären Dateien gefunden")

def process_companies():
    """Hauptfunktion: Liest input_user.xlsx und erstellt output.xlsx mit Daten aus Excel-Dateien - UNTERSTÜTZT RIC UND NAMEN"""
    print("🚀 STARTE VERARBEITUNG (RIC oder NAME)...")

    try:
        # 1. Lese input_user.xlsx
        try:
            df_input = pd.read_excel("excel_data/input_user.xlsx")
            first_row = df_input.iloc[0]

            # KORRIGIERT: Spalte A = Name, Spalte B = RIC
            input_name = str(first_row.iloc[0] if len(first_row) > 0 else "").strip()  # Spalte A
            input_ric = str(first_row.iloc[1] if len(first_row) > 1 else "").strip()   # Spalte B
            is_focus = str(first_row.get("Focus?", "")).strip().lower() == "ja"

            # Sammle gewünschte Kennzahlen
            excel_fields = df_input["Kennzahlen aus Excel"].dropna().astype(str).str.strip().tolist()

            print(f"📋 Input Name (Spalte A): '{input_name}'")
            print(f"📋 Input RIC (Spalte B): '{input_ric}'")
            print(f"📋 Focus: {is_focus}")
            print(f"📋 Gewünschte Kennzahlen: {excel_fields}")

        except Exception as e:
            print(f"❌ Fehler beim Lesen von input_user.xlsx: {e}")
            return []

        # 2. Bestimme Suchstrategie: RIC hat Priorität, dann Name
        if input_ric and input_ric.lower() not in ["", "nan", "none"]:
            # RIC vorhanden - nutze RIC-Suche
            print(f"🔍 PRIORITÄT: RIC-Suche für '{input_ric}'")
            start_company = find_company_by_ric(input_ric)
        elif input_name and input_name.lower() not in ["", "nan", "none"]:
            # Kein RIC, aber Name vorhanden - nutze Name-Suche
            if len(input_name) < 4:
                print("❌ Name-Suche erfordert mindestens 4 Zeichen!")
                return []
            print(f"🔍 FALLBACK: Name-Suche für '{input_name}' (Teilwort-Suche)")
            start_company = find_company_by_name(input_name)
        else:
            print("❌ Weder RIC noch Name im Input gefunden!")
            return []

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
            output_path = "excel_data/output.xlsx"
            df_output = pd.DataFrame(results)
            df_output.to_excel(output_path, index=False)

            print(f"\n✅ OUTPUT GESPEICHERT: {output_path}")
            print(f"📊 {len(results)} Unternehmen mit {len(df_output.columns)} Spalten")

            # Zeige Übersicht
            print(f"\n📋 ERGEBNIS-ÜBERSICHT:")
            for i, result in enumerate(results, 1):
                print(f"\n{i}. {result['Name']} ({result['RIC']})")
                print(f"   Sub-Industry: {result.get('Sub-Industry', 'N/A')}")
                print(f"   Focus: {result.get('Focus', 'N/A')}")

                # Zeige alle Kennzahlen aus Excel
                for field in excel_fields:
                    value = result.get(field, 'N/A')
                    if value != 'N/A' and pd.notna(value):
                        print(f"   {field}: {value}")
                    else:
                        print(f"   {field}: ❌ Nicht gefunden")

        return results

    finally:
        # Bereinige temporäre Dateien nach der Ausführung (wird IMMER ausgeführt)
        cleanup_temp_files()


def find_company_by_ric(ric):
    """Finde Unternehmen anhand des RIC - direkte Index-Positionen"""
    print(f"🔍 RIC-Suche: '{ric}' (RIC=Spalte E, Focus=Spalte D, Sub-Industry=Spalte C)")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        file_path = os.path.join(DATA_DIR, file)

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                # Nur relevante Sheets
                if not any(pattern in sheet_name.lower() for pattern in ["equity", "key", "revenue", "profitability", "financial", "growth", "figures"]):
                    continue

                try:
                    # Lese mit Header=2 (Zeile 3)
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2, nrows=200)

                    # Prüfe ob genügend Spalten vorhanden sind
                    if len(df.columns) < 5:
                        continue

                    # RIC ist in Spalte E (Index 4)
                    ric_col = df.columns[4]

                    # Suche nach dem spezifischen RIC
                    matches = df[df[ric_col].astype(str).str.upper().str.strip() == ric.upper().strip()]

                    if not matches.empty:
                        row = matches.iloc[0]

                        # Direkte Index-Zugriffe
                        sub_industry = str(row.iloc[2]).strip()  # Spalte C
                        focus_value = str(row.iloc[3]).strip()   # Spalte D
                        ric_value = str(row.iloc[4]).strip()     # Spalte E

                        # Finde Name-Spalte (Universe oder Holding)
                        name_value = "Unknown"
                        if len(df.columns) > 1:
                            name_value = str(row.iloc[1]).strip()  # Spalte B (Universe)

                        company = {
                            "Name": name_value,
                            "RIC": ric_value,
                            "Sub-Industry": sub_industry,
                            "Focus": focus_value
                        }
                        print(f"✅ GEFUNDEN: {company['Name']} ({company['RIC']})")
                        print(f"   Sub-Industry (Spalte C): '{company['Sub-Industry']}'")
                        print(f"   Focus (Spalte D): '{company['Focus']}'")
                        return company

                except Exception as e:
                    continue

        except Exception as e:
            continue

    print(f"❌ RIC '{ric}' nicht gefunden")
    return None


def find_companies_by_focus(focus):
    """Suche alle Unternehmen mit gleichem Focus (Spalte D)"""
    companies = []
    print(f"🔍 Focus-Suche in Spalte D: '{focus}'")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        file_path = os.path.join(DATA_DIR, file)
        print(f"📁 Durchsuche: {file}")

        try:
            xls = pd.ExcelFile(file_path)

            for sheet_name in xls.sheet_names:
                # Nur relevante Sheets
                if not any(pattern in sheet_name.lower() for pattern in ["equity", "key", "revenue", "profitability", "financial", "growth", "figures"]):
                    continue

                try:
                    # Lese mit Header=2 (Zeile 3)
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

                    # Prüfe ob genügend Spalten vorhanden sind
                    if len(df.columns) < 5:
                        continue

                    # Suche in Spalte D (Index 3) nach gleichem Focus
                    found_in_sheet = 0
                    for _, row in df.iterrows():
                        if len(row) >= 5 and pd.notna(row.iloc[1]) and pd.notna(row.iloc[3]) and pd.notna(row.iloc[4]):
                            row_focus = str(row.iloc[3]).strip()  # Spalte D

                            if row_focus == focus.strip():
                                company = {
                                    "Name": str(row.iloc[1]).strip(),       # Spalte B
                                    "RIC": str(row.iloc[4]).strip(),        # Spalte E
                                    "Sub-Industry": str(row.iloc[2]).strip(), # Spalte C
                                    "Focus": row_focus                       # Spalte D
                                }

                                # Vermeide Duplikate
                                if not any(c["RIC"] == company["RIC"] for c in companies):
                                    companies.append(company)
                                    found_in_sheet += 1

                    if found_in_sheet > 0:
                        print(f"  📄 {sheet_name}: {found_in_sheet} Unternehmen gefunden")

                except Exception as e:
                    continue

        except Exception as e:
            continue

    print(f"📊 GESAMT: {len(companies)} Unternehmen mit Focus '{focus}' gefunden")
    return companies


def find_companies_by_sub_industry(sub_industry):
    """Suche alle Unternehmen mit gleicher Sub-Industry (Spalte C)"""
    companies = []
    print(f"🔍 Sub-Industry-Suche in Spalte C: '{sub_industry}'")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        file_path = os.path.join(DATA_DIR, file)
        print(f"📁 Durchsuche: {file}")

        try:
            xls = pd.ExcelFile(file_path)

            for sheet_name in xls.sheet_names:
                # Nur relevante Sheets
                if not any(pattern in sheet_name.lower() for pattern in ["equity", "key", "revenue", "profitability", "financial", "growth", "figures"]):
                    continue

                try:
                    # Lese mit Header=2 (Zeile 3)
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

                    # Prüfe ob genügend Spalten vorhanden sind
                    if len(df.columns) < 5:
                        continue

                    # Suche in Spalte C (Index 2) nach gleicher Sub-Industry
                    found_in_sheet = 0
                    for _, row in df.iterrows():
                        if len(row) >= 5 and pd.notna(row.iloc[1]) and pd.notna(row.iloc[2]) and pd.notna(row.iloc[4]):
                            row_sub_industry = str(row.iloc[2]).strip()  # Spalte C

                            if row_sub_industry == sub_industry.strip():
                                company = {
                                    "Name": str(row.iloc[1]).strip(),       # Spalte B
                                    "RIC": str(row.iloc[4]).strip(),        # Spalte E
                                    "Sub-Industry": row_sub_industry,        # Spalte C
                                    "Focus": str(row.iloc[3]).strip()       # Spalte D
                                }

                                # Vermeide Duplikate
                                if not any(c["RIC"] == company["RIC"] for c in companies):
                                    companies.append(company)
                                    found_in_sheet += 1

                    if found_in_sheet > 0:
                        print(f"  📄 {sheet_name}: {found_in_sheet} Unternehmen gefunden")

                except Exception as e:
                    continue

        except Exception as e:
            continue

    print(f"📊 GESAMT: {len(companies)} Unternehmen mit Sub-Industry '{sub_industry}' gefunden")
    return companies


def get_kennzahlen_for_company(ric, fields):
    """Sammelt alle gewünschten Kennzahlen für ein Unternehmen basierend auf RIC (nutzt robusten Import aus excel_kennzahlen.py)"""
    return fetch_excel_kennzahlen_by_ric(ric, fields)

def find_company_by_name(name):
    """Finde Unternehmen anhand des Namens - Suche in Holding/Universe"""
    print(f"🔍 Name-Suche: '{name}' (Teilwort-Suche in Holding/Universe)")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        file_path = os.path.join(DATA_DIR, file)

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                # Nur relevante Sheets
                if not any(pattern in sheet_name.lower() for pattern in ["equity", "key", "revenue", "profitability", "financial", "growth", "figures"]):
                    continue

                try:
                    # Lese mit Header=2 (Zeile 3)
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

                    # Prüfe ob genügend Spalten vorhanden sind
                    if len(df.columns) < 5:
                        continue

                    # Finde Name-Spalte (Universe oder Holding)
                    name_col = df.columns[1]
                    ric_col = df.columns[4]

                    # Suche nach dem Namen (Teilwortsuche, Groß-/Kleinschreibung ignorieren)
                    matches = df[df[name_col].astype(str).str.contains(name, case=False, na=False)]

                    if not matches.empty:
                        row = matches.iloc[0]

                        # Direkte Index-Zugriffe
                        sub_industry = str(row.iloc[2]).strip()  # Spalte C
                        focus_value = str(row.iloc[3]).strip()   # Spalte D
                        ric_value = str(row.iloc[4]).strip()     # Spalte E
                        name_value = str(row.iloc[1]).strip()    # Spalte B (Universe)

                        company = {
                            "Name": name_value,
                            "RIC": ric_value,
                            "Sub-Industry": sub_industry,
                            "Focus": focus_value
                        }
                        print(f"✅ GEFUNDEN: {company['Name']} ({company['RIC']})")
                        print(f"   Sub-Industry (Spalte C): '{company['Sub-Industry']}'")
                        print(f"   Focus (Spalte D): '{company['Focus']}'")
                        return company

                except Exception as e:
                    continue

        except Exception as e:
            continue

    print(f"❌ Name '{name}' nicht gefunden")
    return None
