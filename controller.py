import os
import pandas as pd
from excel_kennzahlen import fetch_excel_kennzahlen_by_ric, fetch_excel_kennzahlen_by_ric_filtered, clear_excel_cache
from refinitiv_integration import get_refinitiv_kennzahlen_for_companies
import glob
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import warnings

# KORRIGIERT: Unterdrücke openpyxl Warnungen über Datums-Formatierung
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

DATA_DIR = "excel_data/data"

def clean_refinitiv_field_name(field_name):
    """
    Entfernt TR. aus Refinitiv-Feldnamen, behält aber Period-Information bei
    Beispiel: TR.EBIT(Period=FY-1) → EBIT(Period=FY-1)
    """
    if field_name.startswith('TR.'):
        return field_name.replace('TR.', '')
    return field_name

def cleanup_temp_files():
    """Bereinigt temporäre Excel-Dateien (~$*.xlsx) nach der Ausführung"""
    print("🧹 BEREINIGE TEMPORÄRE DATEIEN...")

    directories = ["excel_data/", "excel_data/data/", "."]
    deleted_count = 0

    for directory in directories:
        if os.path.exists(directory):
            temp_files = glob.glob(os.path.join(directory, "~$*.xlsx"))
            for temp_file in temp_files:
                try:
                    os.remove(temp_file)
                    deleted_count += 1
                except Exception as e:
                    pass

    if deleted_count > 0:
        print(f"✅ {deleted_count} temporäre Dateien bereinigt")

def process_companies():
    """OPTIMIERTE Hauptfunktion: Schnellere Verarbeitung mit Caching"""
    start_time = time.time()
    print("🚀 STARTE OPTIMIERTE VERARBEITUNG...")

    try:
        # 1. Lese input_user.xlsx (SCHNELL)
        print("📖 Lese input_user.xlsx...")
        df_input = pd.read_excel("excel_data/input_user.xlsx")

        # Kennzahlen aus der ersten Zeile
        first_row = df_input.iloc[0]
        excel_fields = list(dict.fromkeys(df_input["Kennzahlen aus Excel"].dropna().astype(str).str.strip().tolist()))
        refinitiv_fields = list(dict.fromkeys(df_input["Kennzahlen aus Refinitiv"].dropna().astype(str).str.strip().tolist()))

        # Filter-Einstellungen
        sub_industry_filter = str(first_row.get("Sub-Industry", "")).strip().upper()
        focus_filter = str(first_row.get("Focus", "")).strip().upper()

        if focus_filter == "X":
            is_focus = True
            filter_type = "Focus"
        elif sub_industry_filter == "X":
            is_focus = False
            filter_type = "Sub-Industry"
        else:
            is_focus = False
            filter_type = "Sub-Industry (Default)"

        print(f"🎯 Filter: {filter_type}")
        print(f"📋 Excel-Kennzahlen: {len(excel_fields)}")
        print(f"📊 Refinitiv-Kennzahlen: {len(refinitiv_fields)}")

        # 2. SAMMLE ALLE INPUT-UNTERNEHMEN (SCHNELL)
        input_companies = []
        all_gics_sectors = set()

        for index, row in df_input.iterrows():
            input_name = str(row.iloc[0] if len(row) > 0 else "").strip()
            input_ric = str(row.iloc[1] if len(row) > 1 else "").strip()

            gics_sector = ""
            if "GICS Sector" in df_input.columns:
                gics_sector = str(row.get("GICS Sector", "")).strip()

            # Überspringe leere Zeilen
            if not input_name and not input_ric:
                continue
            if input_name.lower() in ["", "nan", "none"] and input_ric.lower() in ["", "nan", "none"]:
                continue

            if gics_sector and gics_sector.lower() not in ["", "nan", "none"]:
                all_gics_sectors.add(gics_sector)

            input_companies.append({
                'name': input_name if input_name.lower() not in ["", "nan", "none"] else None,
                'ric': input_ric if input_ric.lower() not in ["", "nan", "none"] else None,
                'gics_sector': gics_sector if gics_sector.lower() not in ["", "nan", "none"] else None,
                'row_number': index + 1
            })

        print(f"📋 {len(input_companies)} Input-Unternehmen")
        print(f"🏭 GICS Sektoren: {sorted(all_gics_sectors)}")

        # 3. OPTIMIERTE VERARBEITUNG
        all_results = []
        processed_groups = set()
        sector_companies = {}

        # Leere Cache für frische Daten
        clear_excel_cache()

        for i, input_company in enumerate(input_companies, 1):
            print(f"\n🔍 {i}/{len(input_companies)}: Zeile {input_company['row_number']}")

            # Bestimme Suchstrategie
            start_company = None
            if input_company['ric']:
                print(f"   🎯 RIC: {input_company['ric']}")
                start_company = find_company_by_ric(input_company['ric'])
            elif input_company['name']:
                if len(input_company['name']) >= 4:
                    print(f"   🎯 Name: {input_company['name']}")
                    start_company = find_company_by_name(input_company['name'])
                else:
                    print(f"   ❌ Name zu kurz: {input_company['name']}")
                    continue
            else:
                print("   ❌ Weder RIC noch Name")
                continue

            if not start_company:
                print(f"   ❌ Nicht gefunden!")
                continue

            print(f"   ✅ {start_company['Name']} ({start_company['RIC']})")

            # WIEDERHERGESTELLT: Erstelle eindeutigen Gruppenschlüssel für Peer-Group-Verarbeitung
            group_key = f"{start_company.get('Sub-Industry', 'Unknown')}_{start_company.get('Focus', 'Unknown')}"

            # Überspringe, wenn diese Gruppe bereits verarbeitet wurde
            if group_key in processed_groups:
                print(f"   ⏭️  Gruppe '{group_key}' bereits verarbeitet - überspringe")
                continue

            processed_groups.add(group_key)

            # WIEDERHERGESTELLT: Finde alle Peer-Unternehmen derselben Sub-Industry/Focus-Gruppe
            print(f"   🔍 Suche Peer-Group für: Sub-Industry='{start_company.get('Sub-Industry', '')}', Focus='{start_company.get('Focus', '')}'")

            if is_focus:
                peer_companies = find_companies_by_focus(start_company.get('Focus', ''))
            else:
                peer_companies = find_companies_by_sub_industry(start_company.get('Sub-Industry', ''))

            # KORRIGIERT: Füge das Input-Unternehmen selbst hinzu, falls es nicht in der Peer-Liste steht
            input_company_in_peers = any(comp['RIC'] == start_company['RIC'] for comp in peer_companies)
            if not input_company_in_peers:
                print(f"   📌 Input-Unternehmen {start_company['Name']} ({start_company['RIC']}) nicht in Peer-Liste gefunden - füge hinzu")
                peer_companies.insert(0, start_company)  # Füge am Anfang hinzu

            if not peer_companies:
                print(f"   ❌ Keine Peer-Unternehmen gefunden!")
                continue

            print(f"   📊 {len(peer_companies)} Peer-Unternehmen gefunden (inkl. Input-Unternehmen)")

            # Sammle für Refinitiv-Durchschnitte
            sector_key = input_company['gics_sector'] or 'Unknown'
            if sector_key not in sector_companies:
                sector_companies[sector_key] = []
            sector_companies[sector_key].extend(peer_companies)

            # 4. HOLE REFINITIV-DATEN (für alle Peer-Unternehmen)
            refinitiv_data = {}
            if refinitiv_fields:
                print(f"   🔄 Refinitiv für {len(peer_companies)} Unternehmen...")
                refinitiv_data = get_refinitiv_kennzahlen_for_companies(peer_companies, refinitiv_fields)

            # 5. SAMMLE KENNZAHLEN (mit GICS Sector-Filter)
            for j, company in enumerate(peer_companies, 1):
                print(f"     🏢 {j}/{len(peer_companies)}: {company['Name']}")

                # Excel-Kennzahlen mit GICS Sector-Filter (CACHED!)
                gics_sectors_for_search = [input_company['gics_sector']] if input_company['gics_sector'] else None
                excel_kennzahlen = get_kennzahlen_for_company_filtered(company['RIC'], excel_fields, gics_sectors_for_search)

                # Refinitiv-Kennzahlen
                refinitiv_kennzahlen = refinitiv_data.get(company['RIC'], {})

                # Erstelle Ergebnis
                result = {
                    "Name": company['Name'],
                    "RIC": company['RIC'],
                    "Sub-Industry": company.get('Sub-Industry', ''),
                    "Focus": company.get('Focus', ''),
                    "GICS_Sector": input_company['gics_sector'] or '',
                    "Input_Source": f"Zeile {input_company['row_number']}"
                }
                result.update(excel_kennzahlen)
                result.update(refinitiv_kennzahlen)
                all_results.append(result)

                print(f"       ✅ {len(excel_kennzahlen)} Excel + {len(refinitiv_kennzahlen)} Refinitiv")

        # 6. Speichere Output mit schönem Design
        if all_results:
            output_path = "excel_data/output.xlsx"
            df_output = pd.DataFrame(all_results)

            print(f"\n📊 INSGESAMT {len(all_results)} UNTERNEHMEN VERARBEITET")
            print("💾 Speichere in output.xlsx...")

            # KORRIGIERT: Stelle sicher, dass Output-Verzeichnis existiert
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                print(f"📁 Erstelle fehlendes Verzeichnis: {output_dir}")
                os.makedirs(output_dir, exist_ok=True)

            # 🔢 BERECHNE DURCHSCHNITTE FÜR EXCEL-KENNZAHLEN
            print("\n🔢 BERECHNE DURCHSCHNITTE FÜR EXCEL-KENNZAHLEN...")
            df_output_with_averages = calculate_excel_averages(df_output, excel_fields)

            # 🏭 NEU: BERECHNE SEKTOR-SPEZIFISCHE DURCHSCHNITTE FÜR REFINITIV-KENNZAHLEN
            if refinitiv_fields:
                print(f"\n🏭 BERECHNE REFINITIV DURCHSCHNITTE FÜR {len(sector_companies)} SEKTOREN...")

                for sector_name, companies_list in sector_companies.items():
                    if not companies_list or sector_name == 'Unknown':
                        continue

                    # Entferne Duplikate basierend auf RIC
                    unique_companies = {comp['RIC']: comp for comp in companies_list}.values()
                    unique_companies = list(unique_companies)

                    print(f"   🏭 Sektor '{sector_name}': {len(unique_companies)} Unternehmen")

                    from refinitiv_integration import get_sector_average_by_companies
                    sector_averages = get_sector_average_by_companies(unique_companies, refinitiv_fields)

                    if sector_averages:
                        # Füge Sektor-Durchschnitt als neue Zeile hinzu
                        sector_avg_row = {
                            'Name': f'🏭 Ø {sector_name} Sector',
                            'RIC': '',
                            'Sub-Industry': '',
                            'Focus': '',
                            'GICS_Sector': sector_name,
                            'Input_Source': f'Durchschnitt ({sector_name} Sector)'
                        }

                        # Füge alle bestehenden Spalten mit leeren Werten hinzu
                        for col in df_output_with_averages.columns:
                            if col not in sector_avg_row:
                                sector_avg_row[col] = None

                        # Füge Refinitiv-Kennzahlen-Durchschnitte hinzu
                        for field, avg_value in sector_averages.items():
                            cleaned_field_name = clean_refinitiv_field_name(field)
                            sector_avg_row[cleaned_field_name] = avg_value
                            print(f"     📊 {cleaned_field_name}: {avg_value}")

                        # Erstelle DataFrame für neue Zeile und füge hinzu
                        sector_avg_df = pd.DataFrame([sector_avg_row])
                        df_output_with_averages = pd.concat([df_output_with_averages, sector_avg_df], ignore_index=True)

            # KORRIGIERT: Filtere Output-DataFrame, um nur angeforderte Kennzahlen zu behalten
            print(f"\n🔍 FILTERE OUTPUT AUF NUR ANGEFORDERTE KENNZAHLEN...")

            # Basis-Spalten die immer beibehalten werden
            base_columns = ['Name', 'RIC', 'Sub-Industry', 'Focus', 'Input_Source', 'GICS_Sector']

            # Sammle alle erlaubten Spalten
            allowed_columns = base_columns.copy()
            allowed_columns.extend(excel_fields)  # Angeforderte Excel-Kennzahlen

            # Füge Refinitiv-Kennzahlen hinzu (mit und ohne TR. Präfix)
            for ref_field in refinitiv_fields:
                allowed_columns.append(ref_field)  # Original (z.B. TR.EBIT)
                clean_field = clean_refinitiv_field_name(ref_field)
                allowed_columns.append(clean_field)  # Ohne TR. (z.B. EBIT)

            # Filtere DataFrame auf nur erlaubte Spalten
            existing_allowed_columns = [col for col in allowed_columns if col in df_output_with_averages.columns]
            df_output_cleaned = df_output_with_averages[existing_allowed_columns].copy()

            print(f"   📊 Ursprüngliche Spalten: {len(df_output_with_averages.columns)}")
            print(f"   ✅ Gefilterte Spalten: {len(df_output_cleaned.columns)}")
            print(f"   📋 Behaltene Spalten: {list(df_output_cleaned.columns)}")

            # Erstelle schön formatierte Excel-Datei
            create_beautiful_excel_output(df_output_cleaned, output_path, excel_fields, len(all_results))

            print(f"\n✅ SCHÖN FORMATIERTES OUTPUT GESPEICHERT: {output_path}")
            print(f"📊 {len(all_results)} Unternehmen + {len(df_output_cleaned) - len(all_results)} Durchschnittswerte = {len(df_output_cleaned)} Zeilen insgesamt mit {len(df_output_cleaned.columns)} Spalten")

            # Zeige Übersicht
            print(f"\n📋 ERGEBNIS-ÜBERSICHT:")
            for i, result in enumerate(all_results, 1):
                print(f"\n{i}. {result['Name']} ({result['RIC']}) - {result.get('Input_Source', '')}")
                print(f"   Sub-Industry: {result.get('Sub-Industry', 'N/A')}")
                print(f"   Focus: {result.get('Focus', 'N/A')}")

                # Zeige alle Excel-Kennzahlen
                for field in excel_fields:
                    value = result.get(field, 'N/A')
                    if value != 'N/A' and pd.notna(value):
                        print(f"   [Excel] {field}: {value}")
                    else:
                        print(f"   [Excel] {field}: ❌ Nicht gefunden")

                # KORRIGIERT: Finde ALLE Refinitiv-Spalten im DataFrame (nicht nur die ursprünglich angeforderten)
                # Sammle alle Refinitiv-relevanten Spalten aus dem tatsächlichen DataFrame
                actual_refinitiv_columns = []

                # 1. Alle ursprünglich angeforderten Refinitiv-Felder
                for field in refinitiv_fields:
                    actual_refinitiv_columns.append(field)

                # 2. Alle Spalten im result, die wie Refinitiv-Felder aussehen
                for key in result.keys():
                    # Überspringt Basis-Spalten und Excel-Kennzahlen
                    if key not in ['Name', 'RIC', 'Sub-Industry', 'Focus', 'Input_Source'] and key not in excel_fields:
                        # Prüft, ob es ein potentielles Refinitiv-Feld ist
                        if (key.startswith('TR.') or
                            any(key.upper() == ref_field.replace('TR.', '').upper() for ref_field in refinitiv_fields)):
                            if key not in actual_refinitiv_columns:
                                actual_refinitiv_columns.append(key)

                # Entferne Duplikate und behalte Reihenfolge
                actual_refinitiv_columns = list(dict.fromkeys(actual_refinitiv_columns))

                # Zeige alle gefundenen Refinitiv-Kennzahlen
                for field in actual_refinitiv_columns:
                    # Suche nach der Spalte im Result
                    found_value = None
                    found_key = None

                    # Direkte Suche nach dem Feld
                    if field in result:
                        found_value = result[field]
                        found_key = field
                    else:
                        # Erweiterte Suche für ursprünglich angeforderten Felder
                        cleaned_field = field.replace("TR.", "") if field.startswith("TR.") else field
                        if cleaned_field in result:
                            found_value = result[cleaned_field]
                            found_key = cleaned_field
                        else:
                            # Fuzzy-Suche nach ähnlichen Feldern
                            for key, value in result.items():
                                if (field.lower() in key.lower() or
                                    cleaned_field.lower() in key.lower() or
                                    key.lower() in field.lower()):
                                    found_value = value
                                    found_key = key
                                    break

                    if found_value is not None and pd.notna(found_value) and str(found_value).strip() != '':
                        # Bestimme Label für Ausgabe
                        if field in refinitiv_fields:
                            display_label = f"[Refinitiv] {field}"
                        else:
                            display_label = f"[Refinitiv*] {field}"  # * für neu erstellte Spalten

                        if found_key != field:
                            print(f"   {display_label} (als '{found_key}'): {found_value}")
                        else:
                            print(f"   {display_label}: {found_value}")
                    else:
                        print(f"   [Refinitiv] {field}: ❌ Nicht gefunden")

            # Zeige Consumer Discretionary Sector-Durchschnitte für Refinitiv-Kennzahlen
            if refinitiv_fields and sector_averages:
                print(f"\n🏭 CONSUMER DISCRETIONARY SECTOR-DURCHSCHNITTE (REFINITIV):")
                for field, avg_value in sector_averages.items():
                    # Finde den ursprünglichen Feldnamen
                    original_field = None
                    for ref_field in refinitiv_fields:
                        clean_ref = ref_field.replace('TR.', '') if ref_field.startswith('TR.') else ref_field
                        if field == clean_ref or field.lower() == clean_ref.lower():
                            original_field = ref_field
                            break

                    if original_field:
                        print(f"   📈 {original_field}: {avg_value:,.4f} (Sektor-Durchschnitt GICS 25)")
                    else:
                        print(f"   📈 {field}: {avg_value:,.4f} (Sektor-Durchschnitt GICS 25)")

        return all_results

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

                        # KORRIGIERT: Robuste Name-Extraktion aus Spalte A (Holding) oder B (Universe)
                        name_value = "Unknown"

                        # Versuche zuerst Spalte A (Holding)
                        if len(df.columns) > 0:
                            holding_name = str(row.iloc[0]).strip()
                            if holding_name and holding_name.lower() not in ['nan', 'none', ''] and len(holding_name) > 2:
                                name_value = holding_name

                        # Falls Spalte A leer oder ungültig, versuche Spalte B (Universe)
                        if name_value == "Unknown" and len(df.columns) > 1:
                            universe_name = str(row.iloc[1]).strip()
                            if universe_name and universe_name.lower() not in ['nan', 'none', ''] and len(universe_name) > 2:
                                name_value = universe_name

                        # Falls beide Spalten leer, verwende generischen Fallback-Namen
                        if name_value == "Unknown":
                            name_value = f"Company_{ric_value}"

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

    # Prüfe 4-Zeichen-Regel
    if len(name) < 4:
        print(f"❌ Name '{name}' zu kurz (mindestens 4 Zeichen erforderlich)")
        return None

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

                    # KORRIGIERT: Suche sowohl in Spalte A (Holding) als auch Spalte B (Universe)
                    holding_col = df.columns[0]  # Spalte A (Holding)
                    universe_col = df.columns[1]  # Spalte B (Universe)

                    # Suche in beiden Spalten
                    holding_matches = df[df[holding_col].astype(str).str.contains(name, case=False, na=False)]
                    universe_matches = df[df[universe_col].astype(str).str.contains(name, case=False, na=False)]

                    # Kombiniere beide Ergebnisse, bevorzuge aber Holding-Treffer
                    matches = holding_matches if not holding_matches.empty else universe_matches

                    if not matches.empty:
                        row = matches.iloc[0]

                        # Direkte Index-Zugriffe
                        holding_value = str(row.iloc[0]).strip()   # Spalte A (Holding)
                        universe_value = str(row.iloc[1]).strip()  # Spalte B (Universe)
                        sub_industry = str(row.iloc[2]).strip()    # Spalte C
                        focus_value = str(row.iloc[3]).strip()     # Spalte D
                        ric_value = str(row.iloc[4]).strip()       # Spalte E

                        # Bestimme den Namen (Holding hat Priorität)
                        if holding_value and holding_value != 'nan' and len(holding_value.strip()) > 2:
                            name_value = holding_value
                            found_in = "Holding"
                        else:
                            name_value = universe_value
                            found_in = "Universe"

                        company = {
                            "Name": name_value,
                            "RIC": ric_value,
                            "Sub-Industry": sub_industry,
                            "Focus": focus_value
                        }
                        print(f"✅ GEFUNDEN: {company['Name']} ({company['RIC']}) in {found_in}-Spalte")
                        print(f"   Sub-Industry (Spalte C): '{company['Sub-Industry']}'")
                        print(f"   Focus (Spalte D): '{company['Focus']}'")
                        return company

                except Exception as e:
                    continue

        except Exception as e:
            continue

    print(f"❌ Name '{name}' nicht gefunden")
    return None

def create_beautiful_excel_output(df, output_path, excel_fields, actual_company_count=None):
    """Erstellt eine wunderschön formatierte Excel-Datei mit professionellem Design"""
    print("🎨 ERSTELLE SCHÖNES EXCEL-DESIGN...")

    # Speichere DataFrame als Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Financial Analysis', index=False)

    # Lade Workbook für Formatierung
    wb = load_workbook(output_path)
    ws = wb['Financial Analysis']

    # 🎨 FARB-SCHEMA (NUR ALTERNIERENDE FARBEN)
    header_fill = PatternFill(start_color="1f4e79", end_color="1f4e79", fill_type="solid")  # Dunkles Blau für Header
    alternating_fill = PatternFill(start_color="f8f9fa", end_color="f8f9fa", fill_type="solid")  # Sehr helles Grau
    white_fill = PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")  # Weiß

    # 📝 SCHRIFT-STILE
    header_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    company_font = Font(name="Calibri", size=11, bold=True, color="1f4e79")
    data_font = Font(name="Calibri", size=10, color="2f2f2f")

    # 📐 ALIGNMENT
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_alignment = Alignment(horizontal="right", vertical="center")

    # 🔳 BORDERS
    thin_border = Border(
        left=Side(style="thin", color="b0b0b0"),
        right=Side(style="thin", color="b0b0b0"),
        top=Side(style="thin", color="b0b0b0"),
        bottom=Side(style="thin", color="b0b0b0")
    )

    thick_border = Border(
        left=Side(style="medium", color="1f4e79"),
        right=Side(style="medium", color="1f4e79"),
        top=Side(style="medium", color="1f4e79"),
        bottom=Side(style="medium", color="1f4e79")
    )

    # 1️⃣ HEADER-ZEILE FORMATIEREN
    print("  🎯 Formatiere Header...")

    # Berechne dynamische Header-Höhe basierend auf Zeilenumbrüchen
    max_lines = 1
    for col_num, cell in enumerate(ws[1], 1):
        col_name = df.columns[col_num - 1]

        # Zähle Zeilenumbrüche in Spalten-Namen
        line_count = col_name.count('\n') + 1
        max_lines = max(max_lines, line_count)

        # Formatiere Header-Zelle
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = thick_border

    # Setze dynamische Header-Höhe (15 Pixel pro Zeile + 10 Pixel Padding)
    dynamic_header_height = max_lines * 15 + 10
    ws.row_dimensions[1].height = max(25, dynamic_header_height)  # Minimum 25 Pixel

    print(f"  📏 Header-Höhe: {ws.row_dimensions[1].height}px ({max_lines} Zeilen)")

    # 2️⃣ SPALTEN-KATEGORIEN BESTIMMEN
    company_cols = ['Name', 'RIC']  # Unternehmensdaten
    category_cols = ['Sub-Industry', 'Focus']  # Kategorien
    metric_cols = [col for col in df.columns if col not in company_cols + category_cols]  # Kennzahlen

    # 3️⃣ DATENZEILEN FORMATIEREN
    print("  🎯 Formatiere Datenzeilen...")
    for row_num in range(2, len(df) + 2):
        # Alternierend gefärbte Zeilen für bessere Lesbarkeit
        is_even = (row_num % 2 == 0)
        row_fill = alternating_fill if is_even else white_fill

        for col_num, cell in enumerate(ws[row_num], 1):
            col_name = df.columns[col_num - 1]

            # Basis-Formatierung
            cell.border = thin_border
            cell.font = data_font
            cell.fill = row_fill  # Nur alternierende Farben

            # Spezielle Formatierung je Spalten-Typ
            if col_name in company_cols:
                cell.font = company_font if col_name == 'Name' else Font(name="Calibri", size=10, bold=True, color="1f4e79")
                cell.alignment = left_alignment
            elif col_name in category_cols:
                cell.alignment = center_alignment
            elif col_name in metric_cols:
                cell.alignment = right_alignment

                # Formatiere Zahlen schön
                if cell.value and str(cell.value).replace('.', '').replace('-', '').isdigit():
                    try:
                        num_val = float(cell.value)
                        if abs(num_val) >= 1:
                            cell.number_format = '#,##0.00'  # Mit Tausender-Trennzeichen
                        else:
                            cell.number_format = '0.0000'    # Mehr Dezimalstellen für kleine Zahlen
                    except:
                        pass
            else:
                cell.alignment = left_alignment

        # Zeilenhöhe optimieren
        ws.row_dimensions[row_num].height = 20

    # 4️⃣ SPALTENBREITEN AUTOMATISCH ANPASSEN
    print("  🎯 Optimiere Spaltenbreiten...")
    for col_num, column in enumerate(ws.columns, 1):
        col_letter = ws.cell(row=1, column=col_num).column_letter
        col_name = df.columns[col_num - 1]

        # Berechne optimale Breite basierend auf Inhalt
        max_length = 0
        for cell in column:
            if cell.value:
                cell_length = len(str(cell.value))
                max_length = max(max_length, cell_length)

        # Setze minimale und maximale Breiten je nach Spalten-Typ
        if col_name == 'Name':
            width = min(max(max_length + 2, 25), 40)  # Name: 25-40 Zeichen
        elif col_name == 'RIC':
            width = min(max(max_length + 2, 12), 15)  # RIC: 12-15 Zeichen
        elif col_name in category_cols:
            width = min(max(max_length + 2, 15), 25)  # Kategorien: 15-25 Zeichen
        elif col_name in metric_cols:
            width = min(max(max_length + 2, 12), 18)  # Kennzahlen: 12-18 Zeichen
        else:
            width = min(max(max_length + 2, 10), 30)  # Standard: 10-30 Zeichen

        ws.column_dimensions[col_letter].width = width

    # 5️⃣ CONDITIONAL FORMATTING ENTFERNT
    # (Keine Farbvergleiche mehr für Kennzahlen-Spalten)

    # 6️⃣ TITEL UND METADATA HINZUFÜGEN
    print("  🎯 Füge Titel hinzu...")
    # Neue Zeile oben einfügen für Titel
    ws.insert_rows(1)

    # Titel erstellen
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"📊 FINANCIAL ANALYSIS REPORT - {df['Sub-Industry'].iloc[0] if len(df) > 0 else 'PEER ANALYSIS'}"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="1f4e79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Titel über alle Spalten mergen
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))

    # Titel-Zeile höher machen
    ws.row_dimensions[1].height = 35

    # Titel-Hintergrund
    for col in range(1, len(df.columns) + 1):
        ws.cell(row=1, column=col).fill = PatternFill(start_color="f2f2f2", end_color="f2f2f2", fill_type="solid")
        ws.cell(row=1, column=col).border = thick_border

    # KORRIGIERT: Header-Höhe nach Titel-Einfügung neu setzen (jetzt Zeile 2)
    ws.row_dimensions[2].height = max(25, dynamic_header_height)
    print(f"  📏 Header-Höhe korrigiert: {ws.row_dimensions[2].height}px ({max_lines} Zeilen) - Zeile 2")

    # 7️⃣ FREEZE PANES FÜR BESSERE NAVIGATION
    ws.freeze_panes = "A3"  # Freeze Header und Titel

    # 8️⃣ METADATA AM ENDE HINZUFÜGEN
    last_row = len(df) + 3
    metadata_cell = ws.cell(row=last_row, column=1)

    # KORRIGIERT: Verwende die tatsächliche Anzahl der Unternehmen (ohne Durchschnitte)
    company_count = actual_company_count if actual_company_count is not None else len(df)

    metadata_cell.value = f"📅 Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} | 📊 Companies: {company_count}"
    metadata_cell.font = Font(name="Calibri", size=9, italic=True, color="666666")
    metadata_cell.alignment = left_alignment

    # Metadata über mehrere spalten mergen
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=min(6, len(df.columns)))

    # Speichere formatierte Datei
    wb.save(output_path)

    print("  ✨ Excel-Formatierung abgeschlossen!")
    print(f"  📋 Titel: {title_cell.value}")
    print(f"  📊 {len(df)} Unternehmen formatiert")
    print(f"  📈 {len(metric_cols)} Kennzahlen mit Conditional Formatting")
    print(f"  💾 Datei gespeichert: {output_path}")

def calculate_excel_averages(df, excel_fields):
    """Berechnet die Durchschnitte für Excel-Kennzahlen nach Sub-Industry und Focus-Gruppen"""
    print("🔢 BERECHNE DURCHSCHNITTE FÜR EXCEL-KENNZAHLEN...")

    # Filtere nur die Spalten, die mit Excel-Kennzahlen gefüllt sind
    excel_columns = [field for field in excel_fields if field in df.columns]

    if not excel_columns:
        print("⚠️ Keine Excel-Kennzahlen gefunden, überspringe Durchschnittsberechnung")
        return df

    print(f"📊 Berechne Durchschnitte für: {excel_columns}")

    # Konvertiere Excel-Kennzahlen zu numerischen Werten
    df_numeric = df.copy()
    for col in excel_columns:
        df_numeric[col] = pd.to_numeric(df_numeric[col], errors='coerce')

    # 1. SUB-INDUSTRY DURCHSCHNITTE (ALLE UNTERNEHMEN AUS DEN EXCEL-DATEIEN)
    print("   🏭 Berechne Sub-Industry Durchschnitte (alle verfügbaren Unternehmen)...")

    # Hole alle eindeutigen Sub-Industries aus dem Output
    unique_sub_industries = df_numeric['Sub-Industry'].dropna().unique()

    for sub_industry in unique_sub_industries:
        if sub_industry and sub_industry.strip():
            print(f"     🔍 Suche alle Unternehmen der Sub-Industry: '{sub_industry}'")

            # Hole ALLE Unternehmen dieser Sub-Industry aus den Excel-Dateien
            all_companies_in_sub_industry = find_companies_by_sub_industry(sub_industry)

            if len(all_companies_in_sub_industry) > 1:
                # Sammle Excel-Kennzahlen für ALLE Unternehmen der Sub-Industry
                all_sub_industry_data = []
                print(f"       📋 Verarbeite {len(all_companies_in_sub_industry)} Unternehmen...")

                for i, company in enumerate(all_companies_in_sub_industry, 1):
                    if i <= 5 or i % 20 == 0:  # Zeige nur jeden 20. nach den ersten 5
                        print(f"         {i}/{len(all_companies_in_sub_industry)}: {company['Name']}")

                    company_data = get_kennzahlen_for_company(company['RIC'], excel_columns)
                    if company_data:
                        # Füge Basis-Informationen hinzu
                        company_data.update({
                            'Name': company['Name'],
                            'RIC': company['RIC'],
                            'Sub-Industry': company.get('Sub-Industry', ''),
                            'Focus': company.get('Focus', '')
                        })
                        all_sub_industry_data.append(company_data)

                if all_sub_industry_data:
                    # Erstelle DataFrame für alle Sub-Industry Unternehmen
                    df_sub_industry = pd.DataFrame(all_sub_industry_data)

                    # Konvertiere zu numerischen Werten
                    for col in excel_columns:
                        df_sub_industry[col] = pd.to_numeric(df_sub_industry[col], errors='coerce')

                    # Berechne Durchschnitte
                    avg_row = {
                        'Name': f'💼 Ø {sub_industry}',
                        'RIC': '',
                        'Sub-Industry': sub_industry,
                        'Focus': '',
                        'Input_Source': 'Durchschnitt (Branche)'
                    }

                    for col in excel_columns:
                        valid_values = df_sub_industry[col].dropna()
                        if len(valid_values) > 0:
                            avg_row[col] = valid_values.mean()
                            print(f"       📈 {col}: {avg_row[col]:.4f} (aus {len(valid_values)} von {len(df_sub_industry)} Unternehmen)")
                        else:
                            avg_row[col] = None

                    # Füge Durchschnitts-Zeile hinzu
                    df = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)
                    print(f"       ✅ Sub-Industry Durchschnitt hinzugefügt: {sub_industry} ({len(df_sub_industry)} Unternehmen)")

    # 2. FOCUS-GRUPPEN DURCHSCHNITTE (nur wenn Focus-Werte vorhanden)
    focus_values = df_numeric['Focus'].dropna()
    focus_values = focus_values[focus_values != '']

    # KORRIGIERT: Filtere auch "nan" und "None" Strings heraus
    focus_values = focus_values[focus_values.astype(str).str.lower() != 'nan']
    focus_values = focus_values[focus_values.astype(str).str.lower() != 'none']

    if len(focus_values) > 0:
        print("   🎯 Berechne Focus-Gruppen Durchschnitte...")

        # KORRIGIERT: Filtere DataFrame vor Gruppierung
        valid_focus_df = df_numeric[
            (df_numeric['Focus'] != '') &
            (df_numeric['Focus'].notna()) &
            (df_numeric['Focus'].astype(str).str.lower() != 'nan') &
            (df_numeric['Focus'].astype(str).str.lower() != 'none')
        ]

        focus_groups = valid_focus_df.groupby('Focus')

        for focus, group in focus_groups:
            if len(group) > 1:  # Nur wenn mehr als 1 Unternehmen
                print(f"     🎯 Focus-Gruppe '{focus}': {len(group)} Unternehmen")

                avg_row = {
                    'Name': f'🎯 Ø {focus}',
                    'RIC': '',
                    'Sub-Industry': '',
                    'Focus': focus,
                    'Input_Source': 'Durchschnitt'
                }

                # Berechne Durchschnitt für jede Excel-Kennzahl
                for col in excel_columns:
                    valid_values = group[col].dropna()
                    if len(valid_values) > 0:
                        avg_row[col] = valid_values.mean()
                        print(f"     📈 {col}: {avg_row[col]:.4f} (aus {len(valid_values)} Werten)")
                    else:
                        avg_row[col] = None

                # Füge Durchschnitts-Zeile hinzu
                df = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)
                print(f"     ✅ Focus-Gruppen Durchschnitt hinzugefügt: {focus}")
    else:
        print("   ⚠️ Keine Focus-Gruppen gefunden, überspringe Focus-Durchschnitte")

    print(f"✅ Durchschnittsberechnung abgeschlossen")
    return df

def get_kennzahlen_for_company_filtered(ric, fields, gics_sectors=None):
    """
    Sammelt alle gewünschten Kennzahlen für ein Unternehmen basierend auf RIC
    mit GICS Sector-Filter für Excel-Dateien

    Args:
        ric: Reuters Instrument Code
        fields: Liste der gewünschten Kennzahlen
        gics_sectors: Liste von GICS Sektoren für Datei-Filterung
    """
    return fetch_excel_kennzahlen_by_ric_filtered(ric, fields, gics_sectors)
