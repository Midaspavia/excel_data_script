import os
import pandas as pd
from excel_kennzahlen import fetch_excel_kennzahlen_by_ric
from refinitiv_integration import get_refinitiv_kennzahlen_for_companies
import glob
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows

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
    """Hauptfunktion: Liest input_user.xlsx und erstellt output.xlsx mit Daten aus Excel-Dateien UND Refinitiv-Kennzahlen"""
    print("🚀 STARTE VERARBEITUNG (MEHRERE RICs/Namen + Excel + Refinitiv)...")

    try:
        # 1. Lese input_user.xlsx
        try:
            df_input = pd.read_excel("excel_data/input_user.xlsx")

            # Kennzahlen aus der ersten Zeile für alle verwenden
            first_row = df_input.iloc[0]
            excel_fields_raw = df_input["Kennzahlen aus Excel"].dropna().astype(str).str.strip().tolist()
            refinitiv_fields_raw = df_input["Kennzahlen aus Refinitiv"].dropna().astype(str).str.strip().tolist()

            # KORRIGIERT: Entferne Duplikate aus Kennzahlen-Listen (behält Reihenfolge bei)
            excel_fields = list(dict.fromkeys(excel_fields_raw))
            refinitiv_fields = list(dict.fromkeys(refinitiv_fields_raw))

            # Zeige Duplikat-Bereinigung an
            if len(excel_fields_raw) != len(excel_fields):
                print(f"🔧 Excel-Kennzahlen: {len(excel_fields_raw)} → {len(excel_fields)} (Duplikate entfernt)")
            if len(refinitiv_fields_raw) != len(refinitiv_fields):
                print(f"🔧 Refinitiv-Kennzahlen: {len(refinitiv_fields_raw)} → {len(refinitiv_fields)} (Duplikate entfernt)")

            # Filter-Einstellungen aus der ersten Zeile
            sub_industry_filter = str(first_row.get("Sub-Industry", "")).strip().upper()
            focus_filter = str(first_row.get("Focus", "")).strip().upper()

            # Bestimme Filter-Typ basierend auf X-Markierung
            # KORRIGIERT: Focus hat Priorität vor Sub-Industry
            if focus_filter == "X":
                is_focus = True
                filter_type = "Focus"
            elif sub_industry_filter == "X":
                is_focus = False
                filter_type = "Sub-Industry"
            else:
                # Fallback: Standard ist Sub-Industry
                is_focus = False
                filter_type = "Sub-Industry (Default)"

            print(f"🎯 Filter-Typ: {filter_type}")
            print(f"📋 Gewünschte Excel-Kennzahlen: {excel_fields}")
            print(f"📊 Gewünschte Refinitiv-Kennzahlen: {refinitiv_fields}")

            # 2. Sammle alle Input-Unternehmen aus allen Zeilen
            input_companies = []
            for index, row in df_input.iterrows():
                input_name = str(row.iloc[0] if len(row) > 0 else "").strip()  # Spalte A
                input_ric = str(row.iloc[1] if len(row) > 1 else "").strip()   # Spalte B

                # Überspringe leere Zeilen
                if not input_name and not input_ric:
                    continue
                if input_name.lower() in ["", "nan", "none"] and input_ric.lower() in ["", "nan", "none"]:
                    continue

                input_companies.append({
                    'name': input_name if input_name.lower() not in ["", "nan", "none"] else None,
                    'ric': input_ric if input_ric.lower() not in ["", "nan", "none"] else None,
                    'row_number': index + 1
                })

            print(f"📋 {len(input_companies)} Input-Unternehmen gefunden")

        except Exception as e:
            print(f"❌ Fehler beim Lesen von input_user.xlsx: {e}")
            return []

        # 3. Verarbeite jedes Input-Unternehmen
        all_results = []
        processed_groups = set()  # Verhindere Duplikate bei gleichen Gruppen

        for i, input_company in enumerate(input_companies, 1):
            print(f"\n🔍 VERARBEITE {i}/{len(input_companies)}: Zeile {input_company['row_number']}")

            # Bestimme Suchstrategie: RIC hat Priorität, dann Name
            if input_company['ric']:
                print(f"   🎯 RIC-Suche für '{input_company['ric']}'")
                start_company = find_company_by_ric(input_company['ric'])
            elif input_company['name']:
                if len(input_company['name']) < 4:
                    print(f"   ❌ Name '{input_company['name']}' zu kurz (min. 4 Zeichen)")
                    continue
                print(f"   🎯 Name-Suche für '{input_company['name']}'")
                start_company = find_company_by_name(input_company['name'])
            else:
                print("   ❌ Weder RIC noch Name vorhanden")
                continue

            if not start_company:
                print(f"   ❌ Unternehmen nicht gefunden!")
                continue

            print(f"   ✅ Gefunden: {start_company['Name']} ({start_company['RIC']})")

            # Bestimme Gruppe für Filterung
            if is_focus and start_company.get('Focus') and str(start_company.get('Focus')).strip().lower() not in ['', 'nan', 'none']:
                group_key = f"focus_{start_company['Focus']}"
                if group_key in processed_groups:
                    print(f"   ⏭️  Focus-Gruppe '{start_company['Focus']}' bereits verarbeitet")
                    continue
                processed_groups.add(group_key)
                print(f"   🎯 Focus-Modus: Suche nach Focus-Gruppe '{start_company['Focus']}'")
                peer_companies = find_companies_by_focus(start_company['Focus'])
            elif is_focus and (not start_company.get('Focus') or str(start_company.get('Focus')).strip().lower() in ['', 'nan', 'none']):
                # FALLBACK: Wenn Focus-Filter gewählt, aber Unternehmen hat keinen Focus-Wert
                group_key = f"subindustry_{start_company.get('Sub-Industry')}"
                if group_key in processed_groups:
                    print(f"   ⏭️  Sub-Industry '{start_company.get('Sub-Industry')}' bereits verarbeitet")
                    continue
                processed_groups.add(group_key)
                print(f"   ⚠️  Focus-Filter gewählt, aber Unternehmen hat keinen Focus-Wert")
                print(f"   🔄 Fallback auf Sub-Industry-Modus: Suche nach Sub-Industry '{start_company.get('Sub-Industry')}'")
                peer_companies = find_companies_by_sub_industry(start_company.get('Sub-Industry'))
            else:
                group_key = f"subindustry_{start_company.get('Sub-Industry')}"
                if group_key in processed_groups:
                    print(f"   ⏭️  Sub-Industry '{start_company.get('Sub-Industry')}' bereits verarbeitet")
                    continue
                processed_groups.add(group_key)
                print(f"   🎯 Sub-Industry-Modus: Suche nach Sub-Industry '{start_company.get('Sub-Industry')}'")
                peer_companies = find_companies_by_sub_industry(start_company.get('Sub-Industry'))

            print(f"   📊 {len(peer_companies)} Unternehmen der Gruppe gefunden")

            # 4. Hole Refinitiv-Kennzahlen für diese Gruppe (falls vorhanden)
            refinitiv_data = {}
            if refinitiv_fields:
                print(f"   🔄 Hole Refinitiv-Kennzahlen für {len(peer_companies)} Unternehmen...")
                refinitiv_data = get_refinitiv_kennzahlen_for_companies(peer_companies, refinitiv_fields)

            # 5. Sammle Kennzahlen für alle Unternehmen der Gruppe
            for j, company in enumerate(peer_companies, 1):
                print(f"     🏢 {j}/{len(peer_companies)}: {company['Name']} ({company['RIC']})")

                # Sammle Excel-Kennzahlen
                excel_kennzahlen = get_kennzahlen_for_company(company['RIC'], excel_fields)

                # Sammle Refinitiv-Kennzahlen
                refinitiv_kennzahlen = refinitiv_data.get(company['RIC'], {})

                # Erstelle Ergebnis
                result = {
                    "Name": company['Name'],
                    "RIC": company['RIC'],
                    "Sub-Industry": company.get('Sub-Industry', ''),
                    "Focus": company.get('Focus', ''),
                    "Input_Source": f"Zeile {input_company['row_number']}"  # Markiere Herkunft
                }
                result.update(excel_kennzahlen)
                result.update(refinitiv_kennzahlen)
                all_results.append(result)

                print(f"       ✅ {len(excel_kennzahlen)} Excel + {len(refinitiv_kennzahlen)} Refinitiv Kennzahlen")

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

            # 🏭 BERECHNE CONSUMER DISCRETIONARY SECTOR DURCHSCHNITTE FÜR REFINITIV-KENNZAHLEN
            if refinitiv_fields:
                print("\n🏭 BERECHNE CONSUMER DISCRETIONARY SECTOR DURCHSCHNITTE...")
                from refinitiv_integration import get_consumer_discretionary_sector_average
                sector_averages = get_consumer_discretionary_sector_average(refinitiv_fields)

                if sector_averages:
                    print(f"   🔍 DEBUG: Verfügbare Spalten: {list(df_output_with_averages.columns)}")
                    print(f"   🔍 DEBUG: Berechnete Durchschnitte: {list(sector_averages.keys())}")

                    # Füge Sector-Durchschnitt als neue Zeile hinzu
                    sector_avg_row = {
                        'Name': '🏭 Ø Consumer Discretionary Sector',
                        'RIC': f'AVG_GICS25_{len(sector_averages)}',
                        'Sub-Industry': '',
                        'Focus': '',
                        'Input_Source': 'Durchschnitt (GICS Sector 25)'
                    }

                    # Füge alle bestehenden Spalten mit leeren Werten hinzu
                    for col in df_output_with_averages.columns:
                        if col not in sector_avg_row:
                            sector_avg_row[col] = None

                    # Füge Refinitiv-Kennzahlen-Durchschnitte hinzu
                    for field, avg_value in sector_averages.items():
                        # Erstelle eine Liste möglicher Spaltennamen
                        possible_column_names = [
                            field,  # Original: "EBIT"
                            field.replace('TR.', ''),  # Ohne TR.: "EBIT"
                            field.upper(),  # Großbuchstaben: "EBIT"
                            field.lower(),  # Kleinbuchstaben: "ebit"
                        ]

                        # Wenn es ein TR.-Feld ist, füge auch TR.-Varianten hinzu
                        if field.startswith('TR.'):
                            clean_field = field.replace('TR.', '')
                            possible_column_names.extend([
                                clean_field,
                                clean_field.upper(),
                                clean_field.lower()
                            ])

                        found_column = None

                        # Suche nach exakter Übereinstimmung
                        for possible_name in possible_column_names:
                            if possible_name in df_output_with_averages.columns:
                                found_column = possible_name
                                print(f"   🎯 EXAKT gefunden: {field} → {possible_name}")
                                break

                        # Wenn nicht gefunden, suche nach Teilstring-Übereinstimmungen
                        if not found_column:
                            for col in df_output_with_averages.columns:
                                for possible_name in possible_column_names:
                                    if (possible_name.lower() in col.lower() or
                                        col.lower() in possible_name.lower()):
                                        found_column = col
                                        print(f"   🎯 TEILSTRING gefunden: {field} → {col}")
                                        break
                                if found_column:
                                    break

                        if found_column:
                            sector_avg_row[found_column] = avg_value
                            print(f"   📈 {found_column}: {avg_value:,.4f} (Sector-Durchschnitt)")
                        else:
                            # Fallback: Erstelle neue Spalte
                            clean_field = field.replace('TR.', '') if field.startswith('TR.') else field
                            sector_avg_row[clean_field] = avg_value
                            print(f"   ⚠️  NEUE SPALTE: {clean_field}: {avg_value:,.4f} (Sector-Durchschnitt)")

                    # Füge Sector-Durchschnitts-Zeile zum DataFrame hinzu
                    df_output_with_averages = pd.concat([df_output_with_averages, pd.DataFrame([sector_avg_row])], ignore_index=True)
                    print(f"   ✅ Consumer Discretionary Sector-Durchschnitt hinzugefügt")

                    print(f"   🔍 DEBUG: Finale Spalten: {list(df_output_with_averages.columns)}")
            # Erstelle schön formatierte Excel-Datei
            all_refinitiv_fields = []
            if refinitiv_fields:
                all_refinitiv_fields = refinitiv_fields
            create_beautiful_excel_output(df_output_with_averages, output_path, excel_fields)

            print(f"\n✅ SCHÖN FORMATIERTES OUTPUT GESPEICHERT: {output_path}")
            print(f"📊 {len(df_output_with_averages)} Zeilen insgesamt (inkl. Durchschnitte) mit {len(df_output_with_averages.columns)} Spalten")

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
                for field in all_refinitiv_fields:
                    actual_refinitiv_columns.append(field)

                # 2. Alle Spalten im result, die wie Refinitiv-Felder aussehen
                for key in result.keys():
                    # Überspringt Basis-Spalten und Excel-Kennzahlen
                    if key not in ['Name', 'RIC', 'Sub-Industry', 'Focus', 'Input_Source'] and key not in excel_fields:
                        # Prüft, ob es ein potentielles Refinitiv-Feld ist
                        if (key.startswith('TR.') or
                            any(key.upper() == ref_field.replace('TR.', '').upper() for ref_field in all_refinitiv_fields) or
                            key.upper() in ['EBIT', 'EBITDA', 'TOTALRETURN', 'TOTALASSETS']):  # Häufige Refinitiv-Felder
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
                        # Erweiterte Suche für ursprünglich angeforderte Felder
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
                        if field in all_refinitiv_fields:
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

def create_beautiful_excel_output(df, output_path, excel_fields):
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
    metadata_cell.value = f"📅 Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} | 📊 Companies: {len(df)} | 🔍 Analysis Type: {'Focus Group' if 'Focus' in df.columns else 'Sub-Industry'}"
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
                        'RIC': f'Branche-Ø ({len(df_sub_industry)} Unternehmen)',
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

    if len(focus_values) > 0:
        print("   🎯 Berechne Focus-Gruppen Durchschnitte...")
        focus_groups = df_numeric[df_numeric['Focus'] != ''].groupby('Focus')

        for focus, group in focus_groups:
            if len(group) > 1:  # Nur wenn mehr als 1 Unternehmen
                avg_row = {
                    'Name': f'🎯 Ø {focus}',
                    'RIC': f'AVG_FOC_{len(group)}',
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
