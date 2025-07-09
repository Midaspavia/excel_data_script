import os
import pandas as pd
from excel_kennzahlen import fetch_excel_kennzahlen_by_ric
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

        # 5. Speichere Output mit schönem Design
        if results:
            output_path = "excel_data/output.xlsx"
            df_output = pd.DataFrame(results)

            # Erstelle schön formatierte Excel-Datei
            create_beautiful_excel_output(df_output, output_path, excel_fields)

            print(f"\n✅ SCHÖN FORMATIERTES OUTPUT GESPEICHERT: {output_path}")
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

def create_beautiful_excel_output(df, output_path, excel_fields):
    """Erstellt eine wunderschön formatierte Excel-Datei mit professionellem Design"""
    print("🎨 ERSTELLE SCHÖNES EXCEL-DESIGN...")

    # Speichere DataFrame als Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Financial Analysis', index=False)

    # Lade Workbook für Formatierung
    wb = load_workbook(output_path)
    ws = wb['Financial Analysis']

    # 🎨 FARB-SCHEMA (Elegantes Blau-Grün Design)
    header_fill = PatternFill(start_color="1f4e79", end_color="1f4e79", fill_type="solid")  # Dunkles Blau
    company_fill = PatternFill(start_color="d9e2f3", end_color="d9e2f3", fill_type="solid")  # Helles Blau
    metrics_fill = PatternFill(start_color="e2efda", end_color="e2efda", fill_type="solid")  # Helles Grün
    alternating_fill = PatternFill(start_color="f8f9fa", end_color="f8f9fa", fill_type="solid")  # Sehr helles Grau

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
        row_fill = alternating_fill if is_even else PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")

        for col_num, cell in enumerate(ws[row_num], 1):
            col_name = df.columns[col_num - 1]

            # Basis-Formatierung
            cell.border = thin_border
            cell.font = data_font

            # Spezielle Formatierung je Spalten-Typ
            if col_name in company_cols:
                cell.fill = company_fill if col_name == 'Name' else row_fill
                cell.font = company_font if col_name == 'Name' else Font(name="Calibri", size=10, bold=True, color="1f4e79")
                cell.alignment = left_alignment
            elif col_name in category_cols:
                cell.fill = row_fill
                cell.alignment = center_alignment
            elif col_name in metric_cols:
                cell.fill = metrics_fill if not is_even else row_fill
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
                cell.fill = row_fill
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

    # 5️⃣ CONDITIONAL FORMATTING FÜR KENNZAHLEN
    print("  🎯 Füge Conditional Formatting hinzu...")
    for col_num, col_name in enumerate(df.columns, 1):
        if col_name in metric_cols:
            col_letter = ws.cell(row=1, column=col_num).column_letter
            data_range = f"{col_letter}2:{col_letter}{len(df)+1}"

            # Farbverlauf für bessere Visualisierung (Grün für hohe, Rot für niedrige Werte)
            color_scale = ColorScaleRule(
                start_type='min', start_color='ffcccc',  # Helles Rot
                mid_type='percentile', mid_value=50, mid_color='ffffcc',  # Helles Gelb
                end_type='max', end_color='ccffcc'  # Helles Grün
            )
            ws.conditional_formatting.add(data_range, color_scale)

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

    # Metadata über mehrere Spalten mergen
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=min(6, len(df.columns)))

    # Speichere formatierte Datei
    wb.save(output_path)

    print("  ✨ Excel-Formatierung abgeschlossen!")
    print(f"  📋 Titel: {title_cell.value}")
    print(f"  📊 {len(df)} Unternehmen formatiert")
    print(f"  📈 {len(metric_cols)} Kennzahlen mit Conditional Formatting")
    print(f"  💾 Datei gespeichert: {output_path}")
