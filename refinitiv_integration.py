import pandas as pd
import refinitiv.data as rd
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

def resolve_field_name(field_expression):
    """Liefert den tatsächlichen Spaltennamen zu einem Refinitiv-Feldausdruck mit Period-Information"""
    try:
        sample = rd.get_data(universe="IBM.N", fields=[field_expression])
        if not sample.empty:
            original_col_name = sample.columns[-1]

            # Extrahiere Period-Information aus dem ursprünglichen Feldausdruck
            if "(Period=" in field_expression:
                # Extrahiere TR.EBIT(Period=FY-1) → EBIT(Period=FY-1)
                base_field = field_expression.replace("TR.", "")
                return base_field
            else:
                # Ohne Period-Information: TR.EBIT → EBIT
                return original_col_name

    except Exception as e:
        print(f"⚠️ Feldauflösung fehlgeschlagen für '{field_expression}': {e}")
        pass

    # Fallback: Entferne nur TR. aber behalte Period-Information
    if field_expression.startswith('TR.'):
        return field_expression.replace('TR.', '')
    return field_expression

def fetch_refinitiv_data(ric_list, field_expressions):
    """Hole Refinitiv-Daten für mehrere RICs und Felder"""
    if not field_expressions:
        return pd.DataFrame()

    results = {}

    for field_expr in field_expressions:
        if not field_expr.strip():
            continue

        # Stelle sicher, dass TR. am Anfang steht
        if not field_expr.startswith('TR.'):
            field_expr = 'TR.' + field_expr

        try:
            print(f"📊 Hole Refinitiv-Daten für Feld: {field_expr}")

            # Hole Daten für alle RICs
            data = rd.get_data(universe=ric_list, fields=[field_expr])

            if not data.empty:
                # Resolva den echten Spaltennamen
                resolved_col_name = resolve_field_name(field_expr)

                # Reset index und bereite DataFrame vor
                data.reset_index(inplace=True)
                data.rename(columns={"Instrument": "RIC"}, inplace=True)
                data['RIC'] = data['RIC'].str.upper()

                # DEBUG: Zeige die tatsächlichen Refinitiv-Daten
                print(f"🔍 DEBUG: Refinitiv-Daten für {field_expr}:")
                print(f"   Spalten: {list(data.columns)}")
                print(f"   Erste 3 Zeilen: {data.head(3).to_dict('records')}")

                # Speichere Daten - verwende den tatsächlichen Spaltennamen aus der API
                data_columns = [col for col in data.columns if col not in ['RIC', 'index']]
                if data_columns:
                    actual_col_name = data_columns[0]  # Nimm die erste Nicht-RIC/Nicht-Index-Spalte
                    ric_data = data.set_index('RIC')[actual_col_name].to_dict()
                    results[resolved_col_name] = ric_data
                    print(f"✅ {resolved_col_name}: {len(data)} Datensätze erhalten (API-Spalte: {actual_col_name})")
                    print(f"   Beispiel-Werte: {dict(list(ric_data.items())[:3])}")
                else:
                    print(f"❌ Keine Datenspalten gefunden für '{field_expr}'")

        except Exception as e:
            print(f"❌ Fehler beim Abrufen von '{field_expr}': {e}")
            continue

    return results

def calculate_gics_average(field_expression, resolved_col_name):
    """Berechne GICS-Durchschnitt für Consumer Discretionary Sektor"""
    try:
        sample = rd.get_data(
            universe='SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.GICSSectorCode,"25"), CURN=USD)',
            fields=[field_expression]
        )
        if not sample.empty:
            col = sample.columns[-1]
            values = pd.to_numeric(sample[col], errors='coerce').dropna()
            if not values.empty:
                # Entferne Ausreißer (5 % und 95 % Quantile)
                lower = values.quantile(0.05)
                upper = values.quantile(0.95)
                values = values[(values >= lower) & (values <= upper)]
                avg = round(values.mean(), 2)
                print(f"📉 GICS-Durchschnitt für {resolved_col_name}: {avg:,}")
                return avg
        return None
    except Exception as e:
        print(f"⚠️ Fehler beim GICS-Durchschnitt: {e}")
        return None

def format_refinitiv_value(value):
    """Formatiere Refinitiv-Werte für die Ausgabe"""
    if pd.isna(value):
        return ""

    try:
        # Versuche numerische Formatierung
        numeric_value = pd.to_numeric(value, errors='coerce')
        if not pd.isna(numeric_value):
            if numeric_value >= 1000000:
                return f"{numeric_value:,.0f}"
            elif numeric_value >= 1000:
                return f"{numeric_value:,.0f}"
            else:
                return f"{numeric_value:.2f}"
        else:
            return str(value)
    except:
        return str(value)

def get_refinitiv_kennzahlen_for_companies(companies, refinitiv_fields):
    """Hole Refinitiv-Kennzahlen für alle Unternehmen"""
    if not refinitiv_fields or not companies:
        return {}

    print(f"🔄 Öffne Refinitiv-Session...")

    try:
        rd.open_session()

        # Sammle alle RICs
        ric_list = [company['RIC'] for company in companies if company.get('RIC')]

        print(f"📊 Hole Refinitiv-Daten für {len(ric_list)} RICs und {len(refinitiv_fields)} Felder")

        # Hole alle Refinitiv-Daten
        refinitiv_data = fetch_refinitiv_data(ric_list, refinitiv_fields)

        # Erstelle Ergebnis-Dictionary
        results = {}
        for company in companies:
            ric = company.get('RIC')
            if ric:
                company_data = {}
                for field_name, field_data in refinitiv_data.items():
                    value = field_data.get(ric.upper())
                    company_data[field_name] = format_refinitiv_value(value)
                results[ric] = company_data

        return results

    except Exception as e:
        print(f"❌ Fehler bei Refinitiv-Datenabfrage: {e}")
        return {}
    finally:
        try:
            rd.close_session()
            print("✅ Refinitiv-Session geschlossen")
        except:
            pass

# NEUE FUNKTIONEN FÜR DYNAMISCHE SEKTOR-ERKENNUNG

def get_gics_sector_mapping():
    """Mapping von GICS Sector Codes zu Namen"""
    return {
        "10": "Energy",
        "15": "Materials",
        "20": "Industrials",
        "25": "Consumer Discretionary",
        "30": "Consumer Staples",
        "35": "Health Care",
        "40": "Financials",
        "45": "Information Technology",
        "50": "Communication Services",
        "55": "Utilities",
        "60": "Real Estate"
    }

def detect_sector_from_excel_files(excel_files):
    """Erkenne GICS-Sektoren aus Excel-Dateinamen"""
    sector_mapping = {
        "consumer": "25",  # Consumer Discretionary
        "housing": "25",   # Housing gehört zu Consumer Discretionary
        "basic consumer": "25",  # Basic Consumer gehört zu Consumer Discretionary
        "health": "35",    # Health Care
        "it": "45",        # Information Technology
        "tech": "45",      # Information Technology (alternative)
        "financial": "40", # Financials
        "energy": "10",    # Energy
        "materials": "15", # Materials (beide Materials-Dateien)
        "chemicals": "15", # Materials Chemicals
        "commodities": "15", # Materials Commodities
        "industrial": "20", # Industrials
        "staples": "30",   # Consumer Staples
        "communication": "50", # Communication Services
        "utilities": "55", # Utilities
        "real_estate": "60" # Real Estate
    }

    detected_sectors = set()
    gics_mapping = get_gics_sector_mapping()

    for file_path in excel_files:
        filename = file_path.lower()
        sector_found = False

        # Prüfe spezielle Fälle zuerst
        if "basic consumer" in filename or "basic_consumer" in filename:
            sector_code = "25"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Basic Consumer")
            sector_found = True
        elif "housing" in filename:
            sector_code = "25"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Housing")
            sector_found = True
        elif "materials" in filename or "chemicals" in filename or "commodities" in filename:
            sector_code = "15"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Materials")
            sector_found = True
        elif "health" in filename:
            sector_code = "35"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Health Care")
            sector_found = True
        elif "it" in filename or "tech" in filename:
            sector_code = "45"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Technology")
            sector_found = True
        elif "utilities" in filename:
            sector_code = "55"
            sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
            detected_sectors.add((sector_code, sector_name))
            print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code}) - Utilities")
            sector_found = True

        # Fallback für andere Keywords
        if not sector_found:
            for keyword, sector_code in sector_mapping.items():
                if keyword in filename:
                    sector_name = gics_mapping.get(sector_code, f"Sector {sector_code}")
                    detected_sectors.add((sector_code, sector_name))
                    print(f"🎯 Erkannt aus '{file_path}': {sector_name} (Code: {sector_code})")
                    break

    return list(detected_sectors)

def get_sector_average(sector_code, sector_name, refinitiv_fields):
    """
    Berechnet den Durchschnitt für ALLE Refinitiv-Kennzahlen über einen spezifischen
    GICS Sector mit Ausreißer-Filterung (5% oben/unten)
    """
    print(f"🏭 BERECHNE {sector_name.upper()} SECTOR DURCHSCHNITTE (Code: {sector_code})...")

    if not refinitiv_fields:
        print("⚠️ Keine Refinitiv-Kennzahlen angegeben")
        return {}

    sector_averages = {}

    try:
        print("🔄 Öffne Refinitiv-Session für Sector-Analyse...")
        rd.open_session()

        print(f"📋 Hole {sector_name} Sektor-Durchschnitte...")

        # Berechne für jede Refinitiv-Kennzahl den Sektor-Durchschnitt
        for field_expr in refinitiv_fields:
            if not field_expr.strip():
                continue

            # Stelle sicher, dass TR. am Anfang steht
            if not field_expr.startswith('TR.'):
                field_expr = 'TR.' + field_expr

            try:
                print(f"📊 Berechne Sektor-Durchschnitt für: {field_expr}")

                # Hole Daten für spezifischen GICS Sector
                sector_data = rd.get_data(
                    universe=f'SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.GICSSectorCode,"{sector_code}"), CURN=USD)',
                    fields=[field_expr]
                )

                if not sector_data.empty:
                    # Bestimme die relevante Datenspalte
                    data_columns = [col for col in sector_data.columns if col not in ['Instrument']]
                    if data_columns:
                        col = data_columns[0]
                        values = pd.to_numeric(sector_data[col], errors='coerce').dropna()

                        if not values.empty and len(values) >= 5:  # Mindestens 5 Datenpunkte
                            # Entferne Ausreißer (5% und 95% Quantile)
                            lower = values.quantile(0.05)
                            upper = values.quantile(0.95)
                            filtered_values = values[(values >= lower) & (values <= upper)]

                            if not filtered_values.empty:
                                avg = round(filtered_values.mean(), 2)
                                sector_averages[field_expr] = avg
                                print(f"   ✅ {field_expr}: {avg:,} (aus {len(filtered_values)}/{len(values)} Werten)")
                            else:
                                print(f"   ⚠️ {field_expr}: Keine Werte nach Ausreißer-Filterung")
                        else:
                            print(f"   ⚠️ {field_expr}: Zu wenige Datenpunkte ({len(values)})")
                    else:
                        print(f"   ❌ {field_expr}: Keine Datenspalten gefunden")
                else:
                    print(f"   ❌ {field_expr}: Keine Sektor-Daten erhalten")

            except Exception as e:
                print(f"   ❌ Fehler bei {field_expr}: {e}")
                continue

        print(f"✅ {len(sector_averages)}/{len(refinitiv_fields)} Sektor-Durchschnitte berechnet für {sector_name}")
        return sector_averages

    except Exception as e:
        print(f"❌ Fehler bei Sektor-Durchschnitts-Berechnung: {e}")
        return {}
    finally:
        try:
            rd.close_session()
            print("✅ Refinitiv-Session geschlossen")
        except:
            pass

# BESTEHENDE FUNKTION FÜR RÜCKWÄRTSKOMPATIBILITÄT
def get_consumer_discretionary_sector_average(refinitiv_fields):
    """Legacy-Funktion für Consumer Discretionary - nutzt neue generische Funktion"""
    return get_sector_average("25", "Consumer Discretionary", refinitiv_fields)
