import pandas as pd
import refinitiv.data as rd
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

# GICS Sektor Mapping zu Refinitiv Codes
GICS_SECTOR_CODES = {
    'Consumer Discretionary': '25',
    'Consumer Staples': '30',
    'Energy': '10',
    'Financials': '40',
    'Health Care': '35',
    'Industrials': '20',
    'Information Technology': '45',
    'Materials': '15',
    'Real Estate': '60',
    'Communication Services': '50',
    'Utilities': '55'
}

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

def calculate_gics_average_for_sector(sector_code, field_expression, resolved_col_name):
    """Berechne GICS-Durchschnitt für einen spezifischen Sektor"""
    try:
        print(f"   📊 Berechne GICS-Durchschnitt für Sektor {sector_code}: {field_expression}")

        sample = rd.get_data(
            universe=f'SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.GICSSectorCode,"{sector_code}"), CURN=USD)',
            fields=[field_expression]
        )
        if not sample.empty:
            col = sample.columns[-1]
            values = pd.to_numeric(sample[col], errors='coerce').dropna()
            if not values.empty and len(values) > 5:  # Mindestens 5 Werte
                # Entferne Ausreißer (5 % und 95 % Quantile)
                lower = values.quantile(0.05)
                upper = values.quantile(0.95)
                filtered_values = values[(values >= lower) & (values <= upper)]
                if not filtered_values.empty:
                    avg = round(filtered_values.mean(), 4)
                    print(f"     ✅ GICS-Durchschnitt für {resolved_col_name}: {avg:,} (aus {len(filtered_values)} von {len(values)} Unternehmen)")
                    return avg
        return None
    except Exception as e:
        print(f"     ⚠️ Fehler beim GICS-Durchschnitt für Sektor {sector_code}: {e}")
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

def get_consumer_discretionary_sector_average(refinitiv_fields):
    """
    Berechnet den Durchschnitt für ALLE Refinitiv-Kennzahlen über den gesamten
    GICS Consumer Discretionary Sector (25) mit Ausreißer-Filterung (5% oben/unten)
    """
    print("🏭 BERECHNE CONSUMER DISCRETIONARY SECTOR DURCHSCHNITTE...")

    if not refinitiv_fields:
        print("⚠️ Keine Refinitiv-Kennzahlen angegeben")
        return {}

    sector_averages = {}

    try:
        print("🔄 Öffne Refinitiv-Session für Sector-Analyse...")
        rd.open_session()

        # Verwende vereinfachte Methode für Sector-Screening
        print("📋 Hole Consumer Discretionary Sektor-Durchschnitte...")

        # Berechne für jede Refinitiv-Kennzahl den Sektor-Durchschnitt
        for field_expr in refinitiv_fields:
            if not field_expr.strip():
                continue

            # Stelle sicher, dass TR. am Anfang steht
            if not field_expr.startswith('TR.'):
                field_expr = 'TR.' + field_expr

            try:
                print(f"📊 Berechne Sektor-Durchschnitt für: {field_expr}")

                # Hole Daten für Consumer Discretionary Sector
                sector_data = rd.get_data(
                    universe='SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.GICSSectorCode,"25"), CURN=USD)',
                    fields=[field_expr]
                )

                if not sector_data.empty:
                    # Resolva den echten Spaltennamen
                    resolved_col_name = resolve_field_name(field_expr)

                    # Finde die richtige Spalte
                    data_col = None
                    for col in sector_data.columns:
                        if col != 'Instrument':
                            data_col = col
                            break

                    if data_col:
                        # Konvertiere zu numerischen Werten
                        values = pd.to_numeric(sector_data[data_col], errors='coerce').dropna()

                        if not values.empty and len(values) > 10:  # Mindestens 10 Werte für sinnvollen Durchschnitt
                            # Entferne Ausreißer (5 % und 95 % Quantile)
                            lower = values.quantile(0.05)
                            upper = values.quantile(0.95)
                            filtered_values = values[(values >= lower) & (values <= upper)]

                            if not filtered_values.empty:
                                avg = round(filtered_values.mean(), 4)
                                sector_averages[resolved_col_name] = avg
                                print(f"   ✅ {resolved_col_name}: {avg:,} (aus {len(filtered_values)} Werten)")
                            else:
                                print(f"   ❌ {resolved_col_name}: Keine Werte nach Filterung")
                        else:
                            print(f"   ❌ {resolved_col_name}: Zu wenig Daten ({len(values)} Werte)")
                    else:
                        print(f"   ❌ {field_expr}: Keine Datenspalte gefunden")
                else:
                    print(f"   ❌ {field_expr}: Keine Sektor-Daten erhalten")

            except Exception as e:
                print(f"   ❌ Fehler bei {field_expr}: {e}")
                continue

        print(f"✅ {len(sector_averages)} Sektor-Durchschnitte berechnet")
        return sector_averages

    except Exception as e:
        print(f"❌ Fehler bei Sektor-Durchschnittsberechnung: {e}")
        return {}
    finally:
        try:
            rd.close_session()
            print("✅ Refinitiv-Session geschlossen")
        except:
            pass

def get_sector_average_by_companies(companies, field_expressions):
    """
    Berechnet Refinitiv-Kennzahlen-Durchschnitte für eine spezifische Liste von Unternehmen

    Args:
        companies: Liste von Unternehmen-Dictionaries mit 'RIC'-Schlüssel
        field_expressions: Liste von Refinitiv-Feldausdrücken

    Returns:
        Dictionary mit Durchschnittswerten für jedes Feld
    """
    if not companies or not field_expressions:
        return {}

    # Extrahiere RICs
    ric_list = [comp['RIC'] for comp in companies if comp.get('RIC')]

    if not ric_list:
        return {}

    print(f"📊 Berechne Durchschnitte für {len(ric_list)} Unternehmen...")

    try:
        rd.open_session()

        # Hole Refinitiv-Daten
        all_data = fetch_refinitiv_data(ric_list, field_expressions)

        if not all_data:
            print("⚠️ Keine Refinitiv-Daten erhalten")
            return {}

        # Berechne Durchschnitte für jedes Feld
        averages = {}

        for field_expr in field_expressions:
            if not field_expr.strip():
                continue

            # Finde die entsprechende Spalte im Dictionary
            resolved_field = resolve_field_name(field_expr)

            # Suche nach dem Feld in den Daten (verschiedene Varianten probieren)
            field_data = None
            for data_key in all_data.keys():
                if (data_key == resolved_field or
                    data_key == field_expr or
                    data_key.replace('TR.', '') == resolved_field or
                    data_key.replace('TR.', '') == field_expr.replace('TR.', '')):
                    field_data = all_data[data_key]
                    break

            if field_data:
                # Konvertiere zu numerischen Werten und berechne Durchschnitt
                values = []
                for ric, value in field_data.items():
                    if pd.notna(value) and str(value).strip() != '':
                        try:
                            # Bereinige Wert falls nötig
                            clean_value = str(value).replace(',', '').replace('%', '')
                            num_val = pd.to_numeric(clean_value, errors='coerce')
                            if pd.notna(num_val):
                                values.append(num_val)
                        except:
                            continue

                if len(values) > 0:
                    avg_value = sum(values) / len(values)
                    averages[resolved_field] = avg_value
                    print(f"   📈 {resolved_field}: {avg_value:.4f} (aus {len(values)} von {len(ric_list)} Unternehmen)")
                else:
                    print(f"   ⚠️ {resolved_field}: Keine gültigen Werte gefunden")
            else:
                print(f"   ❌ {field_expr}: Feld nicht in den Daten gefunden")

        return averages

    except Exception as e:
        print(f"❌ Fehler bei Durchschnittsberechnung: {e}")
        return {}
    finally:
        try:
            rd.close_session()
        except:
            pass

def fetch_refinitiv_sector_averages(sector_name, field_expressions):
    """
    Hole echte Refinitiv-Sektor-Durchschnitte direkt von Refinitiv
    Verwendet keine Einzelunternehmen, sondern die bereits berechneten Sektor-Durchschnitte
    """
    if sector_name not in GICS_SECTOR_CODES:
        print(f"⚠️ GICS-Sektor '{sector_name}' nicht im Mapping gefunden")
        return None

    sector_code = GICS_SECTOR_CODES[sector_name]
    print(f"🔍 Hole Refinitiv-Sektor-Durchschnitte für {sector_name} (GICS: {sector_code})")

    try:
        import refinitiv.data as rd
        rd.open_session()

        sector_averages = {}

        # Verwende einen speziellen Sektor-RIC oder Index für Durchschnittswerte
        # Refinitiv bietet oft Sektor-Indizes an
        sector_rics = {
            'Consumer Discretionary': ['.SPCD', 'XLY', '.DJU5340', 'IYC'],  # Verschiedene Sektor-Indizes
            'Consumer Staples': ['.SPCS', 'XLP', '.DJU5350', 'XLP'],
            'Information Technology': ['.SPIT', 'XLK', '.DJU9530', 'IGV'],
            'Health Care': ['.SPHC', 'XLV', '.DJU4530', 'IHI'],
            'Materials': ['.SPMT', 'XLB', '.DJU1510', 'VAW'],
            'Energy': ['.SPEN', 'XLE', '.DJU1010', 'VDE'],
            'Financials': ['.SPFN', 'XLF', '.DJU4010', 'VFH'],
            'Industrials': ['.SPIN', 'XLI', '.DJU2010', 'VIS'],
            'Utilities': ['.SPUT', 'XLU', '.DJU5510', 'VPU'],
            'Real Estate': ['.SPRE', 'XLRE', '.DJU6010', 'VNQ'],
            'Communication Services': ['.SPCM', 'XLC', '.DJU5010', 'VOX']
        }

        # Fallback: Verwende direkte Sektor-Aggregate-Abfrage
        for field_expr in field_expressions:
            if not field_expr.strip():
                continue

            # Stelle sicher, dass TR. am Anfang steht
            if not field_expr.startswith('TR.'):
                field_expr = 'TR.' + field_expr

            try:
                print(f"   📊 Hole Sektor-Durchschnitt für: {field_expr}")

                # Methode 1: Verwende Sektor-Index falls verfügbar
                sector_value = None
                if sector_name in sector_rics:
                    for sector_ric in sector_rics[sector_name]:
                        try:
                            sector_data = rd.get_data(universe=[sector_ric], fields=[field_expr])
                            if not sector_data.empty and len(sector_data.columns) > 1:
                                data_col = [col for col in sector_data.columns if col != 'Instrument'][0]
                                value = sector_data[data_col].iloc[0]
                                if pd.notna(value):
                                    sector_value = value
                                    print(f"     ✅ Gefunden über Sektor-Index {sector_ric}: {value}")
                                    break
                        except:
                            continue

                # Methode 2: Falls kein Sektor-Index funktioniert, verwende aggregierte Sektor-Abfrage
                if sector_value is None:
                    try:
                        # Spezielle Refinitiv-Syntax für Sektor-Aggregate
                        aggregate_universe = f"GICS({sector_code})"  # Vereinfachte GICS-Syntax

                        # Alternative: Verwende Screening mit Aggregation
                        screen_universe = f"SCREEN(U(IN(Equity(active,public,primary))), IN(TR.GICSSector,{sector_code}), CURN=USD, TOP(500))"

                        aggregate_data = rd.get_data(universe=screen_universe, fields=[field_expr])

                        if not aggregate_data.empty and len(aggregate_data.columns) > 1:
                            data_col = [col for col in aggregate_data.columns if col != 'Instrument'][0]
                            # Berechne Median als robustereren Durchschnitt
                            numeric_values = pd.to_numeric(aggregate_data[data_col], errors='coerce').dropna()
                            if len(numeric_values) > 0:
                                sector_value = numeric_values.median()  # Median ist robuster als Mean
                                print(f"     ✅ Berechnet über Sektor-Screening: {sector_value} (aus {len(numeric_values)} Unternehmen)")

                    except Exception as e:
                        print(f"     ⚠️ Aggregate-Abfrage fehlgeschlagen: {e}")

                # Methode 3: Fallback - verwende den größten ETF des Sektors
                if sector_value is None and sector_name in sector_rics:
                    try:
                        main_etf = sector_rics[sector_name][1]  # Normalerweise XL* ETFs
                        etf_data = rd.get_data(universe=[main_etf], fields=[field_expr])
                        if not etf_data.empty and len(etf_data.columns) > 1:
                            data_col = [col for col in etf_data.columns if col != 'Instrument'][0]
                            value = etf_data[data_col].iloc[0]
                            if pd.notna(value):
                                sector_value = value
                                print(f"     ✅ Fallback über ETF {main_etf}: {value}")
                    except:
                        pass

                if sector_value is not None:
                    clean_field = field_expr.replace('TR.', '')
                    sector_averages[clean_field] = round(float(sector_value), 4)
                    print(f"     ✅ {clean_field}: {sector_value:.4f}")
                else:
                    print(f"     ❌ Keine Sektor-Daten für {field_expr} verfügbar")

            except Exception as e:
                print(f"     ❌ Fehler bei {field_expr}: {e}")
                continue

        rd.close_session()
        return sector_averages if sector_averages else None

    except Exception as e:
        print(f"   ❌ Fehler beim Öffnen der Refinitiv-Session: {e}")
        try:
            rd.close_session()
        except:
            pass
        return None

def get_all_sector_averages(used_sectors, refinitiv_fields):
    """
    Berechnet Refinitiv-Kennzahlen-Durchschnitte für alle verwendeten GICS-Sektoren
    """
    print("🏭 BERECHNE DURCHSCHNITTE FÜR ALLE VERWENDETEN GICS-SEKTOREN...")

    if not refinitiv_fields or not used_sectors:
        print("⚠️ Keine Refinitiv-Kennzahlen oder Sektoren angegeben")
        return {}

    all_sector_averages = {}

    try:
        print("🔄 Öffne Refinitiv-Session für Multi-Sektor-Analyse...")
        rd.open_session()

        for sector_name in used_sectors:
            # Finde den GICS-Code für den Sektor
            sector_code = GICS_SECTOR_CODES.get(sector_name)
            if not sector_code:
                print(f"⚠️ GICS-Code für Sektor '{sector_name}' nicht gefunden")
                continue

            print(f"📋 Berechne Durchschnitte für {sector_name} (GICS {sector_code})...")

            sector_averages = {}

            # Berechne für jede Refinitiv-Kennzahl den Sektor-Durchschnitt
            for field_expr in refinitiv_fields:
                if not field_expr.strip():
                    continue

                # Stelle sicher, dass TR. am Anfang steht
                if not field_expr.startswith('TR.'):
                    field_expr = 'TR.' + field_expr

                try:
                    print(f"   📊 Berechne Sektor-Durchschnitt für: {field_expr}")

                    # Hole Daten für den spezifischen GICS-Sektor
                    sector_data = rd.get_data(
                        universe=f'SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.GICSSectorCode,"{sector_code}"), CURN=USD)',
                        fields=[field_expr]
                    )

                    if not sector_data.empty:
                        # Resolva den echten Spaltennamen
                        resolved_col_name = resolve_field_name(field_expr)

                        # Finde die richtige Spalte
                        data_col = None
                        for col in sector_data.columns:
                            if col != 'Instrument':
                                data_col = col
                                break

                        if data_col:
                            # Konvertiere zu numerischen Werten
                            values = pd.to_numeric(sector_data[data_col], errors='coerce').dropna()

                            if not values.empty and len(values) > 5:  # Mindestens 5 Werte für sinnvollen Durchschnitt
                                # Entferne Ausreißer (5 % und 95 % Quantile)
                                lower = values.quantile(0.05)
                                upper = values.quantile(0.95)
                                filtered_values = values[(values >= lower) & (values <= upper)]

                                if not filtered_values.empty:
                                    avg = round(filtered_values.mean(), 4)
                                    sector_averages[resolved_col_name] = avg
                                    print(f"     ✅ {resolved_col_name}: {avg:,} (aus {len(filtered_values)} von {len(values)} Unternehmen)")
                                else:
                                    print(f"     ❌ {resolved_col_name}: Keine Werte nach Filterung")
                            else:
                                print(f"     ❌ {resolved_col_name}: Zu wenig Daten ({len(values)} Werte)")
                        else:
                            print(f"     ❌ {field_expr}: Keine Datenspalte gefunden")
                    else:
                        print(f"     ❌ {field_expr}: Keine Sektor-Daten erhalten")

                except Exception as e:
                    print(f"     ❌ Fehler bei {field_expr}: {e}")
                    continue

            if sector_averages:
                all_sector_averages[sector_name] = sector_averages
                print(f"   ✅ {len(sector_averages)} Durchschnitte für {sector_name} berechnet")
            else:
                print(f"   ⚠️ Keine Durchschnitte für {sector_name} berechnet")

        print(f"✅ GICS-Sektor-Durchschnitte für {len(all_sector_averages)} Sektoren berechnet")
        return all_sector_averages

    except Exception as e:
        print(f"❌ Fehler bei Multi-Sektor-Durchschnittsberechnung: {e}")
        return {}
    finally:
        try:
            rd.close_session()
            print("✅ Refinitiv-Session geschlossen")
        except:
            pass

