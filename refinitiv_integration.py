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
