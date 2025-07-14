import pandas as pd
import refinitiv.data as rd
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

def resolve_field_name(field_expression):
    """Liefert den tatsächlichen Spaltennamen zu einem Refinitiv-Feldausdruck"""
    try:
        sample = rd.get_data(universe="IBM.N", fields=[field_expression])
        if not sample.empty:
            return sample.columns[-1]
    except Exception as e:
        print(f"⚠️ Feldauflösung fehlgeschlagen für '{field_expression}': {e}")
        pass
    return field_expression  # Fallback

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

                # Speichere Daten
                if resolved_col_name in data.columns:
                    results[resolved_col_name] = data.set_index('RIC')[resolved_col_name].to_dict()
                    print(f"✅ {resolved_col_name}: {len(data)} Datensätze erhalten")
                else:
                    print(f"❌ Spalte '{resolved_col_name}' nicht in Daten gefunden")

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
