# Testkonzept: Excel-Datenverarbeitungsskript

## Teststatus-Legende
- ✅ Getestet und erfolgreich
- ❌ Getestet und fehlgeschlagen
- ⏳ Test ausstehend
- 🔄 In Bearbeitung
- 🚫 Test nicht möglich/relevant

## 1. Äquivalenzklassen für Eingabeparameter

### 1.1 Unternehmens-Identifikation (Spalte A & B)

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.1.1 | Gültiger RIC (Spalte B) | "RL.N", "HRMS.PA" | RIC-basierte Suche, RIC-basierte Filterung | ✅ |
| 1.1.2 | Gültiger Name (Spalte A) | "Hermes", "Ralph Lauren" | Name-basierte Suche, Zuordnung zum richtigen RIC | ✅ |
| 1.1.3 | Name zu kurz (< 4 Zeichen) | "Zara", "VF" | Fehlermeldung, Name zu kurz | ✅ |
| 1.1.4 | Nicht existierender RIC | "XXX.YY" | Fehlermeldung, Unternehmen nicht gefunden | ✅ |
| 1.1.5 | Nicht existierender Name | "Fantasiefirma AG" | Fehlermeldung, Unternehmen nicht gefunden | ✅ |
| 1.1.6 | Leerer RIC und leerer Name | "", "" | Fehlermeldung, kein Unternehmen angegeben | ✅ |
| 1.1.7 | RIC und Name gleichzeitig angegeben | RIC="RL.N", Name="Ralph" | RIC hat Priorität, Name wird ignoriert | ✅ |

### 1.2 Filterkriterien (Sub-Industry vs. Focus)

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.2.1 | Sub-Industry markiert (X) | Sub-Industry="X", Focus="" | Filterung nach Sub-Industry des Startunternehmens | ✅ |
| 1.2.2 | Focus markiert (X) | Sub-Industry="", Focus="X" | Filterung nach Focus-Gruppe des Startunternehmens | ✅ |
| 1.2.3 | Beide markiert | Sub-Industry="X", Focus="X" | Focus hat Priorität (gemäß Implementierung) | ✅ |
| 1.2.4 | Keines markiert | Sub-Industry="", Focus="" | Standard ist Sub-Industry Filterung | ✅ |
| 1.2.5 | Focus markiert aber Unternehmen hat keinen Focus | Focus="X", Unternehmen ohne Focus-Wert | Fallback auf Sub-Industry oder Fehlermeldung | ✅ |

### 1.3 Kennzahlen (Excel und Refinitiv)

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.3.1 | Gültige Excel-Kennzahlen | "ISIN", "P/E", "Free Float" | Daten aus Excel-Dateien extrahiert | ✅ |
| 1.3.2 | Gültige Refinitiv-Kennzahlen | "TR.Revenue", "TR.PriceClose" | Daten aus Refinitiv API geholt | ✅ |
| 1.3.3 | Nicht existierende Excel-Kennzahlen | "XYZ123" | Leerer oder N/A Wert in Ergebnis | ✅ |
| 1.3.4 | Nicht existierende Refinitiv-Kennzahlen | "TR.NonExistent" | Leerer oder N/A Wert in Ergebnis | ✅ |
| 1.3.5 | Keine Kennzahlen angegeben | [] | Minimale Ausgabe (nur Name, RIC, etc.) | ✅ |
| 1.3.6 | Duplikate in Kennzahlen | ["ISIN", "ISIN"] | Duplikate werden ignoriert | ⏳ |
| 1.3.7 | Mix aus gültigen und ungültigen Kennzahlen | ["ISIN", "XYZ123"] | Gültige werden verarbeitet, ungültige ignoriert | ⏳ |

### 1.4 Mehrere Unternehmen im Input

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.4.1 | Mehrere Unternehmen mit gleicher Sub-Industry | "Hermes", "Ralph Lauren" | Alle Unternehmen der gleichen Sub-Industry ohne Duplikate | ⏳ |
| 1.4.2 | Mehrere Unternehmen mit unterschiedlicher Sub-Industry | "Hermes", "Netflix" | Unternehmen beider Sub-Industries | ⏳ |
| 1.4.3 | Mehrere Unternehmen mit gleichem Focus | "Hermes", "Brunello Cucinelli" | Alle Unternehmen des gleichen Focus ohne Duplikate | ⏳ |
| 1.4.4 | Mix aus Sub-Industry und Focus Filterung | Zeile 1: Sub-Industry="X", Zeile 2: Focus="X" | Jedes Unternehmen nach seinem Filter-Kriterium | ⏳ |
| 1.4.5 | Doppelte Unternehmen im Input | "RL.N" in Zeile 1 & 2 | Duplikaterkennung, nur einmal verarbeiten | ⏳ |

## 2. Grenzfälle und spezielle Szenarien

| ID | Szenario | Testfall | Erwartetes Verhalten | Status |
|----|----------|---------|---------------------|--------|
| 2.1 | Große Datenmengen | 100+ Unternehmen in einer Sub-Industry | Korrekte Verarbeitung aller Unternehmen, Speicherverwaltung | ⏳ |
| 2.2 | Fehlende Excel-Dateien | Excel-Datei nicht vorhanden | Robuste Fehlerbehandlung, Fortsetzung mit verfügbaren Daten | ⏳ |
| 2.3 | Leere Excel-Sheets | Excel-Sheet ohne Daten | Robuste Fehlerbehandlung, Fortsetzung mit anderen Sheets | ⏳ |
| 2.4 | Formatfehler in Excel | Falsch formatierte Zellen | Robuste Fehlerbehandlung, bestmögliche Datenextraktion | ⏳ |
| 2.5 | Refinitiv API nicht verfügbar | API-Ausfall | Robuste Fehlerbehandlung, Fortsetzung mit Excel-Daten | ⏳ |
| 2.6 | Refinitiv Ratenlimitierung | API-Limit erreicht | Robuste Fehlerbehandlung, evtl. Wartezeiten | ⏳ |
| 2.7 | Output-Verzeichnis existiert nicht | "data" Verzeichnis fehlt | Verzeichnis erstellen oder Fehlermeldung | ⏳ |
| 2.8 | Output-Datei bereits vorhanden | output.xlsx existiert bereits | Überschreiben oder Backup der alten Datei | ⏳ |

## 3. Kombinationsszenarien

| ID | Szenario | Kombination | Erwartetes Verhalten | Status |
|----|----------|------------|---------------------|--------|
| 3.1 | Mix aus allen Filtertypen | Mehrere Zeilen mit unterschiedlichen Filtern | Korrekte Verarbeitung aller Filter | ⏳ |
| 3.2 | Mix aus Namen und RICs | Einige Zeilen mit RIC, andere mit Namen | Korrekte Identifikation aller Unternehmen | ⏳ |
| 3.3 | Mix aus Excel und Refinitiv | Beide Kennzahlentypen kombiniert | Daten aus beiden Quellen korrekt zusammengeführt | ⏳ |
| 3.4 | Teilweise fehlerhafte Eingaben | Einige gültige, einige ungültige Zeilen | Gültige verarbeiten, Fehler protokollieren | ⏳ |

## 4. Leistungstests

| ID | Testfall | Beschreibung | Akzeptanzkriterien | Status |
|----|----------|-------------|-------------------|--------|
| 4.1 | Viele Kennzahlen | 20+ Kennzahlen pro Unternehmen | Verarbeitung < 5 Min. | ⏳ |
| 4.2 | Viele Unternehmen | 100+ Unternehmen in Ergebnis | Verarbeitung < 5 Min. | ⏳ |
| 4.3 | Große Excel-Dateien | Excel-Dateien > 10 MB | Robuste Verarbeitung | ⏳ |
| 4.4 | Mehrere Unternehmen mit vielen Kennzahlen | 5 Unternehmen, je 10 Kennzahlen | Verarbeitung < 3 Min. | ⏳ |

## 5. Regressionstests

| ID | Testfall | Was geprüft wird | Status |
|----|----------|-----------------|--------|
| 5.1 | Standard-Fall: Hermes + ISIN | Basis-Funktionalität | ⏳ |
| 5.2 | Ralph Lauren RIC + 3 Kennzahlen | Multi-Kennzahlen Verarbeitung | ⏳ |
| 5.3 | 3 Unternehmen + Sub-Industry Filter | Multiple Eingabe mit Sub-Industry | ⏳ |
| 5.4 | 3 Unternehmen + Focus Filter | Multiple Eingabe mit Focus | ⏳ |
| 5.5 | Kennzahlen aus Refinitiv | Refinitiv-Integration | ⏳ |

## 6. Testdaten

### 6.1 Unternehmen für Tests
- Hermes International SCA (HRMS.PA) - Luxury Goods
- Ralph Lauren Corp (RL.N) - Luxury Goods
- Nike Inc (NKE.N) - Footwear
- Netflix Inc (NFLX.O) - Entertainment
- LVMH (LVMH.PA) - Luxury Goods
- Brunello Cucinelli (BCU.MI) - Luxury Goods (High Focus)

### 6.2 Kennzahlen für Tests
- Excel: "ISIN", "P/E", "Free Float", "Market in USD", "Forward P/E"
- Refinitiv: "TR.Revenue", "TR.PriceClose", "TR.TotalReturn"

## 7. Testprotokoll

| Datum | Testfall-ID | Getestet von | Ergebnis | Kommentar |
|-------|------------|--------------|----------|------------|
| YYYY-MM-DD | 1.1.1 | | | |
| YYYY-MM-DD | 1.1.2 | | | |
| YYYY-MM-DD | 1.2.1 | | | |
