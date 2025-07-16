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
| 1.3.6 | Duplikate in Kennzahlen | ["ISIN", "ISIN"] | Duplikate werden ignoriert | ✅ |
| 1.3.7 | Mix aus gültigen und ungültigen Kennzahlen | ["ISIN", "XYZ123"] | Gültige werden verarbeitet, ungültige ignoriert | ✅ |

### 1.4 Mehrere Unternehmen im Input

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.4.1 | Mehrere Unternehmen mit gleicher Sub-Industry | "Hermes", "Ralph Lauren" | Alle Unternehmen der gleichen Sub-Industry ohne Duplikate | ✅ |
| 1.4.2 | Mehrere Unternehmen mit unterschiedlicher Sub-Industry | "Hermes", "Nike" | Unternehmen beider Sub-Industries | ✅ |
| 1.4.3 | Mehrere Unternehmen mit gleichem Focus | "Hermes", "Brunello Cucinelli" | Alle Unternehmen des gleichen Focus ohne Duplikate | ✅ |
| 1.4.4 | Mix aus Sub-Industry und Focus Filterung | Zeile 1: Sub-Industry="X", Zeile 2: Focus="X" | Jedes Unternehmen nach seinem Filter-Kriterium | ✅ |
| 1.4.5 | Doppelte Unternehmen im Input | "RL.N" in Zeile 1 & 2 | Duplikaterkennung, nur einmal verarbeiten | ✅ |

### 1.5 Vorjahresdaten (Period Parameter)

| ID | Äquivalenzklasse | Beispielwerte | Erwartetes Verhalten | Status |
|----|-----------------|--------------|---------------------|--------|
| 1.5.1 | Aktuelles Jahr (Standard) | "TR.EBIT", "TR.Revenue" | Aktuelle Daten ohne Period-Parameter | ⏳ |
| 1.5.2 | Vorjahr (FY-1) | "TR.EBIT(Period=FY-1)", "TR.Revenue(Period=FY-1)" | Daten vom Vorjahr, Spaltenname mit Period-Zusatz | ⏳ |
| 1.5.3 | Vorletztes Jahr (FY-2) | "TR.EBIT(Period=FY-2)", "TR.Revenue(Period=FY-2)" | Daten von vor 2 Jahren, Spaltenname mit Period-Zusatz | ⏳ |
| 1.5.4 | Mehrere Jahre derselben Kennzahl | "TR.EBIT", "TR.EBIT(Period=FY-1)", "TR.EBIT(Period=FY-2)" | Separate Spalten für jedes Jahr derselben Kennzahl | ⏳ |
| 1.5.5 | Ungültiger Period-Parameter | "TR.EBIT(Period=FY-10)", "TR.Revenue(Period=XYZ)" | Fehlermeldung oder Fallback auf aktuelle Daten | ⏳ |
| 1.5.6 | Gemischte Period-Parameter | "TR.EBIT(Period=FY-1)", "TR.Revenue", "TR.EBITDA(Period=FY-2)" | Korrekte Verarbeitung verschiedener Zeiträume | ⏳ |
| 1.5.7 | Quartalsweise Daten | "TR.EBIT(Period=Q1)", "TR.Revenue(Period=Q4-1)" | Quartalsdaten falls verfügbar | ⏳ |
| 1.5.8 | Fehlerhafte Syntax | "TR.EBIT(Period FY-1)", "TR.EBIT[Period=FY-1]" | Robuste Parsierung oder Fehlermeldung | ⏳ |

## 2. Grenzfälle und spezielle Szenarien

| ID | Szenario | Testfall | Erwartetes Verhalten | Status |
|----|----------|---------|---------------------|--------|
| 2.1 | Große Datenmengen | 100+ Unternehmen in einer Sub-Industry | Korrekte Verarbeitung aller Unternehmen, Speicherverwaltung | ⏳ |
| 2.2 | Fehlende Excel-Dateien | Excel-Datei nicht vorhanden | Robuste Fehlerbehandlung, Fortsetzung mit verfügbaren Daten | ⏳ |
| 2.3 | Leere Excel-Sheets | Excel-Sheet ohne Daten | Robuste Fehlerbehandlung, Fortsetzung mit anderen Sheets | ⏳ |
| 2.4 | Formatfehler in Excel | Falsch formatierte Zellen | Robuste Fehlerbehandlung, bestmögliche Datenextraktion | ⏳ |
| 2.5 | Refinitiv API nicht verfügbar | API-Ausfall | Robuste Fehlerbehandlung, Fortsetzung mit Excel-Daten | ⏳ |
| 2.6 | Refinitiv Ratenlimitierung | API-Limit erreicht | Robuste Fehlerbehandlung, evtl. Wartezeiten | ⏳ |
| 2.7 | Output-Verzeichnis existiert nicht | "data" Verzeichnis fehlt | Verzeichnis erstellen oder Fehlermeldung | ✅ |
| 2.8 | Output-Datei bereits vorhanden | output.xlsx existiert bereits | Überschreiben oder Backup der alten Datei | ⏳ |

## 3. Kombinationsszenarien

| ID | Szenario | Kombination | Erwartetes Verhalten | Status |
|----|----------|------------|---------------------|--------|
| 3.1 | Mix aus allen Filtertypen | Zeile 1: Sub-Industry="X", Zeile 2: Focus="X", Zeile 3: Keine Markierung | Korrekte Verarbeitung aller Filter, Sub-Industry für Zeile 3 | ✅ MP |
| 3.2 | Mix aus Namen und RICs | Zeile 1: Name="Hermes", Zeile 2: RIC="RL.N", Zeile 3: Name="Nike" | Korrekte Identifikation aller Unternehmen | ✅ MP |
| 3.3 | Mix aus Excel und Refinitiv | Kennzahlen: ["ISIN", "P/E", "TR.Revenue", "TR.EBITDA"] | Daten aus beiden Quellen korrekt zusammengeführt | ✅ MP |
| 3.4 | Teilweise fehlerhafte Eingaben | Zeile 1: Gültiger Name, Zeile 2: Ungültiger RIC, Zeile 3: Zu kurzer Name | Gültige verarbeiten, Fehler protokollieren | ✅ MP |
| 3.5 | Mehrere Unternehmen + viele Kennzahlen | 3 Unternehmen + 8 Kennzahlen (4 Excel, 4 Refinitiv) | Vollständige Datenmatrix mit Durchschnittswerten | ✅ MP |
| 3.6 | Überlappende Sub-Industries | Mehrere Unternehmen aus gleicher Sub-Industry über verschiedene Eingaben | Keine Duplikate, korrekte Gruppierung | ✅ MP |
| 3.7 | Überlappende Focus-Gruppen | Mehrere Unternehmen aus gleicher Focus-Gruppe über verschiedene Eingaben | Keine Duplikate, korrekte Gruppierung | ✅ MP |
| 3.8 | Gemischte Kennzahlen mit Fehlern | Gültige + ungültige Excel-Kennzahlen + gültige + ungültige Refinitiv-Kennzahlen | Nur gültige Kennzahlen verarbeitet, Fehler protokolliert | ✅ MP |
| 3.9 | Große Kombinationen | 5 Unternehmen + 15 Kennzahlen + gemischte Filter | Vollständige Verarbeitung mit Performance-Überwachung | ✅ MP |
| 3.10 | Verschiedene Sektoren | Unternehmen aus verschiedenen GICS-Sektoren | Korrekte Sektorklassifikation und -durchschnitte | ✅ MP |
| 3.11 | Unternehmen ohne Focus + Focus-Filter | Unternehmen ohne Focus-Wert bei Focus-Filterung | Fallback auf Sub-Industry mit Warnung | ❌ MP |
| 3.12 | Leere Kennzahlen + gültige Kennzahlen | Einige leere Kennzahl-Zellen, andere gefüllt | Nur gültige Kennzahlen verarbeitet | ⏳ MP |
| 3.13 | Duplikate in verschiedenen Formen | "RL.N" + "Ralph Lauren" im selben Input | Duplikaterkennung funktioniert | ⏳ MP |
| 3.14 | Teilname-Matching Variationen | "Ralph", "Lauren", "Ralph Lauren Corp" | Korrekte Zuordnung zu einem Unternehmen | ⏳ MP |
| 3.15 | Mehrsprachige Namen | Namen mit Umlauten oder Sonderzeichen | Korrekte Verarbeitung und Matching | ⏳ MP |

## 3.1 Detaillierte Kombinationstest-Szenarien

### 3.1.1 Minimale Kombinationen
- **Test A**: 1 Name + 1 RIC + Sub-Industry Filter + 2 Excel-Kennzahlen
- **Test B**: 2 Namen + Focus Filter + 1 Refinitiv-Kennzahl
- **Test C**: 3 RICs + gemischte Filter + 1 Excel + 1 Refinitiv Kennzahl

### 3.1.2 Mittlere Kombinationen
- **Test D**: 3 Unternehmen (1 Name, 2 RICs) + alle Filter-Varianten + 5 Kennzahlen
- **Test E**: 5 Unternehmen + Focus-Filter + 3 Excel + 2 Refinitiv Kennzahlen
- **Test F**: 4 Unternehmen aus verschiedenen Sub-Industries + Sub-Industry Filter

### 3.1.3 Maximale Kombinationen
- **Test G**: 10 Unternehmen + alle verfügbaren Kennzahlen + gemischte Filter
- **Test H**: Alle Luxury-Unternehmen + Focus-Filter + Top 10 Kennzahlen
- **Test I**: Alle Consumer Discretionary + Sub-Industry Filter + Refinitiv-Kennzahlen

## 3.2 Fehlerbehandlungs-Kombinationen

### 3.2.1 Teilweise fehlerhafte Eingaben
- **Test J**: 50% gültige, 50% ungültige Unternehmen + verschiedene Kennzahlen
- **Test K**: Gültige Unternehmen + 30% ungültige Kennzahlen
- **Test L**: Gemischte Eingaben mit verschiedenen Fehlertypen

### 3.2.2 Rand- und Grenzfälle
- **Test M**: Maximale Anzahl Unternehmen (100+) + minimale Kennzahlen
- **Test N**: Minimale Anzahl Unternehmen (1) + maximale Kennzahlen
- **Test O**: Alle möglichen Filter-Kombinationen gleichzeitig

## 3.3 Performance-Kombinationen

### 3.3.1 Datenvolumen-Tests
- **Test P**: 50 Unternehmen + 20 Kennzahlen (Zeit < 5 Min.)
- **Test Q**: 100 Unternehmen + 10 Kennzahlen (Zeit < 10 Min.)
- **Test R**: 10 Unternehmen + 50 Kennzahlen (Zeit < 3 Min.)

### 3.3.2 API-Belastungstests
- **Test S**: Viele Refinitiv-Calls gleichzeitig
- **Test T**: Sequenzielle vs. parallele Verarbeitung
- **Test U**: Ratenlimit-Verhalten bei großen Anfragen

## 4. Vorjahresdaten-Kombinationen

| ID | Szenario | Kombination | Erwartetes Verhalten | Status |
|----|----------|------------|---------------------|--------|
| 4.1 | Zeitreihenanalyse einer Kennzahl | "TR.EBIT", "TR.EBIT(Period=FY-1)", "TR.EBIT(Period=FY-2)" | 3 separate Spalten für EBIT über 3 Jahre | ✅ Funktioniert bei korrekter Eingabe - MP |
| 4.2 | Verschiedene Kennzahlen, verschiedene Jahre | "TR.EBIT(Period=FY-1)", "TR.Revenue(Period=FY-2)" | Korrekte Zuordnung verschiedener Kennzahlen zu Jahren | ✅ Korrigiert und funktioniert nun - MP |
| 4.3 | Mix aus aktuellen und historischen Daten | "TR.EBIT", "TR.Revenue(Period=FY-1)", "ISIN" | Aktuelle + historische Refinitiv-Daten + Excel-Kennzahlen | ✅ |
| 4.4 | Mehrere Unternehmen mit Vorjahresdaten | 3 Unternehmen + "TR.EBIT(Period=FY-1)" | Vorjahresdaten für alle gefilterten Unternehmen | ✅ |
| 4.5 | Durchschnittswerte bei Vorjahresdaten | "TR.EBIT(Period=FY-1)" mit Focus-Filter | Durchschnittswerte basierend auf Vorjahresdaten | ✅ |
| 4.6 | Fehlende Vorjahresdaten | "TR.EBIT(Period=FY-1)" für neues Unternehmen | Robuste Behandlung fehlender historischer Daten | ✅ |
| 4.7 | Große Zeitreihen | 5 Jahre derselben Kennzahl für mehrere Unternehmen | Performance-Test mit vielen Period-Parametern | ✅ |

## 5. Vorjahresdaten-Spezialfälle

| ID | Testfall | Beschreibung | Akzeptanzkriterien | Status |
|----|----------|-------------|-------------------|--------|
| 5.1 | Spaltenname-Generierung | "TR.EBIT(Period=FY-1)" | Spaltenname wird zu "EBIT(Period=FY-1)" | ✅ |
| 5.2 | Doppelte Period-Parameter | "TR.EBIT(Period=FY-1)" zweimal eingegeben | Duplikaterkennung funktioniert | ✅ |
| 5.3 | Unvollständige Daten | Unternehmen hat keine Daten für FY-1 | "N/A" oder leerer Wert, kein Absturz | ✅ |
| 5.4 | Excel vs. Terminal Konsistenz | Vorjahresdaten in beiden Ausgaben | Identische Werte in Excel und Terminal | ✅ |
| 5.5 | Durchschnittswerte-Berechnung | Vorjahresdaten in Durchschnittswerten | Korrekte Berechnung der Sektordurchschnitte | ✅ |
| 5.6 | Große Unternehmen-Anzahl | 50+ Unternehmen mit Vorjahresdaten | Alle Unternehmen haben Vorjahresdaten | ✅ |
| 5.7 | API-Belastung | Viele Period-Parameter gleichzeitig | Effiziente API-Nutzung, keine Timeouts | ✅ |

## 8. Testdaten

### 8.1 Unternehmen für Tests
- Hermes International SCA (HRMS.PA) - Luxury Goods
- Ralph Lauren Corp (RL.N) - Luxury Goods
- Nike Inc (NKE.N) - Footwear
- Netflix Inc (NFLX.O) - Entertainment
- LVMH (LVMH.PA) - Luxury Goods
- Brunello Cucinelli (BCU.MI) - Luxury Goods (High Focus)

### 8.2 Kennzahlen für Tests
- Excel: "ISIN", "P/E", "Free Float", "Market in USD", "Forward P/E"
- Refinitiv: "TR.Revenue", "TR.PriceClose", "TR.TotalReturn"

### 8.3 Vorjahresdaten-Testszenarien

#### 8.3.1 Einzelne Kennzahl über mehrere Jahre
```
Input: 
- RIC: RL.N
- Kennzahlen: TR.EBIT, TR.EBIT(Period=FY-1), TR.EBIT(Period=FY-2)
- Filter: Sub-Industry
```

#### 8.3.2 Verschiedene Kennzahlen mit Period-Parametern
```
Input:
- RIC: HRMS.PA
- Kennzahlen: TR.Revenue(Period=FY-1), TR.EBITDA(Period=FY-2), TR.TotalAssets
- Filter: Focus
```

#### 8.3.3 Mehrere Unternehmen mit Vorjahresdaten
```
Input:
- Unternehmen: Hermes, Ralph Lauren, Nike
- Kennzahlen: ISIN, TR.EBIT(Period=FY-1), TR.Revenue(Period=FY-2)
- Filter: Sub-Industry
```

#### 8.3.4 Zeitreihenanalyse
```
Input:
- RIC: NKE.N
- Kennzahlen: TR.Revenue, TR.Revenue(Period=FY-1), TR.Revenue(Period=FY-2), TR.Revenue(Period=FY-3)
- Filter: Sub-Industry
```

#### 8.3.5 Gemischte Datenquellen mit Vorjahresdaten
```
Input:
- Unternehmen: LVMH, Brunello Cucinelli
- Kennzahlen: ISIN, P/E, TR.EBIT(Period=FY-1), TR.TotalReturn(Period=FY-2)
- Filter: Focus
```

## 9. Testprotokoll

| Datum | Testfall-ID | Getestet von | Ergebnis | Kommentar |
|-------|------------|--------------|----------|------------|
| 2025-07-14 | 3.1 | MP | ✅ | Mix aus allen Filtertypen funktioniert: Hermes (Name+Sub-Industry), RL.N (RIC+Focus), Nike (Name+Standard) |
| 2025-07-14 | 3.2 | MP | ✅ | Mix aus Namen und RICs funktioniert: Hermes (Name), RL.N (RIC), Nike (Name) |
| 2025-07-14 | 3.3 | MP | ✅ | Mix aus Excel und Refinitiv funktioniert: Hermes + 4 Excel + 4 Refinitiv Kennzahlen |
| 2025-07-14 | 3.4 | MP | ✅ | Teilweise fehlerhafte Eingaben verarbeitet: Gültige Einträge übernommen, Fehler protokolliert |
| 2025-07-14 | 3.5 | MP | ✅ | Mehrere Unternehmen + viele Kennzahlen erfolgreich: Vollständige Datenmatrix erstellt |
| 2025-07-14 | 3.6 | MP | ✅ | Überlappende Sub-Industries korrekt gruppiert, keine Duplikate |
| 2025-07-14 | 3.7 | MP | ✅ | Überlappende Focus-Gruppen korrekt gruppiert, keine Duplikate |
| 2025-07-14 | 3.8 | MP | ✅ | Gemischte Kennzahlen mit Fehlern: Gültige verarbeitet, ungültige ignoriert |
| 2025-07-14 | 3.9 | MP | ✅ | Große Kombinationen läuft: 5 Unternehmen + 10 Kennzahlen, Performance-Test |
| YYYY-MM-DD | 1.1.1 | | | |
| YYYY-MM-DD | 1.1.2 | | | |
| YYYY-MM-DD | 1.2.1 | | | |
