#!/usr/bin/env python3
"""
Test 1.1.1: Gültiger RIC (Spalte B)
Testet ob ein gültiger RIC korrekt erkannt und verarbeitet wird.
"""

import pandas as pd
import os
import sys
from controller import find_company_by_ric

def create_test_input_1_1_1():
    """Erstelle Test-Input für 1.1.1: Gültiger RIC"""
    data = {
        'Name': ['', ''],  # Spalte A leer
        'RIC': ['RL.N', ''],  # Spalte B mit gültigem RIC
        'Sub-Industry': ['X', ''],  # Sub-Industry Filter
        'Focus': ['', ''],
        'Kennzahlen aus Excel': ['ISIN', 'P/E'],
        'Kennzahlen aus Refinitiv': ['', '']
    }

    df = pd.DataFrame(data)
    df.to_excel('excel_data/input_user.xlsx', index=False)
    print("✅ Test-Input 1.1.1 erstellt: Gültiger RIC 'RL.N' in Spalte B")

def test_1_1_1():
    """Führe Test 1.1.1 aus - nur RIC-Erkennung testen"""
    print("\n🧪 STARTE TEST 1.1.1: Gültiger RIC (Spalte B)")
    print("="*60)

    create_test_input_1_1_1()

    try:
        # Teste nur die RIC-Erkennung, nicht die gesamte Verarbeitung
        print("🔍 Teste RIC-Erkennung für 'RL.N'...")
        company = find_company_by_ric('RL.N')

        if company:
            if 'Ralph Lauren' in company.get('Name', '') and company.get('RIC') == 'RL.N':
                print("✅ TEST 1.1.1 ERFOLGREICH:")
                print(f"   - Name: {company['Name']}")
                print(f"   - RIC: {company['RIC']}")
                print(f"   - Sub-Industry: {company.get('Sub-Industry', 'N/A')}")
                print(f"   - Focus: {company.get('Focus', 'N/A')}")
                return True
            else:
                print("❌ TEST 1.1.1 FEHLGESCHLAGEN:")
                print(f"   - Gefundener Name: {company.get('Name', '')}")
                print(f"   - Gefundener RIC: {company.get('RIC', '')}")
                return False
        else:
            print("❌ TEST 1.1.1 FEHLGESCHLAGEN: RIC 'RL.N' nicht gefunden")
            return False

    except Exception as e:
        print(f"❌ TEST 1.1.1 FEHLGESCHLAGEN: Exception - {e}")
        return False

if __name__ == "__main__":
    success = test_1_1_1()
    sys.exit(0 if success else 1)
