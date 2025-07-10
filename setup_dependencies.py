#!/usr/bin/env python3
"""
Setup-Skript zur Überprüfung und Installation aller notwendigen Abhängigkeiten
für das Excel-Datenverarbeitungsprojekt mit Refinitiv-Integration.
"""

import sys
import subprocess
import importlib
import os

def check_python_version():
    """Überprüft die Python-Version"""
    print("🔍 Überprüfe Python-Version...")
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 oder höher ist erforderlich!")
        print(f"   Aktuelle Version: {sys.version}")
        return False
    print(f"✅ Python-Version OK: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    return True

def install_package(package_name, import_name=None):
    """Installiert ein Python-Package"""
    if import_name is None:
        import_name = package_name

    try:
        importlib.import_module(import_name)
        print(f"✅ {package_name} ist bereits installiert")
        return True
    except ImportError:
        print(f"📦 Installiere {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"✅ {package_name} erfolgreich installiert")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Fehler beim Installieren von {package_name}: {e}")
            return False

def test_package_functionality():
    """Testet die Funktionalität der installierten Pakete"""
    print("\n🧪 Teste Paket-Funktionalität...")

    # Test Pandas
    try:
        import pandas as pd
        df = pd.DataFrame({'test': [1, 2, 3]})
        print("✅ Pandas funktioniert korrekt")
    except Exception as e:
        print(f"❌ Pandas-Test fehlgeschlagen: {e}")
        return False

    # Test Openpyxl (für Excel)
    try:
        import openpyxl
        print("✅ Openpyxl (Excel-Support) funktioniert korrekt")
    except Exception as e:
        print(f"❌ Openpyxl-Test fehlgeschlagen: {e}")
        return False

    # Test Refinitiv Data
    try:
        import refinitiv.data as rd
        print("✅ Refinitiv Data Library importiert")

        # Test ob Desktop Session verfügbar ist
        try:
            session = rd.open_session()
            if session:
                print("✅ Refinitiv Desktop Session erfolgreich geöffnet")
                rd.close_session()
            else:
                print("⚠️  Refinitiv Desktop Session konnte nicht geöffnet werden")
                print("    Stelle sicher, dass Refinitiv Workspace läuft")
        except Exception as e:
            print(f"⚠️  Refinitiv Session-Test: {e}")
            print("    Das ist normal, wenn Refinitiv Workspace nicht läuft")

    except Exception as e:
        print(f"❌ Refinitiv Data-Test fehlgeschlagen: {e}")
        return False

    # Test xlsxwriter (für schöne Excel-Ausgaben)
    try:
        import xlsxwriter
        print("✅ XlsxWriter funktioniert korrekt")
    except Exception as e:
        print(f"❌ XlsxWriter-Test fehlgeschlagen: {e}")
        return False

    return True

def check_refinitiv_workspace():
    """Überprüft ob Refinitiv Workspace läuft"""
    print("\n🔍 Überprüfe Refinitiv Workspace...")

    try:
        import refinitiv.data as rd
        session = rd.open_session()
        if session:
            print("✅ Refinitiv Workspace läuft und ist verbunden")
            rd.close_session()
            return True
        else:
            print("❌ Refinitiv Workspace ist nicht verbunden")
            return False
    except Exception as e:
        print(f"❌ Refinitiv Workspace-Fehler: {e}")
        return False

def main():
    """Hauptfunktion zur Ausführung der Setup-Checks"""
    print("🔧 Setup-Überprüfung für Excel-Datenverarbeitungsprojekt\n")

    # Prüfe Python-Version
    if not check_python_version():
        print("\n❌ Setup fehlgeschlagen: Python-Version nicht kompatibel")
        return False

    # Installiere Abhängigkeiten
    packages = [
        ("pandas", None),
        ("openpyxl", None),
        ("xlsxwriter", None),
        ("refinitiv-data", "refinitiv.data")  # Füge refinitiv-data Paket hinzu
    ]

    success = True
    for package_name, import_name in packages:
        if not install_package(package_name, import_name):
            success = False

    if not success:
        print("\n⚠️ Einige Pakete konnten nicht installiert werden")
        print("   Bitte installiere sie manuell mit pip install <paketname>")

    # Teste Funktionalität
    if not test_package_functionality():
        print("\n⚠️ Einige Pakete funktionieren nicht wie erwartet")

    # Prüfe Refinitiv Workspace
    check_refinitiv_workspace()

    print("\n✅ Setup abgeschlossen!")
    return True

if __name__ == "__main__":
    main()
