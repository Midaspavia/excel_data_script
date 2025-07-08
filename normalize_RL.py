import os
import pandas as pd
import refinitiv.data as rd

DATA_DIR = "data/excel_data"

def replace_rl_with_rln():
    print("🔧 Suche nach 'RL' und ersetze mit 'RL.N' in allen Dateien...")

    for file in os.listdir(DATA_DIR):
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue

        path = os.path.join(DATA_DIR, file)
        xls = pd.ExcelFile(path)
        writer = pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace")

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            changed = False

            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    value = str(df.iat[row_idx, col_idx])
                    if value.strip().upper() == "RL":
                        df.iat[row_idx, col_idx] = "RL.N"
                        print(f"✅ Ersetzt 'RL' → 'RL.N' in Datei '{file}', Sheet '{sheet_name}', Zelle ({row_idx+1}, {col_idx+1})")
                        changed = True

            if changed:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        writer.close()

if __name__ == "__main__":
    replace_rl_with_rln()