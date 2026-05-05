import pandas as pd
xls = pd.ExcelFile("data/kennedy_mira_consolidado.xlsx")
for sheet in xls.sheet_names:
    df = pd.read_excel("data/kennedy_mira_consolidado.xlsx", sheet_name=sheet)
    for c in df.columns:
        try:
            vals = set(df[c].dropna().astype(str).str.strip().str.upper().unique())
            if "LIBERTAD RELIGIOSA" in vals or "POLITICO COMUNITARIA" in vals:
                print(f"MATCH IN SHEET: {sheet}, COLUMN: {c}")
        except Exception as e:
            pass
