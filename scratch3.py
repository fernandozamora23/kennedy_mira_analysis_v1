import pandas as pd
df = pd.read_excel("data/kennedy_mira_consolidado.xlsx", sheet_name="actividades_campana")
print("COLUMNS:")
print(df.columns.tolist())
for c in df.columns:
    try:
        vals = set(df[c].dropna().astype(str).unique())
        if "LIBERTAD RELIGIOSA" in vals or "POLITICO COMUNITARIA" in vals:
            print("MATCH COLUMN:", c)
    except Exception as e:
        pass
