import pandas as pd
import os

orig_path = "/Users/fernandozamora/Documents/Kennedy/CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx"
cons_path = "/Users/fernandozamora/Documents/Kennedy/kennedy_mira_consolidado_actualizado_tecnologia.xlsx"
out_path = "/Users/fernandozamora/Documents/Kennedy/kennedy_mira_analysis/data/kennedy_mira_consolidado.xlsx"

# 1. Build mapping from original
mapping = {}
xls = pd.ExcelFile(orig_path)
for sheet in ["AGENDA GENERAL CON CANDIDATOS", "AGENDA PARALELA", "Cronograma Kennedy enero"]:
    df = pd.read_excel(orig_path, sheet_name=sheet)
    # The columns might have leading/trailing spaces
    col_obs = None
    col_est = None
    for c in df.columns:
        c_upper = str(c).strip().upper()
        if "DETALLE" in c_upper or "OBSERVACION" in c_upper:
            col_obs = c
        if "ESTRATEGIA" in c_upper:
            col_est = c
    if col_obs and col_est:
        for _, row in df.iterrows():
            obs = str(row[col_obs]).strip()
            est = str(row[col_est]).strip().upper()
            if obs and obs != "nan":
                mapping[obs] = est

# 2. Update consolidated
xls_cons = pd.ExcelFile(cons_path)
sheets = {}
for sheet in xls_cons.sheet_names:
    sheets[sheet] = pd.read_excel(cons_path, sheet_name=sheet)

act_df = sheets["actividades_campana"]

# Function to map based on observation
def get_estrategia(obs):
    obs = str(obs).strip()
    return mapping.get(obs, "N/A")

act_df["ESTRATEGIA"] = act_df["OBSERVACIONES"].apply(get_estrategia)
sheets["actividades_campana"] = act_df

# 3. Save
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    for sheet, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet, index=False)

# Also update the actual "tecnologia" file in parent dir to keep it consistent
with pd.ExcelWriter(cons_path, engine="openpyxl") as writer:
    for sheet, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet, index=False)

print("Updated successfully")
