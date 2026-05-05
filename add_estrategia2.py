import pandas as pd

orig_path = "data/CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx"
cons_path = "data/kennedy_mira_consolidado.xlsx"

print("Cargando mapeo de estrategias...")
mapping = {}
xls = pd.ExcelFile(orig_path)
for sheet in ["AGENDA GENERAL CON CANDIDATOS", "AGENDA PARALELA", "Cronograma Kennedy enero"]:
    if sheet in xls.sheet_names:
        df = pd.read_excel(orig_path, sheet_name=sheet)
        
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

print(f"Mapeadas {len(mapping)} estrategias.")

print("Actualizando consolidado...")
xls_cons = pd.ExcelFile(cons_path)
sheets = {}
for sheet in xls_cons.sheet_names:
    sheets[sheet] = pd.read_excel(cons_path, sheet_name=sheet)

act_df = sheets["actividades_campana"]

def get_estrategia(obs):
    obs = str(obs).strip()
    return mapping.get(obs, "AFINIDAD") # Default a afinidad para que las demas se filtren o N/A

act_df["ESTRATEGIA"] = act_df["OBSERVACIONES"].apply(get_estrategia)
sheets["actividades_campana"] = act_df

# Verificar cuantas quedaron de politico comunitaria y libertad religiosa
counts = act_df["ESTRATEGIA"].value_counts()
print("\nDistribución final de ESTRATEGIA en actividades:")
print(counts)

valid_count = act_df["ESTRATEGIA"].isin(["LIBERTAD RELIGIOSA", "POLITICO COMUNITARIA", "POLÍTICO COMUNITARIA"]).sum()
print(f"\nActividades que se mostrarán en el dashboard: {valid_count}")

print("\nGuardando Excel consolidado actualizado...")
with pd.ExcelWriter(cons_path, engine="openpyxl") as writer:
    for sheet, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet, index=False)

print("¡Listo! El archivo kennedy_mira_consolidado.xlsx ha sido actualizado.")
