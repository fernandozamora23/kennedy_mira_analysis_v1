import pandas as pd
path = "data/kennedy_mira_consolidado.xlsx"
df = pd.read_excel(path, sheet_name="resumen_general")
print(df)
