import sys
import pandas as pd
df = pd.read_excel('data/kennedy_mira_consolidado.xlsx', sheet_name='MESAS_BARRIO_UPZ')
print(list(df.columns))
print(df[df['IGLESIA'] == 'VALLADOLID'][['MESA_ID', 'IGLESIA']].head())
