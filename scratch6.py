import pandas as pd
xls = pd.ExcelFile("CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx")
for sheet in ['AGENDA GENERAL CON CANDIDATOS', 'AGENDA PARALELA', 'Cronograma Kennedy enero']:
    df = pd.read_excel("CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx", sheet_name=sheet)
    print(f"--- {sheet} ---")
    print(df.columns.tolist())
