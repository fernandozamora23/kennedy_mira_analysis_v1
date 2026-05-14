import pandas as pd
import numpy as np
df = pd.DataFrame({'a': [1, 2, np.nan], 'b': ['x', 'y', np.nan]})
for col in df.columns:
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = df[col].map(lambda x: str(x))
    else:
        df[col] = df[col].fillna("").astype(str)
print(df)
