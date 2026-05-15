import pandas as pd
from app import kennedy_dataframe
df = pd.DataFrame({'A': [1, 2, 3], 'B': ['x', 'y', 'z']})
try:
    kennedy_dataframe(df)
    print("SUCCESS")
except Exception as e:
    print("FAILED", e)
