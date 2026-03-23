from pathlib import Path
import pandas as pd
filepath = Path("工作簿2.csv")
df_raw = pd.read_csv(filepath, low_memory=False)