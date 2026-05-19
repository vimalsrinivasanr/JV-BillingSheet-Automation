import sys
import pandas as pd
path = sys.argv[1]
df = pd.read_excel(path, sheet_name=0, dtype=str)
for c in df.columns:
    if 'recharge' in str(c).lower() or 'mark' in str(c).lower():
        print('---', c)
        print(df[c].head(10).tolist())
print('Rows:', len(df))
