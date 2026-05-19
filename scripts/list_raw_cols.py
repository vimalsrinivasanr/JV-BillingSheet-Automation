import sys
import pandas as pd
path = sys.argv[1]
df = pd.read_excel(path, sheet_name=0, dtype=str)
print('\n'.join(df.columns.tolist()))
print('Rows:', len(df))
