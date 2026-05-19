import sys
import pandas as pd
p = sys.argv[1]
try:
    # header is now on the first row
    df = pd.read_excel(p, sheet_name='JV', header=0)
except Exception as e:
    print('Failed to read JV sheet:', e); sys.exit(1)
if 'Amount' not in df.columns:
    print('Amount column not found. Columns:', df.columns.tolist()); sys.exit(1)
vals = df['Amount'].fillna(0.0).astype(float)
print('Total rows in JV data:', len(vals))
print('Sample Amounts (first 40):')
print(vals.head(40).tolist())
print('Non-zero counts: positive=', (vals>0).sum(), 'negative=', (vals<0).sum())
print('Sum of Amounts:', vals.sum())
