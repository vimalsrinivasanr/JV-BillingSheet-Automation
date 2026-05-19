import sys
import pandas as pd
path = sys.argv[1]
df = pd.read_excel(path, sheet_name='Normalized', dtype=str)
cols = ['Fixed CTC','Total billable amount','Recharge - Payroll','Recharge - Manager','Recharge - Leadership','Recharge - Desk Cost','Recharge - Retirals','Mark up']
for c in cols:
    print(f"-- {c} --")
    if c in df.columns:
        vals = df[c].head(10).tolist()
        print(vals)
    else:
        print('MISSING')
print('Rows:', len(df))
