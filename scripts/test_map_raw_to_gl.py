import pandas as pd
import sys, os
norm = sys.argv[1]
raw = sys.argv[2]
ndf = pd.read_excel(norm, sheet_name='Normalized', dtype=str)
rdf = pd.read_excel(raw, sheet_name=0, dtype=str)
print('Normalized rows', len(ndf))
print('Raw rows', len(rdf))
raw_cols = [str(c).strip().lower() for c in rdf.columns]
print('raw cols sample:', raw_cols)

def pick_raw(kw_list):
    for i,h in enumerate(raw_cols):
        for kw in kw_list:
            if kw in h:
                s = pd.to_numeric(rdf.iloc[:,i], errors='coerce').fillna(0.0)
                print('matched', rdf.columns[i], 'non-zero count', (s.abs()>0.009).sum(), 'sum', s.sum())
                return s
    print('no match for', kw_list)
    return pd.Series([0.0]*len(ndf))

p = pick_raw(['payroll','recharge  general cost - payroll','recharge payroll'])
print('p head', p.head(12).tolist())
