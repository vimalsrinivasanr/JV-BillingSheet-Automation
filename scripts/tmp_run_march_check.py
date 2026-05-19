from normalizer import BillingNormalizer
from engine import JVEngine
import os
import pandas as pd

base = '/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation'
inp = os.path.join(base, 'Input Data March.xlsx')

n = BillingNormalizer(log_callback=print)
normalized_path, report = n.normalize(inp)
print('NORMALIZED_PATH=', normalized_path)

config = {
    'MONTH_LABEL': 'Mar26',
    'MONTH_END_DATE': '31032026',
    'COMPANY_CODE': 6000,
    'API_KEY': ''
}
eng = JVEngine(config)
rows = eng.run_processing(normalized_path, log_callback=print)
out = os.path.join(base, 'SAP_JV_Upload_Mar26_TEST.xlsx')
eng.write_excel(rows, out, log_callback=print)
print('JV_OUTPUT_PATH=', out)
print('JV_ROWS=', len(rows))

jv = pd.read_excel(out, sheet_name='JV', header=2)
amount = pd.to_numeric(jv['Amount'], errors='coerce').fillna(0)
non_empty = jv[~jv['Reference'].isna()]
print('DATA_ROWS_WITH_REFERENCE=', len(non_empty))
print('AMOUNT_SUM=', round(float(amount.sum()), 2))
print('UNIQUE_INVOICES=', jv['Reference.1'].dropna().astype(str).nunique())
