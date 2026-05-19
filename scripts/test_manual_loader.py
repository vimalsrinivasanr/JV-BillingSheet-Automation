import importlib.util
import pandas as pd
from pathlib import Path

# Import JVEngine directly from scripts/engine.py by path to avoid package import issues
spec = importlib.util.spec_from_file_location("engine_mod", "scripts/engine.py")
engine_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine_mod)
JVEngine = engine_mod.JVEngine

print('Creating sample manual input...')
path = Path('scripts/test_manual_input.xlsx')

df = pd.DataFrame({
    'Workday ID': ['W1','W2'],
    'Classification': ['Billable','Billable'],
    'Billed/ Unbilled': ['Billed','Billed'],
    'Invoice No.': ['INV1','INV1'],
    'EmpNo': ['E1','E2'],
    'Capability Center': ['CC1','CC1'],
    'Recharge - Payroll': [100.0,200.0],
    'Recharge - Manager': [0.0,50.0],
    'Recharge - Leadership': [0.0,0.0],
    'Recharge - Desk Cost': [0.0,0.0],
    'Recharge - Retirals': [0.0,0.0],
    'Mark up': [10.0,20.0],
})

df.to_excel(path, index=False)
print(f'Wrote sample file: {path}')

engine = JVEngine({
    'MONTH_LABEL': "Apr'26",
    'MONTH_END_DATE': '30042026',
})

print('Running engine...')
rows = engine.run_processing(str(path), print)
print(f'Rows produced: {len(rows)}')

out = Path('scripts/test_output.xlsx')
engine.write_excel(rows, str(out), print)
print(f'Output written: {out}')
