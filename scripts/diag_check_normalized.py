import sys
import pandas as pd

if len(sys.argv) < 2:
    print("Usage: python scripts/diag_check_normalized.py <path_to_normalized_xlsx>")
    sys.exit(1)

path = sys.argv[1]
print(f"Inspecting: {path}")
try:
    df = pd.read_excel(path, sheet_name='Normalized', dtype=str)
except Exception as e:
    print(f"Failed to read Normalized sheet: {e}")
    sys.exit(1)

# Columns we expect
gl_cols = ['Recharge - Payroll','Recharge - Manager','Recharge - Leadership','Recharge - Desk Cost','Recharge - Retirals','Mark up']
for c in gl_cols:
    if c not in df.columns:
        print(f"MISSING column in Normalized: {c}")

print(f"Total rows in Normalized: {len(df)}")

# Convert GL cols to numeric and summarize
for c in gl_cols:
    if c in df.columns:
        nums = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        nz = (nums.abs() > 0.009).sum()
        s = nums.sum()
        print(f"{c}: non-zero count={nz}, sum={s}")

# Count invoices with any non-zero GL total
if 'Invoice No.' in df.columns:
    gl_sum = sum((pd.to_numeric(df.get(c, 0), errors='coerce').fillna(0.0) for c in gl_cols))
    inv_with_amount = df.loc[gl_sum.abs() > 0.009, 'Invoice No.'].nunique()
    print(f"Unique invoices with any GL amount: {inv_with_amount}")
else:
    print("Invoice No. column not present in Normalized sheet.")

print("Done.")
