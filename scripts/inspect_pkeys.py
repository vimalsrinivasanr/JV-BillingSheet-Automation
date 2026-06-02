import openpyxl
from collections import Counter

manual_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/manual_output/Output.xlsx"
wb = openpyxl.load_workbook(manual_path, data_only=False)
ws = wb.active

pkeys = []
accounts = []
formulas_in_pkey = []
matching_patterns = []

for r in range(4, 200): # Sample first 200 rows (skipping header rows 1-3)
    pk_val = ws.cell(row=r, column=11).value # Column K is 11
    acc_val = ws.cell(row=r, column=12).value # Column L is 12
    amt_val = ws.cell(row=r, column=10).value # Column J is 10
    
    if pk_val is not None:
        pkeys.append(str(pk_val))
        if str(pk_val).startswith('='):
            formulas_in_pkey.append((r, pk_val, acc_val, amt_val))
        else:
            matching_patterns.append((pk_val, acc_val, amt_val))
        accounts.append(str(acc_val))

print("=== Posting Key analysis of Output.xlsx ===")
print("Unique Posting Keys found in first 200 rows:", Counter(pkeys))
print("Unique Accounts found in first 200 rows:", Counter(accounts))
print(f"Formulas found in Posting Key column: {len([x for x in pkeys if x.startswith('=')])}")
if formulas_in_pkey:
    print("\nSample formulas in Posting Key column:")
    for row, formula, acc, amt in formulas_in_pkey[:10]:
        print(f"Row {row}: PKey={formula} | Account={acc} | Amount={amt}")
else:
    print("\nNo formulas in Posting Key column. Here are some sample rows (PKey, Account, Amount):")
    for pk, acc, amt in matching_patterns[:15]:
        print(f"PKey={pk} | Account={acc} | Amount={amt}")
