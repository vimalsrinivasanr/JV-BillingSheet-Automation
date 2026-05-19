import openpyxl

manual_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/manual_output/Output Apr'26.xlsx"
wb = openpyxl.load_workbook(manual_path)
ws = wb.active

for r in range(3990, 4005):
    row_vals = [ws.cell(r, c).value for c in range(1, 14)]
    print(f"Row {r}: {row_vals}")
