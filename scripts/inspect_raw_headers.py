import pandas as pd

manual_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/manual_output/Output Apr'26.xlsx"
df_raw = pd.read_excel(manual_path, sheet_name=0, header=None).iloc[:15]
print(df_raw.to_string())
