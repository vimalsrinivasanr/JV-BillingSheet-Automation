import pandas as pd
import numpy as np

input_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/Input Apr'26.xlsx"
manual_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/manual_output/Output Apr'26.xlsx"

print("--- Reading Input File ---")
try:
    xls_in = pd.ExcelFile(input_path)
    print("Sheets in Input:", xls_in.sheet_names)
    df_in = pd.read_excel(input_path, sheet_name=0)
    print("Input Shape:", df_in.shape)
    print("Input Columns:", list(df_in.columns)[:20])
except Exception as e:
    print("Error reading input:", e)

print("\n--- Reading Manual Output File ---")
try:
    xls_out = pd.ExcelFile(manual_path)
    print("Sheets in Output:", xls_out.sheet_names)
    df_out = pd.read_excel(manual_path, sheet_name=0, header=2)
    print("Output Shape:", df_out.shape)
    print("Output Columns:", list(df_out.columns))
    # Let's inspect unique values in Column J (Amount) and see if there are negative/positive values
    amt_col = df_out.iloc[:, 9] # Column J is index 9
    print("First 20 Amount values in manual output:")
    print(amt_col.head(20))
    print("Amount description in manual output:")
    print(pd.to_numeric(amt_col, errors='coerce').describe())
except Exception as e:
    print("Error reading output:", e)
