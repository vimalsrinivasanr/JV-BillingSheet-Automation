import pandas as pd
p = r'C:\Users\vimalsrinivasan.r\Desktop\JV-BillingSheet-Automation\output\input file\SAP_JV_Upload_Run.xlsx'
print(pd.ExcelFile(p).sheet_names)
