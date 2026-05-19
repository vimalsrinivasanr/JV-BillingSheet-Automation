import pandas as pd
import openpyxl

def load_jv_file(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    sheet_name = wb.sheetnames[0]
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    if "Reference" not in df.columns:
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=2)
    df_clean = df[df["Reference"].notna()].copy()
    df_clean["Reference.1"] = df_clean["Reference.1"].astype(str).str.strip()
    return df_clean

def inspect_invoice(invoice_no):
    manual_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/manual_output/Output Apr'26.xlsx"
    auto_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/SAP_JV_Upload_Run.xlsx"

    manual_df = load_jv_file(manual_path)
    auto_df = load_jv_file(auto_path)

    m_rows = manual_df[manual_df["Reference.1"] == invoice_no]
    a_rows = auto_df[auto_df["Reference.1"] == invoice_no]

    with open("scripts/invoice_comparison.txt", "w") as f:
        f.write(f"=== Manual rows for {invoice_no} ({len(m_rows)} rows) ===\n")
        f.write(m_rows[["Ref Key 3 (20)", "Posting Key", "Account", "Amount"]].to_string() + "\n\n")

        f.write(f"=== Auto rows for {invoice_no} ({len(a_rows)} rows) ===\n")
        f.write(a_rows[["Ref Key 3 (20)", "Posting Key", "Amount", "Account", "Amount"]].to_string() + "\n")

if __name__ == "__main__":
    inspect_invoice("30001435.0")
