import os
import sys

# Ensure scripts folder is on path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from normalizer import BillingNormalizer
from engine import JVEngine

def run_april():
    input_file = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/Input Apr'26.xlsx"
    normalized_file = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/Input Apr'26_NORMALIZED.xlsx"
    output_file = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/SAP_JV_Upload_Run.xlsx"

    print("=== [1/2] RUNNING STAGE 1: NORMALIZER ===")
    # Initialize normalizer with a custom print logger
    norm = BillingNormalizer(log_callback=print)
    
    # We monkey-patch the built-in input() function in normalizer if it asks for sheet selection.
    # Let's see if the file has multiple sheets first.
    import openpyxl
    wb = openpyxl.load_workbook(input_file, read_only=True)
    sheet_names = wb.sheetnames
    print(f"Sheet names in input file: {sheet_names}")
    
    # Let's mock input to select 'Billing sheet' or first sheet
    import builtins
    original_input = builtins.input
    builtins.input = lambda prompt: "1"  # Auto select first sheet if asked
    
    try:
        norm.normalize(input_file, normalized_file)
    finally:
        builtins.input = original_input

    print("\n=== [2/2] RUNNING STAGE 2: JV ENGINE ===")
    config = {
        "MONTH_LABEL": "Apr'26",
        "MONTH_END_DATE": "30042026",
        "COMPANY_CODE": 6000,
        "CURRENCY": "INR",
        "DOC_TYPE": "SA",
        "COST_CENTER": "682490C510",
        "PROFIT_CENTER": "682490C5",
    }
    engine = JVEngine(config=config)
    rows = engine.run_processing(normalized_file, log_callback=print)
    engine.write_excel(rows, output_file, log_callback=print)
    print(f"\nDone! Output written to {output_file}")

if __name__ == "__main__":
    run_april()
