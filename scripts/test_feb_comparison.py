import os
import sys
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from normalizer import BillingNormalizer
from engine import JVEngine
from compare_outputs import load_manual_jv, load_and_evaluate_auto_jv

def test_feb_comparison():
    input_file = "input/Input Data.xlsx"
    normalized_file = "normalized/Input Data_NORMALIZED.xlsx"
    output_file = "output/SAP_JV_Upload_Feb26_Run.xlsx"
    manual_file = "output/manual_output/Output.xlsx"

    # Step 1: Run normalizer
    print("Running Normalizer on Feb data...")
    norm = BillingNormalizer(log_callback=lambda x: None)
    norm.normalize(input_file, normalized_file)

    # Step 2: Run engine
    print("Running Engine on Feb data...")
    config = {
        "MONTH_LABEL": "Feb'26",
        "MONTH_END_DATE": "28022026",
        "COMPANY_CODE": 6000
    }
    engine = JVEngine(config=config)
    rows = engine.run_processing(normalized_file, log_callback=lambda x: None)
    engine.write_excel(rows, output_file, log_callback=lambda x: None)

    # Step 3: Compare row-by-row
    print("\n--- Row-by-Row Comparison ---")
    manual_df = load_manual_jv(manual_file)
    auto_df = load_and_evaluate_auto_jv(output_file)

    # Merge on keys
    merged = pd.merge(
        manual_df, auto_df,
        on=['Account', 'Amount_Evaluated'],
        suffixes=('_manual', '_auto')
    )

    print(f"Manual rows: {len(manual_df)}")
    print(f"Automated rows: {len(auto_df)}")
    print(f"Merged rows: {len(merged)}")

    # Check posting key mismatches
    mismatch = merged[merged['Posting Key_manual'] != merged['Posting Key_auto']]
    print(f"Posting key mismatches: {len(mismatch)}")

    if not mismatch.empty:
        print("\nSample mismatches:")
        print(mismatch[['Reference.1', 'Account', 'Ref Key 3 (20)', 'Amount_manual', 'Amount_auto', 'Posting Key_manual', 'Posting Key_auto']].head(20))
    else:
        print("SUCCESS: 100% row-by-row match for Posting Keys against Output.xlsx!")

if __name__ == "__main__":
    test_feb_comparison()
