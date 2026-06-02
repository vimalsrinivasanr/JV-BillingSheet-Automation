"""
Model: JVGenerator
Handles JV file generation logic from normalized billing data.
"""

import os
import tempfile
import pandas as pd
from scripts.engine import JVEngine

class JVGenerator:
    def __init__(self, billing_df):
        self.billing_df = billing_df
        self.jv_df = None
        self.engine = None
        self.rows = None

    def generate(self):
        if self.billing_df is None or self.billing_df.empty:
            return None

        # Configuration for engine
        config = {
            "MONTH_LABEL": "Feb'26",
            "MONTH_END_DATE": "28022026",
            "COMPANY_CODE": 6000
        }
        self.engine = JVEngine(config)

        # Write self.billing_df to a temporary Excel file under sheet "Normalized"
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file = os.path.join(temp_dir, "temp_normalized.xlsx")
            with pd.ExcelWriter(temp_file, engine="openpyxl") as writer:
                self.billing_df.to_excel(writer, sheet_name="Normalized", index=False)
            
            # Run engine processing to get rows
            self.rows = self.engine.run_processing(temp_file, log_callback=lambda x: None)
            
        # Also build a DataFrame from the rows for standard save fallback
        cols = [
            "Reference", "Document Date", "Document Type", "Company Code", "Posting Date",
            "Reference.1", "Document Header Text", "Currency", "Exchange rate", "Amount",
            "Posting Key", "Account", "Special G/L ind.", "Cost Center", "Internal Order",
            "Profit Center", "Business Area", "Assignment Number (20)", "Item Text (50)",
            "Ref Key 1", "Ref Key 2", "Ref Key 3 (20)", "Material", "Trading Partner",
            "Tax Code", "Withholding tax code", "Withholding tax base amount in document currency",
            "Customer", "Contracts", "Revenue Period", "Core Consultant", "Revenue Month",
            "Reversal Date", "LEDGER", "WT CODE1", "WT Amount", "Inovice Receipt Date",
        ]
        if self.rows:
            # We filter out empty/blank separator rows for the raw df representation
            valid_rows = [r for r in self.rows if any(v is not None for v in r.values())]
            self.jv_df = pd.DataFrame(valid_rows)
            # Ensure columns are named correctly
            self.jv_df.columns = cols[:len(self.jv_df.columns)]
        else:
            self.jv_df = pd.DataFrame(columns=cols)

        return None

    def save(self, output_path):
        if self.engine is not None and self.rows is not None:
            self.engine.write_excel(self.rows, output_path, log_callback=lambda x: None)
        elif self.jv_df is not None:
            self.jv_df.to_excel(output_path, index=False)
