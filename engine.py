import pandas as pd
import os
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

class JVEngine:
    """Core logic engine to handle data cleaning, formula calculation, and JV generation."""

    def __init__(self, config=None):
        self.config = config or {}
        self.MONTH_LABEL = self.config.get("MONTH_LABEL", "Feb'26")
        self.MONTH_END_DATE = self.config.get("MONTH_END_DATE", "28022026")
        self.COMPANY_CODE = self.config.get("COMPANY_CODE", 6000)
        self.CURRENCY = self.config.get("CURRENCY", "INR")
        self.DOC_TYPE = self.config.get("DOC_TYPE", "SA")
        self.MAX_LINES_PER_JV = 999
        self.CREDIT_ACCOUNT = 500003
        self.DEBIT_POSTING_KEY = "50"
        self.CREDIT_POSTING_KEY = "40"
        self.COST_CENTER = self.config.get("COST_CENTER", "682490C510")
        self.PROFIT_CENTER = self.config.get("PROFIT_CENTER", "682490C5")

        self.GL_COL_NAMES = ["gl_742234", "gl_742238", "gl_742235", "gl_742236", "gl_742237", "gl_842028"]
        self.GL_CODES = [742234, 742238, 742235, 742236, 742237, 842028]

    @staticmethod
    def d2(val):
        """Round val to exactly 2 decimal places using Decimal."""
        if val is None or pd.isna(val): return 0.0
        try:
            return float(Decimal(str(val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
        except:
            return 0.0

    def calculate_virtual_columns(self, raw_data, df):
        """Internal rules engine for missing columns CB-CM."""
        def s(idx): 
            if idx >= len(raw_data.columns): return pd.Series([0.0]*len(df))
            return pd.to_numeric(raw_data.iloc[:, idx], errors="coerce").fillna(0.0)

        # Derived logic based on user's reference file
        df["billed_status"] = "Billed" 
        df["ic_code"] = raw_data.iloc[:, 80] if len(raw_data.columns) > 80 else "UNKNOWN"
        df["invoice_no"] = raw_data.iloc[:, 82] if len(raw_data.columns) > 82 else ""
        df["emp_no_ref"] = df["workday_id"] # Formula: =A4
        df["cap_center_ref"] = df["cap_center"] # Formula: =G4

        # Financial Formulas (Summing A-AQ raw data)
        df["gl_742234"] = s(14) + s(16) + s(19) + s(20) + s(21) + s(22) + s(23)
        df["gl_742238"] = s(26) + s(30)
        df["gl_742235"] = s(25)
        df["gl_742236"] = s(35)
        df["gl_742237"] = s(17) + s(18)
        df["gl_842028"] = s(34) + s(36)
        return df

    def run_processing(self, filepath, log_callback=print, api_key=None):
        """Full processing pipeline with AI logic mapping."""
        log_callback(f"Loading {os.path.basename(filepath)}...")
        raw = pd.read_excel(filepath, sheet_name="Billing sheet", header=None, dtype=str)
        col_names = raw.iloc[2].tolist()
        data = raw.iloc[3:].copy()
        data.columns = col_names
        data = data.reset_index(drop=True)

        mapping = None
        if api_key:
            from ai_mapper import AIMapper
            log_callback("Connecting to Gemini AI for intelligent mapping...")
            mapper = AIMapper(api_key)
            mapping = mapper.analyze_template(data.head(5))
        
        if mapping:
            log_callback("AI Mapping successful!")
        else:
            log_callback("AI mapping skipped or failed. Using robust fallback rules...")
            # Traditional fallback based on standard indices
            mapping = {
                "workday_id": 0, "cap_center": 6, "legal_entity": 41, "classification": 42,
                "billed_status": 79, "ic_code": 80, "invoice_no": 82, "emp_no_ref": 83, "cap_center_ref": 84
            }

        df = pd.DataFrame()
        df["workday_id"] = data.iloc[:, mapping["workday_id"]]
        df["cap_center"] = data.iloc[:, mapping["cap_center"]]
        df["legal_entity"] = data.iloc[:, mapping["legal_entity"]]
        df["classification"] = data.iloc[:, mapping["classification"]]
        df["billed_status"] = data.iloc[:, mapping["billed_status"]] if "billed_status" in mapping else "Billed"
        df["ic_code"] = data.iloc[:, mapping["ic_code"]] if "ic_code" in mapping else "UNKNOWN"
        df["invoice_no"] = data.iloc[:, mapping["invoice_no"]] if "invoice_no" in mapping else ""
        df["emp_no_ref"] = data.iloc[:, mapping["emp_no_ref"]] if "emp_no_ref" in mapping else df["workday_id"]
        df["cap_center_ref"] = data.iloc[:, mapping["cap_center_ref"]] if "cap_center_ref" in mapping else df["cap_center"]

        # Financial Columns (Check if they exist; if not, calculate)
        if len(data.columns) > 85:
            # Traditional index-based loading
            for i, name in enumerate(self.GL_COL_NAMES):
                df[name] = pd.to_numeric(data.iloc[:, 85+i], errors="coerce").fillna(0.0)
        else:
            log_callback("Missing financial columns (CB-CM). Applying auto-calculation formulas...")
            df = self.calculate_virtual_columns(data, df)

        log_callback("Cleaning and filtering...")
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
        
        df = df[df["workday_id"].notna() & ~df["workday_id"].isin(["nan", "None", ""])]
        df = df[(df["classification"] == "Billable") & (df["billed_status"] == "Billed")]
        df = df[df["invoice_no"].notna() & ~df["invoice_no"].isin(["", "nan", "None"])]
        
        for col in self.GL_COL_NAMES:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            
        df = df.sort_values(by=["invoice_no", "workday_id"]).reset_index(drop=True)
        
        log_callback(f"Building JV rows for {len(df)} employees...")
        rows = self._build_rows(df)
        return rows

    def _build_rows(self, df):
        doc_header = f"Revenue Reclass {self.MONTH_LABEL}"
        all_invoices = df["invoice_no"].unique()
        rows = []
        serial_counter = 1

        for inv in all_invoices:
            inv_df = df[df["invoice_no"] == inv]
            ic_code = inv_df.iloc[0]["ic_code"]
            debits = []
            for _, emp in inv_df.iterrows():
                for gl_col, gl_code in zip(self.GL_COL_NAMES, self.GL_CODES):
                    if abs(emp[gl_col]) < 0.001: continue
                    debits.append({
                        "Reference": 0, "Document Date": self.MONTH_END_DATE, "Document Type": self.DOC_TYPE,
                        "Company Code": self.COMPANY_CODE, "Posting Date": self.MONTH_END_DATE,
                        "Reference.1": inv, "Document Header Text": doc_header, "Currency": self.CURRENCY,
                        "Amount": self.d2(-abs(emp[gl_col])), "Posting Key": self.DEBIT_POSTING_KEY,
                        "Account": gl_code, "Cost Center": self.COST_CENTER, "Profit Center": self.PROFIT_CENTER,
                        "Assignment Number (20)": inv, "Item Text (50)": doc_header,
                        "Ref Key 1": ic_code, "Ref Key 2": emp["cap_center_ref"], "Ref Key 3 (20)": emp["emp_no_ref"],
                        "Inovice Receipt Date": self.MONTH_END_DATE
                    })

            if not debits: continue
            
            # Batching logic
            batch_max = self.MAX_LINES_PER_JV - 1
            i = 0
            while i < len(debits):
                batch = debits[i:i+batch_max]
                batch_credit = sum(Decimal(str(abs(d["Amount"]))) for d in batch)
                
                credit_row = {
                    "Reference": serial_counter, "Document Date": self.MONTH_END_DATE, "Document Type": self.DOC_TYPE,
                    "Company Code": self.COMPANY_CODE, "Posting Date": self.MONTH_END_DATE,
                    "Reference.1": inv, "Document Header Text": doc_header, "Currency": self.CURRENCY,
                    "Amount": self.d2(float(batch_credit)), "Posting Key": self.CREDIT_POSTING_KEY,
                    "Account": 500003, "Cost Center": self.COST_CENTER, "Profit Center": self.PROFIT_CENTER,
                    "Assignment Number (20)": inv, "Item Text (50)": doc_header,
                    "Ref Key 1": ic_code, "Inovice Receipt Date": self.MONTH_END_DATE
                }
                for d in batch: d["Reference"] = serial_counter
                
                rows.append(credit_row)
                rows.extend(batch)
                rows.append({k: None for k in credit_row.keys()}) # spacer
                serial_counter += 1
                i += batch_max
        return rows
