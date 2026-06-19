import os
import tempfile
import pandas as pd
import pytest
from scripts.normalizer import BillingNormalizer
from scripts.engine import JVEngine

def test_standalone_billable_cost():
    # 1. Create a mock input Excel file containing only a Billable Cost sheet
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, "Inputs_Billable_Cost.xlsx")
        
        bc_data = {
            "Ref Key 1": ["IC Code", "US_SAP_L", "US_SAP_L"],
            "Assignment Number (20)": ["Invoice No.", "30001507", "30001507"],
            "742118": ["Billable Cost", "-1000", "-500"],
            "Ref Key 3": ["Inv/ Bill No.", "290982", "290982"],
            "Ref Key 2": ["Billable Cost", "Billable Cost", "Billable Cost"]
        }
        df_bc = pd.DataFrame(bc_data)
        with pd.ExcelWriter(input_path, engine="openpyxl") as writer:
            df_bc.to_excel(writer, sheet_name="Billable Cost", index=False)
            
        # 2. Run Normalizer
        norm = BillingNormalizer(log_callback=lambda x: None)
        normalized_path, _ = norm.normalize(input_path)
        
        assert os.path.exists(normalized_path)
        # Check sheet names
        xls = pd.ExcelFile(normalized_path)
        assert "Billable Cost" in xls.sheet_names
        
        # 3. Run JVEngine in standalone mode with start serial number 46
        config = {
            "MONTH_LABEL": "May'26",
            "MONTH_END_DATE": "31052026",
            "COMPANY_CODE": 6000,
            "START_SERIAL_NO": 46
        }
        engine = JVEngine(config=config)
        rows = engine.run_processing(normalized_path, log_callback=lambda x: None)
        
        # We expect:
        # - 1 credit row (Account 500003, Posting Key 40, Amount +1500)
        # - 1 debit row (Account 742118, Posting Key 50, Amount -1500, Ref Key 2 = "290982", Ref Key 3 = "Billable Cost")
        # - 1 spacer row
        assert len(rows) == 3
        
        # Credit row
        cr = rows[0]
        assert cr["Reference"] == 46
        assert cr["Account"] == 500003
        assert cr["Posting Key"] == "40"
        
        # Debit row
        db = rows[1]
        assert db["Reference"] == 46
        assert db["Account"] == 742118
        assert db["Posting Key"] == "50"
        assert db["Amount"] == -1500.0
        assert db["Ref Key 1"] == "US_SAP_L"
        assert db["Ref Key 2"] == "290982"
        assert db["Ref Key 3 (20)"] == "Billable Cost"
        
        # Spacer row
        sp = rows[2]
        assert all(v is None for v in sp.values())

def test_combined_billing_and_billable_cost():
    # Create a mock input containing both Normalized sheet and Billable Cost sheet
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, "Combined_Normalized.xlsx")
        
        # 1. Standard Normalized data
        std_data = {
            "Workday ID": ["W0001"],
            "Capability Center": ["HR"],
            "Legal Entity": ["Entity1"],
            "Classification": ["Billable"],
            "Billed/ Unbilled": ["Billed"],
            "IC Code": ["IC1"],
            "Invoice No.": ["INV1"],
            "EmpNo (ref)": ["E1"],
            "Capability Center (ref)": ["HR"],
            "Recharge - Payroll": [1000.0],
            "Recharge - Manager": [0.0],
            "Recharge - Leadership": [0.0],
            "Recharge - Desk Cost": [0.0],
            "Recharge - Retirals": [0.0],
            "Mark up": [0.0],
        }
        df_std = pd.DataFrame(std_data)
        
        # 2. Billable Cost data
        bc_data = {
            "Ref Key 1": ["IC Code", "US_SAP_L"],
            "Assignment Number (20)": ["Invoice No.", "30001507"],
            "742118": ["Billable Cost", "-2000"],
            "Ref Key 3": ["Inv/ Bill No.", "290982"],
            "Ref Key 2": ["Billable Cost", "Billable Cost"]
        }
        df_bc = pd.DataFrame(bc_data)
        
        with pd.ExcelWriter(input_path, engine="openpyxl") as writer:
            df_std.to_excel(writer, sheet_name="Normalized", index=False)
            df_bc.to_excel(writer, sheet_name="Billable Cost", index=False)
            
        config = {
            "MONTH_LABEL": "May'26",
            "MONTH_END_DATE": "31052026",
            "COMPANY_CODE": 6000,
            "START_SERIAL_NO": 10
        }
        engine = JVEngine(config=config)
        rows = engine.run_processing(input_path, log_callback=lambda x: None)
        
        # We expect:
        # - Standard rows starting at Reference 10 (Credit + Debit + Spacer)
        # - Billable Cost rows starting at Reference 11 (Credit + Debit + Spacer)
        
        # Let's verify standard JV rows (index 0 to 2)
        assert rows[0]["Reference"] == 10
        assert rows[0]["Account"] == 500003
        assert rows[0]["Amount"] == "=-SUM(J3:J3)"
        assert rows[1]["Reference"] == 10
        assert rows[1]["Account"] == 742234 # Payroll
        assert rows[1]["Amount"] == -1000.0
        
        # Let's verify billable cost rows (index 3 to 5)
        assert rows[3]["Reference"] == 11
        assert rows[3]["Account"] == 500003
        assert rows[3]["Amount"] == "=-SUM(J6:J6)"
        assert rows[4]["Reference"] == 11
        assert rows[4]["Account"] == 742118
        assert rows[4]["Amount"] == -2000.0


def test_custom_sheet_selection():
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, "Custom_Input.xlsx")

        # Mock standard billing data
        std_data = {
            "Workday ID": ["W9999"],
            "Capability Center": ["Admin"],
            "Legal Entity": ["Entity9"],
            "Classification": ["Billable"],
            "Billed/ Unbilled": ["Billed"],
            "IC Code": ["IC9"],
            "Invoice No.": ["INV9"],
            "EmpNo (ref)": ["E9"],
            "Capability Center (ref)": ["Admin"],
            "Recharge - Payroll": [500.0],
            "Recharge - Manager": [0.0],
            "Recharge - Leadership": [0.0],
            "Recharge - Desk Cost": [0.0],
            "Recharge - Retirals": [0.0],
            "Mark up": [0.0],
        }
        df_std = pd.DataFrame(std_data)

        # Mock billable cost data
        bc_data = {
            "Ref Key 1": ["IC Code", "US_SAP_L"],
            "Assignment Number (20)": ["Invoice No.", "30001507"],
            "742118": ["Billable Cost", "-3000"],
            "Ref Key 3": ["Inv/ Bill No.", "290982"],
            "Ref Key 2": ["Billable Cost", "Billable Cost"]
        }
        df_bc = pd.DataFrame(bc_data)

        # Mock other data that should not be processed
        df_other = pd.DataFrame({"Trash": [1, 2, 3]})

        with pd.ExcelWriter(input_path, engine="openpyxl") as writer:
            df_std.to_excel(writer, sheet_name="My Billing Sheet", index=False)
            df_bc.to_excel(writer, sheet_name="My Billable Cost", index=False)
            df_other.to_excel(writer, sheet_name="Other Sheet", index=False)

        # 1. Normalize with explicit custom sheet names
        norm = BillingNormalizer(log_callback=lambda x: None)
        normalized_path, _ = norm.normalize(
            input_path,
            standard_sheet="My Billing Sheet",
            billable_cost_sheet="My Billable Cost"
        )

        assert os.path.exists(normalized_path)
        xls = pd.ExcelFile(normalized_path)
        assert "Normalized" in xls.sheet_names
        assert "Billable Cost" in xls.sheet_names

        # Run Engine on this
        config = {
            "MONTH_LABEL": "May'26",
            "MONTH_END_DATE": "31052026",
            "COMPANY_CODE": 6000,
            "START_SERIAL_NO": 1
        }
        engine = JVEngine(config=config)
        rows = engine.run_processing(
            normalized_path,
            standard_sheet="Normalized",
            billable_cost_sheet="Billable Cost",
            log_callback=lambda x: None
        )

        # Expect Standard JV (1 cr, 1 dr, 1 spacer) + Billable Cost JV (1 cr, 1 dr, 1 spacer) = 6 rows
        assert len(rows) == 6
        assert rows[0]["Reference"] == 1
        assert rows[0]["Account"] == 500003
        assert rows[0]["Amount"] == "=-SUM(J3:J3)"
        assert rows[3]["Reference"] == 2
        assert rows[3]["Account"] == 500003
        assert rows[3]["Amount"] == "=-SUM(J6:J6)"

        # 2. Normalize and skip Standard Billing sheet
        norm2 = BillingNormalizer(log_callback=lambda x: None)
        normalized_path2, _ = norm2.normalize(
            input_path,
            standard_sheet="[Skip / None]",
            billable_cost_sheet="My Billable Cost"
        )
        xls2 = pd.ExcelFile(normalized_path2)
        assert "Billable Cost" in xls2.sheet_names

        # Run Engine on this skipping standard
        rows2 = engine.run_processing(
            normalized_path2,
            standard_sheet="[Skip / None]",
            billable_cost_sheet="Billable Cost",
            log_callback=lambda x: None
        )
        # Only Billable Cost JV rows should be present (1 cr, 1 dr, 1 spacer) = 3 rows
        assert len(rows2) == 3
        assert rows2[0]["Reference"] == 1
        assert rows2[0]["Account"] == 500003
        assert rows2[0]["Amount"] == "=-SUM(J3:J3)"
        assert rows2[1]["Account"] == 742118
