import os
import pytest
from models.billing_sheet import BillingSheet
from models.jv_generator import JVGenerator
import pandas as pd

def test_full_pipeline(tmp_path):
    # Use a real or sample input file
    sample_file = os.path.join('input', 'Input Data.xlsx')
    sheet = BillingSheet(sample_file)
    df = sheet.load()
    sheet.apply_business_rules()  # Stub, expand as needed
    assert not df.empty

    # Generate JV
    gen = JVGenerator(df)
    gen.generate()  # Stub, expand as needed
    # Save to a temp output file
    out_path = tmp_path / 'test_jv_output.xlsx'
    gen.save(str(out_path))
    assert os.path.exists(out_path)
    # Optionally, check output file is a valid Excel file
    out_df = pd.read_excel(out_path)
    assert not out_df.empty
