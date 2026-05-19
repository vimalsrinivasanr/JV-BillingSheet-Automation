import pytest
from models.billing_sheet import BillingSheet
import os

def test_load_valid_file():
    # Use a sample file from input/ for testing
    sample_file = os.path.join('input', 'Input Data.xlsx')
    sheet = BillingSheet(sample_file)
    df = sheet.load()
    assert df is not None
    assert not df.empty

def test_load_missing_file():
    with pytest.raises(FileNotFoundError):
        sheet = BillingSheet('input/does_not_exist.xlsx')
        sheet.load()

def test_validate_stub():
    # This is a stub; expand with real validation logic
    sample_file = os.path.join('input', 'Input Data.xlsx')
    sheet = BillingSheet(sample_file)
    sheet.load()
    assert sheet.validate() is None
