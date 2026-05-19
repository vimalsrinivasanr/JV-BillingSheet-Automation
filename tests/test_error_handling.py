import pytest
from models.billing_sheet import BillingSheet
from models.jv_generator import JVGenerator

def test_billing_sheet_invalid_file():
    with pytest.raises(Exception):
        BillingSheet('input/invalid_file.txt').load()

def test_jv_generator_with_empty_df():
    import pandas as pd
    df = pd.DataFrame()
    gen = JVGenerator(df)
    # Should not raise, but generate() should handle empty gracefully
    assert gen.generate() is None
