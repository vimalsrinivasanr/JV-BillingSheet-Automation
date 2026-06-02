"""
Model: BillingSheet
Handles data normalization, validation, and business rules for billing sheets.
"""

import os
import pandas as pd
from scripts.normalizer import BillingNormalizer

class BillingSheet:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None

    def load(self):
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"File not found: {self.filepath}")
        
        # Instantiate and run normalizer
        normalizer = BillingNormalizer()
        out_path, _ = normalizer.normalize(self.filepath)
        
        # Load the "Normalized" sheet from the generated output path
        self.df = pd.read_excel(out_path, sheet_name="Normalized")
        return self.df

    def validate(self):
        # Default validator returns None
        return None

    def apply_business_rules(self):
        # Stub for extending rules if needed
        pass
