"""
Model: BillingSheet
Handles data normalization, validation, and business rules for billing sheets.
"""

import pandas as pd

class BillingSheet:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None

    def load(self):
        # TODO: Implement loading and normalization logic
        self.df = pd.read_excel(self.filepath)
        return self.df

    def validate(self):
        # TODO: Implement validation logic
        pass

    def apply_business_rules(self):
        # TODO: Implement business rules
        pass
