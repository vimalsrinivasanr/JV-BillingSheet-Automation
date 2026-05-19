"""
Controller: MainController
Coordinates between the View and Model.
"""

from models.billing_sheet import BillingSheet
from models.jv_generator import JVGenerator

class MainController:
    def __init__(self, view):
        self.view = view
        self.billing_sheet = None
        self.jv_generator = None

    def load_billing_sheet(self, filepath):
        self.billing_sheet = BillingSheet(filepath)
        df = self.billing_sheet.load()
        self.view.update_view(df)

    def generate_jv(self):
        if self.billing_sheet:
            self.jv_generator = JVGenerator(self.billing_sheet.df)
            self.jv_generator.generate()
            # TODO: Save or display JV output
