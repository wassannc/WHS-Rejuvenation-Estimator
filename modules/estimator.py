"""
Estimation Engine
Version 1

Loads the Excel template and writes mapped values into it.
"""

from openpyxl import load_workbook
from config import TEMPLATE_FILE


class EstimateGenerator:

    def __init__(self):
        self.workbook = load_workbook(TEMPLATE_FILE)

    def get_sheet(self, sheet_name):
        return self.workbook[sheet_name]

    def write_value(self, sheet_name, cell, value):
        sheet = self.get_sheet(sheet_name)
        sheet[cell] = value

    def save(self, filename):
        self.workbook.save(filename)
