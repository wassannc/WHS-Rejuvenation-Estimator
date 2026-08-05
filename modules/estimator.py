"""
Estimator Module
Reads the estimate template,
writes values into Excel,
and saves a new estimate.
"""

import os
from openpyxl import load_workbook

from config import TEMPLATE_FILE


class EstimateGenerator:

    def __init__(self):
        self.workbook = load_workbook(TEMPLATE_FILE)

    def write_value(self, sheet_name, cell, value):
        sheet = self.workbook[sheet_name]
        sheet[cell] = value

    def save(self, filename):
        os.makedirs("output", exist_ok=True)
        self.workbook.save(filename)
