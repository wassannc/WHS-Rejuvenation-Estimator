"""
Estimator Module
"""

import os
from openpyxl import load_workbook

from config import TEMPLATE_FILE
from modules.mapper import FieldMapper


class EstimateGenerator:

    def __init__(self):
        self.workbook = load_workbook(TEMPLATE_FILE)
        self.mapper = FieldMapper()

    def write_value(self, sheet_name, cell, value):
        sheet = self.workbook[sheet_name]
        sheet[cell] = value

    def populate_sheet(self, sheet_name, record):

        mapping = self.mapper.get_sheet_mapping(sheet_name)

        for _, row in mapping.iterrows():

            if str(row["Data Source"]).strip() != "ODK":
                continue

            odk_field = str(row["ODK Field"]).strip()

            if odk_field not in record:
                continue

            value = record[odk_field]

            self.write_value(
                row["Workbook Sheet"],
                row["Cell"],
                value
            )
        # Fixed values
        self.write_value(
            "Input Data Sheet-G",
            "C10",
            "Rejuvenation of Water Harvesting Structure"
        )

    def save(self, filename):
        os.makedirs("output", exist_ok=True)
        self.workbook.save(filename)
