"""
Dynamic Mapping Engine
Reads field_mapping.xlsx and provides mapping information.
"""

import pandas as pd
from config import MAPPING_FILE


class FieldMapper:

    def __init__(self):
        self.mapping = pd.read_excel(MAPPING_FILE, sheet_name="Sheet1")
        self.repeat_mapping = pd.read_excel(
            MAPPING_FILE,
            sheet_name="Repeat Mapping"
        )

        # remove completely blank rows
        self.mapping = self.mapping.dropna(how="all")

        # clean column names
        self.mapping.columns = (
            self.mapping.columns
            .astype(str)
            .str.strip()
        )

    def get_sheet_mapping(self, sheet_name):
        """
        Returns mapping rows for a specific sheet.
        """
        return self.mapping[
            self.mapping["Workbook Sheet"] == sheet_name
        ]

    def get_odk_fields(self):

        return self.mapping[
            self.mapping["Data Source"] == "ODK"
        ]

    def get_formula_fields(self):

        return self.mapping[
            self.mapping["Data Source"] == "Formula"
        ]

    def get_gis_fields(self):

        return self.mapping[
            self.mapping["Data Source"] == "GIS"
        ]
    def get_repeat_mapping(self):
        """
        Returns repeat mapping rows from the Repeat Mapping sheet.
        """
        return self.mapping
