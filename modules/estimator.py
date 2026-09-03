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

            if row["Cell"] == "C7":
                print("Writing C7:", value)

            self.write_value(
                row["Workbook Sheet"],
                row["Cell"],
                value
            )

    def write_fixed_values(self, record):
        # Kharif extent from ODK
        extent = record.get("whs-extent_kharif", "")

        # Format extent
        try:
            extent = round(float(extent), 2)
        except Exception:
            pass

        # Name of Work
        self.write_value(
            "Input Data Sheet-G",
            "C10",
            "Rejuvenation of Water Harvesting Structure"
        )

        # Expected Outcome
        outcome = (
            f"Assured irrigation to {extent} acres of land;\n"
            f"Additional income of ₹25,000 to ₹40,000 per acre;\n"
            f"Ecological development in the village through agro-ecological farming practices."
        )

        self.write_value(
            "Input Data Sheet-G",
            "C11",
            outcome
        )
        self.write_value(
            "Input Data Sheet-T",
            "C66",
            "Kothavalasa"
        )

        self.write_value(
            "Input Data Sheet-T",
            "C70",
            "Mamidipalli"
        )
    def populate_gwr_repeat(self, repeat_records):
        """
        Write GWR repeat records into the Repeat Details sheet.
        """

        sheet = self.workbook["Repeat Details"]

        # Start writing from row 2
        output_row = 2

        # GWR CSV filename from ODK ZIP export
        gwr_filename = "2.Rejuvenation_works-gwr_.csv"

        if gwr_filename not in repeat_records:
            return

        gwr_df = repeat_records[gwr_filename]

        for record_no, (_, record) in enumerate(
            gwr_df.iterrows(), start=1
        ):

            # Repeat Group
            sheet.cell(output_row, 2).value = "GWR"

            # Parameter / Work
            sheet.cell(output_row, 3).value = "Guide Wall Repair"

            # Record No
            sheet.cell(output_row, 4).value = record_no

            # Side
            sheet.cell(output_row, 5).value = record.get(
                "gwr_side", ""
            )

            # Chainage From
            sheet.cell(output_row, 6).value = record.get(
                "chainage_gwr_from", ""
            )

            # Chainage To
            sheet.cell(output_row, 7).value = record.get(
                "chainage_gwr_to", ""
            )

            # Length
            sheet.cell(output_row, 8).value = record.get(
                "avg_length_gwr", ""
            )

            output_row += 1
    def save(self, filename):

        os.makedirs("output", exist_ok=True)

        # Recalculate Excel formulas when the workbook is opened
        self.workbook.calculation.fullCalcOnLoad = True
        self.workbook.calculation.forceFullCalc = True
        self.workbook.calculation.calcMode = "auto"

        self.workbook.save(filename)
