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
    def setup_gwr_formulas(self):
        """
        Create Excel 2019+ compatible formulas that consolidate
        GWR records from Repeat Details into Input Data Sheet-T.
        """

        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]

        # -------------------------------------------------
        # Helper columns
        # J = Right From
        # K = Right To
        # L = Right Length
        # M = Left From
        # N = Left To
        # O = Left Length
        # -------------------------------------------------

        for row in range(2, 501):

            if row == 2:
                # First data row

                repeat_sheet.cell(row, 10).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),$F{row},"")'
                )

                repeat_sheet.cell(row, 11).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),$G{row},"")'
                )

                repeat_sheet.cell(row, 12).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),$H{row},0)'
                )

                repeat_sheet.cell(row, 13).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),$F{row},"")'
                )

                repeat_sheet.cell(row, 14).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),$G{row},"")'
                )

                repeat_sheet.cell(row, 15).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),$H{row},0)'
                )

            else:
                # Subsequent rows: build cumulative chainage text

                repeat_sheet.cell(row, 10).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),'
                    f'IF(J{row-1}<>"",J{row-1}&"; ","")&$F{row},'
                    f'J{row-1})'
                )

                repeat_sheet.cell(row, 11).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),'
                    f'IF(K{row-1}<>"",K{row-1}&"; ","")&$G{row},'
                    f'K{row-1})'
                )

                repeat_sheet.cell(row, 12).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="right"),'
                    f'L{row-1}+$H{row},'
                    f'L{row-1})'
                )

                repeat_sheet.cell(row, 13).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),'
                    f'IF(M{row-1}<>"",M{row-1}&"; ","")&$F{row},'
                    f'M{row-1})'
                )

                repeat_sheet.cell(row, 14).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),'
                    f'IF(N{row-1}<>"",N{row-1}&"; ","")&$G{row},'
                    f'N{row-1})'
                )

                repeat_sheet.cell(row, 15).value = (
                    f'=IF(AND($B{row}="GWR",'
                    f'LOWER($E{row})="left"),'
                    f'O{row-1}+$H{row},'
                    f'O{row-1})'
                )

        # -------------------------------------------------
        # Sheet-T
        # GWR Right = row 40
        # GWR Left  = row 41
        # -------------------------------------------------

        target_sheet["C40"] = "='Repeat Details'!J500"
        target_sheet["D40"] = "='Repeat Details'!K500"
        target_sheet["E40"] = "='Repeat Details'!L500"

        target_sheet["C41"] = "='Repeat Details'!M500"
        target_sheet["D41"] = "='Repeat Details'!N500"
        target_sheet["E41"] = "='Repeat Details'!O500"

        # -------------------------------------------------
        # Hide helper columns
        # -------------------------------------------------

        for column in ["J", "K", "L", "M", "N", "O"]:
            repeat_sheet.column_dimensions[column].hidden = True
    def save(self, filename):

        os.makedirs("output", exist_ok=True)

        # Recalculate Excel formulas when the workbook is opened
        self.workbook.calculation.fullCalcOnLoad = True
        self.workbook.calculation.forceFullCalc = True
        self.workbook.calculation.calcMode = "auto"

        self.workbook.save(filename)
