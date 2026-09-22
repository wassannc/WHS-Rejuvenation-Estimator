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
            return 2

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

            # Length = Chainage To - Chainage From
            sheet.cell(output_row, 8).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )

            output_row += 1
        return output_row
    
    def populate_ncg_repeat(self, repeat_records, start_row=2):
        """
        Write NCG repeat records into Repeat Details
        and populate NCG totals/chainages in Input Data Sheet-T.
        """
    
        sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        ncg_filename = "2.Rejuvenation_works-ncg_.csv"
    
        if ncg_filename not in repeat_records:
            return start_row
    
        ncg_df = repeat_records[ncg_filename]
    
        output_row = start_row
    
        left_lengths = []
        right_lengths = []
    
        left_from = []
        left_to = []
    
        right_from = []
        right_to = []
    
        for record_no, (_, record) in enumerate(
            ncg_df.iterrows(), start=1
        ):
    
            side = str(record.get("guidewalls_side", "")).strip().lower()
    
            chainage_from = record.get("chainage_ncg_from", "")
            chainage_to = record.get("chainage_ncg_to", "")
    
            sheet.cell(output_row, 2).value = "NCG"
            sheet.cell(output_row, 3).value = "New Canal Guidewall"
            sheet.cell(output_row, 4).value = record_no
            sheet.cell(output_row, 5).value = side
            sheet.cell(output_row, 6).value = chainage_from
            sheet.cell(output_row, 7).value = chainage_to
    
            # Length = Chainage To - Chainage From
            sheet.cell(output_row, 8).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            # Collect values for Sheet-T
            try:
                cf = float(chainage_from)
                ct = float(chainage_to)
                length = ct - cf
    
                if side == "left":
                    left_lengths.append(length)
                    left_from.append(str(chainage_from))
                    left_to.append(str(chainage_to))
    
                elif side == "right":
                    right_lengths.append(length)
                    right_from.append(str(chainage_from))
                    right_to.append(str(chainage_to))
    
            except (ValueError, TypeError):
                pass
    
            output_row += 1
    
        # -------------------------------------------------
        # Populate Input Data Sheet-T
        # -------------------------------------------------
    
        # LEFT NCG -> Row 19
        if left_lengths:
            target_sheet["E19"] = sum(left_lengths)
            target_sheet["C19"] = "; ".join(left_from)
            target_sheet["D19"] = "; ".join(left_to)
        else:
            target_sheet["E19"] = 0
            target_sheet["C19"] = ""
            target_sheet["D19"] = ""
    
        # RIGHT NCG -> Row 20
        if right_lengths:
            target_sheet["E20"] = sum(right_lengths)
            target_sheet["C20"] = "; ".join(right_from)
            target_sheet["D20"] = "; ".join(right_to)
        else:
            target_sheet["E20"] = 0
            target_sheet["C20"] = ""
            target_sheet["D20"] = ""
    
        return output_row

    def populate_cghi_repeat(self, repeat_records, start_row=2):
        """
        Write Canal Guidewall Height Increase repeat records
        into Repeat Details.
        """
    
        sheet = self.workbook["Repeat Details"]
    
        cghi_filename = "2.Rejuvenation_works-Canal_guidewall_height_increase_.csv"
    
        if cghi_filename not in repeat_records:
            return start_row
    
        cghi_df = repeat_records[cghi_filename]
    
        output_row = start_row
    
        for record_no, (_, record) in enumerate(
            cghi_df.iterrows(), start=1
        ):
    
            sheet.cell(output_row, 2).value = "CGHI"
            sheet.cell(output_row, 3).value = "Increasing height of guide wall"
            sheet.cell(output_row, 4).value = record_no
            sheet.cell(output_row, 5).value = record.get(
                "canal_guidewall_height_increase_side",
                ""
            )

            sheet.cell(output_row, 6).value = record.get(
                "chainage_canal_guidewall_height_increase_from",
                ""
            )
    
            sheet.cell(output_row, 7).value = record.get(
                "chainage_canal_guidewall_height_increase_to",
                ""
            )
    
            # Length = Chainage To - Chainage From
            sheet.cell(output_row, 8).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            output_row += 1
    
        return output_row
        
    def populate_gwbjl_repeat(self, repeat_records, start_row=2):
        """
        Write GWBJL repeat records into Repeat Details
        and populate E50/E51 in Input Data Sheet-T.
    
        Leak 1 -> E50
        Leak 2 -> E51
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        gwbjl_filename = "2.Rejuvenation_works-gwbjl_.csv"
    
        if gwbjl_filename not in repeat_records:
            return start_row
    
        gwbjl_df = repeat_records[gwbjl_filename]
    
        output_row = start_row
    
        leak1_total = 0
        leak2_total = 0
    
        for record_no, (_, record) in enumerate(
            gwbjl_df.iterrows(), start=1
        ):
    
            # ---------------------------------------------
            # Common information
            # ---------------------------------------------
    
            repeat_sheet.cell(output_row, 2).value = "GWBJL"
            repeat_sheet.cell(
                output_row, 3
            ).value = "Hunch through bed and guidewall joint"
    
            repeat_sheet.cell(output_row, 4).value = record_no
            repeat_sheet.cell(output_row, 5).value = "right"
    
            # ---------------------------------------------
            # Chainage
            # ---------------------------------------------
    
            chainage_from = record.get(
                "leak1_gwbjl_chinage_from", ""
            )
    
            chainage_to = record.get(
                "leak1_gwbjl_chinage_to", ""
            )
    
            repeat_sheet.cell(output_row, 6).value = chainage_from
            repeat_sheet.cell(output_row, 7).value = chainage_to
    
            # ---------------------------------------------
            # Leak 1
            # ---------------------------------------------
    
            leak1 = record.get(
                "leakage_canal_length_gwbjl_leak1",
                ""
            )
    
            try:
                leak1_value = float(leak1) if leak1 not in ["", None] else 0
            except (ValueError, TypeError):
                leak1_value = 0
    
            leak1_total += leak1_value
    
            leak1_total += leak1_value
            repeat_sheet.cell(output_row, 8).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            output_row += 1
    
            # ---------------------------------------------
            # Leak 2
            # ---------------------------------------------
    
            leak2 = record.get(
                "leakage_canal_length_gwbjl_leak2",
                ""
            )
    
            try:
                leak2_value = float(leak2) if leak2 not in ["", None] else 0
            except (ValueError, TypeError):
                leak2_value = 0
    
            leak2_total += leak2_value
    
            repeat_sheet.cell(output_row, 2).value = "GWBJL"
            repeat_sheet.cell(
                output_row, 3
            ).value = "Hunch through bed and guidewall joint"
    
            repeat_sheet.cell(output_row, 4).value = record_no
            repeat_sheet.cell(output_row, 5).value = "left"
    
            repeat_sheet.cell(
                output_row, 6
            ).value = record.get(
                "leak1_gwbjl_chinage_from", ""
            )
    
            repeat_sheet.cell(
                output_row, 7
            ).value = record.get(
                "leak1_gwbjl_chinage_to", ""
            )
    
            leak2_total += leak2_value
            repeat_sheet.cell(output_row, 8).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            output_row += 1
    
        # ---------------------------------------------
        # ---------------------------------------------
        # Populate Sheet-T from Repeat Details
        # ---------------------------------------------
        
        # Right canal - Row 50
        target_sheet["C50"] = "='Repeat Details'!J500"
        target_sheet["D50"] = "='Repeat Details'!K500"
        
        target_sheet["E50"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"GWBJL",'
            '\'Repeat Details\'!$E$2:$E$500,"right")'
        )
                            
        # Left canal - Row 51
        target_sheet["C51"] = "='Repeat Details'!L500"
        target_sheet["D51"] = "='Repeat Details'!M500"
        
        target_sheet["E51"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"GWBJL",'
            '\'Repeat Details\'!$E$2:$E$500,"left")'
        )

    
        return output_row

    def populate_ltcb_repeat(self, repeat_records, start_row=2):
        """
        Write LTCB repeat records into Repeat Details.
    
        Leak 1 -> Right canal -> E52
        Leak 2 -> Left canal  -> E53
    
        Length is taken from ODK and written to Repeat Details H.
        Sheet-T totals are formula-based so manual additions
        in Repeat Details are also included.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        ltcb_filename = "2.Rejuvenation_works-ltcb_.csv"
    
        if ltcb_filename not in repeat_records:
            return start_row
    
        ltcb_df = repeat_records[ltcb_filename]
    
        output_row = start_row
    
        for record_no, (_, record) in enumerate(
            ltcb_df.iterrows(), start=1
        ):
    
            # -------------------------------------------------
            # RIGHT CANAL - Leak 1
            # -------------------------------------------------
    
            leak1 = record.get(
                "ltcb_-canal_damaged_length_leak1",
                ""
            )
    
            repeat_sheet.cell(output_row, 2).value = "LTCB"
            repeat_sheet.cell(
                output_row, 3
            ).value = "New bed over the existing bed"
            repeat_sheet.cell(output_row, 4).value = record_no
            repeat_sheet.cell(output_row, 5).value = "right"
    
            repeat_sheet.cell(
                output_row, 6
            ).value = record.get(
                "leak1_ltcb_chinage_from",
                ""
            )
    
            repeat_sheet.cell(
                output_row, 7
            ).value = record.get(
                "leak1_ltcb_chinage_to",
                ""
            )
    
            repeat_sheet.cell(
                output_row, 8
            ).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            output_row += 1
    
            # -------------------------------------------------
            # LEFT CANAL - Leak 2
            # -------------------------------------------------
    
            leak2 = record.get(
                "canal_damaged_length_leak2",
                ""
            )
    
            repeat_sheet.cell(output_row, 2).value = "LTCB"
            repeat_sheet.cell(
                output_row, 3
            ).value = "New bed over the existing bed"
            repeat_sheet.cell(output_row, 4).value = record_no
            repeat_sheet.cell(output_row, 5).value = "left"
    
            repeat_sheet.cell(
                output_row, 6
            ).value = record.get(
                "leak1_ltcb_chinage_from",
                ""
            )
    
            repeat_sheet.cell(
                output_row, 7
            ).value = record.get(
                "leak1_ltcb_chinage_to",
                ""
            )
    
            repeat_sheet.cell(
                output_row, 8
            ).value = (
                f'=IF(AND(F{output_row}<>"",G{output_row}<>""),'
                f'G{output_row}-F{output_row},"")'
            )
    
            output_row += 1
    
        # -------------------------------------------------
        # Sheet-T formulas
        # These include future manual Repeat Details entries
        # -------------------------------------------------
    
        target_sheet["E52"] = (
            '=SUMIFS('
            "'Repeat Details'!$H$2:$H$500,"
            "'Repeat Details'!$B$2:$B$500,\"LTCB\","
            "'Repeat Details'!$E$2:$E$500,\"right\""
            ')'
        )
    
        target_sheet["E53"] = (
            '=SUMIFS('
            "'Repeat Details'!$H$2:$H$500,"
            "'Repeat Details'!$B$2:$B$500,\"LTCB\","
            "'Repeat Details'!$E$2:$E$500,\"left\""
            ')'
        )
    
        # -------------------------------------------------
        # First chainage for each side
        # -------------------------------------------------
    
        target_sheet["C52"] = (
            '=IFERROR(INDEX(\'Repeat Details\'!$F$2:$F$500,'
            'MATCH(1,'
            '(\'Repeat Details\'!$B$2:$B$500="LTCB")*'
            '(LOWER(\'Repeat Details\'!$E$2:$E$500)="right"),'
            '0)),"")'
        )
    
        target_sheet["D52"] = (
            '=IFERROR(INDEX(\'Repeat Details\'!$G$2:$G$500,'
            'MATCH(1,'
            '(\'Repeat Details\'!$B$2:$B$500="LTCB")*'
            '(LOWER(\'Repeat Details\'!$E$2:$E$500)="right"),'
            '0)),"")'
        )
    
        target_sheet["C53"] = (
            '=IFERROR(INDEX(\'Repeat Details\'!$F$2:$F$500,'
            'MATCH(1,'
            '(\'Repeat Details\'!$B$2:$B$500="LTCB")*'
            '(LOWER(\'Repeat Details\'!$E$2:$E$500)="left"),'
            '0)),"")'
        )
    
        target_sheet["D53"] = (
            '=IFERROR(INDEX(\'Repeat Details\'!$G$2:$G$500,'
            'MATCH(1,'
            '(\'Repeat Details\'!$B$2:$B$500="LTCB")*'
            '(LOWER(\'Repeat Details\'!$E$2:$E$500)="left"),'
            '0)),"")'
        )
    
        return output_row
        
    def setup_ltcb_formulas(self):
        """
        Consolidate LTCB chainages and lengths into Input Data Sheet-T.
        Supports ODK records and manually added Repeat Details records.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        # -------------------------------------------------
        # Helper columns for LTCB
        # N = Right Chainage From
        # O = Right Chainage To
        # P = Left Chainage From
        # Q = Left Chainage To
        # -------------------------------------------------
    
        for row in range(2, 501):
    
            if row == 2:
    
                repeat_sheet.cell(row, 14).value = (
                    f'=IF(AND($B{row}="LTCB",LOWER($E{row})="right"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 15).value = (
                    f'=IF(AND($B{row}="LTCB",LOWER($E{row})="right"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 16).value = (
                    f'=IF(AND($B{row}="LTCB",LOWER($E{row})="left"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 17).value = (
                    f'=IF(AND($B{row}="LTCB",LOWER($E{row})="left"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
            else:
    
                # RIGHT - Chainage From
                repeat_sheet.cell(row, 14).value = (
                    f'=IF(AND($B{row}="LTCB",'
                    f'LOWER($E{row})="right"),'
                    f'IF(N{row-1}="",TEXT($F{row},"0.0"),'
                    f'IF($F{row}="",N{row-1},'
                    f'N{row-1}&", "&TEXT($F{row},"0.0"))),'
                    f'N{row-1})'
                )
    
                # RIGHT - Chainage To
                repeat_sheet.cell(row, 15).value = (
                    f'=IF(AND($B{row}="LTCB",'
                    f'LOWER($E{row})="right"),'
                    f'IF(O{row-1}="",TEXT($G{row},"0.0"),'
                    f'IF($G{row}="",O{row-1},'
                    f'O{row-1}&", "&TEXT($G{row},"0.0"))),'
                    f'O{row-1})'
                )
    
                # LEFT - Chainage From
                repeat_sheet.cell(row, 16).value = (
                    f'=IF(AND($B{row}="LTCB",'
                    f'LOWER($E{row})="left"),'
                    f'IF(P{row-1}="",TEXT($F{row},"0.0"),'
                    f'IF($F{row}="",P{row-1},'
                    f'P{row-1}&", "&TEXT($F{row},"0.0"))),'
                    f'P{row-1})'
                )
    
                # LEFT - Chainage To
                repeat_sheet.cell(row, 17).value = (
                    f'=IF(AND($B{row}="LTCB",'
                    f'LOWER($E{row})="left"),'
                    f'IF(Q{row-1}="",TEXT($G{row},"0.0"),'
                    f'IF($G{row}="",Q{row-1},'
                    f'Q{row-1}&", "&TEXT($G{row},"0.0"))),'
                    f'Q{row-1})'
                )
    
        # -------------------------------------------------
        # LTCB Right canal - Row 52
        # -------------------------------------------------
    
        target_sheet["C52"] = "='Repeat Details'!N500"
        target_sheet["D52"] = "='Repeat Details'!O500"
    
        target_sheet["E52"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"LTCB",'
            '\'Repeat Details\'!$E$2:$E$500,"right")'
        )
    
        # -------------------------------------------------
        # LTCB Left canal - Row 53
        # -------------------------------------------------
    
        target_sheet["C53"] = "='Repeat Details'!P500"
        target_sheet["D53"] = "='Repeat Details'!Q500"
    
        target_sheet["E53"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"LTCB",'
            '\'Repeat Details\'!$E$2:$E$500,"left")'
        )
    
        # Hide helper columns
        for column in ["N", "O", "P", "Q"]:
            repeat_sheet.column_dimensions[column].hidden = True
        
    def setup_gwbjl_formulas(self):
        """
        Consolidate GWBJL chainages and lengths
        into Input Data Sheet-T.
    
        Right side -> Row 50
        Left side  -> Row 51
    
        Uses dedicated helper columns V:Y
        so it does not interfere with GWR or CGHI.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        # -------------------------------------------------
        # Helper columns for GWBJL
        #
        # V = Right From
        # W = Right To
        # X = Left From
        # Y = Left To
        # -------------------------------------------------
    
        for row in range(2, 501):
    
            if row == 2:
    
                # RIGHT - From
                repeat_sheet.cell(row, 22).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="right"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                # RIGHT - To
                repeat_sheet.cell(row, 23).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="right"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
                # LEFT - From
                repeat_sheet.cell(row, 24).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="left"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                # LEFT - To
                repeat_sheet.cell(row, 25).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="left"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
            else:
    
                # RIGHT - From
                repeat_sheet.cell(row, 22).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="right"),'
                    f'IF(V{row-1}="",TEXT($F{row},"0.0"),'
                    f'V{row-1}&", "&TEXT($F{row},"0.0")),'
                    f'V{row-1})'
                )
    
                # RIGHT - To
                repeat_sheet.cell(row, 23).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="right"),'
                    f'IF(W{row-1}="",TEXT($G{row},"0.0"),'
                    f'W{row-1}&", "&TEXT($G{row},"0.0")),'
                    f'W{row-1})'
                )
    
                # LEFT - From
                repeat_sheet.cell(row, 24).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="left"),'
                    f'IF(X{row-1}="",TEXT($F{row},"0.0"),'
                    f'X{row-1}&", "&TEXT($F{row},"0.0")),'
                    f'X{row-1})'
                )
    
                # LEFT - To
                repeat_sheet.cell(row, 25).value = (
                    f'=IF(AND($B{row}="GWBJL",'
                    f'LOWER($E{row})="left"),'
                    f'IF(Y{row-1}="",TEXT($G{row},"0.0"),'
                    f'Y{row-1}&", "&TEXT($G{row},"0.0")),'
                    f'Y{row-1})'
                )
    
        # -------------------------------------------------
        # Sheet-T
        # -------------------------------------------------
    
        # Right -> Row 50
        target_sheet["C50"] = "='Repeat Details'!V500"
        target_sheet["D50"] = "='Repeat Details'!W500"
    
        # Left -> Row 51
        target_sheet["C51"] = "='Repeat Details'!X500"
        target_sheet["D51"] = "='Repeat Details'!Y500"
    
        # -------------------------------------------------
        # Keep existing E50/E51 length formulas
        # -------------------------------------------------
    
        target_sheet["E50"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"GWBJL",'
            '\'Repeat Details\'!$E$2:$E$500,"right")'
        )
    
        target_sheet["E51"] = (
            '=SUMIFS(\'Repeat Details\'!$H$2:$H$500,'
            '\'Repeat Details\'!$B$2:$B$500,"GWBJL",'
            '\'Repeat Details\'!$E$2:$E$500,"left")'
        )
    
        # -------------------------------------------------
        # Hide helper columns
        # -------------------------------------------------
    
        for column in ["V", "W", "X", "Y"]:
            repeat_sheet.column_dimensions[column].hidden = True
    
    def setup_gwr_formulas(self):
        """
        Populate GWR chainage and length directly into Input Data Sheet-T.
    
        GWR:
            Right side -> Row 40
            Left side  -> Row 41
    
        This version avoids Excel helper-column formulas because
        Excel recalculation was causing blank / incorrect values.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        right_from = []
        right_to = []
        right_length = 0.0
    
        left_from = []
        left_to = []
        left_length = 0.0
    
        # ---------------------------------------------------------
        # Read actual GWR records from Repeat Details
        # ---------------------------------------------------------
    
        for row in range(2, 501):
    
            repeat_group = repeat_sheet.cell(row, 2).value
            side = repeat_sheet.cell(row, 5).value
            chain_from = repeat_sheet.cell(row, 6).value
            chain_to = repeat_sheet.cell(row, 7).value
            length = repeat_sheet.cell(row, 8).value
    
            if not repeat_group:
                continue
    
            if str(repeat_group).strip().upper() != "GWR":
                continue
    
            side = str(side or "").strip().lower()
    
            # Convert numeric values safely
            try:
                chain_from = float(chain_from) if chain_from not in ("", None) else None
            except:
                chain_from = None
    
            try:
                chain_to = float(chain_to) if chain_to not in ("", None) else None
            except:
                chain_to = None
    
            try:
                length = float(length) if length not in ("", None) else 0.0
            except:
                length = 0.0
    
            # -----------------------------------------------------
            # RIGHT
            # -----------------------------------------------------
            if side == "right":
    
                if chain_from is not None:
                    right_from.append(chain_from)
    
                if chain_to is not None:
                    right_to.append(chain_to)
    
                right_length += length
    
            # -----------------------------------------------------
            # LEFT
            # -----------------------------------------------------
            elif side == "left":
    
                if chain_from is not None:
                    left_from.append(chain_from)
    
                if chain_to is not None:
                    left_to.append(chain_to)
    
                left_length += length
    
        # ---------------------------------------------------------
        # Helper to format chainages
        # ---------------------------------------------------------
    
        def format_chainages(values):
    
            if not values:
                return ""
    
            formatted = []
    
            for value in values:
                if float(value).is_integer():
                    formatted.append(str(int(value)))
                else:
                    formatted.append(f"{value:.1f}")
    
            return "; ".join(formatted)
    
        # ---------------------------------------------------------
        # RIGHT -> Input Data Sheet-T Row 40
        # ---------------------------------------------------------
    
        target_sheet["C40"] = format_chainages(right_from)
        target_sheet["D40"] = format_chainages(right_to)
        target_sheet["E40"] = right_length
    
        # ---------------------------------------------------------
        # LEFT -> Input Data Sheet-T Row 41
        # ---------------------------------------------------------
    
        target_sheet["C41"] = format_chainages(left_from)
        target_sheet["D41"] = format_chainages(left_to)
        target_sheet["E41"] = left_length
    
        # ---------------------------------------------------------
        # Clear old helper columns J:O
        # ---------------------------------------------------------
    
        for row in range(2, 501):
            for col in range(10, 16):
                repeat_sheet.cell(row, col).value = None
    
        # Hide helper columns
        for column in ["J", "K", "L", "M", "N", "O"]:
            repeat_sheet.column_dimensions[column].hidden = True
        
    def setup_cghi_formulas(self):
        """
        Consolidate CGHI records from Repeat Details
        separately for Right and Left sides.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        # -------------------------------------------------
        # Helper columns
        #
        # P = Right From
        # Q = Right To
        # R = Right Length
        # S = Left From
        # T = Left To
        # U = Left Length
        # -------------------------------------------------
    
        for row in range(2, 501):
    
            if row == 2:
    
                # RIGHT
                repeat_sheet.cell(row, 16).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 17).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 18).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'$H{row},0)'
                )
    
                # LEFT
                repeat_sheet.cell(row, 19).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'TEXT($F{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 20).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'TEXT($G{row},"0.0"),"")'
                )
    
                repeat_sheet.cell(row, 21).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'$H{row},0)'
                )
    
            else:
    
                # RIGHT - From
                repeat_sheet.cell(row, 16).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'IF(P{row-1}="",TEXT($F{row},"0.0"),'
                    f'P{row-1}&"; "&TEXT($F{row},"0.0")),'
                    f'P{row-1})'
                )
    
                # RIGHT - To
                repeat_sheet.cell(row, 17).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'IF(Q{row-1}="",TEXT($G{row},"0.0"),'
                    f'Q{row-1}&"; "&TEXT($G{row},"0.0")),'
                    f'Q{row-1})'
                )
    
                # RIGHT - Length
                repeat_sheet.cell(row, 18).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="right"),'
                    f'R{row-1}+$H{row},'
                    f'R{row-1})'
                )
    
                # LEFT - From
                repeat_sheet.cell(row, 19).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'IF(S{row-1}="",TEXT($F{row},"0.0"),'
                    f'S{row-1}&"; "&TEXT($F{row},"0.0")),'
                    f'S{row-1})'
                )
    
                # LEFT - To
                repeat_sheet.cell(row, 20).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'IF(T{row-1}="",TEXT($G{row},"0.0"),'
                    f'T{row-1}&"; "&TEXT($G{row},"0.0")),'
                    f'T{row-1})'
                )
    
                # LEFT - Length
                repeat_sheet.cell(row, 21).value = (
                    f'=IF(AND($B{row}="CGHI",'
                    f'LOWER($E{row})="left"),'
                    f'U{row-1}+$H{row},'
                    f'U{row-1})'
                )
    
        # -------------------------------------------------
        # Sheet-T
        # -------------------------------------------------
    
        # Right side → Row 42
        target_sheet["C42"] = "='Repeat Details'!P500"
        target_sheet["D42"] = "='Repeat Details'!Q500"
        target_sheet["E42"] = "='Repeat Details'!R500"
    
        # Left side → Row 43
        target_sheet["C43"] = "='Repeat Details'!S500"
        target_sheet["D43"] = "='Repeat Details'!T500"
        target_sheet["E43"] = "='Repeat Details'!U500"
    
        # -------------------------------------------------
        # Hide helper columns
        # -------------------------------------------------
    
        for column in ["P", "Q", "R", "S", "T", "U"]:
            repeat_sheet.column_dimensions[column].hidden = True
            
    def write_final_gwr_values(self):
        """
        Write consolidated GWR values directly into Input Data Sheet-T.
        Avoids Excel formula/cached-value issues.
        """
    
        repeat_sheet = self.workbook["Repeat Details"]
        target_sheet = self.workbook["Input Data Sheet-T"]
    
        # -------------------------------------------------
        # Read final helper values
        # -------------------------------------------------
        right_from = repeat_sheet["J500"].value
        right_to = repeat_sheet["K500"].value
        right_length = repeat_sheet["L500"].value
    
        left_from = repeat_sheet["M500"].value
        left_to = repeat_sheet["N500"].value
        left_length = repeat_sheet["O500"].value
    
        # -------------------------------------------------
        # Write RIGHT GWR → Row 40
        # -------------------------------------------------
        target_sheet["C40"] = right_from if right_from not in (None, "") else ""
        target_sheet["D40"] = right_to if right_to not in (None, "") else ""
        target_sheet["E40"] = right_length if right_length not in (None, "") else 0
    
        # -------------------------------------------------
        # Write LEFT GWR → Row 41
        # -------------------------------------------------
        target_sheet["C41"] = left_from if left_from not in (None, "") else ""
        target_sheet["D41"] = left_to if left_to not in (None, "") else ""
        target_sheet["E41"] = left_length if left_length not in (None, "") else 0
    
        # -------------------------------------------------
        # Qty = Length × Breadth × Depth
        # -------------------------------------------------
        target_sheet["H40"] = (
            f'=IF(OR(E40="",F40="",G40=""),0,E40*F40*G40)'
        )
    
        target_sheet["H41"] = (
            f'=IF(OR(E41="",F41="",G41=""),0,E41*F41*G41)'
        )
    def save(self, filename):

        os.makedirs("output", exist_ok=True)

        # Recalculate Excel formulas when the workbook is opened
        self.workbook.calculation.fullCalcOnLoad = True
        self.workbook.calculation.forceFullCalc = True
        self.workbook.calculation.calcMode = "auto"

        self.workbook.save(filename)
