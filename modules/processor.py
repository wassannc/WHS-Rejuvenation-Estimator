"""
processor.py

Converts ODK data into a Structure Object.
"""

import pandas as pd


class StructureProcessor:

    def __init__(self, basic_df, repair_df):
        self.basic_df = basic_df
        self.repair_df = repair_df

    def get_structure(self, village):

        # -------- Basic Information --------

        basic = self.basic_df[
            self.basic_df["village"] == village
        ]

        if basic.empty:
            raise Exception(f"{village} not found")

        basic = basic.iloc[0].to_dict()

        # -------- Repairs --------

        repairs = self.repair_df[
            self.repair_df["village"] == village
        ]

        structure = {

            "basic": basic,

            "repairs": repairs,

            "gis": {},

            "calculations": {},

            "estimate": {}

        }

        return structure
