"""
Repair Data Processor
"""

import pandas as pd


class RepairProcessor:

    def __init__(self, repairs_df):
        self.repairs = repairs_df

    def filter_structure(
        self,
        district,
        block,
        gp,
        village
    ):

        df = self.repairs[
            (self.repairs["basic_details_repairs-district"] == district) &
            (self.repairs["basic_details_repairs-block"] == block) &
            (self.repairs["basic_details_repairs-gp"] == gp) &
            (self.repairs["basic_details_repairs-village"] == village)
        ]

        return df.reset_index(drop=True)
