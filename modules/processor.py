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

    def filter_lead(
        self,
        lead_df,
        district,
        block,
        gp,
        village
    ):

        df = lead_df[
            (lead_df["basic_details_lead-district"] == district) &
            (lead_df["basic_details_lead-block"] == block) &
            (lead_df["basic_details_lead-gp"] == gp) &
            (lead_df["basic_details_lead-village"] == village)
        ]

        return df.reset_index(drop=True)
