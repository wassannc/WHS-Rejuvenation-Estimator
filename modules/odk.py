"""
ODK Central Data Module
"""

import requests
import pandas as pd
from io import StringIO, BytesIO
from zipfile import ZipFile
from requests.auth import HTTPBasicAuth

from config import (
    ODK_URL,
    USERNAME,
    PASSWORD,
    PROJECT_ID,
    BASIC_FORM_ID,
    REPAIR_FORM_ID,
    LEAD_FORM_ID,
    DISCHARGE_FORM_ID,
)


class ODKCentral:

    def __init__(self):
        self.base_url = ODK_URL
        self.project_id = PROJECT_ID
        self.auth = HTTPBasicAuth(USERNAME, PASSWORD)

    def get_form_data(self, form_id):

        url = (
            f"{self.base_url}/v1/projects/"
            f"{self.project_id}/forms/{form_id}/submissions.csv"
        )

        response = requests.get(url, auth=self.auth)

        if response.status_code != 200:
            raise Exception(
                f"ODK Error {response.status_code}\n"
                f"{response.text}"
            )

        try:
            return pd.read_csv(StringIO(response.text))

        except Exception:
            raise Exception(
                "Unable to read CSV\n\n"
                + response.text[:500]
            )

    def get_form_export_files(self, form_id):
        """
        Download the complete ODK CSV ZIP export.

        Diagnostic version:
        - Shows each CSV filename and row count.
        - For the root CSV, shows key-related fields.
        - For repeat CSVs, shows PARENT_KEY and KEY values.

        This is used to confirm how repeat records connect
        to the main repair submission.
        """

        url = (
            f"{self.base_url}/v1/projects/"
            f"{self.project_id}/forms/{form_id}/submissions.csv.zip"
        )

        response = requests.get(url, auth=self.auth)

        if response.status_code != 200:
            raise Exception(
                f"ODK ZIP Export Error {response.status_code}\n"
                f"{response.text}"
            )

        export_details = []

        with ZipFile(BytesIO(response.content)) as z:

            for filename in z.namelist():

                if not filename.lower().endswith(".csv"):
                    continue

                with z.open(filename) as file:
                    df = pd.read_csv(file)

                details = {
                    "filename": filename,
                    "rows": len(df),
                }

                # Show only columns useful for identifying
                # the connection between root and repeat records.
                key_columns = [
                    col for col in df.columns
                    if (
                        "KEY" in str(col).upper()
                        or "INSTANCE" in str(col).upper()
                    )
                ]

                details["key_columns"] = key_columns

                if key_columns and len(df) > 0:
                    details["key_samples"] = (
                        df[key_columns]
                        .head(5)
                        .fillna("")
                        .astype(str)
                        .to_dict("records")
                    )
                else:
                    details["key_samples"] = []

                export_details.append(details)

        return export_details

    def get_basic_information(self):
        return self.get_form_data(BASIC_FORM_ID)

    def get_repairs(self):
        return self.get_form_data(REPAIR_FORM_ID)

    def get_lead(self):
        return self.get_form_data(LEAD_FORM_ID)

    def get_discharge(self):
        return self.get_form_data(DISCHARGE_FORM_ID)
