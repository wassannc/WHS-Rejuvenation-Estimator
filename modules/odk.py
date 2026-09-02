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

        For each CSV file in the ZIP, return:
        - filename
        - number of rows
        - column names

        This is a diagnostic step to identify the repeat-table
        structure and the fields needed to connect repeat records
        with the main repair submission.
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

                # Ignore media files and directories.
                if not filename.lower().endswith(".csv"):
                    continue

                with z.open(filename) as file:
                    df = pd.read_csv(file)

                export_details.append(
                    {
                        "filename": filename,
                        "rows": len(df),
                        "columns": list(df.columns),
                    }
                )

        return export_details

    def get_basic_information(self):
        return self.get_form_data(BASIC_FORM_ID)

    def get_repairs(self):
        return self.get_form_data(REPAIR_FORM_ID)

    def get_lead(self):
        return self.get_form_data(LEAD_FORM_ID)

    def get_discharge(self):
        return self.get_form_data(DISCHARGE_FORM_ID)
