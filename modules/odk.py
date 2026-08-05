"""
ODK Central Data Module
-----------------------
This module is responsible for:
1. Connecting to ODK Central
2. Downloading submissions
3. Returning clean pandas DataFrames
"""

import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
from config import (
    ODK_URL,
    USERNAME,
    PASSWORD,
    PROJECT_ID,
    BASIC_FORM_ID,
    REPAIR_FORM_ID,
)


class ODKCentral:

    def __init__(self):
        self.base_url = ODK_URL
        self.project_id = PROJECT_ID
        self.auth = HTTPBasicAuth(USERNAME, PASSWORD)

    def get_form_data(self, form_id):
        """
        Download one form as CSV from ODK Central
        """

        url = (
            f"{self.base_url}/v1/projects/"
            f"{self.project_id}/forms/{form_id}.svc/Submissions.csv"
        )

        response = requests.get(url, auth=self.auth)

        if response.status_code != 200:
            raise Exception(
                f"Unable to download {form_id}\n"
                f"Status Code: {response.status_code}"
            )

        from io import StringIO

        return pd.read_csv(StringIO(response.text))

    def get_basic_information(self):
        return self.get_form_data(BASIC_FORM_ID)

    def get_repairs(self):
        return self.get_form_data(REPAIR_FORM_ID)
