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

    from io import StringIO

    try:
        return pd.read_csv(StringIO(response.text))
    except Exception:
        raise Exception(
            "Unable to read CSV.\n\n"
            + response.text[:500]
        )
