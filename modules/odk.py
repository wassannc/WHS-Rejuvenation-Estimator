# ODK repeat retrieval helper
# This version joins repeat CSV records to the selected parent repair
# using root KEY == repeat PARENT_KEY.

import pandas as pd
from io import BytesIO
from zipfile import ZipFile


def get_repeat_records_from_export(zip_bytes, parent_key):
    """
    Read repeat records from an ODK Central ZIP export.

    Parameters
    ----------
    zip_bytes : bytes
        ZIP file returned by ODK Central.
    parent_key : str
        Root repair KEY / meta-instanceID.

    Returns
    -------
    dict
        {csv_filename: DataFrame} containing only rows belonging
        to the selected repair.
    """
    results = {}

    with ZipFile(BytesIO(zip_bytes)) as z:
        for filename in z.namelist():
            if not filename.lower().endswith(".csv"):
                continue

            with z.open(filename) as f:
                df = pd.read_csv(f, dtype=str)

            if "PARENT_KEY" not in df.columns:
                continue

            matched = df[df["PARENT_KEY"].astype(str).str.strip() ==
                         str(parent_key).strip()].copy()

            results[filename] = matched.reset_index(drop=True)

    return results
