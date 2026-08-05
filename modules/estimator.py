from modules.mapper import FieldMapper

class EstimateGenerator:

    ...

    def populate_from_record(self, record):

        mapper = FieldMapper()

        for _, row in mapper.mapping.iterrows():

            # Skip blank mapping rows
            if str(row["Data Source"]).strip() != "ODK":
                continue

            sheet = row["Workbook Sheet"]
            cell = row["Cell"]
            odk_field = row["ODK Field"]

            if odk_field in record:

                value = record[odk_field]

                self.write_value(sheet, cell, value)
