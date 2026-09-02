import streamlit as st
import pandas as pd

from modules.odk import ODKCentral
from modules.estimator import EstimateGenerator
from modules.processor import RepairProcessor
import os

st.set_page_config(
    page_title="WHS Rejuvenation Estimation System",
    layout="wide"
)

st.title("WHS Rejuvenation Estimation")

# ---------------------------------------
# Load Basic Information
# ---------------------------------------
@st.cache_data(ttl=300)
def load_basic_data():
    odk = ODKCentral()
    return odk.get_basic_information()
    
basic = load_basic_data()
from modules.mapper import FieldMapper

mapper_test = FieldMapper()

st.write("### 🔎 Repeat Mapping Test")
st.dataframe(mapper_test.get_repeat_mapping())

st.success(f"{len(basic)} structures loaded")

# ---------------------------------------
# District
# ---------------------------------------

districts = sorted(
    basic["geo-district"].dropna().unique()
)

district = st.selectbox(
    "District",
    districts
)

# ---------------------------------------
# Block
# ---------------------------------------

block_df = basic[
    basic["geo-district"] == district
]

blocks = sorted(
    block_df["geo-block"].dropna().unique()
)

block = st.selectbox(
    "Block",
    blocks
)

# ---------------------------------------
# GP
# ---------------------------------------

gp_df = block_df[
    block_df["geo-block"] == block
]

gps = sorted(
    gp_df["geo-gp"].dropna().unique()
)

gp = st.selectbox(
    "Gram Panchayat",
    gps
)

# ---------------------------------------
# Village
# ---------------------------------------

village_df = gp_df[
    gp_df["geo-gp"] == gp
]

villages = sorted(
    village_df["geo-village"].dropna().unique()
)

village = st.selectbox(
    "Village",
    villages
)
st.divider()

st.subheader("📋 Structure Information")

structure = village_df[
    village_df["geo-village"] == village
].iloc[0]

col1, col2 = st.columns(2)

with col1:
    st.write(f"**District:** {district}")
    st.write(f"**Block:** {block}")
    st.write(f"**GP:** {gp}")
    st.write(f"**Village:** {village}")

with col2:
    st.write(f"**Latitude:** {structure['geo-village_gps-Latitude']}")
    st.write(f"**Longitude:** {structure['geo-village_gps-Longitude']}")
    st.write(f"**Altitude:** {structure['geo-village_gps-Altitude']}")

if st.button("📄 Generate Estimate", type="primary"):

    with st.spinner("Loading ODK data..."):

        odk = ODKCentral()

        repairs = odk.get_repairs()
        st.write("### 🔎 Repeat Fields Test")

        repeat_fields = [
            "ncg_-chainage_ncg_from",
            "ncg_-chainage_ncg_to",
            "guidewalls_side",
            "length_ncg",
        
            "Canal_guidewall_height_increase_-chainage_canal_guidewall_height_increase_from",
            "Canal_guidewall_height_increase_-chainage_canal_guidewall_height_increase_to",
            "length_canal_guidewall_height_increase",
        
            "gwr_-gwr_side",
            "gwr_-chainage_gwr_from",
            "gwr_-chainage_gwr_to",
            "avg_length_gwr",
        
            "stop_leak_bodywall_repeat_-chainage_leak_la_sl_from",
            "stop_leak_bodywall_repeat_-chainage_leak_la_sl_to",
            "stop_leak_bodywall_repeat_-avg_length_la_leak1_sl",
        
            "nboe_-canal_side",
            "nboe_-chainage_from_nboe",
            "nboe_-chainage_to_nboe",
            "nboe_-length_nboe",
        ]
        
        available_fields = [
            field for field in repeat_fields
            if field in repairs.columns
        ]
        
        st.dataframe(
            repairs[available_fields],
            use_container_width=True
        )
        lead = odk.get_lead()

        discharge = odk.get_discharge()

    processor = RepairProcessor(repairs)

    village_repairs = processor.filter_structure(
        district,
        block,
        gp,
        village
    )
    
    village_lead = processor.filter_lead(
        lead,
        district,
        block,
        gp,
        village
    )
    
    village_discharge = processor.filter_discharge(
        discharge,
        district,
        block,
        gp,
        village
    )

    st.success(f"Found {len(village_lead)} lead record(s)")
    st.success(f"Found {len(village_discharge)} discharge record(s)")
    st.success(f"Found {len(village_repairs)} repair record(s)")

    try:
        # Create estimate workbook
        estimator = EstimateGenerator()
        # -------------------------
        # Sheet-G
        # -------------------------
        record = structure.to_dict()

        estimator.populate_sheet(
            "Input Data Sheet-G",
            record
        )

        estimator.write_fixed_values(record)

        st.success("✅ Input Data Sheet-G populated")

        # -------------------------
        # Sheet-T
        # -------------------------
        # -------------------------
        # Repairs
        # -------------------------
        if len(village_repairs) > 0:

            repair_record = village_repairs.iloc[0].to_dict()

            estimator.populate_sheet(
                "Input Data Sheet-T",
                repair_record
            )

            st.success("✅ Repair data populated")

        else:

            st.warning("No repair records found.")

        # -------------------------
        # Lead Statement
        # -------------------------
        if len(village_lead) > 0:

            lead_record = village_lead.iloc[0].to_dict()

            estimator.populate_sheet(
                "Input Data Sheet-T",
                lead_record
            )

            st.success("✅ Lead Statement populated")

        else:

            st.warning("No Lead Statement found.")
        # -------------------------
        # Discharge
        # -------------------------
        if len(village_discharge) > 0:

            discharge_record = village_discharge.iloc[0].to_dict()

            estimator.populate_sheet(
                "Input Data Sheet-T",
                discharge_record
             )

            st.success("✅ Discharge data populated")

        else:

            st.warning("No Discharge record found.")

    except Exception as e:
        st.error(f"Populate Error: {e}")
        st.stop()

    output_file = os.path.join(
        "output",
        f"{village}_Estimate.xlsx"
    )

    estimator.save(output_file)

    st.success("✅ Estimate workbook created successfully!")

    import os

    if os.path.exists(output_file):
        st.success("✅ File created successfully!")

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download Estimate",
                data=f,
                file_name=f"{village}_Estimate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ File was NOT created.")
