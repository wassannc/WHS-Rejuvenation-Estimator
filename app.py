import streamlit as st
import pandas as pd

from modules.odk import ODKCentral

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

    from modules.estimator import EstimateGenerator
    import os

    with st.spinner("Loading repair assessments..."):

        odk = ODKCentral()
        repairs = odk.get_repairs()

    # Filter repair records for selected structure
    village_repairs = repairs[
        (repairs["basic_details_repairs-district"] == district) &
        (repairs["basic_details_repairs-block"] == block) &
        (repairs["basic_details_repairs-gp"] == gp) &
        (repairs["basic_details_repairs-village"] == village)
    ]

    st.success(f"Found {len(village_repairs)} repair record(s)")

    # Create estimate workbook
    estimator = EstimateGenerator()
    # Convert selected structure into a dictionary
    record = structure.to_dict()

    try:
        estimator.populate_sheet(
            "Input Data Sheet-G",
            record
        )
        st.success("✅ Input Data Sheet-G populated")

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

    st.write(f"Saved to: {output_file}")

    if os.path.exists(output_file):
        st.success("✅ File exists!")
    else:
        st.error("❌ File was NOT created.")
