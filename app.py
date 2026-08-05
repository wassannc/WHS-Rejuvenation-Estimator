import streamlit as st
import pandas as pd

from modules.odk import ODKCentral

st.set_page_config(
    page_title="WHS Rejuvenation Estimation System",
    layout="wide"
)

st.title("🏗 WHS Rejuvenation Estimation System")

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

st.info("Repairs: Not Loaded")
st.divider()

if st.button("📄 Generate Estimate", type="primary"):

    with st.spinner("Loading repair assessments..."):
        odk = ODKCentral()
        repairs = odk.get_repairs()

    st.success(f"{len(repairs)} repair records downloaded")

    st.subheader("Repairs Form Columns")
    st.write(repairs.columns.tolist())

    st.dataframe(repairs.head())

    st.stop()

st.write("---")
