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

odk = ODKCentral()
basic = odk.get_basic_information()

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

st.write("---")

st.write("### Selected Structure")

st.write(f"**District :** {district}")
st.write(f"**Block :** {block}")
st.write(f"**GP :** {gp}")
st.write(f"**Village :** {village}")
