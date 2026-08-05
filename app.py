import streamlit as st
import pandas as pd

from modules.odk import ODKCentral

st.set_page_config(
    page_title="WHS Rejuvenation Estimation System",
    layout="wide"
)

st.title("🏗 WHS Rejuvenation Estimation System")

if st.button("Connect to ODK"):

    odk = ODKCentral()

    with st.spinner("Downloading Basic Information..."):

        basic = odk.get_basic_information()

    st.success(f"{len(basic)} records downloaded")

    st.dataframe(basic.head())
