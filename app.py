import streamlit as st

st.set_page_config(
    page_title="WHS Rejuvenation Estimator",
    layout="wide"
)

st.title("🏗 WHS Rejuvenation Estimation System")

st.markdown("---")

st.header("Generate Estimate")

village = st.text_input("Enter Village Name")

if st.button("Generate Estimate"):

    st.info(f"Generating estimate for {village}...")

    # Next step:
    # 1. Read ODK
    # 2. Process
    # 3. Generate Workbook

    st.success("Coming in next step...")
