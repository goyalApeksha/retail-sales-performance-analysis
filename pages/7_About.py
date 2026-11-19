"""UI-08 About and metric definitions."""

import streamlit as st

st.set_page_config(page_title="About | Retail Sales Performance", page_icon="ℹ️", layout="wide")
st.title("About this project")
st.caption("UI-08 · ABES MCA Data Analytics")
st.info("Retail Sales Performance Analysis is an academic dashboard built with synthetic Indian retail data. It does not use real customer or personal data.")
st.write("**Prepared by:** Apeksha Goel · Roll 2500320140030 · MCA 2nd Year DA, ABES")
st.write("**Mentor:** Kalpana Dwivedi")
st.write("**Synthetic data period:** January 2023–December 2024 · all monetary values are INR.")
st.subheader("Metric definitions")
st.markdown("""
- **Net revenue:** quantity × sale-time unit price × (1 − discount rate), rounded at line level.
- **Average order value:** eligible net revenue ÷ distinct eligible order IDs.
- **Profit margin:** (aggregate net revenue − aggregate line cost) ÷ aggregate net revenue × 100.
- **Return rate:** returned eligible sales lines ÷ eligible sales lines × 100.
- **Target variance:** actual net revenue − target at matching month/store/product grain.
- **Sales velocity:** units sold ÷ calendar days in the selected date range.
""")
st.subheader("Limitations")
st.write("Patterns are illustrative and generated for learning. Correlation is descriptive and does not establish causation. Monthly targets repeated in the analysis view must be aggregated from the target fact table independently.")
st.caption("Project baseline version 1.0 · synthetic-only academic decision-support tool.")
