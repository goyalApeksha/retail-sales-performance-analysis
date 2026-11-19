"""UI-06 Data preview and filtered CSV download."""

import pandas as pd
import streamlit as st

from src.data_loader import serialize_filtered_csv
from src.page_utils import render_page_context, show_active_slice

frame, dates, regions, categories = render_page_context("Data and Download", "UI-06 · Filtered synthetic records; preview is limited to 100 rows")
show_active_slice(dates, regions, categories, len(frame))
st.info("This project uses synthetic data only. Downloaded rows follow the shared date, region and category filters.")
st.write(f"Filtered sales lines: **{len(frame):,}**")
st.dataframe(frame.head(100), use_container_width=True, hide_index=True)
if not frame.empty:
    csv_text = serialize_filtered_csv(frame)
    st.download_button(
        "Download filtered CSV", data=csv_text,
        file_name="retail_sales_filtered_synthetic.csv", mime="text/csv",
        use_container_width=True,
    )
with st.expander("Analysis view columns"):
    dictionary = pd.DataFrame({
        "Column": frame.columns,
        "Description": ["Joined synthetic sales analysis field"] * len(frame.columns),
    })
    st.dataframe(dictionary, use_container_width=True, hide_index=True)
