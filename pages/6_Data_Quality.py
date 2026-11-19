"""UI-07 Data-quality report and validation checks."""

from pathlib import Path
import logging

import pandas as pd
import streamlit as st

from src.data_loader import read_quality_report
from src.page_utils import render_page_context, show_active_slice
from src.validation import quality_summary, validate_sales_frame

logger = logging.getLogger(__name__)

frame, dates, regions, categories = render_page_context("Data Quality", "UI-07 · Generator issues, validation checks, and cleaning disclosure")
show_active_slice(dates, regions, categories, len(frame))
root = Path(__file__).resolve().parents[1]
report_path = root / "data" / "validation_report.json"
try:
    report = read_quality_report(report_path)
except (OSError, ValueError) as exc:
    logger.exception("Failed to read the synthetic-data validation report: %s", exc)
    st.error("The synthetic-data validation report is unavailable or invalid.")
    st.caption("Regenerate or restore data/validation_report.json using the approved synthetic-data generator.")
    report = {}

if report:
    st.metric("Generator validation", report.get("validation_status", "Unknown"))
    counts = report.get("actual_database_row_counts", {})
    st.subheader("SQLite table counts")
    st.dataframe(pd.DataFrame([{"Table": name, "Rows": count} for name, count in counts.items()]), use_container_width=True, hide_index=True)
    raw = report.get("raw_quality_issues", {})
    st.subheader("Intentional raw-input quality issues")
    a, b, c = st.columns(3)
    a.metric("Missing payment values", raw.get("missing_payment_method_cells", 0))
    b.metric("Duplicate business groups", raw.get("duplicate_group_count", 0))
    outliers = raw.get("valid_range_statistical_outliers", [])
    c.metric("Illustrative price outliers", len(outliers))
    st.caption(f"The constrained SQLite database imputes missing payment method to mode: {raw.get('imputation_for_database', 'not reported')}. Raw CSV/XLSX inputs remain the cleaning-practice source.")
    with st.expander("Declared validation checks"):
        checks = report.get("range_and_key_checks", {})
        st.dataframe(pd.DataFrame([{"Check": key, "Result": "PASS" if value else "REVIEW"} for key, value in checks.items()]), use_container_width=True, hide_index=True)
    local_issues = validate_sales_frame(frame)
    summary = quality_summary(local_issues)
    if summary.empty:
        st.success("Loaded analysis view passes required sales-frame checks.")
    else:
        st.warning("Review validation findings in the selected slice.")
        st.dataframe(summary, use_container_width=True, hide_index=True)
    with st.expander("Illustrative outlier details"):
        st.dataframe(pd.DataFrame(outliers), use_container_width=True, hide_index=True)
