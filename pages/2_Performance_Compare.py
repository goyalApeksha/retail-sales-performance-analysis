"""UI-03 Performance Compare."""

from datetime import date
import calendar

import pandas as pd
import streamlit as st

from src.analysis import compare_actual_to_target, grouped_aov
from src.charts import grouped_bar, render_chart, target_variance_chart
from src.data_loader import load_target_data
from src.page_utils import render_page_context, show_active_slice, show_empty

frame, dates, regions, categories = render_page_context("Performance Compare", "UI-03 · Actual-versus-target and average order value")
show_active_slice(dates, regions, categories, len(frame))
if not show_empty(frame):
    grain_label = st.selectbox("Compare target by", ["Store", "Product", "Region"], key="target_grain")
    targets = load_target_data()
    targets["target_month"] = pd.to_datetime(targets["target_month"])
    start_month = pd.Timestamp(dates[0]).to_period("M")
    end_month = pd.Timestamp(dates[1]).to_period("M")
    target_month_period = targets["target_month"].dt.to_period("M")
    targets = targets.loc[target_month_period.between(start_month, end_month)]
    if regions:
        targets = targets.loc[targets["region"].isin(regions)]
    if categories:
        targets = targets.loc[targets["category"].isin(categories)]
    if grain_label == "Store":
        grain = ("sales_month", "store_id", "product_id")
        label_col = "store_name"
    elif grain_label == "Product":
        grain = ("sales_month", "store_id", "product_id")
        label_col = "product_name"
    else:
        grain = ("sales_month", "region")
        label_col = "region"
    comparison = compare_actual_to_target(frame, targets, grain=grain)
    if grain_label == "Store":
        labels = frame[["store_id", "store_name"]].drop_duplicates()
        comparison = comparison.merge(labels, on="store_id", how="left")
        comparison = comparison.groupby(["store_id", "store_name"], as_index=False)[["actual", "target", "variance"]].sum()
        comparison["variance_pct"] = comparison["variance"].div(comparison["target"].replace(0, float("nan"))) * 100
    elif grain_label == "Product":
        labels = frame[["product_id", "product_name"]].drop_duplicates()
        comparison = comparison.merge(labels, on="product_id", how="left")
        comparison = comparison.groupby(["product_id", "product_name"], as_index=False)[["actual", "target", "variance"]].sum()
        comparison["variance_pct"] = comparison["variance"].div(comparison["target"].replace(0, float("nan"))) * 100
    else:
        comparison = comparison.groupby("region", as_index=False)[["actual", "target", "variance"]].sum()
        comparison["variance_pct"] = comparison["variance"].div(comparison["target"].replace(0, float("nan"))) * 100
    st.subheader("Actual versus target variance")
    st.caption("Targets are summed from the target fact table independently at matching month/store/product grain. Partial-month actuals are compared with the full monthly target.")
    if dates[0].day != 1 or dates[1].day != calendar.monthrange(dates[1].year, dates[1].month)[1]:
        st.warning("The selected range includes a partial month; actual sales are compared with that month's full target.")
    if not comparison.empty:
        render_chart(target_variance_chart(comparison, label_col))
        st.dataframe(comparison, use_container_width=True, hide_index=True)
    st.subheader("Average order value by region and segment")
    aov = grouped_aov(frame, ("region", "customer_segment"))
    render_chart(grouped_bar(aov, "region", "aov", "customer_segment", "AOV by region and segment", "AOV (INR / order)"))
    st.dataframe(aov, use_container_width=True, hide_index=True)
