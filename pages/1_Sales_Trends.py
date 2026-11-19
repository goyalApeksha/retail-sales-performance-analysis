"""UI-02 Sales Trends."""

import pandas as pd
import streamlit as st
from datetime import date
from dateutil.relativedelta import relativedelta

from src.analysis import monthly_revenue_growth
from src.charts import monthly_quantity_chart, monthly_revenue_chart, render_chart, seasonality_heatmap
from src.page_utils import render_page_context, show_active_slice, show_empty

frame, dates, regions, categories = render_page_context("Sales Trends", "UI-02 · Monthly revenue growth and seasonality")
show_active_slice(dates, regions, categories, len(frame))
if not show_empty(frame):
    data_start = pd.to_datetime(frame["sale_date"]).min().date()
    data_end = pd.to_datetime(frame["sale_date"]).max().date()
    default_a = (max(data_start, data_end - relativedelta(months=2)), data_end)
    default_b = (max(data_start, default_a[0] - relativedelta(years=1)), max(data_start, data_end - relativedelta(years=1)))
    if "compare_period_a" not in st.session_state:
        st.session_state["compare_period_a"] = default_a
    if "compare_period_b" not in st.session_state:
        st.session_state["compare_period_b"] = default_b
    st.subheader("Compare two periods")
    p1, p2 = st.columns(2)
    with p1:
        period_a = st.date_input("Period A", min_value=data_start, max_value=data_end, key="compare_period_a")
    with p2:
        period_b = st.date_input("Period B", min_value=data_start, max_value=data_end, key="compare_period_b")
    if len(period_a) == 2 and len(period_b) == 2:
        dates_col = pd.to_datetime(frame["sale_date"])
        revenue_a = frame.loc[dates_col.between(pd.Timestamp(period_a[0]), pd.Timestamp(period_a[1])), "net_revenue"].sum()
        revenue_b = frame.loc[dates_col.between(pd.Timestamp(period_b[0]), pd.Timestamp(period_b[1])), "net_revenue"].sum()
        growth = ((revenue_a - revenue_b) / revenue_b * 100) if revenue_b else None
        a, b, c = st.columns(3)
        a.metric(f"Period A · {period_a[0]} to {period_a[1]}", f"₹{revenue_a:,.2f}")
        b.metric(f"Period B · {period_b[0]} to {period_b[1]}", f"₹{revenue_b:,.2f}")
        c.metric("Change A vs B", "Unavailable" if growth is None else f"{growth:.2f}%")

    trend = monthly_revenue_growth(frame)
    window = st.radio("Trend window", ["Latest 12 months", "Full selected range"], horizontal=True, key="trend_window")
    display_trend = trend.tail(12).copy() if window == "Latest 12 months" else trend.copy()
    visible_months = set(display_trend["sales_month"].tolist())
    display_frame = frame.loc[frame["sales_month"].astype(str).isin(visible_months)] if "sales_month" in frame else frame.loc[pd.to_datetime(frame["sale_date"]).dt.to_period("M").astype(str).isin(visible_months)]
    st.subheader("Revenue and growth")
    render_chart(monthly_revenue_chart(display_frame))
    if not display_trend.empty:
        current = display_trend.iloc[-1]
        m1, m2 = st.columns(2)
        m1.metric("Latest month net revenue", f"₹{current['net_revenue']:,.2f}")
        mom = current["mom_growth_pct"]
        m2.metric("Latest MoM growth", "Unavailable" if pd.isna(mom) else f"{mom:.2f}%")
        st.dataframe(display_trend, use_container_width=True, hide_index=True)
    st.subheader("Sales volume")
    render_chart(monthly_quantity_chart(display_frame))
    render_chart(seasonality_heatmap(frame))
    st.caption("The latest 12 months means the last 12 months available in the selected synthetic data. Growth is unavailable when the prior-period denominator is zero or absent.")
