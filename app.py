"""UI-01 Overview page and multipage Streamlit entrypoint."""

import streamlit as st

from src.analysis import calculate_kpis, product_revenue
from src.charts import monthly_revenue_chart, product_rank_chart, render_chart
from src.page_utils import render_page_context, show_active_slice, show_empty
from src.ui import render_kpi_card

st.set_page_config(page_title="Retail Sales Performance", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
frame, dates, regions, categories = render_page_context("Overview and KPIs", "UI-01 · A filtered health check for retail sales performance")
show_active_slice(dates, regions, categories, len(frame))

if not show_empty(frame):
    kpis = calculate_kpis(frame)
    cols = st.columns(4)
    labels = [
        ("Net revenue", kpis["net_revenue"], "INR"),
        ("Average order value", kpis["aov"], "INR / order"),
        ("Profit margin", kpis["profit_margin_pct"], "%"),
        ("Return rate", kpis["return_rate_pct"], "% of eligible sales lines"),
    ]
    for col, (label, value, unit) in zip(cols, labels):
        text = "Unavailable" if value is None else (f"₹{value:,.2f}" if unit.startswith("INR") else f"{value:.2f}%")
        with col:
            render_kpi_card(label, text, help_text=unit)
    left, right = st.columns([1.6, 1])
    with left:
        render_chart(monthly_revenue_chart(frame))
    with right:
        st.subheader("Leading products")
        ranking = product_revenue(frame, top_n=5)
        render_chart(product_rank_chart(ranking, "Top products in selected slice"))
    st.caption("Margin uses aggregate profit divided by aggregate net revenue. Return rate is returned eligible sales lines divided by eligible sales lines.")
