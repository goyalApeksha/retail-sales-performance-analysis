"""UI-04 Product and Category EDA."""

import math

import streamlit as st

from src.analysis import category_margin, pareto_analysis, product_revenue, sales_velocity
from src.charts import category_margin_chart, pareto_chart, product_rank_chart, render_chart
from src.page_utils import render_page_context, show_active_slice, show_empty

frame, dates, regions, categories = render_page_context("Product and Category EDA", "UI-04 · Ranking, profitability, Pareto contribution and sales velocity")
show_active_slice(dates, regions, categories, len(frame))
if not show_empty(frame):
    limit = st.slider("Products to show", min_value=5, max_value=20, value=10, step=5)
    st.subheader("Product revenue ranking")
    ranking = product_revenue(frame, top_n=limit)
    render_chart(product_rank_chart(ranking))
    st.dataframe(ranking, use_container_width=True, hide_index=True)
    st.subheader("Category profitability")
    margins = category_margin(frame)
    render_chart(category_margin_chart(margins))
    st.dataframe(margins, use_container_width=True, hide_index=True)
    st.subheader("Pareto contribution")
    pareto = pareto_analysis(frame)
    if pareto["net_revenue"].sum() > 0:
        top_count = max(1, math.ceil(len(pareto) * 0.20))
        top_share = pareto.head(top_count)["net_revenue"].sum() / pareto["net_revenue"].sum() * 100
        st.metric(f"Revenue share from top 20% of products ({top_count} of {len(pareto)})", f"{top_share:.1f}%")
        render_chart(pareto_chart(pareto.head(limit)))
        crossing = pareto.loc[pareto["cumulative_share_pct"] >= 80]
        if not crossing.empty:
            st.info(f"The first 80% crossing occurs after {int(crossing.index[0]) + 1} products ({crossing.iloc[0]['cumulative_share_pct']:.1f}% cumulative).")
        st.caption("Top-product count is rounded up to include a whole product. The chart also marks the smallest set reaching 80%.")
    else:
        st.info("Pareto share is unavailable because the selection has no positive net revenue.")
    st.subheader("Sales velocity")
    days = (dates[1] - dates[0]).days + 1
    velocity = sales_velocity(frame, days)
    st.caption(f"Velocity denominator: {days} calendar days in the selected date range.")
    st.dataframe(velocity, use_container_width=True, hide_index=True)
