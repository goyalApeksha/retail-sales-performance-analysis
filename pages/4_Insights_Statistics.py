"""UI-05 Insights and Statistics."""

import streamlit as st

from src.analysis import payment_summary, returns_by_category
from src.charts import discount_margin_scatter, grouped_bar, render_chart, return_rate_chart
from src.page_utils import render_page_context, show_active_slice, show_empty
from src.statistics import discount_margin_correlation, payment_aov_group_test

frame, dates, regions, categories = render_page_context("Insights and Statistics", "UI-05 · Descriptive evidence with sample and denominator context")
show_active_slice(dates, regions, categories, len(frame))
if not show_empty(frame):
    st.subheader("Payment method use and order value")
    payment = payment_summary(frame)
    left, right = st.columns(2)
    with left:
        render_chart(grouped_bar(payment, "payment_method", "order_count", None, "Distinct orders by payment method", "Orders"))
    with right:
        render_chart(grouped_bar(payment, "payment_method", "aov", None, "AOV by payment method", "AOV (INR / order)"))
    st.dataframe(payment, use_container_width=True, hide_index=True)
    group_test = payment_aov_group_test(frame)
    if group_test["status"] == "valid":
        st.caption(f"Optional Kruskal–Wallis comparison across {group_test['n_groups']} eligible payment groups: H={group_test['statistic']:.3f}, p={group_test['p_value']:.4f}. This is descriptive, not causal.")
    else:
        st.caption("Optional payment-group test is unavailable because fewer than two groups have enough observations.")

    st.subheader("Discount and margin")
    render_chart(discount_margin_scatter(frame))
    correlation = discount_margin_correlation(frame)
    if not correlation.empty:
        for _, row in correlation.iterrows():
            if row["status"] == "valid":
                st.write(f"{row['category']}: Pearson r={row['pearson_r']:.3f}, p={row['p_value']:.4f}, n={row['n']}.")
            else:
                st.info(f"{row['category']}: correlation withheld (n={row['n']}); {row['reason']}")
        st.caption("Correlation describes association in this synthetic sample and does not establish causation.")

    st.subheader("Returns by category")
    returns = returns_by_category(frame)
    render_chart(return_rate_chart(returns))
    st.dataframe(returns, use_container_width=True, hide_index=True)
