"""Consistent Plotly chart builders for the app's analytical screens."""

from __future__ import annotations

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:  # Core data views remain usable if local dependency setup is incomplete.
    px = None
    go = None

PALETTE = ["#1F4E79", "#0F766E", "#D97706", "#6B7280", "#7C3AED"]


def _style(fig, title: str):
    fig.update_layout(title=title, colorway=PALETTE, template="plotly_white", margin=dict(l=18, r=18, t=55, b=20), legend_title_text="")
    fig.update_xaxes(showgrid=True, gridcolor="#E8EDF3")
    fig.update_yaxes(showgrid=False)
    return fig


def _fallback(frame: pd.DataFrame, title: str, kind: str, x: str, y: str, color: str | None = None):
    return {"fallback": True, "frame": frame.copy(), "title": title, "kind": kind, "x": x, "y": y, "color": color}


def render_chart(chart) -> None:
    """Render a Plotly figure, or readable tabular/native fallback if Plotly is absent."""
    if not isinstance(chart, dict) or not chart.get("fallback"):
        st.plotly_chart(chart, use_container_width=True)
        return
    st.caption(f"{chart['title']} · native fallback (Plotly is not installed)")
    frame = chart["frame"]
    kind, x, y, color = chart["kind"], chart["x"], chart["y"], chart["color"]
    if kind == "line" and x in frame and y in frame:
        st.line_chart(frame.set_index(x)[y])
    elif kind == "bar" and x in frame and y in frame:
        st.bar_chart(frame.set_index(x)[y])
    elif kind == "scatter" and x in frame and y in frame:
        st.scatter_chart(frame, x=x, y=y, color=color if color in frame else None)
    else:
        st.dataframe(frame, use_container_width=True, hide_index=True)


def monthly_revenue_chart(frame: pd.DataFrame, title: str = "Monthly net revenue"):
    grouped = frame.groupby("sales_month", as_index=False)["net_revenue"].sum().sort_values("sales_month")
    if px is None:
        return _fallback(grouped, title, "line", "sales_month", "net_revenue")
    fig = px.line(grouped, x="sales_month", y="net_revenue", markers=True, labels={"sales_month": "Month", "net_revenue": "Net revenue (INR)"})
    return _style(fig, title)


def monthly_quantity_chart(frame: pd.DataFrame, title: str = "Monthly units sold"):
    work = frame.copy()
    if "sales_month" not in work:
        work["sales_month"] = pd.to_datetime(work["sale_date"]).dt.to_period("M").astype(str)
    grouped = work.groupby("sales_month", as_index=False)["quantity"].sum().sort_values("sales_month")
    if px is None:
        return _fallback(grouped, title, "line", "sales_month", "quantity")
    fig = px.line(grouped, x="sales_month", y="quantity", markers=True, labels={"sales_month": "Month", "quantity": "Units"})
    return _style(fig, title)


def product_rank_chart(frame: pd.DataFrame, title: str = "Products ranked by net revenue"):
    shown = frame.sort_values("net_revenue", ascending=True)
    if px is None:
        return _fallback(shown, title, "table", "product_name", "net_revenue")
    fig = px.bar(shown, x="net_revenue", y="product_name", color="category", orientation="h", labels={"net_revenue": "Net revenue (INR)", "product_name": "Product"})
    return _style(fig, title)


def category_margin_chart(frame: pd.DataFrame):
    if px is None:
        return _fallback(frame, "Profit margin by category", "bar", "category", "margin_pct")
    fig = px.bar(frame.sort_values("margin_pct"), x="margin_pct", y="category", orientation="h", color="margin_pct", color_continuous_scale=["#D97706", "#0F766E"], labels={"margin_pct": "Aggregate margin (%)", "category": "Category"})
    return _style(fig, "Profit margin by category")


def pareto_chart(frame: pd.DataFrame):
    if go is None:
        return _fallback(frame, "Pareto contribution by product", "table", "product_name", "net_revenue")
    fig = go.Figure()
    fig.add_bar(x=frame["product_name"], y=frame["net_revenue"], name="Net revenue", marker_color="#1F4E79")
    fig.add_scatter(x=frame["product_name"], y=frame["cumulative_share_pct"], name="Cumulative share", yaxis="y2", mode="lines+markers", line=dict(color="#D97706", width=3))
    fig.add_hline(y=80, yref="y2", line_dash="dash", line_color="#6B7280", annotation_text="80%")
    fig.update_layout(yaxis=dict(title="Net revenue (INR)"), yaxis2=dict(title="Cumulative revenue share (%)", overlaying="y", side="right", range=[0, 105]), xaxis_title="Product", barmode="group")
    return _style(fig, "Pareto contribution by product")


def target_variance_chart(frame: pd.DataFrame, label: str = "store_name"):
    column = label if label in frame else "product_id"
    if px is None:
        return _fallback(frame, "Actual-versus-target variance", "bar", column, "variance")
    fig = px.bar(frame.sort_values("variance"), x="variance", y=column, orientation="h", color="variance", color_continuous_scale=["#B91C1C", "#D1D5DB", "#0F766E"], color_continuous_midpoint=0, labels={"variance": "Actual minus target (INR)"})
    fig.add_vline(x=0, line_color="#172B4D", line_width=1)
    return _style(fig, "Actual-versus-target variance")


def grouped_bar(frame: pd.DataFrame, x: str, y: str, color: str | None, title: str, y_label: str):
    if px is None:
        return _fallback(frame, title, "bar", x, y, color)
    fig = px.bar(frame, x=x, y=y, color=color, barmode="group", labels={x: x.replace("_", " ").title(), y: y_label})
    return _style(fig, title)


def discount_margin_scatter(frame: pd.DataFrame):
    if px is None:
        return _fallback(frame, "Discount and line margin association", "scatter", "discount_pct", "margin_pct", "category")
    fig = px.scatter(frame, x="discount_pct", y="margin_pct", color="category", hover_data=["sales_line_id", "net_revenue"], labels={"discount_pct": "Discount rate", "margin_pct": "Line margin (%)"})
    return _style(fig, "Discount and line margin association")


def return_rate_chart(frame: pd.DataFrame):
    if px is None:
        return _fallback(frame, "Returns by category", "bar", "category", "return_rate_pct")
    fig = px.bar(frame.sort_values("return_rate_pct"), x="return_rate_pct", y="category", orientation="h", hover_data=["returned_count", "eligible_count"], labels={"return_rate_pct": "Return rate (%)", "category": "Category"})
    return _style(fig, "Returns by category")


def seasonality_heatmap(frame: pd.DataFrame):
    work = frame.copy()
    work["year"] = pd.to_datetime(work["sale_date"]).dt.year
    work["calendar_month"] = pd.to_datetime(work["sale_date"]).dt.month
    matrix = work.pivot_table(index="year", columns="calendar_month", values="quantity", aggfunc="sum", fill_value=0)
    if px is None:
        display = matrix.reset_index()
        display.columns = [str(column) for column in display.columns]
        return _fallback(display, "Monthly volume by calendar year", "table", "year", "calendar_month")
    fig = px.imshow(matrix, aspect="auto", color_continuous_scale="Teal", labels={"x": "Calendar month", "y": "Year", "color": "Units"})
    return _style(fig, "Monthly volume by calendar year")
