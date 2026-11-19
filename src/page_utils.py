"""Shared app-page setup, filter state, and safe rendering helpers."""

from __future__ import annotations

from datetime import date
import logging
import sqlite3

import pandas as pd
import streamlit as st

from src.analysis import apply_filters
from src.data_loader import load_analysis_data_fresh
from src.ui import render_header

logger = logging.getLogger(__name__)


def load_page_data() -> pd.DataFrame:
    try:
        return load_analysis_data_fresh()
    except (OSError, ValueError, RuntimeError, sqlite3.Error) as exc:
        logger.exception("The synthetic analysis data failed its load/validation check: %s", exc)
        st.error("The synthetic analysis data could not be loaded or validated. Check data/app.db and the setup instructions.")
        st.caption("The app requires the committed synthetic database with the documented schema.")
        st.stop()
    return pd.DataFrame()


def _reset_filters(start: date, end: date, regions: list[str], categories: list[str]) -> None:
    st.session_state["filter_date_range"] = (start, end)
    st.session_state["filter_regions"] = regions
    st.session_state["filter_categories"] = categories


def render_shared_filters(frame: pd.DataFrame) -> tuple[pd.DataFrame, tuple[date, date], list[str], list[str]]:
    if frame.empty:
        return frame.copy(), (date.today(), date.today()), [], []
    min_date = pd.to_datetime(frame["sale_date"]).min().date()
    max_date = pd.to_datetime(frame["sale_date"]).max().date()
    regions = sorted(frame["region"].dropna().astype(str).unique().tolist())
    categories = sorted(frame["category"].dropna().astype(str).unique().tolist())
    if "filter_date_range" not in st.session_state:
        st.session_state["filter_date_range"] = (min_date, max_date)
    if "filter_regions" not in st.session_state:
        st.session_state["filter_regions"] = regions
    if "filter_categories" not in st.session_state:
        st.session_state["filter_categories"] = categories

    st.sidebar.markdown("### Shared filters")
    selected_dates = st.sidebar.date_input(
        "Date range", min_value=min_date, max_value=max_date,
        key="filter_date_range", help="Both endpoints are included.",
    )
    selected_regions = st.sidebar.multiselect("Region", regions, key="filter_regions")
    selected_categories = st.sidebar.multiselect("Product category", categories, key="filter_categories")
    st.sidebar.button("Reset filters", on_click=_reset_filters, args=(min_date, max_date, regions, categories), use_container_width=True)

    if not isinstance(selected_dates, (tuple, list)) or len(selected_dates) != 2:
        st.sidebar.warning("Choose a start and end date.")
        return frame.iloc[0:0].copy(), (min_date, max_date), selected_regions, selected_categories
    start, end = selected_dates
    if start > end:
        st.sidebar.error("Start date must be on or before end date.")
        return frame.iloc[0:0].copy(), (start, end), selected_regions, selected_categories
    filtered = apply_filters(frame, (start, end), selected_regions, selected_categories)
    return filtered, (start, end), selected_regions, selected_categories


def render_page_context(title: str, subtitle: str = "") -> tuple[pd.DataFrame, tuple[date, date], list[str], list[str]]:
    render_header()
    st.title(title)
    if subtitle:
        st.caption(subtitle)
    st.caption("Synthetic retail dataset · values in INR · date, region and category filters are shared across analytical pages.")
    raw = load_page_data()
    return render_shared_filters(raw)


def show_active_slice(dates: tuple[date, date], regions: list[str], categories: list[str], row_count: int) -> None:
    st.caption(f"Selected: {dates[0]} to {dates[1]} · {len(regions)} region(s) · {len(categories)} category(ies) · {row_count:,} sales lines")


def show_empty(frame: pd.DataFrame, detail: str = "Try widening the date range or selecting additional regions/categories.") -> bool:
    if frame.empty:
        st.info("No eligible sales lines match the current selection.")
        st.caption(detail)
        return True
    return False
