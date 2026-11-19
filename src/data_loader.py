"""Data access and cleaning boundary for the Streamlit application."""

from __future__ import annotations

from pathlib import Path
import json

import pandas as pd
import streamlit as st

from src.db import DEFAULT_DB_PATH, read_connection
from src.validation import validate_database


@st.cache_data(show_spinner=False, ttl=300)
def load_analysis_data(db_path: str | Path = DEFAULT_DB_PATH, freshness_key: int | None = None) -> pd.DataFrame:
    """Load the approved joined view; cache only serializable data, never connections."""
    path = Path(db_path).expanduser().resolve()
    stamp = freshness_key if freshness_key is not None else path.stat().st_mtime_ns
    del stamp  # Included in the cache key; file freshness invalidates cached content.
    with read_connection(path) as connection:
        issues = validate_database(connection)
        if issues:
            raise ValueError("Database validation failed: " + "; ".join(str(x["message"]) for x in issues[:5]))
        frame = pd.read_sql_query("SELECT * FROM vw_sales_analysis ORDER BY sale_date, sales_line_id", connection)
    if frame.empty:
        return frame
    frame["sale_date"] = pd.to_datetime(frame["sale_date"], errors="raise")
    for column in ("gross_revenue", "discount_amount", "net_revenue", "line_cost", "profit_amount", "margin_pct", "monthly_target"):
        if column in frame:
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
    return frame


@st.cache_data(show_spinner=False, ttl=300)
def load_analysis_data_fresh(db_path: str | Path = DEFAULT_DB_PATH) -> pd.DataFrame:
    """Cache wrapper whose key always includes the current DB modification time."""
    path = Path(db_path).expanduser().resolve()
    return load_analysis_data(str(path), path.stat().st_mtime_ns)


@st.cache_data(show_spinner=False, ttl=300)
def load_target_data(db_path: str | Path = DEFAULT_DB_PATH) -> pd.DataFrame:
    """Load monthly targets with dimensions, at their own target grain."""
    path = Path(db_path).expanduser().resolve()
    with read_connection(path) as connection:
        return pd.read_sql_query(
            """SELECT t.target_month, t.store_id, s.store_name, s.region,
                      t.product_id, p.product_name, p.category, t.monthly_target
                 FROM fact_sales_target AS t
                 JOIN dim_store AS s ON s.store_id=t.store_id
                 JOIN dim_product AS p ON p.product_id=t.product_id
                ORDER BY t.target_month, t.store_id, t.product_id""",
            connection,
        )


def clean_sales_dataframe(frame: pd.DataFrame) -> tuple[pd.DataFrame, str | None]:
    """Impute missing payment method with mode without changing the input frame."""
    cleaned = frame.copy(deep=True)
    if "payment_method" not in cleaned or cleaned.empty or not cleaned["payment_method"].isna().any():
        return cleaned, None
    mode_values = cleaned["payment_method"].dropna().astype(str).value_counts()
    if mode_values.empty:
        return cleaned, None
    mode = sorted(mode_values.index, key=lambda value: (-mode_values[value], value))[0]
    cleaned["payment_method"] = cleaned["payment_method"].fillna(mode)
    return cleaned, mode


def serialize_filtered_csv(frame: pd.DataFrame) -> str:
    """Serialize a caller-filtered frame to UTF-8-friendly CSV text without index."""
    return frame.to_csv(index=False, lineterminator="\n")


def read_quality_report(report_path: str | Path) -> dict[str, object]:
    """Read the synthetic generator's JSON quality report."""
    path = Path(report_path)
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)
