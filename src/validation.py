"""Schema, data-quality, and analysis-frame validation helpers."""

from __future__ import annotations

from collections.abc import Iterable
import sqlite3

import pandas as pd


REQUIRED_OBJECTS = {
    "dim_product": {"product_id", "sku", "product_name", "category", "subcategory", "list_price", "standard_unit_cost", "active_flag"},
    "dim_store": {"store_id", "store_code", "store_name", "region", "city", "active_flag"},
    "fact_sales": {"sales_line_id", "order_id", "sale_date", "product_id", "store_id", "customer_segment", "payment_method", "quantity", "unit_price_at_sale", "unit_cost_at_sale", "discount_pct", "return_flag"},
    "fact_sales_target": {"target_month", "store_id", "product_id", "monthly_target"},
    "vw_sales_analysis": {"sales_line_id", "order_id", "sale_date", "product_id", "product_name", "category", "store_id", "region", "customer_segment", "payment_method", "quantity", "discount_pct", "net_revenue", "line_cost", "profit_amount", "margin_pct", "return_flag", "sales_month", "calendar_year", "calendar_quarter"},
}
REQUIRED_SALES_COLUMNS = {"sales_line_id", "sale_date", "product_id", "store_id", "quantity", "discount_pct", "payment_method"}


def _issue(code: str, row_index: object, field: str, message: str) -> dict[str, object]:
    return {"code": code, "row_index": row_index, "field": field, "message": message}


def validate_sales_frame(frame: pd.DataFrame) -> list[dict[str, object]]:
    """Return row-level issues using stable codes; this function does not mutate."""
    issues: list[dict[str, object]] = []
    for column in sorted(REQUIRED_SALES_COLUMNS - set(frame.columns)):
        issues.append(_issue("MISSING_COLUMN", None, column, f"Required column '{column}' is absent."))
    if not REQUIRED_SALES_COLUMNS.issubset(frame.columns):
        return issues

    for idx in frame.index[frame[list(REQUIRED_SALES_COLUMNS)].isna().any(axis=1)]:
        issues.append(_issue("MISSING_REQUIRED", idx, "required", "A required sales field is missing."))
    dates = pd.to_datetime(frame["sale_date"], errors="coerce", format="mixed")
    for idx in frame.index[dates.isna()]:
        issues.append(_issue("INVALID_DATE", idx, "sale_date", "Date is missing or cannot be parsed."))
    quantity = pd.to_numeric(frame["quantity"], errors="coerce")
    for idx in frame.index[quantity.isna() | (quantity <= 0) | (quantity % 1 != 0)]:
        issues.append(_issue("INVALID_QUANTITY", idx, "quantity", "Quantity must be a positive whole number."))
    discount = pd.to_numeric(frame["discount_pct"], errors="coerce")
    for idx in frame.index[discount.isna() | (discount < 0) | (discount > 0.80)]:
        issues.append(_issue("INVALID_DISCOUNT", idx, "discount_pct", "Discount must be between 0 and 0.80."))
    duplicate_mask = frame["sales_line_id"].notna() & frame["sales_line_id"].duplicated(keep=False)
    for idx in frame.index[duplicate_mask]:
        issues.append(_issue("DUPLICATE_KEY", idx, "sales_line_id", "Sales line key occurs more than once."))
    return issues


def validate_database(connection: sqlite3.Connection) -> list[dict[str, object]]:
    """Validate required SQLite objects/columns and referential integrity."""
    issues: list[dict[str, object]] = []
    objects = {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view')")}
    for name, required_columns in REQUIRED_OBJECTS.items():
        if name not in objects:
            issues.append(_issue("MISSING_OBJECT", None, name, f"Required table/view '{name}' is absent."))
            continue
        actual = {row[1] for row in connection.execute(f'PRAGMA table_info("{name}")')}
        for column in sorted(required_columns - actual):
            issues.append(_issue("MISSING_COLUMN", None, f"{name}.{column}", "Required column is absent."))
    if "fact_sales" in objects and "dim_product" in objects and "dim_store" in objects:
        fk_rows = connection.execute("PRAGMA foreign_key_check").fetchall()
        for row in fk_rows:
            issues.append(_issue("INVALID_FOREIGN_KEY", row[1], row[2], "Foreign-key integrity check failed."))
    return issues


def quality_summary(issues: Iterable[dict[str, object]]) -> pd.DataFrame:
    """Summarize validation issues by stable issue code."""
    rows = list(issues)
    if not rows:
        return pd.DataFrame(columns=["check", "affected_rows", "status"])
    frame = pd.DataFrame(rows)
    result = frame.groupby("code", dropna=False).size().rename("affected_rows").reset_index()
    result = result.rename(columns={"code": "check"})
    result["status"] = "Review"
    return result
