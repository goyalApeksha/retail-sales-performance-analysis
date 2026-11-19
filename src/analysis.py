"""Pure analytical functions for the approved retail-sales measures."""

from __future__ import annotations

from collections.abc import Sequence

import numpy as np
import pandas as pd


def calculate_kpis(frame: pd.DataFrame) -> dict[str, float | None]:
    """Return revenue, AOV, aggregate margin, and line-based return rate."""
    if frame.empty:
        return {"net_revenue": 0.0, "aov": None, "profit_margin_pct": None, "return_rate_pct": None}
    revenue = pd.to_numeric(frame.get("net_revenue", pd.Series(dtype=float)), errors="coerce").fillna(0)
    total_revenue = float(revenue.sum())
    cost = pd.to_numeric(frame.get("line_cost", pd.Series(0, index=frame.index)), errors="coerce").fillna(0)
    order_count = int(frame["order_id"].nunique(dropna=True)) if "order_id" in frame else 0
    return_flags = pd.to_numeric(frame.get("return_flag", pd.Series(0, index=frame.index)), errors="coerce")
    eligible_returns = return_flags.dropna()
    return_rate = (float(eligible_returns.sum()) / len(eligible_returns) * 100) if len(eligible_returns) else None
    margin = ((total_revenue - float(cost.sum())) / total_revenue * 100) if total_revenue != 0 else None
    aov = total_revenue / order_count if order_count else None
    return {
        "net_revenue": total_revenue,
        "aov": float(aov) if aov is not None else None,
        "profit_margin_pct": float(margin) if margin is not None else None,
        "return_rate_pct": float(return_rate) if return_rate is not None else None,
    }


def monthly_revenue_growth(frame: pd.DataFrame) -> pd.DataFrame:
    """Aggregate revenue by month and calculate guarded MoM/YoY growth."""
    columns = ["sales_month", "net_revenue", "mom_growth_pct", "yoy_growth_pct"]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    work = frame.copy()
    dates = pd.to_datetime(work["sale_date"], errors="coerce")
    work = work.loc[dates.notna()].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)
    work["sales_month"] = dates.loc[work.index].dt.to_period("M").astype(str)
    monthly = work.groupby("sales_month", as_index=False)["net_revenue"].sum().sort_values("sales_month")
    full_months = pd.period_range(monthly["sales_month"].iloc[0], monthly["sales_month"].iloc[-1], freq="M").astype(str)
    monthly = (monthly.set_index("sales_month").reindex(full_months, fill_value=0)
               .rename_axis("sales_month").reset_index())
    monthly["mom_growth_pct"] = monthly["net_revenue"].pct_change(fill_method=None) * 100
    monthly["yoy_growth_pct"] = monthly["net_revenue"].pct_change(periods=12, fill_method=None) * 100
    monthly.replace([np.inf, -np.inf], np.nan, inplace=True)
    return monthly[columns].reset_index(drop=True)


def compare_actual_to_target(
    sales: pd.DataFrame,
    targets: pd.DataFrame,
    grain: Sequence[str] = ("sales_month", "store_id", "product_id"),
) -> pd.DataFrame:
    """Aggregate sales and target facts independently, then join at matching grain."""
    sales_work = sales.copy()
    if "sales_month" not in sales_work and "sale_date" in sales_work:
        sales_work["sales_month"] = pd.to_datetime(sales_work["sale_date"], errors="coerce").dt.to_period("M").astype(str)
    target_work = targets.copy()
    if "sales_month" not in target_work and "target_month" in target_work:
        target_work["sales_month"] = pd.to_datetime(target_work["target_month"], errors="coerce").dt.to_period("M").astype(str)
    keys = list(grain)
    actual = sales_work.groupby(keys, dropna=False, as_index=False)["net_revenue"].sum().rename(columns={"net_revenue": "actual"})
    target = target_work.groupby(keys, dropna=False, as_index=False)["monthly_target"].sum().rename(columns={"monthly_target": "target"})
    result = actual.merge(target, on=keys, how="outer", validate="one_to_one")
    result[["actual", "target"]] = result[["actual", "target"]].fillna(0.0)
    result["variance"] = result["actual"] - result["target"]
    result["variance_pct"] = (result["variance"] / result["target"].replace(0, np.nan) * 100)
    return result.sort_values(keys).reset_index(drop=True)


def product_revenue(frame: pd.DataFrame, top_n: int | None = None) -> pd.DataFrame:
    result = frame.groupby(["product_id", "product_name", "category"], as_index=False)["net_revenue"].sum()
    result = result.sort_values("net_revenue", ascending=False).reset_index(drop=True)
    return result.head(top_n) if top_n else result


def category_margin(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.groupby("category", as_index=False).agg(net_revenue=("net_revenue", "sum"), line_cost=("line_cost", "sum"))
    result["profit_amount"] = result["net_revenue"] - result["line_cost"]
    result["margin_pct"] = result["profit_amount"].div(result["net_revenue"].replace(0, np.nan)) * 100
    return result.sort_values("margin_pct", ascending=False, na_position="last").reset_index(drop=True)


def grouped_aov(frame: pd.DataFrame, groups: Sequence[str] = ("region", "customer_segment")) -> pd.DataFrame:
    grouped = frame.groupby(list(groups), dropna=False).agg(net_revenue=("net_revenue", "sum"), order_count=("order_id", "nunique")).reset_index()
    grouped["aov"] = grouped["net_revenue"].div(grouped["order_count"].replace(0, np.nan))
    return grouped


def monthly_volume(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["sales_month", "quantity"])
    work = frame.assign(sales_month=pd.to_datetime(frame["sale_date"]).dt.to_period("M").astype(str))
    return work.groupby("sales_month", as_index=False)["quantity"].sum().sort_values("sales_month").reset_index(drop=True)


def pareto_analysis(frame: pd.DataFrame) -> pd.DataFrame:
    result = product_revenue(frame)
    total = float(result["net_revenue"].sum()) if not result.empty else 0.0
    result["cumulative_revenue"] = result["net_revenue"].cumsum()
    result["cumulative_share_pct"] = result["cumulative_revenue"] / total * 100 if total else np.nan
    return result


def payment_summary(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.groupby("payment_method", dropna=False).agg(
        net_revenue=("net_revenue", "sum"), order_count=("order_id", "nunique")
    ).reset_index()
    result["aov"] = result["net_revenue"].div(result["order_count"].replace(0, np.nan))
    return result.sort_values("order_count", ascending=False).reset_index(drop=True)


def returns_by_category(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.groupby("category", dropna=False).agg(
        returned_count=("return_flag", "sum"), eligible_count=("return_flag", "count")
    ).reset_index().assign(return_rate_pct=lambda x: x["returned_count"].div(x["eligible_count"].replace(0, np.nan)) * 100)


def sales_velocity(frame: pd.DataFrame, period_days: int) -> pd.DataFrame:
    if period_days <= 0:
        raise ValueError("The selected period must contain at least one calendar day.")
    result = frame.groupby(["product_id", "product_name", "category"], as_index=False)["quantity"].sum()
    result["units_per_day"] = result["quantity"] / period_days
    return result.sort_values("units_per_day", ascending=False).reset_index(drop=True)


def apply_filters(
    frame: pd.DataFrame,
    date_range: tuple[object, object] | None = None,
    regions: Sequence[str] | None = None,
    categories: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Return a filtered copy; empty selections mean all values."""
    result = frame.copy()
    if date_range and len(date_range) == 2:
        dates = pd.to_datetime(result["sale_date"])
        result = result.loc[dates.between(pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]))]
    if regions:
        result = result.loc[result["region"].isin(regions)]
    if categories:
        result = result.loc[result["category"].isin(categories)]
    return result.copy()
