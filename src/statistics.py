"""Statistical summaries with explicit validity checks."""

from __future__ import annotations

import numpy as np
import pandas as pd
from scipy.stats import kruskal, pearsonr


def discount_margin_correlation(frame: pd.DataFrame, min_n: int = 5) -> pd.DataFrame:
    """Return Pearson r/p by category only when sample and variation are valid."""
    columns = ["category", "n", "pearson_r", "p_value", "status", "reason"]
    if frame.empty or "category" not in frame:
        return pd.DataFrame(columns=columns)
    rows = []
    for category, group in frame.groupby("category", dropna=False):
        pairs = group[["discount_pct", "margin_pct"]].apply(pd.to_numeric, errors="coerce").dropna()
        n = len(pairs)
        row = {"category": category, "n": n, "pearson_r": np.nan, "p_value": np.nan, "status": "valid", "reason": ""}
        if n < min_n:
            row.update(status="insufficient", reason=f"Requires at least {min_n} eligible observations.")
        elif pairs["discount_pct"].nunique() < 2 or pairs["margin_pct"].nunique() < 2:
            row.update(status="constant", reason="Discount or margin is constant in the eligible sample.")
        else:
            result = pearsonr(pairs["discount_pct"], pairs["margin_pct"])
            row.update(pearson_r=float(result.statistic), p_value=float(result.pvalue))
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def payment_aov_group_test(frame: pd.DataFrame, min_group_n: int = 2) -> dict[str, float | int | str | None]:
    """Optional Kruskal-Wallis test of per-order value by payment method."""
    orders = frame.groupby(["payment_method", "order_id"], dropna=True, as_index=False)["net_revenue"].sum()
    groups = [pd.to_numeric(group["net_revenue"], errors="coerce").dropna().to_numpy()
              for _, group in orders.groupby("payment_method")]
    groups = [values for values in groups if len(values) >= min_group_n]
    if len(groups) < 2:
        return {"status": "insufficient", "n_groups": len(groups), "statistic": None, "p_value": None}
    result = kruskal(*groups)
    return {"status": "valid", "n_groups": len(groups), "statistic": float(result.statistic), "p_value": float(result.pvalue)}


def descriptive_statistics(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Provide count, mean, median, standard deviation and quartiles."""
    numeric = frame[columns].apply(pd.to_numeric, errors="coerce")
    return numeric.describe(percentiles=[0.25, 0.5, 0.75]).T.rename(columns={"25%": "q1", "50%": "median", "75%": "q3"})
