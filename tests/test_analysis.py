"""Calculation contracts; src.analysis.py and src.statistics.py are built in Step 15."""
import pandas as pd
import pytest

from src.analysis import apply_filters, calculate_kpis, monthly_revenue_growth, compare_actual_to_target
from src.statistics import discount_margin_correlation


@pytest.fixture
def tiny_sales():
    # Hand-calculated values: net=250, cost=140, profit=110, orders=2,
    # AOV=125, returned lines=1/3, aggregate margin=44%.
    return pd.DataFrame([
        {"sales_line_id": 1, "order_id": "O-1", "sale_date": "2023-01-10", "sales_month": "2023-01", "product_id": "P-1", "category": "Grocery", "region": "North", "customer_segment": "Regular", "payment_method": "UPI", "net_revenue": 100.0, "line_cost": 50.0, "margin_pct": 50.0, "discount_pct": 0.0, "quantity": 1, "return_flag": 0},
        {"sales_line_id": 2, "order_id": "O-1", "sale_date": "2023-01-11", "sales_month": "2023-01", "product_id": "P-2", "category": "Grocery", "region": "North", "customer_segment": "Regular", "payment_method": "UPI", "net_revenue": 50.0, "line_cost": 20.0, "margin_pct": 60.0, "discount_pct": 0.1, "quantity": 1, "return_flag": 1},
        {"sales_line_id": 3, "order_id": "O-2", "sale_date": "2023-02-10", "sales_month": "2023-02", "product_id": "P-1", "category": "Grocery", "region": "South", "customer_segment": "Premium", "payment_method": "Card", "net_revenue": 100.0, "line_cost": 70.0, "margin_pct": 30.0, "discount_pct": 0.2, "quantity": 2, "return_flag": 0},
    ])


def test_all_four_dashboard_kpis_use_correct_denominators(tiny_sales):
    result = calculate_kpis(tiny_sales)
    assert result["net_revenue"] == pytest.approx(250.0)
    assert result["aov"] == pytest.approx(125.0)
    assert result["profit_margin_pct"] == pytest.approx(44.0)
    assert result["return_rate_pct"] == pytest.approx(100 / 3)


def test_monthly_growth_is_chronological_and_guards_zero_prior_month():
    frame = pd.DataFrame({
        "sale_date": ["2023-01-05", "2023-02-05", "2023-03-05"],
        "net_revenue": [0.0, 10.0, 15.0],
    })
    result = monthly_revenue_growth(frame).set_index("sales_month")
    assert list(result.index) == ["2023-01", "2023-02", "2023-03"]
    assert pd.isna(result.loc["2023-02", "mom_growth_pct"])
    assert result.loc["2023-03", "mom_growth_pct"] == pytest.approx(50.0)


def test_valid_category_correlation_reports_sample_count_and_result():
    frame = pd.DataFrame({
        "category": ["Grocery"] * 5,
        "discount_pct": [0.0, 0.1, 0.2, 0.3, 0.4],
        "margin_pct": [10.0, 20.0, 30.0, 40.0, 50.0],
    })
    result = discount_margin_correlation(frame, min_n=5).set_index("category")
    assert result.loc["Grocery", "n"] == 5
    assert result.loc["Grocery", "status"] == "valid"
    assert result.loc["Grocery", "pearson_r"] == pytest.approx(1.0)
    assert result.loc["Grocery", "p_value"] == pytest.approx(0.0, abs=1e-10)


@pytest.mark.parametrize("discounts,margins,n,reason", [
    ([0.1, 0.2, 0.3, 0.4], [10, 20, 30, 40], 4, "insufficient"),
    ([0.2] * 5, [10, 20, 30, 40, 50], 5, "constant"),
])
def test_invalid_correlation_is_withheld(discounts, margins, n, reason):
    frame = pd.DataFrame({"category": ["Clothing"] * n,
                          "discount_pct": discounts, "margin_pct": margins})
    result = discount_margin_correlation(frame, min_n=5).set_index("category")
    row = result.loc["Clothing"]
    assert row["n"] == n
    assert row["status"] == reason
    assert pd.isna(row["pearson_r"])
    assert pd.isna(row["p_value"])


def test_actual_target_aggregation_uses_matching_grain_without_line_multiplication():
    sales = pd.DataFrame({
        "sale_date": ["2023-01-10", "2023-01-11"],
        "store_id": ["S-1", "S-1"], "product_id": ["P-1", "P-1"],
        "net_revenue": [100.0, 150.0],
    })
    targets = pd.DataFrame({
        "target_month": ["2023-01-01"], "store_id": ["S-1"],
        "product_id": ["P-1"], "monthly_target": [200.0],
    })
    result = compare_actual_to_target(sales, targets).iloc[0]
    assert result["actual"] == pytest.approx(250.0)
    assert result["target"] == pytest.approx(200.0)
    assert result["variance"] == pytest.approx(50.0)
    assert result["variance_pct"] == pytest.approx(25.0)


def test_shared_filters_apply_date_region_and_category_together(tiny_sales):
    filtered = apply_filters(
        tiny_sales,
        ("2023-01-01", "2023-01-31"),
        regions=["North"],
        categories=["Grocery"],
    )
    assert filtered["sales_line_id"].tolist() == [1, 2]
    assert filtered["region"].eq("North").all()
    assert filtered["category"].eq("Grocery").all()


def test_kpi_calculation_guards_empty_and_zero_denominators():
    empty = pd.DataFrame(columns=["net_revenue", "line_cost", "order_id", "return_flag"])
    result = calculate_kpis(empty)
    assert result["net_revenue"] == 0
    assert result["aov"] is None
    assert result["profit_margin_pct"] is None
    assert result["return_rate_pct"] is None

    zero = pd.DataFrame([{
        "net_revenue": 0.0, "line_cost": 10.0, "order_id": "O-1", "return_flag": 0
    }])
    result = calculate_kpis(zero)
    assert result["aov"] == 0
    assert result["profit_margin_pct"] is None
    assert result["return_rate_pct"] == 0
