"""Generate the Step 8 synthetic retail dataset and its validation artifacts.

Reproducible with seed 42. The Step 8 workbook is the schema authority. Raw
CSV/Excel exports preserve deliberate data-quality defects for cleaning practice;
SQLite applies the documented missing-value imputation before loading the exact
Step 8 NOT NULL / CHECK / FK schema. Duplicate business lines and valid-range
statistical outliers are retained in both forms and reported.

Requires Python 3.12, numpy, and openpyxl.
Run from the repository root: python src/generate_data.py
"""

from __future__ import annotations

import csv
import json
import sqlite3
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SEED = 42
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
SCHEMA_WORKBOOK = PROJECT_ROOT / "docs" / "08_Database_Design.xlsx"
EXCEL_PATH = DATA_DIR / "retail_sales_performance_analysis_data.xlsx"
DB_PATH = DATA_DIR / "app.db"
REPORT_JSON = DATA_DIR / "validation_report.json"
REPORT_TXT = DATA_DIR / "validation_report.txt"

TABLE_COLUMNS: dict[str, list[str]] = {
    "dim_product": [
        "product_id", "sku", "product_name", "category", "subcategory",
        "list_price", "standard_unit_cost", "active_flag",
    ],
    "dim_store": [
        "store_id", "store_code", "store_name", "region", "city", "active_flag",
    ],
    "fact_sales": [
        "sales_line_id", "order_id", "sale_date", "product_id", "store_id",
        "customer_segment", "payment_method", "quantity", "unit_price_at_sale",
        "unit_cost_at_sale", "discount_pct", "return_flag",
    ],
    "fact_sales_target": [
        "target_month", "store_id", "product_id", "monthly_target",
    ],
}

CATEGORY_PRODUCTS = {
    "Electronics": [
        ("Smartphone", "Aarohi 5G Phone", 24999.0),
        ("Audio", "Madhav Wireless Buds", 3499.0),
        ("Accessories", "Tara Fast Charger", 1299.0),
        ("Appliances", "Ishaan LED Lamp", 2199.0),
    ],
    "Clothing": [
        ("Ethnic Wear", "Meera Cotton Kurta", 1899.0),
        ("Casual Wear", "Kabir Linen Shirt", 1599.0),
        ("Footwear", "Saanvi Walking Sandal", 1299.0),
        ("Accessories", "Reyansh Woven Scarf", 699.0),
    ],
    "Grocery": [
        ("Staples", "Anaya Basmati Rice 5kg", 649.0),
        ("Spices", "Vihaan Masala Box", 299.0),
        ("Beverages", "Kiara Filter Coffee Pack", 389.0),
        ("Snacks", "Arjun Millet Snack Pack", 179.0),
    ],
    "Home & Kitchen": [
        ("Cookware", "Diya Steel Cook Set", 2499.0),
        ("Storage", "Aarav Stackable Jars", 899.0),
        ("Textiles", "Myra Cotton Bedsheet", 1799.0),
        ("Decor", "Rohan Ceramic Planter", 799.0),
    ],
    "Sports": [
        ("Fitness", "Ira Yoga Mat", 999.0),
        ("Outdoor", "Dev Cricket Bat", 2299.0),
        ("Fitness", "Nisha Training Bands", 549.0),
        ("Outdoor", "Aditya Trek Bottle", 749.0),
    ],
}

REGION_CITIES = {
    "North": [("Indrapur", "Aarohi Indrapur Bazaar"), ("Nandipur", "Madhav Nandipur Market")],
    "South": [("Suryanagar", "Tara Suryanagar Bazaar"), ("Devnagar", "Ishaan Devnagar Market")],
    "East": [("Kesavgarh", "Meera Kesavgarh Bazaar"), ("Shantipur", "Kabir Shantipur Market")],
    "West": [("Navapur", "Saanvi Navapur Bazaar"), ("Vasantgarh", "Reyansh Vasantgarh Market")],
    "Central": [("Ratanpur", "Anaya Ratanpur Bazaar"), ("Padmapur", "Vihaan Padmapur Market")],
}

REGIONS = ["North", "South", "East", "West", "Central"]
CATEGORIES = list(CATEGORY_PRODUCTS)
SEGMENTS = ["Regular", "Premium", "Wholesale"]
PAYMENT_METHODS = ["UPI", "Card", "Cash", "Net Banking"]
QUALITY_MISSING_IDS = [14, 23, 70, 102, 170, 188]
QUALITY_OUTLIER_IDS = [44, 101, 222]
QUALITY_DUPLICATES = [(238, 8), (239, 79), (240, 153)]  # duplicate line ID -> source line ID


def _row(**kwargs: object) -> dict[str, object]:
    return kwargs


def build_products() -> list[dict[str, object]]:
    rng = np.random.default_rng(SEED + 1)
    rows: list[dict[str, object]] = []
    number = 1
    for category, products in CATEGORY_PRODUCTS.items():
        for subcategory, name, base_price in products:
            price = round(base_price * float(rng.uniform(0.94, 1.06)), 2)
            cost_rate = float(rng.uniform(0.54, 0.72))
            rows.append(_row(
                product_id=f"P{number:03d}", sku=f"SYN-{number:04d}",
                product_name=name, category=category, subcategory=subcategory,
                list_price=price, standard_unit_cost=round(price * cost_rate, 2), active_flag=1,
            ))
            number += 1
    return rows


def build_stores() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    n = 1
    for region, locations in REGION_CITIES.items():
        for city, name in locations:
            rows.append(_row(
                store_id=f"ST{n:03d}", store_code=f"SYN-ST-{n:02d}",
                store_name=name, region=region, city=city, active_flag=1,
            ))
            n += 1
    return rows


def _weighted_choice(rng: np.random.Generator, values: list[object], weights: list[float]):
    weights_array = np.asarray(weights, dtype=float)
    weights_array /= weights_array.sum()
    return values[int(rng.choice(len(values), p=weights_array))]


def build_sales(products: list[dict[str, object]], stores: list[dict[str, object]]) -> list[dict[str, object]]:
    rng = np.random.default_rng(SEED + 2)
    products_by_category: dict[str, list[dict[str, object]]] = defaultdict(list)
    for product in products:
        products_by_category[str(product["category"])].append(product)
    store_by_region: dict[str, list[dict[str, object]]] = defaultdict(list)
    for store in stores:
        store_by_region[str(store["region"])].append(store)

    rows: list[dict[str, object]] = []
    line_id = 1
    for month_offset in range(24):
        year = 2023 + month_offset // 12
        month = month_offset % 12 + 1
        month_start = date(year, month, 1)
        days_in_month = (date(year + (month == 12), month % 12 + 1, 1) - month_start).days

        for _ in range(10):  # exactly 10 synthetic order lines per month
            q4 = month in (11, 12)
            # Pattern A: electronics form a visibly larger share in Nov-Dec.
            category_weights = [0.39 if q4 and c == "Electronics" else 0.17 if q4 else
                                0.16 if c == "Electronics" else 0.21 for c in CATEGORIES]
            category = str(_weighted_choice(rng, CATEGORIES, category_weights))
            product = products_by_category[category][int(rng.integers(0, 4))]

            # Pattern B: North and West stores receive higher sales propensity.
            region_weights = [1.28 if r in ("North", "West") else 0.82 if r == "East" else 1.0 for r in REGIONS]
            region = str(_weighted_choice(rng, REGIONS, region_weights))
            store_group = store_by_region[region]
            store = store_group[int(rng.integers(0, len(store_group)))]

            day = int(rng.integers(0, days_in_month))
            sale_date = month_start + timedelta(days=day)
            year_uplift = 1.10 if year == 2024 else 1.0
            seasonal_qty = 1.35 if q4 else 1.0
            weekday_qty = 1.12 if sale_date.weekday() >= 5 else 1.0
            qty_mean = 1.45 * year_uplift * seasonal_qty * weekday_qty
            quantity = int(np.clip(rng.poisson(qty_mean) + 1, 1, 8))

            list_price = float(product["list_price"])
            base_price = list_price * float(rng.uniform(0.94, 1.04))
            # Pattern C: discounts are more common in selected low season months;
            # the sales view's margin calculation therefore tends to fall as discount rises.
            discount_choices = [0.00, 0.05, 0.10, 0.15, 0.20, 0.25]
            discount_weights = [0.22, 0.25, 0.23, 0.16, 0.10, 0.04]
            if month in (3, 7, 8):
                discount_weights = [0.10, 0.20, 0.25, 0.22, 0.15, 0.08]
            discount = float(_weighted_choice(rng, discount_choices, discount_weights))
            year_price_factor = 1.025 if year == 2024 else 1.0
            unit_price = round(base_price * year_price_factor, 2)
            unit_cost = round(list_price * float(rng.uniform(0.53, 0.73)), 2)

            segment = str(_weighted_choice(rng, SEGMENTS, [0.65, 0.23, 0.12]))
            payment = str(_weighted_choice(rng, PAYMENT_METHODS, [0.43, 0.31, 0.16, 0.10]))
            return_probability = 0.11 if category in ("Clothing", "Electronics") else 0.055
            return_flag = int(rng.random() < return_probability)
            rows.append(_row(
                sales_line_id=line_id,
                order_id=f"ORD-{year}-{month:02d}-{line_id:04d}",
                sale_date=sale_date.isoformat(), product_id=product["product_id"],
                store_id=store["store_id"], customer_segment=segment,
                payment_method=payment, quantity=quantity, unit_price_at_sale=unit_price,
                unit_cost_at_sale=unit_cost, discount_pct=discount, return_flag=return_flag,
            ))
            line_id += 1

    # Three intentional duplicate business records. The generated line key stays unique;
    # every other sales field duplicates an earlier line exactly.
    by_id = {int(row["sales_line_id"]): row for row in rows}
    for duplicate_id, source_id in QUALITY_DUPLICATES:
        clone = dict(by_id[source_id])
        clone["sales_line_id"] = duplicate_id
        rows[duplicate_id - 1] = clone

    # Three intentionally unusual but schema-valid unit-price values.
    for line_id in QUALITY_OUTLIER_IDS:
        row = rows[line_id - 1]
        product = next(p for p in products if p["product_id"] == row["product_id"])
        multiplier = {44: 4.2, 101: 4.8, 222: 5.5}[line_id]
        row["unit_price_at_sale"] = round(float(product["list_price"]) * multiplier, 2)

    # Missing raw payment methods are intentionally isolated to six non-key cells (2.5% of fact rows).
    for line_id in QUALITY_MISSING_IDS:
        rows[line_id - 1]["payment_method"] = None
    return rows


def build_targets(sales: list[dict[str, object]], products: list[dict[str, object]],
                  stores: list[dict[str, object]]) -> list[dict[str, object]]:
    rng = np.random.default_rng(SEED + 3)
    product_by_id = {str(p["product_id"]): p for p in products}
    actuals: dict[tuple[str, str, str], float] = defaultdict(float)
    for row in sales:
        month = str(row["sale_date"])[:7] + "-01"
        product = product_by_id[str(row["product_id"])]
        net = int(row["quantity"]) * float(row["unit_price_at_sale"]) * (1 - float(row["discount_pct"]))
        actuals[(month, str(row["store_id"]), str(product["product_id"]))] += net

    rows: list[dict[str, object]] = []
    for month_offset in range(24):
        year = 2023 + month_offset // 12
        month = month_offset % 12 + 1
        target_month = f"{year}-{month:02d}-01"
        seasonal = 1.22 if month in (11, 12) else 0.96 if month in (3, 7, 8) else 1.0
        trend = 1.06 if year == 2024 else 1.0
        for store in stores:
            for product in products:
                key = (target_month, str(store["store_id"]), str(product["product_id"]))
                # About 0.05 sales lines are expected for each product-store-month cell.
                expected = float(product["list_price"]) * 2.0 * 0.92 * (10 / (len(stores) * len(products)))
                target = expected * seasonal * trend * float(rng.uniform(0.90, 1.16))
                actual = actuals.get(key, 0.0)
                # For active cells, targets are usually moderately above observed actuals,
                # with a small set of naturally exceeded targets for comparison work.
                if actual > 0:
                    target = max(target, actual * float(rng.uniform(0.98, 1.28)))
                rows.append(_row(
                    target_month=target_month, store_id=store["store_id"],
                    product_id=product["product_id"], monthly_target=round(target, 2),
                ))
    return rows


def read_schema_ddl() -> str:
    if not SCHEMA_WORKBOOK.exists():
        raise FileNotFoundError(f"Required Step 8 schema workbook not found: {SCHEMA_WORKBOOK}")
    workbook = load_workbook(SCHEMA_WORKBOOK, read_only=True, data_only=True)
    try:
        sheet = workbook["SQLite_DDL"]
        statements = []
        for (value,) in sheet.iter_rows(min_row=4, max_col=1, values_only=True):
            if value is None or str(value).strip() == "SQL statement":
                continue
            statements.append(str(value))
        ddl = "\n".join(statements)
        required = ["CREATE TABLE dim_product", "CREATE TABLE dim_store", "CREATE TABLE fact_sales (",
                    "CREATE TABLE fact_sales_target", "CREATE VIEW vw_sales_analysis"]
        missing = [token for token in required if token not in ddl]
        if missing:
            raise ValueError(f"Step 8 DDL sheet is incomplete; missing: {missing}")
        return ddl
    finally:
        workbook.close()


def raw_sales_with_missing(sales: list[dict[str, object]]) -> list[dict[str, object]]:
    return [dict(row) for row in sales]


def impute_missing_for_database(sales: list[dict[str, object]]) -> tuple[list[dict[str, object]], str]:
    observed = [str(row["payment_method"]) for row in sales if row["payment_method"] is not None]
    counts = Counter(observed)
    mode = sorted(counts, key=lambda value: (-counts[value], value))[0]
    clean = []
    for row in sales:
        copy = dict(row)
        if copy["payment_method"] is None:
            copy["payment_method"] = mode
        clean.append(copy)
    return clean, mode


def write_csv(table: str, rows: list[dict[str, object]]) -> None:
    path = DATA_DIR / f"{table}.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=TABLE_COLUMNS[table], extrasaction="raise")
        writer.writeheader()
        for row in rows:
            writer.writerow({key: "" if row.get(key) is None else row.get(key) for key in TABLE_COLUMNS[table]})


def write_excel(tables: dict[str, list[dict[str, object]]]) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_fill = PatternFill("solid", fgColor="17365D")
    for table_name, rows in tables.items():
        sheet = workbook.create_sheet(table_name)
        sheet.append(TABLE_COLUMNS[table_name])
        for row in rows:
            sheet.append([row.get(column) for column in TABLE_COLUMNS[table_name]])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.row_dimensions[1].height = 25
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for column_cells in sheet.columns:
            values = [str(cell.value) if cell.value is not None else "" for cell in list(column_cells)[:100]]
            width = min(max(max((len(value) for value in values), default=12) + 2, 12), 34)
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = width
        sheet.sheet_view.showGridLines = False
    workbook.save(EXCEL_PATH)
    workbook.close()


def load_database(ddl: str, tables: dict[str, list[dict[str, object]]]) -> tuple[sqlite3.Connection, str]:
    if DB_PATH.exists():
        DB_PATH.unlink()
    connection = sqlite3.connect(DB_PATH)
    connection.execute("PRAGMA foreign_keys = ON")
    connection.executescript(ddl)
    for table_name, rows in tables.items():
        columns = TABLE_COLUMNS[table_name]
        placeholders = ",".join("?" for _ in columns)
        column_names = ",".join(f'"{c}"' for c in columns)
        sql = f'INSERT INTO "{table_name}" ({column_names}) VALUES ({placeholders})'
        connection.executemany(sql, [tuple(row[c] for c in columns) for row in rows])
    connection.commit()
    return connection, "Step 8 SQLite_DDL sheet executed without modification."


def validate(connection: sqlite3.Connection, raw_tables: dict[str, list[dict[str, object]]],
             db_tables: dict[str, list[dict[str, object]]], imputed_mode: str) -> dict[str, object]:
    expected = {"dim_product": 20, "dim_store": 10, "fact_sales": 240, "fact_sales_target": 4800}
    counts = {name: int(connection.execute(f'SELECT COUNT(*) FROM "{name}"').fetchone()[0])
              for name in TABLE_COLUMNS}
    fk_violations = connection.execute("PRAGMA foreign_key_check").fetchall()
    fk_enforcement_enabled = int(connection.execute("PRAGMA foreign_keys").fetchone()[0]) == 1
    view_columns = [row[1] for row in connection.execute("PRAGMA table_info(vw_sales_analysis)")]
    view_count = int(connection.execute("SELECT COUNT(*) FROM vw_sales_analysis").fetchone()[0])
    range_checks = {
        "product_prices_nonnegative": connection.execute(
            "SELECT COUNT(*) FROM dim_product WHERE list_price < 0 OR standard_unit_cost < 0").fetchone()[0] == 0,
        "sales_dates_valid_and_in_range": connection.execute(
            "SELECT COUNT(*) FROM fact_sales WHERE date(sale_date) IS NULL OR sale_date NOT BETWEEN '2023-01-01' AND '2024-12-31'").fetchone()[0] == 0,
        "sales_quantity_positive": connection.execute(
            "SELECT COUNT(*) FROM fact_sales WHERE quantity <= 0").fetchone()[0] == 0,
        "sales_prices_costs_nonnegative": connection.execute(
            "SELECT COUNT(*) FROM fact_sales WHERE unit_price_at_sale < 0 OR unit_cost_at_sale < 0").fetchone()[0] == 0,
        "discount_between_zero_and_0_80": connection.execute(
            "SELECT COUNT(*) FROM fact_sales WHERE discount_pct NOT BETWEEN 0 AND 0.80").fetchone()[0] == 0,
        "returns_and_active_flags_binary": connection.execute(
            "SELECT (SELECT COUNT(*) FROM fact_sales WHERE return_flag NOT IN (0,1)) + (SELECT COUNT(*) FROM dim_product WHERE active_flag NOT IN (0,1)) + (SELECT COUNT(*) FROM dim_store WHERE active_flag NOT IN (0,1))").fetchone()[0] == 0,
        "targets_valid_month_and_nonnegative": connection.execute(
            "SELECT COUNT(*) FROM fact_sales_target WHERE date(target_month) IS NULL OR substr(target_month,9,2)!='01' OR monthly_target < 0 OR target_month NOT BETWEEN '2023-01-01' AND '2024-12-01'").fetchone()[0] == 0,
        "allowed_categories_segments_payments_regions": connection.execute(
            "SELECT (SELECT COUNT(*) FROM dim_product WHERE category NOT IN ('Electronics','Clothing','Grocery','Home & Kitchen','Sports')) + (SELECT COUNT(*) FROM dim_store WHERE region NOT IN ('North','South','East','West','Central')) + (SELECT COUNT(*) FROM fact_sales WHERE customer_segment NOT IN ('Regular','Premium','Wholesale') OR payment_method NOT IN ('UPI','Card','Cash','Net Banking'))").fetchone()[0] == 0,
        "required_columns_not_null": all(
            connection.execute(f'SELECT COUNT(*) FROM "{table}" WHERE ' + " OR ".join(f'"{col}" IS NULL' for col in columns)).fetchone()[0] == 0
            for table, columns in TABLE_COLUMNS.items()),
        "primary_and_unique_keys": (
            connection.execute("SELECT COUNT(*) FROM (SELECT product_id FROM dim_product GROUP BY product_id HAVING COUNT(*)>1)").fetchone()[0] == 0
            and connection.execute("SELECT COUNT(*) FROM (SELECT sku FROM dim_product GROUP BY sku HAVING COUNT(*)>1)").fetchone()[0] == 0
            and connection.execute("SELECT COUNT(*) FROM (SELECT store_id FROM dim_store GROUP BY store_id HAVING COUNT(*)>1)").fetchone()[0] == 0
            and connection.execute("SELECT COUNT(*) FROM (SELECT store_code FROM dim_store GROUP BY store_code HAVING COUNT(*)>1)").fetchone()[0] == 0
            and connection.execute("SELECT COUNT(*) FROM (SELECT sales_line_id FROM fact_sales GROUP BY sales_line_id HAVING COUNT(*)>1)").fetchone()[0] == 0
            and connection.execute("SELECT COUNT(*) FROM (SELECT target_month,store_id,product_id FROM fact_sales_target GROUP BY target_month,store_id,product_id HAVING COUNT(*)>1)").fetchone()[0] == 0
        ),
    }
    sales = raw_tables["fact_sales"]
    missing_count = sum(row["payment_method"] is None for row in sales)
    duplicate_groups = []
    sales_cols_no_pk = [col for col in TABLE_COLUMNS["fact_sales"] if col != "sales_line_id"]
    grouped: dict[tuple[object, ...], list[int]] = defaultdict(list)
    for row in sales:
        # Treat blanks and null markers as equal for detecting business duplicates.
        grouped[tuple(row[col] for col in sales_cols_no_pk)].append(int(row["sales_line_id"]))
    for ids in grouped.values():
        if len(ids) > 1:
            duplicate_groups.append(sorted(ids))
    outliers = [{"sales_line_id": i,
                 "unit_price_at_sale": float(sales[i-1]["unit_price_at_sale"]),
                 "list_price": float(next(p["list_price"] for p in raw_tables["dim_product"] if p["product_id"] == sales[i-1]["product_id"])),
                 "ratio_to_list_price": round(float(sales[i-1]["unit_price_at_sale"]) / float(next(p["list_price"] for p in raw_tables["dim_product"] if p["product_id"] == sales[i-1]["product_id"])), 2)}
                for i in QUALITY_OUTLIER_IDS]
    view_column_check = 30 <= len(view_columns) <= 40 and len(view_columns) == 35
    passed = (all(counts.get(name) == count for name, count in expected.items())
              and not fk_violations and fk_enforcement_enabled
              and all(range_checks.values()) and view_column_check)
    year_totals = {str(year): float(total) for year, total in connection.execute(
        "SELECT calendar_year,SUM(net_revenue) FROM vw_sales_analysis GROUP BY calendar_year")}
    yoy_change_pct = round((year_totals["2024"] / year_totals["2023"] - 1) * 100, 1)
    q4_avg, q4_rows = connection.execute(
        "SELECT AVG(net_revenue),COUNT(*) FROM vw_sales_analysis WHERE calendar_quarter='Q4'").fetchone()
    non_q4_avg, non_q4_rows = connection.execute(
        "SELECT AVG(net_revenue),COUNT(*) FROM vw_sales_analysis WHERE calendar_quarter!='Q4'").fetchone()
    q4_electronics_share = float(connection.execute(
        "SELECT 100.0*SUM(CASE WHEN category='Electronics' THEN net_revenue ELSE 0 END)/SUM(net_revenue) FROM vw_sales_analysis WHERE calendar_quarter='Q4'").fetchone()[0])
    non_q4_electronics_share = float(connection.execute(
        "SELECT 100.0*SUM(CASE WHEN category='Electronics' THEN net_revenue ELSE 0 END)/SUM(net_revenue) FROM vw_sales_analysis WHERE calendar_quarter!='Q4'").fetchone()[0])
    region_line_counts = {str(region): int(n) for region, n in connection.execute(
        "SELECT region,COUNT(*) FROM vw_sales_analysis GROUP BY region")}
    correlation_data = np.asarray(connection.execute(
        "SELECT discount_pct,margin_pct FROM vw_sales_analysis ORDER BY sales_line_id").fetchall(), dtype=float)
    discount_margin_corr = round(float(np.corrcoef(correlation_data[:, 0], correlation_data[:, 1])[0, 1]), 3)
    return {
        "project": "Retail Sales Performance Analysis",
        "seed": SEED,
        "schema_source": "docs/08_Database_Design.xlsx | SQLite_DDL",
        "validation_status": "PASS" if passed else "FAIL",
        "expected_row_counts": expected,
        "actual_database_row_counts": counts,
        "foreign_key_check": {
            "status": "PASS" if not fk_violations and fk_enforcement_enabled else "FAIL",
            "enforcement_enabled_on_generator_connection": fk_enforcement_enabled,
            "violations": [list(x) for x in fk_violations],
        },
        "range_and_key_checks": range_checks,
        "joined_view": {"name": "vw_sales_analysis", "rows": view_count, "column_count": len(view_columns), "columns": view_columns},
        "raw_quality_issues": {
            "missing_payment_method_cells": missing_count,
            "missing_rate_by_sales_rows_pct": round(100 * missing_count / expected["fact_sales"], 2),
            "missing_sales_line_ids": QUALITY_MISSING_IDS,
            "imputation_for_database": {"field": "payment_method", "method": "mode of observed raw payment_method values", "value": imputed_mode},
            "duplicate_business_groups": duplicate_groups,
            "duplicate_group_count": len(duplicate_groups),
            "valid_range_statistical_outliers": outliers,
        },
        "patterns_intentionally_built_in": [
            "2024 sales use a modest year-over-year quantity and price uplift.",
            "November and December have higher unit volumes and a larger Electronics share.",
            "North and West store sales have higher propensity than other regions.",
            "Discounts are more frequent in March, July, and August; greater discounts mechanically reduce calculated gross margin.",
        ],
        "observed_pattern_metrics": {
            "net_revenue_by_year_inr": {year: round(total, 2) for year, total in year_totals.items()},
            "net_revenue_yoy_change_pct": yoy_change_pct,
            "q4_average_net_revenue_per_line_inr": round(float(q4_avg), 2),
            "non_q4_average_net_revenue_per_line_inr": round(float(non_q4_avg), 2),
            "q4_sales_line_count": int(q4_rows),
            "non_q4_sales_line_count": int(non_q4_rows),
            "q4_electronics_share_of_net_revenue_pct": round(q4_electronics_share, 1),
            "non_q4_electronics_share_of_net_revenue_pct": round(non_q4_electronics_share, 1),
            "sales_line_count_by_region": region_line_counts,
            "discount_vs_margin_pearson_correlation": discount_margin_corr,
            "interpretation_note": "Generated signals are descriptive; 240 lines and three intentional price outliers do not support causal claims.",
        },
        "notes": [
            "CSV and Excel contain raw synthetic rows; six payment_method cells are blank for cleaning practice.",
            "SQLite uses the exact Step 8 NOT NULL schema and imputes only those six blank payment_method values using the observed mode.",
            "Three duplicate business lines retain distinct sales_line_id primary keys; duplicates are assessed across all non-key sales fields.",
            "Three high unit_price_at_sale values exceed four times list price. The Step 8 DDL has no upper price cap, so these remain valid-range statistical outliers.",
            "All people-like names and place names are fictional labels; no customer names, IDs, or real personal data are generated.",
        ],
    }


def write_report(report: dict[str, object]) -> None:
    REPORT_JSON.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    lines = [
        "RETAIL SALES PERFORMANCE ANALYSIS - DATA VALIDATION REPORT",
        f"Seed: {report['seed']}",
        f"Schema source: {report['schema_source']}",
        f"Overall status: {report['validation_status']}",
        "",
        "SQLite row counts (expected / actual):",
    ]
    for table, expected in report["expected_row_counts"].items():
        lines.append(f"  {table}: {expected} / {report['actual_database_row_counts'][table]}")
    lines.extend(["", f"Foreign key check: {report['foreign_key_check']['status']} (enforcement enabled during load: {report['foreign_key_check']['enforcement_enabled_on_generator_connection']})", "Range, allowed-value, NOT NULL, and key checks:"])
    for name, passed in report["range_and_key_checks"].items():
        lines.append(f"  {'PASS' if passed else 'FAIL'} - {name}")
    view = report["joined_view"]
    lines.extend(["", f"Joined view: {view['name']} | {view['rows']} rows | {view['column_count']} columns"])
    quality = report["raw_quality_issues"]
    lines.extend([
        "", "Intentional raw-data quality issues:",
        f"  Missing payment_method: {quality['missing_payment_method_cells']} / 240 rows ({quality['missing_rate_by_sales_rows_pct']}%); IDs {quality['missing_sales_line_ids']}",
        f"  Database handling: imputed to mode '{quality['imputation_for_database']['value']}' to satisfy Step 8 NOT NULL and CHECK constraints.",
        f"  Duplicate business groups: {quality['duplicate_group_count']} -> {quality['duplicate_business_groups']}",
        f"  Statistical outliers (unit price / list price): {quality['valid_range_statistical_outliers']}",
        "", "Patterns built into the synthetic data:",
    ])
    for pattern in report["patterns_intentionally_built_in"]:
        lines.append(f"  - {pattern}")
    lines.extend(["", "Observed pattern metrics (for inspection, not causal inference):"])
    for key, value in report["observed_pattern_metrics"].items():
        lines.append(f"  {key}: {value}")
    lines.extend(["", "Interpretation: raw CSV/XLSX are cleaning-practice inputs; app.db is loaded with the exact constrained Step 8 schema."])
    REPORT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ddl = read_schema_ddl()
    products = build_products()
    stores = build_stores()
    sales_raw = build_sales(products, stores)
    targets = build_targets(sales_raw, products, stores)
    raw_tables = {
        "dim_product": products,
        "dim_store": stores,
        "fact_sales": raw_sales_with_missing(sales_raw),
        "fact_sales_target": targets,
    }
    clean_sales, imputed_mode = impute_missing_for_database(sales_raw)
    db_tables = {**raw_tables, "fact_sales": clean_sales}

    for table_name, rows in raw_tables.items():
        write_csv(table_name, rows)
    write_excel(raw_tables)
    connection, _ = load_database(ddl, db_tables)
    report = validate(connection, raw_tables, db_tables, imputed_mode)
    write_report(report)
    connection.close()
    print(f"Validation: {report['validation_status']}")
    print(f"Workbook: {EXCEL_PATH}")
    print(f"SQLite: {DB_PATH}")
    for table_name in TABLE_COLUMNS:
        print(f"CSV: {DATA_DIR / (table_name + '.csv')} ({report['actual_database_row_counts'][table_name]} rows in DB)")
    print(f"Validation report: {REPORT_TXT}")
    print(f"Joined view: {report['joined_view']['column_count']} columns")
    if report["validation_status"] != "PASS":
        raise SystemExit("Generated dataset did not pass all Step 10 checks.")


if __name__ == "__main__":
    main()
