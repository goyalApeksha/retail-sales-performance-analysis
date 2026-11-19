"""Database and generation quality checks for the synthetic retail dataset."""
from pathlib import Path
import sqlite3

import pytest

from src.generate_data import impute_missing_for_database

ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "app.db"
EXPECTED = {
    "dim_product": (20, 8),
    "dim_store": (10, 6),
    "fact_sales": (240, 12),
    "fact_sales_target": (4800, 4),
    "vw_sales_analysis": (240, 35),
}


@pytest.fixture
def db():
    assert DB_PATH.exists(), f"Synthetic SQLite database is missing: {DB_PATH}"
    conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def test_app_database_opens_read_only_and_required_objects_load(db):
    names = {row[0] for row in db.execute(
        "SELECT name FROM sqlite_master WHERE type IN ('table','view')"
    )}
    assert set(EXPECTED).issubset(names)
    assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"


@pytest.mark.parametrize("table,expected", EXPECTED.items())
def test_table_or_view_row_and_column_counts(db, table, expected):
    rows = db.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
    columns = len(db.execute(f'PRAGMA table_info("{table}")').fetchall())
    assert (rows, columns) == expected


@pytest.mark.parametrize("table", ["dim_product", "dim_store", "fact_sales", "fact_sales_target"])
def test_primary_keys_are_unique_and_non_null(db, table):
    info = db.execute(f'PRAGMA table_info("{table}")').fetchall()
    pk = [row[1] for row in sorted((r for r in info if r[5]), key=lambda r: r[5])]
    assert pk, f"{table} must declare a primary key"
    columns = ", ".join(f'"{name}"' for name in pk)
    duplicate = db.execute(
        f'SELECT {columns}, COUNT(*) AS n FROM "{table}" GROUP BY {columns} HAVING n > 1 LIMIT 1'
    ).fetchone()
    assert duplicate is None
    null_predicate = " OR ".join(f'"{name}" IS NULL' for name in pk)
    assert db.execute(f'SELECT COUNT(*) FROM "{table}" WHERE {null_predicate}').fetchone()[0] == 0


def test_all_foreign_keys_are_valid(db):
    assert db.execute("PRAGMA foreign_key_check").fetchall() == []


def test_database_cleaning_imputes_missing_payment_without_losing_rows():
    rows = [
        {"sales_line_id": 1, "payment_method": "UPI"},
        {"sales_line_id": 2, "payment_method": None},
        {"sales_line_id": 3, "payment_method": "Card"},
        {"sales_line_id": 4, "payment_method": "UPI"},
    ]
    cleaned, mode = impute_missing_for_database(rows)
    assert mode == "UPI"
    assert len(cleaned) == len(rows)
    assert all(row["payment_method"] is not None for row in cleaned)
    assert rows[1]["payment_method"] is None  # input was not mutated


def test_analysis_view_does_not_multiply_sales_lines(db):
    base = db.execute("SELECT COUNT(*) FROM fact_sales").fetchone()[0]
    joined = db.execute("SELECT COUNT(*) FROM vw_sales_analysis").fetchone()[0]
    distinct_lines = db.execute("SELECT COUNT(DISTINCT sales_line_id) FROM vw_sales_analysis").fetchone()[0]
    assert joined == base == distinct_lines


def test_missing_database_is_reported_cleanly(tmp_path):
    # Step 15 contract: src.db.open_readonly_connection raises FileNotFoundError
    # before attempting SQLite to avoid creating an empty replacement database.
    from src.db import open_readonly_connection
    missing = tmp_path / "missing.db"
    with pytest.raises(FileNotFoundError):
        open_readonly_connection(missing)


def test_validation_detects_required_quality_issue_types():
    # Step 15 contract: return a list of issue dictionaries with a stable code.
    import pandas as pd
    from src.validation import validate_sales_frame
    frame = pd.DataFrame({
        "sales_line_id": [1, 1], "sale_date": ["not-a-date", "2023-01-01"],
        "quantity": [0, 1], "discount_pct": [0.9, 0.1],
        "product_id": [None, "P-1"], "store_id": ["S-1", "S-1"],
        "payment_method": ["UPI", "Card"],
    })
    codes = {issue["code"] for issue in validate_sales_frame(frame)}
    assert {"DUPLICATE_KEY", "INVALID_DATE", "INVALID_QUANTITY",
            "INVALID_DISCOUNT", "MISSING_REQUIRED"}.issubset(codes)


def test_filtered_csv_serialization_round_trips_only_selected_synthetic_rows():
    # Step 15 contract: serialize_filtered_csv accepts the already-filtered frame
    # and returns UTF-8 CSV text for st.download_button.
    import pandas as pd
    from src.data_loader import serialize_filtered_csv
    filtered = pd.DataFrame({
        "sales_line_id": [10, 11], "region": ["North", "North"],
        "net_revenue": [125.0, 70.0],
    })
    csv_text = serialize_filtered_csv(filtered)
    round_trip = pd.read_csv(__import__("io").StringIO(csv_text))
    assert list(round_trip.columns) == list(filtered.columns)
    assert round_trip["sales_line_id"].tolist() == [10, 11]
    assert round_trip["region"].eq("North").all()
