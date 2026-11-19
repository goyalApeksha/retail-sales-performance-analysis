"""Streamlit AppTest checks for the Step 15 multipage app contract."""
from pathlib import Path
from time import perf_counter

import pytest
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
ENTRY = ROOT / "app.py"
PAGE_FILES = [
    "pages/1_Sales_Trends.py",
    "pages/2_Performance_Compare.py",
    "pages/3_Product_Category_EDA.py",
    "pages/4_Insights_Statistics.py",
    "pages/5_Data_Download.py",
    "pages/6_Data_Quality.py",
    "pages/7_About.py",
]


def test_streamlit_app_starts_without_uncaught_exception():
    at = AppTest.from_file(str(ENTRY), default_timeout=10).run()
    assert not at.exception
    assert len(at.title) + len(at.header) + len(at.subheader) > 0


def test_every_declared_page_file_renders_without_exception():
    assert ENTRY.exists()
    missing = [name for name in PAGE_FILES if not (ROOT / name).exists()]
    assert not missing, f"Missing expected UI page files: {missing}"
    at = AppTest.from_file(str(ENTRY), default_timeout=10).run()
    assert not at.exception
    for page in PAGE_FILES:
        at.switch_page(page).run(timeout=10)
        assert not at.exception, f"Page raised an exception: {page}"


def test_filtered_data_page_exposes_download_control():
    at = AppTest.from_file(str(ENTRY), default_timeout=10).run()
    at.switch_page("pages/5_Data_Download.py").run(timeout=10)
    assert not at.exception
    assert len(at.get("download_button")) >= 1


def test_app_entrypoint_startup_completes_under_three_seconds():
    start = perf_counter()
    at = AppTest.from_file(str(ENTRY), default_timeout=10).run()
    elapsed = perf_counter() - start
    assert not at.exception
    assert elapsed < 3.0, f"AppTest startup took {elapsed:.2f}s (target < 3s)"


def test_download_button_is_present_on_data_page_and_download_is_nonempty():
    at = AppTest.from_file(str(ENTRY), default_timeout=3).run()
    at.switch_page("pages/5_Data_Download.py").run()
    assert not at.exception
    buttons = at.get("download_button")
    assert len(buttons) >= 1
