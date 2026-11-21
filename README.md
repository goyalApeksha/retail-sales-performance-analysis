# Retail Sales Performance Analysis (Project ID: P048)

ABES MCA 2nd Year Data Analytics project by **Apeksha Goel** (Roll No: 2500320140030), mentored by **Mr. Dinesh**.  
**Programme / Batch:** MCA 2nd Year Data Analytics — Batch `ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01` (Academic Session: 2025–2027)  
**Live Deployed App (Streamlit Community Cloud):** [https://retail-sales-performance-analysis.streamlit.app/Performance_Compare](https://retail-sales-performance-analysis.streamlit.app/Performance_Compare)  
**GitHub Repository:** [https://github.com/goyalApeksha/retail-sales-performance-analysis.git](https://github.com/goyalApeksha/retail-sales-performance-analysis.git)  

The project uses a reproducible **synthetic-only** retail-sales dataset. It contains no real customer or personal data.

## What the app answers

The Streamlit app addresses OQ-01–OQ-10 with shared date, region, and category filters:

- **Overview and KPIs:** net revenue, average order value, aggregate profit margin, and line-based return rate.
- **Sales Trends:** monthly revenue and volume, MoM/YoY growth, and calendar-month seasonality.
- **Performance Compare:** actual-versus-target by store, product, or region; AOV by region and customer segment.
- **Product and Category EDA:** revenue ranking, category margin, Pareto contribution, and sales velocity.
- **Insights and Statistics:** payment method counts/AOV, discount-margin association, and return rate by category.
- **Data and Download:** filtered, bounded data preview and synthetic CSV download.
- **Data Quality:** generator validation, raw-input issues, and cleaning disclosure.
- **About:** project purpose, definitions, and limitations.

Correlation describes association in this synthetic sample and does not establish causation. Optional statistical tests must disclose sample limitations.

## Technology

Python 3.12, Streamlit, pandas, NumPy, SciPy, statsmodels, Plotly, SQLAlchemy, SQLite, pytest, and Streamlit AppTest. Tailwind CDN styling is limited to the authored iframe header and KPI fragments; the Streamlit UI keeps a fallback.

## Setup

Install Python 3.12, then run from the repository root in PowerShell:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

The app resolves `data/app.db` relative to the project files and opens it read-only. No database server or environment secrets are required.

## Run the app

```powershell
streamlit run app.py
```

Open the local URL shown by Streamlit. The shared filters apply date range, region, and product category to all analytical screens. Use **Reset filters** in the sidebar to restore the full data range and all categories/regions.

## Run the tests

```powershell
pytest -q
```

The tests cover SQLite object/count baselines, key and foreign-key integrity, cleaning, hand-calculated KPI/target logic, monthly growth guards, correlation validity, app startup, page rendering, download availability, and the AppTest startup target. Manual functional cases TC-F-01–TC-F-24 are documented in [docs/13_Test_Cases.docx](docs/13_Test_Cases.docx); capture execution evidence before marking them Pass.

## Data and metric definitions

The database contains four normalized tables: `dim_product` (20 rows), `dim_store` (10), `fact_sales` (240), and `fact_sales_target` (4,800), plus the 35-column `vw_sales_analysis` view (240 sales lines). Generated files and quality evidence are in `data/`; reproducible generation code is in `src/generate_data.py`.

- **Net revenue:** quantity × sale-time unit price × (1 − discount rate), rounded per line, then summed.
- **AOV:** eligible net order revenue divided by distinct eligible order IDs.
- **Profit margin:** (aggregate net revenue − aggregate line cost) ÷ aggregate net revenue × 100.
- **Target variance:** actual net revenue minus target at matching month/store/product grain.
- **MoM growth:** (current month − prior month) ÷ prior month × 100; unavailable when prior revenue is zero.
- **Return rate:** returned eligible sales lines divided by eligible sales lines × 100.
- **Sales velocity:** selected units divided by calendar days in the selected date range.

Monthly target values repeat on sales lines in the analysis view. Target comparisons therefore aggregate `fact_sales_target` independently; do not sum repeated target values from `vw_sales_analysis`.

## Known data-volume variance

Step 1 success criterion SC-03 states at least 1,000 transaction rows. The later Step 8 design planned 150–300 main-fact rows, and Step 10 generated 240 `fact_sales` rows using seed 42. This implementation follows the later 240-row baseline. Resolve this variance with the project owner before changing the generator or row-count tests; update the validation report and documentation together if the baseline changes.

## Project references

- [Build plan](BUILD_PLAN.md)
- [Traceability matrix](docs/04_Traceability_Matrix.xlsx)
- [BRD](docs/05_BRD.docx)
- [TRD](docs/07_TRD.docx)
- [Database design](docs/08_Database_Design.xlsx)
- [User stories](docs/11_User_Stories.docx)
- [UI framework](docs/12_UI_Framework.docx)
- [Test cases](docs/13_Test_Cases.docx)

## Deployment

Deploy the repository's `app.py` entrypoint from GitHub to Streamlit Community Cloud with Python 3.12 and the pinned `requirements.txt`. Include the synthetic `data/app.db`, theme config, and all source files. Confirm the live URL, all eight screens, filters, database read, charts, CSV download, and empty/error states after deployment; record only verified results.
