# Retail Sales Performance Analysis — Step 14 Build Plan

**AI build input for Step 15** | Owner: Apeksha Goel, 2500320140030, MCA 2nd Year DA, ABES | Mentor: Kalpana Dwivedi | Version 1.0 | 18 November 2026

## 1. Project summary (10 lines)
1. ABES MCA Data Analytics academic project for Retail Sales Performance Analysis, using synthetic retail data only.
2. Replaces scattered manual reports with a consistent analytical view of product, store, target, margin, seasonality, payment, and returns.
3. Primary users are Store Manager SH-01 and Regional Sales Head SH-02; SH-03–SH-07 are also represented.
4. Answers OQ-01–OQ-10 using the approved dataset and shared date, region, and category filters.
5. Business requirements are BR-01–BR-13; functional features FR-01–FR-13; preserve every established ID.
6. Stack: Python 3.12, Streamlit, pandas, NumPy, SciPy/statsmodels, Plotly, SQLAlchemy, SQLite, pytest, GitHub, Streamlit Community Cloud.
7. Data model: dim_product, dim_store, fact_sales, fact_sales_target, plus 35-column vw_sales_analysis.
8. UI comprises UI-01–UI-08; Step 13 defines 24 functional TC-F and 18 technical TC-T cases.
9. Analysis is descriptive; do not state that correlation proves causation. Do not introduce real personal data.
10. Step 15 follows one BUILD task at a time, validates it, maintains the RTM, then prepares and verifies deployment.

## 2. Source documents and project artifacts

Use the following approved items as the source of truth. Step 6 was intentionally skipped by the owner; Step 7 is the TRD. Duplicate Step 1–3 documents are retained copies.
- [Step 1 Project Understanding](docs/01_Project_Understanding_RetailSales_ApekshaGoel.docx)
- [Step 1 duplicate copy](docs/01_name_project_understa.docx)
- [Step 2 HMW Statements](docs/02_HMW_Statements.docx)
- [Step 2 duplicate copy](docs/02_HMW_Statements_RetailSales_ApekshaGoel.docx)
- [Step 3 Needs Analysis](docs/03_Needs_Analysis.docx)
- [Step 3 duplicate copy](docs/03_Needs_Analysis_RetailSales_ApekshaGoel.docx)
- [Step 4 RTM](docs/04_Traceability_Matrix.xlsx)
- [Step 5 BRD](docs/05_BRD.docx)
- [Step 7 TRD](docs/07_TRD.docx)
- [Step 8 Database Design and RTM](docs/08_Database_Design.xlsx)
- [Step 9 ER Diagram document](docs/09_ER_Diagram.docx)
- [Step 9 ER Diagram image](docs/09_ER_Diagram.png)
- [Step 10 generator](src/generate_data.py)
- [Step 10 SQLite database](data/app.db)
- [Step 10 synthetic data workbook](data/retail_sales_performance_analysis_data.xlsx)
- [Step 10 validation report text](data/validation_report.txt)
- [Step 10 validation report JSON](data/validation_report.json)
- [Step 11 User Stories](docs/11_User_Stories.docx)
- [Step 12 UI Framework](docs/12_UI_Framework.docx)
- [Step 13 Test Cases](docs/13_Test_Cases.docx)
- [Current UI helper](src/ui.py)
- [Streamlit theme](.streamlit/config.toml)

### Traceability, assumptions, and change control

- Stable identifiers: OQ-01–OQ-10, SH-01–SH-07, HMW-01–HMW-12, BR-01–BR-13, FR-01–FR-13, US-01–US-16, UI-01–UI-08, TC-F-01–TC-F-24 and TC-T-01–TC-T-18. Never renumber. Add new IDs only at sequence end and update the RTM.
- Canonical RTM: `docs/04_Traceability_Matrix.xlsx`. Keep `docs/08_Database_Design.xlsx` / `Updated_RTM` synchronized. All 18 RTM rows currently contain UI and test IDs; all 13 BRs have functional coverage.
- **Known data-volume variance:** Step 1 SC-03 states ≥1,000 fact rows; Step 8 planned 150–300, and Step 10/current `fact_sales` has 240 rows from seed 42. Default Step 15 to this later approved 240-row baseline unless Apeksha approves a change. Record the variance; if changed, update generator, test expectations, validation report and documents without renumbering.
- `vw_sales_analysis` is at sales-line grain and repeats `monthly_target`. Compare targets by separately aggregating `fact_sales_target` at target-month/store/product grain.
- OQ-08 is descriptive by payment method; optional Kruskal-Wallis must disclose limitations. OQ-09 correlation is association, not causation. Return rate uses eligible sales lines as M-15 specifies.

## 3. Final folder structure

```text
retail-sales-performance-analysis/
├── app.py
├── pages/ (1_Sales_Trends.py, 2_Performance_Compare.py, 3_Product_Category_EDA.py,
│          4_Insights_Statistics.py, 5_Data_Download.py, 6_Data_Quality.py, 7_About.py)
├── src/ (__init__.py, db.py, data_loader.py, validation.py, analysis.py,
│        statistics.py, charts.py, ui.py, generate_data.py)
├── data/ (app.db, dimension/fact CSVs, synthetic workbook, validation reports)
├── tests/ (test_data.py, test_analysis.py, test_app.py)
├── docs/ (approved Steps 1–14 and final ABES deliverables)
├── .streamlit/config.toml
├── requirements.txt
├── README.md
└── .gitignore
```

Existing files include generator, UI helper, synthetic SQLite/data exports, theme, test files and project deliverables. BUILD-01–BUILD-16 fill the remaining modules and deployment artifacts. Keep secrets, venvs, caches and temporary renders out of Git.

## 4. Build tasks — one at a time

Complete each task and its check before continuing. Report files changed, checks/results, and unresolved issues after each task.

### BUILD-01 — Environment and pinned dependencies
- **Goal:** Set up Python 3.12, package pins, repository hygiene, and the base configuration.
- **Files:** `requirements.txt; .gitignore; src/__init__.py; .streamlit/config.toml; README.md (initial setup)`
- **Stories / features / BRs:** Foundation for all FRs and US-01–US-16.
- **Exact AI prompt:** “Create the Python 3.12 environment and requirements.txt using the exact Step 7 pins below. Add src/__init__.py and .gitignore entries for secrets, virtual environments, caches and temporary files, but keep approved synthetic app.db/data artifacts. Preserve the existing Streamlit theme. Add initial install/run instructions. Check availability/compatibility of pins; if any must change, record why and update the TRD before proceeding. Do not change scope or IDs.”
- **How to check:** Install in a clean Python 3.12 environment; import all pinned libraries; inspect tracked files for secrets and ensure app.db remains included.
- **Tests that must pass:** Environment smoke check; TC-T-01 after BUILD-02.

### BUILD-02 — Database loader, validation, and cleaning
- **Goal:** Implement safe SQLite access, analysis loading, and documented data validation/cleaning.
- **Files:** `src/db.py; src/data_loader.py; src/validation.py; tests/test_data.py.`
- **Stories / features / BRs:** US-01, US-02, US-04; FR-02, FR-12; BR-01, BR-04, BR-13.
- **Exact AI prompt:** “Implement project-root-relative read-only SQLite access, context-managed connections, parameterized queries, and data loading for the approved schema. Validate schema, types, dates, required values, keys, foreign keys, ranges and coverage before metrics. Apply the documented deterministic missing payment_method cleaning without mutating source files. Provide serialize_filtered_csv for an already-filtered synthetic frame. Preserve the sales-line grain. Never sum repeated monthly_target values from the view; aggregate targets from fact_sales_target independently. Satisfy tests/test_data.py.”
- **How to check:** Compare table/view row and column counts to Step 10, verify unique PKs and PRAGMA foreign_key_check, reconcile 240 view rows to 240 distinct sales_line_id, and compare validation with data/validation_report.json. Verify missing DB produces a clear error.
- **Tests that must pass:** TC-T-01–TC-T-06, TC-T-15, TC-T-16, TC-T-18; TC-F-02, TC-F-21–TC-F-23.

### BUILD-03 — Analysis and statistics functions
- **Goal:** Build calculation functions independently from the UI.
- **Files:** `src/analysis.py; src/statistics.py; tests/test_analysis.py.`
- **Stories / features / BRs:** US-01, US-03, US-05–US-16; FR-01–FR-11, FR-13; BR-01–BR-12.
- **Exact AI prompt:** “Implement calculate_kpis(frame), monthly_revenue_growth(frame), compare_actual_to_target(sales, targets), and discount_margin_correlation(frame, min_n=5), matching the exact return contracts asserted by tests/test_analysis.py. Add reusable functions for product ranking, aggregate margin, grouped AOV, monthly quantity, Pareto, payment summaries, returns, and units-per-calendar-day velocity. Follow approved M-01–M-15 formulas: aggregate ratios after summation, distinct order denominator for AOV, independent target aggregation at month/store/product grain, safe zero denominators, and correlation validity/sample reporting. Do not claim causation.”
- **How to check:** Hand-calculate the small fixtures; verify sorting, chronological trends, target matching grain, margins from sums, distinct-order AOV, return denominator, valid/invalid correlations, and empty/zero guards.
- **Tests that must pass:** TC-T-07–TC-T-11, TC-T-17; TC-F-08–TC-F-20.

### BUILD-04 — Shared UI components and filter state
- **Goal:** Build reusable components, filter handling, chart styling, and responsive shell.
- **Files:** `src/ui.py; src/charts.py; app.py shared helpers; tests/test_app.py.`
- **Stories / features / BRs:** US-02, US-03, US-04; FR-11, FR-12; UI-01–UI-08; BR-10, BR-13.
- **Exact AI prompt:** “Implement shared filters using session keys filter_date_range, filter_regions and filter_categories, initialized from actual SQLite values with a reset-to-default action. Provide one filter function used on UI-01–UI-07; keep screen-specific controls separate. Add reusable KPI, chart, insight, filter summary, empty and error components. Use Streamlit/Plotly for analytical data; restrict Tailwind CDN iframe to static header/KPI markup, HTML-escape all dynamic text and provide a native fallback. Use accessible labels, consistent INR, and mobile-friendly layouts.”
- **How to check:** Verify defaults, combined filters, reset, state persistence across screens, no stale totals, empty states, readable narrow layout and CDN failure fallback.
- **Tests that must pass:** TC-F-01–TC-F-07, TC-F-24; TC-T-12, TC-T-13.

### BUILD-05 — UI-01 Overview and KPIs
- **Goal:** Create the landing summary with the four agreed KPIs.
- **Files:** `app.py; src/analysis.py; src/ui.py; src/charts.py.`
- **Stories / features / BRs:** US-03, US-04; FR-11, FR-12; UI-01; BR-10, BR-13.
- **Exact AI prompt:** “Build UI-01 Overview with net revenue, AOV, aggregate profit margin %, and line-based return rate. Show active period/filters and a concise trend/target context with a clear empty state. Use the tested shared KPI function and same filtered population for all cards. Do not add unapproved KPI definitions.”
- **How to check:** Reconcile all four card values to filtered rows and hand calculation; change each shared filter and confirm all cards refresh; review empty and zero-denominator outputs.
- **Tests that must pass:** TC-F-03–TC-F-08, TC-F-24; TC-T-07, TC-T-12, TC-T-14, TC-T-17.

### BUILD-06 — UI-02 Sales Trends
- **Goal:** Show monthly revenue growth and sales-volume seasonality.
- **Files:** `pages/1_Sales_Trends.py; src/analysis.py; src/charts.py.`
- **Stories / features / BRs:** US-04, US-06, US-09; FR-03, FR-06, FR-12; UI-02; BR-02, BR-09, BR-13.
- **Exact AI prompt:** “Build UI-02 with monthly net-revenue line and period comparison, MoM/YoY growth, monthly quantity and descriptive calendar-month pattern. Default the rolling 12-month display to the latest 12 months available in the dataset, not today. Label periods and handle zero prior-month revenue as unavailable. Apply shared filters.”
- **How to check:** Verify chronological ordering, period labels, values and date/region/category filtering; exercise zero-prior-month; check mobile layout.
- **Tests that must pass:** TC-F-03–TC-F-07, TC-F-10, TC-F-13, TC-F-24; TC-T-09, TC-T-13.

### BUILD-07 — UI-03 Performance Compare
- **Goal:** Compare targets and AOV at valid aggregation levels.
- **Files:** `pages/2_Performance_Compare.py; src/analysis.py; src/charts.py.`
- **Stories / features / BRs:** US-01, US-04, US-08, US-14; FR-02, FR-05, FR-12; UI-03; BR-01, BR-04, BR-05, BR-13.
- **Exact AI prompt:** “Build actual-vs-target at month/store/product grain; aggregate actuals and fact_sales_target separately, show actual, target, absolute/percentage variance, and guard zero targets. Add AOV by region/customer segment using distinct orders and expose order counts. Apply shared filters.”
- **How to check:** Use the fixture actual=250, target=200, variance=50, variance%=25%; ensure multi-line joins do not multiply either side; verify grouped AOV denominator.
- **Tests that must pass:** TC-F-03–TC-F-07, TC-F-12, TC-F-19, TC-F-24; TC-T-08, TC-T-13, TC-T-17.

### BUILD-08 — UI-04 Product and Category EDA
- **Goal:** Answer product rank, margin, Pareto and velocity questions.
- **Files:** `pages/3_Product_Category_EDA.py; src/analysis.py; src/charts.py.`
- **Stories / features / BRs:** US-04, US-05, US-07, US-10, US-15; FR-01, FR-04, FR-07, FR-12, FR-13; UI-04; BR-03, BR-06, BR-12, BR-13.
- **Exact AI prompt:** “Build descending product net-revenue ranking, category aggregate profit margin, Pareto bars plus cumulative share/80% threshold, and product/category units per calendar day in selected period. Include units, sort, period, filters, denominators and safe empty states.”
- **How to check:** Reconcile totals, ensure margin is ratio of sums, cumulative share is monotonic and velocity uses selected calendar-day denominator.
- **Tests that must pass:** TC-F-03–TC-F-06, TC-F-09, TC-F-11, TC-F-14, TC-F-20, TC-F-24; TC-T-07, TC-T-13.

### BUILD-09 — UI-05 Insights and Statistics
- **Goal:** Present payment, discount/margin, and return analysis with valid statistical evidence.
- **Files:** `pages/4_Insights_Statistics.py; src/analysis.py; src/statistics.py; src/charts.py.`
- **Stories / features / BRs:** US-04, US-11–US-13, US-16; FR-08–FR-10, FR-12; UI-05; BR-07, BR-08, BR-11, BR-13.
- **Exact AI prompt:** “Build separate payment count and AOV comparisons, category discount-vs-margin scatter/correlation with eligible sample count and validity reason, and category returns with numerator/denominator. Suppress invalid Pearson output; describe association without causation. Optional Kruskal-Wallis is secondary and must state sample limitations.”
- **How to check:** Reconcile counts/AOV/rates, match plot and correlation eligible records, verify invalid results withhold coefficient and p-value.
- **Tests that must pass:** TC-F-03–TC-F-07, TC-F-15–TC-F-18, TC-F-24; TC-T-10, TC-T-11, TC-T-13, TC-T-17.

### BUILD-10 — UI-07 Data Quality
- **Goal:** Expose quality checks and cleaning outcomes.
- **Files:** `pages/6_Data_Quality.py; src/validation.py; src/data_loader.py.`
- **Stories / features / BRs:** US-02, US-16; FR-09, FR-12; UI-07; BR-07, BR-13.
- **Exact AI prompt:** “Build UI-07 showing completeness, duplicates, PK/FK, ranges, date coverage and cleaning outcomes. Distinguish raw synthetic input issues from constrained cleaned app.db; show counts, status and treatment. Never rewrite source data from the UI or expose stack traces.”
- **How to check:** Reconcile check totals with Step 10 validation report; inject representative issues and confirm detection and clear messages.
- **Tests that must pass:** TC-F-02, TC-F-07, TC-F-17, TC-F-23; TC-T-04, TC-T-06, TC-T-15, TC-T-16.

### BUILD-11 — UI-06 Data and Download
- **Goal:** Provide a bounded preview and filtered synthetic CSV export.
- **Files:** `pages/5_Data_Download.py; src/data_loader.py; src/ui.py.`
- **Stories / features / BRs:** US-02, US-04; FR-12; UI-06; BR-13.
- **Exact AI prompt:** “Build UI-06 with max 100-row filtered preview, selected date/row summary, synthetic-data notice, data dictionary, and CSV download. Export only an already-filtered frame, with approved columns and clear filename. Never put data into iframe HTML.”
- **How to check:** Reopen export and compare values and filters with source frame; confirm encoding, headers and no personal data.
- **Tests that must pass:** TC-F-02–TC-F-06, TC-F-21, TC-F-22, TC-F-24; TC-T-18.

### BUILD-12 — UI-08 About and navigation
- **Goal:** Complete static context and verify page-to-page flow.
- **Files:** `pages/7_About.py; app.py; pages/*.py.`
- **Stories / features / BRs:** All stories/features in navigation context; UI-01–UI-08.
- **Exact AI prompt:** “Complete UI-08 with project purpose, period, synthetic-data disclosure, KPI definitions, limitations and academic context. Ensure navigation labels and order match UI-01–UI-08 and filters persist on analytical screens. Do not state unverified test or deployment outcomes.”
- **How to check:** Visit every screen from Overview and back; inspect page title, state persistence, empty startup and narrow layout.
- **Tests that must pass:** TC-F-01, TC-F-24; TC-T-12, TC-T-13.

### BUILD-13 — Run full tests, fix defects, record evidence
- **Goal:** Meet quality gates and align the test evidence with implementation.
- **Files:** `tests/*; all app/source files; docs/04_Traceability_Matrix.xlsx; docs/13_Test_Cases.docx.`
- **Stories / features / BRs:** All BR-01–BR-13, FR-01–FR-13, US-01–US-16, UI-01–UI-08.
- **Exact AI prompt:** “Run the test cases named in the RTM; repair root causes rather than weaken acceptance. Capture environment, command, results, defects and retests. Measure AppTest startup against 3 s. Confirm every BR has a test and every manual case has execution evidence. Resolve and document the Step 1 ≥1,000 vs Step 8/10 actual 240 row variance. Preserve the seed-42 240-row baseline unless Apeksha approves changing it; if changed, align generator, counts, validation, test expectations and deliverables without renumbering IDs.”
- **How to check:** Run pytest -q; target all 18 technical tests passing and all 24 functional scenarios executed. Inspect FK checks and test report; no critical/high open defects.
- **Tests that must pass:** TC-F-01–TC-F-24; TC-T-01–TC-T-18.

### BUILD-14 — Write reproducible README
- **Goal:** Make setup, usage, analysis limits and troubleshooting clear to learner/reviewer.
- **Files:** `README.md; requirements.txt; data/validation_report.txt; docs links.`
- **Stories / features / BRs:** All OQs and FRs; reviewer stakeholders SH-06, SH-07.
- **Exact AI prompt:** “Write a README with purpose, synthetic-only statement, Python version, environment setup, install/run/test commands, pages/filters, metric formulas, data/schema summary, known limitations and row-count decision, troubleshooting, deployment notes and links to deliverables. Do not claim unrun tests or deployment.”
- **How to check:** Follow instructions from a clean checkout in Python 3.12 and verify commands/paths.
- **Tests that must pass:** TC-T-01–TC-T-18 and TC-F-01 smoke path.

### BUILD-15 — GitHub and deployment preflight
- **Goal:** Prepare a clean, reproducible deployment branch.
- **Files:** `GitHub branch; requirements.txt; app.py; data/app.db; .streamlit/config.toml; README.md; .gitignore.`
- **Stories / features / BRs:** OBJ-06; SC-02, SC-05; all features.
- **Exact AI prompt:** “Prepare the reviewed repository/branch for Streamlit Community Cloud. Include approved source and synthetic SQLite only; exclude secrets, caches, temporary files and personal data. Confirm app.py entrypoint, Python 3.12, requirements and root-relative DB path work from a clean checkout. Do not deploy before BUILD-13 passes.”
- **How to check:** Inspect repository file list, clean installation, local start, tracked synthetic DB and absence of secrets.
- **Tests that must pass:** TC-T-01, TC-T-12–TC-T-14; TC-F-01, TC-F-22.

### BUILD-16 — Deploy and smoke-test Community Cloud
- **Goal:** Publish verified app and capture real deployment evidence.
- **Files:** `Streamlit Community Cloud app; GitHub branch; README.md; deployment evidence.`
- **Stories / features / BRs:** OBJ-06; SC-01, SC-02, SC-05; all ten OQs.
- **Exact AI prompt:** “Deploy the reviewed branch with app.py and Python 3.12 on Streamlit Community Cloud. Record the actual URL and deployment date. Smoke-test clean startup, SQLite access, navigation, filters, KPI formulas, charts, download, empty/error states, CDN fallback and mobile width. Fix and redeploy failures; never place secrets in source. Update README with observed URL/status only after verification.”
- **How to check:** Open public URL in a fresh browser, complete OQ-01–OQ-10 demonstration and record evidence.
- **Tests that must pass:** TC-F-01–TC-F-24 walkthrough; TC-T-12–TC-T-18 where applicable.

## 5. Definition of Done

- Must BRs BR-01, BR-02, BR-04, BR-06, BR-07, BR-09, BR-10, BR-12 and BR-13 are implemented and traced to accepted cases. Should/Could BR-03, BR-05, BR-08 and BR-11 are implemented for approved full OQ/story scope or have an approved exclusion.
- OQ-01–OQ-10 are demonstrable with synthetic data. KPI formulas, target grain, shared filters, empty states, denominator guards and quality behavior are correct.
- UI-01–UI-08 render and are usable at desktop/narrow widths; navigation and session filters remain consistent.
- `pytest -q` passes all technical cases. All functional cases have an execution result and evidence. Objective OBJ-05 sets ≥90%; release DoD requires all automated tests pass and no open critical/high defect.
- App starts without exception, reads the committed DB, and meets the <3 s AppTest target on a recorded reference environment; `PRAGMA foreign_key_check` is clean.
- The ≥1,000 vs 240 row variance is explicitly resolved and aligned across generator, tests, validation report and final project document.
- Reviewed synthetic-only code is in GitHub; Community Cloud app is publicly reachable and smoke-tested. Actual URL and evidence are recorded.
- README supports clean setup/run/test/troubleshooting. RTM, Step 13 test document, final ABES Word submission and slide deck match actual implementation and do not claim unverified results.

## 6. Commands

Run in PowerShell from the repository root with Python 3.12 available:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
python -m pytest -q
```

If activation is blocked, invoke `.\.venv\Scripts\python.exe -m streamlit run app.py` directly. Resolve the database path from the repository root, not the current working directory.

### Proposed dependency pins from Step 7

```text
streamlit==1.67.0
pandas==3.1.1
numpy==2.5.5
scipy==1.18.1
statsmodels==0.15.0
plotly==7.1.0
SQLAlchemy==2.0.58
pytest==9.2.0
```

These are the TRD baseline proposed for Python 3.12. BUILD-01 checks compatibility; any justified update must be documented in TRD and README before build proceeds.

## 7. Viva notes — 10 likely questions

1. **What problem does this app solve?** It replaces scattered manual reports with a consistent, filterable view of sales, targets, products, trends, margin and returns.
2. **Who are its primary users?** Store Manager SH-01 and Regional Sales Head SH-02; category, operations, finance, analyst and academic reviewers also have defined needs.
3. **Why synthetic data?** It protects privacy, is reproducible, and lets the project demonstrate known patterns without real customer data.
4. **How is net revenue calculated?** Quantity × sale-time unit price × (1 − discount rate), rounded per line and then aggregated.
5. **How is AOV calculated?** Eligible net order revenue divided by distinct eligible order count, not sales-line count.
6. **Why aggregate targets separately?** The target repeats across matching sales lines in the joined view; summing there multiplies the target.
7. **How is profit margin calculated?** (Total net revenue − total line cost) / total net revenue × 100; undefined if revenue is zero.
8. **How is Pareto analyzed?** Sort products by revenue, calculate cumulative revenue share, and identify where it reaches 80%.
9. **What does correlation tell us?** It measures observed association only; it does not establish causation and requires adequate sample size/variation.
10. **How is the row-count conflict handled?** Later Step 8/10 artifacts specify and generate 240 rows; Step 15 preserves that baseline unless the owner approves changing data and aligned tests/docs.

---

End of Step 14. Step 15 starts BUILD-01 and proceeds one task at a time; this plan itself does not start implementation.