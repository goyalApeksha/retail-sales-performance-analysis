# Data Cleaning & Validation Report
**Project:** Retail Sales Performance Analysis  
**Learner:** Apeksha Goel | **Roll No:** 2500320140030 | MCA 2nd Year Data Analytics, ABES Engineering College  
**Generated:** 20 November 2026, 22:13:15  

---

## 1. Executive Summary & Objective
This report details the automated data hygiene and cleaning pipeline executed on the raw transactional sales dataset. The clean dataset eliminates data quality defects, enforces referential integrity across relational dimension tables, standardizes data types, computes derived retail financial metrics, and produces analysis-ready CSV and Excel deliverables.

---

## 2. Generated Clean Artifacts

| Artifact Name | Format | Rows × Cols | Description |
|---|---|---|---|
| [`data/clean_sales_analysis.csv`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/clean_sales_analysis.csv) | CSV | 240 × 37 | Full clean joined analytical dataset with all dimensions, targets & derived KPIs. |
| [`data/clean_sales_analysis_deduplicated.csv`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/clean_sales_analysis_deduplicated.csv) | CSV | 237 × 37 | Deduplicated analytical dataset excluding redundant POS billing retransmissions. |
| [`data/clean_fact_sales.csv`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/clean_fact_sales.csv) | CSV | 240 × 12 | Normalized fact table with mode-imputed payment methods and zero nulls. |
| [`data/clean_fact_sales_deduplicated.csv`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/clean_fact_sales_deduplicated.csv) | CSV | 237 × 12 | Deduplicated normalized fact table. |
| [`data/retail_sales_clean_data.xlsx`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/retail_sales_clean_data.xlsx) | Excel (.xlsx) | 8 Sheets | Formatted institutional workbook with styling, auto-column widths, and audit log. |
| [`data/clean_dataset_summary.json`](file:///c:/Users/apeks/Downloads/retail%20sales%20performance%20analysis/data/clean_dataset_summary.json) | JSON | - | Machine-readable validation and KPI summary. |

---

## 3. Data Hygiene Transformations Applied

### 3.1. Missing Value Imputation (Payment Method)
- **Defect:** 6 transactions (2.5% of sales lines) contained missing (`NULL`) payment methods due to network packet timeouts at POS terminals.
- **Root Cause Analysis:** Deleting these rows would bias store and regional revenue figures downwards by ₹42,890.
- **Remediation:** Mode imputation was applied. The observed mode across the dataset was **`UPI`** (accounting for 42.7% of known transactions).
- **Audit of Imputed Lines:**
  - Sales Line ID **14** (`ORD-2023-01-0014`, South Region) → Imputed to `UPI`
  - Sales Line ID **23** (`ORD-2023-02-0003`, West Region) → Imputed to `UPI`
  - Sales Line ID **70** (`ORD-2023-05-0010`, North Region) → Imputed to `UPI`
  - Sales Line ID **102** (`ORD-2023-08-0002`, East Region) → Imputed to `UPI`
  - Sales Line ID **170** (`ORD-2024-03-0010`, Central Region) → Imputed to `UPI`
  - Sales Line ID **188** (`ORD-2024-05-0008`, North Region) → Imputed to `UPI`
- **Result:** **0 missing values remain (100% complete).**

### 3.2. Duplicate Business Line Resolution
- **Defect:** 3 redundant transaction groups were identified where the exact same customer transaction was transmitted twice by store billing queues:
  - Line **238** duplicates Line **8** (`ORD-2023-01-0008`)
  - Line **239** duplicates Line **79** (`ORD-2023-06-0009`)
  - Line **240** duplicates Line **153** (`ORD-2024-01-0013`)
- **Remediation:** Both versions are preserved for analytical clarity:
  1. `clean_fact_sales.csv` preserves all 240 primary keys for exact relational parity with SQLite.
  2. `clean_fact_sales_deduplicated.csv` purges the 3 redundant retransmissions, yielding 237 unique commercial transactions.

### 3.3. Statistical Outlier Audit
- **Observation:** 3 transactions feature unusually high unit prices or quantities:
  - Line **44**: Unit price ₹3,602.97 vs. list price ₹857.85 (custom commercial bundle)
  - Line **101**: Unit price ₹6,069.55 vs. list price ₹1,264.49 (premium multi-pack)
  - Line **222**: Unit price ₹140,007.34 vs. list price ₹25,455.88 (enterprise B2B bulk purchase)
- **Remediation:** Retained without clipping. All three lines fall within legitimate wholesale transaction boundaries and preserve genuine commercial variance.

---

## 4. Key Business Metrics of the Clean Dataset

| Metric | Clean Dataset (240 Rows) | Deduplicated (237 Rows) | Unit / Definition |
|---|---|---|---|
| **Total Net Revenue** | **₹20,17,836.97** | **₹20,14,386.60** | Gross Sales minus Discounts |
| **Total Gross Revenue** | **₹22,96,527.87** | **₹22,92,703.01** | Quantity × Unit Price at Sale |
| **Total Gross Profit** | **₹8,77,631.56** | **₹8,76,378.18** | Net Revenue minus Cost of Goods Sold |
| **Gross Margin %** | **43.49%** | **43.51%** | (Profit / Net Revenue) × 100 |
| **Average Order Value (AOV)** | **₹8,514.08** | **₹8,499.52** | Total Net Revenue / Unique Orders |
| **Return Rate %** | **7.92%** | **8.02%** | Returned Units / Sold Units |

---

## 5. Verification Checklist
- [x] **Zero NULLs** in primary and mandatory columns (`sales_line_id`, `product_id`, `store_id`, `payment_method`, `quantity`, `net_revenue`).
- [x] **100% Referential Integrity** — All foreign keys match valid keys in `dim_store` (10 stores) and `dim_product` (20 products).
- [x] **Date Standardization** — All transaction dates conform to ISO-8601 (`YYYY-MM-DD`).
- [x] **Mathematical Integrity** — `gross_revenue - discount_amount == net_revenue` and `net_revenue - line_cost == profit_amount` validated across 100% of rows.
