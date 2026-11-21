# PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
## MASTER BUILD PLAN & SPRINT IMPLEMENTATION SPECIFICATION (BUILD_PLAN.md)

**Candidate:** Harshit | **University Roll No:** 2500320140058 | **Project ID:** P054  
**Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  
**Project Mentor:** Dinesh Sereno | Industry Data Engineering Lead  
**Institution:** ABES Engineering College, Ghaziabad  
**Baseline Release:** Version 1.0 (MVP) | **Date:** November 2026  
**Target Environment:** Python 3.12+, Streamlit Multi-Page, SQLite, Tailwind CSS CDN, Plotly Graph Objects, pytest  

---

## 1. Project Executive Summary & Documentation Index

Project P054 is an enterprise-grade retail sales performance analytics web application built to eliminate the 48-to-72-hour reporting latency typically suffered by multi-unit retail operations. By ingesting normalized daily transactional and target quota data from an ACID-compliant SQLite data layer (`data/app.db`), the system computes monetary and percentage variances, identifies underperforming stores and merchandise lines, isolates footfall vs ticket deficits via algebraic waterfall decomposition, models promotional markdown margin erosion, and forecasts month-end revenue quotas using daily run-rate velocity. The user experience combines an in-memory caching architecture (`st.cache_data`) for sub-second query latency with Tailwind CSS components embedded in Streamlit for an executive-ready presentation.

### Complete Documentation Artifacts Index
* [01_Project_Understanding.docx](docs/01_Project_Understanding.docx): Problem statement, 5 personas, 10 Operational Questions (OQ-01..10), SMART objectives, Scope boundaries.
* [02_HMW_Statements.docx](docs/02_HMW_Statements.docx): 12 stakeholder pain points, 11 "How Might We" statements, top 5 prioritized challenges.
* [03_Needs_Analysis.docx](docs/03_Needs_Analysis.docx): 12 Business Requirements (BR-01..12), MoSCoW prioritization (58% Must), data feasibility audit.
* [04_Traceability_Matrix.xlsx](docs/04_Traceability_Matrix.xlsx): Requirements Traceability Matrix (RTM) linking OQ $\rightarrow$ HMW $\rightarrow$ BR $\rightarrow$ FR $\rightarrow$ DB $\rightarrow$ US $\rightarrow$ UI $\rightarrow$ TC.
* [05_BRD.docx](docs/05_BRD.docx): Business Requirements Document, RACI governance matrix, 8 business rules, risk mitigation catalog.
* [07_TRD.docx](docs/07_TRD.docx): Technical Requirements Document, 5-layer pipeline architecture, pinned dependencies, SQLite schema.
* [08_Database_Design.xlsx](docs/08_Database_Design.xlsx): Constellation relational schema, 3 dimensions, 2 fact tables, 40-column analytical view.
* [09_ER_Diagram.docx](docs/09_ER_Diagram.docx): High-res ER diagram graphic (`09_ER_Diagram.png`), Mermaid DDL relationships, design critique.
* [11_User_Stories.docx](docs/11_User_Stories.docx): 15 User Stories across 5 Epics (55 Story Points), Gherkin acceptance criteria, 2-week sprint plan.
* [12_UI_Framework.docx](docs/12_UI_Framework.docx): 7 screen blueprints, ASCII wireframes, responsive breakpoints, color palette, chart rationale.
* [13_Test_Cases.docx](docs/13_Test_Cases.docx): Quality assurance matrix: 16 Functional Tests (TC-F-01..16), 21 Technical Unit Tests (TC-T-01..21).

---

## 2. Final Target Repository Structure

```text
c:\Users\Sieghart\Desktop\HCL Project\
│
├── app.py                           # Application Entry Point & Global Routing
├── requirements.txt                 # Pinned Production Dependencies
├── README.md                        # Project Setup, Architecture & Operation Guide
├── BUILD_PLAN.md                    # Step-by-Step Engineering Execution Blueprint
│
├── .streamlit/
│   └── config.toml                  # Streamlit Theme & Server Configurations
│
├── data/
│   ├── app.db                       # Primary SQLite Relational Database Engine
│   ├── P054_Daily_Sales_data.xlsx   # Master 6-Sheet Relational & View Excel Seed
│   ├── dim_stores.csv               # 15 Fictional Indian Retail Store Masters
│   ├── dim_categories.csv           # 5 Merchandise Category Dimension Masters
│   ├── dim_staff.csv                # 75 Sales Associate Dimension Masters
│   ├── fact_daily_targets.csv       # 300 Budgeted Daily Target Quota Records
│   ├── fact_daily_sales.csv         # 303 Raw Sales Records (with injected issues)
│   └── vw_daily_sales_target_analysis.csv # 300 x 40 Cleaned Denormalized Flat Export
│
├── src/
│   ├── __init__.py
│   ├── generate_data.py             # Reproducible Synthetic Seed Generator (seed=42)
│   ├── data_loader.py               # Ingestion, Null Imputation, Deduplication & Caching
│   ├── analysis.py                  # Variance Metrics, Waterfall, OLS, Forecast & Z-Scores
│   └── ui.py                        # Tailwind CSS Header, KPI Cards & Alert Callouts
│
├── pages/
│   ├── 01_Overview.py               # UI-01: Executive Cockpit & Store Deficit Table
│   ├── 02_Trends_Seasonality.py     # UI-02: Day-of-Week Heatmap & Cumulative Trajectory
│   ├── 03_Compare.py                # UI-03: Category Breakdown & Staff Leaderboard
│   ├── 04_Root_Cause.py             # UI-04: Waterfall Decomposition & Margin Scatter
│   ├── 05_Forecasting.py            # UI-05: Run-Rate Velocity & Month-End Forecaster
│   ├── 06_Data_Governance.py        # UI-06: Data Quality Audit, Z-Score Outliers & Export
│   └── 07_About.py                  # UI-07: Academic Credentials, Architecture & DDL
│
├── tests/
│   ├── __init__.py
│   ├── test_data.py                 # TC-T-01..10: Schema, Keys, Integrity & Cleaning
│   ├── test_analysis.py             # TC-T-11..17: Math Formulas, Waterfall, OLS, Z-Scores
│   └── test_app.py                  # TC-T-18..21: UI Components, Theme & Query Latency
│
└── docs/                            # Formal Academic & Professional Documentation Suite
    ├── 01_Project_Understanding.docx
    ├── 02_HMW_Statements.docx
    ├── 03_Needs_Analysis.docx
    ├── 04_Traceability_Matrix.xlsx
    ├── 05_BRD.docx
    ├── 07_TRD.docx
    ├── 08_Database_Design.xlsx
    ├── 09_ER_Diagram.docx
    ├── 09_ER_Diagram.png
    ├── 11_User_Stories.docx
    ├── 12_UI_Framework.docx
    ├── 13_Test_Cases.docx
    └── 14_BUILD_PLAN.docx
```

---

## 3. Step-by-Step Engineering Build Tasks (BUILD-01 to BUILD-12)

### Task BUILD-01: Environment Setup & Requirements Pinning
* **Goal:** Create production-pinned `requirements.txt` and install required dependencies.
* **Files:** `requirements.txt`
* **Linked Traceability:** FR-01 | Non-Functional Requirements (TBD Section 8)
* **Action Prompt for AI/Developer:**
  > "Create `requirements.txt` pinning exact versions of `streamlit`, `pandas`, `numpy`, `scipy`, `statsmodels`, `plotly`, `openpyxl`, and `pytest`. Ensure the environment installs cleanly without wheel compile errors."
* **Verification Check:** Run `pip install -r requirements.txt` and verify exit code 0.
* **Tests That Must Pass:** `test_tc_t_20_folder_structure_integrity`

---

### Task BUILD-02: Data Ingestion, Sanitization & Caching Layer
* **Goal:** Create `src/data_loader.py` to ingest SQLite tables, impute missing discount values with `0.0`, purge duplicate transactions, and cache the 40-column analytical view using `@st.cache_data`.
* **Files:** `src/data_loader.py`
* **Linked Traceability:** US-01, US-02, US-03 | FR-01 | BR-01 | OQ-01
* **Action Prompt for AI/Developer:**
  > "Write `src/data_loader.py` with functions `get_db_connection()`, `load_raw_data()`, `clean_data()`, and `@st.cache_data def load_analytical_view()`. Ensure it handles missing SQLite files gracefully, executes `PRAGMA foreign_keys = ON;`, imputes nulls in `discount_amount` with `0.0`, removes duplicate `sales_id` rows, and returns exactly 300 rows and 40 columns."
* **Verification Check:** Run `python -c "from src.data_loader import load_analytical_view; df = load_analytical_view(); print(df.shape)"` -> Prints `(300, 40)`.
* **Tests That Must Pass:** `test_tc_t_01` to `test_tc_t_10`, `test_tc_t_21`

---

### Task BUILD-03: Analytical & Statistical Computation Engine
* **Goal:** Create `src/analysis.py` containing modular calculation functions for variance metrics, algebraic waterfall decomposition, day-of-week seasonality, OLS linear regression, month-end run-rate forecasting, and Z-score outlier flagging.
* **Files:** `src/analysis.py`
* **Linked Traceability:** US-11, US-12, US-14, US-15 | FR-04, FR-09, FR-10, FR-11, FR-12 | BR-07, BR-09, BR-10, BR-11 | OQ-07, OQ-09, OQ-10
* **Action Prompt for AI/Developer:**
  > "Write `src/analysis.py` implementing:
  > 1. `calculate_kpi_summary(df)`: actual revenue, target, dollar variance, % variance, achievement rate, gross margin.
  > 2. `decompose_variance(actual_orders, target_orders, actual_atv, target_atv)`: computes Volume Effect ($) and Price Effect ($) verifying identity.
  > 3. `calculate_run_rate_forecast(df, elapsed_days, total_month_days)`: current run rate, projected total, projected deficit, required catch-up pace.
  > 4. `compute_ols_regression(x, y)`: slope, intercept, Pearson r, p-value.
  > 5. `detect_anomalies_zscore(df, threshold=3.0)`: flags transactions where $|Z| > 3.0$."
* **Verification Check:** Run `pytest tests/test_analysis.py -v` -> All 7 analytical tests pass.
* **Tests That Must Pass:** `test_tc_t_11` through `test_tc_t_17`

---

### Task BUILD-04: Reusable Tailwind UI Component System
* **Goal:** Extend `src/ui.py` to provide standardized HTML iframe renderers for page headers, 5-tile KPI card grids, insight callouts, and consistent Plotly chart theme templates.
* **Files:** `src/ui.py`
* **Linked Traceability:** US-04 | FR-02 | BR-01 | OQ-01
* **Action Prompt for AI/Developer:**
  > "Ensure `src/ui.py` exposes `render_tailwind_header()`, `render_kpi_cards()`, `render_insight_box()`, and `apply_plotly_theme(fig)`. Ensure responsive styling with CDN Tailwind classes, custom font fallback to Inter/system sans, and proper heights in `st.components.v1.html` to prevent scrollbars."
* **Verification Check:** Import functions in a Python shell without errors.
* **Tests That Must Pass:** `test_tc_t_18`

---

### Task BUILD-05: Main Routing Hub & Executive Overview Page
* **Goal:** Implement `app.py` as the multi-page entry point and `pages/01_Overview.py` as UI-01 (Executive Cockpit, KPI cards header, store ranking deficit table, and ATV gauge).
* **Files:** `app.py`, `pages/01_Overview.py`
* **Linked Traceability:** US-04, US-05, US-06, US-10 | FR-02, FR-03, FR-04 | BR-01, BR-02, BR-08 | OQ-01, OQ-02, OQ-08
* **Action Prompt for AI/Developer:**
  > "Build `app.py` with shared sidebar filter controls bound to `st.session_state` (Date Range, Region, Store Format, Store Name, Category). Build `pages/01_Overview.py` displaying the Tailwind header, 5 KPI cards, store achievement ranking table with deficit badges (<85%), and Plotly ATV indicator bullet gauge."
* **Verification Check:** Run `streamlit run app.py` and inspect Overview page in browser.
* **Tests That Must Pass:** `TC-F-01`, `TC-F-02`, `TC-F-04`, `TC-F-05`, `TC-F-06`

---

### Task BUILD-06: Trends, Seasonality & Cumulative Trajectory Page
* **Goal:** Implement `pages/02_Trends_Seasonality.py` (UI-02) featuring the Day-of-Week Seasonality Heatmap and the MTD Cumulative Burn-Up Curve.
* **Files:** `pages/02_Trends_Seasonality.py`
* **Linked Traceability:** US-08, US-13 | FR-06, FR-08 | BR-04, BR-06 | OQ-04, OQ-06
* **Action Prompt for AI/Developer:**
  > "Build `pages/02_Trends_Seasonality.py` with:
  > 1. Plotly Calendar Heatmap: Region/Store on Y-axis vs Monday–Sunday on X-axis, color-coded by achievement rate.
  > 2. MTD Cumulative Trajectory Line Chart: Dual-trace Plotly chart plotting Cumulative Actual Sales vs Target Pace Line with deficit gap shading."
* **Verification Check:** Inspect Seasonality heatmap in browser; verify Tuesday/Wednesday lulls display amber/red.
* **Tests That Must Pass:** `TC-F-07`, `TC-F-08`

---

### Task BUILD-07: Store & Merchandise Category Benchmarking Page
* **Goal:** Implement `pages/03_Compare.py` (UI-03) featuring the Paired Category Bar Chart and Sales Associate Quota Leaderboard.
* **Files:** `pages/03_Compare.py`
* **Linked Traceability:** US-07, US-09 | FR-05, FR-07 | BR-03, BR-05 | OQ-03, OQ-05
* **Action Prompt for AI/Developer:**
  > "Build `pages/03_Compare.py` displaying:
  > 1. Paired horizontal bar chart comparing Actual vs Target revenue across all 5 merchandise categories.
  > 2. Sales Associate Quota Leaderboard table with sales vs quota, achievement rate, and conditional red badges for associates below 85%."
* **Verification Check:** Select store filter; verify staff table filters to associates of that store.
* **Tests That Must Pass:** `TC-F-09`, `TC-F-10`

---

### Task BUILD-08: Root Cause Variance Decomposition & Elasticity Page
* **Goal:** Implement `pages/04_Root_Cause.py` (UI-04) featuring the Volume vs Price Waterfall Decomposition and Discount vs Margin Scatter Plot.
* **Files:** `pages/04_Root_Cause.py`
* **Linked Traceability:** US-11, US-12 | FR-09, FR-10 | BR-07, BR-09 | OQ-07, OQ-09
* **Action Prompt for AI/Developer:**
  > "Build `pages/04_Root_Cause.py` featuring:
  > 1. Plotly Waterfall Chart: Target Revenue $\rightarrow$ Volume Effect $\rightarrow$ Price Effect $\rightarrow$ Actual Revenue.
  > 2. Promotional Discount vs Realized Gross Margin Scatter Plot with OLS linear regression trendline and Pearson r card."
* **Verification Check:** Verify waterfall starting bar matches target and concluding bar matches actual.
* **Tests That Must Pass:** `TC-F-11`, `TC-F-12`, `test_tc_t_13`, `test_tc_t_17`

---

### Task BUILD-09: Predictive Run-Rate & Month-End Forecasting Page
* **Goal:** Implement `pages/05_Forecasting.py` (UI-05) featuring daily velocity gauges, projected full-month shortfall cards, and interactive what-if catch-up pace sliders.
* **Files:** `pages/05_Forecasting.py`
* **Linked Traceability:** US-14 | FR-11 | BR-10 | OQ-10
* **Action Prompt for AI/Developer:**
  > "Build `pages/05_Forecasting.py` displaying:
  > 1. Velocity Summary Cards: Current Daily Average, Projected Month-End Revenue, Expected Quota Shortfall.
  > 2. Catch-up Pace Slider: Simulates required daily run-rate for remaining 10 days to hit 100% quota."
* **Verification Check:** Move slider; verify catch-up quota recalculates dynamically.
* **Tests That Must Pass:** `TC-F-13`, `test_tc_t_15`

---

### Task BUILD-10: Data Governance, Anomaly Inspector & About Pages
* **Goal:** Implement `pages/06_Data_Governance.py` (UI-06) for data quality audit, Z-score outlier table ($Z > 3.0$), Excel export, and `pages/07_About.py` (UI-07) for credentials, DDL, and architecture.
* **Files:** `pages/06_Data_Governance.py`, `pages/07_About.py`
* **Linked Traceability:** US-02, US-15 | FR-12 | BR-11 | OQ-01, OQ-02
* **Action Prompt for AI/Developer:**
  > "Build `pages/06_Data_Governance.py` with data quality metrics (6 nulls imputed, 3 duplicates dropped), Z-score outlier inspection table ($|Z| > 3.0$), outlier exclusion checkbox, and an Excel/CSV download button. Build `pages/07_About.py` displaying Harshit's academic credentials, Dinesh Sereno's mentorship, the 5-layer pipeline architecture, and the SQLite schema DDL."
* **Verification Check:** Click 'Download Cleaned Dataset'; confirm Excel file downloads. Toggle outlier exclusion; confirm regression updates.
* **Tests That Must Pass:** `TC-F-14`, `TC-F-15`, `TC-F-16`, `test_tc_t_16`

---

### Task BUILD-11: Full Automated Test Suite Execution & Verification Gate
* **Goal:** Execute the full test suite via pytest and confirm 100% test passage.
* **Files:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
* **Linked Traceability:** All BRs, FRs, US, and TCs
* **Action Prompt for AI/Developer:**
  > "Run `pytest -v` across all test suites. Confirm that all 21 technical unit tests pass with zero failures and zero unexpected errors."
* **Verification Check:** `pytest -v` exits with code 0 and reports `21 passed`.
* **Tests That Must Pass:** `TC-T-01` through `TC-T-21`

---

### Task BUILD-12: README Documentation & Deployment Packaging
* **Goal:** Write an executive `README.md` with project background, setup commands, architecture diagram, feature screenshots guide, and GitHub push instructions for Streamlit Community Cloud.
* **Files:** `README.md`, `.gitignore`
* **Linked Traceability:** Non-Functional Requirements | Release Gate
* **Action Prompt for AI/Developer:**
  > "Write comprehensive `README.md` with: Overview, Key Features, Tech Stack, Installation Guide, Running the App, Running Tests, Data Model & Schema, and Deployment Instructions."
* **Verification Check:** Readme is well-formatted and all repository links work.

---

## 4. Definition of Done (DoD) Checklist

To declare the build phase complete and ready for deployment:
- [ ] 100% of "Must Have" Business Requirements (BR-01, 02, 03, 06, 08, 09, 10) fully functional.
- [ ] All 10 Operational Questions (OQ-01 to OQ-10) answered by dedicated interactive visuals.
- [ ] All 21 Technical Unit Tests (`pytest -v`) pass with 100% success rate.
- [ ] Query latency for analytical view remains strictly below 1.0 second.
- [ ] Cascading multi-filter sidebar functions smoothly without page reload errors.
- [ ] Zero unhandled exceptions or broken Streamlit tracebacks on edge filter selections.
- [ ] Mobile/phone responsive layout verified without horizontal clipping.
- [ ] Complete `README.md` and `.gitignore` committed to repository.

---

## 5. Developer Command Cheat Sheet

```bash
# 1. Install production dependencies
pip install -r requirements.txt

# 2. Re-generate synthetic data & SQLite database (reproducible seed=42)
python src/generate_data.py

# 3. Execute all 21 automated technical unit tests
pytest -v

# 4. Execute specific test modules
pytest tests/test_data.py -v
pytest tests/test_analysis.py -v
pytest tests/test_app.py -v

# 5. Launch the Streamlit multi-page application locally
streamlit run app.py
```

---

## 6. Master Viva Voce Preparation (10 Likely Questions & Model Answers)

**Q1: What is the core business problem Project P054 solves?**
> **Answer:** It solves the 48-to-72-hour reporting latency in multi-unit retail operations. Instead of waiting for weekly consolidated batch reports to discover store deficits, retail leaders get automated daily variance analysis, immediate deficit alerts (<85%), and actionable root-cause decomposition before month-end quota closure.

**Q2: Explain the difference between fact_daily_targets and fact_daily_sales in your database design.**
> **Answer:** It is a constellation schema with two fact tables at different grains. `fact_daily_targets` holds planned budget quotas generated during quarterly planning. `fact_daily_sales` holds actual operational sales transactions logged by POS systems. They are joined in `vw_daily_sales_target_analysis` on `(transaction_date, store_id, category_id)` to calculate daily variance.

**Q3: Why did you compute derived variance metrics inside a SQL View rather than in Python runtime?**
> **Answer:** Pre-computing variance metrics (`dollar_variance`, `pct_variance`, `volume_effect_dollar`, `price_effect_dollar`) inside SQLite's `vw_daily_sales_target_analysis` guarantees single-source-of-truth consistency across all 7 UI pages and offloads mathematical operations to the C-based database engine, ensuring UI page loads complete in under 0.5 seconds.

**Q4: Explain the mathematical identity behind your Volume vs Price Variance Waterfall.**
> **Answer:** $\text{Volume Effect} = (\text{Actual Orders} - \text{Target Orders}) \times \text{Target ATV}$ and $\text{Price Effect} = (\text{Actual ATV} - \text{Target ATV}) \times \text{Actual Orders}$. Their sum algebraically equals $(\text{Actual Orders} \times \text{Actual ATV}) - (\text{Target Orders} \times \text{Target ATV}) = \text{Actual Revenue} - \text{Target Revenue}$. This proves whether a revenue gap was caused by customer footfall lulls or lower spending per transaction.

**Q5: What is Simpson's Paradox, and where did you encounter it in your analysis?**
> **Answer:** In our discount vs margin elasticity analysis (OQ-07), aggregating all categories together showed a positive correlation because high-margin Apparel (48%) has high discounts (18%), while low-margin Electronics (22%) has low discounts (5%). However, analyzing *within* each category reveals the true commercial relationship: promotional discounts directly erode gross profit margin ($r < -0.30$).

**Q6: Why did you use `st.session_state` instead of standard Streamlit widget parameters?**
> **Answer:** Streamlit re-executes the active page script on every user interaction. When users navigate between multi-page tabs (e.g. from Overview to Root Cause), widget variables are lost. `st.session_state` persists filter criteria (Date Range, Region, Store, Category) across all page navigations.

**Q7: How did you implement Tailwind CSS inside Streamlit, and why?**
> **Answer:** We rendered lightweight HTML cards using `st.components.v1.html` loading the Tailwind CDN. This provides modern styling (responsive grids, colored badges, alert callouts, and clean borders) while letting native Streamlit widgets and Plotly charts handle data processing and interactive visualization.

**Q8: How does your run-rate forecasting model project month-end revenue?**
> **Answer:** It calculates average daily revenue velocity over elapsed trading days ($\text{Daily Run Rate} = \frac{\text{Cumulative Revenue}}{\text{Elapsed Days}}$) and scales it to total month days ($\text{Projected Revenue} = \text{Daily Run Rate} \times \text{Total Month Days}$). It then subtracts the monthly target quota to project the deficit and calculates the required daily run-rate for the remaining days.

**Q9: Why did you inject data quality issues (NaNs, duplicates, outliers) into the synthetic data?**
> **Answer:** Real enterprise POS data always contains imperfections. By injecting 6 missing discount values, 3 duplicate transactional records, and 3 festival flash outliers (>3.0 Z-scores), we prove that our ETL pipeline (`src/data_loader.py`) and data governance page (`UI-06`) can sanitize data and allow analysts to inspect anomalies.

**Q10: What automated testing strategy did you follow before building the application?**
> **Answer:** We followed the Test Pyramid. We built 21 automated technical unit tests in `pytest` covering database schemas, primary/foreign key integrity, cleaning routines, mathematical identities, and latency benchmarks, plus 16 functional acceptance tests (TC-F) derived from Gherkin user story criteria.
