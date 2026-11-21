# Project P054: Daily Sales Actual vs Target Analysis

[![Python](https://img.shields.io/badge/Python-3.12%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.35%2B-FF4B4B.svg)](https://streamlit.io/)
[![SQLite](https://img.shields.io/badge/Database-SQLite3-003B57.svg)](https://www.sqlite.org/)
[![Plotly](https://img.shields.io/badge/Visualizations-Plotly-3F4F75.svg)](https://plotly.com/)
[![TailwindCSS](https://img.shields.io/badge/UI-Tailwind_CSS-38B2AC.svg)](https://tailwindcss.com/)
[![Tests](https://img.shields.io/badge/Tests-21%20Passed-brightgreen.svg)](https://docs.pytest.org/)

An enterprise-grade retail sales operations analytics web application engineered to eliminate the 48-to-72-hour reporting latency in multi-unit retail networks. The system computes daily monetary variances, identifies deficit branches requiring immediate intervention (<85% quota attainment), decomposes root-cause variance into volume vs ticket effects, assesses promotional markdown margin erosion, and forecasts month-end revenues using daily run-rate velocity.

---

## 📌 Project Credentials & Academic Metadata
* **Project ID:** P054
* **Project Title:** Daily Sales Actual vs Target Analysis
* **Candidate / Lead Analyst:** Harshit
* **University Roll No:** 2500320140058
* **Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
* **Academic Program:** Master of Computer Applications (MCA) — 2nd Year (Data Analytics)
* **Institution:** ABES Engineering College, Ghaziabad
* **Project Mentor:** Dinesh Sereno (Industry Data Engineering Lead)
* **Release Version:** 1.0 (Production Release) | **Baseline:** November 2026

---

## 🏗️ Technical Architecture & Pipeline
The application implements a five-layer decoupled architecture:
1. **Relational Storage Engine:** ACID-compliant SQLite (`data/app.db`) with Foreign Keys enforced via `PRAGMA foreign_keys = ON;`.
2. **Data Ingestion & Cleaning Layer (`src/data_loader.py`):** Automatically sanitizes raw POS feeds by imputing missing discount values (NaN $\rightarrow$ 0.0) and purging duplicate transactions.
3. **In-Memory Caching Architecture:** Streamlit `@st.cache_data` caches the 40-column analytical view (`vw_daily_sales_target_analysis`) enabling sub-second query retrieval ($<0.5\text{s}$).
4. **Analytical & Statistical Engine (`src/analysis.py`):** Computes variance metrics, algebraic waterfall decomposition, OLS linear regression, and Z-score anomaly detection ($|Z| > 3.0\sigma$).
5. **Presentation & UX Layer (`src/ui.py` & `pages/`):** Tailwind CSS components embedded inside `st.components.v1.html` for executive banners, KPI cards, and alert callouts, paired with Plotly interactive charts.

---

## 🚀 Quick Start Guide

### 1. Clone & Set Up Environment
```bash
# Clone the repository
git clone https://github.com/<your-username>/p054-sales-actual-vs-target.git
cd p054-sales-actual-vs-target

# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

# Install pinned dependencies
pip install -r requirements.txt
```

### 2. Generate / Verify Synthetic Dataset (Seed 42)
```bash
python src/generate_data.py
```

### 3. Run Automated Pytest Test Suite (21 Tests)
```bash
pytest -v
```

### 4. Launch the Streamlit Multi-Page Web Application
```bash
streamlit run app.py
```
The application will open automatically in your browser at `http://localhost:8501`.

---

## 📊 Application Screens Walkthrough (7 Screens)

| Screen ID | Title | File | Core Capabilities |
|-----------|-------|------|-------------------|
| **UI-01** | Executive Overview | `pages/01_Overview.py` | 5 Tailwind KPI cards, sortable store achievement ranking table, ATV indicator gauge |
| **UI-02** | Trends & Seasonality | `pages/02_Trends_Seasonality.py` | Day-of-week achievement heatmap (Mon–Sun), MTD cumulative sales vs target pace curve |
| **UI-03** | Performance Comparison | `pages/03_Compare.py` | Paired horizontal category bar chart, sales associate leaderboard with coaching alerts |
| **UI-04** | Root Cause & Elasticity | `pages/04_Root_Cause.py` | Algebraic waterfall decomposition (Volume vs Price effect), discount vs margin scatter |
| **UI-05** | Predictive Forecasting | `pages/05_Forecasting.py` | Elapsed daily velocity, projected month-end total, interactive catch-up pace slider |
| **UI-06** | Data Governance & Audit | `pages/06_Data_Governance.py` | ETL cleaning audit log (nulls/duplicates), Z-score outlier table ($|Z| > 3.0$), Excel export |
| **UI-07** | Architecture & About | `pages/07_About.py` | Credentials, five-layer architecture flowchart, and SQLite relational DDL schema |

---

## 📐 Mathematical Formulation & Analytical Identities

### 1. Volume vs Price Variance Waterfall Decomposition (OQ-09)
$$\text{Volume Effect (\$) } = (\text{Actual Orders} - \text{Target Orders}) \times \text{Target ATV}$$
$$\text{Price Effect (\$) } = (\text{Actual ATV} - \text{Target ATV}) \times \text{Actual Orders}$$
$$\text{Total Modeled Variance } = \text{Volume Effect } + \text{Price Effect } \equiv \text{Actual Revenue } - \text{Target Basket Revenue}$$

### 2. Month-End Run-Rate Revenue Forecasting (OQ-10)
$$\text{Daily Run Rate } = \frac{\text{Cumulative Sales}}{\text{Elapsed Trading Days}}$$
$$\text{Projected Month-End Sales } = \text{Daily Run Rate} \times \text{Total Month Days (30)}$$
$$\text{Required Daily Catch-Up Rate } = \frac{\text{Monthly Target} - \text{Cumulative Sales}}{\text{Remaining Days (10)}}$$

### 3. Standardized Statistical Anomaly Detection (OQ-01, OQ-02)
$$Z = \frac{x_i - \mu}{\sigma}$$
* Records with $|Z| \ge 3.0$ standard deviations are isolated as external flash trading events.

---

## 📂 Formal Project Documentation Index
All formal project deliverables generated across the 18-step engineering lifecycle are stored in `docs/`:
* [01_Project_Understanding.docx](docs/01_Project_Understanding.docx)
* [02_HMW_Statements.docx](docs/02_HMW_Statements.docx)
* [03_Needs_Analysis.docx](docs/03_Needs_Analysis.docx)
* [04_Traceability_Matrix.xlsx](docs/04_Traceability_Matrix.xlsx)
* [05_BRD.docx](docs/05_BRD.docx)
* [07_TRD.docx](docs/07_TRD.docx)
* [08_Database_Design.xlsx](docs/08_Database_Design.xlsx)
* [09_ER_Diagram.docx](docs/09_ER_Diagram.docx)
* [11_User_Stories.docx](docs/11_User_Stories.docx)
* [12_UI_Framework.docx](docs/12_UI_Framework.docx)
* [13_Test_Cases.docx](docs/13_Test_Cases.docx)
* [14_BUILD_PLAN.docx](docs/14_BUILD_PLAN.docx)

---

## ☁️ Deployment Guide (Streamlit Community Cloud)
1. Push repository to GitHub: `git push origin main`
2. Log into [Streamlit Community Cloud](https://share.streamlit.io/).
3. Click **"New app"**, select your repository, branch (`main`), and set Main file path to `app.py`.
4. The cloud build installs dependencies from `requirements.txt` and serves the app live.
