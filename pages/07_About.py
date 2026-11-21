"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Screen UI-07: System Architecture, Academic Credentials & About
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import streamlit as st
import os
from src.ui import render_tailwind_header

st.set_page_config(page_title="About & Architecture | P054", page_icon="ℹ️", layout="wide")

render_tailwind_header(
    title="System Architecture, Data Model & Project Credentials",
    subtitle="Academic project documentation, five-layer application architecture pipeline, and relational schema specifications.",
    meta_tag="UI-07 Architecture & About"
)

# -------------------------------------------------------------
# 1. ACADEMIC & MENTORSHIP CREDENTIALS
# -------------------------------------------------------------
st.markdown("### 🎓 **Academic Project & Governance Credentials**")

col_cred1, col_cred2 = st.columns(2)

with col_cred1:
    st.info("""
    **Candidate Information:**
    * **Learner / Data Analyst:** Harshit
    * **University Roll No:** 2500320140058
    * **Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
    * **Academic Program:** Master of Computer Applications (MCA) — 2nd Year
    * **Specialization:** Data Analytics (DA)
    * **Institution:** ABES Engineering College, Ghaziabad
    """)

with col_cred2:
    st.success("""
    **Project & Mentorship Governance:**
    * **Project ID:** P054
    * **Project Title:** Daily Sales Actual vs Target Analysis
    * **Project Mentor:** Dinesh Sereno (Industry Data Engineering Lead)
    * **Domain:** Enterprise Retail Operations & Sales Analytics
    * **Release Version:** 1.0 (Production MVP)
    * **Baseline Trading Period:** November 2026 (20 Consecutive Days)
    """)

st.markdown("---")

# -------------------------------------------------------------
# 2. FIVE-LAYER APPLICATION ARCHITECTURE
# -------------------------------------------------------------
st.markdown("### 🏗️ **Five-Layer Enterprise Software Pipeline**")

st.markdown("""
```text
+---------------------------------------------------------------------------------------------------+
|  LAYER 1: RELATIONAL DATA ENGINE                                                                  |
|  - SQLite (data/app.db) with PRAGMA foreign_keys = ON; Constellation Star Schema                  |
|  - Tables: dim_stores (15), dim_categories (5), dim_staff (75), fact_daily_targets, fact_daily_sales|
+---------------------------------------------------------------------------------------------------+
                                                  │
                                                  ▼
+---------------------------------------------------------------------------------------------------+
|  LAYER 2: DATA INGESTION & SANITIZATION PIPELINE (src/data_loader.py)                             |
|  - Imputes missing discount values (NaN -> 0.0) & drops duplicate sales records by Primary Key    |
|  - Validates foreign keys & non-negative constraints across transactions                          |
+---------------------------------------------------------------------------------------------------+
                                                  │
                                                  ▼
+---------------------------------------------------------------------------------------------------+
|  LAYER 3: IN-MEMORY CACHING ARCHITECTURE                                                          |
|  - Streamlit @st.cache_data decorator loads 40-column analytical view                             |
|  - Sub-second query retrieval (< 0.5s) for instant multi-page filtering                           |
+---------------------------------------------------------------------------------------------------+
                                                  │
                                                  ▼
+---------------------------------------------------------------------------------------------------+
|  LAYER 4: STATISTICAL & ANALYTICAL COMPUTATION ENGINE (src/analysis.py)                           |
|  - Algebraic Variance Decomposition: Volume Effect ($) + Price Effect ($) == Total Dollar Deficit |
|  - Ordinary Least Squares (OLS) Regression: Promotional discount elasticity vs gross margin       |
|  - Standardized Z-Score Anomaly Detector: Flags flash trading shocks (|Z| > 3.0 sigma)           |
|  - Run-Rate Velocity Predictor: Projects month-end quota surpluses/deficits                       |
+---------------------------------------------------------------------------------------------------+
                                                  │
                                                  ▼
+---------------------------------------------------------------------------------------------------+
|  LAYER 5: PRESENTATION & USER EXPERIENCE LAYER (src/ui.py & pages/)                               |
|  - Embedded Tailwind CSS CDN components inside st.components.v1.html for executive banners & cards|
|  - Interactive Plotly Graph Objects & Multi-Page Cascading Sidebar Navigation                     |
+---------------------------------------------------------------------------------------------------+
```
""")

st.markdown("---")

# -------------------------------------------------------------
# 3. RELATIONAL DDL SPECIFICATION
# -------------------------------------------------------------
st.markdown("### 🗄️ **Relational Schema Definition (SQLite DDL)**")

ddl_code = """-- Constellation Relational Schema: Project P054
CREATE TABLE dim_stores (
    store_id TEXT PRIMARY KEY,
    store_name TEXT NOT NULL,
    region TEXT NOT NULL CHECK(region IN ('North', 'South', 'East', 'West', 'Central')),
    store_format TEXT NOT NULL CHECK(store_format IN ('Flagship', 'Mall', 'HighStreet', 'Outlet')),
    square_footage INTEGER NOT NULL,
    benchmark_daily_target REAL NOT NULL,
    benchmark_target_atv REAL NOT NULL,
    manager_name TEXT NOT NULL,
    opened_date TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE dim_categories (
    category_id TEXT PRIMARY KEY,
    category_name TEXT NOT NULL UNIQUE,
    department_head TEXT NOT NULL,
    target_gross_margin_pct REAL NOT NULL,
    standard_vat_pct REAL NOT NULL
);

CREATE TABLE dim_staff (
    staff_id TEXT PRIMARY KEY,
    staff_name TEXT NOT NULL,
    store_id TEXT NOT NULL,
    department TEXT NOT NULL,
    role_title TEXT NOT NULL,
    base_daily_quota REAL NOT NULL,
    hire_date TEXT NOT NULL,
    status TEXT NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id)
);

CREATE TABLE fact_daily_targets (
    target_id TEXT PRIMARY KEY,
    target_date TEXT NOT NULL,
    store_id TEXT NOT NULL,
    category_id TEXT NOT NULL,
    target_revenue REAL NOT NULL,
    target_orders INTEGER NOT NULL,
    target_atv REAL NOT NULL,
    target_units INTEGER NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id),
    FOREIGN KEY (category_id) REFERENCES dim_categories(category_id)
);

CREATE TABLE fact_daily_sales (
    sales_id TEXT PRIMARY KEY,
    transaction_date TEXT NOT NULL,
    day_of_week TEXT NOT NULL,
    store_id TEXT NOT NULL,
    category_id TEXT NOT NULL,
    staff_id TEXT NOT NULL,
    gross_revenue REAL NOT NULL,
    discount_amount REAL,
    net_revenue REAL NOT NULL,
    cost_of_goods_sold REAL NOT NULL,
    gross_profit REAL NOT NULL,
    order_count INTEGER NOT NULL,
    units_sold INTEGER NOT NULL,
    realized_atv REAL NOT NULL,
    cash_revenue REAL NOT NULL,
    card_digital_revenue REAL NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id),
    FOREIGN KEY (category_id) REFERENCES dim_categories(category_id),
    FOREIGN KEY (staff_id) REFERENCES dim_staff(staff_id)
);"""

st.code(ddl_code, language="sql")
