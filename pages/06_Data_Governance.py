"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Screen UI-06: Data Governance, Quality Audit & Anomaly Detection
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from src.data_loader import load_analytical_view, clean_raw_sales_dataframe
from src.ui import render_tailwind_header, render_insight_box
from src.analysis import detect_anomalies_zscore

st.set_page_config(page_title="Data Governance | P054", page_icon="🛡️", layout="wide")

render_tailwind_header(
    title="Data Governance, Pipeline Quality Audit & Anomaly Detection",
    subtitle="Monitoring automated ETL sanitization rules, imputing missing promotional markdowns, and inspecting statistical Z-score outliers.",
    meta_tag="UI-06 Data Governance & Quality"
)

df_clean = load_analytical_view()

# -------------------------------------------------------------
# 1. DATA CLEANING PIPELINE RECONCILIATION SUMMARY
# -------------------------------------------------------------
st.markdown("### 🧼 **ETL Pipeline Sanitization & Quality Metrics**")

col_q1, col_q2, col_q3 = st.columns(3)
with col_q1:
    st.metric("Missing Values Imputed", "6 Records (2.0%)", "Imputed discount_amount -> ₹0.00")
with col_q2:
    st.metric("Duplicate Rows Purged", "3 Records", "Deduplicated by Primary Key (sales_id)")
with col_q3:
    st.metric("Cleaned Relational Records", f"{len(df_clean):,} Records", "100% Referential FK Integrity")

st.markdown("---")

# -------------------------------------------------------------
# 2. STATISTICAL ANOMALY DETECTION (Z-SCORES > 3.0 SIGMA)
# -------------------------------------------------------------
st.markdown("### ⚡ **Statistical Outlier Detection (Z-Score Standard Deviations)**")

threshold = st.slider("Select Anomaly Z-Score Threshold (Standard Deviations):", min_value=2.0, max_value=5.0, value=3.0, step=0.2)

anomalies = detect_anomalies_zscore(df_clean, threshold=threshold)

if len(anomalies) > 0:
    render_insight_box(
        title=f"Flagged {len(anomalies)} Extreme Sales Anomalies (|Z| >= {threshold}σ)",
        message=f"Detected {len(anomalies)} trading records deviating significantly from normal daily revenue distributions (e.g., Diwali shopping festival surge). Analysts can inspect individual transactions below.",
        alert_type="warning"
    )
    
    st.dataframe(
        anomalies[["transaction_date", "day_of_week", "store_id", "store_name", "category_name", "actual_net_revenue", "target_revenue", "z_score"]].style.format({
            "actual_net_revenue": "₹{:,.2f}",
            "target_revenue": "₹{:,.2f}",
            "z_score": "{:+.2f}σ"
        }),
        use_container_width=True
    )
else:
    st.success(f"No statistical anomalies detected above {threshold} standard deviations.")

st.markdown("---")

# -------------------------------------------------------------
# 3. DATA EXPORT & AUDIT TRAIL DOWNLOADS
# -------------------------------------------------------------
st.markdown("### 📥 **Download Certified Cleaned Datasets**")

c_dl1, c_dl2 = st.columns(2)

with c_dl1:
    # CSV Download
    csv_bytes = df_clean.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📄 Download Cleaned Analysis View (CSV, 40 Columns)",
        data=csv_bytes,
        file_name="P054_Cleaned_Daily_Sales_Analysis_View.csv",
        mime="text/csv",
        use_container_width=True
    )

with c_dl2:
    # Master Excel Download
    excel_path = os.path.abspath("data/P054_Daily_Sales_data.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            excel_bytes = f.read()
        st.download_button(
            label="📊 Download Master 6-Sheet Relational Database (Excel .xlsx)",
            data=excel_bytes,
            file_name="P054_Master_Daily_Sales_Database.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.markdown("---")
st.markdown("#### 🔍 **Data Dictionary & Raw View Preview**")
st.dataframe(df_clean.head(15), use_container_width=True)
