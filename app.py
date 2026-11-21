"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Main Streamlit Application Entry Point (app.py)
Author: Harshit (Roll No: 2500320140058) | Mentor: Dinesh Sereno
"""

import streamlit as st
import pandas as pd
from datetime import date
from src.data_loader import load_analytical_view, load_dimension_tables, filter_dataset
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.analysis import calculate_kpi_summary

# Configure Streamlit Page
st.set_page_config(
    page_title="P054: Daily Sales Actual vs Target Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize Global Session State Filters
if "date_range" not in st.session_state:
    st.session_state["date_range"] = (date(2026, 11, 1), date(2026, 11, 20))
if "selected_regions" not in st.session_state:
    st.session_state["selected_regions"] = ["All Regions"]
if "selected_formats" not in st.session_state:
    st.session_state["selected_formats"] = ["All Formats"]
if "selected_stores" not in st.session_state:
    st.session_state["selected_stores"] = ["All Stores"]
if "selected_categories" not in st.session_state:
    st.session_state["selected_categories"] = ["All Categories"]

# Load Data Layer via st.cache_data
df_master = load_analytical_view()
dim_stores, dim_categories, dim_staff = load_dimension_tables()

# ==========================================
# SIDEBAR NAVIGATION & CASCADING FILTERS
# ==========================================
with st.sidebar:
    st.markdown("### 🏬 **Retail Sales Operations**")
    st.caption("Project P054 | Harshit (ABES) | Mentor: Dinesh Sereno")
    st.divider()
    
    st.markdown("#### 🔍 **Global Performance Filters**")
    
    # 1. Date Range Filter
    min_date = df_master["transaction_date"].min().date()
    max_date = df_master["transaction_date"].max().date()
    date_val = st.date_input(
        "Trading Date Range",
        value=st.session_state["date_range"],
        min_value=min_date,
        max_value=max_date,
        key="global_date_picker"
    )
    if isinstance(date_val, tuple) and len(date_val) == 2:
        st.session_state["date_range"] = date_val
    elif isinstance(date_val, list) and len(date_val) == 2:
        st.session_state["date_range"] = tuple(date_val)

    # 2. Region Filter
    region_options = ["All Regions"] + sorted(list(dim_stores["region"].unique()))
    sel_region = st.selectbox("Geographic Region", region_options, index=0)
    st.session_state["selected_regions"] = [sel_region] if sel_region != "All Regions" else ["All Regions"]

    # 3. Store Format Filter
    format_options = ["All Formats"] + sorted(list(dim_stores["store_format"].unique()))
    sel_format = st.selectbox("Store Format", format_options, index=0)
    st.session_state["selected_formats"] = [sel_format] if sel_format != "All Formats" else ["All Formats"]

    # 4. Cascading Store Selection
    eligible_stores = dim_stores.copy()
    if sel_region != "All Regions":
        eligible_stores = eligible_stores[eligible_stores["region"] == sel_region]
    if sel_format != "All Formats":
        eligible_stores = eligible_stores[eligible_stores["store_format"] == sel_format]
        
    store_options = ["All Stores"] + sorted(list(eligible_stores["store_name"].unique()))
    sel_store = st.selectbox("Retail Store Branch", store_options, index=0)
    st.session_state["selected_stores"] = [sel_store] if sel_store != "All Stores" else ["All Stores"]

    # 5. Merchandise Category Filter
    cat_options = ["All Categories"] + sorted(list(dim_categories["category_name"].unique()))
    sel_cat = st.selectbox("Merchandise Category", cat_options, index=0)
    st.session_state["selected_categories"] = [sel_cat] if sel_cat != "All Categories" else ["All Categories"]

    st.divider()
    st.markdown("#### ⚡ **System Telemetry**")
    st.caption("Engine: SQLite + Caching Layer (`@st.cache_data`)")
    st.caption(f"Active Master Records: {len(df_master):,} rows | 40 cols")
    
    if st.button("🔄 Flush In-Memory Cache", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Apply Global Filters to Active Slice
filtered_df = filter_dataset(
    df_master,
    date_range=st.session_state["date_range"],
    regions=st.session_state["selected_regions"],
    formats=st.session_state["selected_formats"],
    stores=st.session_state["selected_stores"],
    categories=st.session_state["selected_categories"]
)

# ==========================================
# HOME PAGE / EXECUTIVE DASHBOARD LANDING
# ==========================================
render_tailwind_header(
    title="Daily Sales Actual vs Target Analysis Cockpit",
    subtitle="Enterprise performance intelligence monitoring daily monetary variance, quota achievement rates, and basket depth across multi-unit retail networks.",
    meta_tag="P054 Enterprise Retail Operations Cockpit"
)

# Compute Executive KPIs
kpi = calculate_kpi_summary(filtered_df)

if len(filtered_df) == 0:
    render_insight_box(
        title="Zero Matching Transactions",
        message="No transactions match the selected filter combination. Please expand your date range or reset regional filters.",
        alert_type="warning"
    )
else:
    # 5-Tile Tailwind KPI Card Data
    kpi_cards = [
        {
            "label": "Actual Net Sales",
            "value": f"₹{kpi['actual_revenue']:,.2f}",
            "delta": f"{kpi['achievement_rate']:.1f}% Quota",
            "is_positive": kpi["achievement_rate"] >= 100.0,
            "subtext": f"{kpi['actual_orders']:,} Orders"
        },
        {
            "label": "Target Quota",
            "value": f"₹{kpi['target_revenue']:,.2f}",
            "delta": "Planned Base",
            "is_positive": True,
            "subtext": f"{kpi['target_orders']:,} Target Orders"
        },
        {
            "label": "Dollar Variance",
            "value": f"{'+' if kpi['dollar_variance']>=0 else ''}₹{kpi['dollar_variance']:,.2f}",
            "delta": f"{kpi['pct_variance']:+.1f}%",
            "is_positive": kpi["dollar_variance"] >= 0,
            "subtext": "Net Revenue Gap"
        },
        {
            "label": "Achievement Status",
            "value": f"{kpi['achievement_rate']:.2f}%",
            "delta": kpi["status"],
            "is_positive": kpi["achievement_rate"] >= 85.0,
            "subtext": "Alert: <85% Deficit"
        },
        {
            "label": "Realized Margin",
            "value": f"{kpi['realized_margin_pct']:.1f}%",
            "delta": f"₹{kpi['gross_profit']:,.0f} Profit",
            "is_positive": kpi["realized_margin_pct"] >= 35.0,
            "subtext": "Gross Profit Margin"
        }
    ]
    render_kpi_cards(kpi_cards)

    st.markdown("---")
    
    # Executive Summary Columns
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.markdown("### 🏆 **Store Performance Ranking & Deficit Summary (OQ-02)**")
        store_summary = filtered_df.groupby(["store_id", "store_name", "region", "store_format"]).agg(
            actual_sales=("actual_net_revenue", "sum"),
            target_sales=("target_revenue", "sum"),
            orders=("actual_orders", "sum")
        ).reset_index()
        
        store_summary["dollar_variance"] = store_summary["actual_sales"] - store_summary["target_sales"]
        store_summary["achievement_pct"] = (store_summary["actual_sales"] / store_summary["target_sales"]) * 100.0
        store_summary["status"] = store_summary["achievement_pct"].apply(
            lambda x: "🟢 EXCEEDING" if x >= 100 else ("🟡 AT RISK" if x >= 85 else "🔴 CRITICAL DEFICIT")
        )
        store_summary = store_summary.sort_values(by="achievement_pct", ascending=True)
        
        st.dataframe(
            store_summary.style.format({
                "actual_sales": "₹{:,.2f}",
                "target_sales": "₹{:,.2f}",
                "dollar_variance": "₹{:,.2f}",
                "achievement_pct": "{:.1f}%"
            }).background_gradient(subset=["achievement_pct"], cmap="RdYlGn", vmin=75, vmax=115),
            use_container_width=True,
            height=320
        )

    with col2:
        st.markdown("### 🎯 **Average Transaction Value (ATV) Depth (OQ-08)**")
        st.metric("Realized Basket ATV", f"₹{kpi['actual_atv']:,.2f}", delta=f"{kpi['actual_atv'] - kpi['target_atv']:+,.2f} vs Target")
        st.metric("Target Basket Benchmark", f"₹{kpi['target_atv']:,.2f}")
        
        if kpi["actual_atv"] < kpi["target_atv"]:
            render_insight_box(
                title="Basket Depth Deficit Detected",
                message=f"Realized ATV (₹{kpi['actual_atv']:,.0f}) is lagging behind target (₹{kpi['target_atv']:,.0f}). Store staff should focus on promotional cross-selling and multi-item bundling.",
                alert_type="danger"
            )
        else:
            render_insight_box(
                title="Basket Depth On Track",
                message=f"Realized ATV (₹{kpi['actual_atv']:,.0f}) exceeds planned benchmark (₹{kpi['target_atv']:,.0f}), indicating healthy transaction value depth.",
                alert_type="success"
            )

    st.markdown("---")
    st.info("👈 **Navigate using the multi-page menu on the sidebar to inspect Trends, Category Comparisons, Root Cause Decomposition, Run-Rate Forecasting, and Data Quality Audits.**")
