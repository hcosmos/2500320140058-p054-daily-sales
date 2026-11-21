"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Screen UI-03: Performance Comparison & Associate Leaderboard
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_analytical_view, filter_dataset
from src.ui import render_tailwind_header, render_insight_box
from src.analysis import calculate_staff_leaderboard

st.set_page_config(page_title="Compare Performance | P054", page_icon="⚖️", layout="wide")

render_tailwind_header(
    title="Store, Category & Staff Performance Comparison",
    subtitle="Benchmarking revenue quota contribution across merchandise categories and evaluating individual sales associate quota realization.",
    meta_tag="UI-03 Performance Comparison"
)

df_master = load_analytical_view()

filtered_df = filter_dataset(
    df_master,
    date_range=st.session_state.get("date_range"),
    regions=st.session_state.get("selected_regions"),
    formats=st.session_state.get("selected_formats"),
    stores=st.session_state.get("selected_stores"),
    categories=st.session_state.get("selected_categories")
)

if len(filtered_df) == 0:
    render_insight_box("Zero Records", "No records found for active filters.", "warning")
else:
    # -------------------------------------------------------------
    # 1. MERCHANDISE CATEGORY CONTRIBUTION (OQ-03)
    # -------------------------------------------------------------
    st.markdown("### 🛍️ **Merchandise Category Target vs Actual Attribution (OQ-03)**")
    
    cat_summary = filtered_df.groupby(["category_id", "category_name", "department_head"]).agg(
        actual_revenue=("actual_net_revenue", "sum"),
        target_revenue=("target_revenue", "sum"),
        gross_profit=("gross_profit", "sum"),
        orders=("actual_orders", "sum")
    ).reset_index()

    cat_summary["dollar_variance"] = cat_summary["actual_revenue"] - cat_summary["target_revenue"]
    cat_summary["achievement_pct"] = (cat_summary["actual_revenue"] / cat_summary["target_revenue"]) * 100.0
    cat_summary["realized_margin_pct"] = (cat_summary["gross_profit"] / cat_summary["actual_revenue"]) * 100.0

    # Paired Grouped Bar Chart
    fig_cat = go.Figure()
    fig_cat.add_trace(go.Bar(
        y=cat_summary["category_name"],
        x=cat_summary["target_revenue"],
        name="Target Quota",
        orientation="h",
        marker_color="#94A3B8"
    ))
    fig_cat.add_trace(go.Bar(
        y=cat_summary["category_name"],
        x=cat_summary["actual_revenue"],
        name="Actual Net Sales",
        orientation="h",
        marker_color="#1E3A8A"
    ))

    fig_cat.update_layout(
        barmode="group",
        title="Category Comparison: Target Quota vs Actual Sales Revenue (INR)",
        xaxis_title="Revenue (₹)",
        yaxis_title="Merchandise Category",
        margin=dict(l=20, r=20, t=40, b=20),
        height=320,
        plot_bgcolor="white"
    )
    st.plotly_chart(fig_cat, use_container_width=True)

    # Category Summary Table
    st.dataframe(
        cat_summary.style.format({
            "actual_revenue": "₹{:,.2f}",
            "target_revenue": "₹{:,.2f}",
            "dollar_variance": "₹{:,.2f}",
            "achievement_pct": "{:.1f}%",
            "gross_profit": "₹{:,.2f}",
            "realized_margin_pct": "{:.1f}%"
        }).background_gradient(subset=["achievement_pct"], cmap="RdYlGn", vmin=80, vmax=110),
        use_container_width=True
    )

    st.markdown("---")

    # -------------------------------------------------------------
    # 2. SALES ASSOCIATE QUOTA LEADERBOARD (OQ-05)
    # -------------------------------------------------------------
    st.markdown("### 🧑‍💼 **Sales Associate Quota Attainment Leaderboard (OQ-05)**")
    
    staff_df = calculate_staff_leaderboard(filtered_df)
    
    underperformers = staff_df[staff_df["achievement_pct"] < 85.0]
    
    if len(underperformers) > 0:
        render_insight_box(
            title=f"Actionable Staff Coaching Alert: {len(underperformers)} Associates Below 85% Quota",
            message=f"Identified {len(underperformers)} floor sales staff operating significantly below quota baseline. Store supervisors should inspect shift attendance and provide customer engagement coaching.",
            alert_type="danger"
        )

    st.dataframe(
        staff_df.style.format({
            "total_sales": "₹{:,.2f}",
            "daily_quota": "₹{:,.2f}",
            "total_quota": "₹{:,.2f}",
            "achievement_pct": "{:.1f}%"
        }).applymap(
            lambda v: "color: #EF4444; font-weight: bold;" if "Coaching" in str(v) else "color: #10B981;",
            subset=["coaching_flag"]
        ),
        use_container_width=True,
        height=380
    )
