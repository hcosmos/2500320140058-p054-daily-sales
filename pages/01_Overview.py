"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Screen UI-01: Executive Overview & KPI Cockpit
Author: Harshit (Roll No: 2500320140058) | Mentor: Dinesh Sereno
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_analytical_view, filter_dataset
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.analysis import calculate_kpi_summary

st.set_page_config(page_title="Executive Overview | P054", page_icon="🏢", layout="wide")

render_tailwind_header(
    title="Executive Overview & Store Deficit Cockpit",
    subtitle="High-level enterprise actual vs target revenue tracking, percentage variances, and branch operational health rankings.",
    meta_tag="UI-01 Executive Overview"
)

df_master = load_analytical_view()

# Apply Session State Filters
filtered_df = filter_dataset(
    df_master,
    date_range=st.session_state.get("date_range"),
    regions=st.session_state.get("selected_regions"),
    formats=st.session_state.get("selected_formats"),
    stores=st.session_state.get("selected_stores"),
    categories=st.session_state.get("selected_categories")
)

kpi = calculate_kpi_summary(filtered_df)

if len(filtered_df) == 0:
    render_insight_box("Zero Records", "No records found for current filter selection.", "warning")
else:
    # 5-Tile KPI Cards
    cards = [
        {"label": "Actual Revenue", "value": f"₹{kpi['actual_revenue']:,.2f}", "delta": f"{kpi['achievement_rate']:.1f}% Quota", "is_positive": kpi["achievement_rate"]>=100, "subtext": f"{kpi['actual_orders']:,} Orders"},
        {"label": "Target Quota", "value": f"₹{kpi['target_revenue']:,.2f}", "delta": "Target Baseline", "is_positive": True, "subtext": f"{kpi['target_orders']:,} Target Orders"},
        {"label": "Monetary Variance", "value": f"{'+' if kpi['dollar_variance']>=0 else ''}₹{kpi['dollar_variance']:,.2f}", "delta": f"{kpi['pct_variance']:+.1f}%", "is_positive": kpi["dollar_variance"]>=0, "subtext": "Net Deficit/Surplus"},
        {"label": "Quota Achievement", "value": f"{kpi['achievement_rate']:.2f}%", "delta": kpi["status"], "is_positive": kpi["achievement_rate"]>=85, "subtext": "Alert: <85% Deficit"},
        {"label": "Gross Margin %", "value": f"{kpi['realized_margin_pct']:.1f}%", "delta": f"₹{kpi['gross_profit']:,.0f} Margin", "is_positive": kpi["realized_margin_pct"]>=35, "subtext": "Realized Margin"}
    ]
    render_kpi_cards(cards)

    st.markdown("---")

    col_chart, col_gauge = st.columns([3, 2])

    with col_chart:
        st.markdown("#### 📊 **Store Quota Attainment Ranking (OQ-02)**")
        store_df = filtered_df.groupby("store_name").agg(
            actual=("actual_net_revenue", "sum"),
            target=("target_revenue", "sum")
        ).reset_index()
        store_df["achievement_pct"] = (store_df["actual"] / store_df["target"]) * 100.0
        store_df = store_df.sort_values(by="achievement_pct", ascending=True)

        fig_bar = px.bar(
            store_df,
            x="achievement_pct",
            y="store_name",
            orientation="h",
            color="achievement_pct",
            color_continuous_scale=["#EF4444", "#F59E0B", "#10B981"],
            text_auto=".1f",
            title="Store Target Achievement Rate (%)",
            labels={"achievement_pct": "Achievement Rate (%)", "store_name": "Store Branch"}
        )
        fig_bar.add_vline(x=100.0, line_dash="dash", line_color="#1E3A8A", annotation_text="100% Quota Target")
        fig_bar.add_vline(x=85.0, line_dash="dot", line_color="#EF4444", annotation_text="85% Deficit Line")
        fig_bar.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=380, plot_bgcolor="white")
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_gauge:
        st.markdown("#### 🎯 **Average Transaction Value (ATV) Benchmark Gauge (OQ-08)**")
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=kpi["actual_atv"],
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Realized Basket Size (₹)", 'font': {'size': 16, 'color': '#1E3A8A'}},
            delta={'reference': kpi["target_atv"], 'increasing': {'color': "#10B981"}, 'decreasing': {'color': "#EF4444"}},
            gauge={
                'axis': {'range': [None, max(5000, kpi["target_atv"] * 1.5)]},
                'bar': {'color': "#1E3A8A"},
                'steps': [
                    {'range': [0, kpi["target_atv"] * 0.85], 'color': "#FEE2E2"},
                    {'range': [kpi["target_atv"] * 0.85, kpi["target_atv"]], 'color': "#FEF3C7"},
                    {'range': [kpi["target_atv"], max(5000, kpi["target_atv"] * 1.5)], 'color': "#D1FAE5"}
                ],
                'threshold': {
                    'line': {'color': "#10B981", 'width': 4},
                    'thickness': 0.75,
                    'value': kpi["target_atv"]
                }
            }
        ))
        fig_gauge.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=380)
        st.plotly_chart(fig_gauge, use_container_width=True)

    st.markdown("---")
    st.markdown("#### 📋 **Store Operational Deficit Audit Grid**")
    grid_df = filtered_df.groupby(["store_id", "store_name", "region", "store_format"]).agg(
        actual_revenue=("actual_net_revenue", "sum"),
        target_revenue=("target_revenue", "sum"),
        actual_orders=("actual_orders", "sum"),
        target_orders=("target_orders", "sum")
    ).reset_index()
    grid_df["dollar_variance"] = grid_df["actual_revenue"] - grid_df["target_revenue"]
    grid_df["achievement_pct"] = (grid_df["actual_revenue"] / grid_df["target_revenue"]) * 100.0
    grid_df["alert_status"] = grid_df["achievement_pct"].apply(
        lambda x: "🟢 On Target" if x >= 100 else ("🟡 Watchlist" if x >= 85 else "🔴 CRITICAL DEFICIT")
    )
    grid_df = grid_df.sort_values(by="achievement_pct", ascending=True)

    st.dataframe(
        grid_df.style.format({
            "actual_revenue": "₹{:,.2f}",
            "target_revenue": "₹{:,.2f}",
            "dollar_variance": "₹{:,.2f}",
            "achievement_pct": "{:.2f}%",
            "actual_orders": "{:,}",
            "target_orders": "{:,}"
        }),
        use_container_width=True
    )
