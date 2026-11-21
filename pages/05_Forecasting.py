"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Screen UI-05: Predictive Run-Rate & Month-End Forecasting
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from src.data_loader import load_analytical_view, filter_dataset
from src.ui import render_tailwind_header, render_insight_box
from src.analysis import calculate_run_rate_forecast

st.set_page_config(page_title="Forecasting | P054", page_icon="🔮", layout="wide")

render_tailwind_header(
    title="Predictive Run-Rate Velocity & Month-End Forecaster",
    subtitle="Forecasting full-month revenue realization based on elapsed daily trading pace and computing required catch-up velocity.",
    meta_tag="UI-05 Predictive Analytics"
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
    # Compute Forecast
    elapsed_trading_days = filtered_df["transaction_date"].dt.date.nunique()
    total_month_days = 30
    
    fc = calculate_run_rate_forecast(filtered_df, elapsed_days=elapsed_trading_days, total_month_days=total_month_days)

    st.markdown("### 🎯 **Month-End Quota Attainment Projections (OQ-10)**")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Elapsed Daily Velocity", f"₹{fc['daily_run_rate']:,.0f} / day", f"{fc['elapsed_days']} of {fc['total_month_days']} Days")
    with c2:
        st.metric("Projected Month-End Sales", f"₹{fc['projected_month_end']:,.0f}", f"Target: ₹{fc['monthly_target']:,.0f}")
    with c3:
        st.metric("Projected Quota Deficit", f"{'+' if fc['projected_deficit']>=0 else ''}₹{fc['projected_deficit']:,.0f}", f"{(fc['projected_deficit']/fc['monthly_target'])*100:+.1f}% Gap")
    with c4:
        st.metric("Required Catch-Up Rate", f"₹{fc['required_daily_pace']:,.0f} / day", f"{fc['remaining_days']} Days Remaining")

    if fc["is_deficit"]:
        render_insight_box(
            title=f"Critical Month-End Deficit Projected: Shortfall of ₹{abs(fc['projected_deficit']):,.0f}",
            message=f"At current daily velocity (₹{fc['daily_run_rate']:,.0f}/day), the network will fail to reach target quotas. To achieve 100% quota by Day 30, daily sales velocity must accelerate by {((fc['required_daily_pace']-fc['daily_run_rate'])/fc['daily_run_rate'])*100:+.1f}% to ₹{fc['required_daily_pace']:,.0f}/day across remaining {fc['remaining_days']} trading days.",
            alert_type="danger"
        )
    else:
        render_insight_box(
            title="Surplus Run-Rate Projected",
            message=f"Current daily sales pace is on track to exceed monthly budgeted quotas with a projected surplus of ₹{fc['projected_deficit']:,.0f}.",
            alert_type="success"
        )

    st.markdown("---")

    # Interactive Scenario Simulation
    st.markdown("### 🎛️ **Scenario Simulation: Interactive Daily Catch-Up Pace Slider**")
    
    boost_pct = st.slider("Simulate Promotional Revenue Acceleration (%) for Remaining Days:", min_value=-20, max_value=50, value=15, step=5)
    
    simulated_daily_pace = fc["daily_run_rate"] * (1.0 + boost_pct / 100.0)
    simulated_additional_sales = simulated_daily_pace * fc["remaining_days"]
    simulated_total_sales = fc["cumulative_sales"] + simulated_additional_sales
    simulated_deficit = simulated_total_sales - fc["monthly_target"]

    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        st.metric("Simulated Pace", f"₹{simulated_daily_pace:,.0f} / day", f"{boost_pct:+d}% Acceleration")
    with sc2:
        st.metric("Simulated Month Total", f"₹{simulated_total_sales:,.0f}")
    with sc3:
        st.metric("Simulated Outcome", f"{'+' if simulated_deficit>=0 else ''}₹{simulated_deficit:,.0f}", delta_color="normal" if simulated_deficit>=0 else "inverse")

    # Visual Horizon Comparison
    fig_proj = go.Figure()
    fig_proj.add_trace(go.Bar(
        x=["Planned Monthly Target", "Status-Quo Run-Rate Projection", "Simulated Scenario (+{}%)".format(boost_pct)],
        y=[fc["monthly_target"], fc["projected_month_end"], simulated_total_sales],
        marker_color=["#1E3A8A", "#EF4444" if fc["is_deficit"] else "#10B981", "#3B82F6"],
        text=[f"₹{fc['monthly_target']:,.0f}", f"₹{fc['projected_month_end']:,.0f}", f"₹{simulated_total_sales:,.0f}"],
        textposition="outside"
    ))
    fig_proj.update_layout(
        title="Full-Month Revenue Comparison: Target vs Status Quo vs Simulated Scenario (INR)",
        yaxis_title="Total Month Sales (₹)",
        margin=dict(l=20, r=20, t=40, b=20),
        height=380,
        plot_bgcolor="white"
    )
    st.plotly_chart(fig_proj, use_container_width=True)
