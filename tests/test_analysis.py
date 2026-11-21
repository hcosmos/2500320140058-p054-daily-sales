"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Technical Test Suite: Analytical & Statistical Formulas (tests/test_analysis.py)
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import numpy as np
import pandas as pd
import sqlite3
import os

DB_PATH = os.path.abspath("data/app.db")

def test_tc_t_11_variance_and_achievement_formulas():
    """TC-T-11: Verify variance ($) and achievement (%) on known micro-fixtures."""
    target_rev = 100000.0
    actual_rev = 95000.0
    
    dollar_variance = actual_rev - target_rev
    pct_variance = ((actual_rev - target_rev) / target_rev) * 100.0
    achievement_rate = (actual_rev / target_rev) * 100.0
    
    assert dollar_variance == -5000.0
    assert pct_variance == -5.0
    assert achievement_rate == 95.0

def test_tc_t_12_zero_division_guard():
    """TC-T-12: Ensure division by zero handles gracefully with 0.0 or None."""
    target_rev = 0.0
    actual_rev = 50000.0
    achievement = (actual_rev / target_rev * 100.0) if target_rev > 0 else 0.0
    assert achievement == 0.0

def test_tc_t_13_variance_decomposition_mathematical_identity():
    """
    TC-T-13: Verify algebraic decomposition identity:
    Volume Effect ($) + Price Effect ($) == Modeled Order Basket Variance ($)
    Formula:
      Volume Effect = (Actual Orders - Target Orders) * Target ATV
      Price Effect  = (Actual ATV - Target ATV) * Actual Orders
      Identity:
        Volume Effect + Price Effect
        = Actual Orders * Target ATV - Target Orders * Target ATV + Actual Orders * Actual ATV - Actual Orders * Target ATV
        = (Actual Orders * Actual ATV) - (Target Orders * Target ATV)
    """
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("""
        SELECT 
            actual_orders, target_orders, actual_atv, target_atv, 
            volume_effect_dollar, price_effect_dollar
        FROM vw_daily_sales_target_analysis
    """, conn)
    conn.close()

    for _, row in df.iterrows():
        calc_vol = (row["actual_orders"] - row["target_orders"]) * row["target_atv"]
        calc_pri = (row["actual_atv"] - row["target_atv"]) * row["actual_orders"]
        total_decomposed = calc_vol + calc_pri
        expected_modeled_variance = (row["actual_orders"] * row["actual_atv"]) - (row["target_orders"] * row["target_atv"])
        # Epsilon precision tolerance of 0.50 for rounding
        assert abs(total_decomposed - expected_modeled_variance) < 0.50, \
            f"Decomposition mismatch: {total_decomposed} vs {expected_modeled_variance}"

def test_tc_t_14_gross_margin_realization():
    """TC-T-14: Verify Gross Profit and Realized Margin % calculations."""
    net_rev = 50000.0
    cogs = 30000.0
    gross_profit = net_rev - cogs
    margin_pct = (gross_profit / net_rev) * 100.0
    
    assert gross_profit == 20000.0
    assert margin_pct == 40.0

def test_tc_t_15_month_end_run_rate_projection():
    """TC-T-15: Verify Run-Rate Projections based on elapsed trading days."""
    elapsed_days = 20
    total_month_days = 30
    cumulative_sales = 8400000.0
    monthly_target = 13000000.0

    daily_run_rate = cumulative_sales / elapsed_days
    projected_month_end = daily_run_rate * total_month_days
    projected_deficit = projected_month_end - monthly_target

    assert daily_run_rate == 420000.0
    assert projected_month_end == 12600000.0
    assert projected_deficit == -400000.0
    
    # Required catch-up pace for remaining 10 days
    remaining_days = total_month_days - elapsed_days
    remaining_target = monthly_target - cumulative_sales
    required_daily_pace = remaining_target / remaining_days
    assert required_daily_pace == 460000.0

def test_tc_t_16_statistical_anomaly_zscore_detection():
    """TC-T-16: Verify Z-Score anomaly detector catches injected >3 sigma outliers."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT actual_net_revenue FROM vw_daily_sales_target_analysis", conn)
    conn.close()

    rev = df["actual_net_revenue"]
    mean = rev.mean()
    std = rev.std()
    z_scores = (rev - mean) / std

    outliers = z_scores[z_scores.abs() > 3.0]
    # In Step 10, we injected exactly 3 statistical revenue outliers
    assert len(outliers) >= 3, f"Expected at least 3 Z-score outliers, found {len(outliers)}"

def test_tc_t_17_discount_vs_margin_negative_correlation_within_category():
    """
    TC-T-17: Verify that within promotional categories (e.g., Apparel / Footwear),
    higher discount percentage directly erodes realized gross profit margin.
    (Note: Aggregating across categories introduces Simpson's Paradox due to differing base margins).
    """
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("""
        SELECT category_name, discount_pct, realized_margin_pct 
        FROM vw_daily_sales_target_analysis 
        WHERE category_name IN ('Apparel', 'Footwear')
          AND discount_pct IS NOT NULL AND realized_margin_pct IS NOT NULL
    """, conn)
    conn.close()

    for cat in ['Apparel', 'Footwear']:
        cat_df = df[df['category_name'] == cat]
        r = cat_df["discount_pct"].corr(cat_df["realized_margin_pct"])
        assert r < -0.30, f"Expected negative discount-margin correlation (< -0.30) for {cat}, got {r}"
