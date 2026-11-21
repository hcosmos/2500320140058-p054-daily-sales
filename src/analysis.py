"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Analytical, Statistical & Mathematical Computation Engine (src/analysis.py)
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import numpy as np
import pandas as pd
from scipy import stats

def calculate_kpi_summary(df: pd.DataFrame) -> dict:
    """
    Computes enterprise and slice-level KPI aggregates with safe zero guards.
    """
    if len(df) == 0:
        return {
            "actual_revenue": 0.0,
            "target_revenue": 0.0,
            "dollar_variance": 0.0,
            "pct_variance": 0.0,
            "achievement_rate": 0.0,
            "gross_profit": 0.0,
            "realized_margin_pct": 0.0,
            "actual_orders": 0,
            "target_orders": 0,
            "actual_atv": 0.0,
            "target_atv": 0.0,
            "status": "NO_DATA"
        }
        
    act_rev = float(df["actual_net_revenue"].sum())
    tgt_rev = float(df["target_revenue"].sum())
    gross_profit = float(df["gross_profit"].sum())
    
    act_orders = int(df["actual_orders"].sum())
    tgt_orders = int(df["target_orders"].sum())
    
    dollar_var = act_rev - tgt_rev
    pct_var = (dollar_var / tgt_rev * 100.0) if tgt_rev > 0 else 0.0
    ach_rate = (act_rev / tgt_rev * 100.0) if tgt_rev > 0 else 0.0
    margin_pct = (gross_profit / act_rev * 100.0) if act_rev > 0 else 0.0
    
    act_atv = (act_rev / act_orders) if act_orders > 0 else 0.0
    tgt_atv = (tgt_rev / tgt_orders) if tgt_orders > 0 else 0.0
    
    status = "ON_TRACK" if ach_rate >= 100.0 else ("WARNING" if ach_rate >= 85.0 else "CRITICAL")
    
    return {
        "actual_revenue": act_rev,
        "target_revenue": tgt_rev,
        "dollar_variance": dollar_var,
        "pct_variance": pct_var,
        "achievement_rate": ach_rate,
        "gross_profit": gross_profit,
        "realized_margin_pct": margin_pct,
        "actual_orders": act_orders,
        "target_orders": tgt_orders,
        "actual_atv": act_atv,
        "target_atv": tgt_atv,
        "status": status
    }

def decompose_variance(df: pd.DataFrame) -> dict:
    """
    Decomposes total revenue variance into Volume Effect ($) and Price/ATV Effect ($).
    Formula:
      Volume Effect = (Actual Orders - Target Orders) * Target ATV
      Price Effect  = (Actual ATV - Target ATV) * Actual Orders
      Total Modeled Variance = Volume Effect + Price Effect
    """
    if len(df) == 0:
        return {"volume_effect": 0.0, "price_effect": 0.0, "total_variance": 0.0}
        
    act_orders = df["actual_orders"].sum()
    tgt_orders = df["target_orders"].sum()
    act_rev = df["actual_net_revenue"].sum()
    tgt_rev = df["target_revenue"].sum()
    
    tgt_atv = (tgt_rev / tgt_orders) if tgt_orders > 0 else 0.0
    act_atv = (act_rev / act_orders) if act_orders > 0 else 0.0
    
    vol_effect = (act_orders - tgt_orders) * tgt_atv
    price_effect = (act_atv - tgt_atv) * act_orders
    total_modeled = vol_effect + price_effect
    
    return {
        "volume_effect": round(vol_effect, 2),
        "price_effect": round(price_effect, 2),
        "total_variance": round(act_rev - tgt_rev, 2),
        "modeled_variance": round(total_modeled, 2),
        "actual_orders": int(act_orders),
        "target_orders": int(tgt_orders),
        "actual_atv": round(act_atv, 2),
        "target_atv": round(tgt_atv, 2)
    }

def calculate_day_of_week_seasonality(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregates trading metrics across chronological weekdays (Monday to Sunday).
    """
    dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    grouped = df.groupby("day_of_week").agg(
        actual_revenue=("actual_net_revenue", "sum"),
        target_revenue=("target_revenue", "sum"),
        total_orders=("actual_orders", "sum"),
        avg_atv=("actual_atv", "mean")
    ).reindex(dow_order).reset_index()
    
    grouped["achievement_rate_pct"] = np.where(
        grouped["target_revenue"] > 0,
        (grouped["actual_revenue"] / grouped["target_revenue"]) * 100.0,
        0.0
    )
    grouped["dollar_variance"] = grouped["actual_revenue"] - grouped["target_revenue"]
    
    return grouped

def calculate_staff_leaderboard(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregates performance by sales associate with coaching alert threshold (<85%).
    """
    grouped = df.groupby(["staff_id", "staff_name", "store_name", "staff_role"]).agg(
        total_sales=("actual_net_revenue", "sum"),
        daily_quota=("base_daily_quota", "first"),
        days_active=("transaction_date", "nunique")
    ).reset_index()
    
    grouped["total_quota"] = grouped["daily_quota"] * grouped["days_active"]
    grouped["achievement_pct"] = np.where(
        grouped["total_quota"] > 0,
        (grouped["total_sales"] / grouped["total_quota"]) * 100.0,
        0.0
    )
    grouped["coaching_flag"] = np.where(grouped["achievement_pct"] < 85.0, "YES (Coaching Required)", "NO (On Track)")
    grouped = grouped.sort_values(by="achievement_pct", ascending=True)
    
    return grouped

def compute_discount_margin_elasticity(df: pd.DataFrame):
    """
    Fits Ordinary Least Squares (OLS) linear regression between discount % and margin %.
    Returns slope, intercept, Pearson r, p-value, and regression line points.
    """
    valid = df.dropna(subset=["discount_pct", "realized_margin_pct"])
    if len(valid) < 3:
        return {"slope": 0.0, "intercept": 0.0, "r_value": 0.0, "p_value": 1.0, "std_err": 0.0}
        
    x = valid["discount_pct"].values
    y = valid["realized_margin_pct"].values
    
    res = stats.linregress(x, y)
    
    return {
        "slope": res.slope,
        "intercept": res.intercept,
        "r_value": res.rvalue,
        "r_squared": res.rvalue ** 2,
        "p_value": res.pvalue,
        "std_err": res.stderr
    }

def calculate_run_rate_forecast(df: pd.DataFrame, elapsed_days: int = 20, total_month_days: int = 30) -> dict:
    """
    Calculates run-rate velocity and projects month-end total sales vs targets.
    """
    act_rev = df["actual_net_revenue"].sum()
    tgt_rev = df["target_revenue"].sum()
    
    # Scale budgeted target to full 30 days
    monthly_target = (tgt_rev / max(1, elapsed_days)) * total_month_days
    
    daily_run_rate = act_rev / max(1, elapsed_days)
    projected_month_end = daily_run_rate * total_month_days
    projected_deficit = projected_month_end - monthly_target
    
    remaining_days = max(1, total_month_days - elapsed_days)
    remaining_quota = max(0.0, monthly_target - act_rev)
    required_daily_pace = remaining_quota / remaining_days
    
    return {
        "elapsed_days": elapsed_days,
        "total_month_days": total_month_days,
        "remaining_days": remaining_days,
        "cumulative_sales": act_rev,
        "monthly_target": monthly_target,
        "daily_run_rate": daily_run_rate,
        "projected_month_end": projected_month_end,
        "projected_deficit": projected_deficit,
        "required_daily_pace": required_daily_pace,
        "is_deficit": projected_deficit < 0
    }

def detect_anomalies_zscore(df: pd.DataFrame, threshold: float = 3.0) -> pd.DataFrame:
    """
    Identifies statistical revenue spikes/dips where absolute Z-Score exceeds threshold.
    """
    out_df = df.copy()
    mean_rev = out_df["actual_net_revenue"].mean()
    std_rev = out_df["actual_net_revenue"].std()
    
    if std_rev == 0 or np.isnan(std_rev):
        out_df["z_score"] = 0.0
    else:
        out_df["z_score"] = (out_df["actual_net_revenue"] - mean_rev) / std_rev
        
    anomalies = out_df[out_df["z_score"].abs() >= threshold].sort_values(by="z_score", ascending=False)
    return anomalies
