"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Data Ingestion, Cleaning & In-Memory Caching Module (src/data_loader.py)
Author: Harshit (Roll No: 2500320140058) | Mentor: Dinesh Sereno
"""

import os
import sqlite3
import pandas as pd
import numpy as np

try:
    import streamlit as st
    cache_decorator = st.cache_data
except ImportError:
    # Fallback dummy decorator for testing without streamlit
    def cache_decorator(func):
        return func

DB_PATH = os.path.abspath("data/app.db")
RAW_SALES_PATH = os.path.abspath("data/fact_daily_sales.csv")

def get_db_connection():
    """Returns an active SQLite connection with Foreign Keys enabled."""
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Please run src/generate_data.py first.")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def clean_raw_sales_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Sanitizes raw transactional feed:
    1. Imputes NaN values in discount_amount with 0.0 (unrecorded markdowns).
    2. Drops duplicate records by primary key (sales_id).
    3. Recomputes realized_atv and gross_profit for sanitized integrity.
    """
    df = df_raw.copy()
    
    # 1. Impute missing discounts
    df["discount_amount"] = df["discount_amount"].fillna(0.0)
    
    # 2. Deduplicate primary key
    df = df.drop_duplicates(subset=["sales_id"], keep="first")
    
    # 3. Defensive clipping of financial metrics
    df["net_revenue"] = df["net_revenue"].clip(lower=0.0)
    df["gross_profit"] = df["net_revenue"] - df["cost_of_goods_sold"]
    df["realized_atv"] = df["net_revenue"] / df["order_count"].replace(0, 1)
    
    return df

@cache_decorator
def load_analytical_view() -> pd.DataFrame:
    """
    Loads and caches the 40-column analytical view from SQLite.
    Execution time: < 0.5s for 300 rows.
    """
    conn = get_db_connection()
    query = "SELECT * FROM vw_daily_sales_target_analysis ORDER BY transaction_date ASC, store_id ASC;"
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    # Clean up and enforce datatypes
    df["transaction_date"] = pd.to_datetime(df["transaction_date"])
    df["actual_orders"] = df["actual_orders"].astype(int)
    df["target_orders"] = df["target_orders"].astype(int)
    df["units_sold"] = df["units_sold"].astype(int)
    
    # Ensure zero-division safety on percentage columns
    df["achievement_rate_pct"] = np.where(
        df["target_revenue"] > 0,
        (df["actual_net_revenue"] / df["target_revenue"]) * 100.0,
        0.0
    )
    df["pct_variance"] = np.where(
        df["target_revenue"] > 0,
        ((df["actual_net_revenue"] - df["target_revenue"]) / df["target_revenue"]) * 100.0,
        0.0
    )
    df["discount_pct"] = np.where(
        df["gross_revenue"] > 0,
        (df["discount_amount"] / df["gross_revenue"]),
        0.0
    )
    df["realized_margin_pct"] = np.where(
        df["actual_net_revenue"] > 0,
        (df["gross_profit"] / df["actual_net_revenue"]) * 100.0,
        0.0
    )
    
    return df

@cache_decorator
def load_dimension_tables():
    """Loads reference dimension dataframes for stores, categories, and staff."""
    conn = get_db_connection()
    stores = pd.read_sql("SELECT * FROM dim_stores ORDER BY store_id", conn)
    categories = pd.read_sql("SELECT * FROM dim_categories ORDER BY category_id", conn)
    staff = pd.read_sql("SELECT * FROM dim_staff ORDER BY staff_id", conn)
    conn.close()
    return stores, categories, staff

def filter_dataset(df: pd.DataFrame, date_range=None, regions=None, formats=None, stores=None, categories=None):
    """
    Cascading filter applicator for Streamlit dashboard views.
    Returns filtered subset of the 40-column view.
    """
    filtered = df.copy()
    
    if date_range and len(date_range) == 2:
        start_d, end_d = date_range
        filtered = filtered[(filtered["transaction_date"].dt.date >= start_d) & 
                            (filtered["transaction_date"].dt.date <= end_d)]
        
    if regions and "All Regions" not in regions:
        filtered = filtered[filtered["region"].isin(regions)]
        
    if formats and "All Formats" not in formats:
        filtered = filtered[filtered["store_format"].isin(formats)]
        
    if stores and "All Stores" not in stores:
        filtered = filtered[filtered["store_name"].isin(stores)]
        
    if categories and "All Categories" not in categories:
        filtered = filtered[filtered["category_name"].isin(categories)]
        
    return filtered
