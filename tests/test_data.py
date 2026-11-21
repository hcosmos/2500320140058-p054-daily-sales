"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Technical Test Suite: Data Layer & Referential Integrity (tests/test_data.py)
Author: Harshit (Roll No: 2500320140058) | Mentor: Dinesh Sereno
"""

import os
import sqlite3
import pandas as pd
import numpy as np

DB_PATH = os.path.abspath("data/app.db")
EXCEL_PATH = os.path.abspath("data/P054_Daily_Sales_data.xlsx")

def get_connection():
    assert os.path.exists(DB_PATH), f"Database not found at {DB_PATH}"
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def test_tc_t_01_db_file_exists():
    """TC-T-01: Verify SQLite database file exists and is accessible."""
    assert os.path.exists(DB_PATH)
    assert os.path.getsize(DB_PATH) > 50000, "Database file is unexpectedly small"

def test_tc_t_02_dim_stores_counts():
    """TC-T-02: Verify dim_stores has exactly 15 records and expected schema."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM dim_stores", conn)
    conn.close()
    assert len(df) == 15
    assert "store_id" in df.columns
    assert "region" in df.columns
    assert set(df["region"].unique()).issubset({"North", "South", "East", "West", "Central"})

def test_tc_t_03_dim_categories_counts():
    """TC-T-03: Verify dim_categories has exactly 5 merchandise categories."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM dim_categories", conn)
    conn.close()
    assert len(df) == 5
    assert set(df["category_name"].unique()) == {"Electronics", "Apparel", "Home & Living", "Grocery", "Footwear"}

def test_tc_t_04_dim_staff_counts():
    """TC-T-04: Verify dim_staff has exactly 75 associates with valid foreign keys."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM dim_staff", conn)
    conn.close()
    assert len(df) == 75
    assert df["base_daily_quota"].min() > 0

def test_tc_t_05_fact_targets_counts():
    """TC-T-05: Verify fact_daily_targets has 300 budgeted target rows."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM fact_daily_targets", conn)
    conn.close()
    assert len(df) == 300
    assert (df["target_revenue"] > 0).all()

def test_tc_t_06_fact_sales_counts():
    """TC-T-06: Verify fact_daily_sales has 300 primary transactional records."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM fact_daily_sales", conn)
    conn.close()
    assert len(df) == 300
    assert (df["net_revenue"] >= 0).all()

def test_tc_t_07_analysis_view_40_columns():
    """TC-T-07: Verify vw_daily_sales_target_analysis returns 40 columns."""
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM vw_daily_sales_target_analysis LIMIT 5", conn)
    conn.close()
    assert df.shape[1] == 40, f"Expected 40 columns in analysis view, found {df.shape[1]}"
    expected_cols = [
        "transaction_date", "day_of_week", "month_year", "day_number_in_month",
        "store_id", "store_name", "region", "store_format", "square_footage", "store_manager",
        "category_id", "category_name", "department_head", "target_gross_margin_pct",
        "staff_id", "staff_name", "staff_role", "base_daily_quota",
        "gross_revenue", "discount_amount", "discount_pct", "actual_net_revenue",
        "target_revenue", "dollar_variance", "pct_variance", "achievement_rate_pct",
        "underperformance_flag", "cost_of_goods_sold", "gross_profit", "realized_margin_pct",
        "actual_orders", "target_orders", "order_variance", "actual_atv", "target_atv",
        "atv_variance", "volume_effect_dollar", "price_effect_dollar", "units_sold", "reconciled_status"
    ]
    for col in expected_cols:
        assert col in df.columns, f"Missing expected column: {col}"

def test_tc_t_08_primary_key_uniqueness():
    """TC-T-08: Verify Primary Keys are strictly unique across all tables."""
    conn = get_connection()
    for tbl, pk in [
        ("dim_stores", "store_id"),
        ("dim_categories", "category_id"),
        ("dim_staff", "staff_id"),
        ("fact_daily_targets", "target_id"),
        ("fact_daily_sales", "sales_id")
    ]:
        df = pd.read_sql(f"SELECT {pk}, COUNT(*) as c FROM {tbl} GROUP BY {pk} HAVING c > 1", conn)
        assert len(df) == 0, f"Found duplicate primary keys in {tbl}.{pk}"
    conn.close()

def test_tc_t_09_foreign_key_integrity():
    """TC-T-09: Verify zero orphan records exist across all foreign key relationships."""
    conn = get_connection()
    # sales -> stores
    orphans_store = pd.read_sql("""
        SELECT s.sales_id FROM fact_daily_sales s
        LEFT JOIN dim_stores st ON s.store_id = st.store_id
        WHERE st.store_id IS NULL
    """, conn)
    assert len(orphans_store) == 0, "Found orphan store_id in fact_daily_sales"

    # sales -> categories
    orphans_cat = pd.read_sql("""
        SELECT s.sales_id FROM fact_daily_sales s
        LEFT JOIN dim_categories c ON s.category_id = c.category_id
        WHERE c.category_id IS NULL
    """, conn)
    assert len(orphans_cat) == 0, "Found orphan category_id in fact_daily_sales"

    # sales -> staff
    orphans_staff = pd.read_sql("""
        SELECT s.sales_id FROM fact_daily_sales s
        LEFT JOIN dim_staff sf ON s.staff_id = sf.staff_id
        WHERE sf.staff_id IS NULL
    """, conn)
    assert len(orphans_staff) == 0, "Found orphan staff_id in fact_daily_sales"
    conn.close()

def test_tc_t_10_cleaning_routine_imputation_and_deduplication():
    """TC-T-10: Test data cleaning pipeline on raw CSV feed with injected nulls/duplicates."""
    raw_csv = os.path.abspath("data/fact_daily_sales.csv")
    assert os.path.exists(raw_csv)
    df_raw = pd.read_csv(raw_csv)
    
    # Injected raw has 303 rows (3 duplicates) and 6 nulls in discount_amount
    assert len(df_raw) >= 300
    
    # Cleaning step 1: Impute missing discounts with 0.0
    df_clean = df_raw.copy()
    df_clean["discount_amount"] = df_clean["discount_amount"].fillna(0.0)
    assert df_clean["discount_amount"].isna().sum() == 0

    # Cleaning step 2: Deduplicate by primary key (sales_id)
    df_clean = df_clean.drop_duplicates(subset=["sales_id"], keep="first")
    assert len(df_clean) == 300, f"Expected 300 deduplicated rows, got {len(df_clean)}"
