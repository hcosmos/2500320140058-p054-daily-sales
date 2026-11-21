"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Technical Test Suite: App Architecture & Performance (tests/test_app.py)
Author: Harshit (Roll No: 2500320140058) | Mentor: Dinesh Sereno
"""

import os
import sys
import time
import sqlite3
import pandas as pd
import pytest

def test_tc_t_18_ui_module_import():
    """TC-T-18: Verify src/ui.py module exists and exports required UI rendering helpers."""
    pytest.importorskip("streamlit", reason="Streamlit will be installed during Step 15 build phase")
    sys.path.insert(0, os.path.abspath("."))
    from src import ui
    assert hasattr(ui, "render_tailwind_header"), "Missing render_tailwind_header in src/ui.py"
    assert hasattr(ui, "render_kpi_cards"), "Missing render_kpi_cards in src/ui.py"
    assert hasattr(ui, "render_insight_box"), "Missing render_insight_box in src/ui.py"

def test_tc_t_19_streamlit_config_toml_theme():
    """TC-T-19: Verify .streamlit/config.toml exists and specifies correct primaryColor."""
    config_path = os.path.abspath(".streamlit/config.toml")
    assert os.path.exists(config_path), ".streamlit/config.toml is missing"
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    assert 'primaryColor = "#1E3A8A"' in content, "Incorrect primary color in config.toml"
    assert 'backgroundColor = "#F8FAFC"' in content, "Incorrect background color in config.toml"

def test_tc_t_20_folder_structure_integrity():
    """TC-T-20: Verify mandatory folder structure exists."""
    for folder in ["data", "docs", "src", "tests", ".streamlit"]:
        assert os.path.isdir(folder), f"Directory {folder} does not exist"

def test_tc_t_21_query_performance_under_3_seconds():
    """TC-T-21: Benchmark query time: analytical view loading must complete in < 1.0 second."""
    db_path = os.path.abspath("data/app.db")
    start_time = time.time()
    conn = sqlite3.connect(db_path)
    df = pd.read_sql("SELECT * FROM vw_daily_sales_target_analysis", conn)
    conn.close()
    duration = time.time() - start_time
    assert len(df) == 300
    assert duration < 1.0, f"Query took {duration:.3f}s, exceeding 1.0s limit"
