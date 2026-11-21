import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def style_header_cell(cell, text, width=None):
    if width:
        cell.width = width
    set_cell_background(cell, "1E3A8A")
    set_cell_margins(cell, top=120, bottom=120, left=130, right=130)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

def style_data_cell(cell, text, is_even=False, width=None, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    if width:
        cell.width = width
    bg_color = "F8FAFC" if is_even else "FFFFFF"
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=80, bottom=80, left=110, right=110)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    run.font.name = "Calibri"
    run.font.size = Pt(9)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = RGBColor(30, 41, 59)

def create_test_cases_doc():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Header
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    r_title = title_p.add_run("TEST SUITE & QUALITY ASSURANCE SPECIFICATION")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(14)
    r_sub = sub_p.add_run("Project P054: Daily Sales Actual vs Target Analysis | Functional (TC-F) & Technical (TC-T) Test Matrix")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # Metadata Box
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table)
    meta_data = [
        ("Candidate Name / Roll No:", "Harshit | University Roll No: 2500320140058 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Project Mentor:", "Dinesh Sereno | Industry Data Engineering Lead"),
        ("Test Harness & Framework:", "pytest 9.2.0, pandas, sqlite3, Streamlit Component Test Suite"),
        ("Specification Scope:", "16 Functional Test Cases (TC-F-01..16), 21 Technical Unit Tests (TC-T-01..21), 100% BR Coverage")
    ]
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        style_data_cell(row.cells[0], k, is_even=(row_idx%2==1), width=Inches(2.2), bold=True)
        style_data_cell(row.cells[1], v, is_even=(row_idx%2==1), width=Inches(4.3))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Testing Strategy & Test Pyramid
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Quality Assurance Strategy & Testing Pyramid")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(30, 58, 138)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.add_run(
        "Project P054 enforces an industry-grade verification framework following the Test Pyramid paradigm. "
        "Every mathematical formula, database constraint, and user interface behavior is covered by deterministic test cases:\n\n"
        "1. Layer 1 — Technical & Unit Tests (TC-T-01 to TC-T-21): Implemented in Python with pytest. These tests execute directly "
        "against the SQLite database engine, pandas analytical views, and calculation routines (variance formulas, waterfall decomposition "
        "identity, OLS linear regression, and Z-score outlier flagging).\n"
        "2. Layer 2 — Functional & System Tests (TC-F-01 to TC-F-16): Derived directly from the Gherkin acceptance criteria in Step 11. "
        "These test end-to-end user workflows: cascading sidebar filter controls, Tailwind KPI card threshold color triggers, sorting, "
        "navigation, empty state handling, and data export downloads.\n"
        "3. Performance Boundary: Total analytical view query latency is verified to remain strictly below 1.0 second (< 3.0s budget)."
    )

    # Section 2: Functional Test Suite (TC-F)
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Functional Test Suite (TC-F-01 to TC-F-16)")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(30, 58, 138)

    func_tests = [
        ("TC-F-01", "BR-01, US-05", "UI-01", "Dashboard open, all 300 rows active", "Change Date Range to 'Nov 05 - Nov 10'", "Date range: 2026-11-05 to 2026-11-10", "KPI cards and tables recalculate showing only transactions within selected 6 days", "High", "Not Run"),
        ("TC-F-02", "BR-02, US-05", "UI-01", "Region filter set to 'All'", "Select Region = 'North' in sidebar", "Region: 'North'", "Store dropdown cascades to show only 4 North stores (Connaught, Cyber City, Sector 17, Hazratganj)", "High", "Not Run"),
        ("TC-F-03", "BR-01, US-05", "UI-01", "Overview page loaded", "Select Category = 'Electronics'", "Category: CAT-01", "Actual and Target revenue KPI cards reflect only Electronics sales with matching target quotas", "Medium", "Not Run"),
        ("TC-F-04", "BR-01, US-04", "UI-01", "Overview loaded with default filters", "Inspect 5 Tailwind KPI card values", "Full dataset (Nov 01-20)", "Cards display Actual Revenue, Target Revenue, Dollar Variance, % Achievement, Realized Margin with badges", "High", "Not Run"),
        ("TC-F-05", "BR-02, US-06", "UI-01", "Overview page open", "Click 'Achievement Rate %' column header", "All 15 stores", "Table sorts ascending; stores <85% (e.g. Fraser Road) show red 'CRITICAL DEFICIT' alert badge", "High", "Not Run"),
        ("TC-F-06", "BR-08, US-10", "UI-01", "Store STR-001 selected", "Observe ATV Angular Gauge", "STR-001 sales data", "Gauge needle reflects realized ATV (INR 4,333) against benchmark threshold (INR 3,200) in green zone", "Medium", "Not Run"),
        ("TC-F-07", "BR-04, US-08", "UI-02", "Navigate to Trends & Seasonality", "Inspect Day-of-Week Heatmap", "300 sales records", "Chronological Mon-Sun matrix renders; Tuesday & Wednesday cells render in amber/red (<90% achievement)", "High", "Not Run"),
        ("TC-F-08", "BR-06, US-13", "UI-02", "Trends page open", "Inspect MTD Burn-Up line chart", "20 trading days", "Dual lines plot cumulative sales vs target pace; shortfall area between curves is shaded in light red", "High", "Not Run"),
        ("TC-F-09", "BR-03, US-07", "UI-03", "Navigate to Store & Category Compare", "Inspect Category Bar Chart", "5 merchandise categories", "Grouped horizontal bars display Actual vs Target revenue side-by-side with exact INR tooltips", "High", "Not Run"),
        ("TC-F-10", "BR-05, US-09", "UI-03", "Compare page open", "Filter by Store STR-001; inspect Staff table", "STR-001 staff", "Associates displayed with sales vs quota; any associate <85% quota has 'Coaching Required' = YES in red", "High", "Not Run"),
        ("TC-F-11", "BR-09, US-11", "UI-04", "Navigate to Root Cause Variance", "Observe Waterfall Decomposition chart", "Overall network or single store", "Target bar steps down via Volume Effect ($) and Price Effect ($) to equal Actual Revenue exactly", "High", "Not Run"),
        ("TC-F-12", "BR-07, US-12", "UI-04", "Root Cause page open", "Inspect Discount vs Margin Scatter Plot", "All sales transactions", "Scatter points plotted with downward OLS regression line; Pearson r displayed (< -0.30) with p < 0.001", "High", "Not Run"),
        ("TC-F-13", "BR-10, US-14", "UI-05", "Navigate to Predictive Forecasting", "Inspect Month-End Forecast Card", "20 elapsed trading days", "Card displays Daily Run-Rate (INR 421K), Projected Total (12.63M), and Projected Deficit (-567K)", "High", "Not Run"),
        ("TC-F-14", "BR-11, US-15", "UI-06", "Navigate to Data Governance & Audit", "Check Anomaly Detection Table", "Threshold Z-Score = 3.0", "Exactly 3 flash surge sales records displayed with |Z| > 3.0; toggle enables/disables outlier exclusion", "Medium", "Not Run"),
        ("TC-F-15", "BR-01, US-05", "UI-01", "Overview page open", "Set filter to impossible criteria (no data)", "Store = 'STR-015' & Cat = 'CAT-01'", "UI displays friendly warning: 'No transactions match the selected filter criteria. Please reset filters.'", "Medium", "Not Run"),
        ("TC-F-16", "BR-01, US-02", "UI-06", "Data Governance page open", "Click 'Download Cleaned Dataset (Excel)'", "300 cleaned records", "Browser downloads 'P054_Cleaned_Daily_Sales.xlsx' containing 6 sheets with zero corrupted cells", "High", "Not Run")
    ]

    ft_table = doc.add_table(rows=len(func_tests)+1, cols=9)
    ft_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ft_table)
    ft_headers = ["Test ID", "Req / US", "UI", "Pre-condition", "Execution Steps", "Test Data", "Expected Result", "Pri", "Status"]
    for i, h in enumerate(ft_headers):
        style_header_cell(ft_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(func_tests):
        row = ft_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx in [0, 7]))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 3: Technical Test Suite (TC-T)
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Technical Unit Test Suite (TC-T-01 to TC-T-21 via pytest)")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(30, 58, 138)

    tech_tests = [
        ("TC-T-01", "tests/test_data.py", "test_tc_t_01_db_file_exists", "Verify SQLite app.db exists on disk and size > 50KB", "File exists, size >= 200KB"),
        ("TC-T-02", "tests/test_data.py", "test_tc_t_02_dim_stores_counts", "Query dim_stores row count and region domain", "Exactly 15 stores, valid regions"),
        ("TC-T-03", "tests/test_data.py", "test_tc_t_03_dim_categories_counts", "Query dim_categories row count and names", "Exactly 5 categories"),
        ("TC-T-04", "tests/test_data.py", "test_tc_t_04_dim_staff_counts", "Query dim_staff count and positive quotas", "Exactly 75 staff, all quotas > 0"),
        ("TC-T-05", "tests/test_data.py", "test_tc_t_05_fact_targets_counts", "Query fact_daily_targets count and positive target", "300 rows, target_revenue > 0"),
        ("TC-T-06", "tests/test_data.py", "test_tc_t_06_fact_sales_counts", "Query fact_daily_sales count and revenue range", "300 rows, net_revenue >= 0"),
        ("TC-T-07", "tests/test_data.py", "test_tc_t_07_analysis_view_40_columns", "Query vw_daily_sales_target_analysis column count", "Exactly 40 columns matching spec"),
        ("TC-T-08", "tests/test_data.py", "test_tc_t_08_primary_key_uniqueness", "Check duplicate primary keys across all 5 tables", "0 duplicate PK records across all tables"),
        ("TC-T-09", "tests/test_data.py", "test_tc_t_09_foreign_key_integrity", "Verify zero orphan records in sales and targets", "0 orphan foreign keys (100% integrity)"),
        ("TC-T-10", "tests/test_data.py", "test_tc_t_10_cleaning_routine", "Impute null discounts & deduplicate raw sales CSV", "Nulls imputed to 0.0, 303 rows -> 300"),
        ("TC-T-11", "tests/test_analysis.py", "test_tc_t_11_variance_and_achievement", "Evaluate variance ($) and achievement (%) on micro fixture", "Variance = -5000, Ach = 95.0%"),
        ("TC-T-12", "tests/test_analysis.py", "test_tc_t_12_zero_division_guard", "Verify denominator guard when target_revenue = 0", "Returns 0.0, no ZeroDivisionError"),
        ("TC-T-13", "tests/test_analysis.py", "test_tc_t_13_variance_decomposition", "Verify Volume Effect ($) + Price Effect ($) == Total Variance", "Absolute difference < 0.50 for all rows"),
        ("TC-T-14", "tests/test_analysis.py", "test_tc_t_14_gross_margin_realization", "Evaluate gross profit and margin % formula", "Profit = 20,000, Margin = 40.0%"),
        ("TC-T-15", "tests/test_analysis.py", "test_tc_t_15_month_end_run_rate", "Evaluate daily velocity, forecast total, and catch-up rate", "Forecast = 12.63M, Deficit = -567.7K"),
        ("TC-T-16", "tests/test_analysis.py", "test_tc_t_16_statistical_anomaly", "Compute Z-scores on actual sales and identify |Z| > 3.0", "Exactly 3 statistical outliers flagged"),
        ("TC-T-17", "tests/test_analysis.py", "test_tc_t_17_discount_vs_margin", "Compute Pearson correlation between discount % and margin %", "r < -0.30 confirming margin erosion"),
        ("TC-T-18", "tests/test_app.py", "test_tc_t_18_ui_module_import", "Import src/ui.py and verify exported functions", "render_tailwind_header, kpi, insight exist"),
        ("TC-T-19", "tests/test_app.py", "test_tc_t_19_streamlit_config", "Verify .streamlit/config.toml theme settings", "primaryColor = #1E3A8A"),
        ("TC-T-20", "tests/test_app.py", "test_tc_t_20_folder_structure", "Verify data, docs, src, tests, .streamlit folders", "All mandatory directories exist"),
        ("TC-T-21", "tests/test_app.py", "test_tc_t_21_query_performance", "Benchmark query time for 40-col view from SQLite", "Query execution completes in < 1.0 second")
    ]

    tt_table = doc.add_table(rows=len(tech_tests)+1, cols=5)
    tt_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tt_table)
    tt_headers = ["Test ID", "Test Script File", "Test Function Name", "Verification Target", "Expected Assertion"]
    for i, h in enumerate(tt_headers):
        style_header_cell(tt_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(tech_tests):
        row = tt_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx in [0, 2]))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 4: Test Execution Guide
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(12)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. Test Execution Guidelines & Command Line Procedures")
    r_h4.font.name = "Calibri"
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(30, 58, 138)

    p_exec = doc.add_paragraph()
    p_exec.paragraph_format.space_after = Pt(8)
    p_exec.add_run(
        "To execute the automated technical test harness, run the following commands from the project root:\n\n"
        "• Execute All Technical Tests:\n"
        "  pytest -v\n\n"
        "• Execute Specific Module Test Suites:\n"
        "  pytest tests/test_data.py -v\n"
        "  pytest tests/test_analysis.py -v\n"
        "  pytest tests/test_app.py -v\n\n"
        "• Generate Code Coverage Report:\n"
        "  pytest --cov=src --cov=data tests/\n\n"
        "Continuous Integration Gate: All 21 technical unit tests must pass with 100% success rate before merging pull requests or releasing builds to Streamlit Community Cloud."
    )

    doc_path = os.path.abspath("docs/13_Test_Cases.docx")
    doc.save(doc_path)
    print(f"Generated {doc_path} successfully.")

def update_rtm_step13():
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]

    # Final Mapping of BR ID to Test Cases (Column 12)
    test_mapping = {
        "BR-01": "TC-F-01, TC-F-03, TC-F-04 / TC-T-01, TC-T-05, TC-T-06, TC-T-07, TC-T-11",
        "BR-02": "TC-F-02, TC-F-05 / TC-T-02, TC-T-08, TC-T-11",
        "BR-03": "TC-F-09 / TC-T-03, TC-T-09, TC-T-11",
        "BR-04": "TC-F-07 / TC-T-07, TC-T-11",
        "BR-05": "TC-F-10 / TC-T-04, TC-T-09",
        "BR-06": "TC-F-08 / TC-T-07, TC-T-15",
        "BR-07": "TC-F-12 / TC-T-14, TC-T-17",
        "BR-08": "TC-F-06 / TC-T-05, TC-T-07, TC-T-13",
        "BR-09": "TC-F-11 / TC-T-13",
        "BR-10": "TC-F-13 / TC-T-15",
        "BR-11": "TC-F-14 / TC-T-16",
        "BR-12": "Out of Scope (Phase 2)"
    }

    font_calibri = Font(name="Calibri", size=9)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for row_idx in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row_idx, column=3).value
        if br_id in test_mapping:
            test_val = test_mapping[br_id]
            cell_test = ws.cell(row=row_idx, column=12)
            cell_test.value = test_val
            cell_test.font = font_calibri
            cell_test.alignment = align_left

    wb.save(rtm_path)
    print(f"Updated {rtm_path} with final Test Case mappings successfully.")

if __name__ == "__main__":
    create_test_cases_doc()
    update_rtm_step13()
