import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=130, right=130):
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
    set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
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

def add_code_block(doc, title, code_text):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(2)
    r = p_title.add_run(title)
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_before = Pt(0)
    p_box.paragraph_format.space_after = Pt(8)
    p_box.paragraph_format.line_spacing = Pt(10)
    r_code = p_box.add_run(code_text)
    r_code.font.name = "Consolas"
    r_code.font.size = Pt(8)
    r_code.font.color.rgb = RGBColor(15, 23, 42)

def create_build_plan_doc():
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
    r_title = title_p.add_run("MASTER BUILD PLAN & SPRINT SPECIFICATION")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(14)
    r_sub = sub_p.add_run("Project P054: Daily Sales Actual vs Target Analysis | Task-by-Task Implementation Architecture")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # Metadata Box
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table)
    meta_data = [
        ("Candidate Name / Roll No:", "Harshit | Roll No: 2500320140058 (MCA 2nd Year DA, ABES)"),
        ("Project Mentor:", "Dinesh Sereno | Industry Data Engineering Lead"),
        ("Execution Blueprint:", "BUILD_PLAN.md & 14_BUILD_PLAN.docx | 12 Discrete Engineering Tasks (BUILD-01..12)"),
        ("Baseline Release & Date:", "Version 1.0 (Build Ready) | November 2026")
    ]
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        style_data_cell(row.cells[0], k, is_even=(row_idx%2==1), width=Inches(2.2), bold=True)
        style_data_cell(row.cells[1], v, is_even=(row_idx%2==1), width=Inches(4.3))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Executive Summary & Document Index
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Executive Summary & Document Artifacts Index")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(30, 58, 138)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.add_run(
        "Project P054 is an enterprise-grade retail sales performance analytics web application built to eliminate the "
        "48-to-72-hour reporting latency typically suffered by multi-unit retail operations. By ingesting normalized daily transactional "
        "and target quota data from an ACID-compliant SQLite data layer (data/app.db), the system computes monetary and percentage variances, "
        "identifies underperforming stores and merchandise lines, isolates footfall vs ticket deficits via algebraic waterfall decomposition, "
        "models promotional markdown margin erosion, and forecasts month-end revenue quotas using daily run-rate velocity. The user experience "
        "combines an in-memory caching architecture (st.cache_data) for sub-second query latency with Tailwind CSS components embedded in "
        "Streamlit for an executive-ready presentation."
    )

    doc_index = [
        ("Step 1", "01_Project_Understanding.docx", "Problem statement, 5 personas, 10 Operational Questions (OQ-01..10), SMART objectives"),
        ("Step 2", "02_HMW_Statements.docx", "12 stakeholder pain points, 11 How Might We statements, top 5 prioritized challenges"),
        ("Step 3", "03_Needs_Analysis.docx", "12 Business Requirements (BR-01..12), MoSCoW prioritization (58% Must), data feasibility"),
        ("Step 4", "04_Traceability_Matrix.xlsx", "Requirements Traceability Matrix (RTM) linking OQ -> HMW -> BR -> FR -> DB -> US -> UI -> TC"),
        ("Step 5", "05_BRD.docx", "Business Requirements Document, RACI matrix, 8 business rules, risk mitigation catalog"),
        ("Step 7", "07_TRD.docx", "Technical Requirements Document, 5-layer pipeline architecture, pinned stack, SQLite schema"),
        ("Step 8", "08_Database_Design.xlsx", "Constellation relational schema, 3 dimensions, 2 fact tables, 40-column analytical view"),
        ("Step 9", "09_ER_Diagram.docx / .png", "High-res ER diagram graphic (matplotlib), Mermaid DDL relationships, design critique"),
        ("Step 10", "data/ + generate_data.py", "Synthetic dataset generator (seed=42), SQLite app.db, 5 patterns, 3 cleaning issues"),
        ("Step 11", "11_User_Stories.docx", "15 User Stories across 5 Epics (55 Story Points), Gherkin acceptance criteria, 2-week plan"),
        ("Step 12", "12_UI_Framework.docx", "7 screen blueprints, ASCII wireframes, responsive breakpoints, Tailwind classes, charts"),
        ("Step 13", "13_Test_Cases.docx + tests/", "Quality assurance matrix: 16 Functional Tests (TC-F), 21 Technical Unit Tests (TC-T)")
    ]

    idx_table = doc.add_table(rows=len(doc_index)+1, cols=3)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(idx_table)
    idx_headers = ["Step", "Document Artifact", "Scope & Content Summary"]
    for i, h in enumerate(idx_headers):
        style_header_cell(idx_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(doc_index):
        row = idx_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 2: Final Folder Structure
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Target Directory & Repository Structure")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(30, 58, 138)

    tree_text = """c:\\Users\\Sieghart\\Desktop\\HCL Project\\
|-- app.py                           # Application Entry Point & Global Routing
|-- requirements.txt                 # Pinned Production Dependencies
|-- README.md                        # Project Setup, Architecture & Operation Guide
|-- BUILD_PLAN.md                    # Step-by-Step Engineering Execution Blueprint
|-- .streamlit/
|   `-- config.toml                  # Streamlit Theme & Server Configurations
|-- data/
|   |-- app.db                       # Primary SQLite Relational Database Engine
|   |-- P054_Daily_Sales_data.xlsx   # Master 6-Sheet Relational & View Excel Seed
|   |-- dim_stores.csv               # 15 Fictional Indian Retail Store Masters
|   |-- dim_categories.csv           # 5 Merchandise Category Dimension Masters
|   |-- dim_staff.csv                # 75 Sales Associate Dimension Masters
|   |-- fact_daily_targets.csv       # 300 Budgeted Daily Target Quota Records
|   |-- fact_daily_sales.csv         # 303 Raw Sales Records (with injected issues)
|   `-- vw_daily_sales_target_analysis.csv # 300 x 40 Cleaned Denormalized Export
|-- src/
|   |-- __init__.py
|   |-- generate_data.py             # Reproducible Synthetic Seed Generator (seed=42)
|   |-- data_loader.py               # Ingestion, Null Imputation, Deduplication & Caching
|   |-- analysis.py                  # Variance Metrics, Waterfall, OLS, Forecast & Z-Scores
|   `-- ui.py                        # Tailwind CSS Header, KPI Cards & Alert Callouts
|-- pages/
|   |-- 01_Overview.py               # UI-01: Executive Cockpit & Store Deficit Table
|   |-- 02_Trends_Seasonality.py     # UI-02: Day-of-Week Heatmap & Cumulative Trajectory
|   |-- 03_Compare.py                # UI-03: Category Breakdown & Staff Leaderboard
|   |-- 04_Root_Cause.py             # UI-04: Waterfall Decomposition & Margin Scatter
|   |-- 05_Forecasting.py            # UI-05: Run-Rate Velocity & Month-End Forecaster
|   |-- 06_Data_Governance.py        # UI-06: Data Quality Audit, Z-Score Outliers & Export
|   `-- 07_About.py                  # UI-07: Academic Credentials, Architecture & DDL
|-- tests/
|   |-- __init__.py
|   |-- test_data.py                 # TC-T-01..10: Schema, Keys, Integrity & Cleaning
|   |-- test_analysis.py             # TC-T-11..17: Math Formulas, Waterfall, OLS, Z-Scores
|   `-- test_app.py                  # TC-T-18..21: UI Components, Theme & Query Latency
`-- docs/                            # Formal Academic & Professional Documentation Suite"""

    add_code_block(doc, "Repository Architecture Layout", tree_text)

    # Section 3: Build Tasks (BUILD-01 to BUILD-12)
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Step-by-Step Engineering Build Tasks (BUILD-01 to BUILD-12)")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(30, 58, 138)

    tasks_data = [
        ("BUILD-01", "Environment Setup & Requirements Pinning", "requirements.txt", "FR-01", "pip install -r requirements.txt", "Exit code 0", "TC-T-20"),
        ("BUILD-02", "Data Ingestion, Sanitization & Caching", "src/data_loader.py", "US-01..03, FR-01", "Build load_analytical_view() with @st.cache_data", "Returns (300, 40) df", "TC-T-01..10, 21"),
        ("BUILD-03", "Analytical & Statistical Engine", "src/analysis.py", "US-11, 12, 14, 15", "Build KPI, waterfall, OLS regression, run-rate, Z-scores", "All 7 math tests pass", "TC-T-11..17"),
        ("BUILD-04", "Reusable Tailwind UI Components", "src/ui.py", "US-04, FR-02", "Export render_tailwind_header, kpi_cards, insight_box", "Import without errors", "TC-T-18"),
        ("BUILD-05", "Main Routing & Overview Page", "app.py, pages/01_Overview.py", "US-04..06, 10", "Shared sidebar state + Overview KPI cards & ranking table", "Visual check in browser", "TC-F-01, 02, 04, 05, 06"),
        ("BUILD-06", "Trends & Seasonality Page", "pages/02_Trends_Seasonality.py", "US-08, 13", "Day-of-Week heatmap + MTD cumulative burn-up curve", "Heatmap renders Mon-Sun", "TC-F-07, 08"),
        ("BUILD-07", "Store & Category Compare Page", "pages/03_Compare.py", "US-07, 09", "Paired category horizontal bar + staff quota leaderboard", "Staff filter by store", "TC-F-09, 10"),
        ("BUILD-08", "Root Cause & Elasticity Page", "pages/04_Root_Cause.py", "US-11, 12", "Volume vs Price waterfall + Discount vs margin scatter", "Waterfall identity holds", "TC-F-11, 12, TC-T-13, 17"),
        ("BUILD-09", "Predictive Forecasting Page", "pages/05_Forecasting.py", "US-14, FR-11", "Velocity run-rate cards + month-end deficit slider", "Catch-up rate updates", "TC-F-13, TC-T-15"),
        ("BUILD-10", "Data Governance & About Pages", "pages/06_Data_Gov, 07_About.py", "US-02, 15", "Z-score outlier table + Excel export + About credentials", "File downloads, Z>3.0", "TC-F-14..16, TC-T-16"),
        ("BUILD-11", "Full Test Suite Execution", "tests/*.py", "All US / FR / BR", "Execute pytest -v across all modules", "21 passed (100%)", "TC-T-01..21"),
        ("BUILD-12", "README & Deployment Packaging", "README.md, .gitignore", "Governance", "Write comprehensive README + Cloud deployment guide", "All markdown links work", "DoD Release Gate")
    ]

    task_table = doc.add_table(rows=len(tasks_data)+1, cols=7)
    task_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(task_table)
    t_headers = ["Task ID", "Task Title", "Files", "Req / US", "Action Prompt", "Manual Check", "Tests"]
    for i, h in enumerate(t_headers):
        style_header_cell(task_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(tasks_data):
        row = task_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 4: Definition of Done & Commands
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(12)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. Definition of Done & Command Reference")
    r_h4.font.name = "Calibri"
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(30, 58, 138)

    p_dod = doc.add_paragraph()
    p_dod.paragraph_format.space_after = Pt(6)
    p_dod.add_run(
        "Definition of Done (DoD) Criteria for Step 15 Build Phase:\n"
        "[x] 100% of 'Must Have' Business Requirements (BR-01, 02, 03, 06, 08, 09, 10) fully functional.\n"
        "[x] All 10 Operational Questions (OQ-01..10) answered with interactive Plotly visual charts.\n"
        "[x] All 21 Technical Unit Tests (pytest -v) pass with 100% success rate.\n"
        "[x] Analytical view query latency strictly below 1.0 second on local SQLite database.\n"
        "[x] Shared multi-filter sidebar persists criteria across all 7 pages via st.session_state.\n"
        "[x] Zero unhandled exceptions or broken tracebacks on edge filter selections.\n"
        "[x] Complete README.md and .gitignore committed to repository."
    )

    cmd_text = """# Developer Command Reference:
# 1. Install production dependencies
pip install -r requirements.txt

# 2. Re-generate synthetic data & SQLite database (seed=42)
python src/generate_data.py

# 3. Execute all automated technical unit tests
pytest -v

# 4. Launch the Streamlit multi-page application locally
streamlit run app.py"""

    add_code_block(doc, "Developer Commands", cmd_text)

    # Section 5: Master Viva Voce Preparation (10 Questions)
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(12)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run("5. Master Viva Voce Preparation (10 Likely Questions & High-Impact Answers)")
    r_h5.font.name = "Calibri"
    r_h5.font.size = Pt(14)
    r_h5.font.bold = True
    r_h5.font.color.rgb = RGBColor(30, 58, 138)

    viva_questions = [
        ("Q1: What business problem does P054 solve?",
         "It eliminates the 48-to-72-hour reporting latency in multi-unit retail by providing automated daily variance analysis, immediate deficit alerts (<85%), and actionable root-cause decomposition before month-end closure."),
        ("Q2: Why use a constellation schema with two fact tables?",
         "Planned target quotas (fact_daily_targets) and actual operational transactions (fact_daily_sales) operate at different grains. They are joined in vw_daily_sales_target_analysis on (transaction_date, store_id, category_id) to calculate daily variance."),
        ("Q3: Why pre-compute variance metrics inside a SQL View?",
         "Pre-computing variance metrics inside SQLite guarantees single-source-of-truth consistency across all 7 pages and leverages the C-based database engine for sub-second query latency."),
        ("Q4: Explain the mathematical identity behind your Waterfall Decomposition.",
         "Volume Effect = (Actual Orders - Target Orders) * Target ATV; Price Effect = (Actual ATV - Target ATV) * Actual Orders. Their sum algebraically equals Actual Revenue - Target Revenue, isolating footfall vs ticket drivers."),
        ("Q5: Where did you encounter Simpson's Paradox in this project?",
         "Across categories, discounts and margins appeared positively correlated because high-margin Apparel (48%) has high discounts (18%). Analyzing within categories revealed the true reality: discounts directly erode profit margins (r < -0.30)."),
        ("Q6: Why is st.session_state required for sidebar filters?",
         "Streamlit re-executes page scripts on user interaction. Without session state, navigating between pages resets filters. st.session_state preserves filter selections across all page navigations."),
        ("Q7: How did you implement Tailwind CSS in Streamlit?",
         "Lightweight HTML containers rendered via st.components.v1.html load the Tailwind CDN to display modern cards, badges, and alerts while native Streamlit and Plotly handle interactive visuals."),
        ("Q8: How does the run-rate model project month-end revenue?",
         "It calculates daily velocity (Cumulative Sales / Elapsed Days) and scales it to total month days (Velocity * 30), subtracting the monthly quota to determine projected deficits and required catch-up pace."),
        ("Q9: Why inject data quality issues into the synthetic dataset?",
         "Real POS feeds always contain imperfections. Injecting 6 nulls, 3 duplicates, and 3 outliers validates that our ETL cleaning logic and data governance audit page handle messy data properly."),
        ("Q10: What automated testing strategy was implemented?",
         "Following the Test Pyramid, we implemented 21 technical unit tests in pytest (schema, keys, cleaning, formulas, latency) and 16 functional acceptance tests derived from Gherkin user stories.")
    ]

    viva_table = doc.add_table(rows=len(viva_questions)+1, cols=2)
    viva_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(viva_table)
    v_headers = ["Viva Question & Topic", "Model Technical Answer"]
    for i, h in enumerate(v_headers):
        style_header_cell(viva_table.rows[0].cells[i], h)

    for r_idx, (q, a) in enumerate(viva_questions):
        row = viva_table.rows[r_idx+1]
        style_data_cell(row.cells[0], q, is_even=(r_idx%2==1), width=Inches(2.5), bold=True)
        style_data_cell(row.cells[1], a, is_even=(r_idx%2==1), width=Inches(4.0))

    doc_path = os.path.abspath("docs/14_BUILD_PLAN.docx")
    doc.save(doc_path)
    print(f"Generated {doc_path} successfully.")

if __name__ == "__main__":
    create_build_plan_doc()
