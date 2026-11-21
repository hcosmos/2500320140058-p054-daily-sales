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

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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
    set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

def style_data_cell(cell, text, is_even=False, width=None, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    if width:
        cell.width = width
    bg_color = "F8FAFC" if is_even else "FFFFFF"
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = RGBColor(30, 41, 59)

def create_user_stories_doc():
    doc = Document()

    # Page Margins: 1 inch
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Header Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    r_title = title_p.add_run("USER STORIES & AGILE SPECIFICATION")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(16)
    r_sub = sub_p.add_run("Project P054: Daily Sales Actual vs Target Analysis | Agile Scrum Release 1.0")
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
        ("Document Release & Date:", "Version 1.0 | Sprint Ready Baseline | November 2026"),
        ("Artifact Scope:", "15 User Stories (55 Story Points), 5 Epics, Gherkin Criteria, Sprint Backlog")
    ]
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        style_data_cell(row.cells[0], k, is_even=(row_idx%2==1), width=Inches(2.2), bold=True)
        style_data_cell(row.cells[1], v, is_even=(row_idx%2==1), width=Inches(4.3))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Executive Summary & Agile Framework
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Executive Summary & Agile Delivery Framework")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(30, 58, 138)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.add_run(
        "This Agile Specification Document defines the 15 user stories comprising the Minimum Viable Product (MVP) "
        "for Project P054: Daily Sales Actual vs Target Analysis. These stories bridge the high-level business goals "
        "defined in the Business Requirements Document (BRD) and the architectural constraints of the Technical Requirements "
        "Document (TRD). Each story adopts the industry-standard Connextra template ('As a <persona>, I want <goal> so that <benefit>') "
        "and is paired with comprehensive Given-When-Then Gherkin acceptance criteria to facilitate automated test-driven development (TDD).\n\n"
        "The project follows a 2-week Sprint Plan (10 business trading days), scoping 55 story points with each user story scoped "
        "to be completely implementable and verifiable within 2 to 4 engineering hours (under half a day)."
    )

    # Section 2: Epics Overview
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Epic Architecture & Functional Categorization")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(30, 58, 138)

    epics_table = doc.add_table(rows=6, cols=4)
    epics_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(epics_table)
    headers = ["Epic ID", "Epic Name", "Scope & Description", "Total Points"]
    for i, h in enumerate(headers):
        style_header_cell(epics_table.rows[0].cells[i], h)

    epic_rows = [
        ("EPIC-DATA", "Data Pipeline & Quality Layer", "Database connectivity (SQLite), seed data sanitization, null value imputation, duplicate removal, and high-performance cached analytical view integration.", "8 pts (3 Stories)"),
        ("EPIC-DASH", "Core Executive Dashboard", "Tailwind CSS KPI card header, responsive interactive sidebar filters (date, region, store, category), and high-level branch ranking deficit table.", "11 pts (3 Stories)"),
        ("EPIC-ANLY", "Dimensional & Behavioral Analysis", "Category contribution matrix, day-of-week seasonality footfall analysis, sales staff quota leaderboard, and Average Transaction Value benchmark gauge.", "16 pts (4 Stories)"),
        ("EPIC-INST", "Root Cause & Predictive Insights", "Volume vs Price variance waterfall decomposition, promotional discount elasticity scatter plot, MTD cumulative trajectory, and run-rate forecast engine.", "15 pts (4 Stories)"),
        ("EPIC-QUAL", "Governance & Anomaly Detection", "Automated Z-score statistical outlier flagging to separate isolated shocks from persistent structural deficit patterns.", "5 pts (1 Story)")
    ]
    for row_idx, r_data in enumerate(epic_rows):
        row = epics_table.rows[row_idx+1]
        for col_idx, text in enumerate(r_data):
            style_data_cell(row.cells[col_idx], text, is_even=(row_idx%2==1), bold=(col_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 3: Master User Stories Table
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Master User Stories Catalog (15 Stories, 100% OQ & BR Coverage)")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(30, 58, 138)

    stories_summary = [
        ("US-01", "Data Analyst", "Ingest clean SQLite data and execute 40-col view", "FR-01", "BR-01", "OQ-01", "Must", "3"),
        ("US-02", "Data Analyst", "Impute null discounts & remove duplicate transactions", "FR-01", "BR-01", "OQ-01", "Must", "2"),
        ("US-03", "Data Analyst", "Cache analytical view with st.cache_data for <1s latency", "FR-01", "BR-01", "OQ-01", "Must", "3"),
        ("US-04", "Commercial VP", "View Tailwind executive KPI cards for total actual vs target", "FR-02", "BR-01", "OQ-01", "Must", "3"),
        ("US-05", "Store Manager", "Filter all dashboard metrics by Date, Region, Format & Store", "FR-03", "BR-01", "OQ-01, OQ-02", "Must", "3"),
        ("US-06", "Operations Lead", "Sortable Store Ranking Table with absolute & % deficit alerts", "FR-04", "BR-02", "OQ-02", "Must", "5"),
        ("US-07", "Category Head", "Merchandise Category Bar Chart comparing target vs actual revenue", "FR-05", "BR-03", "OQ-03", "Must", "5"),
        ("US-08", "Store Manager", "Day-of-Week Seasonality Heatmap to pinpoint weekday footfall lag", "FR-06", "BR-04", "OQ-04", "Should", "3"),
        ("US-09", "Store Supervisor", "Staff Quota Attainment Leaderboard with <85% coaching alerts", "FR-07", "BR-05", "OQ-05", "Should", "5"),
        ("US-10", "Category Head", "Average Transaction Value (ATV) Benchmark Gauge against targets", "FR-04", "BR-08", "OQ-08", "Must", "3"),
        ("US-11", "Business Analyst", "Volume vs Price Variance Waterfall Decomposition chart", "FR-10", "BR-09", "OQ-09", "Must", "5"),
        ("US-12", "Merchandise VP", "Discount vs Margin Scatter Plot to identify gross profit erosion", "FR-09", "BR-07", "OQ-07", "Should", "5"),
        ("US-13", "Financial Analyst", "MTD Cumulative Burn-Up Curve vs target velocity trajectory", "FR-08", "BR-06", "OQ-06", "Must", "5"),
        ("US-14", "Commercial VP", "Month-End Revenue Run-Rate Projection with deficit forecaster", "FR-11", "BR-10", "OQ-10", "Must", "5"),
        ("US-15", "Data Analyst", "Statistical Anomaly Detection (Z-Score > 3.0) outlier table", "FR-12", "BR-11", "OQ-01, OQ-02", "Could", "5")
    ]

    master_table = doc.add_table(rows=len(stories_summary)+1, cols=8)
    master_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(master_table)
    m_headers = ["US ID", "Persona", "Summary Goal", "FR ID", "BR ID", "OQ", "Pri", "Pts"]
    for i, h in enumerate(m_headers):
        style_header_cell(master_table.rows[0].cells[i], h)

    for row_idx, r_data in enumerate(stories_summary):
        row = master_table.rows[row_idx+1]
        for col_idx, text in enumerate(r_data):
            style_data_cell(row.cells[col_idx], text, is_even=(row_idx%2==1), bold=(col_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 4: Detailed User Stories with Gherkin Acceptance Criteria
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. Detailed User Stories & Gherkin Acceptance Criteria")
    r_h4.font.name = "Calibri"
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(30, 58, 138)

    detailed_stories = [
        {
            "id": "US-01",
            "title": "SQLite Database Connection & Master View Execution",
            "epic": "EPIC-DATA (Data Pipeline)",
            "persona": "Business Data Analyst (SH-05)",
            "statement": "As a Business Data Analyst, I want the system to reliably connect to data/app.db and query vw_daily_sales_target_analysis so that the application operates on 100% verified, normalized relational data.",
            "fr": "FR-01 (Data Ingestion & Integrity)",
            "br": "BR-01 | OQ-01",
            "priority": "Must Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: Successful Database Ingestion",
                 "Given the application starts and data/app.db exists,\n"
                 "When load_data() is executed,\n"
                 "Then exactly 300 records with 40 distinct columns are loaded into a pandas DataFrame without SQL connection timeout."),
                ("Scenario 2: Missing Database File Fallback",
                 "Given data/app.db is missing or corrupted,\n"
                 "When load_data() is executed,\n"
                 "Then the system gracefully catches the FileNotFoundError and displays a user-friendly error message instructing the user to run src/generate_data.py.")
            ]
        },
        {
            "id": "US-02",
            "title": "Data Sanitization, Null Handling & Deduplication",
            "epic": "EPIC-DATA (Data Pipeline)",
            "persona": "Business Data Analyst (SH-05)",
            "statement": "As a Business Data Analyst, I want the data ingestion engine to impute missing discount values and remove duplicate transactional rows so that downstream KPI metrics are mathematically pure.",
            "fr": "FR-01 (Data Ingestion & Cleaning)",
            "br": "BR-01 | OQ-01",
            "priority": "Must Have | 2 Story Points",
            "criteria": [
                ("Scenario 1: Missing Discount Imputation",
                 "Given raw transactional rows contain NaN in discount_amount,\n"
                 "When the clean_sales_data() function executes,\n"
                 "Then all NaN values are imputed with 0.0 and zero nulls remain in the final DataFrame."),
                ("Scenario 2: Duplicate Row Removal",
                 "Given the raw feed contains 3 duplicate sales rows,\n"
                 "When clean_sales_data() executes,\n"
                 "Then exactly 3 duplicate records are purged, leaving exactly 300 unique primary sales records.")
            ]
        },
        {
            "id": "US-03",
            "title": "Analytical View In-Memory Caching (st.cache_data)",
            "epic": "EPIC-DATA (Data Pipeline)",
            "persona": "Business Data Analyst (SH-05)",
            "statement": "As a Business Data Analyst, I want cleaned dataset frames cached in memory using Streamlit cache decorators so that subsequent filter clicks reload in under 1 second.",
            "fr": "FR-01 (Performance & Caching)",
            "br": "BR-01 | OQ-01",
            "priority": "Must Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: Cached Retrieval on Filter Change",
                 "Given the data has been loaded and initial cache generated,\n"
                 "When any sidebar filter is toggled,\n"
                 "Then the data frame is served from st.cache_data without re-reading the SQLite database from disk in < 0.5 seconds.")
            ]
        },
        {
            "id": "US-04",
            "title": "Executive Tailwind KPI Cards Header",
            "epic": "EPIC-DASH (Core Executive Dashboard)",
            "persona": "Commercial Sales Director / Regional VP (SH-03)",
            "statement": "As a Commercial Sales Director, I want high-level Tailwind-styled KPI cards displaying Total Actual Sales, Total Target, Net Variance ($), Achievement Rate (%), and Realized Gross Margin so that I can assess enterprise health in under 5 seconds.",
            "fr": "FR-02 (Executive KPI Header)",
            "br": "BR-01 | OQ-01",
            "priority": "Must Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: KPI Cards Render Correct Totals",
                 "Given the user opens the Overview dashboard page,\n"
                 "When data is rendered,\n"
                 "Then 5 distinct cards display Actual Revenue, Target Revenue, Dollar Variance, % Achievement, and Margin with appropriate green/red delta badges based on whether achievement >= 100%."),
                ("Scenario 2: Underperformance Color Alert",
                 "Given the filtered achievement rate is less than 85%,\n"
                 "When the KPI card renders,\n"
                 "Then the achievement badge renders with red background and an alert icon indicating critical deficit status.")
            ]
        },
        {
            "id": "US-05",
            "title": "Multi-Dimensional Interactive Sidebar Filters",
            "epic": "EPIC-DASH (Core Executive Dashboard)",
            "persona": "Branch Store Manager (SH-01)",
            "statement": "As a Branch Store Manager, I want to filter the entire dashboard by Date Range, Region, Store Format, Specific Store, and Category so that I can isolate metrics specific to my operational responsibility.",
            "fr": "FR-03 (Dynamic Sidebar Navigation)",
            "br": "BR-01, BR-02 | OQ-01, OQ-02",
            "priority": "Must Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: Filter Cascading",
                 "Given the user selects Region = 'North',\n"
                 "When the Store dropdown is inspected,\n"
                 "Then only North region stores (Connaught, Cyber City, Sector 17, Hazratganj) appear in the store selection list."),
                ("Scenario 2: Synchronized Visual Updates",
                 "Given any filter selection changes,\n"
                 "When the application re-runs,\n"
                 "Then all KPI cards, tables, and charts instantly update to reflect the filtered slice.")
            ]
        },
        {
            "id": "US-06",
            "title": "Store Target Achievement Ranking & Deficit Table",
            "epic": "EPIC-DASH (Core Executive Dashboard)",
            "persona": "Store Operations Manager (SH-01)",
            "statement": "As a Store Operations Manager, I want a sortable data table of all retail stores ranked by Target Achievement Rate (%) and Absolute Deficit ($) so that I can immediately pinpoint branches requiring operational intervention.",
            "fr": "FR-04 (Store Ranking & Deficit Table)",
            "br": "BR-02 | OQ-02",
            "priority": "Must Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Store Table Ranking",
                 "Given the Store Performance view is active,\n"
                 "When the user views the ranking table,\n"
                 "Then stores are sorted by Achievement Rate ascending (worst performers first), with deficit highlighted in bold red."),
                ("Scenario 2: Deficit Status Badge",
                 "Given a store has an achievement rate < 85%,\n"
                 "When the table renders,\n"
                 "Then a 'CRITICAL DEFICIT' tag is displayed in the Status column.")
            ]
        },
        {
            "id": "US-07",
            "title": "Merchandise Category Quota Contribution Bar Chart",
            "epic": "EPIC-ANLY (Dimensional Analysis)",
            "persona": "Category Head / Merchandise Manager (SH-02)",
            "statement": "As a Category Head, I want a grouped horizontal bar chart comparing Actual vs Target revenue across all 5 merchandise categories so that I can see which departments are driving or dragging store performance.",
            "fr": "FR-05 (Category Contribution Analysis)",
            "br": "BR-03 | OQ-03",
            "priority": "Must Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Grouped Bar Chart Display",
                 "Given the Category Analysis view is selected,\n"
                 "When the chart renders,\n"
                 "Then each category displays paired bars for Actual Net Sales and Target Revenue with exact INR values formatted in tooltips."),
                ("Scenario 2: Hover Details",
                 "Given the user hovers over the Electronics bar,\n"
                 "When the tooltip triggers,\n"
                 "Then it displays Gross Revenue, Discount Amount, Net Revenue, Target, and % Attainment.")
            ]
        },
        {
            "id": "US-08",
            "title": "Day-of-Week Seasonality & Footfall Lag Heatmap",
            "epic": "EPIC-ANLY (Dimensional Analysis)",
            "persona": "Store Operations Manager (SH-01)",
            "statement": "As a Store Operations Manager, I want an interactive heatmap showing sales achievement rates across Monday through Sunday so that I can optimize associate rostering and identify weekday trading lulls.",
            "fr": "FR-06 (Day-of-Week Seasonality)",
            "br": "BR-04 | OQ-04",
            "priority": "Should Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: Weekday Heatmap Visualization",
                 "Given the Seasonality tab is selected,\n"
                 "When the chart renders,\n"
                 "Then days of the week are displayed chronologically (Monday to Sunday) with color scale mapping achievement rate (deep red for <85%, deep green for >105%)."),
                ("Scenario 2: Lag Confirmation",
                 "Given the synthetic trading dataset,\n"
                 "When Tuesday and Wednesday cells are inspected,\n"
                 "Then they render with red/amber shading confirming weekday trading volume deficits.")
            ]
        },
        {
            "id": "US-09",
            "title": "Sales Staff Quota Attainment Leaderboard",
            "epic": "EPIC-ANLY (Dimensional Analysis)",
            "persona": "Category Head / Floor Supervisor (SH-02)",
            "statement": "As a Floor Supervisor, I want a staff leaderboard showing associate revenue vs base daily quotas with automated coaching flags (<85%) so that I can provide targeted shift coaching.",
            "fr": "FR-07 (Staff Performance Leaderboard)",
            "br": "BR-05 | OQ-05",
            "priority": "Should Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Staff Coaching Alert",
                 "Given a staff member's total sales / total quota < 0.85,\n"
                 "When the staff table renders,\n"
                 "Then the Coaching Required flag displays 'YES' in high-contrast red font."),
                ("Scenario 2: Store Specific Filtering",
                 "Given the user filters by Store STR-001,\n"
                 "When the staff leaderboard renders,\n"
                 "Then only associates belonging to STR-001 are displayed.")
            ]
        },
        {
            "id": "US-10",
            "title": "Average Transaction Value (ATV) Benchmark Gauge",
            "epic": "EPIC-ANLY (Dimensional Analysis)",
            "persona": "Category Head / Floor Supervisor (SH-02)",
            "statement": "As a Category Head, I want an ATV indicator gauge showing realized basket size versus targeted ticket benchmark (₹) so that I can determine if customer basket depth meets expectations.",
            "fr": "FR-04 (Store & Basket Analysis)",
            "br": "BR-08 | OQ-08",
            "priority": "Must Have | 3 Story Points",
            "criteria": [
                ("Scenario 1: Gauge Rendering",
                 "Given the store selection is active,\n"
                 "When the ATV widget renders,\n"
                 "Then it displays current realized ATV, target ATV benchmark, and the delta variance with threshold coloring.")
            ]
        },
        {
            "id": "US-11",
            "title": "Volume vs Price Variance Waterfall Decomposition",
            "epic": "EPIC-INST (Root Cause Insights)",
            "persona": "Business Data Analyst (SH-05)",
            "statement": "As a Business Data Analyst, I want an algebraic waterfall chart decomposing sales variance into Volume Effect (order deficit) and Price Effect (ATV deficit) so that management knows whether to fix footfall or basket size.",
            "fr": "FR-10 (Root Cause Variance Decomposition)",
            "br": "BR-09 | OQ-09",
            "priority": "Must Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Mathematical Identity Reconciliation",
                 "Given any store or category selection,\n"
                 "When the waterfall chart calculates,\n"
                 "Then Volume Effect Dollar + Price Effect Dollar exactly equals Total Net Revenue Variance."),
                ("Scenario 2: Waterfall Visual Representation",
                 "Given target revenue is the starting bar and actual revenue is the ending bar,\n"
                 "When the chart renders,\n"
                 "Then intermediate floating bars display the positive or negative contribution of Volume and ATV.")
            ]
        },
        {
            "id": "US-12",
            "title": "Promotional Discount Elasticity & Margin Scatter Plot",
            "epic": "EPIC-INST (Root Cause Insights)",
            "persona": "Commercial Sales Director (SH-03)",
            "statement": "As a Commercial Sales Director, I want a scatter plot of Discount % versus Realized Gross Margin % with an OLS regression trendline so that I can identify if heavy markdowns are destroying gross profits.",
            "fr": "FR-09 (Discount vs Margin Analysis)",
            "br": "BR-07 | OQ-07",
            "priority": "Should Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Scatter Plot with OLS Trendline",
                 "Given the Pricing & Margin page is open,\n"
                 "When data is plotted,\n"
                 "Then each transaction slice appears as a scatter point with a downward sloping regression line showing negative correlation between discount and margin."),
                ("Scenario 2: Correlation Metric Display",
                 "Given the scatter plot is displayed,\n"
                 "When statistical summary card renders,\n"
                 "Then Pearson's correlation coefficient r and p-value are explicitly shown.")
            ]
        },
        {
            "id": "US-13",
            "title": "MTD Cumulative Trajectory & Target Pace Curve",
            "epic": "EPIC-INST (Root Cause Insights)",
            "persona": "Retail Financial Planner (SH-04)",
            "statement": "As a Retail Financial Planner, I want a dual-line chart showing cumulative actual sales vs cumulative target budget over elapsed trading days so that I can assess whether revenue velocity is on pace.",
            "fr": "FR-08 (MTD Cumulative Burn-Up)",
            "br": "BR-06 | OQ-06",
            "priority": "Must Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Cumulative Trajectory Line Chart",
                 "Given trading days 1 to 20,\n"
                 "When the MTD chart renders,\n"
                 "Then the cumulative target line increases monotonically and the cumulative actual line tracks actual realized cumulative revenue."),
                ("Scenario 2: Pace Deficit Shading",
                 "Given actual cumulative sales fall below target cumulative sales,\n"
                 "When the chart renders,\n"
                 "Then the area between lines is shaded in light red highlighting the cumulative gap.")
            ]
        },
        {
            "id": "US-14",
            "title": "Month-End Revenue Run-Rate Projection Engine",
            "epic": "EPIC-INST (Root Cause Insights)",
            "persona": "Commercial Sales Director (SH-03)",
            "statement": "As a Commercial Sales Director, I want a run-rate forecasting module projecting full-month revenue based on current average daily velocity so that I can anticipate monthly quota surpluses or deficits before month-end.",
            "fr": "FR-11 (Run-Rate Projection Engine)",
            "br": "BR-10 | OQ-10",
            "priority": "Must Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Accurate Run-Rate Formula",
                 "Given 20 elapsed trading days with cumulative revenue R and 30 total month days,\n"
                 "When the projection runs,\n"
                 "Then projected revenue equals (R / 20) * 30 and projected deficit equals Projected Revenue - Monthly Target."),
                ("Scenario 2: Executive Summary Badge",
                 "Given projected deficit < 0,\n"
                 "When the forecast card renders,\n"
                 "Then a warning card states 'Projected Month-End Shortfall of ₹X' with required daily catch-up run rate.")
            ]
        },
        {
            "id": "US-15",
            "title": "Statistical Anomaly Detection using Z-Scores",
            "epic": "EPIC-QUAL (Governance & Anomaly Detection)",
            "persona": "Business Data Analyst (SH-05)",
            "statement": "As a Business Data Analyst, I want an anomaly detection table identifying daily sales records where absolute Z-score of net revenue exceeds 3.0 standard deviations so that I can isolate external flash events from regular trends.",
            "fr": "FR-12 (Statistical Anomaly Detection)",
            "br": "BR-11 | OQ-01, OQ-02",
            "priority": "Could Have | 5 Story Points",
            "criteria": [
                ("Scenario 1: Outlier Identification",
                 "Given the dataset contains 3 flash surge sales records (>3 sigma),\n"
                 "When the anomaly detection module executes,\n"
                 "Then exactly those 3 records are flagged with Z-scores > 3.0 and displayed in the Anomaly Inspection Table."),
                ("Scenario 2: Analyst Exclusion Toggle",
                 "Given an analyst reviews anomalies,\n"
                 "When 'Exclude Outliers from Baseline Trend' checkbox is checked,\n"
                 "Then all downstream regression trendlines automatically recalculate excluding the 3 outlier records.")
            ]
        }
    ]

    for us in detailed_stories:
        p_story_head = doc.add_paragraph()
        p_story_head.paragraph_format.space_before = Pt(12)
        p_story_head.paragraph_format.space_after = Pt(2)
        r_sh = p_story_head.add_run(f"Story {us['id']}: {us['title']}")
        r_sh.font.name = "Calibri"
        r_sh.font.size = Pt(12)
        r_sh.font.bold = True
        r_sh.font.color.rgb = RGBColor(30, 58, 138)

        st_table = doc.add_table(rows=6, cols=2)
        st_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(st_table)

        meta_rows = [
            ("Epic & Persona:", f"{us['epic']} | {us['persona']}"),
            ("User Story Statement:", f"\"{us['statement']}\""),
            ("Traceability Mapping:", f"Feature: {us['fr']} | BR: {us['br']}"),
            ("Priority & Estimation:", us['priority'])
        ]
        for r_idx, (k, v) in enumerate(meta_rows):
            row = st_table.rows[r_idx]
            style_data_cell(row.cells[0], k, is_even=(r_idx%2==1), width=Inches(1.8), bold=True)
            style_data_cell(row.cells[1], v, is_even=(r_idx%2==1), width=Inches(4.7))

        # Row 4 & 5: Acceptance Criteria
        row_crit_lbl = st_table.rows[4]
        style_data_cell(row_crit_lbl.cells[0], "Acceptance Criteria:", width=Inches(1.8), bold=True)
        crit_text_combined = ""
        for title, gherkin in us["criteria"]:
            crit_text_combined += f"• {title}:\n{gherkin}\n\n"
        crit_text_combined = crit_text_combined.strip()
        style_data_cell(row_crit_lbl.cells[1], crit_text_combined, width=Inches(4.7))

        # Definition of Done
        row_dod = st_table.rows[5]
        style_data_cell(row_dod.cells[0], "Definition of Done (DoD):", is_even=True, width=Inches(1.8), bold=True)
        dod_text = "Unit tests pass in pytest; Streamlit component renders < 1s; no console errors; clean code committed."
        style_data_cell(row_dod.cells[1], dod_text, is_even=True, width=Inches(4.7))

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 5: Sprint Planning & Execution Schedule (2-Week / 10-Day Plan)
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run("5. Two-Week Sprint Implementation Schedule (Day-by-Day)")
    r_h5.font.name = "Calibri"
    r_h5.font.size = Pt(14)
    r_h5.font.bold = True
    r_h5.font.color.rgb = RGBColor(30, 58, 138)

    p_sprint_intro = doc.add_paragraph()
    p_sprint_intro.paragraph_format.space_after = Pt(8)
    p_sprint_intro.add_run(
        "The project is structured into two 1-week sprints (5 working days each, total 10 development days). "
        "Sprint 1 focuses on data engineering foundations, Tailwind executive KPI cards, and core operational reporting. "
        "Sprint 2 focuses on advanced variance decomposition, statistical elasticity modeling, run-rate projections, and anomaly governance."
    )

    sprint_table = doc.add_table(rows=11, cols=5)
    sprint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sprint_table)
    s_headers = ["Sprint", "Day", "Stories Planned", "Daily Focus & Deliverable", "Story Pts"]
    for i, h in enumerate(s_headers):
        style_header_cell(sprint_table.rows[0].cells[i], h)

    sprint_days = [
        ("Sprint 1", "Day 1", "US-01, US-02", "SQLite connection, seed data ingestion, NaN imputation, deduplication", "5 pts"),
        ("Sprint 1", "Day 2", "US-03, US-04", "Streamlit caching (st.cache_data), Tailwind executive KPI cards header", "6 pts"),
        ("Sprint 1", "Day 3", "US-05", "Interactive multi-dimensional filter sidebar with cascading selections", "3 pts"),
        ("Sprint 1", "Day 4", "US-06, US-10", "Store ranking deficit table, alert flags, and ATV benchmark gauge", "8 pts"),
        ("Sprint 1", "Day 5", "US-07", "Merchandise category contribution horizontal grouped bar chart", "5 pts"),
        ("Sprint 2", "Day 6", "US-08, US-09", "Day-of-Week seasonality heatmap and sales staff quota leaderboard", "8 pts"),
        ("Sprint 2", "Day 7", "US-11", "Volume vs Price variance waterfall root-cause decomposition module", "5 pts"),
        ("Sprint 2", "Day 8", "US-12", "Promotional discount elasticity & gross profit margin scatter plot", "5 pts"),
        ("Sprint 2", "Day 9", "US-13, US-14", "MTD cumulative burn-up trajectory and month-end run-rate projection engine", "10 pts"),
        ("Sprint 2", "Day 10", "US-15", "Statistical anomaly detection table (Z-score > 3.0), QA verification & sign-off", "5 pts")
    ]

    for row_idx, r_data in enumerate(sprint_days):
        row = sprint_table.rows[row_idx+1]
        for col_idx, text in enumerate(r_data):
            style_data_cell(row.cells[col_idx], text, is_even=(row_idx%2==1), bold=(col_idx in [0, 1]))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 6: Story Points & Velocity Analysis
    h6 = doc.add_paragraph()
    h6.paragraph_format.space_before = Pt(14)
    h6.paragraph_format.space_after = Pt(4)
    r_h6 = h6.add_run("6. Velocity, Estimation & Risk Governance")
    r_h6.font.name = "Calibri"
    r_h6.font.size = Pt(14)
    r_h6.font.bold = True
    r_h6.font.color.rgb = RGBColor(30, 58, 138)

    p_velo = doc.add_paragraph()
    p_velo.paragraph_format.space_after = Pt(8)
    p_velo.add_run(
        "• Total Estimated Story Points: 55 Story Points across 15 user stories.\n"
        "• Sprint Velocity: 27.5 points / 5-day sprint (average 5.5 story points per day, corresponding to 1 senior developer / data analyst).\n"
        "• Fibonacci Estimation Scale: 1 (Trivial config), 2 (Simple cleaning routine), 3 (Standard UI widget/card), 5 (Complex interactive chart or analytical algorithm), 8 (Full multi-table module - decomposed into 3+5).\n"
        "• Traceability Guarantee: Every story maps 1:1 to at least one Business Requirement (BR), one Technical Functional Requirement (FR), and directly answers at least one Operational Question (OQ).\n"
        "• Contingency Buffer: Day 10 incorporates a 0.5-day buffer for end-to-end integration testing and automated regression verification using pytest."
    )

    doc_path = os.path.abspath("docs/11_User_Stories.docx")
    doc.save(doc_path)
    print(f"Generated {doc_path} successfully.")

def update_rtm_excel():
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]

    # Mapping based on BR ID:
    # Feature (Col 8) and User Story (Col 10)
    # Row indices in Traceability_Matrix:
    # Row 4 is header: OQ | HMW | Req ID (BR) | Requirement Statement | Stakeholder (SH) | Priority | Status | Feature (Step 6) | DB table.column (Step 8) | User Story (Step 11) | UI Screen (Step 12) | Test Case (TC-F / TC-T) (Step 13)
    
    mapping = {
        "BR-01": ("FR-01, FR-02, FR-03", "US-01, US-02, US-03, US-04, US-05"),
        "BR-02": ("FR-03, FR-04", "US-05, US-06"),
        "BR-03": ("FR-05", "US-07"),
        "BR-04": ("FR-06", "US-08"),
        "BR-05": ("FR-07", "US-09"),
        "BR-06": ("FR-08", "US-13"),
        "BR-07": ("FR-09", "US-12"),
        "BR-08": ("FR-04", "US-10"),
        "BR-09": ("FR-10", "US-11"),
        "BR-10": ("FR-11", "US-14"),
        "BR-11": ("FR-12", "US-15"),
        "BR-12": ("FR-01 (Phase 2)", "Out of Scope (Phase 2)")
    }

    font_calibri = Font(name="Calibri", size=9)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for row_idx in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row_idx, column=3).value
        if br_id in mapping:
            fr_val, us_val = mapping[br_id]
            # Column 8: Feature (Step 6/7)
            cell_feat = ws.cell(row=row_idx, column=8)
            cell_feat.value = fr_val
            cell_feat.font = font_calibri
            cell_feat.alignment = align_left

            # Column 10: User Story (Step 11)
            cell_us = ws.cell(row=row_idx, column=10)
            cell_us.value = us_val
            cell_us.font = font_calibri
            cell_us.alignment = align_left

    wb.save(rtm_path)
    print(f"Updated {rtm_path} with User Story & Feature mappings successfully.")

if __name__ == "__main__":
    create_user_stories_doc()
    update_rtm_excel()
