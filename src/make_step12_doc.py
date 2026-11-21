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

def add_wireframe_block(doc, title, ascii_art):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(10)
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
    r_code = p_box.add_run(ascii_art)
    r_code.font.name = "Consolas"
    r_code.font.size = Pt(7.5)
    r_code.font.color.rgb = RGBColor(15, 23, 42)

def create_ui_framework_doc():
    doc = Document()

    # 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("UI ARCHITECTURE & WIREFRAME FRAMEWORK")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Project P054: Daily Sales Actual vs Target Analysis | Multi-Page Streamlit + Tailwind CDN Architecture")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # Document Control Box
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table)
    meta_data = [
        ("Candidate / Roll No:", "Harshit | University Roll No: 2500320140058 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Project Mentor:", "Dinesh Sereno | Industry Data Engineering Lead"),
        ("UI Technology Stack:", "Streamlit Multi-Page, Tailwind CSS (via iframe CDN), Plotly Graph Objects, openpyxl"),
        ("Document Status & Date:", "Version 1.0 (UI Baseline Specification) | November 2026")
    ]
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        style_data_cell(row.cells[0], k, is_even=(row_idx%2==1), width=Inches(2.2), bold=True)
        style_data_cell(row.cells[1], v, is_even=(row_idx%2==1), width=Inches(4.3))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Executive Summary & UI Design Principles
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Executive UI Architecture & Design Principles")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(30, 58, 138)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.add_run(
        "The user interface for Project P054 is designed around three core principles: instant cognitive clarity, zero visual clutter, "
        "and sub-second analytical interactivity. Rather than traditional clunky BI dashboards, the app uses a dual-engine architecture: "
        "Streamlit multi-page routing for computational workflows, coupled with a Tailwind CSS component design system rendered inside "
        "st.components.v1.html containers. This gives executive stakeholders responsive, card-based KPI banners and styled alert badges "
        "while retaining high-performance native Plotly interactive visualizations.\n\n"
        "Key Design Tenets:\n"
        "1. 5-Second Executive Assessment: Key revenue KPIs and deficit status are immediately visible without scrolling.\n"
        "2. Stateful Global Filter Sidebar: A shared filter state (st.session_state) preserves user selections across multi-page tab switches.\n"
        "3. Semantic Color Palette: Emerald Green (#10B981) for quota achievement, Crimson Rose (#EF4444) for operational deficits, and Corporate Navy (#1E3A8A) for structure.\n"
        "4. Progressive Diagnostic Drilldown: Navigating from Executive Overview (UI-01) to Deep Dive Root Cause (UI-04) and Forecasting (UI-05)."
    )

    # Section 2: UI Screens Catalog
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Application Screen Directory (7 Production Screens)")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(30, 58, 138)

    screens_data = [
        ("UI-01", "Executive Overview", "High-level enterprise summary, total actual vs target revenue, deficit alerts, and top/bottom store rankings.", "Tailwind KPI Header, Store Ranking Table, ATV Gauge", "Date, Region, Store, Category", "US-01, US-04, US-05, US-06, US-10", "OQ-01, OQ-02, OQ-08"),
        ("UI-02", "Trends & Seasonality", "Temporal performance tracking, day-of-week footfall lag, and MTD cumulative burn-up pace vs target curve.", "Day-of-Week Heatmap, MTD Burn-Up Dual Line Chart", "Date Range, Store Format", "US-08, US-13", "OQ-04, OQ-06"),
        ("UI-03", "Store & Category Compare", "Cross-store benchmark comparisons, merchandise category contribution, and associate sales leaderboard.", "Grouped Horizontal Bar Chart, Staff Quota Table", "Region, Store, Department", "US-07, US-09", "OQ-03, OQ-05"),
        ("UI-04", "Root Cause & Elasticity", "Algebraic variance decomposition (Volume vs Price effect) and promotional discount elasticity scatter plot.", "Waterfall Decomposition Chart, OLS Regression Scatter", "Store, Merchandise Category", "US-11, US-12", "OQ-07, OQ-09"),
        ("UI-05", "Predictive Forecasting", "Daily run-rate velocity calculator, full-month projected revenue, and required daily catch-up quotas.", "Run-rate Gauge, Forecast Deficit Card, Bar Chart", "Store Selection, Target Month", "US-14", "OQ-10"),
        ("UI-06", "Data Governance & Quality", "Full dataset inspection, data cleaning audit log (nulls/duplicates), and Z-score anomaly detector table.", "Data Grid, Anomaly Filter Table, Excel Export Button", "Outlier Z-score Threshold, Filter", "US-02, US-15", "OQ-01, OQ-02"),
        ("UI-07", "System Architecture & About", "Project metadata, ABES academic credentials, mentor attribution, ER schema viewer, and data dictionary.", "Architecture Flowchart, Data Dictionary Table, Links", "None (Static Reference)", "All US (Technical Baseline)", "Governance")
    ]

    screens_table = doc.add_table(rows=len(screens_data)+1, cols=7)
    screens_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(screens_table)
    sc_headers = ["UI ID", "Screen Name", "Core Purpose", "Key Components", "Filters", "User Stories", "OQs"]
    for i, h in enumerate(sc_headers):
        style_header_cell(screens_table.rows[0].cells[i], h)

    for r_idx, r_vals in enumerate(screens_data):
        row = screens_table.rows[r_idx+1]
        for c_idx, val in enumerate(r_vals):
            style_data_cell(row.cells[c_idx], val, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 3: ASCII Wireframes per Screen
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Detailed ASCII Wireframes (Screens UI-01 through UI-07)")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(30, 58, 138)

    # Wireframe UI-01
    ascii_ui01 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR (st.session_state)     |  UI-01: EXECUTIVE OVERVIEW & KPI COCKPIT                            |
|---------------------------------+---------------------------------------------------------------------|
| [X] Retail Sales Analytics      |  +---------------------------------------------------------------+  |
|                                 |  | [TAILWIND HEADER] P054 Daily Sales Actual vs Target Dashboard |  |
| FILTER CRITERIA:                |  | Release 1.0 (MVP) | Harshit (ABES) | Mentor: Dinesh Sereno        |  |
| Date Range: [Nov 01 - Nov 20]   |  +---------------------------------------------------------------+  |
| Region:     [All Regions    v]  |                                                                     |
| Store Fmt:  [All Formats    v]  |  [TAILWIND KPI CARD GRID - 5 TILES]                                 |
| Store Name: [Connaught Mall v]  |  +---------------+ +---------------+ +---------------+ +----------+  |
| Category:   [All Categories v]  |  | ACTUAL REVENUE| | TARGET REVENUE| | VARIANCE ($)  | | ACHIEVE %|  |
|                                 |  | INR 8,421,500 | | INR 8,850,000 | | -INR 428,500  | | 95.16%   |  |
| [x] Live Cache Invalidation     |  | ^ +4.2% DoD   | | Quota Base    | | v -4.84% Gap  | | [DEFICIT]|  |
|                                 |  +---------------+ +---------------+ +---------------+ +----------+  |
| NAVIGATION (Pages):             |                                                                     |
| > 01 Overview                   |  +-------------------------------------+ +--------------------------+  |
|   02 Trends & Seasonality       |  | STORE TARGET RANKING TABLE          | | ATV BASKET SIZE BENCHMARK|  |
|   03 Store & Category Compare   |  | (fact_daily_sales vs targets)       | | (actual vs target ATV)   |  |
|   04 Root Cause Variance        |  | Store | Target | Actual | %Ach |Stat| | [Plotly Indicator Gauge] |  |
|   05 Predictive Forecast        |  | Bandra| 2.10M  | 2.25M  | 107% |PASS| | Realized: INR 2,840       |  |
|   06 Data Quality & Audit       |  | Conngt| 1.80M  | 1.72M  |  95% |WARN| | Target:   INR 2,950       |  |
|   07 About & Architecture       |  | Fraser| 620K   | 496K   |  80% |CRIT| | Status:   -3.7% Ticket Gap|  |
+---------------------------------+----------------------------------------+----------------------------+"""
    add_wireframe_block(doc, "Screen UI-01: Executive Overview & KPI Cockpit Wireframe", ascii_ui01)

    # Wireframe UI-02
    ascii_ui02 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-02: TRENDS & SEASONALITY ANALYTICS                              |
|---------------------------------+---------------------------------------------------------------------|
| [X] Active Filter: Nov 01-20    |  +---------------------------------------------------------------+  |
| Scope: All Stores               |  | [TAILWIND HEADER] Day-of-Week Seasonality & Cumulative Traject|  |
|                                 |  +---------------------------------------------------------------+  |
| Seasonality Metric:             |                                                                     |
| (*) Target Achievement %        |  DAY-OF-WEEK SEASONALITY HEATMAP (OQ-04)                             |
| ( ) Net Revenue ($)             |  [Plotly Heatmap: Region (Y-Axis) x DayOfWeek Mon-Sun (X-Axis)]     |
|                                 |  +---------------------------------------------------------------+  |
| Cumulative Target View:         |  | Region  | Mon  | Tue   | Wed   | Thu  | Fri  | Sat   | Sun   | |  |
| [X] Show Target Pace Curve      |  | North   | 94%  | 86%   | 88%   | 96%  | 102% | 114%  | 112%  | |  |
| [X] Shade Cumulative Gap        |  | West    | 98%  | 89%   | 87%   | 97%  | 105% | 118%  | 115%  | |  |
|                                 |  | East    | 84%  | 78%   | 79%   | 85%  | 89%  | 95%   | 92%   | |  |
| QUICK INSIGHT BOX:              |  +---------------------------------------------------------------+  |
| "Tuesdays & Wednesdays reflect  |                                                                     |
| severe weekday footfall lag     |  MTD CUMULATIVE TRAJECTORY VS BUDGET PACE (OQ-06)                   |
| across East & Central outlets." |  [Plotly Dual-Line: Cumulative Sales (Solid) vs Target Pace (Dotted)]|
|                                 |  +---------------------------------------------------------------+  |
|                                 |  | Cumulative Revenue (INR Millions)                             |  |
|                                 |  |  8M |                                         / Target Pace   |  |
|                                 |  |  6M |                             /-- Actual Sales (Lagging) |  |
|                                 |  |  4M |                     /------/                            |  |
|                                 |  |  2M |             /------/                                    |  |
|                                 |  |  0M +-------------+------------+------------+---------------+  |
|                                 |  |     Day 01        Day 05       Day 10       Day 20 (Nov)    |  |
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-02: Trends & Seasonality Analytics Wireframe", ascii_ui02)

    # Wireframe UI-03
    ascii_ui03 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-03: STORE & MERCHANDISE PERFORMANCE COMPARISON                  |
|---------------------------------+---------------------------------------------------------------------|
| Compare Dimension:              |  +---------------------------------------------------------------+  |
| (*) Category Breakdown          |  | [TAILWIND HEADER] Merchandise Category Contribution & Staff   |  |
| ( ) Store Benchmarking          |  +---------------------------------------------------------------+  |
|                                 |                                                                     |
| Category Filter: [All 5 Categories] CATEGORY TARGET VS ACTUAL CONTRIBUTION (OQ-03)                    |
| Store Selection: [Connaught Mall] [Plotly Horizontal Paired Bar Chart: Category vs Target & Actual]  |
|                                 |  +---------------------------------------------------------------+  |
| Staff Leaderboard Min Quota:    |  | Category     | Actual Net Sales vs Target Quota (INR Lakhs)   |  |
| [Slider: INR 20,000]            |  | Electronics  | [========================] Actual INR 32.4L    |  |
|                                 |  |              | [==============================] Target INR 30.0L|  |
| Sort By:                        |  | Apparel      | [====================] Actual INR 21.2L        |  |
| (*) % Achievement (Ascending)   |  |              | [========================] Target INR 24.5L    |  |
| ( ) Net Volume Sold             |  | Home&Living  | [==============] Actual INR 14.1L             |  |
|                                 |  +---------------------------------------------------------------+  |
|                                 |                                                                     |
|                                 |  SALES ASSOCIATE QUOTA LEADERBOARD (OQ-05)                          |
|                                 |  Staff ID  | Associate Name   | Store | Quota | Sales | Ach % |Coac|  |
|                                 |  STF-104   | Rahul Sharma     | STR01 | 22.0K | 24.5K | 111%  | NO |  |
|                                 |  STF-105   | Sneha Iyer       | STR01 | 20.0K | 15.2K |  76%  |YES!|  |
|                                 |  STF-102   | Rohan Verma      | STR02 | 32.0K | 26.8K |  83%  |YES!|  |
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-03: Store & Merchandise Performance Comparison Wireframe", ascii_ui03)

    # Wireframe UI-04
    ascii_ui04 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-04: ROOT CAUSE VARIANCE DECOMPOSITION & ELASTICITY              |
|---------------------------------+---------------------------------------------------------------------|
| Variance Attribution Level:     |  +---------------------------------------------------------------+  |
| (*) Entire Network              |  | [TAILWIND HEADER] Mathematical Variance Decomposition & Elasticity|
| ( ) Single Store                |  +---------------------------------------------------------------+  |
|                                 |                                                                     |
| Target Store: [Fraser Road]     |  VOLUME VS PRICE VARIANCE WATERFALL (OQ-09)                         |
|                                 |  [Plotly Waterfall: Target -> Volume Effect -> Price Effect -> Actual]
| Discount Analysis Mode:         |  +---------------------------------------------------------------+  |
| (*) Margin Erosion Scatter      |  | INR   |  [Target: 8.85M]                                      |  |
| ( ) Discount Distribution       |  | 8.8M  |    |                                                  |  |
|                                 |  | 8.6M  |    +--- [Volume Deficit: -INR 310K]                   |  |
| Regression Model:               |  | 8.4M  |           |                                           |  |
| [X] Ordinary Least Squares (OLS)|  | 8.2M  |           +--- [Price/ATV Deficit: -INR 118.5K]       |  |
| [X] Display Pearson r & p-value |  | 8.0M  |                  |                                    |  |
|                                 |  |       |                  +---> [Actual Net Sales: 8.42M]      |  |
| STATISTICAL SUMMARY:            |  +---------------------------------------------------------------+  |
| Pearson r: -0.684               |                                                                     |
| p-value:   < 0.001 (Sig. @ 99%) |  PROMOTIONAL DISCOUNT ELASTICITY VS GROSS MARGIN (OQ-07)            |
| "Each 5% discount increment     |  [Plotly Scatter Plot: Discount % (X-Axis) vs Gross Margin % (Y)]    |
| causes 2.1% gross margin drop." |  Points plotted with OLS downward trendline confirming margin drop. |
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-04: Root Cause Variance Decomposition & Elasticity Wireframe", ascii_ui04)

    # Wireframe UI-05
    ascii_ui05 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-05: PREDICTIVE RUN-RATE & MONTH-END FORECASTING                 |
|---------------------------------+---------------------------------------------------------------------|
| Forecast Horizon:               |  +---------------------------------------------------------------+  |
| Elapsed Days:  20 Days          |  | [TAILWIND HEADER] Daily Run-Rate Velocity & Month-End Forecaster|  |
| Total Days:    30 Days          |  +---------------------------------------------------------------+  |
|                                 |                                                                     |
| Scenario Simulation:            |  VELOCITY RUN-RATE GAUGES & SUMMARY CARDS (OQ-10)                   |
| Quota Target: INR 13,200,000    |  +--------------------------------+ +------------------------------+ |
| Actual to Date: INR 8,421,500   |  | HISTORICAL RUN-RATE (20 DAYS)  | | REQUIRED RUN-RATE (10 DAYS)  | |
| Daily Average: INR 421,075/day  |  | Current: INR 421,075 / day     | | Needed:  INR 477,850 / day   | |
|                                 |  | Projected Month: 12.63M        | | Deficit: -INR 567,750 (4.3%) | |
| Adjust Catch-Up Run-Rate:       |  +--------------------------------+ +------------------------------+ |
| [Slider: +13.5% needed/day]     |                                                                     |
|                                 |  PROJECTED REVENUE TRAJECTORY (Plotly Line + Forecast Area)         |
| PROJECTION ALERT:               |  +---------------------------------------------------------------+  |
| "At current velocity, month-end |  | 14M |                                      / Target (13.2M)   |  |
| revenue will close at 12.63M    |  | 12M |                         /-----------/ Proj (12.63M)   |  |
| leaving a shortfall of 567.7K.  |  | 10M |             /----------/ [Projected Shortfall Area]   |  |
| Required daily run rate must    |  |  8M | -----------/ (Actual Nov 01-20)                         |  |
| increase to INR 477.8K."        |  |     Day 01                 Day 20 (Today)       Day 30 (EOM)  |  |
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-05: Predictive Run-Rate & Month-End Forecasting Wireframe", ascii_ui05)

    # Wireframe UI-06
    ascii_ui06 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-06: DATA GOVERNANCE, QUALITY AUDIT & EXPORT                     |
|---------------------------------+---------------------------------------------------------------------|
| Data View Mode:                 |  +---------------------------------------------------------------+  |
| (*) Cleaned Analysis View       |  | [TAILWIND HEADER] Data Governance Audit & Anomaly Detection   |  |
| ( ) Raw Sales Feed              |  +---------------------------------------------------------------+  |
| ( ) Dimension Stores            |                                                                     |
|                                 |  DATA CLEANING PIPELINE RECONCILIATION SUMMARY                      |
| Anomaly Threshold:              |  +-------------------+ +--------------------+ +--------------------+  |
| [Slider Z-Score: 3.0 Sigma]     |  | MISSING IMPUTED   | | DUPLICATES REMOVED | | ANOMALIES FLAGGED  |  |
|                                 |  | 6 null discounts  | | 3 duplicate rows   | | 3 Diwali spikes    |  |
| Download Format:                |  | Imputed with 0.0  | | Purged from feed   | | Z-Score > 3.20     |  |
| [ Button: Export Excel (.xlsx)] |  +-------------------+ +--------------------+ +--------------------+  |
| [ Button: Export Cleaned CSV  ] |                                                                     |
|                                 |  STATISTICAL OUTLIER INSPECTION TABLE (OQ-01, OQ-02 Anomaly Filter) |
| [x] Exclude Anomalies from View |  Date       | Store ID | Category | Net Sales | Mean   | StdDev | Z-Score |
|                                 |  2026-11-04 | STR-005  | CAT-01   | 342,100   | 92,400 | 72,100 | +3.46   |
|                                 |  2026-11-12 | STR-001  | CAT-02   | 388,400   | 92,400 | 72,100 | +4.10   |
|                                 |  2026-11-18 | STR-009  | CAT-01   | 326,800   | 92,400 | 72,100 | +3.25   |
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-06: Data Governance, Quality Audit & Export Wireframe", ascii_ui06)

    # Wireframe UI-07
    ascii_ui07 = """+-------------------------------------------------------------------------------------------------------+
|  SIDEBAR                        |  UI-07: SYSTEM ARCHITECTURE & PROJECT ABOUT                         |
|---------------------------------+---------------------------------------------------------------------|
| [X] About P054 Project          |  +---------------------------------------------------------------+  |
|                                 |  | [TAILWIND HEADER] Project Information & Architecture Guide    |  |
| Quick Links:                    |  +---------------------------------------------------------------+  |
| - Project Understanding (Step 1)|                                                                     |
| - Technical Blueprint (Step 7)  |  ACADEMIC & PROJECT CREDENTIALS                                     |
| - Database Design (Step 8)      |  • Candidate:  Harshit (Roll No: 2500320140058)                     |
| - ER Diagram (Step 9)           |  • Program:    MCA 2nd Year Data Analytics (DA), ABES Institute     |
| - User Stories (Step 11)        |  • Mentor:     Dinesh Sereno (Industry Lead)                        |
|                                 |  • GitHub:     https://github.com/harshit-abes/p054-sales-analysis  |
| REPOSITORY TREE:                |                                                                     |
| .streamlit/config.toml          |  FIVE-LAYER APPLICATION ARCHITECTURE PIPELINE                       |
| data/app.db                     |  [Layer 1: Storage]  SQLite relational store + CSV flat seeds        |
| src/ui.py                       |  [Layer 2: Ingest]   pandas schema validation & cleaning engine      |
| app.py                          |  [Layer 3: Cache]    st.cache_data in-memory tabular store           |
| pages/01_Overview.py            |  [Layer 4: Compute]  NumPy / SciPy variance & OLS regression engines |
| ...                             |  [Layer 5: UI/UX]    Tailwind CDN Header/Cards + Plotly Graph Objects|
+---------------------------------+---------------------------------------------------------------------+"""
    add_wireframe_block(doc, "Screen UI-07: System Architecture & Project About Wireframe", ascii_ui07)

    # Section 4: Global Layout & Responsive Design
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. Global Layout, Responsive Rules & Color Palette")
    r_h4.font.name = "Calibri"
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(30, 58, 138)

    palette_table = doc.add_table(rows=6, cols=4)
    palette_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(palette_table)
    pal_headers = ["Color Role", "Hex Code", "Tailwind Class", "Semantic Purpose"]
    for i, h in enumerate(pal_headers):
        style_header_cell(palette_table.rows[0].cells[i], h)

    pal_data = [
        ("Primary Corporate", "#1E3A8A", "bg-blue-900 / text-blue-900", "Top banners, headers, table headers, primary chart traces"),
        ("Secondary Neutral", "#F8FAFC", "bg-slate-50", "Application background, neutral card bodies, table zebra rows"),
        ("Card Surface", "#FFFFFF", "bg-white / border-slate-200", "Floating KPI card containers, Plotly chart bounding boxes"),
        ("Success Accent", "#10B981", "bg-emerald-500 / text-emerald-600", "Quota achievement >= 100%, positive variance, upward trends"),
        ("Critical Deficit", "#EF4444", "bg-rose-500 / text-rose-600", "Quota achievement < 85%, negative revenue gaps, coaching alerts")
    ]
    for r_idx, vals in enumerate(pal_data):
        row = palette_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    p_resp = doc.add_paragraph()
    p_resp.paragraph_format.space_after = Pt(8)
    p_resp.add_run(
        "Mobile & Phone View Strategy:\n"
        "• Desktop View (>1024px): 5-column KPI cards grid, 2-column paired Plotly visual layouts, full-width data tables.\n"
        "• Tablet View (768px-1024px): 3-column KPI card grid, single-column stacked Plotly charts.\n"
        "• Phone / Mobile View (<768px): 2-column wrapped KPI card grid, 100% width mobile charts with touch-friendly tooltips, scrollable data tables with sticky primary keys."
    )

    # Section 5: Reusable Tailwind Components
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run("5. Reusable UI Components Architecture (src/ui.py)")
    r_h5.font.name = "Calibri"
    r_h5.font.size = Pt(14)
    r_h5.font.bold = True
    r_h5.font.color.rgb = RGBColor(30, 58, 138)

    p_comp = doc.add_paragraph()
    p_comp.paragraph_format.space_after = Pt(8)
    p_comp.add_run(
        "To enforce UI consistency, all styling logic is encapsulated inside `src/ui.py`. "
        "Streamlit pages invoke three standardized functions:\n"
        "1. render_tailwind_header(title, subtitle, meta_tag): Renders an executive banner with live candidate/mentor metadata and subtle pulse animation.\n"
        "2. render_kpi_cards(kpi_data): Renders responsive cards with positive/negative delta pills, border highlights, and formatted currency labels.\n"
        "3. render_insight_box(title, message, alert_type): Renders callouts in success, warning, danger, or info themes with semantic icons."
    )

    # Section 6: Best Chart for each Operational Question
    h6 = doc.add_paragraph()
    h6.paragraph_format.space_before = Pt(14)
    h6.paragraph_format.space_after = Pt(4)
    r_h6 = h6.add_run("6. Best Chart Selection Rationale per Operational Question (OQ-01 to OQ-10)")
    r_h6.font.name = "Calibri"
    r_h6.font.size = Pt(14)
    r_h6.font.bold = True
    r_h6.font.color.rgb = RGBColor(30, 58, 138)

    oq_chart_data = [
        ("OQ-01", "Enterprise Target vs Actual Revenue", "Tailwind KPI Card Grid + Bullet Chart", "Provides instantaneous assessment of total enterprise quota attainment without visual clutter."),
        ("OQ-02", "Store Ranking & Deficit Identification", "Horizontal Bar Chart + Color-Coded Table", "Horizontal bars allow reading 15 store names clearly without text rotation; red highlights flags deficits instantly."),
        ("OQ-03", "Merchandise Category Contribution", "Grouped Paired Bar Chart", "Directly compares Actual vs Target side-by-side across all 5 merchandise categories in distinct colors."),
        ("OQ-04", "Day-of-Week Seasonality Patterns", "Calendar Matrix Heatmap (Mon-Sun)", "Color gradient (red-yellow-green) immediately isolates low-traffic weekdays (Tuesday/Wednesday lag)."),
        ("OQ-05", "Sales Staff Quota Attainment", "Leaderboard Table with Progress Bars", "Allows sorting 75 associates by performance; inline progress bars make underperformance (<85%) stand out."),
        ("OQ-06", "MTD Cumulative Velocity vs Target Pace", "Dual-Line Burn-Up Chart with Gap Shading", "Clear visual comparison of cumulative actual sales slope against budgeted target pace over 20 elapsed days."),
        ("OQ-07", "Promotional Discount vs Gross Margin", "Scatter Plot with OLS Linear Regression Trendline", "Exposes the trade-off and Pearson correlation (r) between aggressive discounting and profit erosion."),
        ("OQ-08", "Average Transaction Value (ATV) Benchmarks", "Angular Indicator Bullet Gauge", "Shows actual basket size against benchmark threshold with distinct red/yellow/green achievement zones."),
        ("OQ-09", "Variance Root Cause (Volume vs Price)", "Waterfall Decomposition Chart", "Mathematically proves how Volume Deficit + Price Deficit add up to the Total Dollar Deficit."),
        ("OQ-10", "Month-End Run-Rate Projections", "Dual-Bar Run-Rate Comparison + Horizon Line", "Projects month-end totals based on 20-day historical average against the required catch-up pace.")
    ]

    oq_table = doc.add_table(rows=len(oq_chart_data)+1, cols=4)
    oq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(oq_table)
    oq_headers = ["OQ ID", "Operational Question", "Best Chart / Visual Type", "Visual & Analytical Rationale"]
    for i, h in enumerate(oq_headers):
        style_header_cell(oq_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(oq_chart_data):
        row = oq_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 7: Streamlit Theme Configuration (.streamlit/config.toml)
    h7 = doc.add_paragraph()
    h7.paragraph_format.space_before = Pt(14)
    h7.paragraph_format.space_after = Pt(4)
    r_h7 = h7.add_run("7. Streamlit Global Configuration (.streamlit/config.toml)")
    r_h7.font.name = "Calibri"
    r_h7.font.size = Pt(14)
    r_h7.font.bold = True
    r_h7.font.color.rgb = RGBColor(30, 58, 138)

    toml_code = """[theme]
primaryColor = "#1E3A8A"              # Deep Corporate Navy
backgroundColor = "#F8FAFC"           # Slate-50 Light Background
secondaryBackgroundColor = "#FFFFFF"  # Pure White Containers / Sidebar
textColor = "#1E293B"                 # Slate-800 High-Contrast Text
font = "sans serif"                   # Clean Inter / System Sans

[server]
port = 8501
enableCORS = false
enableXsrfProtection = true
maxUploadSize = 50

[browser]
gatherUsageStats = false"""

    add_wireframe_block(doc, "Configuration Specification: .streamlit/config.toml", toml_code)

    doc_path = os.path.abspath("docs/12_UI_Framework.docx")
    doc.save(doc_path)
    print(f"Generated {doc_path} successfully.")

def update_rtm_step12():
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]

    # Mapping of BR ID to UI Screen (Column 11)
    ui_mapping = {
        "BR-01": "UI-01 (Executive Overview & KPI Cockpit)",
        "BR-02": "UI-01 (Store Ranking Table), UI-03 (Comparison)",
        "BR-03": "UI-03 (Store & Category Compare)",
        "BR-04": "UI-02 (Trends & Seasonality Heatmap)",
        "BR-05": "UI-03 (Staff Quota Attainment Leaderboard)",
        "BR-06": "UI-02 (MTD Cumulative Burn-Up Trajectory)",
        "BR-07": "UI-04 (Promotional Discount vs Margin Scatter)",
        "BR-08": "UI-01, UI-03 (ATV Benchmark Gauge)",
        "BR-09": "UI-04 (Volume vs Price Variance Waterfall)",
        "BR-10": "UI-05 (Predictive Run-Rate & Forecast)",
        "BR-11": "UI-06 (Data Governance & Anomaly Detection)",
        "BR-12": "Out of Scope (Phase 2)"
    }

    font_calibri = Font(name="Calibri", size=9)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for row_idx in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row_idx, column=3).value
        if br_id in ui_mapping:
            ui_val = ui_mapping[br_id]
            cell_ui = ws.cell(row=row_idx, column=11)
            cell_ui.value = ui_val
            cell_ui.font = font_calibri
            cell_ui.alignment = align_left

    wb.save(rtm_path)
    print(f"Updated {rtm_path} with UI Screen mappings successfully.")

if __name__ == "__main__":
    create_ui_framework_doc()
    update_rtm_step12()
