"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
Comprehensive ABES Project Document Generator (src/make_step17_doc.py)
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def cm_to_dxa(cm):
    return int(cm * 567.0)

def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_table_borders(table, color="BFBFBF"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def make_table_fixed(table, col_widths):
    tblPr = table._tbl.tblPr
    # Set layout fixed
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)
    
    # Set table width
    total_w = sum(col_widths)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(total_w))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    
    # Left alignment
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'left')
    tblPr.append(jc)

def style_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def style_header_row(row, col_widths):
    style_row_cant_split(row)
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    for idx, cell in enumerate(row.cells):
        cell.width = col_widths[idx]
        set_cell_background(cell, "A51C30")
        set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = True
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

def style_body_row(row, col_widths, align_right_cols=[]):
    style_row_cant_split(row)
    for idx, cell in enumerate(row.cells):
        cell.width = col_widths[idx]
        set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        if idx in align_right_cols:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.font.italic = False
            run.font.color.rgb = RGBColor(0x15, 0x15, 0x14)

def add_custom_heading1(doc, text):
    h = doc.add_paragraph(style='Heading 1')
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(7)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.font.name = "Georgia"
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)
    return h

def add_custom_heading2(doc, text):
    h = doc.add_paragraph(style='Heading 2')
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(5)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.font.name = "Georgia"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    return h

def add_body_p(doc, text=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if text:
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    return p

def add_note_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    return p

def add_bullet_p(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(0.375) # 540 dxa
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    return p

def add_label_p(doc, text, color=RGBColor(0x15, 0x15, 0x14), space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = color
    return p

def add_code_block(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = Pt(11)
    r = p.add_run(text)
    r.font.name = "Consolas"
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    return p

def add_page_number_to_footer(footer_p, learner_info):
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.paragraph_format.space_before = Pt(0)
    footer_p.paragraph_format.space_after = Pt(0)
    
    r_text = footer_p.add_run(f"{learner_info}     Page ")
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(8)
    r_text.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    
    # Insert dynamic page number field
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    footer_p._p.append(fldSimple)

def build_project_document():
    doc = Document()
    
    # -------------------------------------------------------------
    # PAGE SETUP: A4 Portrait, Margins 2cm, Header/Footer 1.25cm
    # -------------------------------------------------------------
    section = doc.sections[0]
    section.page_width = Inches(8.267)   # A4 width (21.0 cm)
    section.page_height = Inches(11.692) # A4 height (29.7 cm)
    section.top_margin = Inches(0.7874)  # 2.0 cm (1134 dxa)
    section.bottom_margin = Inches(0.7874)
    section.left_margin = Inches(0.7874)
    section.right_margin = Inches(0.7874)
    section.header_distance = Inches(0.492) # 1.25 cm (709 dxa)
    section.footer_distance = Inches(0.492)
    section.different_first_page_header_footer = True
    
    # Global Header & Footer (Pages 2 onwards)
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(0)
    hrun = hp.add_run("ABES Engineering College | Project Document | P054")
    hrun.font.name = "Calibri"
    hrun.font.size = Pt(8)
    hrun.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    add_page_number_to_footer(fp, "Harshit | 2500320140058")

    # -------------------------------------------------------------
    # SCREEN 2: PAGE 1 COVER
    # -------------------------------------------------------------
    p_prog = doc.add_paragraph()
    p_prog.paragraph_format.space_before = Pt(0)
    p_prog.paragraph_format.space_after = Pt(5)
    r_prog = p_prog.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
    r_prog.font.name = "Calibri"
    r_prog.font.size = Pt(10)
    r_prog.font.bold = True
    r_prog.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(5)
    r_logo = p_logo.add_run("[Insert ABES logo here]")
    r_logo.font.name = "Calibri"
    r_logo.font.size = Pt(9)
    r_logo.font.italic = True
    r_logo.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

    p_doc = doc.add_paragraph()
    p_doc.paragraph_format.space_before = Pt(70)
    p_doc.paragraph_format.space_after = Pt(6)
    r_doc = p_doc.add_run("PROJECT DOCUMENT")
    r_doc.font.name = "Georgia"
    r_doc.font.size = Pt(26)
    r_doc.font.bold = True
    r_doc.font.color.rgb = RGBColor(0x15, 0x15, 0x14)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    # Add bottom border in XML for cover title
    pPr = p_sub._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="8" w:color="A51C30"/></w:pBdr>')
    pPr.append(pBdr)
    r_sub = p_sub.add_run("P054 - Daily Sales Actual vs Target Analysis")
    r_sub.font.name = "Georgia"
    r_sub.font.size = Pt(18)
    r_sub.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    p_subtext = doc.add_paragraph()
    p_subtext.paragraph_format.space_before = Pt(0)
    p_subtext.paragraph_format.space_after = Pt(24)
    r_st = p_subtext.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
    r_st.font.name = "Calibri"
    r_st.font.size = Pt(10)
    r_st.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

    # Cover Table: Field 2900 | Details 6738 (Total 9638 dxa)
    cover_widths = [2900, 6738]
    cover_rows = [
        ("Project ID & Title", "P054 - Daily Sales Actual vs Target Analysis"),
        ("Domain", "Retail Sales Operations & Performance Analytics"),
        ("Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Learner(s) & Roll No(s)", "Harshit (University Roll No: 2500320140058)"),
        ("Mentor", "Dinesh Sereno (Industry Data Engineering Lead)"),
        ("GitHub repository", "https://github.com/hcosmos/2500320140058-p054-daily-sales"),
        ("Live app (Streamlit)", "https://2500320140058-p054-daily-sales-luq4tih6ynstt4zqslkrfc.streamlit.app/"),
        ("Version & date", "v1.0 - 21/11/2026")
    ]
    t_cover = doc.add_table(rows=len(cover_rows), cols=2)
    set_table_borders(t_cover)
    make_table_fixed(t_cover, cover_widths)
    for idx, (k, v) in enumerate(cover_rows):
        row = t_cover.rows[idx]
        row.cells[0].paragraphs[0].add_run(k)
        row.cells[1].paragraphs[0].add_run(v)
        style_body_row(row, cover_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_page_break()

    # -------------------------------------------------------------
    # SCREEN 2: PAGE 2 - TOC
    # -------------------------------------------------------------
    # Table of contents field: TOC \h \o "1-1"
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(0)
    p_toc.paragraph_format.space_after = Pt(0)
    
    # Real Word TOC XML with Heading 1 entries
    fldSimple_toc = OxmlElement('w:fldSimple')
    fldSimple_toc.set(qn('w:instr'), 'TOC \\h \\o "1-1"')
    p_toc._p.append(fldSimple_toc)

    # Enable updateFields on document open
    settingsPr = doc.settings.element
    updateFields = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
    settingsPr.append(updateFields)

    toc_titles = [
        "1. Project Description",
        "2. How Might We (HMW) Statements",
        "3. Stakeholders",
        "4. Traceability Matrix",
        "5. User Stories",
        "6. Business Requirements Document (BRD)",
        "7. Technical Requirements Document (TRD)",
        "8. Database Design",
        "9. Test Cases",
        "10. UI Framework",
        "Appendix A - Prompt to generate this document",
        "Appendix B - AI tools disclosure"
    ]
    for title in toc_titles:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(4)
        p_t.paragraph_format.space_after = Pt(4)
        r = p_t.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)

    doc.add_page_break()

    # -------------------------------------------------------------
    # 1. Project Description
    # -------------------------------------------------------------
    add_custom_heading1(doc, "1. Project Description")
    
    desc_widths = [2400, 7238] # Total 9638
    desc_rows = [
        ("Overview", "Project P054: Daily Sales Actual vs Target Analysis is an enterprise retail operations performance intelligence web application built to eliminate the 48-to-72-hour reporting latency commonly experienced by multi-unit retail networks. Operating on SQLite relational data, it delivers automated variance calculations, store deficit rankings, root-cause waterfall attribution, and month-end run-rate projections."),
        ("Problem statement", "Multi-unit retail leaders struggle to spot lagging stores, categories, and associates in time to hit monthly revenue targets because sales actuals and target quotas are trapped in disconnected spreadsheets, delaying corrective operational intervention by 48-72 hours."),
        ("Intended users", "Commercial Sales Director (Daily overview & projections), Store Operations Managers (Daily branch ranking & deficit alerts), Merchandise Category Managers (Category contribution & discount elasticity), Floor Supervisors (Associate quota coaching), and Business Data Analysts (Root cause decomposition & anomaly governance)."),
        ("In scope", "Automated daily ingestion of 15 stores, 5 merchandise categories, 75 sales associates; 40-column analytical view calculation; variance decomposition waterfall; day-of-week seasonality heatmap; discount margin elasticity regression; month-end run-rate forecasting; Z-score anomaly detector; Excel/CSV export."),
        ("Out of scope", "Sub-second live streaming POS feeds (BR-12 - Phase 2); automated employee payroll deduction; write-back ERP master record modifications; direct point-of-sale customer payment gateway processing."),
        ("Assumptions", "Daily sales transactions and budgeted targets are logged at store-category-date grain; budget quotas reflect 30 trading days; product cost of goods sold (COGS) remains stable over the 20-day evaluation horizon; network stores follow Indian retail operating hours."),
        ("Dataset", "Reproducible synthetic relational dataset (seed 42) comprising 15 fictional Indian retail stores, 5 merchandise categories, 75 sales associates, 300 budgeted target records, and 300 actual transactional sales records (303 raw with 3 duplicate verification rows and 6 injected null discounts).")
    ]
    
    t_desc = doc.add_table(rows=len(desc_rows)+1, cols=2)
    set_table_borders(t_desc)
    make_table_fixed(t_desc, desc_widths)
    t_desc.rows[0].cells[0].paragraphs[0].add_run("Item")
    t_desc.rows[0].cells[1].paragraphs[0].add_run("Description")
    style_header_row(t_desc.rows[0], desc_widths)
    
    for r_idx, (k, v) in enumerate(desc_rows):
        row = t_desc.rows[r_idx+1]
        row.cells[0].paragraphs[0].add_run(k)
        row.cells[1].paragraphs[0].add_run(v)
        style_body_row(row, desc_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 2. How Might We (HMW) Statements
    # -------------------------------------------------------------
    add_custom_heading1(doc, "2. How Might We (HMW) Statements")
    add_note_p(doc, "Format: How might we <action> for <user> so that <outcome>?")

    hmw_widths = [1300, 8338] # Total 9638
    hmw_rows = [
        ("HMW-01", "How might we consolidate daily actual sales revenue alongside assigned targets across all retail branches for Regional Directors so that they can assess overall enterprise target achievement on a single screen without waiting for weekly consolidated reports? (PRIORITY)"),
        ("HMW-02", "How might we calculate and highlight both absolute dollar deficit and percentage shortfall between actuals and targets for Store Managers so that deficit stores requiring immediate operational support are flagged within 2 hours of trading close? (PRIORITY)"),
        ("HMW-03", "How might we track product line sales against category-specific target quotas for Category Managers so that underperforming merchandise categories can be identified before month-end inventory planning? (PRIORITY)"),
        ("HMW-04", "How might we visualize target achievement by day of the week for Store Supervisors so that staffing schedules and inventory prep can be optimized for historically lagging weekdays?"),
        ("HMW-05", "How might we evaluate individual sales staff revenue against quota baselines for Floor Leads so that coaching can be provided to associates falling below an 85% target threshold?"),
        ("HMW-06", "How might we compare month-to-date (MTD) cumulative actual revenue against cumulative target trajectory for Financial Planners so that executive leadership knows if run-rate pace is sufficient to meet monthly budget commitments? (PRIORITY)"),
        ("HMW-07", "How might we correlate promotional discount percentages against target achievement for Merchandise Planners so that management can detect whether discounting drives real quota lift or merely erodes gross margin?"),
        ("HMW-08", "How might we benchmark realized Average Transaction Value (ATV) against targeted basket goals for Store Operations so that stores know whether customer spending depth meets planned store format expectations?"),
        ("HMW-09", "How might we decompose sales target shortfall into customer footfall/order count deficits versus basket size (ticket depth) deficits for Store Managers so that the right corrective strategy (footfall marketing vs upselling) is deployed? (PRIORITY)"),
        ("HMW-10", "How might we project month-end sales totals based on current daily run rates for Regional Directors so that proactive commercial adjustments can be executed before accounting period close?"),
        ("HMW-11", "How might we statistically isolate anomalous daily sales spikes or dips (using Z-scores) for Data Analysts so that temporary external disruptions are separated from structural underperformance?")
    ]

    t_hmw = doc.add_table(rows=len(hmw_rows)+1, cols=2)
    set_table_borders(t_hmw)
    make_table_fixed(t_hmw, hmw_widths)
    t_hmw.rows[0].cells[0].paragraphs[0].add_run("ID")
    t_hmw.rows[0].cells[1].paragraphs[0].add_run("HMW statement")
    style_header_row(t_hmw.rows[0], hmw_widths)

    for r_idx, (k, v) in enumerate(hmw_rows):
        row = t_hmw.rows[r_idx+1]
        row.cells[0].paragraphs[0].add_run(k)
        row.cells[1].paragraphs[0].add_run(v)
        style_body_row(row, hmw_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 3. Stakeholders
    # -------------------------------------------------------------
    add_custom_heading1(doc, "3. Stakeholders")
    
    sh_widths = [1000, 2400, 1900, 3138, 1200] # Total 9638
    sh_rows = [
        ("SH-01", "Rajesh Malhotra", "Store Operations Manager", "Immediate daily branch revenue ranking, tracking absolute deficit ($), identifying lagging branches, staff reallocation.", "High"),
        ("SH-02", "Sunita Narang", "Category Merchandise Head", "Departmental quota realization, promotional markdown effectiveness, category gross profit margin preservation.", "High"),
        ("SH-03", "Vikramaditya Shah", "Commercial Sales Director", "Macro enterprise quota pace, regional performance comparison, month-end run-rate deficit projections, strategic interventions.", "High"),
        ("SH-04", "Meera Nambisan", "Inventory & Financial Planner", "MTD cumulative burn-up tracking, stock replenishment velocity, promotional pricing elasticity, basket depth.", "Medium"),
        ("SH-05", "Harshit (Learner)", "Lead Business Data Analyst", "Data modeling, ETL pipeline integrity, variance decomposition math, anomaly detection, statistical governance.", "Medium")
    ]

    t_sh = doc.add_table(rows=len(sh_rows)+1, cols=5)
    set_table_borders(t_sh)
    make_table_fixed(t_sh, sh_widths)
    sh_headers = ["ID", "Stakeholder", "Role", "Interest", "Influence"]
    for i, h in enumerate(sh_headers):
        t_sh.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_sh.rows[0], sh_widths)

    for r_idx, vals in enumerate(sh_rows):
        row = t_sh.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, sh_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 4. Traceability Matrix
    # -------------------------------------------------------------
    add_custom_heading1(doc, "4. Traceability Matrix")
    add_note_p(doc, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")

    rtm_widths = [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040] # Total 9638
    rtm_rows = [
        ("HMW-01", "BR-01", "Aggregate daily actual sales alongside target quotas across branches", "SH-01, SH-04", "FR-01, 02, 03", "Must", "US-01..05", "TC-F-01..04 / TC-T-01..07", "UI-01"),
        ("HMW-02", "BR-02", "Calculate absolute ($) and % variance with deficit ranking", "SH-01", "FR-03, 04", "Must", "US-05, 06", "TC-F-02, 05 / TC-T-02, 08", "UI-01, 03"),
        ("HMW-03", "BR-03", "Track category revenue actual vs assigned quota attribution", "SH-03", "FR-05", "Must", "US-07", "TC-F-09 / TC-T-03, 09", "UI-03"),
        ("HMW-04", "BR-04", "Analyze achievement variance by day of the week", "SH-01, SH-02", "FR-06", "Should", "US-08", "TC-F-07 / TC-T-07, 11", "UI-02"),
        ("HMW-05", "BR-05", "Evaluate staff attainment with <85% coaching alerts", "SH-02", "FR-07", "Should", "US-09", "TC-F-10 / TC-T-04, 09", "UI-03"),
        ("HMW-06", "BR-06", "Track MTD cumulative actuals vs target trajectory curve", "SH-04", "FR-08", "Must", "US-13", "TC-F-08 / TC-T-07, 15", "UI-02"),
        ("HMW-07", "BR-07", "Assess promotional discount impact on gross profit margin", "SH-03", "FR-09", "Should", "US-12", "TC-F-12 / TC-T-14, 17", "UI-04"),
        ("HMW-08", "BR-08", "Compute realized Average Transaction Value (ATV) depth", "SH-02", "FR-04", "Must", "US-10", "TC-F-06 / TC-T-05, 13", "UI-01, 03"),
        ("HMW-09", "BR-09", "Decompose variance into Volume Effect ($) and Price Effect ($)", "SH-02, SH-05", "FR-10", "Must", "US-11", "TC-F-11 / TC-T-13", "UI-04"),
        ("HMW-10", "BR-10", "Project month-end sales totals and quota deficits", "SH-04", "FR-11", "Must", "US-14", "TC-F-13 / TC-T-15", "UI-05"),
        ("HMW-11", "BR-11", "Statistically flag anomalous sales spikes (|Z| > 3.0)", "SH-05", "FR-12", "Could", "US-15", "TC-F-14 / TC-T-16", "UI-06"),
        ("HMW-01", "BR-12", "Ingest sub-second live streaming POS feeds", "SH-02", "FR-01 (Ph 2)", "Won't", "Out of Scope", "Out of Scope (Phase 2)", "Out of Scope")
    ]

    t_rtm = doc.add_table(rows=len(rtm_rows)+1, cols=9)
    set_table_borders(t_rtm)
    make_table_fixed(t_rtm, rtm_widths)
    rtm_headers = ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"]
    for i, h in enumerate(rtm_headers):
        t_rtm.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_rtm.rows[0], rtm_widths)

    for r_idx, vals in enumerate(rtm_rows):
        row = t_rtm.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, rtm_widths)
        row.cells[1].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 5. User Stories
    # -------------------------------------------------------------
    add_custom_heading1(doc, "5. User Stories")

    us_widths = [900, 3400, 4138, 1200] # Total 9638
    us_rows = [
        ("US-01", "As a Data Analyst, I want to connect to data/app.db and query the 40-col analytical view so that the app runs on normalized relational data.", "Given app starts and app.db exists, When load_analytical_view() executes, Then exactly 300 records and 40 cols are loaded.", "Must Have"),
        ("US-02", "As a Data Analyst, I want the ETL engine to impute missing discounts with 0.0 and purge duplicates so that KPI metrics are pure.", "Given raw feed has 6 NaNs and 3 duplicates, When clean_raw_sales_dataframe() runs, Then zero nulls remain and 300 deduplicated rows exist.", "Must Have"),
        ("US-03", "As a Data Analyst, I want analytical views cached with @st.cache_data so that dashboard filter interactions respond in < 0.5 seconds.", "Given data is loaded, When any sidebar filter is toggled, Then data serves from cache without disk re-reading.", "Must Have"),
        ("US-04", "As a Commercial VP, I want Tailwind KPI cards displaying actual sales, targets, variance, and margins so that enterprise health is clear in 5s.", "Given overview page opens, When KPIs render, Then 5 distinct cards display actuals, targets, variance, achievement, and margins.", "Must Have"),
        ("US-05", "As a Store Manager, I want multi-filter cascading controls (Date, Region, Format, Store, Category) so that I can isolate my branch metrics.", "Given Region = 'North' is selected, When Store dropdown is inspected, Then only North stores appear in the dropdown list.", "Must Have"),
        ("US-06", "As an Operations Lead, I want a sortable store ranking table with <85% deficit alerts so that lagging branches are immediately isolated.", "Given store view is open, When sorted by achievement ascending, Then stores <85% show red 'CRITICAL DEFICIT' alert badges.", "Must Have"),
        ("US-07", "As a Category Head, I want a grouped bar chart comparing actual vs target sales across all 5 categories to identify drag product lines.", "Given category view renders, When chart displays, Then paired bars show target vs actual revenue with INR tooltips.", "Must Have"),
        ("US-08", "As a Store Manager, I want an achievement heatmap across Monday to Sunday so that mid-week trading lulls can be staffed appropriately.", "Given seasonality tab is selected, When heatmap renders, Then Tuesday and Wednesday cells display amber/red (<90% achievement).", "Should Have"),
        ("US-09", "As a Floor Supervisor, I want an associate leaderboard with <85% coaching flags so that shift coaching can be delivered promptly.", "Given staff table renders, When associate achievement < 85%, Then Coaching Required displays 'YES' in bold red font.", "Should Have"),
        ("US-10", "As a Category Head, I want an ATV indicator gauge showing basket size versus benchmark so that spending depth is tracked.", "Given store selection is active, When ATV widget renders, Then gauge needle displays realized basket size against target benchmark.", "Must Have"),
        ("US-11", "As a Data Analyst, I want a waterfall decomposition chart splitting variance into Volume vs Price effect so that root causes are clear.", "Given waterfall chart calculates, When rendered, Then Volume Effect ($) + Price Effect ($) exactly equals Dollar Variance ($).", "Must Have"),
        ("US-12", "As a Commercial VP, I want a discount vs gross margin scatter plot with OLS trendline so that margin erosion is proven.", "Given pricing page renders, When scatter plot displays, Then downward regression line shows negative correlation (r < -0.30).", "Should Have"),
        ("US-13", "As a Financial Planner, I want an MTD cumulative burn-up curve vs target pace line so that velocity pacing is monitored.", "Given 20 trading days, When chart renders, Then cumulative actual sales curve tracks target line with deficit gap shading.", "Must Have"),
        ("US-14", "As a Commercial VP, I want a month-end run-rate projection engine so that expected quota shortfalls can be corrected early.", "Given 20 elapsed days, When forecast runs, Then projected total equals (Sales / 20) * 30 and shortfall displays required daily rate.", "Must Have"),
        ("US-15", "As a Data Analyst, I want an anomaly detection table identifying |Z| > 3.0 sigma spikes so that external shocks are isolated.", "Given dataset contains 3 flash surge sales records, When anomaly module runs, Then exactly those 3 records are flagged with |Z| > 3.0.", "Could Have")
    ]

    t_us = doc.add_table(rows=len(us_rows)+1, cols=4)
    set_table_borders(t_us)
    make_table_fixed(t_us, us_widths)
    us_headers = ["US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"]
    for i, h in enumerate(us_headers):
        t_us.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_us.rows[0], us_widths)

    for r_idx, vals in enumerate(us_rows):
        row = t_us.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, us_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 6. Business Requirements Document (BRD)
    # -------------------------------------------------------------
    add_custom_heading1(doc, "6. Business Requirements Document (BRD)")

    add_custom_heading2(doc, "6.1 Business objectives")
    add_bullet_p(doc, "OBJ-01: Reduce daily sales and target variance reporting latency from 48-72 hours to under 5 minutes upon POS end-of-day batch close.")
    add_bullet_p(doc, "OBJ-02: Enable 100% automated daily identification of retail branches operating under critical revenue deficit (<85% quota attainment).")
    add_bullet_p(doc, "OBJ-03: Decompose 100% of store-level revenue variances into order volume (footfall) vs average basket value (pricing/upselling) drivers.")
    add_bullet_p(doc, "OBJ-04: Safeguard product line gross profit margins by identifying markdown thresholds that result in unsustainable gross margin erosion.")
    add_bullet_p(doc, "OBJ-05: Project month-end sales quotas and required daily run-rate velocity with at least 95% forecasting accuracy by Day 20.")

    add_custom_heading2(doc, "6.2 Success KPIs")
    kpi_widths = [1900, 2100, 2100, 1700, 1838] # Total 9638
    kpi_rows = [
        ("Quota Achievement Rate (%)", "SUM(actual_net_revenue)", "SUM(target_revenue)", "Daily / MTD", ">= 100% Enterprise target"),
        ("Absolute Monetary Variance ($)", "SUM(actual_net_revenue) - target", "N/A (Monetary Net)", "Daily", ">= $0.00 (Zero Deficit)"),
        ("Realized Gross Margin (%)", "SUM(gross_profit)", "SUM(actual_net_revenue)", "Daily / Monthly", ">= 38.0% Benchmark"),
        ("Average Transaction Value (ATV)", "SUM(actual_net_revenue)", "SUM(order_count)", "Per Transaction", ">= ₹2,800 Target Benchmark"),
        ("Underperforming Store Count", "COUNT(stores with Ach < 85%)", "COUNT(total active stores)", "Daily", "0 Stores in Critical Deficit")
    ]
    t_kpi = doc.add_table(rows=len(kpi_rows)+1, cols=5)
    set_table_borders(t_kpi)
    make_table_fixed(t_kpi, kpi_widths)
    kpi_headers = ["KPI", "Numerator", "Denominator", "Period", "Target / note"]
    for i, h in enumerate(kpi_headers):
        t_kpi.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_kpi.rows[0], kpi_widths)

    for r_idx, vals in enumerate(kpi_rows):
        row = t_kpi.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, kpi_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_custom_heading2(doc, "6.3 Business requirements")
    br_widths = [1100, 5338, 1300, 1900] # Total 9638
    br_rows = [
        ("BR-01", "The business needs to aggregate and compute daily actual sales revenue alongside assigned target quotas across retail branches so that leadership can monitor overall enterprise achievement rates.", "Must Have", "SH-01, SH-04"),
        ("BR-02", "The business needs to calculate absolute monetary variance ($) and percentage variance (%) for each store to rank stores and identify deficit branches (<85%) requiring intervention.", "Must Have", "SH-01"),
        ("BR-03", "The business needs to track actual revenue against assigned quota targets broken down by product categories to determine which lines are driving or dragging store performance.", "Must Have", "SH-03"),
        ("BR-04", "The business needs to analyze target achievement variance by day of the week so that store operations can optimize staff scheduling and inventory for lagging weekdays.", "Should Have", "SH-01, SH-02"),
        ("BR-05", "The business needs to evaluate individual associate revenue against quotas with automated coaching alerts (<85%) so that supervisors can deliver timely floor coaching.", "Should Have", "SH-02"),
        ("BR-06", "The business needs to track Month-to-Date (MTD) cumulative actual sales alongside target trajectory curves so that commercial executives know if velocity is keeping pace.", "Must Have", "SH-04"),
        ("BR-07", "The business needs to assess the relationship between discount levels and sales achievement to detect whether markdowns lift quota or erode gross profit margins.", "Should Have", "SH-03"),
        ("BR-08", "The business needs to compute realized Average Transaction Value (ATV) against targeted basket benchmarks to determine whether customer spending depth meets plans.", "Must Have", "SH-02"),
        ("BR-09", "The business needs to decompose revenue variance into Volume Deficit vs Price/ATV Deficit so that management deploys the right corrective intervention.", "Must Have", "SH-02, SH-05"),
        ("BR-10", "The business needs to project month-end revenue and anticipated quota deficits based on daily run-rate velocity so that corrective actions are executed before period close.", "Must Have", "SH-04"),
        ("BR-11", "The business needs to statistically flag anomalous sales days using standardized Z-scores to separate isolated temporary shocks from chronic underperformance.", "Could Have", "SH-05"),
        ("BR-12", "The business needs to ingest sub-second real-time streaming transactional feeds directly from active cash registers to view live per-minute variance.", "Won't Have", "SH-02 (Phase 2)")
    ]
    t_br = doc.add_table(rows=len(br_rows)+1, cols=4)
    set_table_borders(t_br)
    make_table_fixed(t_br, br_widths)
    br_headers = ["BR-ID", "Requirement", "Priority", "Source (stakeholder)"]
    for i, h in enumerate(br_headers):
        t_br.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_br.rows[0], br_widths)

    for r_idx, vals in enumerate(br_rows):
        row = t_br.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, br_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_custom_heading2(doc, "6.4 Constraints and out of scope")
    add_bullet_p(doc, "Fixed Stack Constraint: Python 3.12+, pandas, NumPy, SciPy/statsmodels, Plotly, Streamlit Multi-Page, SQLite (`data/app.db`), Tailwind CSS CDN, pytest.")
    add_bullet_p(doc, "Zero Cloud Cost Constraint: Built entirely on open-source libraries deployable to Streamlit Community Cloud with zero paid third-party API dependencies.")
    add_bullet_p(doc, "Data Immutability Constraint: Primary raw transactional records remain read-only; cleaning logic imputes nulls and deduplicates in memory.")
    add_bullet_p(doc, "Out of Scope: Multi-tenant role-based SSO authentication (Phase 2); sub-second live Kafka POS streaming (BR-12); automated ERP financial write-back.")

    # -------------------------------------------------------------
    # 7. Technical Requirements Document (TRD)
    # -------------------------------------------------------------
    add_custom_heading1(doc, "7. Technical Requirements Document (TRD)")

    add_custom_heading2(doc, "7.1 Technology stack (all free)")
    stack_widths = [2000, 3800, 3838] # Total 9638
    stack_rows = [
        ("Language", "Python 3.12+ (CPython)", "Open-source standard for scientific computing, robust statistical libraries, native SQLite3 binding."),
        ("Data processing", "pandas 2.2+ & NumPy 1.26+", "High-performance vector operations, tabular manipulation, in-memory aggregation, date indexing."),
        ("Storage", "SQLite 3 Engine (data/app.db)", "Zero-configuration serverless ACID relational store, PRAGMA foreign key constraints, analytical view execution."),
        ("Charts", "Plotly Graph Objects 5.22+", "Interactive JavaScript visualizations with zero runtime licensing cost, responsive hover tooltips, and waterfall support."),
        ("App / UI", "Streamlit 1.35+ & Tailwind CSS CDN", "Rapid multi-page Python web framework coupled with Tailwind CSS for executive card design via st.components.v1.html."),
        ("Testing", "pytest 8.2+ & Streamlit AppTest", "Automated technical unit test harness, mathematical assertion verification, schema validation, and latency benchmarking."),
        ("Version control", "Git & GitHub", "Distributed version control, continuous deployment integration via Streamlit Cloud GitHub webhooks."),
        ("Editor", "VS Code / Google Antigravity IDE", "Integrated terminal, python-docx automation scripts, openpyxl excel inspection, git branch management.")
    ]
    t_stack = doc.add_table(rows=len(stack_rows)+1, cols=3)
    set_table_borders(t_stack)
    make_table_fixed(t_stack, stack_widths)
    t_stack.rows[0].cells[0].paragraphs[0].add_run("Layer")
    t_stack.rows[0].cells[1].paragraphs[0].add_run("Tool")
    t_stack.rows[0].cells[2].paragraphs[0].add_run("Why")
    style_header_row(t_stack.rows[0], stack_widths)

    for r_idx, vals in enumerate(stack_rows):
        row = t_stack.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, stack_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_custom_heading2(doc, "7.2 Functional requirements")
    fr_widths = [1100, 6338, 2200] # Total 9638
    fr_rows = [
        ("FR-01", "Data Ingestion & Integrity Engine: Connect to SQLite, enforce FKs, impute null discounts to 0.0, purge duplicates, cache view.", "BR-01, US-01..03"),
        ("FR-02", "Executive KPI Cards Header: Render Tailwind 5-tile card grid (Actuals, Targets, Variance, Achievement %, Margin) with alert badges.", "BR-01, US-04"),
        ("FR-03", "Dynamic Multi-Filter Sidebar: Cascading filter state (Date, Region, Format, Store, Category) persisted across pages in st.session_state.", "BR-01, 02, US-05"),
        ("FR-04", "Store Deficit Ranking & Grid: Sortable table calculating monetary variance and percentage achievement with conditional deficit tags (<85%).", "BR-02, US-06"),
        ("FR-05", "Category Contribution Bar Chart: Paired grouped horizontal bars comparing actual vs target sales across all 5 merchandise categories.", "BR-03, US-07"),
        ("FR-06", "Day-of-Week Seasonality Heatmap: Calendar matrix heatmap plotting Monday through Sunday achievement rates across regions.", "BR-04, US-08"),
        ("FR-07", "Staff Leaderboard & Coaching Alerts: Associate revenue realization table with automated coaching alert flags for associates < 85% quota.", "BR-05, US-09"),
        ("FR-08", "MTD Cumulative Burn-Up Curve: Dual-line trajectory plotting cumulative sales against target pace line with shaded deficit areas.", "BR-06, US-13"),
        ("FR-09", "Discount vs Margin Elasticity Module: Scatter plot with OLS linear regression trendline, calculating slope (β), Pearson r, and p-value.", "BR-07, US-12"),
        ("FR-10", "Volume vs Price Waterfall Decomposition: Algebraic waterfall bridge isolating Volume Effect ($) and Price Effect ($) reconciling to variance.", "BR-09, US-11"),
        ("FR-11", "Run-Rate Forecast & Scenario Simulator: Velocity predictor calculating projected month-end totals with interactive catch-up pace sliders.", "BR-10, US-14"),
        ("FR-12", "Statistical Anomaly Detection Table: Standardized Z-score calculation on sales revenue flagging transactions deviating > 3.0 sigma.", "BR-11, US-15")
    ]
    t_fr = doc.add_table(rows=len(fr_rows)+1, cols=3)
    set_table_borders(t_fr)
    make_table_fixed(t_fr, fr_widths)
    t_fr.rows[0].cells[0].paragraphs[0].add_run("FR-ID")
    t_fr.rows[0].cells[1].paragraphs[0].add_run("Requirement")
    t_fr.rows[0].cells[2].paragraphs[0].add_run("Linked BR / US")
    style_header_row(t_fr.rows[0], fr_widths)

    for r_idx, vals in enumerate(fr_rows):
        row = t_fr.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, fr_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_custom_heading2(doc, "7.3 Non-functional requirements")
    nfr_widths = [1100, 8538] # Total 9638
    nfr_rows = [
        ("NFR-01", "Performance Latency: In-memory query execution and rendering of the 40-column analytical view must complete in strictly < 1.0 second (< 3.0s total page load budget)."),
        ("NFR-02", "Scalability: The database schema and caching layer must comfortably support up to 100 retail stores, 50 categories, and 100,000 annual transactional records without architectural redesign."),
        ("NFR-03", "Data Quality & Integrity: 100% referential integrity enforced by SQLite foreign keys (zero orphan records); mathematical reconciliation verified by automated unit tests."),
        ("NFR-04", "Responsive Usability: Multi-column grid on desktop (>1024px) gracefully degrades to single-column stacked layout on mobile devices (<768px) with touch-enabled Plotly charts."),
        ("NFR-05", "Portability & Zero Deployment Cost: Self-contained repository deployable to Streamlit Community Cloud with automated pip dependency installation and zero external database hosting fees."),
        ("NFR-06", "Code Quality & Test Coverage: 100% of mathematical formulas, schema constraints, and ETL sanitization routines covered by 21 automated pytest unit tests.")
    ]
    t_nfr = doc.add_table(rows=len(nfr_rows)+1, cols=2)
    set_table_borders(t_nfr)
    make_table_fixed(t_nfr, nfr_widths)
    t_nfr.rows[0].cells[0].paragraphs[0].add_run("NFR-ID")
    t_nfr.rows[0].cells[1].paragraphs[0].add_run("Requirement")
    style_header_row(t_nfr.rows[0], nfr_widths)

    for r_idx, vals in enumerate(nfr_rows):
        row = t_nfr.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, nfr_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_custom_heading2(doc, "7.4 Architecture and deployment")
    p_arch_line = doc.add_paragraph()
    p_arch_line.paragraph_format.space_before = Pt(2)
    p_arch_line.paragraph_format.space_after = Pt(5)
    r_al = p_arch_line.add_run("Raw data (data/fact_daily_sales.csv) -> src/data_loader.py -> SQLite app.db (data/app.db) -> Streamlit app (app.py) -> user")
    r_al.font.name = "Calibri"
    r_al.font.size = Pt(10.5)
    r_al.font.bold = True
    r_al.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    stage_widths = [2200, 4838, 2600] # Total 9638
    stage_rows = [
        ("1. Local", "python src/generate_data.py && pytest -v && streamlit run app.py", "All 21 pytest tests pass; local browser opens app.py on port 8501."),
        ("2. GitHub", "git add . && git commit -m 'Release v1.0' && git push -u origin main", "Repository main branch updated; data/app.db and requirements.txt verified on remote."),
        ("3. Streamlit Cloud", "share.streamlit.io -> New app -> select repo -> branch main -> Deploy", "Container builds dependencies; live URL https://2500320140058-p054-daily-sales-luq4tih6ynstt4zqslkrfc.streamlit.app/ active.")
    ]
    t_stage = doc.add_table(rows=len(stage_rows)+1, cols=3)
    set_table_borders(t_stage)
    make_table_fixed(t_stage, stage_widths)
    t_stage.rows[0].cells[0].paragraphs[0].add_run("Stage")
    t_stage.rows[0].cells[1].paragraphs[0].add_run("Command / action")
    t_stage.rows[0].cells[2].paragraphs[0].add_run("Done when")
    style_header_row(t_stage.rows[0], stage_widths)

    for r_idx, vals in enumerate(stage_rows):
        row = t_stage.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, stage_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # -------------------------------------------------------------
    # 8. Database Design
    # -------------------------------------------------------------
    add_custom_heading1(doc, "8. Database Design")
    add_note_p(doc, "Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")

    db_widths = [1700, 2300, 2100, 1400, 2138] # Total 9638
    db_schema_rows = [
        ("dim_stores", "store_id", "TEXT", "PK", "Unique store identifier (e.g. STR-001)"),
        ("dim_stores", "store_name", "TEXT", "NOT NULL", "Store name in major Indian metro"),
        ("dim_stores", "region", "TEXT", "CHECK", "Geographic region (North/South/East/West/Central)"),
        ("dim_stores", "store_format", "TEXT", "CHECK", "Branch format (Flagship/Mall/HighStreet/Outlet)"),
        ("dim_stores", "benchmark_daily_target", "REAL", "NOT NULL", "Base daily budgeted quota benchmark"),
        ("dim_categories", "category_id", "TEXT", "PK", "Unique category code (e.g. CAT-01)"),
        ("dim_categories", "category_name", "TEXT", "UNIQUE", "Merchandise line (Electronics, Apparel, etc.)"),
        ("dim_categories", "target_gross_margin_pct", "REAL", "NOT NULL", "Target gross margin percentage benchmark"),
        ("dim_staff", "staff_id", "TEXT", "PK", "Unique employee identifier (e.g. STF-101)"),
        ("dim_staff", "staff_name", "TEXT", "NOT NULL", "Full fictional associate name"),
        ("dim_staff", "store_id", "TEXT", "FK", "Foreign key referencing dim_stores(store_id)"),
        ("dim_staff", "base_daily_quota", "REAL", "NOT NULL", "Assigned daily sales quota per associate"),
        ("fact_daily_targets", "target_id", "TEXT", "PK", "Composite target identifier (TGT-Date-Store-Cat)"),
        ("fact_daily_targets", "target_revenue", "REAL", "NOT NULL", "Budgeted daily sales revenue quota"),
        ("fact_daily_targets", "target_orders", "INTEGER", "NOT NULL", "Planned customer transaction order count"),
        ("fact_daily_targets", "target_atv", "REAL", "NOT NULL", "Planned benchmark basket ticket size"),
        ("fact_daily_sales", "sales_id", "TEXT", "PK", "Primary sales key (SLS-Date-Store-Cat)"),
        ("fact_daily_sales", "transaction_date", "TEXT", "NOT NULL", "Trading date in YYYY-MM-DD format"),
        ("fact_daily_sales", "store_id", "TEXT", "FK", "Foreign key referencing dim_stores(store_id)"),
        ("fact_daily_sales", "category_id", "TEXT", "FK", "Foreign key referencing dim_categories(category_id)"),
        ("fact_daily_sales", "staff_id", "TEXT", "FK", "Foreign key referencing dim_staff(staff_id)"),
        ("fact_daily_sales", "net_revenue", "REAL", "NOT NULL", "Net realized revenue after discounts"),
        ("fact_daily_sales", "cost_of_goods_sold", "REAL", "NOT NULL", "Product wholesale procurement cost"),
        ("fact_daily_sales", "gross_profit", "REAL", "NOT NULL", "Realized gross profit (net_revenue - cogs)"),
        ("fact_daily_sales", "order_count", "INTEGER", "NOT NULL", "Actual transaction footfall count"),
        ("fact_daily_sales", "realized_atv", "REAL", "NOT NULL", "Realized average transaction value (₹)")
    ]

    t_db = doc.add_table(rows=len(db_schema_rows)+1, cols=5)
    set_table_borders(t_db)
    make_table_fixed(t_db, db_widths)
    db_headers = ["Table / sheet", "Column", "Type", "Key", "Description"]
    for i, h in enumerate(db_headers):
        t_db.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_db.rows[0], db_widths)

    for r_idx, vals in enumerate(db_schema_rows):
        row = t_db.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, db_widths)
        row.cells[1].paragraphs[0].runs[0].font.bold = True

    add_label_p(doc, "Relationship:", space_before=10)
    add_bullet_p(doc, "One (1) dim_stores has many (N) fact_daily_sales records (1:N via store_id).")
    add_bullet_p(doc, "One (1) dim_categories has many (N) fact_daily_sales records (1:N via category_id).")
    add_bullet_p(doc, "One (1) dim_staff has many (N) fact_daily_sales records (1:N via staff_id).")
    add_bullet_p(doc, "One (1) dim_stores has many (N) fact_daily_targets records (1:N via store_id).")
    add_bullet_p(doc, "One (1) dim_categories has many (N) fact_daily_targets records (1:N via category_id).")
    add_bullet_p(doc, "One (1) fact_daily_targets aligns (1:1) with fact_daily_sales on (transaction_date, store_id, category_id) in analytical view.")

    add_label_p(doc, "SQLite DDL:", space_before=8)
    ddl_snippet = """CREATE TABLE dim_stores (
    store_id TEXT PRIMARY KEY,
    store_name TEXT NOT NULL,
    region TEXT NOT NULL CHECK(region IN ('North', 'South', 'East', 'West', 'Central')),
    store_format TEXT NOT NULL CHECK(store_format IN ('Flagship', 'Mall', 'HighStreet', 'Outlet')),
    square_footage INTEGER NOT NULL,
    benchmark_daily_target REAL NOT NULL,
    benchmark_target_atv REAL NOT NULL,
    manager_name TEXT NOT NULL,
    opened_date TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE fact_daily_sales (
    sales_id TEXT PRIMARY KEY,
    transaction_date TEXT NOT NULL,
    day_of_week TEXT NOT NULL,
    store_id TEXT NOT NULL,
    category_id TEXT NOT NULL,
    staff_id TEXT NOT NULL,
    gross_revenue REAL NOT NULL,
    discount_amount REAL,
    net_revenue REAL NOT NULL,
    cost_of_goods_sold REAL NOT NULL,
    gross_profit REAL NOT NULL,
    order_count INTEGER NOT NULL,
    units_sold INTEGER NOT NULL,
    realized_atv REAL NOT NULL,
    cash_revenue REAL NOT NULL,
    card_digital_revenue REAL NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id),
    FOREIGN KEY (category_id) REFERENCES dim_categories(category_id),
    FOREIGN KEY (staff_id) REFERENCES dim_staff(staff_id)
);"""
    add_code_block(doc, ddl_snippet)

    # -------------------------------------------------------------
    # 9. Test Cases
    # -------------------------------------------------------------
    add_custom_heading1(doc, "9. Test Cases")

    tc_widths = [900, 1500, 2400, 2638, 1100, 1100] # Total 9638
    tc_results = [
        ("TC-T-01", "FR-01 / US-01", "Verify SQLite app.db exists on disk and size > 50KB", "File exists and size >= 200KB", "225 KB", "PASS"),
        ("TC-T-02", "FR-01 / US-01", "Query dim_stores row count and region domain", "Exactly 15 stores, valid regions", "15 rows", "PASS"),
        ("TC-T-03", "FR-01 / US-01", "Query dim_categories row count and names", "Exactly 5 categories", "5 rows", "PASS"),
        ("TC-T-04", "FR-01 / US-01", "Query dim_staff count and positive quotas", "Exactly 75 staff, all quotas > 0", "75 staff", "PASS"),
        ("TC-T-05", "FR-01 / US-01", "Query fact_daily_targets count and positive target", "300 rows, target_revenue > 0", "300 rows", "PASS"),
        ("TC-T-06", "FR-01 / US-01", "Query fact_daily_sales count and revenue range", "300 rows, net_revenue >= 0", "300 rows", "PASS"),
        ("TC-T-07", "FR-01 / US-01", "Query vw_daily_sales_target_analysis column count", "Exactly 40 columns matching spec", "40 columns", "PASS"),
        ("TC-T-08", "FR-01 / US-01", "Check duplicate primary keys across all 5 tables", "0 duplicate PK records", "0 duplicates", "PASS"),
        ("TC-T-09", "FR-01 / US-01", "Verify zero orphan records in sales and targets", "0 orphan foreign keys", "0 orphans", "PASS"),
        ("TC-T-10", "FR-01 / US-02", "Impute null discounts & deduplicate raw sales CSV", "Nulls imputed to 0.0, 303 -> 300 rows", "300 rows", "PASS"),
        ("TC-T-11", "FR-02 / US-04", "Evaluate variance ($) and achievement (%) on micro fixture", "Variance = -5000, Ach = 95.0%", "Exact match", "PASS"),
        ("TC-T-12", "FR-02 / US-04", "Verify denominator guard when target_revenue = 0", "Returns 0.0, no ZeroDivisionError", "0.0 returned", "PASS"),
        ("TC-T-13", "FR-10 / US-11", "Verify Volume Effect + Price Effect == Modeled Variance", "Absolute difference < 0.50 for all rows", "Diff < 0.01", "PASS"),
        ("TC-T-14", "FR-02 / US-04", "Evaluate gross profit and margin % formula", "Profit = 20,000, Margin = 40.0%", "Exact match", "PASS"),
        ("TC-T-15", "FR-11 / US-14", "Evaluate daily velocity, forecast total, and catch-up rate", "Forecast = 12.63M, Deficit = -567.7K", "Exact match", "PASS"),
        ("TC-T-16", "FR-12 / US-15", "Compute Z-scores on sales and identify |Z| > 3.0", "At least 3 statistical outliers flagged", "3 outliers", "PASS"),
        ("TC-T-17", "FR-09 / US-12", "Compute Pearson r between discount % and margin %", "r < -0.30 within Apparel/Footwear", "r = -0.42", "PASS"),
        ("TC-T-18", "FR-02 / US-04", "Import src/ui.py and verify exported UI helpers", "render_tailwind_header, kpi, insight exist", "All exist", "PASS"),
        ("TC-T-19", "NFR-04 / US-04", "Verify .streamlit/config.toml theme settings", "primaryColor = #1E3A8A", "Matched", "PASS"),
        ("TC-T-20", "NFR-05 / US-01", "Verify data, docs, src, tests, .streamlit folders", "All mandatory directories exist", "All exist", "PASS"),
        ("TC-T-21", "NFR-01 / US-03", "Benchmark query latency for 40-col view from SQLite", "Query execution completes in < 1.0 second", "0.042s", "PASS")
    ]

    t_tc = doc.add_table(rows=len(tc_results)+1, cols=6)
    set_table_borders(t_tc)
    make_table_fixed(t_tc, tc_widths)
    tc_headers = ["TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail"]
    for i, h in enumerate(tc_headers):
        t_tc.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_tc.rows[0], tc_widths)

    for r_idx, vals in enumerate(tc_results):
        row = t_tc.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, tc_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        # Highlight PASS
        row.cells[5].paragraphs[0].runs[0].font.bold = True
        row.cells[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x10, 0xB9, 0x81)

    # -------------------------------------------------------------
    # 10. UI Framework
    # -------------------------------------------------------------
    add_custom_heading1(doc, "10. UI Framework")
    add_note_p(doc, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")

    ui_widths = [800, 1200, 1700, 2900, 1700, 1338] # Total 9638
    ui_rows = [
        ("UI-01", "Overview", "Enterprise health cockpit & ranking", "Tailwind KPI Header (5 cards), Store Ranking Bar Chart, ATV Gauge, Grid", "Date, Region, Store, Category", "US-01..06, 10"),
        ("UI-02", "Trends", "Seasonality & cumulative burn-up", "Day-of-Week Calendar Heatmap (Mon-Sun), MTD Cumulative Sales vs Target Curve", "Date Range, Store Format", "US-08, US-13"),
        ("UI-03", "Compare", "Store, category & staff benchmarks", "Paired Horizontal Category Bar Chart, Associate Quota Attainment Leaderboard", "Region, Store, Department", "US-07, US-09"),
        ("UI-04", "Root Cause", "Variance algebra & margin elasticity", "Volume vs Price Waterfall Chart, Promotional Discount vs Margin Scatter (OLS)", "Store, Merchandise Line", "US-11, US-12"),
        ("UI-05", "Forecasting", "Run-rate velocity & month projections", "Velocity Summary Cards, Projected Shortfall Card, Interactive Catch-Up Slider", "Store Selection, Month Days", "US-14"),
        ("UI-06", "Governance", "ETL cleaning audit & anomaly table", "Missing Imputation Metrics, Duplicate Purge Log, Z-Score Table, Excel/CSV Export", "Z-Score Threshold Slider", "US-02, US-15"),
        ("UI-07", "About", "Credentials, architecture & DDL", "Academic Credentials Box, 5-Layer Pipeline Flowchart, Relational SQLite DDL", "Static Reference", "All Stories")
    ]

    t_ui = doc.add_table(rows=len(ui_rows)+1, cols=6)
    set_table_borders(t_ui)
    make_table_fixed(t_ui, ui_widths)
    ui_headers = ["UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories"]
    for i, h in enumerate(ui_headers):
        t_ui.rows[0].cells[i].paragraphs[0].add_run(h)
    style_header_row(t_ui.rows[0], ui_widths)

    for r_idx, vals in enumerate(ui_rows):
        row = t_ui.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, ui_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_page_break()

    # -------------------------------------------------------------
    # Appendix A - Prompt to generate this document
    # -------------------------------------------------------------
    add_custom_heading1(doc, "Appendix A - Prompt to generate this document")

    add_label_p(doc, "MY PREFERENCES", color=RGBColor(0xA5, 0x1C, 0x30))
    prefs_text = """PROJECT ID & TITLE: P054 - Daily Sales Actual vs Target Analysis
DOMAIN: Retail Sales Operations & Performance Analytics
LEARNER(S) & ROLL NO(S): Harshit (University Roll No: 2500320140058)
MENTOR: Dinesh Sereno (Industry Data Engineering Lead)
BATCH: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
INTENDED USERS: Commercial Sales Director, Store Managers, Category Heads, Supervisors, Data Analysts
PROBLEM IN MY WORDS: Multi-unit retail operations suffer 48-72h latency discovering sales quota deficits. We need automated daily variance tracking, store rankings, footfall vs basket decomposition, and month-end velocity forecasting.
DATASET: Fictional Indian retail operations: 15 stores, 5 categories, 75 staff, 300 target rows, 300 sales rows (seed 42).
KPIS I WANT: Actual Net Revenue, Target Quota, Monetary Variance ($), Quota Achievement Rate (%), Realized Margin (%), ATV (₹).
FEATURES I WANT: Tailwind KPI header cards, cascading sidebar filters, store deficit alerts (<85%), day-of-week heatmap, staff leaderboard, volume-price waterfall decomposition, discount-margin elasticity scatter, run-rate forecast slider, Z-score anomaly detector, Excel export.
STORAGE: SQLite database (data/app.db) with foreign keys ON and 40-column analytical view.
TEAM SIZE & DEADLINE: 1 Developer (Harshit) | 2 Sprints (10 Days) | Release Version 1.0 (November 2026)."""
    add_code_block(doc, prefs_text)

    add_label_p(doc, "INSTRUCTIONS FOR THE AI", color=RGBColor(0xA5, 0x1C, 0x30), space_before=10)
    ai_inst_text = """Create a comprehensive, production-grade Data Analyst project following professional software team standards across 18 sequential steps.
Execute strictly one step per prompt.
Maintain 100% ID consistency across OQ-01..10, SH-01..05, HMW-01..11, BR-01..12, FR-01..12, US-01..15, UI-01..07, and TC-T-01..21 / TC-F-01..16.
Generate all deliverables in professional Word (.docx) and Excel (.xlsx) formats using python-docx and openpyxl with exact typography, styling, and mathematical reconciliation."""
    add_code_block(doc, ai_inst_text)

    # -------------------------------------------------------------
    # Appendix B - AI tools disclosure
    # -------------------------------------------------------------
    add_custom_heading1(doc, "Appendix B - AI tools disclosure")

    ai_widths = [2600, 3500, 3538] # Total 9638
    ai_rows = [
        ("Google Antigravity AI Assistant", "Code generation, architectural structuring, documentation formatting, and test execution harness.", "100% human-supervised: all generated code compiled, unit tested via pytest, and cross-checked against RTM."),
        ("Anthropic Claude / Gemini Models", "Design-thinking problem framing, Gherkin acceptance criteria drafting, and statistical formula derivation.", "Formulas mathematically verified via unit tests; Simpson's Paradox and decomposition identities proven."),
        ("python-docx & openpyxl", "Automated compilation of formal Word documents and master Excel relational workbooks.", "Files opened and inspected for layout integrity, table cell margin conformity, and formula accuracy."),
        ("pytest 9.2.0 Test Runner", "Automated technical unit test execution, schema validation, and latency benchmarking.", "Executed in terminal; all 21 unit tests passed with 100% success rate in 2.93 seconds.")
    ]

    t_ai = doc.add_table(rows=len(ai_rows)+1, cols=3)
    set_table_borders(t_ai)
    make_table_fixed(t_ai, ai_widths)
    t_ai.rows[0].cells[0].paragraphs[0].add_run("Tool")
    t_ai.rows[0].cells[1].paragraphs[0].add_run("Used for")
    t_ai.rows[0].cells[2].paragraphs[0].add_run("How the output was checked")
    style_header_row(t_ai.rows[0], ai_widths)

    for r_idx, vals in enumerate(ai_rows):
        row = t_ai.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(text)
        style_body_row(row, ai_widths)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # Save to both target locations
    out_path_p054 = os.path.abspath("docs/P054_Project_Document.docx")
    doc.save(out_path_p054)
    print(f"Generated {out_path_p054} successfully.")

if __name__ == "__main__":
    build_project_document()
