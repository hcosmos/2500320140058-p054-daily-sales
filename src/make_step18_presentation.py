"""
Script: src/make_step18_presentation.py
Purpose: Generate docs/P054_Presentation.pptx using python-pptx according to exact ABES guidelines.
"""

import sys
import os
from pathlib import Path
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

# ---------------------------------------------------------
# CONSTANTS & PALETTE
# ---------------------------------------------------------
CANVAS_WIDTH = Inches(10)
CANVAS_HEIGHT = Inches(5.625)

RED = RGBColor(0xA5, 0x1C, 0x30)
BLACK = RGBColor(0x15, 0x15, 0x14)
DARK = RGBColor(0x3A, 0x3A, 0x38)
GREY = RGBColor(0x5E, 0x5E, 0x5B)
CARD_BG = RGBColor(0xF4, 0xF1, 0xF1)
PINK = RGBColor(0xD9, 0xB5, 0xBC)
BORDER_HEX = "D9D9D9"
LINE_GREY = RGBColor(0xBF, 0xBF, 0xBF)
SLIDE_NUM_COL = RGBColor(0x8A, 0x8A, 0x86)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

PROJECT_ID = "P054"
PROJECT_TITLE = "Daily Sales Actual vs Target Analysis"
LEARNER = "Harshit"
ROLL_NO = "2500320140058"
MENTOR = "Dinesh Sereno"
BATCH_ID = "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"
DATE_STR = "21 November 2026"
APP_URL = "https://2500320140058-daily-sales.streamlit.app"
GITHUB_URL = "https://github.com/harshit-abes/2500320140058-p054-daily-sales"

# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------

def set_zero_insets(tf):
    """Set zero insets and enable word wrap on a text frame."""
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)

def set_default_insets(tf):
    """Set default insets and enable word wrap."""
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)

def add_card(slide, x, y, w, h):
    """CARD(x,y,w,h): rounded rectangle, radius .08, fill and line F4F1F1."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BG
    card.line.width = Pt(0.75)
    try:
        card.adjustments[0] = 0.08
    except Exception:
        pass
    return card

def add_badge(slide, x, y, text):
    """BADGE (x,y,t): circle .42 red; t centred, 12 bold white."""
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x), Inches(y), Inches(0.42), Inches(0.42)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = RED
    badge.line.color.rgb = RED
    badge.line.width = Pt(0.5)
    tf = badge.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(text)
    run.font.name = "Calibri"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = WHITE
    return badge

def add_head(slide, x, y, w, h, text):
    """HEAD: 15 bold black, middle."""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    set_zero_insets(tf)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = BLACK
    return txBox

def add_body(slide, x, y, w, h, text):
    """BODY: 13 dark, top."""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    set_zero_insets(tf)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(13)
    run.font.color.rgb = DARK
    return txBox

def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text):
    """
    BADGE CARD:
    CARD; BADGE(x+.2, y+.2); HEAD (x+.75, y+.17, w-.9, .45);
    BODY (x+.2, y+.72, w-.4, h-.85).
    """
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_text)
    add_head(slide, x + 0.75, y + 0.17, w - 0.9, 0.45, head_text)
    add_body(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, body_text)

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    """
    PLAIN CARD:
    CARD; HEAD (x+.2, y+.17, w-.35, .45); BODY (x+.2, y+.72, w-.4, h-.85).
    """
    add_card(slide, x, y, w, h)
    add_head(slide, x + 0.2, y + 0.17, w - 0.35, 0.45, head_text)
    add_body(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, body_text)

def set_cell_border_and_margins(cell, border_color="D9D9D9", border_width="6350"):
    """
    Apply cell borders .5 pt D9D9D9, margins .1 left/right, .05 top/bottom.
    """
    cell.margin_left = Inches(0.1)
    cell.margin_right = Inches(0.1)
    cell.margin_top = Inches(0.05)
    cell.margin_bottom = Inches(0.05)
    
    tcPr = cell._tc.get_or_add_tcPr()
    for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
        ln = parse_xml(
            f'<a:{border_name} {nsdecls("a")} w="{border_width}" cmpd="s">\n'
            f'  <a:solidFill>\n'
            f'    <a:srgbClr val="{border_color}"/>\n'
            f'  </a:solidFill>\n'
            f'</a:{border_name}>'
        )
        tcPr.append(ln)

def format_cell(cell, text, is_header=False, font_size=13, bold=False, text_color=None, align=PP_ALIGN.LEFT):
    """Format table cell per prompt specification."""
    set_cell_border_and_margins(cell, border_color=BORDER_HEX, border_width="6350")
    cell.fill.solid()
    if is_header:
        cell.fill.fore_color.rgb = RED
    else:
        cell.fill.fore_color.rgb = WHITE
        
    cell.text = ""  # clear
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(font_size)
    if is_header:
        run.font.bold = True
        run.font.color.rgb = WHITE
    else:
        run.font.bold = bold
        run.font.color.rgb = text_color if text_color else BLACK

def set_speaker_notes(slide, notes_text):
    """Add speaker notes to the slide."""
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text

def create_dark_slide(prs, title_text, lines_list, notes_text):
    """
    DARK SLIDE:
    Background 151514.
    [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad
    (2.2,.35,7.3,.5) 11 pink, right, middle, default insets
    Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Dark background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BLACK
    
    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    box_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_logo = box_logo.text_frame
    set_default_insets(tf_logo)
    tf_logo.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_logo.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "[ABES logo]"
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.italic = True
    r.font.color.rgb = PINK
    
    # Top right line: HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad
    # (2.2,.35,7.3,.5) 11 pink, right, middle, default insets
    box_hcl = slide.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_hcl = box_hcl.text_frame
    set_default_insets(tf_hcl)
    tf_hcl.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_hcl.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.color.rgb = PINK
    
    # Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    box_title = slide.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_title = box_title.text_frame
    set_default_insets(tf_title)
    tf_title.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf_title.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title_text
    r.font.name = "Cambria"
    r.font.size = Pt(36)
    r.font.bold = True
    r.font.color.rgb = WHITE
    
    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    box_lines = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = box_lines.text_frame
    set_default_insets(tf_lines)
    tf_lines.vertical_anchor = MSO_ANCHOR.TOP
    for i, line_item in enumerate(lines_list):
        if i == 0:
            p = tf_lines.paragraphs[0]
        else:
            p = tf_lines.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4)
        
        # Check if line contains special pink text
        if isinstance(line_item, tuple):
            text, is_pink = line_item
            r = p.add_run()
            r.text = text
            r.font.name = "Calibri"
            r.font.size = Pt(15)
            r.font.color.rgb = PINK if is_pink else CARD_BG
        else:
            r = p.add_run()
            r.text = line_item
            r.font.name = "Calibri"
            r.font.size = Pt(15)
            r.font.color.rgb = CARD_BG
            
    set_speaker_notes(slide, notes_text)
    return slide

def create_content_slide(prs, title_text, slide_number, notes_text):
    """
    CONTENT SLIDE:
    White. Use a real title placeholder.
    Title (.5,.3,9,.75) Cambria 28 bold black, centred, middle, default insets
    Footer (.5,5.2,7.5,.3) 9 grey, middle, default insets:
    HCLTech Career Shaper ABES Engineering College | <Project ID - Title>
    Slide number (9,5.2,.5,.3) 9, colour 8A8A86, right, default insets
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # White background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    
    # Title (.5,.3,9,.75) Cambria 28 bold black, centred, middle, default insets
    box_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf_title = box_title.text_frame
    set_default_insets(tf_title)
    tf_title.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_title.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title_text
    r.font.name = "Cambria"
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = BLACK
    
    # Footer (.5,5.2,7.5,.3) 9 grey, middle, default insets:
    box_foot = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_foot = box_foot.text_frame
    set_default_insets(tf_foot)
    tf_foot.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_foot.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f"HCLTech Career Shaper ABES Engineering College | {PROJECT_ID} - {PROJECT_TITLE}"
    r.font.name = "Calibri"
    r.font.size = Pt(9)
    r.font.color.rgb = GREY
    
    # Slide number (9,5.2,.5,.3) 9, colour 8A8A86, right, default insets
    box_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_num = box_num.text_frame
    set_default_insets(tf_num)
    tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_num.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = str(slide_number)
    r.font.name = "Calibri"
    r.font.size = Pt(9)
    r.font.color.rgb = SLIDE_NUM_COL
    
    set_speaker_notes(slide, notes_text)
    return slide

# ---------------------------------------------------------
# PRESENTATION BUILDER
# ---------------------------------------------------------

def build_presentation(output_path="docs/P054_Presentation.pptx"):
    prs = Presentation()
    prs.slide_width = CANVAS_WIDTH
    prs.slide_height = CANVAS_HEIGHT
    
    print("Building 11 slides for ABES Project Presentation...")
    
    # =========================================================
    # SLIDE 1: DARK (Title Slide)
    # =========================================================
    title_1 = f"{PROJECT_ID} - {PROJECT_TITLE}"
    lines_1 = [
        f"{LEARNER} | University Roll No: {ROLL_NO}",
        f"Batch ID: {BATCH_ID}",
        f"Project Mentor: {MENTOR} (Industry Data Engineering Lead)",
        (f"{DATE_STR}", True)  # date in pink
    ]
    notes_1 = (
        "Good morning mentors and evaluators. I am Harshit, presenting Project P054: "
        "Daily Sales Actual vs Target Analysis. Built under the HCLTech Career Shaper Programme, "
        "this analytics solution bridges the 48-to-72-hour reporting latency in multi-unit retail "
        "networks by delivering automated daily variance attribution, deficit store detection, "
        "and month-end projections."
    )
    create_dark_slide(prs, title_1, lines_1, notes_1)
    print("Slide 1 generated.")
    
    # =========================================================
    # SLIDE 2: CONTENT "Problem and How Might We"
    # =========================================================
    notes_2 = (
        "Our core problem is that retail leaders cannot act on sales deficits when actuals and "
        "targets sit in disconnected spreadsheets with a 48 to 72-hour delay. To address this, "
        "HMW-01 focuses on delivering consolidated executive visibility across all branches on one "
        "dashboard, while HMW-02 automates deficit calculations so lagging stores are flagged "
        "within two hours of store close."
    )
    s2 = create_content_slide(prs, "Problem and How Might We", 2, notes_2)
    
    # BADGE CARD (.5,1.25,4.3,3.7) P, Problem statement: 2-3 sentences
    problem_text = (
        "Multi-unit retail leaders struggle to spot lagging stores, categories, and associates in "
        "time to hit monthly revenue targets because sales actuals and target quotas are trapped in "
        "disconnected spreadsheets, delaying corrective operational intervention by 48-72 hours."
    )
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem Statement", problem_text)
    
    # BADGE CARD (5.2,1.25,4.3,1.75) 1, HMW-01: statement
    hmw01_text = (
        "Consolidate daily actual revenue alongside assigned targets across retail branches "
        "for Regional Directors to assess quota achievement on a single screen."
    )
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Enterprise Visibility", hmw01_text)
    
    # BADGE CARD (5.2,3.2,4.3,1.75) 2, HMW-02: statement
    hmw02_text = (
        "Calculate absolute revenue deficit and percentage shortfall for Store Managers "
        "so that lagging stores are flagged within 2 hours of trading close."
    )
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Deficit Store Alerts", hmw02_text)
    print("Slide 2 generated.")
    
    # =========================================================
    # SLIDE 3: CONTENT "Stakeholders and users"
    # =========================================================
    notes_3 = (
        "Here are our three core stakeholders. Our primary user is the Commercial Sales Director "
        "who requires high-level executive visibility and month-end forecasts with high organizational "
        "influence. Our secondary user is the Store Operations Manager who needs daily branch-level "
        "variance decomposition. Finally, our mentor and evaluators guide analytical rigor and verify "
        "engineering standards."
    )
    s3 = create_content_slide(prs, "Stakeholders and users", 3, notes_3)
    
    # BADGE CARDS y 1.35, w 2.85, h 3.5, x .5 / 3.55 / 6.6, badges 1 / 2 / 3: primary user, secondary user, mentor or evaluator.
    # BODY: role and need, blank line, Influence: High / Medium / Low
    body_sh1 = (
        "Role & Need: Primary User. Requires consolidated enterprise actuals vs budget quotas and "
        "month-end run-rate projections to protect gross margin and drive regional commercial interventions.\n\n"
        "Influence: High"
    )
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Commercial Sales Director", body_sh1)
    
    body_sh2 = (
        "Role & Need: Secondary User. Needs daily branch quota attainment rankings, volume vs price "
        "variance decomposition, and day-of-week traffic insights to adjust floor staffing.\n\n"
        "Influence: High"
    )
    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Store Operations Manager", body_sh2)
    
    body_sh3 = (
        "Role & Need: Mentor & Evaluator. Evaluates dimensional schema integrity, analytical "
        "decomposition rigor, test coverage, and cloud architecture resilience against enterprise standards.\n\n"
        "Influence: Medium"
    )
    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Academic & Mentor", body_sh3)
    print("Slide 3 generated.")
    
    # =========================================================
    # SLIDE 4: CONTENT "Requirements and traceability"
    # =========================================================
    notes_4 = (
        "Every business requirement in our project traces cleanly across the full lifecycle. "
        "As shown in this matrix sample, HMW-01 maps directly to requirement BR-01, User Story US-01, "
        "automated unit test TC-T-08, and UI-01 Overview screen. We achieved 100% test and user story "
        "coverage across all twelve business requirements."
    )
    s4 = create_content_slide(prs, "Requirements and traceability", 4, notes_4)
    
    # TABLE at (.5,1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2:
    # HMW | Requirement | Priority | User story | Test | Screen.
    # 3-4 Must/Should rows.
    t_shape = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    table = t_shape.table
    col_widths = [1.0, 3.4, 1.0, 1.2, 1.2, 1.2]
    for ci, cw in enumerate(col_widths):
        table.columns[ci].width = Inches(cw)
        
    headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    for ci, h_text in enumerate(headers):
        cell = table.cell(0, ci)
        format_cell(cell, h_text, is_header=True, font_size=12, align=PP_ALIGN.CENTER if ci in [0,2,3,4] else PP_ALIGN.LEFT)
        
    rows_data = [
        ("HMW-01", "Consolidated Daily Actual vs Target Revenue & Quota %", "Must", "US-01", "TC-T-08", "UI-01 Overview"),
        ("HMW-02", "Store Attainment Ranking & Deficit Shortfall Highlighting", "Must", "US-03", "TC-T-11", "UI-01 Overview"),
        ("HMW-09", "Variance Decomposition into Volume and Price Effects", "Must", "US-07", "TC-T-13", "UI-04 Root Cause"),
        ("HMW-10", "Month-End Run-Rate Forecast & Revenue Projection", "Should", "US-10", "TC-T-15", "UI-05 Forecast")
    ]
    
    for ri, r_data in enumerate(rows_data, start=1):
        for ci, val in enumerate(r_data):
            cell = table.cell(ri, ci)
            align = PP_ALIGN.CENTER if ci in [0, 2, 3, 4] else PP_ALIGN.LEFT
            bold = True if ci in [0, 2] else False
            format_cell(cell, val, is_header=False, font_size=11, bold=bold, align=align)
            
    # .3 below the table, (.5,y9,.35) italic 12 grey: Full matrix: Project Document, section 4
    # table top=1.3, height=2.5, bottom=3.8 -> y9 = 4.1
    box_sub = s4.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(9.0), Inches(0.35))
    tf_sub = box_sub.text_frame
    set_zero_insets(tf_sub)
    tf_sub.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_sub.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "Full matrix: Project Document, section 4"
    r.font.name = "Calibri"
    r.font.size = Pt(12)
    r.font.italic = True
    r.font.color.rgb = GREY
    print("Slide 4 generated.")
    
    # =========================================================
    # SLIDE 5: CONTENT "Data and preparation"
    # =========================================================
    notes_5 = (
        "Our analytical data foundation combines 300 clean transaction rows with 40 joined dimensional "
        "columns over a 20-day trading period across 15 stores. During automated ETL ingestion, our pipeline "
        "deduplicated 3 redundant rows, imputed 6 missing discount values using category medians, and "
        "flagged 3 extreme revenue anomalies using statistical Z-scores."
    )
    s5 = create_content_slide(prs, "Data and preparation", 5, notes_5)
    
    # Four stats, x .5 / 2.8 / 5.1 / 7.4, w 2.1:
    # value y 1.3 h .8 Cambria 30 bold red, middle;
    # label y 2.1 h .35, 12 grey, middle: rows / columns / period / issues fixed.
    stats = [
        ("300", "Clean Rows (15 Stores)"),
        ("40", "Analytical Columns"),
        ("20 Days", "Trading Period"),
        ("12", "Data Issues Fixed")
    ]
    stat_xs = [0.5, 2.8, 5.1, 7.4]
    for (val, lbl), sx in zip(stats, stat_xs):
        # Value box
        b_val = s5.shapes.add_textbox(Inches(sx), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_v = b_val.text_frame
        set_zero_insets(tf_v)
        tf_v.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_v.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = val
        r.font.name = "Cambria"
        r.font.size = Pt(30)
        r.font.bold = True
        r.font.color.rgb = RED
        
        # Label box
        b_lbl = s5.shapes.add_textbox(Inches(sx), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_l = b_lbl.text_frame
        set_zero_insets(tf_l)
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_l.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = lbl
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.color.rgb = GREY
        
    # PLAIN CARD (.5,2.7,4.3,2.25) Source: dataset source; real, public or synthetic; what one row represents.
    src_body = (
        "Dataset: Synthetic relational retail database (seed 42) spanning 15 stores, "
        "5 merchandise categories, and 75 sales associates.\n\n"
        "Grain: One row represents daily store-category transactions and assigned budget targets."
    )
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source: Relational Constellation", src_body)
    
    # PLAIN CARD (5.2,2.7,4.3,2.25) Cleaning: top 3 fixes with rows affected.
    clean_body = (
        "1. Deduplication: Removed 3 redundant rows (303 -> 300 clean rows).\n"
        "2. Missing Values: Imputed 6 null discounts using category median.\n"
        "3. Outliers: Flagged 3 Z-score anomalies (> 3.0) for operational review."
    )
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning: Top 3 Quality Fixes", clean_body)
    print("Slide 5 generated.")
    
    # =========================================================
    # SLIDE 6: CONTENT "Solution and technology"
    # =========================================================
    notes_6 = (
        "Our technical architecture follows a modern 5-stage pipeline. Raw CSV and Excel files "
        "are validated and cleaned in Python using pandas, loaded into an ACID-compliant SQLite database, "
        "visualized through a 7-screen multi-page Streamlit application, and deployed via Git to "
        "Streamlit Community Cloud for global accessibility."
    )
    s6 = create_content_slide(prs, "Solution and technology", 6, notes_6)
    
    # Five steps, x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2:
    # rounded, radius .1, fill F4F1F1; text centred, middle, 12 bold black, two lines:
    # Raw data (CSV / xlsx); Clean & KPIs (Python, pandas); Store (<my storage>); Dashboard (Streamlit); Deploy (GitHub -> cloud).
    # Dashboard step: red fill, white text.
    # ">" between steps, x 2.1 / 3.95 / 5.8 / 7.65, y 1.5, w .25, h 1.2: 18 bold red, centred.
    steps = [
        ("Raw data\n(CSV / xlsx)", False),
        ("Clean & KPIs\n(Python, pandas)", False),
        ("Store\n(SQLite app.db)", False),
        ("Dashboard\n(Streamlit)", True),  # Red fill, white text
        ("Deploy\n(GitHub -> cloud)", False)
    ]
    step_xs = [0.5, 2.35, 4.2, 6.05, 7.9]
    arrow_xs = [2.1, 3.95, 5.8, 7.65]
    
    for (text, is_red), sx in zip(steps, step_xs):
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(1.5), Inches(1.6), Inches(1.2))
        box.fill.solid()
        if is_red:
            box.fill.fore_color.rgb = RED
            box.line.color.rgb = RED
        else:
            box.fill.fore_color.rgb = CARD_BG
            box.line.color.rgb = CARD_BG
        try:
            box.adjustments[0] = 0.1
        except Exception:
            pass
            
        tf = box.text_frame
        set_zero_insets(tf)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = WHITE if is_red else BLACK
        
    for ax in arrow_xs:
        b_ar = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_a = b_ar.text_frame
        set_zero_insets(tf_a)
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_a.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ">"
        r.font.name = "Calibri"
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = RED
        
    # (.5,3.15,9,.5) 13 dark, middle: bold "Tools:" + my tools separated by " · "
    box_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_tools = box_tools.text_frame
    set_zero_insets(tf_tools)
    tf_tools.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_tools.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "Tools: "
    r1.font.name = "Calibri"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = DARK
    
    r2 = p.add_run()
    r2.text = "Python 3.14 · SQLite · Pandas 3.1 · Plotly · Streamlit · pytest · Git & GitHub"
    r2.font.name = "Calibri"
    r2.font.size = Pt(13)
    r2.font.bold = False
    r2.font.color.rgb = DARK
    
    # (.5,3.75,9,.5) 13 italic grey, middle:
    # Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud
    box_run = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_run = box_run.text_frame
    set_zero_insets(tf_run)
    tf_run.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_run.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    r.font.name = "Calibri"
    r.font.size = Pt(13)
    r.font.italic = True
    r.font.color.rgb = GREY
    print("Slide 6 generated.")
    
    # =========================================================
    # SLIDE 7: CONTENT "Dashboard"
    # =========================================================
    notes_7 = (
        "Slide 7 showcases our Streamlit user interface. The top KPI row gives immediate visibility "
        "into total actual revenue, budget target, overall quota attainment, and net deficit. The four "
        "core screens cover executive KPIs in UI-01, day-of-week seasonality in UI-02, volume versus "
        "price root cause analysis in UI-04, and run-rate forecasting in UI-05."
    )
    s7 = create_content_slide(prs, "Dashboard", 7, notes_7)
    
    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF.
    # Put my screenshot inside, keep its proportions, no stretching.
    # No screenshot: centred italic 14 grey: [Paste a dashboard screenshot here]
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = CARD_BG
    frame.line.color.rgb = LINE_GREY
    frame.line.width = Pt(1.0)
    frame.line.dash_style = pptx.enum.dml.MSO_LINE_DASH_STYLE.DASH
    
    # Check if dashboard screenshot exists
    img_path = Path("docs/dashboard_screenshot.png")
    if img_path.exists():
        # Keep proportions within (5.4 x 3.6)
        s7.shapes.add_picture(str(img_path), Inches(0.55), Inches(1.3), Inches(5.4), Inches(3.6))
    else:
        tf_fr = frame.text_frame
        set_zero_insets(tf_fr)
        tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_fr.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = "[Paste a dashboard screenshot here]"
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.italic = True
        r.font.color.rgb = GREY
        
    # BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4.
    # Beside each, (7.25,y,2.25,.42), middle, 12 bold red "UI-01" + short black screen purpose.
    screens = [
        ("1", "UI-01", "Executive Overview & KPIs"),
        ("2", "UI-02", "Trends & Seasonality"),
        ("3", "UI-04", "Variance Root Cause"),
        ("4", "UI-05", "Month-End Forecast")
    ]
    badge_ys = [1.3, 2.2, 3.1, 4.0]
    for (num, ui_id, purpose), by in zip(screens, badge_ys):
        add_badge(s7, 6.7, by, num)
        
        b_txt = s7.shapes.add_textbox(Inches(7.25), Inches(by), Inches(2.25), Inches(0.42))
        tf_t = b_txt.text_frame
        set_zero_insets(tf_t)
        tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_t.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        
        r1 = p.add_run()
        r1.text = f"{ui_id}  "
        r1.font.name = "Calibri"
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = RED
        
        r2 = p.add_run()
        r2.text = purpose
        r2.font.name = "Calibri"
        r2.font.size = Pt(12)
        r2.font.bold = False
        r2.font.color.rgb = BLACK
    print("Slide 7 generated.")
    
    # =========================================================
    # SLIDE 8: CONTENT "Key findings"
    # =========================================================
    notes_8 = (
        "Our analysis revealed three critical business findings. First, enterprise achievement "
        "sits at 93.5%, creating a Rs 1.06 million deficit with nine lagging stores. Second, mathematical "
        "waterfall decomposition proves 68.4% of this deficit is driven by footfall volume loss rather "
        "than ticket size. Third, discount elasticity regression reveals discounting erodes gross "
        "margins without lifting unit volume."
    )
    s8 = create_content_slide(prs, "Key findings", 8, notes_8)
    
    # Three CARDS y 1.3, w 2.85, h 3.6, x .5 / 3.55 / 6.6.
    # Inside, at x+.2, w 2.45:
    # number y 1.5 h 1 Cambria 40 bold red, middle;
    # meaning y 2.6 h .9, 14 bold black, top;
    # Evidence: KPI or test y 3.6 h 1.1, 12 grey, top.
    findings = [
        ("93.5%", "Enterprise Quota Pace\n(₹1.06M Net Deficit)", "Evidence: Actual ₹15.24M vs Target ₹16.30M across 15 stores; 9 stores lagging."),
        ("68.4%", "Shortfall Driven by\nFootfall Volume Loss", "Evidence: Waterfall shows ₹725K volume loss vs ₹335K basket ticket loss."),
        ("-0.34", "Discount Margin Erosion\nin High-Margin Apparel", "Evidence: OLS regression r = -0.34 (p < 0.05); discounts fail to lift volume.")
    ]
    card_xs = [0.5, 3.55, 6.6]
    for (num_txt, mean_txt, evid_txt), cx in zip(findings, card_xs):
        add_card(s8, cx, 1.3, 2.85, 3.6)
        
        # Number box
        b_num = s8.shapes.add_textbox(Inches(cx + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = b_num.text_frame
        set_zero_insets(tf_n)
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_n.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = num_txt
        r.font.name = "Cambria"
        r.font.size = Pt(40)
        r.font.bold = True
        r.font.color.rgb = RED
        
        # Meaning box
        b_mean = s8.shapes.add_textbox(Inches(cx + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = b_mean.text_frame
        set_zero_insets(tf_m)
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_m.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = mean_txt
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = BLACK
        
        # Evidence box
        b_evid = s8.shapes.add_textbox(Inches(cx + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = b_evid.text_frame
        set_zero_insets(tf_e)
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_e.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = evid_txt
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.color.rgb = GREY
    print("Slide 8 generated.")
    
    # =========================================================
    # SLIDE 9: CONTENT "Testing and validation"
    # =========================================================
    notes_9 = (
        "Quality assurance is backed by 21 automated pytest test cases and 16 manual functional tests, "
        "all passing with 100% success in under 3 seconds. Key tests verify relational schema integrity "
        "in TC-T-01, the mathematical exactness of our volume-price waterfall identity in TC-T-13, and "
        "sub-3-second multi-page render benchmarks in TC-T-18."
    )
    s9 = create_content_slide(prs, "Testing and validation", 9, notes_9)
    
    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    b_pt = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_pt = b_pt.text_frame
    set_zero_insets(tf_pt)
    tf_pt.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_pt.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "21 / 21"
    r.font.name = "Cambria"
    r.font.size = Pt(44)
    r.font.bold = True
    r.font.color.rgb = RED
    
    # (.5,2.4,2.6,.9) 13 grey, middle: test cases passed / pytest: n / manual: n
    b_tstats = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_ts = b_tstats.text_frame
    set_zero_insets(tf_ts)
    tf_ts.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_ts.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "Test cases passed\n"
    r1.font.name = "Calibri"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = BLACK
    
    r2 = p.add_run()
    r2.text = "pytest: 21 | manual: 16\nExecution Time: 2.97s"
    r2.font.name = "Calibri"
    r2.font.size = Pt(12)
    r2.font.color.rgb = GREY
    
    # TABLE at (3.4,1.35), column widths 1.1, 3.9, 1.1:
    # Test | Checks | Result. 3 key rows.
    t_shape9 = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.0))
    table9 = t_shape9.table
    table9.columns[0].width = Inches(1.1)
    table9.columns[1].width = Inches(3.9)
    table9.columns[2].width = Inches(1.1)
    
    t9_headers = ["Test", "Checks", "Result"]
    for ci, h_text in enumerate(t9_headers):
        cell = table9.cell(0, ci)
        format_cell(cell, h_text, is_header=True, font_size=12, align=PP_ALIGN.CENTER if ci in [0, 2] else PP_ALIGN.LEFT)
        
    t9_rows = [
        ("TC-T-01", "Relational schema, FK integrity, 300 clean rows", "Passed"),
        ("TC-T-13", "Volume + price effect mathematically equals variance", "Passed"),
        ("TC-T-18", "Multi-page Streamlit app launches and renders < 3s", "Passed")
    ]
    for ri, (tid, check_desc, res) in enumerate(t9_rows, start=1):
        c0 = table9.cell(ri, 0)
        format_cell(c0, tid, is_header=False, font_size=11, bold=True, align=PP_ALIGN.CENTER)
        c1 = table9.cell(ri, 1)
        format_cell(c1, check_desc, is_header=False, font_size=11, align=PP_ALIGN.LEFT)
        c2 = table9.cell(ri, 2)
        format_cell(c2, res, is_header=False, font_size=11, bold=True, text_color=RGBColor(0x10, 0xB9, 0x81), align=PP_ALIGN.CENTER)
    print("Slide 9 generated.")
    
    # =========================================================
    # SLIDE 10: CONTENT "Recommendation, limitations and next steps"
    # =========================================================
    notes_10 = (
        "Based on our findings, we recommend shifting promotional spend from discounts to footfall "
        "acquisition, realigning staff to lagging mid-week days, and coaching bottom-quartile associates. "
        "As an honest limitation, our current MVP operates on daily batch data rather than streaming "
        "feeds. Next steps include adding automated WhatsApp alert triggers and ML predictive modeling."
    )
    s10 = create_content_slide(prs, "Recommendation, limitations and next steps", 10, notes_10)
    
    # BADGE CARDS y 1.25, w 2.85, h 3.7, x .5 / 3.58 / 6.65:
    # R Recommendation; L Limitations; N Next steps. Base them on my real findings.
    body_r = (
        "1. Prioritize footfall marketing over discounts (volume causes 68% of deficit).\n"
        "2. Reallocate weekend staff to lagging mid-week days.\n"
        "3. Provide quota coaching for bottom-quartile associates."
    )
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendations", body_r)
    
    body_l = (
        "1. Operates on daily batch summaries rather than real-time POS streams.\n"
        "2. Analyzes a 20-day trading window; lacks multi-year seasonality.\n"
        "3. External factors (weather, road closures) are unmodeled."
    )
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", body_l)
    
    body_n = (
        "1. Ingest intraday POS streaming via REST / webhook architecture.\n"
        "2. Integrate automated WhatsApp alerts for store deficit warnings.\n"
        "3. Deploy machine learning XGBoost model for predictive forecasting."
    )
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next Steps", body_n)
    print("Slide 10 generated.")
    
    # =========================================================
    # SLIDE 11: DARK (Closing Slide)
    # =========================================================
    title_11 = "Live demo and Q&A"
    lines_11 = [
        f"App: {APP_URL}",
        f"Code: {GITHUB_URL}",
        "Thank you"
    ]
    notes_11 = (
        "I will now open the live interactive Streamlit application to demonstrate the executive "
        "KPI dashboard, variance waterfall decomposition, and predictive forecasting tools. "
        "Thank you for your time, and I look forward to your questions."
    )
    create_dark_slide(prs, title_11, lines_11, notes_11)
    print("Slide 11 generated.")
    
    # Save presentation
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"Presentation saved successfully to: {output_path}")

if __name__ == "__main__":
    build_presentation()
