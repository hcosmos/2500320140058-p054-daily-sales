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

def create_deployment_guide_doc():
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
    r_title = title_p.add_run("STREAMLIT COMMUNITY CLOUD DEPLOYMENT GUIDE")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(14)
    r_sub = sub_p.add_run("Project P054: Daily Sales Actual vs Target Analysis | Cloud Release & Production Operations")
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
        ("Deployment Target:", "Streamlit Community Cloud (Linux container, Python 3.12+, GitHub integration)"),
        ("Release Status & Date:", "Version 1.0 (Production Live) | November 2026")
    ]
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        style_data_cell(row.cells[0], k, is_even=(row_idx%2==1), width=Inches(2.2), bold=True)
        style_data_cell(row.cells[1], v, is_even=(row_idx%2==1), width=Inches(4.3))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Pre-Flight Cloud Readiness Audit
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Pre-Flight Cloud Readiness Audit Checklist")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(30, 58, 138)

    audit_items = [
        ("Root app.py Presence", "app.py exists in project root directory as the designated execution entry point.", "VERIFIED (Pass)"),
        ("Pinned requirements.txt", "Pinned production dependencies (streamlit, pandas, plotly, scipy, statsmodels).", "VERIFIED (Pass)"),
        ("Database File Committed", "SQLite database data/app.db (225 KB) is committed to Git repository.", "VERIFIED (Pass)"),
        ("Resilient Path Resolution", "Paths use Path(__file__).resolve().parent.parent / 'data' / 'app.db' for Linux containers.", "VERIFIED (Pass)"),
        ("Zero Hardcoded Secrets", "No API keys, database credentials, or secret tokens hardcoded in repository.", "VERIFIED (Pass)"),
        ("Git Ignore Rules", ".gitignore excludes .venv/, pycache, and OS metadata files.", "VERIFIED (Pass)"),
        ("Primary Git Branch", "Repository uses industry standard 'main' branch for automated Cloud webhook builds.", "VERIFIED (Pass)")
    ]

    audit_table = doc.add_table(rows=len(audit_items)+1, cols=3)
    audit_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(audit_table)
    a_headers = ["Audit Requirement", "Technical Verification Criteria", "Status"]
    for i, h in enumerate(a_headers):
        style_header_cell(audit_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(audit_items):
        row = audit_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 2: Step-by-Step GitHub & Streamlit Cloud Deployment Procedures
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Step-by-Step GitHub & Streamlit Cloud Deployment Procedures")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(30, 58, 138)

    p_steps = doc.add_paragraph()
    p_steps.paragraph_format.space_after = Pt(6)
    p_steps.add_run(
        "Follow these exact steps to push the project to GitHub and deploy to Streamlit Community Cloud:\n\n"
        "Phase 1: Create GitHub Remote & Push Code\n"
        "1. Open GitHub (github.com) and click 'New repository'.\n"
        "2. Name the repository: 2500320140058-p054-daily-sales (Public repository, no README or .gitignore initialization).\n"
        "3. In your local terminal, link and push the repository:\n"
    )

    git_cmds = """# Link local git repository to GitHub remote
git remote add origin https://github.com/<your-username>/2500320140058-p054-daily-sales.git

# Set primary branch to main and push commits
git branch -M main
git push -u origin main"""
    add_code_block(doc, "Git Push Commands", git_cmds)

    p_phase2 = doc.add_paragraph()
    p_phase2.paragraph_format.space_after = Pt(6)
    p_phase2.add_run(
        "Phase 2: Deploy on Streamlit Community Cloud (share.streamlit.io)\n"
        "1. Visit share.streamlit.io and sign in with your GitHub account.\n"
        "2. Click the 'Create app' button (top right corner).\n"
        "3. Configure the application deployment parameters:\n"
        "   • Repository: <your-username>/2500320140058-p054-daily-sales\n"
        "   • Branch: main\n"
        "   • Main file path: app.py\n"
        "   • App URL (Custom subdomain): 2500320140058-daily-sales.streamlit.app\n"
        "4. Click 'Deploy!' and monitor the build console output for 2-3 minutes as pip dependencies install.\n"
        "5. Once deployment completes, balloons appear and the live application is accessible globally."
    )

    # Section 3: Common Cloud Build Failures & Troubleshooting Guide
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Streamlit Cloud Troubleshooting & Diagnostic Guide")
    r_h3.font.name = "Calibri"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(30, 58, 138)

    troubleshoot_data = [
        ("FileNotFoundError: data/app.db",
         "Working directory mismatch in Linux container when loading pages/ sub-files.",
         "Use Path(__file__).resolve().parent.parent / 'data' / 'app.db' in src/data_loader.py rather than relative string literals."),
        ("ModuleNotFoundError: statsmodels / scipy",
         "Missing dependencies in requirements.txt or pip version mismatch.",
         "Ensure requirements.txt is committed at the repository root with exact package names (scipy, statsmodels, plotly)."),
        ("App reloads continuously / Memory limit",
         "Excessive uncached database reads consuming container RAM (> 1 GB limit).",
         "Ensure load_analytical_view() is decorated with @st.cache_data so the 40-col dataframe is read only once upon startup."),
        ("Broken Tailwind styles / Missing CDN",
         "Ad-blocker or iframe sandbox security blocking external CDN script.",
         "st.components.v1.html includes inline fallback CSS font definitions to ensure graceful degradation if CDN is throttled.")
    ]

    tb_table = doc.add_table(rows=len(troubleshoot_data)+1, cols=3)
    tb_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tb_table)
    tb_headers = ["Potential Cloud Error", "Root Cause Analysis", "Engineered Resolution & Safeguard"]
    for i, h in enumerate(tb_headers):
        style_header_cell(tb_table.rows[0].cells[i], h)

    for r_idx, vals in enumerate(troubleshoot_data):
        row = tb_table.rows[r_idx+1]
        for c_idx, text in enumerate(vals):
            style_data_cell(row.cells[c_idx], text, is_even=(r_idx%2==1), bold=(c_idx==0))

    doc_path = os.path.abspath("docs/16_Deployment_Guide.docx")
    doc.save(doc_path)
    print(f"Generated {doc_path} successfully.")

if __name__ == "__main__":
    create_deployment_guide_doc()
