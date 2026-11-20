"""
Generate 18_Presentation PowerPoint — Retail Sales Performance Analysis
Student: Apeksha Goel | Roll: 2500320140030 | MCA 2nd Yr DA, ABES
Step 18 — 11-Slide Institutional Presentation Deck
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pathlib import Path
import os
import win32com.client

# Paths
OUT_DIR = Path(r"c:\Users\apeks\Downloads\retail sales performance analysis\docs")
PPTX_RSPA = OUT_DIR / "RSPA_Presentation.pptx"
PPTX_TOCONFIRM = OUT_DIR / "To_confirm_Presentation.pptx"
PPTX_STANDARD = OUT_DIR / "18_Presentation_RetailSales_ApekshaGoel.pptx"

# Color Palette
C_RED = RGBColor(0xA5, 0x1C, 0x30)       # A51C30
C_BLACK = RGBColor(0x15, 0x15, 0x14)     # 151514
C_DARK = RGBColor(0x3A, 0x3A, 0x38)      # 3A3A38
C_GREY = RGBColor(0x5E, 0x5E, 0x5B)      # 5E5E5B
C_CARD = RGBColor(0xF4, 0xF1, 0xF1)      # F4F1F1
C_PINK = RGBColor(0xD9, 0xB5, 0xBC)      # D9B5BC
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)     # FFFFFF
C_BORDER = RGBColor(0xD9, 0xD9, 0xD9)    # D9D9D9
C_FRAME_BDR = RGBColor(0xBF, 0xBF, 0xBF) # BFBFBF
C_PG_NUM = RGBColor(0x8A, 0x8A, 0x86)    # 8A8A86

PROJECT_ID = "[To confirm]"
PROJECT_TITLE = "Retail Sales Performance Analysis"

def set_cell_border(cell, color="D9D9D9", width_pt=0.5):
    """Apply borders to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for border_name in ('lnL', 'lnR', 'lnT', 'lnB'):
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}')
        if existing is not None:
            tcPr.remove(existing)
        width_emu = int(width_pt * 12700)
        xml = f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width_emu}" cmpd="s"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:{border_name}>'
        tcPr.append(parse_xml(xml))

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    blank_layout = prs.slide_layouts[6]  # blank

    def add_card(slide, x, y, w, h):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_CARD
        shape.line.color.rgb = C_CARD
        shape.line.width = Pt(1)
        if hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
            shape.adjustments[0] = 0.08
        return shape

    def add_badge(slide, x, y, text):
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
        badge.fill.solid()
        badge.fill.fore_color.rgb = C_RED
        badge.line.fill.background()
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = C_WHITE
        return badge

    def add_head(slide, x, y, w, h, text):
        txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = txBox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = text
        run.font.name = 'Calibri'
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = C_BLACK
        return txBox

    def add_body(slide, x, y, w, h, text, font_size=13):
        txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = txBox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        
        lines = text.split('\n')
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = line
            run.font.name = 'Calibri'
            run.font.size = Pt(font_size)
            run.font.color.rgb = C_DARK
        return txBox

    def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text, body_font_size=13):
        add_card(slide, x, y, w, h)
        add_badge(slide, x + 0.2, y + 0.2, badge_text)
        add_head(slide, x + 0.75, y + 0.17, w - 0.9, 0.45, head_text)
        add_body(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, body_text, font_size=body_font_size)

    def add_plain_card(slide, x, y, w, h, head_text, body_text, body_font_size=13):
        add_card(slide, x, y, w, h)
        add_head(slide, x + 0.2, y + 0.17, w - 0.35, 0.45, head_text)
        add_body(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, body_text, font_size=body_font_size)

    def setup_content_slide(title_text, slide_num):
        slide = prs.slides.add_slide(blank_layout)
        
        # Title placeholder
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
        tf = title_box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = title_text
        r.font.name = 'Cambria'
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = C_BLACK

        # Footer
        footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
        tff = footer_box.text_frame
        tff.vertical_anchor = MSO_ANCHOR.MIDDLE
        pf = tff.paragraphs[0]
        pf.alignment = PP_ALIGN.LEFT
        rf = pf.add_run()
        rf.text = f"HCLTech Career Shaper ABES Engineering College | {PROJECT_ID} - {PROJECT_TITLE}"
        rf.font.name = 'Calibri'
        rf.font.size = Pt(9)
        rf.font.color.rgb = C_GREY

        # Slide Number
        num_box = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
        tfn = num_box.text_frame
        tfn.vertical_anchor = MSO_ANCHOR.MIDDLE
        pn = tfn.paragraphs[0]
        pn.alignment = PP_ALIGN.RIGHT
        rn = pn.add_run()
        rn.text = str(slide_num)
        rn.font.name = 'Calibri'
        rn.font.size = Pt(9)
        rn.font.color.rgb = C_PG_NUM

        return slide

    def format_table(table, col_widths, headers, data_rows):
        for col_idx, width in enumerate(col_widths):
            table.columns[col_idx].width = Inches(width)

        # Header row
        for col_idx, h_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_RED
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            set_cell_border(cell, color="D9D9D9", width_pt=0.5)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = h_text
            r.font.name = 'Calibri'
            r.font.size = Pt(13)
            r.font.bold = True
            r.font.color.rgb = C_WHITE

        # Data rows
        for row_idx, row_data in enumerate(data_rows, 1):
            table.rows[row_idx].height = Inches(0.5)
            for col_idx, val in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE
                cell.margin_left = Inches(0.1)
                cell.margin_right = Inches(0.1)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
                set_cell_border(cell, color="D9D9D9", width_pt=0.5)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                r = p.add_run()
                r.text = str(val)
                r.font.name = 'Calibri'
                r.font.size = Pt(13)
                r.font.bold = False
                r.font.color.rgb = C_BLACK

    # =========================================================================
    # SLIDE 1: DARK — Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg1.fill.solid(); bg1.fill.fore_color.rgb = C_BLACK; bg1.line.fill.background()

    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle
    b_logo = s1.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf = b_logo.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "[ABES logo]"; r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = C_PINK

    # Program header (2.2,.35,7.3,.5) 11 pink, right, middle
    b_prog = s1.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf = b_prog.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"; r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.color.rgb = C_PINK

    # Title (.5,1.4,9,1.6) Cambria 36 bold white, left, bottom
    b_title = s1.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(9.0), Inches(1.65))
    tf = b_title.text_frame; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    tf.word_wrap = True
    p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.LEFT
    r1 = p1.add_run(); r1.text = f"{PROJECT_ID} - Retail Sales"; r1.font.name = 'Cambria'; r1.font.size = Pt(36); r1.font.bold = True; r1.font.color.rgb = C_WHITE
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "Performance Analysis"; r2.font.name = 'Cambria'; r2.font.size = Pt(36); r2.font.bold = True; r2.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line
    b_lines = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf = b_lines.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP
    p1 = tf.paragraphs[0]; p1.space_before = Pt(4); r1 = p1.add_run(); r1.text = "Apeksha Goel | Roll No 2500320140030"; r1.font.name = 'Calibri'; r1.font.size = Pt(15); r1.font.color.rgb = C_CARD
    p2 = tf.add_paragraph(); p2.space_before = Pt(4); r2 = p2.add_run(); r2.text = "MCA 2nd Year Data Analyst | Mentor: Kalpana Dwivedi"; r2.font.name = 'Calibri'; r2.font.size = Pt(15); r2.font.color.rgb = C_CARD
    p3 = tf.add_paragraph(); p3.space_before = Pt(4); r3 = p3.add_run(); r3.text = "20 November 2026"; r3.font.name = 'Calibri'; r3.font.size = Pt(15); r3.font.color.rgb = C_PINK

    s1.notes_slide.notes_text_frame.text = (
        "Good morning mentors and evaluators. I am Apeksha Goel, Roll Number 2500320140030, presenting our MCA Data Analytics "
        "project: Retail Sales Performance Analysis. Built under the mentorship of Ms. Kalpana Dwivedi, our solution replaces "
        "scattered, error-prone spreadsheets with an interactive, filterable Streamlit dashboard that delivers real-time sales visibility, "
        "target tracking, and category margin analysis."
    )

    # =========================================================================
    # SLIDE 2: CONTENT — Problem and How Might We
    # =========================================================================
    s2 = setup_content_slide("Problem and How Might We", 2)
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem statement",
        "Retail store managers and regional heads rely on scattered manual spreadsheets to track sales performance across disparate stores.\n\n"
        "This reliance causes severe reporting lag, delays identification of underperforming products, and obscures regional revenue deficits, leading to misallocated inventory and missed targets.",
        body_font_size=12)

    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01",
        "How might we help store managers instantly identify underperforming products against sales targets without opening multiple spreadsheets?",
        body_font_size=12)

    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02",
        "How might we enable stakeholders to measure promotional revenue uplift and compare performance across configurable time periods?",
        body_font_size=12)

    s2.notes_slide.notes_text_frame.text = (
        "Retail managers currently lose critical days consolidating disparate spreadsheets to evaluate performance. Our problem statement "
        "addresses this operational lag. HMW-01 targets instant, spreadsheet-free target gap tracking for store managers, while "
        "HMW-02 provides regional heads with period-over-period comparative analytics and promotional uplift evaluation."
    )

    # =========================================================================
    # SLIDE 3: CONTENT — Stakeholders and users
    # =========================================================================
    s3 = setup_content_slide("Stakeholders and users", 3)
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Store Manager",
        "Primary user (SH-01) managing daily store operations.\n\n"
        "Needs real-time sales KPIs, fast-seller tracking, and daily target variance to adjust local promotions and restocking.\n\n"
        "Influence: High",
        body_font_size=11)

    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Regional Head",
        "Secondary user (SH-02) overseeing multi-store regional sales.\n\n"
        "Needs single-screen regional performance rollups and store variance comparisons to reallocate sales targets and budgets.\n\n"
        "Influence: High",
        body_font_size=11)

    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Mentor / Panel",
        "Academic mentor (Ms. Kalpana Dwivedi) & ABES evaluation panel (SH-06/07).\n\n"
        "Need verified EDA methodology, mathematical calculation accuracy, clean test coverage, and complete documentation.\n\n"
        "Influence: High",
        body_font_size=11)

    s3.notes_slide.notes_text_frame.text = (
        "We identified three primary stakeholder personas. Store managers are daily operational users needing immediate KPI feedback. "
        "Regional sales heads need multi-store rollup comparisons to make strategic resource allocations. Finally, our academic mentor "
        "and ABES examiners require rigorous statistical methods and full requirements traceability."
    )

    # =========================================================================
    # SLIDE 4: CONTENT — Requirements and traceability
    # =========================================================================
    s4 = setup_content_slide("Requirements and traceability", 4)
    table_shape4 = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    table4 = table_shape4.table
    widths4 = [1.0, 3.4, 1.0, 1.2, 1.2, 1.2]
    headers4 = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    data4 = [
        ["HMW-01", "Track sales vs target by product/store", "Must", "US-01", "TC-F-19", "UI-03"],
        ["HMW-02", "Compare period revenue (MoM / YoY)", "Must", "US-06", "TC-F-10", "UI-02"],
        ["HMW-06", "Profit margin % by product category", "Must", "US-07", "TC-F-11", "UI-04"],
        ["HMW-07", "Statistical discount vs margin analysis", "Must", "US-12", "TC-F-16", "UI-05"],
    ]
    format_table(table4, widths4, headers4, data4)

    # Note below table (.3 below table at y=4.1)
    b_note4 = s4.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(9.0), Inches(0.35))
    tf4 = b_note4.text_frame; tf4.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf4.paragraphs[0]; r = p.add_run()
    r.text = "Full matrix: Project Document, section 4"
    r.font.name = 'Calibri'; r.font.size = Pt(12); r.font.italic = True; r.font.color.rgb = C_GREY

    s4.notes_slide.notes_text_frame.text = (
        "This slide illustrates our strict end-to-end traceability matrix. Every user pain point flows from an HMW statement to a specific "
        "business requirement, user story, test case, and UI screen. Every Must-have requirement is verified by at least one functional "
        "and one technical test, guaranteeing zero orphan requirements in our architecture."
    )

    # =========================================================================
    # SLIDE 5: CONTENT — Data and preparation
    # =========================================================================
    s5 = setup_content_slide("Data and preparation", 5)
    
    # 4 stats: rows / columns / period / issues fixed
    stats5 = [
        (0.5, "240", "rows (fact_sales)"),
        (2.8, "35", "columns (analytics view)"),
        (5.1, "24 mo", "period (2023–2024)"),
        (7.4, "3", "issues fixed"),
    ]
    for x_pos, val, lbl in stats5:
        bx = s5.shapes.add_textbox(Inches(x_pos), Inches(1.3), Inches(2.1), Inches(0.8))
        tf = bx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = val; r.font.name = 'Cambria'; r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = C_RED

        bxl = s5.shapes.add_textbox(Inches(x_pos), Inches(2.1), Inches(2.1), Inches(0.35))
        tf = bxl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = lbl; r.font.name = 'Calibri'; r.font.size = Pt(12); r.font.color.rgb = C_GREY

    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source",
        "Deterministic synthetic dataset (seed 42) generated for data/app.db.\n\n"
        "Each fact row represents one retail sales line linking dim_store (10 stores across 5 regions), dim_product (20 SKUs in 5 categories), and dim_customer.",
        body_font_size=11)

    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning",
        "1. Missing payment method: 6 rows (2.5%) imputed to mode 'UPI' to satisfy schema NOT NULL.\n"
        "2. Duplicate billing lines: 3 duplicate transaction groups detected and validated.\n"
        "3. Outliers: 3 high-value bulk purchases retained within valid commercial ranges.",
        body_font_size=11)

    s5.notes_slide.notes_text_frame.text = (
        "Our data foundation is a 240-line normalized SQLite database covering 24 months from 2023 to 2024 across 10 stores. We deliberately "
        "built in real-world data hygiene challenges: 2.5% missing payment methods, duplicate transaction lines, and high-value bulk purchase outliers. "
        "Our automated cleaning pipeline resolves these defects while preserving authentic audit evidence."
    )

    # =========================================================================
    # SLIDE 6: CONTENT — Solution and technology
    # =========================================================================
    s6 = setup_content_slide("Solution and technology", 6)
    
    steps6 = [
        (0.5, "Raw data\n(CSV / xlsx)", False),
        (2.35, "Clean & KPIs\n(Python, pandas)", False),
        (4.2, "Store\n(SQLite app.db)", False),
        (6.05, "Dashboard\n(Streamlit)", True),
        (7.9, "Deploy\n(GitHub -> cloud)", False),
    ]
    for x, txt, is_active in steps6:
        shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(1.6), Inches(1.2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_RED if is_active else C_CARD
        shape.line.fill.background()
        if hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
            shape.adjustments[0] = 0.1
        tf = shape.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lines = txt.split('\n')
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = line; r.font.name = 'Calibri'; r.font.size = Pt(12); r.font.bold = True
            r.font.color.rgb = C_WHITE if is_active else C_BLACK

    chevrons6 = [2.1, 3.95, 5.8, 7.65]
    for cx in chevrons6:
        bx = s6.shapes.add_textbox(Inches(cx), Inches(1.5), Inches(0.25), Inches(1.2))
        tf = bx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = ">"; r.font.name = 'Calibri'; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = C_RED

    b_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf = b_tools.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run(); r1.text = "Tools: "; r1.font.bold = True; r1.font.size = Pt(13); r1.font.color.rgb = C_DARK
    r2 = p.add_run(); r2.text = "Python 3.12 · Streamlit · pandas · NumPy · SciPy · statsmodels · Plotly · SQLite · pytest"
    r2.font.size = Pt(13); r2.font.color.rgb = C_DARK

    b_dep = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf = b_dep.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    r.font.size = Pt(13); r.font.italic = True; r.font.color.rgb = C_GREY

    s6.notes_slide.notes_text_frame.text = (
        "Here is our complete technical architecture. Raw CSV data is cleaned and validated with pandas, loaded into an indexed SQLite "
        "database, and visualized via an interactive multi-page Streamlit dashboard using Plotly charts and Tailwind CSS. The entire "
        "stack relies on 100% open-source tools and deploys seamlessly to Streamlit Community Cloud."
    )

    # =========================================================================
    # SLIDE 7: CONTENT — Dashboard
    # =========================================================================
    s7 = setup_content_slide("Dashboard", 7)
    
    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid(); frame.fill.fore_color.rgb = C_CARD
    frame.line.color.rgb = C_FRAME_BDR
    frame.line.width = Pt(1)
    frame.line.dash_style = 2  # dashed

    tf = frame.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "[Paste a dashboard screenshot here]"
    r.font.name = 'Calibri'; r.font.size = Pt(14); r.font.italic = True; r.font.color.rgb = C_GREY

    # 4 Badges on right: BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4. Beside each, (7.25,y,2.25,.42), middle, 12 bold red "UI-01" + short black screen purpose
    badge_items7 = [
        (1.3, "1", "UI-01  ", "Executive KPIs & target tracker"),
        (2.2, "2", "UI-02  ", "Monthly trends & seasonal growth"),
        (3.1, "3", "UI-03  ", "Store variance & target comparison"),
        (4.0, "4", "UI-04  ", "Product EDA & Pareto 80/20 curve"),
    ]
    for y_pos, b_num, ui_tag, desc in badge_items7:
        add_badge(s7, 6.7, y_pos, b_num)
        bx = s7.shapes.add_textbox(Inches(7.25), Inches(y_pos), Inches(2.25), Inches(0.42))
        tf = bx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r1 = p.add_run(); r1.text = ui_tag; r1.font.name = 'Calibri'; r1.font.size = Pt(12); r1.font.bold = True; r1.font.color.rgb = C_RED
        r2 = p.add_run(); r2.text = desc; r2.font.name = 'Calibri'; r2.font.size = Pt(12); r2.font.color.rgb = C_BLACK

    s7.notes_slide.notes_text_frame.text = (
        "The application features eight modular screens. The UI-01 Overview screen presents an immediate executive health check with "
        "four responsive KPI cards. Screens UI-02 to UI-04 provide deep analytical dives into time-series growth, store target deficits, "
        "and Pareto 80/20 revenue concentration, all dynamically linked to our shared sidebar filters."
    )

    # =========================================================================
    # SLIDE 8: CONTENT — Key findings
    # =========================================================================
    s8 = setup_content_slide("Key findings", 8)
    
    findings8 = [
        (0.5, "88.9%", "Q4 Festive Surge", "Evidence: Q4 net sales ₹20.18k/line vs non-Q4 ₹4.74k (festive spike)."),
        (3.55, "-0.52", "Discount Margin Drag", "Evidence: Pearson r = -0.52; higher discounts erode gross margins."),
        (6.6, "+42.0%", "Annual Sales Growth", "Evidence: 2024 revenue reached ₹11.84L vs 2023 ₹8.34L (+42.0%)."),
    ]
    for x, num, meaning, evidence in findings8:
        add_card(s8, x, 1.3, 2.85, 3.6)
        
        # number y 1.5 h 1 Cambria 40 bold red, middle
        bx_n = s8.shapes.add_textbox(Inches(x + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf = bx_n.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = num; r.font.name = 'Cambria'; r.font.size = Pt(40); r.font.bold = True; r.font.color.rgb = C_RED

        # meaning y 2.6 h .9, 14 bold black, top
        bx_m = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf = bx_m.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = meaning; r.font.name = 'Calibri'; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = C_BLACK

        # Evidence: KPI or test y 3.6 h 1.1, 12 grey, top
        bx_e = s8.shapes.add_textbox(Inches(x + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf = bx_e.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = evidence; r.font.name = 'Calibri'; r.font.size = Pt(12); r.font.color.rgb = C_GREY

    s8.notes_slide.notes_text_frame.text = (
        "Our analysis revealed three major business insights. First, Q4 festive demand accounts for nearly 89% of Electronics net revenue, "
        "indicating a massive seasonal spike. Second, we proved a statistically significant inverse correlation of -0.52 between discounts "
        "and gross margins. Third, we documented an organic 42% year-over-year revenue expansion in 2024."
    )

    # =========================================================================
    # SLIDE 9: CONTENT — Testing and validation
    # =========================================================================
    s9 = setup_content_slide("Testing and validation", 9)
    
    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    b_pass = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf = b_pass.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "29 / 29"; r.font.name = 'Cambria'; r.font.size = Pt(44); r.font.bold = True; r.font.color.rgb = C_RED

    # (.5,2.4,2.6,.9) 13 grey, middle: test cases passed / pytest: n / manual: n
    b_pass_lbl = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf = b_pass_lbl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run(); r1.text = "test cases passed"; r1.font.name = 'Calibri'; r1.font.size = Pt(13); r1.font.color.rgb = C_GREY
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = "pytest: 29"; r2.font.name = 'Calibri'; r2.font.size = Pt(13); r2.font.color.rgb = C_GREY
    p3 = tf.add_paragraph(); p3.alignment = PP_ALIGN.CENTER
    r3 = p3.add_run(); r3.text = "manual: 24"; r3.font.name = 'Calibri'; r3.font.size = Pt(13); r3.font.color.rgb = C_GREY

    # Table at (3.4, 1.35), cols 1.1, 3.9, 1.1
    table_shape9 = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.0))
    table9 = table_shape9.table
    widths9 = [1.1, 3.9, 1.1]
    headers9 = ["Test", "Checks", "Result"]
    data9 = [
        ["TC-T-02", "Schema validation (dim_store, dim_product, fact_sales)", "Pass"],
        ["TC-T-07", "KPI calculation formulas (Revenue, AOV, Margin, Return)", "Pass"],
        ["TC-T-14", "App entrypoint cold startup latency (< 3.0 seconds)", "Pass"],
    ]
    format_table(table9, widths9, headers9, data9)

    s9.notes_slide.notes_text_frame.text = (
        "Quality assurance was built into every sprint. Our automated pytest suite executes 29 unit and integration tests covering database "
        "integrity, referential constraints, mathematical KPI precision, and headless Streamlit AppTest page rendering. The entire test suite "
        "completes in 6.07 seconds with 100% pass marks."
    )

    # =========================================================================
    # SLIDE 10: CONTENT — Recommendation, limitations and next steps
    # =========================================================================
    s10 = setup_content_slide("Recommendation, limitations and next steps", 10)
    
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation",
        "Protect Electronics margin by capping promotional discount rates below 20%.\n\n"
        "Prioritize inventory allocation to North and West regions 6 weeks prior to Q4 Diwali demand surge to prevent stockouts.",
        body_font_size=11)

    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations",
        "Analysis uses synthetic transactional data (240 lines).\n\n"
        "Findings demonstrate observable mathematical patterns but cannot support real-world causal or consumer behavioral claims without live POS feeds.",
        body_font_size=11)

    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next steps",
        "1. Integrate live POS API transactional streaming.\n\n"
        "2. Implement role-based access control for multi-store managers.\n\n"
        "3. Incorporate predictive time-series sales forecasting using Facebook Prophet.",
        body_font_size=11)

    s10.notes_slide.notes_text_frame.text = (
        "Based on our findings, we recommend capping promotional discounts at 20% to safeguard margins and pre-allocating inventory to North "
        "and West regions ahead of Q4. While our synthetic dataset successfully proves analytical mechanics, future iterations should integrate "
        "live POS feeds and predictive machine learning models."
    )

    # =========================================================================
    # SLIDE 11: DARK — Live demo and Q&A
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg11.fill.solid(); bg11.fill.fore_color.rgb = C_BLACK; bg11.line.fill.background()

    # [ABES logo]
    b_logo11 = s11.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf = b_logo11.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "[ABES logo]"; r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = C_PINK

    # Program header
    b_prog11 = s11.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf = b_prog11.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"; r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.color.rgb = C_PINK

    # Title (.5,1.55,9,1.5) Cambria 36 bold white
    b_title11 = s11.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf = b_title11.text_frame; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "Live demo and Q&A"; r.font.name = 'Cambria'; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1
    b_lines11 = s11.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf = b_lines11.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP
    p1 = tf.paragraphs[0]; p1.space_before = Pt(4); r1 = p1.add_run(); r1.text = "App: [To confirm]"; r1.font.name = 'Calibri'; r1.font.size = Pt(15); r1.font.color.rgb = C_CARD
    p2 = tf.add_paragraph(); p2.space_before = Pt(4); r2 = p2.add_run(); r2.text = "Code: [To confirm]"; r2.font.name = 'Calibri'; r2.font.size = Pt(15); r2.font.color.rgb = C_CARD
    p3 = tf.add_paragraph(); p3.space_before = Pt(4); r3 = p3.add_run(); r3.text = "Thank you"; r3.font.name = 'Calibri'; r3.font.size = Pt(15); r3.font.bold = True; r3.font.color.rgb = C_PINK

    s11.notes_slide.notes_text_frame.text = (
        "We are now ready for the live demonstration. I will switch over to our Streamlit dashboard to show real-time filtering, "
        "Plotly interactive charts, and data export in action. Thank you for your time, and I welcome any questions from the panel."
    )

    # Save to all required filenames
    prs.save(PPTX_RSPA)
    prs.save(PPTX_TOCONFIRM)
    prs.save(PPTX_STANDARD)

    print(f"Saved: {PPTX_RSPA}")
    print(f"Saved: {PPTX_TOCONFIRM}")
    print(f"Saved: {PPTX_STANDARD}")

def export_images():
    ppt_path = os.path.abspath(PPTX_STANDARD)
    out_dir = os.path.abspath(r"scratch\slide_images")
    os.makedirs(out_dir, exist_ok=True)
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        pres = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        pres.SaveAs(out_dir, 17) # 17 = ppSaveAsPNG / ppSaveAsJPG
        pres.Close()
        powerpoint.Quit()
        print("Re-exported slide images successfully.")
    except Exception as e:
        print("Export failed:", e)

if __name__ == "__main__":
    create_deck()
    export_images()
