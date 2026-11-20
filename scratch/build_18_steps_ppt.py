"""
Build 18-Step Complete Project Presentation Deck (PowerPoint & PDF)
Project: Retail Sales Performance Analysis
Student: Apeksha Goel | Roll No: 2500320140030 | MCA 2nd Year DA, ABES
Mentor: Ms. Kalpana Dwivedi
Covers all 18 SDLC Steps in a unified 21-slide master presentation.
"""

import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
import win32com.client

# Paths
DOCS_DIR = Path(r"c:\Users\apeks\Downloads\retail sales performance analysis\docs")
PPTX_MASTER = DOCS_DIR / "Retail_Sales_Performance_Analysis_18_Steps.pptx"
PPTX_ALT1 = DOCS_DIR / "18_Steps_Complete_Presentation.pptx"
PPTX_ALT2 = DOCS_DIR / "18_Steps_Presentation_RetailSales_ApekshaGoel.pptx"
PDF_OUT = DOCS_DIR / "18_Steps_Complete_Presentation.pdf"

ER_IMG_PATH = DOCS_DIR / "09_ER_Diagram.png"

# Color Palette (ABES specifications)
C_RED = RGBColor(0xA5, 0x1C, 0x30)       # A51C30
C_BLACK = RGBColor(0x15, 0x15, 0x14)     # 151514
C_DARK = RGBColor(0x3A, 0x3A, 0x38)      # 3A3A38
C_GREY = RGBColor(0x5E, 0x5E, 0x5B)      # 5E5E5B
C_CARD = RGBColor(0xF4, 0xF1, 0xF1)      # F4F1F1
C_PINK = RGBColor(0xD9, 0xB5, 0xBC)      # D9B5BC
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)     # FFFFFF
C_BORDER = RGBColor(0xD9, 0xD9, 0xD9)    # D9D9D9
C_PG_NUM = RGBColor(0x8A, 0x8A, 0x86)    # 8A8A86

PROJECT_ID = "[To confirm]"
PROJECT_TITLE = "Retail Sales Performance Analysis"

def set_cell_border(cell, color="D9D9D9", width_pt=0.5):
    tcPr = cell._tc.get_or_add_tcPr()
    for border_name in ('lnL', 'lnR', 'lnT', 'lnB'):
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}')
        if existing is not None:
            tcPr.remove(existing)
        width_emu = int(width_pt * 12700)
        xml = f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width_emu}" cmpd="s"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:{border_name}>'
        tcPr.append(parse_xml(xml))

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    blank_layout = prs.slide_layouts[6]

    def add_card(slide, x, y, w, h):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_CARD
        shape.line.color.rgb = C_CARD
        shape.line.width = Pt(1)
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
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = C_BLACK
        return txBox

    def add_body(slide, x, y, w, h, text, font_size=11):
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

    def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text, body_font_size=11):
        add_card(slide, x, y, w, h)
        add_badge(slide, x + 0.18, y + 0.18, badge_text)
        add_head(slide, x + 0.68, y + 0.15, w - 0.8, 0.45, head_text)
        add_body(slide, x + 0.18, y + 0.68, w - 0.36, max(0.2, h - 0.8), body_text, font_size=body_font_size)

    def add_plain_card(slide, x, y, w, h, head_text, body_text, body_font_size=11):
        add_card(slide, x, y, w, h)
        if h <= 0.85:
            txBox = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.08), Inches(w - 0.4), Inches(max(0.2, h - 0.16)))
            tf = txBox.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r1 = p.add_run(); r1.text = head_text + ": "; r1.font.name = 'Calibri'; r1.font.bold = True; r1.font.size = Pt(body_font_size + 0.5); r1.font.color.rgb = C_BLACK
            r2 = p.add_run(); r2.text = body_text; r2.font.name = 'Calibri'; r2.font.size = Pt(body_font_size); r2.font.color.rgb = C_DARK
        else:
            add_head(slide, x + 0.2, y + 0.15, w - 0.4, 0.45, head_text)
            add_body(slide, x + 0.2, y + 0.68, w - 0.4, max(0.2, h - 0.8), body_text, font_size=body_font_size)

    def setup_content_slide(title_text, slide_num):
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9.0), Inches(0.7))
        tf = title_box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = title_text
        r.font.name = 'Cambria'
        r.font.size = Pt(26)
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

    def format_table(table, col_widths, headers, data_rows, font_size=11):
        for col_idx, width in enumerate(col_widths):
            table.columns[col_idx].width = Inches(width)

        # Header row
        for col_idx, h_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_RED
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            set_cell_border(cell, color="D9D9D9", width_pt=0.5)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = h_text
            r.font.name = 'Calibri'
            r.font.size = Pt(font_size)
            r.font.bold = True
            r.font.color.rgb = C_WHITE

        # Data rows
        for row_idx, row_data in enumerate(data_rows, 1):
            table.rows[row_idx].height = Inches(0.4)
            for col_idx, val in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.04)
                set_cell_border(cell, color="D9D9D9", width_pt=0.5)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                r = p.add_run()
                r.text = str(val)
                r.font.name = 'Calibri'
                r.font.size = Pt(font_size)
                r.font.bold = False
                r.font.color.rgb = C_BLACK

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Theme)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg1.fill.solid(); bg1.fill.fore_color.rgb = C_BLACK; bg1.line.fill.background()

    b_logo = s1.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf = b_logo.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run(); r.text = "[ABES logo]"; r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = C_PINK

    b_prog = s1.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf = b_prog.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"; r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.color.rgb = C_PINK

    b_title = s1.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(9.0), Inches(1.65))
    tf = b_title.text_frame; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    p1 = tf.paragraphs[0]; r1 = p1.add_run(); r1.text = f"{PROJECT_ID} - Retail Sales"; r1.font.name = 'Cambria'; r1.font.size = Pt(36); r1.font.bold = True; r1.font.color.rgb = C_WHITE
    p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = "Performance Analysis (18 Steps)"; r2.font.name = 'Cambria'; r2.font.size = Pt(36); r2.font.bold = True; r2.font.color.rgb = C_WHITE

    b_lines = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf = b_lines.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP
    p1 = tf.paragraphs[0]; p1.space_before = Pt(4); r1 = p1.add_run(); r1.text = "Apeksha Goel | Roll No: 2500320140030"; r1.font.name = 'Calibri'; r1.font.size = Pt(15); r1.font.color.rgb = C_CARD
    p2 = tf.add_paragraph(); p2.space_before = Pt(4); r2 = p2.add_run(); r2.text = "MCA 2nd Year Data Analyst | Mentor: Ms. Kalpana Dwivedi"; r2.font.name = 'Calibri'; r2.font.size = Pt(15); r2.font.color.rgb = C_CARD
    p3 = tf.add_paragraph(); p3.space_before = Pt(4); r3 = p3.add_run(); r3.text = "20 November 2026 — Complete 18-Step SDLC Presentation"; r3.font.name = 'Calibri'; r3.font.size = Pt(15); r3.font.color.rgb = C_PINK

    s1.notes_slide.notes_text_frame.text = (
        "Good morning panel and mentors. I am Apeksha Goel, presenting our MCA Data Analytics capstone project: "
        "Retail Sales Performance Analysis. This deck systematically walks through every single stage of our 18-step "
        "software engineering and analytics lifecycle, from initial problem framing down to live cloud deployment and test verification."
    )

    # =========================================================================
    # SLIDE 2: 18-Step Delivery Roadmap
    # =========================================================================
    s2 = setup_content_slide("18-Step SDLC Delivery Roadmap", 2)
    phases = [
        ("Phase 1: Needs & Scope", "Steps 1–6\n• Problem & 10 OQs\n• 12 HMW Statements\n• MoSCoW Needs (13 BRs)\n• Bidirectional RTM (v1–v5)\n• BRD & PRD MVP Scope", 0.5),
        ("Phase 2: Architecture", "Steps 7–10\n• TRD (17 FRs, 8 NFRs)\n• Star Schema (data/app.db)\n• Crow's Foot ERD (1:N)\n• 35-col SQL Analysis View\n• Synthetic Pipeline (seed 42)", 2.8),
        ("Phase 3: Design & Test", "Steps 11–13\n• 17 User Stories (54 pts)\n• UI-01..08 Wireframes\n• Tailwind CSS via CDN\n• 24 Functional Tests\n• 18 Technical Pytest Cases", 5.1),
        ("Phase 4: Build & Cloud", "Steps 14–18\n• 16 BUILD Blueprint Tasks\n• Streamlit Multi-Page App\n• 29/29 Pytest (6.07s)\n• Streamlit Cloud Deploy\n• Institutional Docs & Deck", 7.4),
    ]
    for p_title, p_desc, x_pos in phases:
        add_plain_card(s2, x_pos, 1.15, 2.1, 3.8, p_title, p_desc, body_font_size=11)
    
    s2.notes_slide.notes_text_frame.text = (
        "Here is the high-level roadmap of our 18-step delivery process. It spans four rigorous phases: Business Needs & Scope, "
        "Data Architecture & Modeling, Agile Design & Test Engineering, and finally AI-Assisted Implementation, Cloud Deployment, "
        "and Institutional Documentation."
    )

    # =========================================================================
    # SLIDE 3: STEP 1 — Project Understanding & Problem Framing
    # =========================================================================
    s3 = setup_content_slide("STEP 1: Project Understanding & Problem Framing", 3)
    add_badge_card(s3, 0.5, 1.15, 4.3, 3.8, "1", "Operational Problem",
        "Problem Statement:\nRetail store managers and regional heads currently track multi-store performance using scattered spreadsheets and manual reports.\n\n"
        "This reliance causes severe 5-to-10 day reporting latency, obscures store-level target variance, and causes inventory misallocation.\n\n"
        "Ecosystem: 7 Stakeholders (SH-01 Store Mgr, SH-02 Regional Head, SH-06 Mentor, SH-07 Panel).",
        body_font_size=11)

    add_plain_card(s3, 5.2, 1.15, 4.3, 3.8, "10 Opportunity Questions (OQ-01..10)",
        "• OQ-01: Sales variance vs. monthly targets by store\n"
        "• OQ-02: Monthly revenue seasonality & MoM/YoY trend\n"
        "• OQ-03: Gross profit margin % across 5 product categories\n"
        "• OQ-04: Correlation between discount rate and profit margin\n"
        "• OQ-05: Average Order Value (AOV) across stores & regions\n"
        "• OQ-06: Pareto 80/20 revenue concentration across 20 SKUs\n"
        "• OQ-07: Return rate % across categories and payment modes\n"
        "• OQ-08: Customer payment distribution (UPI, Card, Cash)\n"
        "• OQ-09: Persistent store deficit gap flags (< 90% attainment)\n"
        "• OQ-10: Injected billing anomalies and data hygiene fixes",
        body_font_size=10.5)

    s3.notes_slide.notes_text_frame.text = (
        "In Step 1, we defined the business problem and established 10 concrete Opportunity Questions covering store variance, "
        "revenue trends, profit margins, discount impact, and data hygiene, grounded in seven distinct stakeholder personas."
    )

    # =========================================================================
    # SLIDE 4: STEP 2 — "How Might We" Statements
    # =========================================================================
    s4 = setup_content_slide("STEP 2: 'How Might We' Statements", 4)
    add_badge_card(s4, 0.5, 1.15, 4.3, 1.8, "P", "HMW-01: Target Tracking",
        "How might we help store managers instantly identify underperforming products against sales targets without opening multiple spreadsheets?\n\nPriority: High Impact, High Feasibility",
        body_font_size=11)
    add_badge_card(s4, 0.5, 3.15, 4.3, 1.8, "P", "HMW-02: Promotional Uplift",
        "How might we enable stakeholders to measure promotional revenue uplift and compare performance across configurable time periods?\n\nPriority: High Impact, High Feasibility",
        body_font_size=11)
    add_badge_card(s4, 5.2, 1.15, 4.3, 1.8, "P", "HMW-06: Category Margin",
        "How might we show merchandise planners which categories yield the highest gross margin vs discount erosion?\n\nPriority: High Impact, Medium Feasibility",
        body_font_size=11)
    add_badge_card(s4, 5.2, 3.15, 4.3, 1.8, "P", "HMW-07: Discount Correlation",
        "How might we statistically prove the impact of discounting on net product profitability for retail directors?\n\nPriority: High Impact, High Feasibility",
        body_font_size=11)

    s4.notes_slide.notes_text_frame.text = (
        "In Step 2, we reframed pain points into 12 HMW statements. The top four priority statements shown here formed the direct basis "
        "for our analytical features: target gap tracking, promotional uplift, category margins, and discount correlation."
    )

    # =========================================================================
    # SLIDE 5: STEP 3 — Needs Analysis & MoSCoW Prioritization
    # =========================================================================
    s5 = setup_content_slide("STEP 3: Needs Analysis & MoSCoW Prioritization", 5)
    t_s5 = s5.shapes.add_table(5, 4, Inches(0.5), Inches(1.2), Inches(9.0), Inches(2.5)).table
    w_s5 = [1.2, 1.2, 4.6, 2.0]
    h_s5 = ["BR ID", "MoSCoW", "Business Requirement Description", "Quantitative Acceptance"]
    d_s5 = [
        ["BR-01", "Must", "Track actual sales vs pre-set targets by product & store", "Variance % computed in < 50 ms"],
        ["BR-02", "Must", "Analyze monthly revenue trends and 24-month seasonality", "MoM / YoY curve updated on filter"],
        ["BR-03", "Must", "Calculate and rank gross profit margins across categories", "Gross Margin % formula: (Rev-COGS)/Rev"],
        ["BR-04", "Must", "Identify negative correlation between discounts and margins", "Pearson r correlation & p-value shown"],
    ]
    format_table(t_s5, w_s5, h_s5, d_s5)

    add_plain_card(s5, 0.5, 3.9, 9.0, 1.1, "MoSCoW Breakdown",
        "Total Requirements: 13 | 7 Must-Have (core MVP functionality) | 4 Should-Have (AOV, Pareto 80/20, return rates) | 2 Could-Have (predictive forecast, export) | 0 Won't-Have.",
        body_font_size=11)

    s5.notes_slide.notes_text_frame.text = (
        "In Step 3, we prioritized 13 Business Requirements using MoSCoW. Seven Must-Have requirements formed our non-negotiable "
        "MVP foundation, each bound to strict quantitative acceptance metrics."
    )

    # =========================================================================
    # SLIDE 6: STEP 4 — Requirements Traceability Matrix (RTM)
    # =========================================================================
    s6 = setup_content_slide("STEP 4: Requirements Traceability Matrix (RTM)", 6)
    t_s6 = s6.shapes.add_table(5, 6, Inches(0.5), Inches(1.2), Inches(9.0), Inches(2.5)).table
    w_s6 = [1.0, 1.0, 2.8, 1.2, 1.5, 1.5]
    h_s6 = ["HMW", "BR ID", "Feature Description", "Priority", "Test Cases", "UI Screen"]
    d_s6 = [
        ["HMW-01", "BR-01", "Sales vs Target Gap Tracking", "Must", "TC-F-19, TC-T-07", "UI-03 Compare"],
        ["HMW-02", "BR-02", "MoM / YoY Trend Analysis", "Must", "TC-F-10, TC-T-08", "UI-02 Trends"],
        ["HMW-06", "BR-03", "Category Margin Ranking", "Must", "TC-F-11, TC-T-09", "UI-04 Product EDA"],
        ["HMW-07", "BR-04", "Discount vs Margin Correlation", "Must", "TC-F-16, TC-T-10", "UI-05 Insights"],
    ]
    format_table(t_s6, w_s6, h_s6, d_s6)

    add_plain_card(s6, 0.5, 3.9, 9.0, 1.1, "RTM Evolution (v1 to v5 Final)",
        "The RTM maintained 100% bidirectional traceability across 5 versions. Every requirement maps cleanly from an HMW statement down to a database entity, user story, UI screen, and automated pytest test case with zero orphan elements.",
        body_font_size=11)

    s6.notes_slide.notes_text_frame.text = (
        "Step 4 produced the RTM. By mapping every single business requirement to a user story, test case, and UI screen, "
        "we guaranteed zero scope creep and complete testability across our software architecture."
    )

    # =========================================================================
    # SLIDE 7: STEP 5 — Business Requirements Document (BRD)
    # =========================================================================
    s7 = setup_content_slide("STEP 5: Business Requirements Document (BRD)", 7)
    add_badge_card(s7, 0.5, 1.15, 4.3, 3.8, "5", "BRD Governance & RACI",
        "Formal 13-Section Specification:\n"
        "• Governance RACI Matrix: Store Managers (R), Regional Heads (A), Planners (C), Finance (I).\n\n"
        "• Operational Scope: 10 physical stores across 5 geographical regions.\n\n"
        "• Reporting Latency SLA: Sub-second dashboard refresh replacing 5-day manual turnaround.",
        body_font_size=11)

    add_plain_card(s7, 5.2, 1.15, 4.3, 3.8, "Core Mathematical KPI Formulas",
        "• Net Sales = (Qty × Unit Price) × (1 - Discount Rate)\n\n"
        "• COGS = Qty × Unit Cost\n\n"
        "• Gross Profit Margin % = ((Net Sales - COGS) / Net Sales) × 100\n\n"
        "• Target Variance % = ((Actual Sales - Target Sales) / Target Sales) × 100\n\n"
        "• AOV = Total Net Sales / Distinct Orders\n\n"
        "• Return Rate % = (Returned Units / Sold Units) × 100",
        body_font_size=11)

    s7.notes_slide.notes_text_frame.text = (
        "In Step 5, the BRD codified stakeholder governance and standardized our mathematical KPI definitions. Having exact formulas "
        "for Net Sales, Margin %, and Target Variance ensured mathematical consistency between SQL queries and UI cards."
    )

    # =========================================================================
    # SLIDE 8: STEP 6 — Product Requirements Document (PRD) & MVP
    # =========================================================================
    s8 = setup_content_slide("STEP 6: Product Requirements & MVP Scope", 8)
    add_badge_card(s8, 0.5, 1.15, 4.3, 3.8, "6", "MVP Feature Boundaries",
        "Included in MVP (Scope In):\n"
        "• FEAT-01: Executive KPI health check cards\n"
        "• FEAT-02: 24-month interactive time-series line\n"
        "• FEAT-03: Multi-store target deficit comparison\n"
        "• FEAT-04: Category margin & Pareto 80/20 curve\n"
        "• FEAT-05: Pearson correlation scatter analysis\n"
        "• FEAT-06: Raw transactional table & CSV export",
        body_font_size=11)

    add_plain_card(s8, 5.2, 1.15, 4.3, 3.8, "Release Governance & User Journeys",
        "Deferred to Phase 2 (Scope Out):\n"
        "• Live real-time POS transactional streaming API\n"
        "• Multi-role authentication & user write access\n"
        "• Automated inventory reordering webhook\n\n"
        "Primary User Journey:\n"
        "Store Manager opens UI-01 → views top-level target deficit flag → drills down into UI-03 store variance → identifies underperforming SKU in UI-04.",
        body_font_size=11)

    s8.notes_slide.notes_text_frame.text = (
        "Step 6 established our MVP product scope. We focused on core decision workflows—executive KPIs, comparative trends, "
        "and target gap detection—while consciously deferring complex features like live POS streaming to future releases."
    )

    # =========================================================================
    # SLIDE 9: STEP 7 — Technical Requirements Document (TRD)
    # =========================================================================
    s9 = setup_content_slide("STEP 7: Technical Requirements Document (TRD)", 9)
    add_badge_card(s9, 0.5, 1.15, 4.3, 3.8, "7", "Technical Architecture",
        "System Architecture Layers:\n"
        "• Data Tier: Normalized SQLite database (data/app.db) + indexed analytical view (v_sales_analysis).\n\n"
        "• Logic Tier: Pure Python functions (src/analysis.py, src/data_loader.py) with pandas and SciPy.\n\n"
        "• Presentation Tier: Streamlit multi-page frontend styled with isolated Tailwind CSS components.",
        body_font_size=11)

    add_plain_card(s9, 5.2, 1.15, 4.3, 3.8, "Non-Functional Requirements (NFRs)",
        "• NFR-01 (Latency): Cold startup < 3.0s; page filter re-renders < 500 ms.\n\n"
        "• NFR-02 (Integrity): Zero NULL primary keys; foreign key constraints enforced in SQLite.\n\n"
        "• NFR-04 (Portability): 100% open-source Python 3.12 stack; cross-platform compatibility.\n\n"
        "• NFR-05 (Accessibility): WCAG AA contrast ratio compliance using Tailwind tokens.",
        body_font_size=11)

    s9.notes_slide.notes_text_frame.text = (
        "Step 7 detailed the technical blueprint. We decoupled data access, analytics logic, and presentation across three tiers, "
        "and established rigorous Non-Functional Requirements including sub-3-second startup and WCAG AA contrast compliance."
    )

    # =========================================================================
    # SLIDE 10: STEP 8 — Database Design & Schema Architecture
    # =========================================================================
    s10 = setup_content_slide("STEP 8: Database Design & Schema Architecture", 10)
    t_s10 = s10.shapes.add_table(6, 4, Inches(0.5), Inches(1.15), Inches(9.0), Inches(2.6)).table
    w_s10 = [1.6, 1.4, 3.8, 2.2]
    h_s10 = ["Table Name", "Entity Type", "Attributes & Scope", "Primary / Foreign Keys"]
    d_s10 = [
        ["dim_store", "Dimension", "10 stores across 5 regions", "PK: store_id"],
        ["dim_product", "Dimension", "20 SKUs in 5 categories", "PK: product_id"],
        ["dim_customer", "Dimension", "50 customer demographic records", "PK: customer_id"],
        ["fact_sales", "Fact Table", "240 transactional billing rows (2023-24)", "PK: transaction_id, FKs: store, product, cust"],
        ["sales_targets", "Fact / Target", "4,800 monthly store-product targets", "FKs: store_id, product_id, year_month"],
    ]
    format_table(t_s10, w_s10, h_s10, d_s10)

    add_plain_card(s10, 0.5, 3.9, 9.0, 1.1, "Analytical View (v_sales_analysis)",
        "A 35-column denormalized analytical SQL view pre-joins all dimension attributes, computes profit margins, discount amounts, and target variances, delivering sub-50 ms query response times.",
        body_font_size=11)

    s10.notes_slide.notes_text_frame.text = (
        "Step 8 defined our Star Schema in SQLite. By pre-joining dimensions into a 35-column view named v_sales_analysis, "
        "we eliminated repetitive multi-table joins at runtime, achieving sub-50-millisecond query performance."
    )

    # =========================================================================
    # SLIDE 11: STEP 9 — Entity Relationship Diagram (ERD)
    # =========================================================================
    s11 = setup_content_slide("STEP 9: Entity Relationship Diagram (ERD)", 11)
    if ER_IMG_PATH.exists():
        s11.shapes.add_picture(str(ER_IMG_PATH), Inches(0.5), Inches(1.15), width=Inches(5.6))
    
    add_badge_card(s11, 6.3, 1.15, 3.2, 3.8, "9", "Relational Integrity",
        "Crow's Foot Cardinalities:\n\n"
        "• dim_store (1) ──< (N) fact_sales\n"
        "• dim_product (1) ──< (N) fact_sales\n"
        "• dim_customer (1) ──< (N) fact_sales\n"
        "• dim_store (1) ──< (N) sales_targets\n"
        "• dim_product (1) ──< (N) sales_targets\n\n"
        "Enforced via SQLite PRAGMA foreign_keys = ON; prevents orphan sales or unmapped target records.",
        body_font_size=10.5)

    s11.notes_slide.notes_text_frame.text = (
        "Step 9 produced the Crow's Foot ER Diagram. It clearly shows one-to-many relationships from store, product, and customer "
        "dimensions into fact_sales and sales_targets, enforced by active foreign key constraints in SQLite."
    )

    # =========================================================================
    # SLIDE 12: STEP 10 — Synthetic Data Engineering & Hygiene
    # =========================================================================
    s12 = setup_content_slide("STEP 10: Synthetic Data Engineering & Hygiene", 12)
    stats12 = [
        (0.5, "240", "Fact Sales Lines"),
        (2.8, "4,800", "Monthly Targets"),
        (5.1, "Seed 42", "Deterministic Generation"),
        (7.4, "3 Types", "Hygiene Defects Fixed"),
    ]
    for x_pos, val, lbl in stats12:
        bx = s12.shapes.add_textbox(Inches(x_pos), Inches(1.15), Inches(2.1), Inches(0.75))
        tf = bx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = val; r.font.name = 'Cambria'; r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = C_RED

        bxl = s12.shapes.add_textbox(Inches(x_pos), Inches(1.9), Inches(2.1), Inches(0.35))
        tf = bxl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = lbl; r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.color.rgb = C_GREY

    add_plain_card(s12, 0.5, 2.4, 4.3, 2.6, "Injected Data Anomalies",
        "Deliberately built-in edge cases to test pipeline resilience:\n\n"
        "1. Missing Payment Method: 6 rows (2.5%) generated with NULL payment method.\n"
        "2. Duplicate Billing: 3 duplicate transaction line groups intentionally created.\n"
        "3. Bulk Outliers: 3 high-value transactions (quantities 8-15) created.",
        body_font_size=11)

    add_plain_card(s12, 5.2, 2.4, 4.3, 2.6, "Automated Cleaning Solutions",
        "Pipeline Resolution Logic:\n\n"
        "1. Mode Imputation: Missing payment methods imputed to modal 'UPI' to satisfy schema constraints.\n"
        "2. PK Deduplication: Duplicates flagged and resolved during SQLite ingest.\n"
        "3. Outlier Validation: Bulk orders audited and validated within legitimate wholesale limits.",
        body_font_size=11)

    s12.notes_slide.notes_text_frame.text = (
        "In Step 10, we generated 240 deterministic fact lines covering 2023 to 2024. We deliberately injected three data hygiene flaws—missing "
        "payment modes, duplicates, and outliers—and built automated cleaning pipelines to impute and validate them."
    )

    # =========================================================================
    # SLIDE 13: STEP 11 — User Stories & Agile Sprint Planning
    # =========================================================================
    s13 = setup_content_slide("STEP 11: User Stories & Agile Sprint Planning", 13)
    add_badge_card(s13, 0.5, 1.15, 4.3, 3.8, "11", "17 User Stories (54 Points)",
        "Structured into 5 Epics:\n"
        "• Epic 1: Data Ingestion & Storage (US-01..03, 10 pts)\n"
        "• Epic 2: Executive Dashboard (US-04..06, 12 pts)\n"
        "• Epic 3: Comparative Analysis (US-07..09, 11 pts)\n"
        "• Epic 4: Statistical Insights (US-10..13, 13 pts)\n"
        "• Epic 5: Quality & Governance (US-14..17, 8 pts)\n\n"
        "Sample Story US-01:\n'As a Store Manager, I want to view my weekly sales vs target so that I can correct underperformance immediately.'",
        body_font_size=10.5)

    add_plain_card(s13, 5.2, 1.15, 4.3, 3.8, "8-Day Agile Sprint Execution Plan",
        "Sprint 1 (Days 1–2): Data Foundation & Storage\n"
        "• Setup SQLite schema, generate synthetic data, build cleaning pipeline.\n\n"
        "Sprint 2 (Days 3–4): Analytics Core & UI Skeleton\n"
        "• Develop KPI calculation functions and Tailwind UI shell.\n\n"
        "Sprint 3 (Days 5–6): Advanced Visualizations\n"
        "• Build Plotly trendlines, store target variance, and Pareto curves.\n\n"
        "Sprint 4 (Days 7–8): Test Coverage & Cloud Deploy\n"
        "• Execute 29 pytest cases, verify smoke matrix, deploy to Streamlit Cloud.",
        body_font_size=10.5)

    s13.notes_slide.notes_text_frame.text = (
        "Step 11 organized requirements into 17 user stories with Gherkin acceptance criteria across 5 epics, "
        "scheduled over an 8-day sprint cycle that ensured incremental delivery and continuous verification."
    )

    # =========================================================================
    # SLIDE 14: STEP 12 — UI Framework & Wireframe Design
    # =========================================================================
    s14 = setup_content_slide("STEP 12: UI Framework & Wireframe Design", 14)
    screens14 = [
        ("UI-01 Overview", "Executive health check: 4 KPI cards & target flag", 0.5, 1.15),
        ("UI-02 Trends", "Monthly MoM line, YoY growth & seasonality", 2.8, 1.15),
        ("UI-03 Compare", "Store target variance & regional AOV benchmark", 5.1, 1.15),
        ("UI-04 Product EDA", "Category gross margin ranking & Pareto 80/20", 7.4, 1.15),
        ("UI-05 Insights", "Pearson correlation scatter & discount analysis", 0.5, 2.75),
        ("UI-06 Data Explorer", "Interactive raw data grid with CSV export", 2.8, 2.75),
        ("UI-07 Data Quality", "Audit dashboard for missing values & outliers", 5.1, 2.75),
        ("UI-08 About", "System architecture, methodology & metadata", 7.4, 2.75),
    ]
    for s_name, s_desc, x_pos, y_pos in screens14:
        add_plain_card(s14, x_pos, y_pos, 2.1, 1.45, s_name, s_desc, body_font_size=10)

    add_plain_card(s14, 0.5, 4.35, 9.0, 0.75, "UI Technology & State",
        "Tailwind CSS loaded via CDN iframe helper in src/ui.py. Shared sidebar filters synchronized seamlessly across all 8 screens using st.session_state.",
        body_font_size=10.5)

    s14.notes_slide.notes_text_frame.text = (
        "Step 12 designed the user interface. We created eight modular screens styled with Tailwind CSS, sharing global filter state "
        "through Streamlit session state so user selections persist seamlessly across every page."
    )

    # =========================================================================
    # SLIDE 15: STEP 13 — Test Case Specification
    # =========================================================================
    s15 = setup_content_slide("STEP 13: Test Case Specification", 15)
    t_s15 = s15.shapes.add_table(5, 4, Inches(0.5), Inches(1.15), Inches(9.0), Inches(2.6)).table
    w_s15 = [1.2, 2.4, 3.8, 1.6]
    h_s15 = ["Test ID", "Type & Module", "Verification Assertion & Check", "Execution Status"]
    d_s15 = [
        ["TC-F-01", "Functional / UI-01", "Cold start renders 4 KPI cards with valid positive numbers", "Pass (Manual)"],
        ["TC-T-02", "Technical / test_data", "Validate SQLite row counts, schema types, and NOT NULL rules", "Pass (Pytest)"],
        ["TC-T-07", "Technical / test_analysis", "Verify mathematical KPI formulas against micro-test dataset", "Pass (Pytest)"],
        ["TC-T-14", "Technical / test_app", "Streamlit AppTest headless execution latency (< 3.0 seconds)", "Pass (Pytest)"],
    ]
    format_table(t_s15, w_s15, h_s15, d_s15)

    add_plain_card(s15, 0.5, 3.9, 9.0, 1.1, "Test Architecture Summary",
        "Total Test Suite: 42 Cases | 24 Functional Acceptance Tests (TC-F-01..24) covering UI filters, KPIs, and empty result handling | 18 Technical Automated Tests (TC-T-01..18) implemented in pytest.",
        body_font_size=11)

    s15.notes_slide.notes_text_frame.text = (
        "In Step 13, we authored 42 test cases. This included 24 functional acceptance checks for UI behavior and 18 technical automated "
        "unit and integration tests in pytest covering schema validity, formula accuracy, and cold boot speed."
    )

    # =========================================================================
    # SLIDE 16: STEP 14 — AI Build Plan Blueprint
    # =========================================================================
    s16 = setup_content_slide("STEP 14: AI Build Plan Blueprint", 16)
    add_badge_card(s16, 0.5, 1.15, 4.3, 3.8, "14", "BUILD_PLAN.md Architecture",
        "16 Sequential AI Execution Tasks:\n"
        "• BUILD-01: Virtual environment & pinned requirements\n"
        "• BUILD-02: Data generator & cleaning pipeline\n"
        "• BUILD-03: Analytical functions & KPI formulas\n"
        "• BUILD-04: UI shell & Tailwind CSS iframe card helper\n"
        "• BUILD-05..12: Implementation of 8 Streamlit pages\n"
        "• BUILD-13: Automated pytest suite execution\n"
        "• BUILD-14: Documentation & verification audit\n"
        "• BUILD-15..16: Cloud deploy & smoke test matrix",
        body_font_size=10.5)

    add_plain_card(s16, 5.2, 1.15, 4.3, 3.8, "Definition of Done (DoD) Criteria",
        "Strict Completion Gates for Every Task:\n\n"
        "1. All Must-Have business requirements implemented.\n"
        "2. Automated pytest suite runs with 100% pass marks.\n"
        "3. Zero hardcoded local file paths; relative Path used.\n"
        "4. Application cold startup latency verified under 3.0s.\n"
        "5. Git commit executed with conventional task message.\n\n"
        "Result: Entire codebase built predictably with zero regressions.",
        body_font_size=11)

    s16.notes_slide.notes_text_frame.text = (
        "Step 14 created our AI build blueprint. Divided into 16 discrete, test-driven tasks with strict Definition of Done gates, "
        "it allowed rapid, systematic implementation without architectural drift."
    )

    # =========================================================================
    # SLIDE 17: STEP 15 — Build Execution & Pytest Results
    # =========================================================================
    s17 = setup_content_slide("STEP 15: Build Execution & Pytest Results", 17)
    
    # Left stat box
    bx = s17.shapes.add_textbox(Inches(0.5), Inches(1.15), Inches(2.6), Inches(1.0))
    tf = bx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run(); r.text = "29 / 29"; r.font.name = 'Cambria'; r.font.size = Pt(44); r.font.bold = True; r.font.color.rgb = C_RED

    bx_lbl = s17.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(2.6), Inches(0.9))
    tf = bx_lbl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER; r1 = p1.add_run(); r1.text = "100% test cases passed"; r1.font.bold = True; r1.font.size = Pt(13); r1.font.color.rgb = C_DARK
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER; r2 = p2.add_run(); r2.text = "pytest: 29 in 6.07s\nManual checks: 24"; r2.font.size = Pt(12); r2.font.color.rgb = C_GREY

    t_s17 = s17.shapes.add_table(4, 3, Inches(3.4), Inches(1.15), Inches(6.1), Inches(2.0)).table
    w_s17 = [1.8, 3.2, 1.1]
    h_s17 = ["Pytest Test Module", "Focus & Coverage", "Result"]
    d_s17 = [
        ["tests/test_data.py", "8 tests: SQLite schema, PKs, FKs, mode impute", "8/8 Pass"],
        ["tests/test_analysis.py", "13 tests: KPI formulas, YoY, Pearson r, Pareto", "13/13 Pass"],
        ["tests/test_app.py", "8 tests: Headless AppTest boot, filters, export", "8/8 Pass"],
    ]
    format_table(t_s17, w_s17, h_s17, d_s17)

    add_plain_card(s17, 0.5, 3.8, 9.0, 1.2, "Codebase Deliverable Summary",
        "Successfully implemented modular Python package: app.py entrypoint, 8 pages in pages/, analytical engine in src/analysis.py, database manager in src/data_loader.py, and Tailwind components in src/ui.py. All 29 automated tests pass in 6.07 seconds.",
        body_font_size=11)

    s17.notes_slide.notes_text_frame.text = (
        "In Step 15, we completed implementation. Our automated pytest suite ran 29 unit and integration tests across data, "
        "analytics, and headless UI rendering, passing 100% in just 6.07 seconds."
    )

    # =========================================================================
    # SLIDE 18: STEP 16 — Streamlit Cloud Deployment
    # =========================================================================
    s18 = setup_content_slide("STEP 16: Streamlit Cloud Deployment", 18)
    add_badge_card(s18, 0.5, 1.15, 4.3, 3.8, "16", "Cloud Deployment Architecture",
        "Streamlit Community Cloud (share.streamlit.io):\n\n"
        "• Git-Driven CI/CD: Connected to GitHub main branch; automatic rebuild on git push.\n\n"
        "• Pinned Dependencies: requirements.txt with exact versions (streamlit==1.42.0, pandas==2.2.3, plotly==6.0.0).\n\n"
        "• Zero Hardcoded Paths: Relative filesystem resolution via Path(__file__).parent / 'data'.",
        body_font_size=11)

    add_plain_card(s18, 5.2, 1.15, 4.3, 3.8, "7-Point Pre-Flight Smoke Test Matrix",
        "• SMK-01: App boots cleanly without module import errors\n"
        "• SMK-02: SQLite database connects with foreign keys active\n"
        "• SMK-03: Sidebar region & category dropdowns populate\n"
        "• SMK-04: Filter changes trigger instant UI updates (<500ms)\n"
        "• SMK-05: Plotly trendline and bar charts render correctly\n"
        "• SMK-06: CSV raw data export generates valid download file\n"
        "• SMK-07: Responsive layout renders cleanly on mobile view",
        body_font_size=11)

    s18.notes_slide.notes_text_frame.text = (
        "Step 16 packaged our app for Streamlit Community Cloud. We pinned all dependencies, ensured relative file paths, "
        "and established a 7-point smoke test matrix to verify production stability."
    )

    # =========================================================================
    # SLIDE 19: STEP 17 — Official ABES Project Document
    # =========================================================================
    s19 = setup_content_slide("STEP 17: Official ABES Project Document", 19)
    add_badge_card(s19, 0.5, 1.15, 4.3, 3.8, "17", "Institutional Document Specifications",
        "ABES Compliance Criteria:\n\n"
        "• Layout: A4 portrait, 2 cm margins, fixed 9,638 dxa tables.\n"
        "• Typography: Georgia bold headings, Calibri 10.5 pt body text.\n"
        "• Palette: #A51C30 crimson primary header fill, #151514 dark body.\n"
        "• Dynamic Word XML: Live automated TOC and native PAGE fields.\n"
        "• Scope: 30 comprehensive pages covering Steps 1 through 13.",
        body_font_size=11)

    add_plain_card(s19, 5.2, 1.15, 4.3, 3.8, "Delivered Document Artifacts",
        "Delivered Official Files in docs/:\n\n"
        "• 17_Project_Document_RetailSales_ApekshaGoel.docx (Word)\n"
        "• 17_Project_Document_RetailSales_ApekshaGoel.pdf (PDF)\n"
        "• RSPA_Project_Document.docx & To_confirm_Project_Document.docx\n\n"
        "Full verification: Table borders 0.5 pt, no row splits across page breaks, and zero fabricated numbers or links.",
        body_font_size=11)

    s19.notes_slide.notes_text_frame.text = (
        "Step 17 generated our formal 30-page ABES Project Document in Word and PDF formats, adhering exactly to institutional "
        "styling, fixed 9,638 dxa table widths, crimson color tokens, and live Word XML pagination."
    )

    # =========================================================================
    # SLIDE 20: STEP 18 — Presentation Deck & Key Findings
    # =========================================================================
    s20 = setup_content_slide("STEP 18: Key Analytical Findings", 20)
    findings20 = [
        (0.5, "88.9%", "Q4 Festive Surge", "Evidence: Q4 net sales ₹20.18k/line vs non-Q4 ₹4.74k (festive spike in Electronics)."),
        (3.55, "-0.52", "Discount Margin Drag", "Evidence: Pearson r = -0.52; higher promotional discounts directly erode gross margin."),
        (6.6, "+42.0%", "Annual Sales Growth", "Evidence: 2024 revenue reached ₹11.84L vs 2023 ₹8.34L (+42.0% across all 10 stores)."),
    ]
    for x, num, meaning, evidence in findings20:
        add_card(s20, x, 1.15, 2.85, 3.8)
        bx_n = s20.shapes.add_textbox(Inches(x + 0.2), Inches(1.35), Inches(2.45), Inches(0.9))
        tf = bx_n.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = num; r.font.name = 'Cambria'; r.font.size = Pt(38); r.font.bold = True; r.font.color.rgb = C_RED

        bx_m = s20.shapes.add_textbox(Inches(x + 0.2), Inches(2.35), Inches(2.45), Inches(0.8))
        tf = bx_m.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = meaning; r.font.name = 'Calibri'; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = C_BLACK

        bx_e = s20.shapes.add_textbox(Inches(x + 0.2), Inches(3.25), Inches(2.45), Inches(1.4))
        tf = bx_e.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = evidence; r.font.name = 'Calibri'; r.font.size = Pt(11.5); r.font.color.rgb = C_GREY

    s20.notes_slide.notes_text_frame.text = (
        "In Step 18, our final analytical evaluation proved three critical findings: an 88.9% Q4 festive surge in Electronics, "
        "a statistically significant -0.52 discount-to-margin correlation, and 42% organic annual sales expansion."
    )

    # =========================================================================
    # SLIDE 21: Live Demo & Viva Voce Q&A (Dark Theme)
    # =========================================================================
    s21 = prs.slides.add_slide(blank_layout)
    bg21 = s21.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg21.fill.solid(); bg21.fill.fore_color.rgb = C_BLACK; bg21.line.fill.background()

    b_logo21 = s21.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf = b_logo21.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run(); r.text = "[ABES logo]"; r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = C_PINK

    b_prog21 = s21.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf = b_prog21.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"; r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.color.rgb = C_PINK

    b_title21 = s21.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(9.0), Inches(1.5))
    tf = b_title21.text_frame; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    r = tf.paragraphs[0].add_run(); r.text = "Live Demonstration & Viva Q&A"; r.font.name = 'Cambria'; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = C_WHITE

    b_lines21 = s21.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(9.0), Inches(2.0))
    tf = b_lines21.text_frame; tf.vertical_anchor = MSO_ANCHOR.TOP
    p1 = tf.paragraphs[0]; p1.space_before = Pt(4); r1 = p1.add_run(); r1.text = "Streamlit Cloud Live App: [To confirm]"; r1.font.name = 'Calibri'; r1.font.size = Pt(15); r1.font.color.rgb = C_CARD
    p2 = tf.add_paragraph(); p2.space_before = Pt(4); r2 = p2.add_run(); r2.text = "GitHub Repository: [To confirm]"; r2.font.name = 'Calibri'; r2.font.size = Pt(15); r2.font.color.rgb = C_CARD
    p3 = tf.add_paragraph(); p3.space_before = Pt(4); r3 = p3.add_run(); r3.text = "18 Steps Completed | 29/29 Pytest Passed (6.07s) | Thank You"; r3.font.name = 'Calibri'; r3.font.size = Pt(15); r3.font.bold = True; r3.font.color.rgb = C_PINK

    s21.notes_slide.notes_text_frame.text = (
        "We are now ready for the live interactive demonstration. I will switch to our running Streamlit dashboard to show real-time "
        "filtering, responsive Plotly charts, and data export. Thank you mentors and examiners, and I welcome any questions."
    )

    # Save to all target paths
    prs.save(str(PPTX_MASTER))
    prs.save(str(PPTX_ALT1))
    prs.save(str(PPTX_ALT2))
    print(f"Generated 21-Slide Master Presentation: {PPTX_MASTER}")

def compile_ppt_to_pdf():
    ppt_path = os.path.abspath(PPTX_MASTER)
    pdf_path = os.path.abspath(PDF_OUT)
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(ppt_path, WithWindow=False)
        pres.SaveAs(pdf_path, 32) # 32 = ppSaveAsPDF
        pres.Close()
        ppt.Quit()
        print(f"Compiled PPT to PDF successfully: {pdf_path} (Size: {os.path.getsize(pdf_path)} bytes)")
    except Exception as e:
        print(f"Failed to compile PPT to PDF: {e}")

if __name__ == "__main__":
    build_presentation()
    compile_ppt_to_pdf()
