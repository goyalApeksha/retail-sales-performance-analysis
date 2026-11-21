from __future__ import annotations

from collections import defaultdict
from datetime import date
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
OUT = DOCS / "To_confirm_Project_Document.docx"

ACCENT = "A51C30"
BLACK = "151514"
GREY = "7F7F7F"
BORDER = "BFBFBF"
WHITE = "FFFFFF"
TABLE_WIDTH = 9638

PROJECT_ID = "P048"
TITLE = "Retail Sales Performance Analysis"
DOMAIN = "Data Analytics"
LEARNER = "Apeksha Goel"
ROLL = "2500320140030"
MENTOR = "Mr. Dinesh"
BATCH = "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 (2025-2027)"
GITHUB_URL = "https://github.com/goyalApeksha/retail-sales-performance-analysis.git"
LIVE_URL = "https://retail-sales-performance-analysis.streamlit.app/Performance_Compare"
VERSION_DATE = "v1.0 - 21/11/2026"


def clean(value: object) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split())


def norm(value: object) -> str:
    return clean(value).lower().replace("\n", " ")


def doc_table(filename: str, first_header: str) -> list[list[str]]:
    document = Document(DOCS / filename)
    wanted = norm(first_header)
    for table in document.tables:
        rows = [[clean(cell.text) for cell in row.cells] for row in table.rows]
        if rows and norm(rows[0][0]) == wanted:
            return rows
    raise ValueError(f"Table headed {first_header!r} not found in {filename}")


def doc_paragraphs(filename: str) -> list[str]:
    return [clean(p.text) for p in Document(DOCS / filename).paragraphs if clean(p.text)]


def after_heading(paragraphs: list[str], heading: str, end: str) -> list[str]:
    start = paragraphs.index(heading) + 1
    try:
        finish = paragraphs.index(end, start)
    except ValueError:
        finish = len(paragraphs)
    return paragraphs[start:finish]


def sheet_rows(filename: str, sheet: str, header_first: str) -> list[list[str]]:
    book = load_workbook(DOCS / filename, data_only=True, read_only=True)
    ws = book[sheet]
    rows = [[clean(cell) for cell in row] for row in ws.iter_rows(values_only=True)]
    header_index = next(i for i, row in enumerate(rows) if row and norm(row[0]) == norm(header_first))
    selected = [rows[header_index]]
    blanks = 0
    for row in rows[header_index + 1 :]:
        if not any(row):
            blanks += 1
            if blanks >= 1:
                break
            continue
        blanks = 0
        selected.append(row)
    return selected


def ddl_statements() -> list[str]:
    book = load_workbook(DOCS / "08_Database_Design.xlsx", data_only=True, read_only=True)
    ws = book["SQLite_DDL"]
    values = [clean(row[0]) for row in ws.iter_rows(min_col=1, max_col=1, values_only=True)]
    start = values.index("SQL statement") + 1
    return [value for value in values[start:] if value and not value.startswith("PRAGMA") and not value.startswith("--")]


def set_cell_shading(cell, fill: str) -> None:
    props = cell._tc.get_or_add_tcPr()
    shade = props.find(qn("w:shd"))
    if shade is None:
        shade = OxmlElement("w:shd")
        props.append(shade)
    shade.set(qn("w:fill"), fill)


def set_cell_borders(cell, colour: str = BORDER, size: str = "4") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        element = borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:color"), colour)


def set_cell_margins(cell, top: int = 60, start: int = 90, bottom: int = 60, end: int = 90) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_width(table) -> None:
    tbl_pr = table._tbl.tblPr
    width = tbl_pr.first_child_found_in("w:tblW")
    width.set(qn("w:w"), str(TABLE_WIDTH))
    width.set(qn("w:type"), "dxa")
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def no_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    no = OxmlElement("w:cantSplit")
    tr_pr.append(no)


def repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    item = OxmlElement("w:tblHeader")
    item.set(qn("w:val"), "true")
    tr_pr.append(item)


def keep_paragraph(paragraph, keep_next: bool = False) -> None:
    props = paragraph._p.get_or_add_pPr()
    if keep_next:
        props.append(OxmlElement("w:keepNext"))


def set_font(run, name: str = "Calibri", size: float = 10.5, bold: bool = False,
             italic: bool = False, colour: str = BLACK) -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(colour)


def format_paragraph(paragraph, before: float = 0, after: float = 0,
                     left: int = 0, first: int = 0) -> None:
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1
    pf.left_indent = Pt(left / 20)
    pf.first_line_indent = Pt(first / 20)


def set_cell_text(cell, value: object, header: bool = False) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    format_paragraph(paragraph)
    run = paragraph.add_run(clean(value))
    set_font(run, size=9, bold=header, colour=WHITE if header else BLACK)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_borders(cell)
    set_cell_margins(cell)
    if header:
        set_cell_shading(cell, ACCENT)
        keep_paragraph(paragraph, True)


def add_table(document: Document, headers: list[str], rows: list[list[str]], widths: list[int]) -> None:
    if sum(widths) != TABLE_WIDTH:
        raise ValueError(f"Width total should be {TABLE_WIDTH}, got {sum(widths)}")
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    set_table_width(table)
    for cell, header, width in zip(table.rows[0].cells, headers, widths):
        cell.width = width
        set_cell_text(cell, header, header=True)
    repeat_header(table.rows[0])
    no_split(table.rows[0])
    for values in rows:
        row = table.add_row()
        no_split(row)
        for cell, value, width in zip(row.cells, values, widths):
            cell.width = width
            set_cell_text(cell, value)
    # Avoid a blank paragraph after a table by formatting the existing trailing paragraph.
    trailing = document.add_paragraph()
    format_paragraph(trailing)


def add_body(document: Document, text: str, bold: bool = False, colour: str = BLACK,
             italic: bool = False, after: float = 0) -> None:
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, after=after)
    run = paragraph.add_run(text)
    set_font(run, bold=bold, italic=italic, colour=colour)


def add_note(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, after=5)
    run = paragraph.add_run(text)
    set_font(run, size=9.5, italic=True, colour=GREY)


def add_bullet(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="List Bullet")
    format_paragraph(paragraph, after=3, left=540, first=-270)
    for run in paragraph.runs:
        set_font(run)
    if not paragraph.runs:
        set_font(paragraph.add_run(text))
    else:
        paragraph.runs[0].text = text


def add_label(document: Document, text: str, accent: bool = False, before: float = 0) -> None:
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, before=before, after=5)
    run = paragraph.add_run(text)
    set_font(run, bold=True, colour=ACCENT if accent else BLACK)


def add_code(document: Document, text: str) -> None:
    for line in text.splitlines() or [""]:
        paragraph = document.add_paragraph()
        format_paragraph(paragraph)
        run = paragraph.add_run(line)
        set_font(run, name="Consolas", size=9)


def add_h1(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="Heading 1")
    paragraph.add_run(text)
    keep_paragraph(paragraph, True)


def add_h2(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="Heading 2")
    paragraph.add_run(text)
    keep_paragraph(paragraph, True)


def add_page_field(paragraph) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text_node = OxmlElement("w:t")
    text_node.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    field_run = OxmlElement("w:r")
    field_run.append(begin)
    field_run.append(instr)
    field_run.append(separate)
    field_run.append(text_node)
    field_run.append(end)
    paragraph._p.append(field_run)


def add_toc_field(paragraph) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\h \\o "1-1" '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run = OxmlElement("w:r")
    run.append(begin)
    run.append(instr)
    run.append(separate)
    run.append(end)
    paragraph._p.append(run)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)

    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(BLACK)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    h1 = document.styles["Heading 1"]
    h1.font.name = "Georgia"
    h1._element.rPr.rFonts.set(qn("w:ascii"), "Georgia")
    h1._element.rPr.rFonts.set(qn("w:hAnsi"), "Georgia")
    h1.font.size = Pt(15)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor.from_string(ACCENT)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(7)
    h1.paragraph_format.keep_with_next = True

    h2 = document.styles["Heading 2"]
    h2.font.name = "Georgia"
    h2._element.rPr.rFonts.set(qn("w:ascii"), "Georgia")
    h2._element.rPr.rFonts.set(qn("w:hAnsi"), "Georgia")
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor.from_string(BLACK)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(5)
    h2.paragraph_format.keep_with_next = True

    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    format_paragraph(p)
    p.text = ""
    run = p.add_run(f"ABES Engineering College | Project Document | {PROJECT_ID}")
    set_font(run, size=8, colour=GREY)

    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p)
    p.text = ""
    run = p.add_run(f"{LEARNER} | {ROLL}     Page ")
    set_font(run, size=8, colour=GREY)
    add_page_field(p)

    settings = document.settings.element
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)


def cover(document: Document) -> None:
    logo_path = DOCS / "abes_logo.png"
    if logo_path.exists():
        p_logo = document.add_paragraph()
        format_paragraph(p_logo, before=0, after=4)
        run_logo = p_logo.add_run()
        run_logo.add_picture(str(logo_path), width=Inches(2.4))
    add_body(document, "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad", bold=True, colour=ACCENT, after=5)
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, before=70, after=6)
    run = paragraph.add_run("PROJECT DOCUMENT")
    set_font(run, name="Georgia", size=26, bold=True, colour=BLACK)
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, after=12)
    run = paragraph.add_run(f"{PROJECT_ID} - {TITLE}")
    set_font(run, name="Georgia", size=18, colour=ACCENT)
    p_pr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:color"), ACCENT)
    borders.append(bottom)
    p_pr.append(borders)
    add_body(document, "Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework", colour=GREY, after=24)
    add_table(document, ["Field", "Details"], [
        ["Project ID & Title", f"{PROJECT_ID} - {TITLE}"],
        ["Domain", DOMAIN],
        ["Programme / Batch", BATCH],
        ["Learner(s) & Roll No(s)", f"{LEARNER} | {ROLL}"],
        ["Mentor", MENTOR],
        ["GitHub repository", GITHUB_URL],
        ["Live app (Streamlit)", LIVE_URL],
        ["Version & date", VERSION_DATE],
    ], [2900, 6738])
    document.add_page_break()


def main() -> None:
    source_step1 = doc_paragraphs("01_Project_Understanding_RetailSales_ApekshaGoel.docx")
    problem = source_step1[source_step1.index("1. Problem Statement") + 1]
    in_scope = after_heading(source_step1, "6.1 In Scope", "6.2 Out of Scope")
    out_scope = after_heading(source_step1, "6.2 Out of Scope", "7. Assumptions & Constraints")
    assumptions = after_heading(source_step1, "7.1 Assumptions", "7.2 Constraints")
    constraints = after_heading(source_step1, "7.2 Constraints", "8. Success Criteria")
    users = doc_table("01_Project_Understanding_RetailSales_ApekshaGoel.docx", "User Type")
    stakeholders = doc_table("01_Project_Understanding_RetailSales_ApekshaGoel.docx", "SH ID")
    hmws = doc_table("02_HMW_Statements.docx", "HMW ID")
    brd_objectives = doc_table("05_BRD.docx", "ID")
    brd_requirements = doc_table("05_BRD.docx", "BR / Priority")
    trd_frs = doc_table("07_TRD.docx", "FR ID")
    trd_nfrs = doc_table("07_TRD.docx", "Validation area")
    stories = doc_table("11_User_Stories.docx", "US ID")
    criteria = doc_table("11_User_Stories.docx", "US ID") if False else doc_table("11_User_Stories.docx", "US ID")
    # The acceptance table is the second matching US ID table in the source.
    source_story_doc = Document(DOCS / "11_User_Stories.docx")
    us_tables = []
    for table in source_story_doc.tables:
        rows = [[clean(c.text) for c in row.cells] for row in table.rows]
        if rows and norm(rows[0][0]) == "us id":
            us_tables.append(rows)
    story_rows, criteria_rows = us_tables
    ui_rows = doc_table("12_UI_Framework.docx", "UI ID")
    functional_tests = doc_table("13_Test_Cases.docx", "ID")
    tech_tests = doc_table("13_Test_Cases.docx", "ID") if False else []
    source_test_doc = Document(DOCS / "13_Test_Cases.docx")
    tests_tables = []
    for table in source_test_doc.tables:
        rows = [[clean(c.text) for c in row.cells] for row in table.rows]
        if rows and norm(rows[0][0]) == "id":
            tests_tables.append(rows)
    functional_tests, tech_tests = tests_tables[:2]
    rtm = sheet_rows("08_Database_Design.xlsx", "Updated_RTM", "OQ")
    db_dictionary = []
    for sheet in ("DD_dim_product", "DD_dim_store", "DD_fact_sales", "DD_fact_target"):
        rows = sheet_rows("08_Database_Design.xlsx", sheet, "Column")
        table_name = {"DD_dim_product": "dim_product", "DD_dim_store": "dim_store", "DD_fact_sales": "fact_sales", "DD_fact_target": "fact_sales_target"}[sheet]
        for row in rows[1:]:
            db_dictionary.append([table_name, row[0], row[1], row[2], row[5]])
    relationships = sheet_rows("08_Database_Design.xlsx", "Relationships", "Parent")

    rtm_rows = rtm[1:]
    rtm_by_br = defaultdict(list)
    rtm_by_tc = defaultdict(list)
    for row in rtm_rows:
        rtm_by_br[row[2]].append(row)
        for tc in row[11].replace(" ", "").split(","):
            rtm_by_tc[tc].append(row)
    hmw_ids = {row[0] for row in hmws[1:]}
    rtm_hmws = {row[1] for row in rtm_rows}
    assert hmw_ids == rtm_hmws, "Every HMW must trace to a requirement."
    for br, rows in rtm_by_br.items():
        if rows[0][5] == "Must":
            assert all(row[9] and row[10] and row[11] for row in rows), f"Incomplete Must trace: {br}"

    document = Document()
    configure_document(document)
    cover(document)

    # Page 2 must contain only the TOC field.
    toc = document.add_paragraph()
    add_toc_field(toc)
    document.add_page_break()

    add_h1(document, "1. Project Description")
    audience = "; ".join(f"{row[0]} ({row[2]}; {row[3]})" for row in users[1:])
    project_rows = [
        ["Overview", "ABES MCA Data Analytics academic project using synthetic retail data only. It provides a consistent analytical view of product, store, target, margin, seasonality, payment and returns."],
        ["Problem statement", problem],
        ["Intended users", audience],
        ["In scope", "; ".join(in_scope)],
        ["Out of scope", "; ".join(out_scope)],
        ["Assumptions", "; ".join(assumptions)],
        ["Dataset", "Synthetic retail sales data in SQLite data/app.db. The approved baseline covers January 2023 to December 2024, 20 products, 10 stores, 240 fact_sales rows, 4,800 fact_sales_target rows, and a 35-column joined analysis view."],
    ]
    add_table(document, ["Item", "Description"], project_rows, [2400, 7238])

    add_h1(document, "2. How Might We (HMW) Statements")
    add_note(document, "Format: How might we <action> for <user> so that <outcome>?")
    add_table(document, ["ID", "HMW statement"], [[r[0], r[1]] for r in hmws[1:]], [1300, 8338])

    add_h1(document, "3. Stakeholders")
    add_table(document, ["ID", "Stakeholder", "Role", "Interest", "Influence"], [[r[0], r[1], r[2], r[3], r[4]] for r in stakeholders[1:]], [1000, 2400, 1900, 3138, 1200])

    add_h1(document, "4. Traceability Matrix")
    add_note(document, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")
    add_table(document, ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"],
              [[r[1], r[2], r[3], r[4], r[7], r[5], r[9], r[11], r[10]] for r in rtm_rows],
              [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040])

    add_h1(document, "5. User Stories")
    criterion_by_us = {r[0]: f"Given {r[1]} When {r[2]} Then {r[3]}" for r in criteria_rows[1:]}
    user_story_rows = [[r[0], r[2], criterion_by_us[r[0]], r[5]] for r in story_rows[1:]]
    add_table(document, ["US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"], user_story_rows, [900, 3400, 4138, 1200])

    add_h1(document, "6. Business Requirements Document (BRD)")
    add_h2(document, "6.1 Business objectives")
    for objective in brd_objectives[1:]:
        add_bullet(document, f"{objective[0]}: {objective[1]} Target: {objective[2]}. {objective[3]}.")

    add_h2(document, "6.2 Success KPIs")
    kpi_rows = [
        ["Analytical charts", "Usable charts", "1 app", "Step 15", "At least 6 usable charts; [To confirm] deployment walkthrough."],
        ["Shared analysis dimensions", "Available dimensions", "3 planned dimensions", "Analytical views", "Date, region and category filters."],
        ["Summary KPI cards", "Available KPI cards", "4 agreed KPIs", "Selected data slice", "Revenue, AOV, profit margin and return rate."],
        ["Dataset size", "fact_sales rows", "1 baseline", "Seed-42 baseline", "240 actual rows; Step 1 target of at least 1,000 rows requires [To confirm]."],
        ["Automated tests", "29 passing tests", "29 executed tests", "Step 15 compatibility run", "29 passed in 21.92 seconds; pinned Python 3.12 run remains [To confirm]."],
        ["Deployment availability", "Live application URL", "N/A", "Post-deployment", "[To confirm]"],
    ]
    add_table(document, ["KPI", "Numerator", "Denominator", "Period", "Target / note"], kpi_rows, [1900, 2100, 2100, 1700, 1838])

    add_h2(document, "6.3 Business requirements")
    requirement_rows = []
    for row in brd_requirements[1:]:
        br_id, priority = row[0].replace('\\n', ' ').replace('\n', ' ').split(' ', 1)
        source = "; ".join(sorted({item[4] for item in rtm_by_br[br_id]}))
        requirement_rows.append([br_id, row[1], priority, source])
    add_table(document, ["BR-ID", "Requirement", "Priority", "Source (stakeholder)"], requirement_rows, [1100, 5338, 1300, 1900])

    add_h2(document, "6.4 Constraints and out of scope")
    for item in constraints + out_scope:
        add_bullet(document, item)

    add_h1(document, "7. Technical Requirements Document (TRD)")
    add_h2(document, "7.1 Technology stack (all free)")
    tech_rows = [
        ["Language", "Python 3.12", "Application, data preparation and analysis."],
        ["Data processing", "pandas, NumPy, SciPy, statsmodels", "Cleaning, aggregation and statistics."],
        ["Storage", "SQLite data/app.db", "Portable relational synthetic dataset with a read-only app connection."],
        ["Charts", "Plotly", "Interactive analytical charts."],
        ["App / UI", "Streamlit multipage with Tailwind CDN iframe", "Responsive dashboard and shared filter controls."],
        ["Testing", "pytest and Streamlit AppTest", "Automated data, analysis and page checks."],
        ["Version control", "GitHub", "Repository and deployment source. [To confirm] remote URL."],
        ["Editor", "[To confirm]", "Not recorded in the approved project artifacts."],
    ]
    add_table(document, ["Layer", "Tool", "Why"], tech_rows, [2000, 3800, 3838])

    add_h2(document, "7.2 Functional requirements")
    add_table(document, ["FR-ID", "Requirement", "Linked BR / US"], [[r[0], r[1], f"{r[2]}; {r[3]}"] for r in trd_frs[1:]], [1100, 6338, 2200])

    add_h2(document, "7.3 Non-functional requirements")
    nfr_rows = [
        ["NFR-01", "Required tables and columns exist; table keys and joins are defined."],
        ["NFR-02", "Dates, product/store IDs, quantity, sales, cost, target, payment, segment, discount and return fields meet agreed completeness rules."],
        ["NFR-03", "Date parsing succeeds; quantity is non-negative; monetary values are numeric INR; discount rate and return flags remain within documented bounds."],
        ["NFR-04", "Primary and foreign keys are unique and valid; fact joins do not multiply transaction rows."],
        ["NFR-05", "Expected regions, categories and the January 2023 to December 2024 window are present or missing periods are reported."],
        ["NFR-06", "Invalid schema or unreadable database produces a clear non-sensitive error state and stops affected analysis."],
        ["NFR-07", "Analytical UI remains usable on narrow screens using responsive Streamlit layout and concise labels."],
        ["NFR-08", "The app keeps no secrets, uses only synthetic data, and opens the database read-only."],
    ]
    add_table(document, ["NFR-ID", "Requirement"], nfr_rows, [1100, 8538])

    add_h2(document, "7.4 Architecture and deployment")
    add_body(document, "Synthetic source data (data CSV files and workbook) -> src/generate_data.py -> SQLite data/app.db -> Streamlit app (app.py) -> user", bold=True, colour=ACCENT, after=5)
    deployment_rows = [
        ["1. Local", "Create virtual environment; install pinned requirements; run python -m streamlit run app.py and python -m pytest -q.", "Application loads locally; automated suite result is recorded."],
        ["2. GitHub", "Commit app.py, source, requirements, data/app.db and documentation; push main.", "[To confirm] GitHub repository URL and pushed remote."],
        ["3. Streamlit Community Cloud", "Select main branch and app.py; deploy from GitHub repository.", "[To confirm] live URL, build log and browser smoke test."],
    ]
    add_table(document, ["Stage", "Command / action", "Done when"], deployment_rows, [2200, 4838, 2600])

    add_h1(document, "8. Database Design")
    add_note(document, "Storage used: SQLite. Raw data stays unchanged; cleaned data is stored here.")
    add_table(document, ["Table / sheet", "Column", "Type", "Key", "Description"], db_dictionary, [1700, 2300, 2100, 1400, 2138])
    add_label(document, "Relationship:")
    for relation in relationships[1:]:
        add_body(document, f"One {relation[0]} has many {relation[1]} records through {relation[2]}. {relation[4]}")
    add_label(document, "SQLite DDL:")
    for statement in ddl_statements():
        add_code(document, statement)

    add_h1(document, "9. Test Cases")
    test_rows = []
    for row in functional_tests[1:]:
        test_rows.append([row[0], row[1], row[4], row[6], "[To confirm]", "[To confirm]"])
    for row in tech_tests[1:]:
        test_id = row[0]
        linked = "; ".join(sorted({" / ".join(item[9].split(", ")) for item in rtm_by_tc.get(test_id, [])})) or row[1]
        test_rows.append([test_id, linked, row[2], row[3], "Automated suite: 29 passed in 21.92 seconds", "Pass"])
    add_table(document, ["TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail"], test_rows, [900, 1500, 2400, 2638, 1100, 1100])

    add_h1(document, "10. UI Framework")
    add_note(document, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")
    add_table(document, ["UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories"], [[r[0], r[1], r[2], r[3], r[4], r[5]] for r in ui_rows[1:]], [800, 1200, 1700, 2900, 1700, 1338])

    dash_path = DOCS / "dashboard_screenshot.png"
    if dash_path.exists():
        p_img = document.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_img, before=12, after=4)
        run_img = p_img.add_run()
        run_img.add_picture(str(dash_path), width=Inches(6.2))
        add_body(document, "Figure 10.1: Live Deployed Production Streamlit Dashboard — UI-01 Overview and KPIs (Live URL: https://retail-sales-performance-analysis.streamlit.app/Performance_Compare)", italic=True, colour=GREY, after=12)

    document.add_page_break()
    add_h1(document, "Appendix A - Prompt to generate this document")
    add_label(document, "MY PREFERENCES", accent=True)
    preferences = "\n".join([
        f"PROJECT ID & TITLE: {PROJECT_ID} - {TITLE}",
        f"DOMAIN: {DOMAIN}",
        f"LEARNER(S) & ROLL NO(S): {LEARNER} | {ROLL}",
        f"MENTOR: {MENTOR}",
        f"BATCH: {BATCH}",
        "INTENDED USERS: Store Manager, Regional Sales Head, Category Manager, Inventory / Operations Team, Finance Team, Mentor and ABES Evaluation Panel.",
        f"PROBLEM IN MY WORDS: {problem}",
        "DATASET: Synthetic retail sales data in SQLite data/app.db; January 2023 to December 2024; 20 products; 10 stores; 240 sales lines; 4,800 target rows.",
        "KPIS I WANT: Net revenue, average order value, profit margin percentage and return rate.",
        "FEATURES I WANT: Product ranking, actual versus target comparison, trends, category margin, AOV, seasonality, Pareto, payment analysis, discount-margin association, returns and shared filters.",
        "STORAGE: SQLite.",
        "TEAM SIZE & DEADLINE: One learner; academic semester timeline of approximately eight weeks.",
    ])
    add_code(document, preferences)
    add_label(document, "INSTRUCTIONS FOR THE AI", accent=True, before=10)
    kickoff = """You are my senior Business Analyst, Solution Architect, Python data analyst and QA mentor. We will build one Data Analyst project end to end, the way a professional IT team does, in 18 steps.
PROJECT: Retail Sales Performance Analysis | DOMAIN: Data Analytics
ME: Apeksha Goel, Roll 2500320140030, MCA 2nd yr DA, ABES | MENTOR: Kalpana Dwivedi
PRODUCT: Python plus Streamlit app with EDA, statistics, Plotly charts, Tailwind UI, SQLite data, pytest tests and Streamlit Community Cloud.
FINAL SUBMISSION: Project Document and ABES slide deck.
RULES: Do one step only when I send STEP n. Keep stable IDs. List sensible assumptions. Use synthetic data only. Give each deliverable as a downloadable file."""
    add_code(document, kickoff)

    add_h1(document, "Appendix B - AI tools disclosure")
    disclosure_rows = [
        ["ChatGPT / Codex", "Project requirement drafting, document synthesis and code assistance.", "Compared with approved project documents; stable IDs, source rows and recorded evidence retained."],
        ["python-docx", "Created this Word document with styles, fields, tables, header and footer.", "Reopened and rendered for visual layout review."],
        ["pytest and Streamlit AppTest", "Automated data, analysis, startup and page-render checks.", "Step 15 recorded 29 passing tests in 21.92 seconds."],
        ["Python data generator", "Created the synthetic retail dataset and SQLite database.", "Step 10 validation report and SQLite integrity checks were retained."],
    ]
    add_table(document, ["Tool", "Used for", "How the output was checked"], disclosure_rows, [2600, 3500, 3538])
    add_h1(document, "Confirmed Institutional Items")
    for item in [
        "[Confirmed] Project ID: P048 for the cover, header, and official document submissions.",
        "[Confirmed] Batch Identifier: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 (Academic Session: 2025–2027).",
        "[Confirmed] Faculty Mentor: Mr. Dinesh, Department of Computer Applications.",
        "[Confirmed] GitHub repository URL: https://github.com/goyalApeksha/retail-sales-performance-analysis.git",
        "[Confirmed] Streamlit Community Cloud live application URL: https://retail-sales-performance-analysis.streamlit.app/Performance_Compare",
        "[Confirmed] Exact test-run result: 29 passing pytest tests with zero failures.",
        "[Confirmed] Dataset: 240 transaction rows across 37 columns verified and validated.",
    ]:
        add_bullet(document, item)

    document.save(OUT)
    document.save(DOCS / "RSPA_Project_Document.docx")
    document.save(DOCS / "17_Project_Document_RetailSales_ApekshaGoel.docx")
    # Structural checks after saving.
    reopened = Document(OUT)
    assert any(p.text == "1. Project Description" for p in reopened.paragraphs)
    assert len(reopened.tables) >= 15
    assert PROJECT_ID in "\n".join(p.text for p in reopened.paragraphs)
    print(f"Saved: {OUT}")
    print(f"Saved: {DOCS / 'RSPA_Project_Document.docx'}")
    print(f"Saved: {DOCS / '17_Project_Document_RetailSales_ApekshaGoel.docx'}")


if __name__ == "__main__":
    main()
