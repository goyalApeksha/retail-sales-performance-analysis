"""
Merge all 18-step project deliverables into 1 unified master PDF.
Includes:
- Part I: 18-Step Master Project Dossier & Executive Summary (11 pages)
- Part II: Official ABES Capstone Project Document (30 pages)
- Part III: Official Institutional 11-Slide Presentation Deck (11 pages)
With hierarchical bookmarks / outline for effortless navigation.
"""

from pathlib import Path
import pypdf

DOCS_DIR = Path(r"c:\Users\apeks\Downloads\retail sales performance analysis\docs")

PDF_DOSSIER = DOCS_DIR / "18_Steps_Complete_Project_Document.pdf"
PDF_PROJ_DOC = DOCS_DIR / "17_Project_Document_RetailSales_ApekshaGoel.pdf"
PDF_PRESENTATION = DOCS_DIR / "18_Presentation_RetailSales_ApekshaGoel.pdf"

PDF_COMBINED_1 = DOCS_DIR / "Retail_Sales_Performance_Analysis_18_Steps_Complete.pdf"
PDF_COMBINED_2 = DOCS_DIR / "18_Steps_Complete_Project_Document_Combined.pdf"
PDF_COMBINED_3 = DOCS_DIR / "Retail_Sales_Performance_Analysis_One_PDF.pdf"

def merge_pdfs():
    writer = pypdf.PdfWriter()

    # Part I: 18-Step Master Project Dossier
    r_dossier = pypdf.PdfReader(str(PDF_DOSSIER))
    p1_start = len(writer.pages)
    writer.append(r_dossier)
    b_part1 = writer.add_outline_item("Part I: 18-Step Master Project Dossier", p1_start)
    writer.add_outline_item("Cover Page & Project Metadata", p1_start, parent=b_part1)
    writer.add_outline_item("Executive Summary & 18-Step Roadmap Table", p1_start + 1, parent=b_part1)
    writer.add_outline_item("Step 1: Project Understanding & Opportunity Questions (OQ-01..10)", p1_start + 2, parent=b_part1)
    writer.add_outline_item("Step 2: 'How Might We' Statements (HMW-01..12)", p1_start + 2, parent=b_part1)
    writer.add_outline_item("Step 3: Needs Analysis & MoSCoW Prioritization", p1_start + 3, parent=b_part1)
    writer.add_outline_item("Step 4: Requirements Traceability Matrix (RTM)", p1_start + 3, parent=b_part1)
    writer.add_outline_item("Steps 5 & 6: Business & Product Requirements (BRD/PRD)", p1_start + 4, parent=b_part1)
    writer.add_outline_item("Step 7: Technical Requirements Document (TRD)", p1_start + 4, parent=b_part1)
    writer.add_outline_item("Steps 8 & 9: Database Architecture & Relational ERD", p1_start + 5, parent=b_part1)
    writer.add_outline_item("Step 10: Synthetic Data Generation & Data Hygiene", p1_start + 5, parent=b_part1)
    writer.add_outline_item("Step 11: User Stories & Agile Sprint Planning", p1_start + 6, parent=b_part1)
    writer.add_outline_item("Step 12: UI Framework & Screen Wireframes", p1_start + 6, parent=b_part1)
    writer.add_outline_item("Step 13: Test Case Specification (Functional & Technical)", p1_start + 7, parent=b_part1)
    writer.add_outline_item("Steps 14 & 15: AI Build Blueprint & Pytest 29/29 Results", p1_start + 7, parent=b_part1)
    writer.add_outline_item("Step 16: Streamlit Community Cloud Deployment", p1_start + 8, parent=b_part1)
    writer.add_outline_item("Step 17: Official ABES Project Document Synthesis", p1_start + 8, parent=b_part1)
    writer.add_outline_item("Step 18: Presentation Deck & Viva Voce Guide (5 Q&As)", p1_start + 9, parent=b_part1)
    writer.add_outline_item("Appendix: Verification Audit & Complete Artifact Inventory", p1_start + 10, parent=b_part1)

    # Part II: Official ABES Capstone Project Document
    r_proj = pypdf.PdfReader(str(PDF_PROJ_DOC))
    p2_start = len(writer.pages)
    writer.append(r_proj)
    b_part2 = writer.add_outline_item("Part II: Official ABES Institutional Project Document (30 Pages)", p2_start)
    writer.add_outline_item("Cover Page & Institutional Details", p2_start, parent=b_part2)
    writer.add_outline_item("Table of Contents", p2_start + 1, parent=b_part2)
    writer.add_outline_item("1. Project Description", p2_start + 2, parent=b_part2)
    writer.add_outline_item("2. How Might We (HMW) Statements", p2_start + 3, parent=b_part2)
    writer.add_outline_item("3. Stakeholders", p2_start + 4, parent=b_part2)
    writer.add_outline_item("4. Traceability Matrix", p2_start + 4, parent=b_part2)
    writer.add_outline_item("5. User Stories", p2_start + 7, parent=b_part2)
    writer.add_outline_item("6. Business Requirements Document (BRD)", p2_start + 11, parent=b_part2)
    writer.add_outline_item("7. Technical Requirements Document (TRD)", p2_start + 16, parent=b_part2)
    writer.add_outline_item("8. Database Design", p2_start + 20, parent=b_part2)
    writer.add_outline_item("9. Test Cases", p2_start + 23, parent=b_part2)
    writer.add_outline_item("10. UI Framework", p2_start + 27, parent=b_part2)

    # Part III: Institutional Presentation Deck
    r_pres = pypdf.PdfReader(str(PDF_PRESENTATION))
    p3_start = len(writer.pages)
    writer.append(r_pres)
    b_part3 = writer.add_outline_item("Part III: Institutional Presentation Deck (11 Slides)", p3_start)
    slide_titles = [
        "Slide 1: Title & Program Header (Dark)",
        "Slide 2: Problem and How Might We",
        "Slide 3: Stakeholders and Users",
        "Slide 4: Requirements and Traceability",
        "Slide 5: Data and Preparation",
        "Slide 6: Solution and Technology",
        "Slide 7: Dashboard Overview & UI Screens",
        "Slide 8: Key Findings (88.9%, -0.52, +42.0%)",
        "Slide 9: Testing and Validation (29/29 Passed)",
        "Slide 10: Recommendation, Limitations & Next Steps",
        "Slide 11: Live Demo and Q&A (Dark)",
    ]
    for s_idx, s_title in enumerate(slide_titles):
        writer.add_outline_item(s_title, p3_start + s_idx, parent=b_part3)

    # Save to all combined output targets
    for out_path in [PDF_COMBINED_1, PDF_COMBINED_2, PDF_COMBINED_3]:
        with open(out_path, "wb") as f_out:
            writer.write(f_out)
        print(f"Saved: {out_path} ({len(writer.pages)} pages, {out_path.stat().st_size} bytes)")

if __name__ == "__main__":
    merge_pdfs()
