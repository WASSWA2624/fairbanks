#!/usr/bin/env python3
"""Generate FairBanks SF-424 V4.0 fill answers PDF for CDC-RFA-JG-26-0056.

Source blank form:
  applications/cdc-ghs-global/documents to fill/WS01740564-SF424_4_0-V4.0.pdf

Output:
  applications/cdc-ghs-global/documents to fill/WS01740564-SF424_4_0-V4.0_ANSWERS.pdf
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "documents to fill"
OUT = OUT_DIR / "WS01740564-SF424_4_0-V4.0_ANSWERS.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
ORANGE = colors.HexColor("#C45C26")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
CREAM = colors.HexColor("#F7F5F0")
LINE = colors.HexColor("#CED9D8")
WHITE = colors.white
WARN = colors.HexColor("#FFF4E5")

ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
CONTACT = "Racheal Nabukeera"
TITLE = "Managing Director and Co-founder"
EMAIL = "info@fairbanksmedicalcentre.org"
PHONE = "+256 772 849 258"
FAX = "N/A"
STREET1 = "Plot 1423 and 1425 Tirupati Road"
STREET2 = "Fairbanks Medical Centre, Kololo IV, Lugogo"
CITY = "Kampala"
COUNTY = "Kampala Central Division"
PROVINCE = "Kampala"
COUNTRY = "UGANDA"
POSTAL = "CONFIRM Uganda postal code (or leave blank if none)"
TIN_UG = "1053370026"
EIN_FOREIGN = "44-4444444"
COMPANY_NO = "80020003843337"
NSSF = "NS043295"
NOFO = "CDC-RFA-JG-26-0056"
NOFO_TITLE = (
    "Strengthening Global Health Security by improving public health capacity "
    "to detect, notify, and respond to disease outbreaks globally"
)
CFDA = "93.318"
CFDA_TITLE = (
    "Protecting and Improving Health Globally: Building and Strengthening "
    "Public Health Impact, Systems, Capacity, and Security"
)
AGENCY = "Centers for Disease Control and Prevention (CDC)"
AGENCY_DETAIL = "CDC Global Health Center / Division of Global Health Protection"
START = "09/30/2026"
END = "09/29/2031"
FED_ASK = "$7,500,000"
C1_ASK = "$3,000,000"
C2_ASK = "$1,000,000"
C3_ASK = "$1,500,000"
C4_ASK = "$1,000,000"
C5_ASK = "$1,000,000"
BOX15 = (
    "FairBanks strengthens last-mile surveillance across the region for faster "
    "outbreak detection and response, using CHWs/VHTs, FCHIP, 7-1-7 timing, "
    "surge readiness and contingency capacity."
)


def styles():
    base = getSampleStyleSheet()
    s = {
        "cover_kicker": ParagraphStyle(
            "cover_kicker",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9,
            textColor=TEAL,
            spaceAfter=6,
            tracking=0.5,
        ),
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=NAVY,
            spaceAfter=8,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=SLATE,
            spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=12,
            textColor=NAVY,
            spaceBefore=10,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=SLATE,
            spaceAfter=4,
        ),
        "small": ParagraphStyle(
            "small",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=MUTED,
            spaceAfter=3,
        ),
        "cell_label": ParagraphStyle(
            "cell_label",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=NAVY,
        ),
        "cell_value": ParagraphStyle(
            "cell_value",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=SLATE,
        ),
        "note": ParagraphStyle(
            "note",
            parent=base["Normal"],
            fontName="Helvetica-Oblique",
            fontSize=8,
            leading=10,
            textColor=ORANGE,
        ),
        "footer": ParagraphStyle(
            "footer",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            textColor=MUTED,
            alignment=TA_CENTER,
        ),
    }
    return s


def p(text: str, style) -> Paragraph:
    return Paragraph(text.replace("\n", "<br/>"), style)


def section_table(rows, col_widths, sty):
    data = []
    for label, value, note in rows:
        left = p(label, sty["cell_label"])
        right_bits = [p(value, sty["cell_value"])]
        if note:
            right_bits.append(p(note, sty["note"]))
        data.append([left, right_bits])

    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), CREAM),
                ("BACKGROUND", (1, 0), (1, -1), WHITE),
                ("BOX", (0, 0), (-1, -1), 0.6, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return t


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(TEAL)
    canvas.setLineWidth(1.2)
    canvas.line(0.7 * inch, LETTER[1] - 0.55 * inch, LETTER[0] - 0.7 * inch, LETTER[1] - 0.55 * inch)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(
        0.7 * inch,
        LETTER[1] - 0.42 * inch,
        "FairBanks Medical Centre · SF-424 V4.0 answers · CDC-RFA-JG-26-0056",
    )
    canvas.drawRightString(
        LETTER[0] - 0.7 * inch,
        LETTER[1] - 0.42 * inch,
        f"Page {doc.page}",
    )
    canvas.setStrokeColor(LINE)
    canvas.line(0.7 * inch, 0.55 * inch, LETTER[0] - 0.7 * inch, 0.55 * inch)
    canvas.drawCentredString(
        LETTER[0] / 2,
        0.35 * inch,
        "Your health, our mission. · Companion worksheet — enter these values into Grants.gov / Adobe SF-424",
    )
    canvas.restoreState()


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    sty = styles()
    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=LETTER,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
        title="SF-424 V4.0 Answers — FairBanks — CDC-RFA-JG-26-0056",
        author="FAIRBANKS MEDICAL CENTRE LIMITED",
    )
    story = []
    w = LETTER[0] - 1.4 * inch
    cols = [1.55 * inch, w - 1.55 * inch]

    # Cover / how to use
    story.append(p("GRANTS.GOV FORM WORKSHEET", sty["cover_kicker"]))
    story.append(p("Application for Federal Assistance (SF-424) V4.0 — Fill Answers", sty["cover_title"]))
    story.append(
        p(
            f"<b>Applicant:</b> {ORG}<br/>"
            f"<b>Opportunity:</b> {NOFO}<br/>"
            f"<b>Form file:</b> WS01740564-SF424_4_0-V4.0.pdf<br/>"
            f"<b>Prepared:</b> {date.today().isoformat()} · Slogan: Your health, our mission.",
            sty["cover_sub"],
        )
    )
    story.append(
        p(
            "How to use: Open the blank Adobe / Grants.gov SF-424 form and copy each answer below "
            "into the matching box. Orange notes marked <b>CONFIRM</b> must be checked before submit "
            "(especially UEI/SAM, EIN entry for foreign orgs, postal code, and final Year 1 federal total).",
            sty["body"],
        )
    )
    story.append(
        p(
            "Legal note: This PDF is a fill companion. The official SF-424 must still be completed and "
            "certified in Grants.gov by the authorized official. Do not password-protect submission files.",
            sty["small"],
        )
    )

    story.append(p("1–7. Submission header", sty["h1"]))
    story.append(
        section_table(
            [
                (
                    "1. Type of Submission",
                    "Application",
                    "Select Application (not Preapplication / Changed).",
                ),
                (
                    "2. Type of Application",
                    "New",
                    "First submission under this UEI for this NOFO.",
                ),
                (
                    "3. Date Received",
                    "Leave blank",
                    "Grants.gov fills this when you submit.",
                ),
                (
                    "4. Applicant Identifier",
                    "Leave blank (optional internal ID)",
                    "Optional. Example if used: FB-CDC-JG-26-0056",
                ),
                (
                    "5a. Federal Entity Identifier",
                    "Leave blank",
                    "Not assigned yet.",
                ),
                (
                    "5b. Federal Award Identifier",
                    "Leave blank",
                    "New application — no prior award ID.",
                ),
                (
                    "6. Date Received by State",
                    "Leave blank",
                    "Not applicable for this foreign / non-EO-12372 path.",
                ),
                (
                    "7. State Application Identifier",
                    "Leave blank",
                    "Not applicable.",
                ),
            ],
            cols,
            sty,
        )
    )

    story.append(p("8. Applicant Information", sty["h1"]))
    story.append(
        section_table(
            [
                ("8a. Legal Name", ORG, "Must match SAM.gov legal name exactly."),
                (
                    "8b. EIN / TIN",
                    f"{EIN_FOREIGN}",
                    f"CONFIRM: Foreign orgs usually enter {EIN_FOREIGN}. "
                    f"Uganda TIN on file: {TIN_UG}. Do not invent a U.S. EIN.",
                ),
                (
                    "8c. UEI (SAM)",
                    "CONFIRM — enter active SAM.gov UEI",
                    "UEI for the physical Kampala location receiving funds. Register/renew now.",
                ),
                ("8d. Street 1", STREET1, "Required."),
                ("8d. Street 2", STREET2, "Optional but recommended."),
                ("8d. City", CITY, "Required."),
                ("8d. County/Parish", COUNTY, "If field allows."),
                ("8d. State", "N/A / leave blank", "U.S. state not applicable."),
                ("8d. Province", PROVINCE, "Use for Uganda."),
                ("8d. Country", COUNTRY, "Select Uganda in dropdown (UG / UGA)."),
                ("8d. ZIP / Postal Code", POSTAL, "CONFIRM before submit."),
                (
                    "8e. Department / Division",
                    "FairBanks Community Reach / FCHIP (Global Health Security)",
                    "Primary unit carrying out the award activities.",
                ),
                ("8f. Prefix", "Ms.", "CONFIRM preferred prefix."),
                ("8f. First Name", "Racheal", ""),
                ("8f. Middle Name", "Leave blank (or CONFIRM)", ""),
                ("8f. Last Name", "Nabukeera", ""),
                ("8f. Suffix", "Leave blank", ""),
                ("8f. Title", TITLE, "Person contacted on application matters."),
                ("8f. Organizational Affiliation", ORG, ""),
                ("8f. Telephone", PHONE, ""),
                ("8f. Fax", FAX, ""),
                ("8f. Email", EMAIL, "Use same email family as SAM EBiz POC if possible."),
            ],
            cols,
            sty,
        )
    )

    story.append(PageBreak())
    story.append(p("9–15. Agency, CFDA, opportunity, project", sty["h1"]))
    story.append(
        section_table(
            [
                (
                    "9. Type of Applicant (1)",
                    "Q – For-Profit Organization (Other than Small Business)",
                    "CONFIRM dropdown label in Grants.gov. FairBanks is a Uganda private limited company.",
                ),
                (
                    "9. Type of Applicant (2/3)",
                    "Optional: Other – Foreign / non-U.S. entity (Uganda)",
                    "Only if a second type is allowed and improves clarity.",
                ),
                (
                    "9. Other (specify)",
                    "Uganda-registered private limited company (Company No. "
                    f"{COMPANY_NO}; TIN {TIN_UG}; NSSF {NSSF})",
                    "Use if form asks Other Specify.",
                ),
                ("10. Name of Federal Agency", AGENCY, AGENCY_DETAIL),
                ("11. CFDA / Assistance Listing No.", CFDA, ""),
                ("11. CFDA / Assistance Listing Title", CFDA_TITLE, ""),
                ("12. Funding Opportunity Number", NOFO, ""),
                ("12. Funding Opportunity Title", NOFO_TITLE, ""),
                (
                    "13. Competition Identification Number",
                    "Leave blank unless package auto-fills",
                    "Often filled by Grants.gov workspace.",
                ),
                (
                    "13. Competition Identification Title",
                    "Leave blank unless package auto-fills",
                    "",
                ),
                (
                    "14. Areas Affected by Project",
                    "Uganda (primary). East Africa corridor learning under MoH Uganda leadership. "
                    "Additional countries only if LOIs / registration attached (CONFIRM).",
                    "Keep honest — do not claim unregistered field ops.",
                ),
                (
                    "15. Descriptive Title of Applicant's Project",
                    BOX15,
                    "Public USAspending text — describe what the project will do (NOFO requirement).",
                ),
            ],
            cols,
            sty,
        )
    )

    story.append(p("16–18. Districts, dates, Year 1 funding", sty["h1"]))
    story.append(
        section_table(
            [
                (
                    "16a. Applicant Congressional District",
                    "00-000",
                    "Use 00-000 when applicant is outside the United States.",
                ),
                (
                    "16b. Program/Project Congressional District(s)",
                    "00-000",
                    "Project performance outside the United States.",
                ),
                (
                    "17. Proposed Project Start Date",
                    START,
                    "NOFO expected award / start date: 30 September 2026.",
                ),
                (
                    "17. Proposed Project End Date",
                    END,
                    "Five 12-month budget periods (5-year period of performance).",
                ),
                (
                    "18a. Federal",
                    FED_ASK,
                    f"Year 1 total $7.5M (C1–C5). Core Component 1: {C1_ASK}. "
                    "C2–C5 are ABU contingency ($4.5M); SF-424A breaks out components.",
                ),
                ("18b. Applicant", "$0", "No cost share required or proposed."),
                ("18c. State", "$0", ""),
                ("18d. Local", "$0", ""),
                ("18e. Other", "$0", ""),
                ("18f. Program Income", "$0", "CONFIRM if any expected; default none."),
                ("18g. TOTAL", FED_ASK, "Must equal 18a–18f sum."),
            ],
            cols,
            sty,
        )
    )

    story.append(PageBreak())
    story.append(p("19–21. Review, debt, authorized representative", sty["h1"]))
    story.append(
        section_table(
            [
                (
                    "19. EO 12372 Review",
                    "c. Program is not covered by E.O. 12372",
                    "Typical for this CDC global / foreign implementation path. CONFIRM if Grants.gov shows otherwise.",
                ),
                (
                    "19. State review date",
                    "Leave blank",
                    "Only if 19a selected.",
                ),
                (
                    "20. Delinquent on Federal Debt?",
                    "No",
                    "If Yes, explanation attachment required.",
                ),
                (
                    "21. Authorized Representative — Prefix",
                    "Ms.",
                    "CONFIRM.",
                ),
                ("21. First Name", "Racheal", ""),
                ("21. Middle Name", "Leave blank (or CONFIRM)", ""),
                ("21. Last Name", "Nabukeera", ""),
                ("21. Suffix", "Leave blank", ""),
                ("21. Title", TITLE, "Must be authorized to bind the organisation."),
                ("21. Telephone", PHONE, ""),
                ("21. Fax", FAX, ""),
                ("21. Email", EMAIL, ""),
                (
                    "21. Signature of Authorized Representative",
                    "Sign in Grants.gov / Adobe at submit",
                    "Governing-body authorization to sign must be on file.",
                ),
                (
                    "21. Date Signed",
                    "Enter date of electronic / wet signature",
                    f"Submit by 14 August 2026, 11:59 p.m. ET.",
                ),
            ],
            cols,
            sty,
        )
    )

    story.append(p("Year 1 federal ask — component reminder (for SF-424A, not Box 18 lines)", sty["h1"]))
    story.append(
        p(
            "Box 18 on SF-424 is the first budget-period total. Component detail goes on SF-424A "
            "(and a second SF-424A if more than four activity rows).",
            sty["body"],
        )
    )
    comp = Table(
        [
            [
                p("<b>Component</b>", sty["cell_label"]),
                p("<b>Draft ask</b>", sty["cell_label"]),
                p("<b>Ceiling</b>", sty["cell_label"]),
                p("<b>Note</b>", sty["cell_label"]),
            ],
            [
                p("1 Core GHS", sty["cell_value"]),
                p(C1_ASK, sty["cell_value"]),
                p("$3,000,000", sty["cell_value"]),
                p("Expected initial funding (full ceiling)", sty["cell_value"]),
            ],
            [
                p("2 Small-scale outbreak / PHE", sty["cell_value"]),
                p(C2_ASK, sty["cell_value"]),
                p("$5,000,000", sty["cell_value"]),
                p("ABU — moderate PHE surge", sty["cell_value"]),
            ],
            [
                p("3 Large-scale outbreak / PHE", sty["cell_value"]),
                p(C3_ASK, sty["cell_value"]),
                p("$10,000,000", sty["cell_value"]),
                p("ABU — substantial PHE under MoH", sty["cell_value"]),
            ],
            [
                p("4 Small emerging threats", sty["cell_value"]),
                p(C4_ASK, sty["cell_value"]),
                p("$15,000,000", sty["cell_value"]),
                p("ABU — scaled contingency", sty["cell_value"]),
            ],
            [
                p("5 Large emerging threats", sty["cell_value"]),
                p(C5_ASK, sty["cell_value"]),
                p("$20,000,000", sty["cell_value"]),
                p("ABU — SF-424A activity row required", sty["cell_value"]),
            ],
            [
                p("<b>Total feasible (C1–C5)</b>", sty["cell_value"]),
                p(f"<b>{FED_ASK}</b>", sty["cell_value"]),
                p("", sty["cell_value"]),
                p("Equals Box 18a / 18g", sty["cell_value"]),
            ],
        ],
        colWidths=[2.0 * inch, 1.2 * inch, 1.1 * inch, w - 4.3 * inch],
        hAlign="LEFT",
    )
    comp.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), CREAM),
                ("BOX", (0, 0), (-1, -1), 0.6, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E8F3F3")),
            ]
        )
    )
    story.append(comp)

    story.append(Spacer(1, 10))
    story.append(p("Pre-submit checks for this form", sty["h1"]))
    checks = [
        "Active SAM.gov registration + UEI for Kampala physical location receiving funds",
        "Legal name on SF-424 matches SAM.gov character-for-character",
        "Foreign EIN entry confirmed as 44-4444444 (or agency-specific instruction)",
        "Box 15 is a short public description of what the project will do (not a vague title)",
        "Box 18 total matches final SF-424A Year 1 federal total",
        "Congressional districts set to 00-000 for non-U.S. applicant/project",
        "Authorized official (Racheal Nabukeera) certifies in Grants.gov before deadline",
        "Keep distinction clear if any parallel CDC-RFA-JG-26-0054 package is also submitted",
    ]
    for c in checks:
        story.append(p(f"• {c}", sty["body"]))

    story.append(Spacer(1, 8))
    story.append(
        p(
            "Sources used: NOFO CDC-RFA-JG-26-0056; applications/cdc-ghs-global/rules/source_of_truth.mdc; "
            "application_answers.md; company-docs (Company No., TIN, NSSF); SF-424 V4.0 field list from "
            "WS01740564-SF424_4_0-V4.0.pdf.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
