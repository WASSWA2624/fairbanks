#!/usr/bin/env python3
"""Generate FairBanks Project Abstract Summary V2.0 fill answers PDF.

Blank form:
  documents to fill/WS01740564-Project_AbstractSummary_2_0-V2.0.pdf

Output:
  documents to fill/WS01740564-Project_AbstractSummary_2_0-V2.0_ANSWERS.pdf
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "documents to fill"
OUT = OUT_DIR / "WS01740564-Project_AbstractSummary_2_0-V2.0_ANSWERS.pdf"
OUT_ALT = OUT_DIR / "WS01740564-Project_AbstractSummary_2_0-V2.0_ANSWERS_FINAL.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
ORANGE = colors.HexColor("#C45C26")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
CREAM = colors.HexColor("#F7F5F0")
LINE = colors.HexColor("#CED9D8")
WHITE = colors.white

NOFO = "CDC-RFA-JG-26-0056"
CFDA = "93.318"
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
TITLE = (
    "FairBanks strengthens last-mile surveillance across the region for faster "
    "outbreak detection and response, using CHWs/VHTs, FCHIP, 7-1-7 timing, "
    "surge readiness and contingency capacity."
)

# Public abstract: plain language, ASCII only, no PII. Max 4000 chars.
# Human thread: recognise -> connect -> act -> contain (people-centered, MoH-led).
ABSTRACT = (
    "FAIRBANKS MEDICAL CENTRE LIMITED (FairBanks) is a Uganda-registered health "
    "organisation committed to quality healthcare and stronger public health "
    "protection closer to communities. FairBanks operates a licensed medical "
    "centre and FairBanks Community Reach, working with Community Health Workers "
    "(CHWs) and Village Health Teams (VHTs) in Kampala peri-urban communities. "
    "Our slogan is Your Health, Our Mission.\n\n"
    "Under CDC-RFA-JG-26-0056, FairBanks proposes to strengthen last-mile "
    "capacity to recognise, report, and respond to disease threats earlier. Under "
    "Ministry of Health (MoH) leadership, the project will bring surveillance "
    "closer to where people live and seek care, so early warning signals reach "
    "the right people before threats become larger public health emergencies.\n\n"
    "Purpose. Early outbreak signals may first appear in communities or health "
    "facilities but fail to reach formal surveillance systems quickly. FairBanks "
    "will help close these gaps by strengthening community and facility "
    "surveillance and improving the flow of structured information into MoH and "
    "National Integrated Surveillance System (NISS)-aligned pathways through the "
    "FairBanks Community Health Intelligence Platform (FCHIP). FCHIP will "
    "strengthen existing systems rather than create a parallel data silo. Routine "
    "clinical care will continue under FairBanks operations and is not the "
    "purpose of this award.\n\n"
    "Year 1 - Component 1 will: (1) Strengthen surveillance and data quality "
    "through CHWs/VHTs and frontline teams. (2) Improve reporting links among "
    "communities, facilities, laboratories, districts, and MoH. (3) Apply 7-1-7 "
    "principles to improve the speed from detection to notification and early "
    "response. (4) Train and supervise a surge-ready workforce with One Health "
    "awareness. (5) Strengthen community-to-facility referral and follow-up for "
    "priority disease signals. (6) Coordinate with MoH, district authorities, "
    "Kampala Capital City Authority, and other CDC partners. Laboratory and "
    "selected surge activities will be delivered through MoH-aligned partners. "
    "Tools, SOPs, software, and training materials developed under the award "
    "will be available to MoH and CDC for appropriate use.\n\n"
    "Components 2-5 will establish contingency plans and budgets for moderate "
    "and substantial responses to outbreaks, public health emergencies, and "
    "emerging infectious disease threats. These components will remain unfunded "
    "until CDC activates emergency funding. Total Year 1 federal request: USD "
    "7,500,000, including USD 3,000,000 for Component 1.\n\n"
    "Geography. Uganda will be the primary country of implementation in Year 1, "
    "with East African corridor learning and cross-border signal protocols under "
    "MoH Uganda leadership. Additional countries will only be included where "
    "partner documentation and registration requirements are met.\n\n"
    "Expected outcomes include faster detection and response; stronger community "
    "and facility reporting; successful MoH/NISS-aligned data exchange tests; "
    "trained CHWs/VHTs and surge rosters; improved referral and follow-up for "
    "priority signals; and practical tools and learning for government and "
    "partners. About 5-10% of funds will support monitoring, evaluation, and "
    "learning aligned with DGHP indicators agreed with CDC after award.\n\n"
    "Beneficiaries include communities served through FairBanks Community Reach, "
    "CHWs/VHTs, frontline workers, public health decision-makers, and regional "
    "partners. Earlier detection and containment will help protect communities "
    "and strengthen health security across the region.\n\n"
    "FairBanks is applying as an in-country local partner, bringing established "
    "community relationships, frontline health experience, and practical "
    "knowledge of last-mile health challenges. FairBanks will document "
    "local-partner preference criteria required by the funding notice.\n\n"
    "At its heart, this project is about helping communities recognise health "
    "threats earlier, connecting those signals to the public health system, "
    "enabling timely action before a local threat becomes a wider emergency."
)


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            textColor=TEAL, spaceAfter=6,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=16,
            leading=20, textColor=NAVY, spaceAfter=8,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            leading=12, textColor=SLATE, spaceAfter=4,
        ),
        "small": ParagraphStyle(
            "small", parent=base["Normal"], fontName="Helvetica", fontSize=8,
            leading=10, textColor=MUTED, spaceAfter=3,
        ),
        "label": ParagraphStyle(
            "label", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=8,
            leading=10, textColor=NAVY,
        ),
        "value": ParagraphStyle(
            "value", parent=base["Normal"], fontName="Helvetica", fontSize=8.5,
            leading=11, textColor=SLATE,
        ),
        "note": ParagraphStyle(
            "note", parent=base["Normal"], fontName="Helvetica-Oblique", fontSize=8,
            leading=10, textColor=ORANGE,
        ),
        "abstract": ParagraphStyle(
            "abstract", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            leading=12, textColor=SLATE, spaceAfter=4,
        ),
    }


def p(text: str, style) -> Paragraph:
    return Paragraph(text.replace("\n", "<br/>"), style)


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
        "FairBanks · Project Abstract Summary V2.0 answers · CDC-RFA-JG-26-0056",
    )
    canvas.drawRightString(LETTER[0] - 0.7 * inch, LETTER[1] - 0.42 * inch, f"Page {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(0.7 * inch, 0.55 * inch, LETTER[0] - 0.7 * inch, 0.55 * inch)
    canvas.drawCentredString(
        LETTER[0] / 2,
        0.35 * inch,
        "Your health, our mission. · Companion worksheet — paste into Grants.gov / Adobe form",
    )
    canvas.restoreState()


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    chars = len(ABSTRACT)
    assert chars <= 4000, f"Abstract too long: {chars}"

    sty = styles()
    target = OUT
    try:
        # Prefer canonical name; fall back if the file is open/locked.
        if OUT.exists():
            with open(OUT, "ab"):
                pass
    except PermissionError:
        target = OUT_ALT

    doc = SimpleDocTemplate(
        str(target),
        pagesize=LETTER,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
        title="Project Abstract Summary V2.0 Answers — FairBanks — CDC-RFA-JG-26-0056",
        author=ORG,
    )
    w = LETTER[0] - 1.4 * inch
    cols = [1.7 * inch, w - 1.7 * inch]
    story = []

    story.append(p("GRANTS.GOV FORM WORKSHEET", sty["kicker"]))
    story.append(p("Project Abstract Summary V2.0 — Fill Answers", sty["title"]))
    story.append(
        p(
            f"<b>Applicant:</b> {ORG}<br/>"
            f"<b>Opportunity:</b> {NOFO}<br/>"
            f"<b>Form file:</b> WS01740564-Project_AbstractSummary_2_0-V2.0.pdf<br/>"
            f"<b>Prepared:</b> {date.today().isoformat()} · Slogan: Your health, our mission.",
            sty["body"],
        )
    )
    story.append(
        p(
            "How to use: Open the blank Project Abstract Summary form and paste each field below. "
            "The Project Abstract is public (USAspending) — no personal names, emails, phones, or "
            "proprietary detail. Use plain English characters only. Limit: 4,000 characters.",
            sty["body"],
        )
    )
    story.append(p(f"Character count for Project Abstract: <b>{chars}</b> / 4000", sty["note"]))
    story.append(Spacer(1, 6))

    rows = [
        ("Funding Opportunity Number", NOFO, "Must match SF-424 Box 12."),
        (
            "Assistance Listing Number(s)",
            CFDA,
            "Protecting and Improving Health Globally (CFDA 93.318). May auto-fill from workspace.",
        ),
        ("Applicant Name", ORG, "Must match SAM.gov / SF-424 legal name."),
        (
            "Descriptive Title of Applicant's Project",
            TITLE,
            "Same public title as SF-424 Box 15.",
        ),
    ]
    data = []
    for label, value, note in rows:
        data.append(
            [
                p(label, sty["label"]),
                [p(value, sty["value"]), p(note, sty["note"])],
            ]
        )
    t = Table(data, colWidths=cols, hAlign="LEFT")
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), CREAM),
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
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(p("Project Abstract (paste as one block)", sty["label"]))
    story.append(Spacer(1, 4))

    # Show abstract in a bordered table for easy copy
    abs_table = Table(
        [[p(ABSTRACT.replace("\n\n", "\n\n"), sty["abstract"])]],
        colWidths=[w],
        hAlign="LEFT",
    )
    abs_table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.8, TEAL),
                ("BACKGROUND", (0, 0), (-1, -1), WHITE),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(abs_table)
    story.append(Spacer(1, 8))
    story.append(
        p(
            "Checks: form required or application incomplete · match SF-424 title and legal name · "
            "no PII · ASCII only · under 4,000 characters · Year 1 total aligns to $7,500,000.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {target} ({chars} abstract chars)")


if __name__ == "__main__":
    build()
