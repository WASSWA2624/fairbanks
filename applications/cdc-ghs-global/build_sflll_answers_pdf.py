#!/usr/bin/env python3
"""Generate FairBanks SF-LLL (Disclosure of Lobbying Activities) V2.0 answers PDF.

Blank form:
  documents to fill/WS01740564-SFLLL_2_0-V2.0.pdf

Output:
  documents to fill/WS01740564-SFLLL_2_0-V2.0_ANSWERS.pdf
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
OUT = OUT_DIR / "WS01740564-SFLLL_2_0-V2.0_ANSWERS.pdf"
OUT_ALT = OUT_DIR / "WS01740564-SFLLL_2_0-V2.0_ANSWERS_FINAL.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
ORANGE = colors.HexColor("#C45C26")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
CREAM = colors.HexColor("#F7F5F0")
LINE = colors.HexColor("#CED9D8")
WHITE = colors.white
OK_BG = colors.HexColor("#E8F3F3")

ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
STREET1 = "Plot 1423 and 1425 Tirupati Road"
STREET2 = "Fairbanks Medical Centre, Kololo IV, Lugogo"
CITY = "Kampala"
STATE = "N/A (Uganda — foreign entity)"
ZIP = "CONFIRM / leave blank if form allows for foreign address"
CONG = "00-000"
AGENCY = "Centers for Disease Control and Prevention (CDC)"
CFDA = "93.318"
CFDA_TITLE = (
    "Protecting and Improving Health Globally: Building and Strengthening "
    "Public Health Impact, Systems, Capacity, and Security"
)
ACTION_NO = "CDC-RFA-JG-26-0056"
CONTACT = "Racheal Nabukeera"
TITLE = "Managing Director and Co-founder"
PHONE = "+256 772 849 258"


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            textColor=TEAL, spaceAfter=6,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=15,
            leading=19, textColor=NAVY, spaceAfter=8,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=11,
            textColor=NAVY, spaceBefore=8, spaceAfter=5,
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
        "FairBanks · SF-LLL V2.0 answers · CDC-RFA-JG-26-0056",
    )
    canvas.drawRightString(LETTER[0] - 0.7 * inch, LETTER[1] - 0.42 * inch, f"Page {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(0.7 * inch, 0.55 * inch, LETTER[0] - 0.7 * inch, 0.55 * inch)
    canvas.drawCentredString(
        LETTER[0] / 2,
        0.35 * inch,
        "Your health, our mission. · Companion worksheet — Grants.gov / Adobe SF-LLL",
    )
    canvas.restoreState()


def section_table(rows, cols, sty):
    data = []
    for label, value, note in rows:
        right = [p(value, sty["value"])]
        if note:
            right.append(p(note, sty["note"]))
        data.append([p(label, sty["label"]), right])
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
    return t


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    sty = styles()
    target = OUT
    try:
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
        title="SF-LLL V2.0 Answers — FairBanks — CDC-RFA-JG-26-0056",
        author=ORG,
    )
    w = LETTER[0] - 1.4 * inch
    cols = [1.75 * inch, w - 1.75 * inch]
    story = []

    story.append(p("GRANTS.GOV FORM WORKSHEET", sty["kicker"]))
    story.append(p("Disclosure of Lobbying Activities (SF-LLL) V2.0 — Fill Answers", sty["title"]))
    story.append(
        p(
            f"<b>Applicant:</b> {ORG}<br/>"
            f"<b>Opportunity:</b> {ACTION_NO}<br/>"
            f"<b>Form file:</b> WS01740564-SFLLL_2_0-V2.0.pdf<br/>"
            f"<b>Prepared:</b> {date.today().isoformat()} · Slogan: Your health, our mission.",
            sty["body"],
        )
    )

    # Verdict banner
    verdict = Table(
        [[
            p(
                "<b>Primary answer: NO REPORTABLE LOBBYING.</b><br/>"
                "CDC lists SF-LLL as required <b>if applicable</b>. FairBanks has not hired a "
                "paid lobbying registrant to influence this cooperative agreement application. "
                "<b>Do not invent lobbyist names or payments.</b> Authorized official must CONFIRM "
                "before submit. If the Grants.gov package forces the form, use Path B below with N/A.",
                sty["body"],
            )
        ]],
        colWidths=[w],
        hAlign="LEFT",
    )
    verdict.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), OK_BG),
                ("BOX", (0, 0), (-1, -1), 1.0, TEAL),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(verdict)

    story.append(p("Path A — Preferred (no lobbying)", sty["h1"]))
    story.append(
        p(
            "1. Confirm with Racheal Nabukeera / board that FairBanks paid <b>no</b> outside "
            "lobbyist to influence CDC, HHS, Congress, or related officials on this application.<br/>"
            "2. In Grants.gov, leave SF-LLL <b>unattached</b> / mark not applicable if the workspace allows.<br/>"
            "3. Keep the anti-lobbying certification that comes with federal applications "
            "(separate from SF-LLL) — that certifies federal funds will not be used to lobby.<br/>"
            "4. If lobbying is hired later, file SF-LLL promptly (material change / new filing).",
            sty["body"],
        )
    )

    story.append(p("Path B — Only if the workspace requires SF-LLL to be completed", sty["h1"]))
    story.append(
        p(
            "Use these values. For lobbyist boxes, enter <b>N/A</b> (not fictitious people). "
            "Foreign address fields: use Uganda details; State/ZIP as the form allows.",
            sty["body"],
        )
    )
    story.append(
        section_table(
            [
                (
                    "1. Type of Federal Action",
                    "c. Cooperative agreement",
                    "Matches CDC-RFA-JG-26-0056 funding instrument.",
                ),
                (
                    "2. Status of Federal Action",
                    "a. Bid/offer/application",
                    "Application stage — not yet awarded.",
                ),
                (
                    "3. Report Type",
                    "a. Initial filing",
                    "Leave Material Change year/quarter/date blank.",
                ),
                (
                    "4. Reporting Entity",
                    "Prime",
                    "FairBanks is the prime applicant (not a subawardee).",
                ),
                ("4. Name", ORG, "Must match SAM.gov / SF-424."),
                ("4. Street 1", STREET1, ""),
                ("4. Street 2", STREET2, "Optional."),
                ("4. City", CITY, ""),
                ("4. State", STATE, "Foreign entity — CONFIRM dropdown options."),
                ("4. ZIP", ZIP, "CONFIRM."),
                ("4. Congressional District", CONG, "Outside U.S."),
                (
                    "5. Prime (if subawardee)",
                    "Leave blank",
                    "Not applicable — FairBanks is Prime.",
                ),
                ("6. Federal Department/Agency", AGENCY, "Global Health Center / DGHP."),
                ("7. CFDA / Assistance Listing No.", CFDA, ""),
                ("7. Federal program name / CFDA Title", CFDA_TITLE, "May auto-fill."),
                (
                    "8. Federal Action Number",
                    ACTION_NO,
                    "Opportunity / announcement number.",
                ),
                (
                    "9. Award Amount",
                    "Leave blank (or $7,500,000 if form insists)",
                    "Optional at application; no award yet. Do not treat as confirmed award.",
                ),
                (
                    "10.a Lobbying Registrant",
                    "First Name: N/A · Last Name: N/A · Street/City: N/A",
                    "CONFIRM no lobbyist. Do not invent names. State/ZIP N/A if allowed.",
                ),
                (
                    "10.b Individual Performing Services",
                    "First Name: N/A · Last Name: N/A · Address: N/A",
                    "Same — no individual lobbyist retained.",
                ),
                ("11. Certifying Official — Prefix", "Ms.", "CONFIRM."),
                ("11. First Name", "Racheal", ""),
                ("11. Middle Name", "Leave blank (or CONFIRM)", ""),
                ("11. Last Name", "Nabukeera", ""),
                ("11. Title", TITLE, ""),
                ("11. Telephone", PHONE, ""),
                (
                    "11. Signature / Date",
                    "Sign in Grants.gov / Adobe at submit",
                    "Date usually stamped by Grants.gov on submit.",
                ),
            ],
            cols,
            sty,
        )
    )

    story.append(p("Confirm before submit", sty["h1"]))
    for c in [
        "No paid lobbying registrant retained for this CDC application (Path A preferred)",
        "If Path B used: every lobbyist field is N/A — not a fake person",
        "Legal name matches SF-424 and SAM.gov",
        "Type of action = Cooperative agreement; Status = Application; Report = Initial",
        "Authorized official (Racheal Nabukeera) reviews and certifies",
    ]:
        story.append(p(f"• {c}", sty["body"]))

    story.append(Spacer(1, 6))
    story.append(
        p(
            "Sources: CDC-RFA-JG-26-0056 (SF-LLL if applicable); Grants.gov SF-LLL V2.0 instructions; "
            "FairBanks company identity from application pack / company-docs.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {target}")


if __name__ == "__main__":
    build()
