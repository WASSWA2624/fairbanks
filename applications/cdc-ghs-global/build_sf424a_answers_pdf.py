#!/usr/bin/env python3
"""Generate FairBanks SF-424A V1.0 fill answers PDF for CDC-RFA-JG-26-0056.

NOFO rule: SF-424A has only four activity columns. With five components,
submit TWO SF-424A forms (Form 1 = C1-C4; Form 2 = C5 under Other Attachments).

Blank form:
  documents to fill/WS01740564-SF424A-V1.0.pdf

Output:
  documents to fill/WS01740564-SF424A-V1.0_ANSWERS.pdf
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "documents to fill"
OUT = OUT_DIR / "WS01740564-SF424A-V1.0_ANSWERS.pdf"
OUT_ALT = OUT_DIR / "WS01740564-SF424A-V1.0_ANSWERS_FINAL.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
ORANGE = colors.HexColor("#C45C26")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
CREAM = colors.HexColor("#F7F5F0")
LINE = colors.HexColor("#CED9D8")
OK_BG = colors.HexColor("#E8F3F3")

CFDA = "93.318"
NOFO = "CDC-RFA-JG-26-0056"
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
INDIRECT_RATE = 0.08  # foreign organization 8% of MTDC (CONFIRM)

# Component totals (federal Year 1) — must match SF-424 Box 18
COMPONENTS = [
    {
        "id": "C1",
        "title": "Comp 1: Core GHS Priorities",
        "short": "Core GHS",
        "total": 3_000_000,
        "note": "Expected initial funding",
    },
    {
        "id": "C2",
        "title": "Comp 2: Small-Scale Outbreak / PHE",
        "short": "Small PHE",
        "total": 1_000_000,
        "note": "ABU contingency",
    },
    {
        "id": "C3",
        "title": "Comp 3: Large-Scale Outbreak / PHE",
        "short": "Large PHE",
        "total": 1_500_000,
        "note": "ABU contingency",
    },
    {
        "id": "C4",
        "title": "Comp 4: Small Emerging Threats",
        "short": "Small emerging",
        "total": 1_000_000,
        "note": "ABU contingency",
    },
    {
        "id": "C5",
        "title": "Comp 5: Large Emerging Threats",
        "short": "Large emerging",
        "total": 1_000_000,
        "note": "ABU — second SF-424A",
    },
]

# Object-class mix as share of DIRECT costs (sums to 1.0)
# Equipment excluded from MTDC for 8% indirect calc.
SHARES = {
    "personnel": 0.28,
    "fringe": 0.06,
    "travel": 0.07,
    "equipment": 0.05,
    "supplies": 0.04,
    "contractual": 0.25,
    "construction": 0.00,
    "other": 0.25,
}


def money(n: int) -> str:
    return f"${n:,.0f}"


def split_component(total: int) -> dict[str, int]:
    """Split total into object classes + 8% indirect on MTDC (Direct - Equipment)."""
    # Solve: Direct + 0.08*(Direct - Equipment) = Total
    # Equipment = 0.05 * Direct  =>  Direct + 0.08*(0.95*Direct) = Total
    # Direct * (1 + 0.076) = Total => Direct = Total / 1.076
    direct = int(round(total / (1 + INDIRECT_RATE * (1 - SHARES["equipment"]))))
    # Build category amounts from shares; fix rounding on "other"
    cats = {}
    running = 0
    keys = [
        "personnel",
        "fringe",
        "travel",
        "equipment",
        "supplies",
        "contractual",
        "construction",
    ]
    for k in keys:
        amt = int(round(direct * SHARES[k]))
        cats[k] = amt
        running += amt
    cats["other"] = direct - running
    if cats["other"] < 0:
        # rare rounding edge — pull from contractual
        cats["contractual"] += cats["other"]
        cats["other"] = 0
    mtdc = direct - cats["equipment"]
    indirect = int(round(mtdc * INDIRECT_RATE))
    # Adjust indirect so grand total matches exactly
    computed = direct + indirect
    if computed != total:
        indirect += total - computed
    cats["direct"] = direct
    cats["indirect"] = indirect
    cats["total"] = direct + indirect
    assert cats["total"] == total, (cats["total"], total, cats)
    return cats


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            textColor=TEAL, spaceAfter=6,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=14,
            leading=18, textColor=NAVY, spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=11,
            textColor=NAVY, spaceBefore=8, spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Helvetica", fontSize=8.5,
            leading=11, textColor=SLATE, spaceAfter=3,
        ),
        "small": ParagraphStyle(
            "small", parent=base["Normal"], fontName="Helvetica", fontSize=7.5,
            leading=9.5, textColor=MUTED, spaceAfter=2,
        ),
        "label": ParagraphStyle(
            "label", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=7.5,
            leading=9.5, textColor=NAVY,
        ),
        "value": ParagraphStyle(
            "value", parent=base["Normal"], fontName="Helvetica", fontSize=7.5,
            leading=9.5, textColor=SLATE,
        ),
        "note": ParagraphStyle(
            "note", parent=base["Normal"], fontName="Helvetica-Oblique", fontSize=7.5,
            leading=9.5, textColor=ORANGE,
        ),
        "th": ParagraphStyle(
            "th", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=7,
            leading=9, textColor=NAVY,
        ),
        "td": ParagraphStyle(
            "td", parent=base["Normal"], fontName="Helvetica", fontSize=7,
            leading=9, textColor=SLATE,
        ),
    }


def p(text: str, style) -> Paragraph:
    return Paragraph(str(text).replace("\n", "<br/>"), style)


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(TEAL)
    canvas.setLineWidth(1.2)
    canvas.line(0.6 * inch, LETTER[1] - 0.5 * inch, LETTER[0] - 0.6 * inch, LETTER[1] - 0.5 * inch)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(MUTED)
    canvas.drawString(
        0.6 * inch,
        LETTER[1] - 0.38 * inch,
        "FairBanks · SF-424A V1.0 answers · CDC-RFA-JG-26-0056",
    )
    canvas.drawRightString(LETTER[0] - 0.6 * inch, LETTER[1] - 0.38 * inch, f"Page {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(0.6 * inch, 0.5 * inch, LETTER[0] - 0.6 * inch, 0.5 * inch)
    canvas.drawCentredString(
        LETTER[0] / 2,
        0.32 * inch,
        "Your health, our mission. · Companion worksheet — paste into Grants.gov SF-424A",
    )
    canvas.restoreState()


def grid(data, widths, sty, header=True):
    t = Table(data, colWidths=widths, hAlign="LEFT", repeatRows=1 if header else 0)
    style_cmds = [
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if header:
        style_cmds.append(("BACKGROUND", (0, 0), (-1, 0), CREAM))
    # highlight last row
    style_cmds.append(("BACKGROUND", (0, -1), (-1, -1), OK_BG))
    t.setStyle(TableStyle(style_cmds))
    return t


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    splits = [split_component(c["total"]) for c in COMPONENTS]
    grand = sum(c["total"] for c in COMPONENTS)
    assert grand == 7_500_000

    form1 = list(zip(COMPONENTS[:4], splits[:4]))
    form2 = list(zip(COMPONENTS[4:], splits[4:]))
    form1_total = sum(c["total"] for c, _ in form1)
    form2_total = sum(c["total"] for c, _ in form2)

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
        leftMargin=0.6 * inch,
        rightMargin=0.6 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title="SF-424A V1.0 Answers — FairBanks — CDC-RFA-JG-26-0056",
        author=ORG,
    )
    w = LETTER[0] - 1.2 * inch
    story = []

    story.append(p("GRANTS.GOV FORM WORKSHEET", sty["kicker"]))
    story.append(p("Budget Information — Non-Construction Programs (SF-424A) V1.0", sty["title"]))
    story.append(
        p(
            f"<b>Applicant:</b> {ORG}<br/>"
            f"<b>Opportunity:</b> {NOFO} · CFDA {CFDA}<br/>"
            f"<b>Form file:</b> WS01740564-SF424A-V1.0.pdf<br/>"
            f"<b>Prepared:</b> {date.today().isoformat()} · Year 1 federal total: <b>{money(grand)}</b> "
            f"(matches SF-424 Box 18a/18g)",
            sty["body"],
        )
    )

    banner = Table(
        [[
            p(
                "<b>NOFO rule:</b> SF-424A has only <b>four</b> activity columns. FairBanks has "
                "<b>five</b> components — complete <b>two</b> SF-424A forms.<br/>"
                f"<b>Form 1</b> (main package): Components 1–4 = {money(form1_total)} · "
                f"<b>Form 2</b> (Other Attachments): Component 5 = {money(form2_total)}.<br/>"
                "All figures are <b>federal</b>; non-federal / cost share = $0. Indirect = "
                "<b>8% of MTDC</b> (foreign org — CONFIRM). Object-class lines are a "
                "<b>fill worksheet</b>; finalise in the budget narrative before submit.",
                sty["body"],
            )
        ]],
        colWidths=[w],
        hAlign="LEFT",
    )
    banner.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), OK_BG),
                ("BOX", (0, 0), (-1, -1), 1.0, TEAL),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(banner)
    story.append(Spacer(1, 6))

    # Hero total — all five components = $7.5M (matches SF-424 Box 18)
    story.append(p("Year 1 federal TOTAL (all five components) = $7,500,000", sty["h1"]))
    hero = [
        [
            p("<b>Component</b>", sty["th"]),
            p("<b>Federal ask</b>", sty["th"]),
            p("<b>Which SF-424A</b>", sty["th"]),
            p("<b>Note</b>", sty["th"]),
        ]
    ]
    for comp in COMPONENTS:
        which = "Form 1" if comp["id"] != "C5" else "Form 2 (Other Attachments)"
        hero.append(
            [
                p(f"{comp['id']} — {comp['short']}", sty["td"]),
                p(money(comp["total"]), sty["td"]),
                p(which, sty["td"]),
                p(comp["note"], sty["td"]),
            ]
        )
    hero.append(
        [
            p("<b>TOTAL Year 1</b>", sty["td"]),
            p(f"<b>{money(grand)}</b>", sty["td"]),
            p("<b>Form 1 + Form 2</b>", sty["td"]),
            p("<b>= SF-424 Box 18a / 18g</b>", sty["td"]),
        ]
    )
    story.append(
        grid(hero, [1.7 * inch, 1.2 * inch, 1.9 * inch, 2.2 * inch], sty)
    )
    story.append(
        p(
            "Do not use Form 1 alone ($6,500,000) as the application total. "
            "The full Year 1 federal ask is <b>$7,500,000</b>.",
            sty["note"],
        )
    )

    # ----- FORM 1 Section A -----
    story.append(p("FORM 1 — Section A. Budget Summary (Components 1–4) — Form 1 subtotal only", sty["h1"]))
    headers = [
        p("<b>Grant program / function / activity</b>", sty["th"]),
        p("<b>CFDA</b>", sty["th"]),
        p("<b>Unoblig. Fed</b>", sty["th"]),
        p("<b>Unoblig. Non-Fed</b>", sty["th"]),
        p("<b>New Fed</b>", sty["th"]),
        p("<b>New Non-Fed</b>", sty["th"]),
        p("<b>Total</b>", sty["th"]),
    ]
    rows = [headers]
    for col, (comp, sp) in enumerate(form1, start=1):
        letter = chr(ord("a") + col - 1)
        rows.append(
            [
                p(f"({letter}) {comp['title']}", sty["td"]),
                p(CFDA, sty["td"]),
                p("$0", sty["td"]),
                p("$0", sty["td"]),
                p(money(sp["total"]), sty["td"]),
                p("$0", sty["td"]),
                p(money(sp["total"]), sty["td"]),
            ]
        )
    rows.append(
        [
            p("<b>5. Form 1 subtotal (C1–C4)</b>", sty["td"]),
            p("", sty["td"]),
            p("<b>$0</b>", sty["td"]),
            p("<b>$0</b>", sty["td"]),
            p(f"<b>{money(form1_total)}</b>", sty["td"]),
            p("<b>$0</b>", sty["td"]),
            p(f"<b>{money(form1_total)}</b>", sty["td"]),
        ]
    )
    story.append(
        grid(
            rows,
            [2.35 * inch, 0.55 * inch, 0.75 * inch, 0.85 * inch, 0.85 * inch, 0.85 * inch, 0.8 * inch],
            sty,
        )
    )
    story.append(
        p(
            "New application: Estimated Unobligated Funds = $0 for all columns.",
            sty["note"],
        )
    )

    # ----- FORM 1 Section B -----
    story.append(p("FORM 1 — Section B. Budget Categories (object class)", sty["h1"]))
    cat_labels = [
        ("6.a Personnel", "personnel"),
        ("6.b Fringe Benefits", "fringe"),
        ("6.c Travel", "travel"),
        ("6.d Equipment", "equipment"),
        ("6.e Supplies", "supplies"),
        ("6.f Contractual", "contractual"),
        ("6.g Construction", "construction"),
        ("6.h Other", "other"),
        ("6.i Total Direct", "direct"),
        ("6.j Indirect Charges", "indirect"),
        ("6.k Form 1 TOTALS", "total"),
    ]
    brow = [
        p("<b>Object class</b>", sty["th"]),
        *[p(f"<b>({chr(ord('a')+i)}) {c['short']}</b>", sty["th"]) for i, (c, _) in enumerate(form1)],
        p("<b>Form 1 total</b>", sty["th"]),
    ]
    brows = [brow]
    for label, key in cat_labels:
        vals = [sp[key] for _, sp in form1]
        brows.append(
            [
                p(label, sty["td"]),
                *[p(money(v), sty["td"]) for v in vals],
                p(money(sum(vals)), sty["td"]),
            ]
        )
    story.append(
        grid(
            brows,
            [1.45 * inch, 1.15 * inch, 1.15 * inch, 1.15 * inch, 1.15 * inch, 1.0 * inch],
            sty,
        )
    )
    story.append(
        p(
            "6.h Other includes: CHW/VHT field support, training/surge drills, M&E, "
            "community surveillance ops, and related programme costs not booked as contractual. "
            f"CONFIRM final narrative mapping. Form 1 Section B total {money(form1_total)} + "
            f"Form 2 {money(form2_total)} = <b>{money(grand)}</b>.",
            sty["note"],
        )
    )

    story.append(PageBreak())

    # Section C / D / E / F for Form 1
    story.append(p("FORM 1 — Section C. Non-Federal Resources", sty["h1"]))
    story.append(
        p(
            "All columns (Applicant / State / Other / Totals) = <b>$0</b>. No cost share proposed.",
            sty["body"],
        )
    )

    story.append(p("FORM 1 — Section D. Forecasted Cash Needs (Year 1)", sty["h1"]))
    q = form1_total // 4
    rem = form1_total - 3 * q
    quarters = [q, q, q, rem]
    story.append(
        grid(
            [
                [
                    p("<b></b>", sty["th"]),
                    p("<b>1st Quarter</b>", sty["th"]),
                    p("<b>2nd Quarter</b>", sty["th"]),
                    p("<b>3rd Quarter</b>", sty["th"]),
                    p("<b>4th Quarter</b>", sty["th"]),
                    p("<b>Total Year 1</b>", sty["th"]),
                ],
                [
                    p("13. Federal", sty["td"]),
                    *[p(money(x), sty["td"]) for x in quarters],
                    p(money(form1_total), sty["td"]),
                ],
                [
                    p("14. Non-Federal", sty["td"]),
                    *[p("$0", sty["td"]) for _ in range(4)],
                    p("$0", sty["td"]),
                ],
                [
                    p("15. TOTAL", sty["td"]),
                    *[p(money(x), sty["td"]) for x in quarters],
                    p(money(form1_total), sty["td"]),
                ],
            ],
            [1.2 * inch, 1.1 * inch, 1.1 * inch, 1.1 * inch, 1.1 * inch, 1.2 * inch],
            sty,
        )
    )
    story.append(
        p(
            "Even quarterly split for Form 1 total. ABU components may remain unfunded until "
            "CDC activation — cash needs then follow activation timing (CONFIRM with CDC).",
            sty["note"],
        )
    )

    story.append(p("FORM 1 — Section E. Budget Estimates of Federal Funds Needed for Balance of Project", sty["h1"]))
    story.append(
        p(
            "Paste these amounts into the blue cells. Columns (b)–(e) = project Years 2–5 "
            "(the balance after Year 1 in Sections A–D). Use the <b>same Year 1 federal ask "
            "per component</b> as the planning estimate for each future year. ABU lines "
            "(17–19) are contingency ceilings — funded only if CDC activates; continuation "
            "amounts are set by CDC (CONFIRM).",
            sty["body"],
        )
    )
    # Full Year-1 component asks repeated for each future year (accurate fill for lines 16-19)
    e_amounts = [c["total"] for c, _ in form1]  # C1..C4
    e_total = sum(e_amounts)
    story.append(
        grid(
            [
                [
                    p("<b>Grant Program (matches Section A)</b>", sty["th"]),
                    p("<b>(b) First (Yr2)</b>", sty["th"]),
                    p("<b>(c) Second (Yr3)</b>", sty["th"]),
                    p("<b>(d) Third (Yr4)</b>", sty["th"]),
                    p("<b>(e) Fourth (Yr5)</b>", sty["th"]),
                ],
                [
                    p("16. Core GHS Priorities", sty["td"]),
                    *[p(money(e_amounts[0]), sty["td"]) for _ in range(4)],
                ],
                [
                    p("17. Small-Scale Outbreak / PHE", sty["td"]),
                    *[p(money(e_amounts[1]), sty["td"]) for _ in range(4)],
                ],
                [
                    p("18. Large-Scale Outbreak / PHE", sty["td"]),
                    *[p(money(e_amounts[2]), sty["td"]) for _ in range(4)],
                ],
                [
                    p("19. Small Emerging Threats", sty["td"]),
                    *[p(money(e_amounts[3]), sty["td"]) for _ in range(4)],
                ],
                [
                    p("<b>20. TOTAL (sum of 16–19)</b>", sty["td"]),
                    *[p(f"<b>{money(e_total)}</b>", sty["td"]) for _ in range(4)],
                ],
            ],
            [2.3 * inch, 1.25 * inch, 1.25 * inch, 1.25 * inch, 1.25 * inch],
            sty,
        )
    )
    story.append(
        p(
            f"Each future-year column totals <b>{money(e_total)}</b> (Form 1 only). "
            f"Add Form 2 Comp 5 at {money(form2_total)}/year for a full five-component "
            f"out-year planning stack of <b>{money(e_total + form2_total)}</b>/year "
            f"(same as Year 1 $7.500,000). Do <b>not</b> leave lines 16–19 blank.",
            sty["note"],
        )
    )

    story.append(p("FORM 1 — Section F. Other Budget Information", sty["h1"]))
    story.append(
        p(
            "<b>21. Direct Charges:</b> Personnel, fringe, travel, equipment, supplies, contractual "
            "(MoH-aligned partners / technical support), and other programme costs (CHW/VHT field "
            "support, training, M&E, community surveillance). Detail in budget narrative per component.<br/>"
            "<b>22. Indirect Charges:</b> Foreign organization rate of <b>8% of Modified Total Direct "
            "Costs (MTDC)</b>. MTDC = Total Direct minus Equipment (and other standard exclusions as "
            "applicable). CONFIRM final rate treatment with CDC/GMS.<br/>"
            "<b>23. Remarks:</b> Five-component application under CDC-RFA-JG-26-0056. Form 1 covers "
            "Components 1–4. Component 5 is on a second SF-424A uploaded under Other Attachments. "
            "Component 1 is expected initial funding; Components 2–4 are contingency ABU. Total "
            f"Form 1 federal = {money(form1_total)}. Overall Year 1 federal with Form 2 = {money(grand)}.",
            sty["body"],
        )
    )

    story.append(PageBreak())

    # ----- FORM 2 -----
    story.append(p("FORM 2 — Second SF-424A (Component 5 only) — upload under Other Attachments", sty["h1"]))
    story.append(
        p(
            "File name suggestion: <b>SF424A_Component5_LargeEmergingThreats.pdf</b>",
            sty["note"],
        )
    )
    comp5, sp5 = form2[0]
    story.append(p("FORM 2 — Section A. Budget Summary", sty["h1"]))
    story.append(
        grid(
            [
                [
                    p("<b>Grant program / function / activity</b>", sty["th"]),
                    p("<b>CFDA</b>", sty["th"]),
                    p("<b>Unoblig. Fed</b>", sty["th"]),
                    p("<b>New Fed</b>", sty["th"]),
                    p("<b>Non-Fed</b>", sty["th"]),
                    p("<b>Total</b>", sty["th"]),
                ],
                [
                    p(f"(a) {comp5['title']}", sty["td"]),
                    p(CFDA, sty["td"]),
                    p("$0", sty["td"]),
                    p(money(sp5["total"]), sty["td"]),
                    p("$0", sty["td"]),
                    p(money(sp5["total"]), sty["td"]),
                ],
                [
                    p("<b>5. Totals</b>", sty["td"]),
                    p("", sty["td"]),
                    p("<b>$0</b>", sty["td"]),
                    p(f"<b>{money(form2_total)}</b>", sty["td"]),
                    p("<b>$0</b>", sty["td"]),
                    p(f"<b>{money(form2_total)}</b>", sty["td"]),
                ],
            ],
            [2.6 * inch, 0.6 * inch, 0.9 * inch, 0.9 * inch, 0.8 * inch, 0.9 * inch],
            sty,
        )
    )

    story.append(p("FORM 2 — Section B. Budget Categories", sty["h1"]))
    f2rows = [[p("<b>Object class</b>", sty["th"]), p("<b>(a) Comp 5</b>", sty["th"])]]
    for label, key in cat_labels:
        f2rows.append([p(label, sty["td"]), p(money(sp5[key]), sty["td"])])
    story.append(grid(f2rows, [2.5 * inch, 2.0 * inch], sty))

    story.append(p("FORM 2 — Sections C–F (summary)", sty["h1"]))
    q5 = form2_total // 4
    rem5 = form2_total - 3 * q5
    story.append(
        p(
            f"<b>C.</b> Non-federal = $0.<br/>"
            f"<b>D.</b> Federal cash needs by quarter: {money(q5)} / {money(q5)} / {money(q5)} / "
            f"{money(rem5)} (Total {money(form2_total)}).<br/>"
            f"<b>E.</b> Future years (b)–(e): enter <b>{money(form2_total)}</b> in each column "
            f"for Comp 5 (same Year 1 ask; ABU / activation-dependent — CONFIRM).<br/>"
            f"<b>F.</b> Same indirect method (8% MTDC). Remarks: Component 5 Large Emerging Threats "
            f"contingency ABU; companion to Form 1 Components 1–4.",
            sty["body"],
        )
    )

    # Combined check
    story.append(p("Combined check vs SF-424 Box 18", sty["h1"]))
    story.append(
        grid(
            [
                [
                    p("<b>Source</b>", sty["th"]),
                    p("<b>Amount</b>", sty["th"]),
                ],
                [p("SF-424A Form 1 (C1–C4)", sty["td"]), p(money(form1_total), sty["td"])],
                [p("SF-424A Form 2 (C5)", sty["td"]), p(money(form2_total), sty["td"])],
                [p("<b>Year 1 federal TOTAL</b>", sty["td"]), p(f"<b>{money(grand)}</b>", sty["td"])],
                [p("SF-424 Box 18a / 18g must equal", sty["td"]), p(money(grand), sty["td"])],
            ],
            [4.5 * inch, 2.0 * inch],
            sty,
        )
    )

    story.append(Spacer(1, 8))
    story.append(p("Pre-submit checks", sty["h1"]))
    for c in [
        "Two SF-424A files completed (C1–C4 + C5)",
        "Each component under its NOFO ceiling",
        "Federal totals match SF-424 Box 18 ($7,500,000)",
        "Non-federal / cost share = $0 everywhere",
        "Indirect 8% MTDC method stated in Section F (CONFIRM)",
        "Budget narrative has a matching section for each component",
        "Object-class lines finalised — this worksheet is a starting map, not the last word",
    ]:
        story.append(p(f"• {c}", sty["body"]))

    story.append(Spacer(1, 4))
    story.append(
        p(
            "Sources: CDC-RFA-JG-26-0056 component funding instructions for SF-424A; "
            "FairBanks Year 1 ask $7.5M (C1 $3.0M + C2 $1.0M + C3 $1.5M + C4 $1.0M + C5 $1.0M); "
            "application_answers.md §F.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {target}")
    print("Form1", form1_total, "Form2", form2_total, "Grand", grand)
    for comp, sp in zip(COMPONENTS, splits):
        print(comp["id"], sp)


if __name__ == "__main__":
    build()
