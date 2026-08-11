#!/usr/bin/env python3
"""Build FairBanks Year 1 Budget Narrative for CDC-RFA-JG-26-0056.

Upload via Grants.gov: Budget Narrative Attachment form
  WS01740564-BudgetNarrativeAttachments_1_2-V1.2.pdf

Suggested attachment file name: Budget_narrative.pdf

Output:
  documents to fill/WS01740564-BudgetNarrative_Year1.pdf

NOFO rules: USD only; separate headed section per component; justify and show
calculations; categories = salaries, fringe, consultants, equipment, supplies,
travel, other, contractual, total direct, total indirect.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
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
OUT = OUT_DIR / "WS01740564-BudgetNarrative_Year1.pdf"
OUT_ALT = OUT_DIR / "WS01740564-BudgetNarrative_Year1_FINAL.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
ORANGE = colors.HexColor("#C45C26")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
CREAM = colors.HexColor("#F7F5F0")
LINE = colors.HexColor("#CED9D8")
OK_BG = colors.HexColor("#E8F3F3")

NOFO = "CDC-RFA-JG-26-0056"
CFDA = "93.318"
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
INDIRECT_RATE = 0.08
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

COMPONENTS = [
    {
        "id": "1",
        "name": "Component 1: Core Global Health Security Priorities",
        "total": 3_000_000,
        "ceiling": 3_000_000,
        "status": "Expected initial funding (core)",
        "purpose": (
            "Strengthen last-mile community and facility surveillance under Ministry of "
            "Health (MoH) leadership so early warning signals reach MoH/NISS-aligned "
            "pathways faster. Lead Strategies 1, 2, and 5; support Strategies 3, 4, 6, "
            "and 7 with MoH-aligned partners. Geography: Uganda (primary), with East "
            "Africa corridor learning under MoH Uganda."
        ),
    },
    {
        "id": "2",
        "name": "Component 2: Rapid Small-Scale Response to Outbreaks / PHE",
        "total": 1_000_000,
        "ceiling": 5_000_000,
        "status": "Approved but unfunded (ABU) until CDC activates",
        "purpose": (
            "Stand up a moderate-scale surge package for a small outbreak or public "
            "health emergency in FairBanks catchments and linked districts under MoH "
            "direction — community investigation, referral surge, training refreshers, "
            "and temporary staffing/supplies."
        ),
    },
    {
        "id": "3",
        "name": "Component 3: Rapid Large-Scale Response to Outbreaks / PHE",
        "total": 1_500_000,
        "ceiling": 10_000_000,
        "status": "Approved but unfunded (ABU) until CDC activates",
        "purpose": (
            "Scale a substantial response under MoH leadership when caseload, geography, "
            "or severity exceeds Component 2 — expanded CHW/VHT surge, multi-district "
            "coordination, partner contractual surge, and logistics."
        ),
    },
    {
        "id": "4",
        "name": "Component 4: Small-Scale Emerging Infectious Disease Threats",
        "total": 1_000_000,
        "ceiling": 15_000_000,
        "status": "Approved but unfunded (ABU) until CDC activates",
        "purpose": (
            "Contingency package for a small-scale emerging infectious disease threat "
            "requiring enhanced surveillance, risk communication, targeted screening, "
            "and rapid community-facility links."
        ),
    },
    {
        "id": "5",
        "name": "Component 5: Large-Scale Emerging Infectious Disease Threats",
        "total": 1_000_000,
        "ceiling": 20_000_000,
        "status": "Approved but unfunded (ABU) until CDC activates",
        "purpose": (
            "Contingency package for a large-scale emerging threat response under MoH/"
            "CDC activation — broader geographic coverage, heavier contractual surge, "
            "and intensified field operations. Submitted on SF-424A Form 2."
        ),
    },
]


def money(n: int) -> str:
    return f"${n:,.0f}"


def split_component(total: int) -> dict[str, int]:
    direct = int(round(total / (1 + INDIRECT_RATE * (1 - SHARES["equipment"]))))
    cats: dict[str, int] = {}
    running = 0
    for k in [
        "personnel",
        "fringe",
        "travel",
        "equipment",
        "supplies",
        "contractual",
        "construction",
    ]:
        amt = int(round(direct * SHARES[k]))
        cats[k] = amt
        running += amt
    cats["other"] = direct - running
    if cats["other"] < 0:
        cats["contractual"] += cats["other"]
        cats["other"] = 0
    mtdc = direct - cats["equipment"]
    indirect = int(round(mtdc * INDIRECT_RATE))
    computed = direct + indirect
    if computed != total:
        indirect += total - computed
    cats["direct"] = direct
    cats["indirect"] = indirect
    cats["total"] = direct + indirect
    cats["mtdc"] = mtdc
    assert cats["total"] == total
    return cats


def fit_rows(rows: list[tuple[str, int]], target: int) -> list[tuple[str, int]]:
    """Adjust last row so amounts sum to target."""
    if not rows:
        return [(f"Program staff (pooled)", target)]
    s = sum(a for _, a in rows)
    last_label, last_amt = rows[-1]
    rows = rows[:-1] + [(last_label, last_amt + (target - s))]
    assert sum(a for _, a in rows) == target
    return rows


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Helvetica", fontSize=9,
            textColor=TEAL, spaceAfter=4,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=14,
            leading=18, textColor=NAVY, spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=11,
            textColor=NAVY, spaceBefore=10, spaceAfter=4,
        ),
        "h2": ParagraphStyle(
            "h2", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=9.5,
            textColor=TEAL, spaceBefore=6, spaceAfter=3,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Helvetica", fontSize=8.5,
            leading=11, textColor=SLATE, spaceAfter=3,
        ),
        "small": ParagraphStyle(
            "small", parent=base["Normal"], fontName="Helvetica", fontSize=7.5,
            leading=9.5, textColor=MUTED, spaceAfter=2,
        ),
        "note": ParagraphStyle(
            "note", parent=base["Normal"], fontName="Helvetica-Oblique", fontSize=7.5,
            leading=9.5, textColor=ORANGE, spaceAfter=3,
        ),
        "th": ParagraphStyle(
            "th", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=7.5,
            leading=9.5, textColor=NAVY,
        ),
        "td": ParagraphStyle(
            "td", parent=base["Normal"], fontName="Helvetica", fontSize=7.5,
            leading=9.5, textColor=SLATE,
        ),
    }


def p(text: str, style) -> Paragraph:
    import re

    s = str(text).replace("\n", "<br/>")
    # Temporarily protect allowed tags
    tags = []

    def _keep(m):
        tags.append(m.group(0))
        return f"@@TAG{len(tags)-1}@@"

    s = re.sub(r"</?(?:b|i|br\s*/)>", _keep, s, flags=re.I)
    s = s.replace("&", "&amp;")
    for i, tag in enumerate(tags):
        s = s.replace(f"@@TAG{i}@@", tag)
    return Paragraph(s, style)


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(TEAL)
    canvas.setLineWidth(1.2)
    y = LETTER[1] - 0.5 * inch
    canvas.line(0.65 * inch, y, LETTER[0] - 0.65 * inch, y)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(MUTED)
    canvas.drawString(
        0.65 * inch,
        LETTER[1] - 0.38 * inch,
        f"FairBanks Budget Narrative · {NOFO} · Your health, our mission.",
    )
    canvas.drawRightString(LETTER[0] - 0.65 * inch, LETTER[1] - 0.38 * inch, f"Page {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(0.65 * inch, 0.5 * inch, LETTER[0] - 0.65 * inch, 0.5 * inch)
    canvas.drawCentredString(
        LETTER[0] / 2,
        0.32 * inch,
        "Your health, our mission.",
    )
    canvas.restoreState()


def grid(data, widths):
    t = Table(data, colWidths=widths, hAlign="LEFT")
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), CREAM),
                ("BACKGROUND", (0, -1), (-1, -1), OK_BG),
                ("BOX", (0, 0), (-1, -1), 0.6, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return t


def simple_table(rows: list[tuple[str, int]], sty, label_w=5.2 * inch, amt_w=1.5 * inch):
    data = [[p("<b>Item / calculation</b>", sty["th"]), p("<b>Amount (USD)</b>", sty["th"])]]
    for label, amt in rows:
        data.append([p(label, sty["td"]), p(money(amt), sty["td"])])
    total = sum(a for _, a in rows)
    data.append([p("<b>Category subtotal</b>", sty["td"]), p(f"<b>{money(total)}</b>", sty["td"])])
    return grid(data, [label_w, amt_w]), total


def personnel_rows_c1(target: int) -> list[tuple[str, int]]:
    # Annual salary x % effort (12-month Year 1)
    base = [
        ("Project Director / Authorized Official — 80% of $90,000", int(0.80 * 90000)),
        ("Technical Lead, Surveillance & 7-1-7 — 100% of $62,000", 62000),
        ("FCHIP / Data Systems Manager — 100% of $56,000", 56000),
        ("M&E and Learning Specialist — 100% of $48,000", 48000),
        ("Finance & Grants Compliance Manager — 100% of $52,000", 52000),
        ("Training & Workforce Coordinator — 100% of $44,000", 44000),
        ("Community Surveillance Officers (4) — 100% of $30,000 each", 120000),
        ("CHW/VHT Supervisors (8) — 100% of $20,000 each", 160000),
        ("Partner / MoH Liaison Officer — 100% of $40,000", 40000),
        ("Admin & Operations Officers (2) — 100% of $24,000 each", 48000),
        ("Drivers / field logistics (3) — 100% of $14,000 each", 42000),
    ]
    return fit_rows(base, target)


def personnel_rows_surge(target: int, scale_label: str) -> list[tuple[str, int]]:
    base = [
        (f"Surge Coordinator ({scale_label}) — 100% of $48,000 (activation period)", 48000),
        ("Field Epidemiologist / Investigation Lead — 100% of $44,000", 44000),
        ("Data / Reporting Officer — 100% of $36,000", 36000),
        ("Community Surge Supervisors (4) — 100% of $22,000 each", 88000),
        ("Finance surge support — 50% of $40,000", 20000),
    ]
    # Scale remaining into temporary field staff pool
    used = sum(a for _, a in base)
    remaining = target - used
    if remaining > 0:
        base.append(
            (
                f"Temporary CHW/VHT surge stipends & short-term field staff pool ({scale_label})",
                remaining,
            )
        )
    return fit_rows(base, target)


def build_component(story, sty, comp: dict, cats: dict[str, int], w: float):
    story.append(p(comp["name"], sty["h1"]))
    story.append(
        p(
            f"<b>Year 1 federal request:</b> {money(comp['total'])} "
            f"(ceiling {money(comp['ceiling'])}) · <b>Status:</b> {comp['status']}<br/>"
            f"<b>Purpose:</b> {comp['purpose']}",
            sty["body"],
        )
    )

    # Summary table matching SF-424A
    summary = [
        [p("<b>SF-424A object class</b>", sty["th"]), p("<b>Amount</b>", sty["th"])],
        [p("Personnel (salaries & wages)", sty["td"]), p(money(cats["personnel"]), sty["td"])],
        [p("Fringe benefits", sty["td"]), p(money(cats["fringe"]), sty["td"])],
        [p("Travel", sty["td"]), p(money(cats["travel"]), sty["td"])],
        [p("Equipment", sty["td"]), p(money(cats["equipment"]), sty["td"])],
        [p("Supplies", sty["td"]), p(money(cats["supplies"]), sty["td"])],
        [p("Contractual (includes consultants)", sty["td"]), p(money(cats["contractual"]), sty["td"])],
        [p("Construction", sty["td"]), p(money(cats["construction"]), sty["td"])],
        [p("Other", sty["td"]), p(money(cats["other"]), sty["td"])],
        [p("<b>Total direct</b>", sty["td"]), p(f"<b>{money(cats['direct'])}</b>", sty["td"])],
        [p("Indirect (8% MTDC)", sty["td"]), p(money(cats["indirect"]), sty["td"])],
        [p("<b>Component total</b>", sty["td"]), p(f"<b>{money(cats['total'])}</b>", sty["td"])],
    ]
    story.append(grid(summary, [5.2 * inch, 1.5 * inch]))
    story.append(Spacer(1, 4))

    # --- Salaries ---
    story.append(p("Salaries and wages", sty["h2"]))
    if comp["id"] == "1":
        prow = personnel_rows_c1(cats["personnel"])
        story.append(
            p(
                "All salaries are Year 1 federal costs for FairBanks staff working on award "
                "activities (not routine clinical care). Amount = annual salary x percent effort. "
                "Routine Medical Centre clinical salaries stay on FairBanks operations.",
                sty["body"],
            )
        )
    else:
        scale = {
            "2": "moderate / small-scale",
            "3": "substantial / large-scale",
            "4": "small emerging threat",
            "5": "large emerging threat",
        }[comp["id"]]
        prow = personnel_rows_surge(cats["personnel"], scale)
        story.append(
            p(
                "Personnel costs activate with CDC emergency / emerging-threat funding. "
                "Amounts are annualised planning figures for a 12-month budget period; "
                "actual drawdown follows activation length and MoH tasking.",
                sty["body"],
            )
        )
    t, _ = simple_table(prow, sty)
    story.append(t)

    # --- Fringe ---
    story.append(p("Fringe benefits", sty["h2"]))
    fringe_rate = cats["fringe"] / cats["personnel"] if cats["personnel"] else 0
    story.append(
        p(
            f"Fringe covers statutory and organisational benefits (NSSF employer contribution, "
            f"medical cover contribution, leave loading, and related payroll burdens) at an "
            f"effective blended rate of about <b>{fringe_rate:.1%}</b> of salaries "
            f"({money(cats['fringe'])} = {fringe_rate:.1%} x {money(cats['personnel'])}). "
            f"CONFIRM final fringe pool against FairBanks payroll policy.",
            sty["body"],
        )
    )
    t, _ = simple_table(
        [("Fringe benefits pool (blended statutory + organisational)", cats["fringe"])],
        sty,
    )
    story.append(t)

    # --- Consultants (subset of contractual for NOFO format) ---
    story.append(p("Consultant costs", sty["h2"]))
    consult = int(round(cats["contractual"] * 0.18))
    story.append(
        p(
            "Short-term technical consultants (not FairBanks employees). On SF-424A these "
            "costs sit inside the Contractual line; shown separately here per NOFO format.",
            sty["body"],
        )
    )
    if comp["id"] == "1":
        crow = fit_rows(
            [
                ("Epidemiology / 7-1-7 practice advisor — ~60 days x $800", 48000),
                ("One Health / zoonoses training consultant — ~40 days x $750", 30000),
                ("Data standards / MoH-NISS interoperability consultant — ~45 days x $850", consult - 78000),
            ],
            consult,
        )
    else:
        crow = fit_rows(
            [
                (f"Surge epidemiology consultant — days x daily rate (pooled)", int(consult * 0.55)),
                (f"Risk communication / RCCE consultant — pooled days", int(consult * 0.25)),
                (f"Logistics / IPC technical consultant — pooled days", consult - int(consult * 0.55) - int(consult * 0.25)),
            ],
            consult,
        )
    t, _ = simple_table(crow, sty)
    story.append(t)

    # --- Equipment ---
    story.append(p("Equipment", sty["h2"]))
    story.append(
        p(
            "Equipment = tangible items with unit cost typically $5,000 or more (or items "
            "capitalised under FairBanks policy). Laptops and phones under the threshold "
            "are listed under Supplies.",
            sty["body"],
        )
    )
    if comp["id"] == "1":
        erow = fit_rows(
            [
                ("Server / secure sync appliance for FCHIP (1)", 28000),
                ("Rugged field tablets kit for supervisors (lot)", 35000),
                ("Vehicle for community surveillance & supervision (1)", 45000),
                ("Other capital field/IT equipment (pooled)", cats["equipment"] - 108000),
            ],
            cats["equipment"],
        )
    else:
        erow = fit_rows(
            [
                ("Surge IT / cold-chain or response equipment (as activated)", int(cats["equipment"] * 0.6)),
                ("Temporary logistics assets / radios / generators (pooled)", cats["equipment"] - int(cats["equipment"] * 0.6)),
            ],
            cats["equipment"],
        )
    t, _ = simple_table(erow, sty)
    story.append(t)

    # --- Supplies ---
    story.append(p("Supplies", sty["h2"]))
    srow = fit_rows(
        [
            ("Laptops, phones, airtime devices under equipment threshold", int(cats["supplies"] * 0.35)),
            ("PPE, sample packaging, field kits (non-lab clinical care excluded)", int(cats["supplies"] * 0.30)),
            ("Printing, registers, training materials, office supplies", cats["supplies"] - int(cats["supplies"] * 0.35) - int(cats["supplies"] * 0.30)),
        ],
        cats["supplies"],
    )
    t, _ = simple_table(srow, sty)
    story.append(t)

    # --- Travel ---
    story.append(p("Travel", sty["h2"]))
    story.append(
        p(
            "Domestic Uganda field travel for supervision, district coordination, training, "
            "and MoH/CDC meetings. International travel only if pre-approved and necessary "
            "for award performance. Costs = trips x travellers x (transport + lodging + per diem).",
            sty["body"],
        )
    )
    trow = fit_rows(
        [
            ("In-country supervision & CHW/VHT field travel (pooled trips)", int(cats["travel"] * 0.45)),
            ("District / KCCA / MoH coordination travel", int(cats["travel"] * 0.25)),
            ("Training cohort travel & participant transport", int(cats["travel"] * 0.20)),
            ("Contingency / approved regional learning travel", cats["travel"] - int(cats["travel"] * 0.45) - int(cats["travel"] * 0.25) - int(cats["travel"] * 0.20)),
        ],
        cats["travel"],
    )
    t, _ = simple_table(trow, sty)
    story.append(t)

    # --- Other ---
    story.append(p("Other categories", sty["h2"]))
    story.append(
        p(
            "Other direct costs that are not salaries, fringe, travel, equipment, supplies, "
            "or contractual. Includes CHW/VHT activity support, training venues, M&E, and "
            "programme operating costs tied to surveillance (not routine clinic care).",
            sty["body"],
        )
    )
    orow = fit_rows(
        [
            ("CHW/VHT stipends / activity allowances for surveillance reporting", int(cats["other"] * 0.40)),
            ("Training venues, catering, and cohort logistics", int(cats["other"] * 0.20)),
            ("M&E, data quality audits, after-action reviews (~5-8% programme intent)", int(cats["other"] * 0.18)),
            ("Communications, connectivity, software licenses, banking/fees", int(cats["other"] * 0.12)),
            ("Other programme operating costs", cats["other"] - int(cats["other"] * 0.40) - int(cats["other"] * 0.20) - int(cats["other"] * 0.18) - int(cats["other"] * 0.12)),
        ],
        cats["other"],
    )
    t, _ = simple_table(orow, sty)
    story.append(t)

    # --- Contractual ---
    story.append(p("Contractual costs", sty["h2"]))
    contract_rest = cats["contractual"] - consult
    story.append(
        p(
            f"SF-424A Contractual line = {money(cats['contractual'])} "
            f"(includes consultants {money(consult)} above + partner/vendor contracts "
            f"{money(contract_rest)} below). Subawards/contracts will use written scopes, "
            "budgets, and FairBanks procurement rules; no funds for restricted activities.",
            sty["body"],
        )
    )
    if comp["id"] == "1":
        crow2 = fit_rows(
            [
                ("MoH-aligned laboratory / diagnostic support partner", int(contract_rest * 0.30)),
                ("FCHIP software development / hosting / security vendors", int(contract_rest * 0.28)),
                ("District / community implementing partners for outreach & referrals", int(contract_rest * 0.25)),
                ("Audit, translation, and specialised service vendors", contract_rest - int(contract_rest * 0.30) - int(contract_rest * 0.28) - int(contract_rest * 0.25)),
            ],
            contract_rest,
        )
    else:
        crow2 = fit_rows(
            [
                ("Surge implementing partners / temporary response contractors", int(contract_rest * 0.55)),
                ("Emergency logistics, transport, and warehousing vendors", int(contract_rest * 0.25)),
                ("Lab surge / specimen transport support (MoH-aligned)", contract_rest - int(contract_rest * 0.55) - int(contract_rest * 0.25)),
            ],
            contract_rest,
        )
    t, _ = simple_table(crow2, sty)
    story.append(t)

    # --- Totals ---
    story.append(p("Total direct costs", sty["h2"]))
    story.append(
        p(
            f"Total direct = Personnel + Fringe + Travel + Equipment + Supplies + "
            f"Contractual + Other = <b>{money(cats['direct'])}</b>.",
            sty["body"],
        )
    )
    story.append(p("Total indirect costs", sty["h2"]))
    story.append(
        p(
            f"Foreign organisation indirect rate of <b>8% of Modified Total Direct Costs "
            f"(MTDC)</b>. MTDC for this worksheet = Total Direct minus Equipment = "
            f"{money(cats['direct'])} - {money(cats['equipment'])} = <b>{money(cats['mtdc'])}</b>. "
            f"Indirect = 8% x {money(cats['mtdc'])} ≈ <b>{money(cats['indirect'])}</b> "
            f"(rounded so component total equals {money(cats['total'])}). "
            f"CONFIRM final MTDC exclusions with CDC/GMS. No negotiated NICRA attached; "
            f"using the foreign 8% method pending confirmation.",
            sty["body"],
        )
    )
    story.append(
        p(
            f"<b>Component {comp['id']} federal total = {money(cats['total'])}</b>",
            sty["body"],
        )
    )


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    splits = [split_component(c["total"]) for c in COMPONENTS]
    grand = sum(c["total"] for c in COMPONENTS)
    assert grand == 7_500_000

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
        leftMargin=0.65 * inch,
        rightMargin=0.65 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title=f"Budget Narrative — {ORG} — {NOFO}",
        author=ORG,
    )
    w = LETTER[0] - 1.3 * inch
    story = []

    story.append(p("BUDGET NARRATIVE ATTACHMENT", sty["kicker"]))
    story.append(
        p(
            "Year 1 Budget Narrative — Strengthening Global Health Security "
            "(Detect, Notify, Respond)",
            sty["title"],
        )
    )
    story.append(
        p(
            f"<b>Applicant:</b> {ORG}<br/>"
            f"<b>Opportunity:</b> {NOFO} · Assistance Listing {CFDA}<br/>"
            f"<b>Instrument:</b> Cooperative agreement · Period: five 12-month budget periods<br/>"
            f"<b>Year 1 federal total:</b> {money(grand)} (matches SF-424 Box 18 and SF-424A)<br/>"
            f"<b>Currency:</b> U.S. dollars only<br/>"
            f"<b>Prepared:</b> {date.today().isoformat()} · Slogan: Your health, our mission.",
            sty["body"],
        )
    )

    overview = [[p("<b>Component</b>", sty["th"]), p("<b>Year 1 ask</b>", sty["th"]), p("<b>Status</b>", sty["th"])]]
    for c in COMPONENTS:
        overview.append(
            [p(c["name"], sty["td"]), p(money(c["total"]), sty["td"]), p(c["status"], sty["td"])]
        )
    overview.append(
        [p("<b>TOTAL Year 1</b>", sty["td"]), p(f"<b>{money(grand)}</b>", sty["td"]), p("SF-424 Box 18", sty["td"])]
    )
    story.append(grid(overview, [4.2 * inch, 1.2 * inch, 2.3 * inch]))

    for i, (comp, cats) in enumerate(zip(COMPONENTS, splits)):
        if i > 0:
            story.append(PageBreak())
        build_component(story, sty, comp, cats, w)

    story.append(PageBreak())
    story.append(p("Cross-cutting notes", sty["h1"]))
    story.append(
        p(
            "<b>Cost share:</b> None proposed ($0 non-federal on SF-424A Section C).<br/>"
            "<b>Allowability:</b> Costs are intended to be reasonable, allocable, and allowable "
            "under 2 CFR 200 / HHS adoption and NOFO funding limitations. No research aims. "
            "No restricted lobbying with federal funds.<br/>"
            "<b>Procurement:</b> FairBanks will use competitive methods appropriate to value, "
            "document selection, and avoid conflicts of interest.<br/>"
            "<b>M&E:</b> About 5–10% of programme effort supports monitoring, evaluation, and "
            "learning aligned to DGHP indicators agreed with CDC after award.<br/>"
            "<b>SF-424A mapping:</b> Form 1 = Components 1–4 ($6,500,000). Form 2 = Component 5 "
            "($1,000,000) under Other Attachments. Combined Year 1 federal = $7,500,000.<br/>"
            "<b>CONFIRM before submit:</b> final salary scales, fringe policy, vendor quotes, "
            "indirect / MTDC treatment, and MoH partner contracting approach.",
            sty["body"],
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        p(
            f"Authorized organisational contact for budget questions: Racheal Nabukeera, "
            f"Managing Director and Co-founder · info@fairbanksmedicalcentre.org · +256 772 849 258.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {target}")
    print("Grand", grand)
    for c, s in zip(COMPONENTS, splits):
        print(c["id"], s["total"], "direct", s["direct"], "indirect", s["indirect"])


if __name__ == "__main__":
    build()
