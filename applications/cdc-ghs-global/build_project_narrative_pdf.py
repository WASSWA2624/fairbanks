#!/usr/bin/env python3
"""Build FairBanks Project Narrative for CDC-RFA-JG-26-0056.

Upload via: WS01740564-ProjectNarrativeAttachments_1_2-V1.2.pdf
Suggested file name: Project_narrative.pdf

NOFO format (exact scored headings, in order):
  - Background and approach (35)
  - Evaluation and performance measurement plan (30) + Data management plan
  - Organizational capacity (35)
  - Collaborations

Format: PDF, 12-pt, single-spaced, 1-inch margins, page numbers. Max 60 pages
(cover, TOC, acronyms do not count toward the 60).

Output:
  documents to fill/WS01740564-ProjectNarrative.pdf
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "documents to fill"
OUT = OUT_DIR / "WS01740564-ProjectNarrative.pdf"
OUT_ALT = OUT_DIR / "WS01740564-ProjectNarrative_FINAL.pdf"

NAVY = colors.HexColor("#0A1F2E")
TEAL = colors.HexColor("#0D6E6E")
SLATE = colors.HexColor("#1E2F38")
MUTED = colors.HexColor("#52636C")
LINE = colors.HexColor("#CED9D8")
CREAM = colors.HexColor("#F7F5F0")

NOFO = "CDC-RFA-JG-26-0056"
CFDA = "93.318"
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
SLOGAN = "Your health, our mission."


def styles():
    base = getSampleStyleSheet()
    return {
        "cover_kicker": ParagraphStyle(
            "cover_kicker", parent=base["Normal"], fontName="Times-Roman",
            fontSize=11, textColor=TEAL, alignment=TA_CENTER, spaceAfter=8,
        ),
        "cover_title": ParagraphStyle(
            "cover_title", parent=base["Normal"], fontName="Times-Bold",
            fontSize=16, leading=20, textColor=NAVY, alignment=TA_CENTER, spaceAfter=10,
        ),
        "cover_body": ParagraphStyle(
            "cover_body", parent=base["Normal"], fontName="Times-Roman",
            fontSize=12, leading=16, textColor=SLATE, alignment=TA_CENTER, spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Normal"], fontName="Times-Bold",
            fontSize=12, leading=16, textColor=NAVY, spaceBefore=14, spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "h2", parent=base["Normal"], fontName="Times-Bold",
            fontSize=12, leading=15, textColor=TEAL, spaceBefore=10, spaceAfter=6,
        ),
        "h3": ParagraphStyle(
            "h3", parent=base["Normal"], fontName="Times-Bold",
            fontSize=12, leading=15, textColor=SLATE, spaceBefore=8, spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Times-Roman",
            fontSize=12, leading=16, textColor=SLATE, alignment=TA_JUSTIFY, spaceAfter=8,
        ),
        "bullet": ParagraphStyle(
            "bullet", parent=base["Normal"], fontName="Times-Roman",
            fontSize=12, leading=16, textColor=SLATE, leftIndent=14, spaceAfter=3,
        ),
        "toc": ParagraphStyle(
            "toc", parent=base["Normal"], fontName="Times-Roman",
            fontSize=12, leading=18, textColor=SLATE, spaceAfter=4,
        ),
        "th": ParagraphStyle(
            "th", parent=base["Normal"], fontName="Times-Bold",
            fontSize=10, leading=13, textColor=NAVY,
        ),
        "td": ParagraphStyle(
            "td", parent=base["Normal"], fontName="Times-Roman",
            fontSize=10, leading=13, textColor=SLATE,
        ),
        "small": ParagraphStyle(
            "small", parent=base["Normal"], fontName="Times-Italic",
            fontSize=10, leading=13, textColor=MUTED, spaceAfter=6,
        ),
    }


def p(text: str, style) -> Paragraph:
    import re

    s = str(text).replace("\n", "<br/>")
    tags = []

    def _keep(m):
        tags.append(m.group(0))
        return f"@@TAG{len(tags)-1}@@"

    s = re.sub(r"</?(?:b|i|br\s*/)>", _keep, s, flags=re.I)
    s = s.replace("&", "&amp;")
    for i, tag in enumerate(tags):
        s = s.replace(f"@@TAG{i}@@", tag)
    return Paragraph(s, style)


def bullets(items, sty):
    return [
        p(f"• {item}", sty["bullet"]) for item in items
    ]


def table(rows, widths):
    t = Table(rows, colWidths=widths, hAlign="LEFT")
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), CREAM),
                ("BOX", (0, 0), (-1, -1), 0.6, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return t


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Times-Roman", 9)
    canvas.setFillColor(MUTED)
    if doc.page > 1:
        canvas.drawString(1.0 * inch, LETTER[1] - 0.6 * inch, f"{ORG} | {NOFO}")
        canvas.drawRightString(LETTER[0] - 1.0 * inch, LETTER[1] - 0.6 * inch, f"Page {doc.page}")
        canvas.setStrokeColor(LINE)
        canvas.setLineWidth(0.6)
        canvas.line(1.0 * inch, LETTER[1] - 0.7 * inch, LETTER[0] - 1.0 * inch, LETTER[1] - 0.7 * inch)
    canvas.drawCentredString(LETTER[0] / 2, 0.55 * inch, SLOGAN)
    canvas.restoreState()


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
        leftMargin=1.0 * inch,
        rightMargin=1.0 * inch,
        topMargin=1.0 * inch,
        bottomMargin=1.0 * inch,
        title=f"Project Narrative - {ORG} - {NOFO}",
        author=ORG,
    )
    story = []
    usable = LETTER[0] - 2.0 * inch

    # -------- Cover (does not count toward 60) --------
    story.append(Spacer(1, 0.8 * inch))
    story.append(p("PROJECT NARRATIVE", sty["cover_kicker"]))
    story.append(
        p(
            "Strengthening Global Health Security by improving public health capacity "
            "to detect, notify, and respond to disease outbreaks globally",
            sty["cover_title"],
        )
    )
    story.append(p(f"Opportunity number: {NOFO}", sty["cover_body"]))
    story.append(p(f"Assistance Listing: {CFDA}", sty["cover_body"]))
    story.append(Spacer(1, 0.3 * inch))
    story.append(p(f"<b>{ORG}</b>", sty["cover_body"]))
    story.append(p("Uganda-registered local partner | Kampala", sty["cover_body"]))
    story.append(p(SLOGAN, sty["cover_body"]))
    story.append(Spacer(1, 0.35 * inch))
    story.append(
        p(
            "Year 1 federal request: USD 7,500,000<br/>"
            "Component 1 (core): USD 3,000,000 | Components 2-5 (ABU): USD 4,500,000<br/>"
            "Primary country: Uganda | Period of performance: five 12-month budget periods",
            sty["cover_body"],
        )
    )
    story.append(Spacer(1, 0.4 * inch))
    story.append(p(f"Prepared: {date.today().isoformat()}", sty["cover_body"]))
    story.append(p("Contact: info@fairbanksmedicalcentre.org | +256 772 849 258", sty["cover_body"]))
    story.append(PageBreak())

    # -------- TOC (does not count) --------
    story.append(p("Table of contents", sty["h1"]))
    for line in [
        "Background and approach",
        "    Problem and context",
        "    Evidence-based activities and technical approach",
        "    Complementing ongoing response",
        "    Methodology by NOFO strategy",
        "    Year 1 Component 1 work plan",
        "    Components 2-5 contingency work plans",
        "    Countries and delivery location",
        "    Anticipated outcomes",
        "Evaluation and performance measurement plan",
        "    Performance measures and reporting",
        "    Use of findings and dissemination",
        "    Financial and programmatic reporting system",
        "    Data management plan",
        "Organizational capacity",
        "Collaborations",
        "Acronym list",
    ]:
        story.append(p(line, sty["toc"]))
    story.append(PageBreak())

    # -------- Acronyms (does not count) --------
    story.append(p("Acronym list", sty["h1"]))
    acr = [
        ("ABU", "Approved but unfunded"),
        ("AAR", "After-action review"),
        ("CHW", "Community Health Worker"),
        ("DGHP", "Division of Global Health Protection"),
        ("EPMP", "Evaluation and Performance Measurement Plan"),
        ("FCHIP", "FairBanks Community Health Intelligence Platform"),
        ("GHS", "Global Health Security"),
        ("IPC", "Infection prevention and control"),
        ("KCCA", "Kampala Capital City Authority"),
        ("MoH", "Ministry of Health"),
        ("NISS", "National Integrated Surveillance System"),
        ("NOFO", "Notice of Funding Opportunity"),
        ("PHE", "Public health emergency"),
        ("VHT", "Village Health Team"),
        ("7-1-7", "Detect within 7 days, notify within 1 day, respond within 7 days"),
    ]
    for a, b in acr:
        story.append(p(f"<b>{a}</b>: {b}", sty["toc"]))
    story.append(PageBreak())

    # ============================================================
    # BACKGROUND AND APPROACH (35)
    # ============================================================
    story.append(p("Background and approach", sty["h1"]))

    story.append(p("Problem and context", sty["h2"]))
    story.append(
        p(
            f"{ORG} (FairBanks) is a Uganda-registered health organisation in Kampala. "
            "We operate a licensed medical centre and FairBanks Community Reach with "
            "Community Health Workers (CHWs) and Village Health Teams (VHTs) in Bukoto, "
            "Kyebando, Kisaasi, Kamwokya, Kikaaya, and nearby peri-urban communities. "
            f"Our slogan is {SLOGAN}",
            sty["body"],
        )
    )
    story.append(
        p(
            "Uganda faces repeated infectious disease threats shaped by geography, rapid "
            "urban growth, dense peri-urban settlements, and high population movement along "
            "East African corridors. Outbreaks that start in communities can grow before "
            "national systems see a clear signal. Last-mile gaps remain practical and urgent: "
            "community reports, facility encounters, and laboratory referrals often sit in "
            "separate paper or digital silos; district and national teams receive incomplete "
            "or delayed information; and surge capacity at subnational level is uneven when "
            "caseloads rise.",
            sty["body"],
        )
    )
    story.append(
        p(
            "In FairBanks catchments, CHWs and VHTs already visit homes and schools, and the "
            "Medical Centre sees patients every day. Too often, those signals do not reach "
            "formal surveillance pathways quickly enough for early action. Uganda’s progress "
            "on global health security, including integrated surveillance such as the National "
            "Integrated Surveillance System (NISS), creates a clear opportunity: strengthen "
            "the community-to-facility-to-district link so early warning reaches decision-makers "
            "before a local cluster becomes a wider emergency.",
            sty["body"],
        )
    )
    story.append(
        p(
            "This cooperative agreement will not fund routine clinical care. FairBanks will "
            "continue clinic operations on its own resources. Award funds will strengthen "
            "surveillance, data quality, workforce readiness, referral follow-up, and "
            "coordination under Ministry of Health (MoH) leadership. This helps communities "
            "recognise threats earlier, connect signals into public systems, act in time, "
            "and contain risk closer to where people live.",
            sty["body"],
        )
    )

    story.append(p("Evidence-based activities and technical approach", sty["h2"]))
    story.append(
        p(
            "Our approach is grounded in FairBanks Community Reach, the operating cascade "
            "from community members to CHWs/VHTs, Community Reach programmes, the Medical "
            "Centre, research and skills, and economic empowerment, with a continuous "
            "Data and Feedback loop. The FairBanks Community Health Intelligence Platform "
            "(FCHIP) is the intelligence component on that loop. FCHIP supports structured "
            "mobile capture (including offline use), supervisory review, simple GIS views, "
            "and MoH/NISS-aligned exports. It is designed to feed government systems, not "
            "replace them.",
            sty["body"],
        )
    )
    story.append(
        p(
            "Year 1 activities are evidence-informed, realistic for Uganda peri-urban and "
            "district settings, and measurable. We apply time-bound practice consistent with "
            "7-1-7 principles: detect community and facility signals sooner, notify through "
            "agreed channels within one day of confirmation thresholds, and support early "
            "response within a defined window. We combine: (1) CHW/VHT structured reporting; "
            "(2) facility flag and referral tracking; (3) weekly data quality checks; "
            "(4) district dashboards and joint reviews; (5) workforce training and surge "
            "rosters with One Health awareness; and (6) contingency playbooks for Components "
            "2-5 when CDC activates emergency funding.",
            sty["body"],
        )
    )

    story.append(p("Complementing ongoing response", sty["h2"]))
    story.append(
        p(
            "FairBanks will build on, not duplicate, MoH, district/KCCA, and other CDC-funded "
            "partner work. We will map existing surveillance and response actors in Year 1 "
            "Quarter 1, share work plans with MoH and CDC activity managers, and adjust "
            "geography or tasks if DGHP priorities change. Laboratory surge and selected "
            "specialised modules will be delivered with MoH-aligned partners. Software, "
            "standard operating procedures (SOPs), and training materials developed under "
            "the award will be available to MoH and CDC for appropriate use.",
            sty["body"],
        )
    )

    story.append(p("Methodology by NOFO strategy", sty["h2"]))
    story.append(
        p(
            "FairBanks will lead Strategies 1, 2, and 5 in Year 1, and support Strategies "
            "3, 4, 6, and 7 with MoH-aligned partners. The table summarises the Year 1 method.",
            sty["body"],
        )
    )
    strat = [
        [p("<b>Strategy</b>", sty["th"]), p("<b>FairBanks Year 1 method</b>", sty["th"])],
        [
            p("1. Surveillance / systems with governments", sty["td"]),
            p(
                "FCHIP capture; data quality rules; MoH/NISS-aligned exports; simple GIS early warning.",
                sty["td"],
            ),
        ],
        [
            p("2. Interoperable public health data", sty["td"]),
            p(
                "Integrate community, facility, and lab signals; weekly reporting; district dashboards.",
                sty["td"],
            ),
        ],
        [
            p("3. Public health programmes / service links", sty["td"]),
            p(
                "Community-facility links for priority diseases (e.g., malaria/fever clusters, VHFs, mpox, TB, HIV, immunisation campaigns) under MoH guidance.",
                sty["td"],
            ),
        ],
        [
            p("4. Assessment and corrective action", sty["td"]),
            p(
                "Joint district reviews; after-action reviews (AARs) after drills or real events; continuous improvement.",
                sty["td"],
            ),
        ],
        [
            p("5. Workforce", sty["td"]),
            p(
                "CHW/VHT and facility training; supervised practice; surge roster; One Health-aware sessions.",
                sty["td"],
            ),
        ],
        [
            p("6. Networks and surge connect", sty["td"]),
            p(
                "Partner coordination; MoH/CDC information sharing; activation pathways for Components 2-5.",
                sty["td"],
            ),
        ],
        [
            p("7. Small/large outbreak response", sty["td"]),
            p(
                "Components 2-5 contingency work plans and budgets for moderate/substantial PHE and emerging threats.",
                sty["td"],
            ),
        ],
    ]
    story.append(table(strat, [2.3 * inch, usable - 2.3 * inch]))
    story.append(Spacer(1, 8))

    # Component 1 work plan
    story.append(p("Year 1 Component 1 work plan (Core GHS, USD 3,000,000)", sty["h2"]))
    story.append(
        p(
            "Component 1 is expected initial funding. It strengthens last-mile detection, "
            "notification, and early response under MoH leadership.",
            sty["body"],
        )
    )
    c1 = [
        [p("<b>Quarter</b>", sty["th"]), p("<b>Key activities and milestones</b>", sty["th"])],
        [
            p("Q1", sty["td"]),
            p(
                "Staffing and onboarding; MoH/district kickoff; CHW/VHT cohort mapping; FCHIP form "
                "finalisation; baseline 7-1-7 timing measure; partner non-duplication map; "
                "first supervisory visits.",
                sty["td"],
            ),
        ],
        [
            p("Q2", sty["td"]),
            p(
                "Weekly community reporting live in primary catchments; facility flag and referral "
                "tools in use; first MoH/NISS-aligned export test; training wave 1 (surveillance "
                "and One Health awareness); district dashboard prototype.",
                sty["td"],
            ),
        ],
        [
            p("Q3", sty["td"]),
            p(
                "Expand supervised CHW/VHT coverage; second export/interoperability test; joint "
                "district review; surge roster exercise / drill 1 with AAR; strengthen referral "
                "outcome documentation for priority signals.",
                sty["td"],
            ),
        ],
        [
            p("Q4", sty["td"]),
            p(
                "Consolidate data quality targets; training wave 2; drill or real-event AAR 2; "
                "handover package of SOPs/tools for MoH/CDC use; Year 1 performance report; "
                "continuation planning.",
                sty["td"],
            ),
        ],
    ]
    story.append(table(c1, [0.7 * inch, usable - 0.7 * inch]))
    story.append(Spacer(1, 8))

    story.append(p("Components 2-5 contingency work plans (ABU)", sty["h2"]))
    story.append(
        p(
            "Components 2-5 remain approved but unfunded until CDC activates emergency or "
            "emerging-threat funding. Each has a scaled Year 1 budget and a ready activation "
            "sequence under MoH direction. FairBanks will not draw these funds without CDC "
            "activation and agreed tasking.",
            sty["body"],
        )
    )
    abu = [
        [p("<b>Component</b>", sty["th"]), p("<b>Year 1 ask</b>", sty["th"]), p("<b>Activation focus</b>", sty["th"])],
        [
            p("2. Small-scale outbreak / PHE", sty["td"]),
            p("USD 1,000,000", sty["td"]),
            p(
                "Moderate surge: rapid community investigation, referral surge, refresher training, temporary staffing and supplies in affected catchments/districts.",
                sty["td"],
            ),
        ],
        [
            p("3. Large-scale outbreak / PHE", sty["td"]),
            p("USD 1,500,000", sty["td"]),
            p(
                "Substantial surge: multi-district coordination, expanded CHW/VHT surge, heavier contractual logistics, intensified reporting to MoH/CDC.",
                sty["td"],
            ),
        ],
        [
            p("4. Small emerging threats", sty["td"]),
            p("USD 1,000,000", sty["td"]),
            p(
                "Enhanced surveillance, risk communication, targeted screening, and rapid community-facility links for an emerging threat at limited scale.",
                sty["td"],
            ),
        ],
        [
            p("5. Large emerging threats", sty["td"]),
            p("USD 1,000,000", sty["td"]),
            p(
                "Broader geographic coverage and intensified field/contractual surge for a large emerging threat under MoH/CDC activation (SF-424A Form 2).",
                sty["td"],
            ),
        ],
    ]
    story.append(table(abu, [1.8 * inch, 1.2 * inch, usable - 3.0 * inch]))
    story.append(Spacer(1, 6))
    story.append(
        p(
            "Activation sequence (all ABU components): (1) CDC/MoH activation notice; "
            "(2) 48 to 72 hour internal surge stand-up; (3) deploy rostered supervisors and "
            "CHWs/VHTs; (4) open daily situation reporting; (5) execute investigation, "
            "referral, risk communication, and logistics tasks; (6) AAR and cost reporting "
            "within agreed timelines.",
            sty["body"],
        )
    )

    story.append(p("Countries and delivery location", sty["h2"]))
    story.append(
        p(
            "<b>Year 1 primary country:</b> Uganda. FairBanks is incorporated and operating "
            "in Uganda (Company No. 80020003843337; TIN 1053370026; NSSF NS043295), with "
            "principal place of business in Kampala and public community base at "
            "Kyebando-Kisalosalo, Northern Bypass, Kampala.",
            sty["body"],
        )
    )
    story.append(
        p(
            "<b>Regional learning:</b> East Africa corridor learning and cross-border signal "
            "protocols under MoH Uganda leadership. We do not claim unregistered field "
            "operations outside Uganda. Additional countries will be included only with "
            "documented partner letters and registration (CONFIRM before naming).",
            sty["body"],
        )
    )

    story.append(p("Anticipated outcomes", sty["h2"]))
    story.extend(
        bullets(
            [
                "Faster detection and response using 7-1-7 style timing from community signal to district notification.",
                "Stronger weekly community and facility reporting, with successful MoH/NISS-aligned export tests.",
                "Trained CHWs/VHTs and an exercised surge roster.",
                "Improved referral and follow-up documentation for priority disease signals.",
                "Practical tools, SOPs, and learning available for MoH and CDC use.",
                "Contingency capacity (Components 2-5) ready for activation without creating parallel systems.",
            ],
            sty,
        )
    )

    story.append(PageBreak())

    # ============================================================
    # EPMP (30)
    # ============================================================
    story.append(p("Evaluation and performance measurement plan", sty["h1"]))
    story.append(
        p(
            "FairBanks will implement an Evaluation and Performance Measurement Plan (EPMP) "
            "aligned to this NOFO’s logic-model outcomes and DGHP FY26 partner-level "
            "indicators (surveillance, community mitigation, emergency operations, IPC, and "
            "laboratory support as applicable). About 5-10% of funds support monitoring, "
            "evaluation, and learning. A full detailed EPMP will be finalised with CDC within "
            "six months of award, as required.",
            sty["body"],
        )
    )

    story.append(p("Performance measures and reporting", sty["h2"]))
    story.append(
        p(
            "Year 1 project measures (draft targets, to be confirmed with CDC after award):",
            sty["body"],
        )
    )
    measures = [
        [p("<b>Measure</b>", sty["th"]), p("<b>Year 1 target</b>", sty["th"]), p("<b>Frequency</b>", sty["th"])],
        [
            p("Median days from community signal to district notification", sty["td"]),
            p("Baseline in Q1; improve toward 7-1-7 by Q4", sty["td"]),
            p("Monthly / quarterly", sty["td"]),
        ],
        [
            p("Active CHWs/VHTs with complete weekly reports", sty["td"]),
            p("At least 80% by Q4", sty["td"]),
            p("Weekly / quarterly", sty["td"]),
        ],
        [
            p("Priority disease referrals documented with outcome", sty["td"]),
            p("At least 70%", sty["td"]),
            p("Monthly / quarterly", sty["td"]),
        ],
        [
            p("Successful MoH/NISS-aligned export tests", sty["td"]),
            p("At least 2", sty["td"]),
            p("By event / quarterly", sty["td"]),
        ],
        [
            p("After-action reviews (drills or real events)", sty["td"]),
            p("At least 2", sty["td"]),
            p("After each event", sty["td"]),
        ],
        [
            p("Training cohorts completed (CHW/VHT and frontline)", sty["td"]),
            p("At least 2 waves", sty["td"]),
            p("Quarterly", sty["td"]),
        ],
    ]
    story.append(table(measures, [2.6 * inch, 2.2 * inch, usable - 4.8 * inch]))
    story.append(Spacer(1, 8))
    story.append(
        p(
            "Each programme milestone in the Component 1 work plan and each ABU activation "
            "will have linked indicators in financial and programmatic reports. Indicator "
            "definitions, numerators/denominators, and data sources will be locked with CDC "
            "in the detailed EPMP.",
            sty["body"],
        )
    )

    story.append(p("Use of findings and dissemination", sty["h2"]))
    story.append(
        p(
            "FairBanks will use evaluation findings and performance data to: (1) coach "
            "CHW/VHT supervisors; (2) fix data-quality gaps; (3) rebalance outreach and "
            "referral support; (4) improve SOPs and training; and (5) brief MoH, districts/KCCA, "
            "and CDC on what is working. Results will be disseminated through quarterly "
            "review meetings, written reports, AARs, and shared dashboards. These show "
            "the value of last-mile surveillance that feeds government systems.",
            sty["body"],
        )
    )

    story.append(p("Financial and programmatic reporting system", sty["h2"]))
    story.append(
        p(
            "FairBanks will operate a grants finance and programme tracking system able to "
            "generate: (1) quarterly financial reports showing fund disbursement by "
            "component and object class; (2) quarterly programmatic reports with indicator "
            "tables and narrative progress; (3) activation-specific burn and output reports "
            "for Components 2-5 if funded; and (4) annual performance / continuation inputs "
            "as required. Segregation of duties, payment approvals, and documentation will "
            "support timely administration of personnel, grants, and contracts.",
            sty["body"],
        )
    )

    story.append(p("Data management plan", sty["h2"]))
    story.append(
        p(
            "This initial Data Management Plan (DMP) describes how FairBanks will manage "
            "award data. A final DMP will be completed with CDC after award.",
            sty["body"],
        )
    )
    story.extend(
        bullets(
            [
                "Data types: CHW/VHT visit and symptom forms; facility flags; referral outcomes; training records; drill/event AARs; de-identified aggregate surveillance exports.",
                "Systems: FCHIP offline-capable mobile capture with secure sync; facility/programme dashboard; MoH-aligned export formats. No parallel patient silo replacing MoH systems.",
                "Privacy and security: consent-aware collection where required; role-based least-privilege access; anonymisation for analytics; encryption in transit for sync; audit logs for administrative access.",
                "Quality: validation rules at entry; supervisory review; monthly data-quality checks; documented corrections.",
                "Sharing and retention: MoH and CDC receive agreed exports and SOPs; retention follows award terms and applicable Uganda/CDC requirements.",
                "Ownership: tools and documentation developed under the award will be available to MoH and CDC for appropriate use.",
            ],
            sty,
        )
    )

    story.append(PageBreak())

    # ============================================================
    # ORGANIZATIONAL CAPACITY (35)
    # ============================================================
    story.append(p("Organizational capacity", sty["h1"]))
    story.append(
        p(
            f"{ORG} is a Uganda company (No. 80020003843337; TIN 1053370026; NSSF NS043295) "
            "with principal place of business in Kampala. FairBanks brings a live Medical "
            "Centre and Pharmacy; active Community Reach with CHWs/VHTs; and a working FCHIP "
            "MVP (mobile capture, supervisory workflows, GIS roadmap, and a secure path for "
            "facility data exchange). Public community base: Kyebando-Kisalosalo, Northern "
            "Bypass, Kampala. Catchments include Bukoto, Kyebando, Kisaasi, Kamwokya, and Kikaaya.",
            sty["body"],
        )
    )
    story.append(
        p(
            "Leadership: Racheal Nabukeera, Managing Director and Co-founder, with more than 15 "
            "years in Uganda private health leadership; MA Social Sector Planning (Makerere); "
            "doctoral studies in progress; links through Uganda Healthcare Federation networks. "
            "Day-to-day award management will add a technical lead for surveillance, FCHIP/data "
            "manager, M&E specialist, finance and grants compliance manager, training "
            "coordinator, community surveillance officers, and CHW/VHT supervisors (see budget "
            "narrative and staffing attachments).",
            sty["body"],
        )
    )
    story.append(
        p(
            "<b>Local experience and institutional capacity.</b> FairBanks already delivers "
            "community outreach, maternal and child health touchpoints, chronic-disease "
            "screening, school health education, and facility care in the same communities "
            "where surveillance signals arise. Field relationships plus a clinical "
            "anchor form the operational base for GHS last-mile work under MoH "
            "leadership. We will document in-country local partner preference criteria in "
            "the required attachment package.",
            sty["body"],
        )
    )
    story.append(
        p(
            "<b>Operational capacity in Uganda.</b> FairBanks is locally registered, "
            "Luganda/English capable in community settings, and staffed for peri-urban Kampala "
            "operations. Multi-country expansion is not claimed for Year 1 field delivery; "
            "corridor learning stays under MoH Uganda. Additional countries require partner "
            "LOIs and registration before inclusion.",
            sty["body"],
        )
    )
    story.append(
        p(
            "<b>Financial and administrative capacity.</b> FairBanks will strengthen federal "
            "award controls for this cooperative agreement: chart of accounts by component, "
            "procurement documentation, timesheets for effort reporting, bank reconciliations, "
            "and timely financial close. We are honest that FairBanks has not yet been prime "
            "on a multi-million CDC GHS award; the Year 1 plan therefore invests in finance "
            "compliance capacity, clear SOPs, and partner contracting discipline. Supporting "
            "attachments (experience statement; financial capability statement; resumes/JDs; "
            "organizational chart; locally registered organisation documents) will provide "
            "evidence for review.",
            sty["body"],
        )
    )

    # ============================================================
    # COLLABORATIONS
    # ============================================================
    story.append(p("Collaborations", sty["h1"]))
    story.append(
        p(
            "If funded, FairBanks will collaborate with CDC programmes and CDC-funded "
            "organisations, and with organisations not funded by CDC, to avoid duplication "
            "and to work under host-government leadership.",
            sty["body"],
        )
    )
    story.append(p("CDC programmes and CDC-funded organisations", sty["h2"]))
    story.extend(
        bullets(
            [
                "Coordinate with CDC Uganda / DGHP activity managers on work plans, indicators, and activation of Components 2-5.",
                "Share geographic and technical maps with other CDC-funded partners to prevent overlap in the same catchments.",
                "Participate in partner coordination meetings and information-sharing protocols set by CDC and MoH.",
            ],
            sty,
        )
    )
    story.append(p("Host government and other stakeholders", sty["h2"]))
    story.extend(
        bullets(
            [
                "Ministry of Health: leadership for surveillance pathways, NISS-aligned exports, and surge tasking.",
                "District health teams and KCCA: weekly reporting, joint reviews, referral follow-up, and local surge coordination.",
                "Community structures: CHWs/VHTs, local leaders, schools, and community spaces for education and early warning.",
                "MoH-aligned laboratory and implementing partners: contractual support for diagnostics and specialised surge modules.",
                "Civil society, academia, and private providers: as needed for training, evaluation, or non-duplicative service links.",
            ],
            sty,
        )
    )
    story.append(
        p(
            "Collaboration principle: FairBanks feeds government systems and community "
            "response. We do not create a parallel surveillance authority. Tools and learning "
            "from the award remain available for MoH and CDC use.",
            sty["body"],
        )
    )

    story.append(Spacer(1, 16))
    story.append(
        p(
            "End of project narrative. Supporting evidence is provided in separate "
            "attachments as required by the NOFO.",
            sty["small"],
        )
    )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Wrote {target}")


if __name__ == "__main__":
    build()
