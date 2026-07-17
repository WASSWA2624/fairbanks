#!/usr/bin/env python3
"""
U.S. Dept of State — Uganda Health System MOU (DFOP0017890) — Statement of Interest.

Win-win Phase 1 SOI under Addendum E (Uganda), Track 2, Objective 4 digital health.
Run: python applications/dos-uganda/build_dos-uganda_docs.py
"""

from pathlib import Path

PROJECT = Path(__file__).resolve().parent
REPO = PROJECT.parents[1]
ASSETS = REPO / "assets"
OUT = PROJECT / "documents"
SLUG = "dos-uganda"

OUT_DOC = OUT / f"{SLUG}_word.docx"
OUT_PDF = OUT / f"{SLUG}_pdf.pdf"
OUT_PPT = OUT / f"{SLUG}_ppt.pptx"

NAVY, TEAL, ACCENT = "0A1F2E", "0D6E6E", "C45C26"
SLATE, MUTED, CREAM, LINE = "1E2F38", "3A4A54", "F7F5F0", "D0DCDC"
SLOGAN = "Your health, our mission."

PHOTOS = {
    "cover": "cover_hero_cinematic.jpg",
    "architecture": "data_flow_iso_labeled.png",
    "dashboard": "dashboard_demo.png",
    "gis": "gis_hotspots.png",
    "facility": "facility_exterior_sign.jpeg",
    "outreach": "outreach_outdoor_clinic.jpeg",
    "mobile": "outreach_mobile_phone_demo_01.jpg",
    "pharmacy": "pharmacy_staff_laptop_01.jpg",
}

# ---------------------------------------------------------------------------
# Critical details (cover table — does not count toward 5-page narrative)
# ---------------------------------------------------------------------------

PROJECT_TITLE = (
    "Connecting Community Health Visits to Uganda's National Digital Health Systems "
    "(FCHIP)"
)
ORG_NAME = "FairBanks Medical Centre"
ADDENDUM = (
    "Addendum E: Health Foreign Assistance Memorandum of Understanding "
    "(MOU) Implementation in Uganda"
)
COUNTRY = "Uganda"
FEDERAL_SHARE = "$1,850,000"
COST_SHARE = "$0 (none proposed at this stage)"
PROJECT_LENGTH = "36 months"

PROGRAMME = "U.S. Department of State / GHSD — Advancing Global Health (DFOP0017890)"
DOC_TITLE = "Statement of Interest — Phase 1"
SUBTITLE = (
    "Track 2, Objective 4: Strengthen and Scale Interoperable, "
    "Fit-for-Purpose Digital Health Systems"
)
ORG_LINE = "FairBanks Medical Centre  ·  FairBanks Community Reach  ·  FCHIP"

CRITICAL_DETAILS = [
    ["Proposed Project Title", PROJECT_TITLE],
    ["Name of the Organization", ORG_NAME],
    ["Addendum to which the SOI is responding", ADDENDUM],
    ["Target Benefiting Country", COUNTRY],
    ["Total Federal Share Requested", FEDERAL_SHARE],
    ["Total Cost Share (as applicable)", COST_SHARE],
    ["Project Length", PROJECT_LENGTH],
]

# ---------------------------------------------------------------------------
# Narrative content — win-win SOI (APS Phase 1 structure)
# ---------------------------------------------------------------------------

ISSUE = [
    "Uganda is moving through a five-year health transition (April 2026 to December "
    "2030). Clinics, districts, and communities must keep HIV, TB, malaria, maternal "
    "and child health, immunisation, and outbreak services running well while more "
    "work shifts to government-led programmes. That only works if leaders can see "
    "what is happening in communities in time to act - not weeks later from paper books.",
    "Community Health Workers and Village Health Teams already visit homes and send "
    "people for care. Too often, what they record never reaches eCHIS, clinic records, "
    "DHIS2/eIDSR, or the National Data Warehouse in a form people can use. Then "
    "national systems miss rising fever cases, missed antenatal visits, gaps in "
    "immunisation, and medicine pressure. Care stays reactive, follow-up is weak, and "
    "hard-won gains can slip.",
    "There is a clear chance to fix this: help community visit data flow into Uganda's "
    "National Digital Health Architecture (NDHA) so Ministry of Health systems are "
    "fuller and more useful, and so government can take over step by step - without "
    "building a separate parallel system.",
]

SOLUTION = [
    "FairBanks Medical Centre and FairBanks Community Reach submit this Statement of "
    "Interest under Track 2, Objective 4 of Addendum E: strengthen and scale digital "
    "health systems that work together, fit Uganda's needs, and move toward government "
    "ownership and steady day-to-day running.",
    "We want to trial, then hand over, a practical tool - the FairBanks Community "
    "Health Intelligence Platform (FCHIP) - that lets CHWs and VHTs record household "
    "and visit information on phones even offline, sync it safely, match it to national "
    "reporting needs, and help the clinic and district act. FCHIP is not meant to replace "
    "Uganda's national systems. It is a bridge so community work shows up where it "
    "should and helps prevent problems earlier.",
    "We will focus on the links that matter most for our catchments: household and "
    "visit data in line with eCHIS; referral and visit handoff into clinic systems "
    "(such as eAfya or ClinicMaster where used); summary reporting into DHIS2 and "
    "alerts that support eIDSR; readiness to feed the National Data Warehouse; and "
    "CHW/VHT workforce lists that can align with iHRIS where the rules allow. We will "
    "follow NDHA and OpenHIE-style sharing rules, under Ministry of Health guidance.",
]

WIN_WIN_INTRO = (
    "We want every partner to gain something lasting. No one should depend forever "
    "on a private tool that sits outside government systems."
)

WIN_WIN = [
    [
        "Uganda / Ministry of Health",
        "More complete community data in eCHIS, DHIS2/eIDSR, and related national "
        "systems; clearer proof that referrals were closed; and a written plan to "
        "hand over staff roles and running costs to government by the end of the project.",
    ],
    [
        "U.S. Department of State / GHSD",
        "A Ugandan partner with a working clinic and community network that keeps "
        "essential services going while showing digital systems can connect under the "
        "MOU - in ways that can be measured and handed on, supporting U.S. goals around "
        "health security and country self-reliance.",
    ],
    [
        "Districts & facilities",
        "Simple screens to plan outreach, follow referrals, and track programmes; "
        "maps that show hotspots so teams can act before outbreaks hit the ward.",
    ],
    [
        "Communities & CHWs/VHTs",
        "Easy tools that work offline, quicker support from the medical centre, and "
        "care that reaches mothers, children, older people, and people with long-term "
        "illness earlier.",
    ],
    [
        "FairBanks / FCHIP",
        "A real place to test and write down a working data-sharing approach, then "
        "share the guides, training, and handoff plans with Ministry-aligned partners.",
    ],
]

ACTIVITIES = [
    "A1. Put simple offline phone tools in the hands of CHWs and VHTs so household "
    "and visit records match eCHIS ideas across FairBanks Community Reach areas.",
    "A2. Set up safe syncing and clear access rules so community records reach "
    "FairBanks Medical Centre work and authorised district viewers, in line with "
    "Uganda's data protection and security expectations.",
    "A3. Match community indicators to DHIS2 reports and show how summary data can "
    "move for maternal and child health, immunisation, chronic disease screening, "
    "and disease watch (including eIDSR-related alerts).",
    "A4. Give clinics and programmes clear dashboards and hotspot maps that help "
    "trigger outreach, referrals, and supply planning.",
    "A5. Write the sharing rules, data-handling guides, training materials, and a "
    "step-by-step government handoff plan (including staff roles and running costs) "
    "ready for a full proposal if we are invited.",
]

APPROACH = [
    "We start from how FairBanks already works: communities name needs; CHWs and "
    "VHTs visit, teach, and refer; FairBanks Community Reach runs outreach and home "
    "visits; FairBanks Medical Centre provides clinic care, tests, pharmacy, and "
    "referrals; FCHIP helps the information move back and forth so teams can act.",
    "We will build for sharing, not for locking data away. Priority links are eCHIS, "
    "DHIS2/eIDSR, clinic record systems (eAfya or ClinicMaster where used), and the "
    "National Data Warehouse when the national formats are confirmed.",
    "Digital work must support, not distract from, life-saving care. Community forms "
    "will include HIV, TB, and malaria flags where useful, plus maternal and child "
    "pathways, immunisation follow-up, and early outbreak warning.",
    "We will shape forms and alerts with CHWs, clinic staff, and local leaders; train "
    "Ugandan staff who will keep the system going; and write every link clearly for "
    "Ministry of Health review.",
    "We start where FairBanks already works - Bukoto, Kyebando, Kisaasi, Kamwokya, "
    "Kikaaya and nearby communities - and design so other districts with CHW/VHT "
    "networks can copy the approach later.",
]

OUTCOMES = [
    [
        "Save lives",
        "Earlier notice of missed antenatal visits, immunisation gaps, fever clusters, "
        "and lost referrals - followed by CHW action and clinic support.",
    ],
    [
        "Strengthen systems",
        "Working links and sample exports for eCHIS-style records and DHIS2/eIDSR "
        "summaries, with clear guides for the Ministry and partners.",
    ],
    [
        "Work more efficiently",
        "Less double entry and less waiting on paper; screens that help teams move "
        "from a community signal to clinic or district action faster.",
    ],
    [
        "Build self-reliance",
        "A handoff package with guides, standard procedures, training, hosting "
        "options, and a costed path to move running costs onto domestic systems by "
        "project close.",
    ],
    [
        "Support U.S. interest",
        "A clear contribution to the MOU digital shift: community data that connects "
        "under government authority, protects past U.S. health investments, and "
        "strengthens early warning for outbreaks.",
    ],
]

CAPACITY = [
    "We are a Ugandan organisation with a working medical centre - clinic care, "
    "tests, pharmacy, and referrals already running every day.",
    "FairBanks Community Reach is active in maternal and child health, Gericare, "
    "chronic disease screening, and school and workplace outreach.",
    "We already work with CHWs and VHTs in the Kampala-area communities named "
    "above. This is not starting from zero.",
    "We already keep digital clinic records and pharmacy dispensing data that support "
    "day-to-day facility work.",
    "Research, ethics, and skills training are part of how FairBanks works with "
    "communities and partners.",
]

INTEROP_ROWS = [
    ["eCHIS", "Household lists, visit logs, and referral flags from phone tools that work offline"],
    ["Clinic records (eAfya / ClinicMaster)", "Referral handoff and visit links at FairBanks Medical Centre"],
    ["DHIS2 / eIDSR", "Summary reports and disease alerts for maternal/child health, immunisation, outbreaks"],
    ["National Data Warehouse", "Standard data feeds when the national format is confirmed"],
    ["iHRIS", "CHW/VHT worker lists aligned where the national rules allow"],
]

PRINCIPLES = [
    ["Sustainability & country ownership", "Build with MoH/NDHA rules and plan a clear handoff over time"],
    ["Protect essential services", "Digital tools support HIV, TB, malaria, maternal/child care, and immunisation"],
    ["Data for action & accountability", "Good data into government systems that people actually use to decide"],
    ["Integrated, people-centred care", "One community-to-clinic loop; less fragmented work"],
    ["Local partnerships", "Government, CHWs/VHTs, community groups, schools, and NGO partners"],
    ["Evidence-based & context fit", "Try in real communities; measure; improve before wider use"],
    ["Oversight & coordination", "Stay aligned with Ministry priorities and joint review meetings"],
]

PARTNERS = [
    [
        "Ministry of Health (digital health / community health)",
        "Align with NDHA, eCHIS, and national reporting; review links and the handoff plan",
    ],
    [
        "District health teams (pilot areas)",
        "Check summary reports, use dashboards, and discuss readiness",
    ],
    [
        "CHWs, VHTs, and community leaders",
        "Main partners for data; help shape forms and alerts",
    ],
    [
        "Academic partners (to be confirmed later)",
        "Help with learning and testing how systems share data",
    ],
    [
        "NGOs / community groups working with FairBanks Community Reach",
        "Help coordinate outreach and community work",
    ],
]

COST_SHARE_NOTE = (
    "We are not proposing cost share at this stage. FairBanks will keep running its "
    "medical centre and community outreach as normal. If invited to the next stage, "
    "we can discuss any in-kind support or matching help then."
)

CLOSING = (
    "With this Statement of Interest, FairBanks asks to move Track 2, Objective 4 "
    "forward through a digital health approach rooted in community work, built to "
    "connect with government systems, and designed for Ugandan ownership. We hope "
    "to be invited to Phase 2 (full proposal) or joint programme design, where we "
    "will set out workplans, budgets, monitoring, partner letters, and a clear plan "
    "to hand over staffing and running costs."
)

OUTCOMES_INTRO = (
    "Success means people get help earlier, government systems get stronger, work "
    "runs with less waste, Uganda can carry more of the load, and U.S. health "
    "investments stay protected:"
)

PARTNERS_INTRO = (
    "The partners below are early ideas for this stage. Names, contacts, and any "
    "shared budget amounts can change in Phase 2 or joint design if we are invited. "
    "For this Statement of Interest, FairBanks Medical Centre is the lead applicant."
)

APPROACH_LABEL = "How we will work:"
ACTIVITIES_LABEL = "What we will do:"
FIGURE_CAPTION = (
    "Figure 1. From community visits through FCHIP into national digital health systems "
    "(chart; not counted in the page limit)."
)


def asset(name: str) -> Path:
    p = ASSETS / name
    if not p.exists():
        raise FileNotFoundError(p)
    return p


def photo(key: str) -> Path:
    return asset(PHOTOS[key])


def embed(key_or_path, max_px: int = 1500) -> str:
    from PIL import Image as PILImage

    src = photo(key_or_path) if key_or_path in PHOTOS else Path(key_or_path)
    cache = REPO / "tmp" / f"{SLUG}_assets"
    cache.mkdir(parents=True, exist_ok=True)
    out = cache / f"{src.stem}_opt.jpg"
    if out.exists() and out.stat().st_mtime >= src.stat().st_mtime:
        return str(out)
    with PILImage.open(src) as im:
        im = im.convert("RGB")
        iw, ih = im.size
        scale = min(1.0, max_px / float(max(iw, ih)))
        if scale < 1.0:
            im = im.resize((int(iw * scale), int(ih * scale)), PILImage.Resampling.LANCZOS)
        im.save(out, format="JPEG", quality=82, optimize=True)
    return str(out)


def build_docx():
    """Submission-shaped Word: letter, 1\" margins, 12-pt Times New Roman."""
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    from PIL import Image as PILImage

    def rgb(c):
        return RGBColor.from_string(c)

    def font(run, size=12, bold=False, color=None, italic=False):
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        if color:
            run.font.color.rgb = rgb(color)

    def shade(cell, color):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))

    def border(cell, color=LINE):
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="4" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="4" w:color="{color}"/></w:tcBorders>'))

    def para(text, **kw):
        defaults = dict(
            size=12, bold=False, color=None, after=6,
            align=WD_ALIGN_PARAGRAPH.LEFT, italic=False, before=0,
        )
        defaults.update(kw)
        p = doc.add_paragraph()
        p.alignment = defaults["align"]
        p.paragraph_format.space_before = Pt(defaults["before"])
        p.paragraph_format.space_after = Pt(defaults["after"])
        p.paragraph_format.line_spacing = 1.0
        font(
            p.add_run(text),
            size=defaults["size"],
            bold=defaults["bold"],
            color=defaults["color"],
            italic=defaults["italic"],
        )
        return p

    def heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        font(p.add_run(text), size=12, bold=True)

    def bullets(items, size=12):
        for it in items:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.0
            p.clear()
            font(p.add_run(it), size=size)

    def table(headers, rows, widths=None, head_size=10, body_size=10):
        t = doc.add_table(rows=1 + len(rows), cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ""
            font(c.paragraphs[0].add_run(h), size=head_size, bold=True)
            shade(c, "E8EEEE")
            border(c, "888888")
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                c = t.rows[ri + 1].cells[ci]
                c.text = ""
                font(c.paragraphs[0].add_run(str(val)), size=body_size)
                border(c, "888888")
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Inches(w)
        doc.add_paragraph()

    def image(key, width_in=5.5, caption=None):
        path = embed(key)
        with PILImage.open(path) as im:
            iw, ih = im.size
        w = min(width_in, 3.4 * iw / ih)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(w))
        if caption:
            para(caption, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=8)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)

    # Page numbers
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run("Page ")
    font(run, size=10)
    fld1 = parse_xml(
        f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'
    )
    instr = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    fld2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run2 = fp.add_run()
    run2._r.append(fld1)
    run2._r.append(instr)
    run2._r.append(fld2)
    font(run2, size=10)

    # ----- Cover: Critical Details (excluded from page limit) -----
    para(PROGRAMME, size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    para(DOC_TITLE, size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    para(ORG_LINE, size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=4)
    para(SUBTITLE, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=4)
    para(SLOGAN, size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=10)

    heading("Table Listing of Critical Details")
    table(["Item", "Detail"], CRITICAL_DETAILS, widths=[2.4, 4.1], head_size=10, body_size=10)
    doc.add_page_break()

    # ----- Narrative (APS §D) -----
    heading("1. Issue / Challenge / Opportunity")
    for p in ISSUE:
        para(p)

    heading("2. Proposed Solution / Activities")
    for p in SOLUTION:
        para(p)
    para(WIN_WIN_INTRO, italic=True, after=6)
    table(["Who gains", "What they gain"], WIN_WIN, widths=[2.0, 4.5], head_size=10, body_size=10)

    para(ACTIVITIES_LABEL, bold=True, after=4)
    bullets(ACTIVITIES)

    para(APPROACH_LABEL, bold=True, after=4, before=6)
    for p in APPROACH:
        para(p, after=5)

    heading("3. Anticipated Outcomes and Results")
    para(OUTCOMES_INTRO, after=6)
    table(["Outcome", "What success looks like"], OUTCOMES, widths=[1.6, 4.9], head_size=10, body_size=10)

    heading("4. Organizational Capacity")
    bullets(CAPACITY)

    heading("5. How Systems Will Connect (chart)")
    table(["System", "How we connect"], INTEROP_ROWS, widths=[2.2, 4.3], head_size=10, body_size=10)
    image("architecture", width_in=5.8, caption=FIGURE_CAPTION)

    heading("6. Alignment with Addendum E Guiding Principles (chart)")
    table(["Principle", "How FairBanks responds"], PRINCIPLES, widths=[2.2, 4.3], head_size=10, body_size=10)

    heading("7. Partner Roles (early list)")
    para(PARTNERS_INTRO, size=12, after=6)
    table(
        ["Partner", "Proposed role"],
        PARTNERS,
        widths=[2.4, 4.1],
        head_size=10,
        body_size=10,
    )

    heading("8. Resource Contributions and/or Cost Share")
    para(COST_SHARE_NOTE)

    heading("9. Readiness for Phase 2")
    para(CLOSING)
    para(SLOGAN, size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, before=8)

    OUT.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_DOC))
    print(f"DOCX: {OUT_DOC}")


def build_pdf():
    """Submission-shaped PDF: letter, 1\" margins, 12-pt Times-Roman."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.colors import HexColor, black
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether,
    )
    from PIL import Image as PILImage

    slate = HexColor("#" + SLATE)
    line = HexColor("#888888")
    cream = HexColor("#E8EEEE")

    st = getSampleStyleSheet()
    for name, kw in [
        ("CoverProg", dict(fontName="Times-Bold", fontSize=11, leading=14, textColor=black, alignment=TA_CENTER, spaceAfter=4)),
        ("CoverTitle", dict(fontName="Times-Bold", fontSize=14, leading=17, textColor=black, alignment=TA_CENTER, spaceAfter=4)),
        ("CoverOrg", dict(fontName="Times-BoldItalic", fontSize=12, leading=15, textColor=black, alignment=TA_CENTER, spaceAfter=4)),
        ("CoverSub", dict(fontName="Times-Italic", fontSize=11, leading=14, textColor=black, alignment=TA_CENTER, spaceAfter=4)),
        ("Slogan", dict(fontName="Times-BoldItalic", fontSize=12, leading=15, textColor=black, alignment=TA_CENTER, spaceAfter=10)),
        ("H1", dict(fontName="Times-Bold", fontSize=12, leading=15, textColor=black, spaceBefore=10, spaceAfter=6, alignment=TA_LEFT)),
        ("Body", dict(fontName="Times-Roman", fontSize=12, leading=15, textColor=black, alignment=TA_JUSTIFY, spaceAfter=6)),
        ("BodyItalic", dict(fontName="Times-Italic", fontSize=12, leading=15, textColor=black, alignment=TA_JUSTIFY, spaceAfter=6)),
        ("FBullet", dict(fontName="Times-Roman", fontSize=12, leading=15, textColor=black, leftIndent=14, spaceAfter=3)),
        ("Meta", dict(fontName="Times-Italic", fontSize=10, leading=12, textColor=slate, alignment=TA_CENTER, spaceAfter=4)),
        ("CellHead", dict(fontName="Times-Bold", fontSize=10, leading=12, textColor=black)),
        ("CellBody", dict(fontName="Times-Roman", fontSize=10, leading=12, textColor=black)),
        ("Caption", dict(fontName="Times-Italic", fontSize=10, leading=12, textColor=slate, alignment=TA_CENTER, spaceAfter=8)),
        ("Link", dict(fontName="Times-Roman", fontSize=10, leading=12, textColor=slate, spaceAfter=2)),
    ]:
        st.add(ParagraphStyle(name, **kw))

    pw = letter[0] - 2 * inch
    story = []

    def img(key, w=pw * 0.92, cap=None, max_h=2.4 * inch):
        path = embed(key)
        with PILImage.open(path) as pi:
            iw, ih = pi.size
        aspect = ih / iw
        h = min(w * aspect, max_h)
        w = h / aspect if h == max_h else w
        block = [Image(path, width=w, height=h)]
        if cap:
            block.append(Paragraph(cap, st["Caption"]))
        story.append(KeepTogether(block))

    def tbl(headers, rows, widths=None):
        data = [[Paragraph(h, st["CellHead"]) for h in headers]]
        data += [[Paragraph(str(c), st["CellBody"]) for c in row] for row in rows]
        t = Table(data, colWidths=widths or [pw / len(headers)] * len(headers), repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), cream),
            ("GRID", (0, 0), (-1, -1), 0.4, line),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    # Cover
    story.append(Paragraph(PROGRAMME, st["CoverProg"]))
    story.append(Paragraph(DOC_TITLE, st["CoverTitle"]))
    story.append(Paragraph(ORG_LINE, st["CoverOrg"]))
    story.append(Paragraph(SUBTITLE, st["CoverSub"]))
    story.append(Paragraph(SLOGAN, st["Slogan"]))
    story.append(Paragraph("Table Listing of Critical Details", st["H1"]))
    tbl(["Item", "Detail"], CRITICAL_DETAILS, [pw * 0.32, pw * 0.68])
    story.append(PageBreak())

    story.append(Paragraph("1. Issue / Challenge / Opportunity", st["H1"]))
    for p in ISSUE:
        story.append(Paragraph(p, st["Body"]))

    story.append(Paragraph("2. Proposed Solution / Activities", st["H1"]))
    for p in SOLUTION:
        story.append(Paragraph(p, st["Body"]))
    story.append(Paragraph(WIN_WIN_INTRO, st["BodyItalic"]))
    tbl(["Who gains", "What they gain"], WIN_WIN, [pw * 0.28, pw * 0.72])

    story.append(Paragraph(f"<b>{ACTIVITIES_LABEL}</b>", st["Body"]))
    for a in ACTIVITIES:
        story.append(Paragraph(f"• {a}", st["FBullet"]))

    story.append(Paragraph(f"<b>{APPROACH_LABEL}</b>", st["Body"]))
    for p in APPROACH:
        story.append(Paragraph(p, st["Body"]))

    story.append(Paragraph("3. Anticipated Outcomes and Results", st["H1"]))
    story.append(Paragraph(OUTCOMES_INTRO, st["Body"]))
    tbl(["Outcome", "What success looks like"], OUTCOMES, [pw * 0.24, pw * 0.76])

    story.append(Paragraph("4. Organizational Capacity", st["H1"]))
    for c in CAPACITY:
        story.append(Paragraph(f"• {c}", st["FBullet"]))

    story.append(Paragraph("5. How Systems Will Connect (chart)", st["H1"]))
    tbl(["System", "How we connect"], INTEROP_ROWS, [pw * 0.32, pw * 0.68])
    img("architecture", cap=FIGURE_CAPTION, max_h=2.2 * inch)

    story.append(Paragraph("6. Alignment with Addendum E Guiding Principles (chart)", st["H1"]))
    tbl(["Principle", "How FairBanks responds"], PRINCIPLES, [pw * 0.32, pw * 0.68])

    story.append(Paragraph("7. Partner Roles (early list)", st["H1"]))
    story.append(Paragraph(PARTNERS_INTRO, st["Body"]))
    tbl(["Partner", "Proposed role"], PARTNERS, [pw * 0.34, pw * 0.66])

    story.append(Paragraph("8. Resource Contributions and/or Cost Share", st["H1"]))
    story.append(Paragraph(COST_SHARE_NOTE, st["Body"]))

    story.append(Paragraph("9. Readiness for Phase 2", st["H1"]))
    story.append(Paragraph(CLOSING, st["Body"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(SLOGAN, st["Slogan"]))

    def _page(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Times-Roman", 10)
        canvas.drawCentredString(letter[0] / 2, 0.55 * inch, f"Page {doc_.page}")
        canvas.restoreState()

    OUT.mkdir(parents=True, exist_ok=True)
    tmp_pdf = REPO / "tmp" / SLUG / f"{SLUG}_pdf.build.pdf"
    tmp_pdf.parent.mkdir(parents=True, exist_ok=True)
    SimpleDocTemplate(
        str(tmp_pdf),
        pagesize=letter,
        leftMargin=1 * inch,
        rightMargin=1 * inch,
        topMargin=1 * inch,
        bottomMargin=1 * inch,
    ).build(story, onFirstPage=_page, onLaterPages=_page)
    try:
        tmp_pdf.replace(OUT_PDF)
    except OSError:
        import shutil
        try:
            shutil.copy2(tmp_pdf, OUT_PDF)
        except OSError as e:
            print(f"PDF built at {tmp_pdf} (could not write {OUT_PDF}: {e})")
            return
    print(f"PDF: {OUT_PDF}")


def build_pptx():
    """Elegant briefing deck — same win-win story; not the official upload format."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from PIL import Image as PILImage

    def _add_entrance_anims(slide):
        from lxml import etree
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        spids = []
        for shape in slide.shapes:
            try:
                stype = int(shape.shape_type) if shape.shape_type is not None else -1
            except Exception:
                stype = -1
            has_text = bool(getattr(shape, "has_text_frame", False) and shape.has_text_frame)
            if stype == 13 or has_text:
                spids.append(str(shape.shape_id))
            if len(spids) >= 3:
                break
        if not spids:
            return
        sld = slide._element
        for old in sld.findall(f"{{{NS_P}}}timing"):
            sld.remove(old)
        children = []
        nid = 3
        for i, spid in enumerate(spids):
            delay = 0 if i == 0 else 200
            children.append(
                f'<p:par xmlns:p="{NS_P}">'
                f'<p:cTn id="{nid}" presetID="10" presetClass="entr" presetSubtype="0" '
                f'fill="hold" nodeType="withEffect">'
                f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
                f'<p:childTnLst>'
                f'<p:animEffect transition="in" filter="fade">'
                f'<p:cBhvr><p:cTn id="{nid + 1}" dur="450"/>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr>'
                f'</p:animEffect></p:childTnLst></p:cTn></p:par>'
            )
            nid += 2
        xml = (
            f'<p:timing xmlns:p="{NS_P}">'
            f'<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst><p:seq concurrent="true" nextAc="seek">'
            f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            f'<p:par><p:cTn id="{nid}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(children)}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        )
        sld.append(etree.fromstring(xml))

    def C(h):
        return RGBColor.from_string(h)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    SW, SH, blank = prs.slide_width, prs.slide_height, prs.slide_layouts[6]

    def rect(sl, x, y, w, h, fill, line=None):
        s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        s.fill.solid()
        s.fill.fore_color.rgb = C(fill)
        s.line.color.rgb = C(line) if line else C(fill)
        if not line:
            s.line.fill.background()
        return s

    def tb(sl, x, y, w, h, text, size=18, bold=False, color=SLATE, align=PP_ALIGN.LEFT, italic=False):
        box = sl.shapes.add_textbox(x, y, w, h)
        tf = box.text_frame
        tf.word_wrap = True
        h_in = float(h) / 914400.0
        w_in = float(w) / 914400.0
        body = h_in >= 4.0 and 13 <= size <= 22
        if body:
            try:
                tf.anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
        raw = str(text).replace("\r\n", "\n")
        if "\n\n" in raw:
            parts = [p.strip() for p in raw.split("\n\n") if p.strip()]
        else:
            parts = list(raw.split("\n"))
        if not parts:
            parts = [""]
        cpl = max(18, int(w_in * 6.5))
        est = 0.0
        nonempty = 0
        for part in parts:
            if not str(part).strip():
                est += 0.35
                continue
            nonempty += 1
            est += max(1, (len(part) + cpl - 1) // cpl)
        est = max(est, float(nonempty or 1))
        use_size = size
        if body and est > 0:
            fitted = int((h_in * 0.82 * 72) / (est * 1.32))
            use_size = max(15, min(26, fitted))
            if fitted >= size:
                use_size = max(size, use_size)
        gap = max(10, int(use_size * 0.55)) if body else 3
        for i, ln in enumerate(parts):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.space_before = Pt(0)
            p.space_after = Pt(gap if body and i < len(parts) - 1 else 2)
            try:
                p.line_spacing = 1.28 if body else 1.15
            except Exception:
                pass
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(use_size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = C(color)
            r.font.name = "Calibri"
        return box

    def pic(sl, key, x, y, w, h):
        path = embed(key)
        with PILImage.open(path) as im:
            iw, ih = im.size
        aspect = ih / iw
        tw, th = w, int(w * aspect)
        if th > h:
            th, tw = h, int(h / aspect)
        sl.shapes.add_picture(path, x + (w - tw) // 2, y + (h - th) // 2, width=tw, height=th)

    def header(sl, roman, title):
        rect(sl, 0, 0, SW, Inches(0.08), TEAL)
        tb(sl, Inches(0.5), Inches(0.2), Inches(1.2), Inches(0.5), roman, size=28, bold=True, color=ACCENT)
        tb(sl, Inches(1.6), Inches(0.22), Inches(11), Inches(0.55), title, size=22, bold=True, color=NAVY)

    def footer(sl, n):
        rect(sl, 0, SH - Inches(0.28), SW, Inches(0.28), SLATE)
        tb(sl, Inches(0.4), SH - Inches(0.26), Inches(11), Inches(0.22),
           f"FairBanks SOI  |  DFOP0017890 Objective 4  |  {SLOGAN}", size=9, color="FFFFFF")
        tb(sl, SW - Inches(0.8), SH - Inches(0.26), Inches(0.5), Inches(0.22),
           str(n), size=9, color="FFFFFF", align=PP_ALIGN.RIGHT)

    # Cover
    s = prs.slides.add_slide(blank)
    pic(s, "cover", 0, 0, SW, SH)
    rect(s, 0, SH - Inches(3.6), SW, Inches(3.6), NAVY)
    tb(s, Inches(0.55), SH - Inches(3.3), Inches(12), Inches(0.3), PROGRAMME, size=11, bold=True, color=TEAL)
    tb(s, Inches(0.55), SH - Inches(2.85), Inches(12), Inches(0.55), DOC_TITLE, size=26, bold=True, color="FFFFFF")
    tb(s, Inches(0.55), SH - Inches(2.2), Inches(12), Inches(0.35),
       ORG_LINE, size=13, bold=True, italic=True, color="F2C79B")
    tb(s, Inches(0.55), SH - Inches(1.7), Inches(12), Inches(0.45),
       "Objective 4 — Helping community visit data reach eCHIS, clinic records, and DHIS2/eIDSR",
       size=12, italic=True, color="D0E8E8")
    tb(s, Inches(0.55), SH - Inches(1.1), Inches(12), Inches(0.3),
       f"Phase 1 Statement of Interest  |  {FEDERAL_SHARE}  |  {PROJECT_LENGTH}  |  Due 31 July 2026",
       size=12, color="FFFFFF")
    tb(s, Inches(0.55), SH - Inches(0.7), Inches(12), Inches(0.3), SLOGAN, size=13, bold=True, color="FFFFFF")

    win_body = "\n".join(
        f"• {row[0]}: {row[1][:95]}..." if len(row[1]) > 95 else f"• {row[0]}: {row[1]}"
        for row in WIN_WIN
    )
    act_body = "\n".join(f"• {a[:115]}..." if len(a) > 115 else f"• {a}" for a in ACTIVITIES[:4])
    out_body = "\n".join(f"• {r[0]} — {r[1][:85]}" for r in OUTCOMES)

    sections = [
        ("I", "The Problem", ISSUE[0] + "\n\n" + ISSUE[1], "outreach"),
        ("II", "What We Propose", SOLUTION[1], "architecture"),
        ("III", "Who Gains What", win_body, None),
        ("IV", "What We Will Do", act_body, "mobile"),
        ("V", "How We Will Work",
         "Community visits → safe sync → links to eCHIS, clinic records, and DHIS2/eIDSR "
         "→ dashboards and maps → clear handoff plan for the Ministry of Health.",
         "dashboard"),
        ("VI", "What Success Looks Like", out_body, "gis"),
        ("VII", "Our Capacity",
         "\n".join(f"• {c}" for c in CAPACITY[:4]), "facility"),
        ("VIII", "How Systems Connect",
         "Phone tools for eCHIS-style household visits → clinic records "
         "(eAfya/ClinicMaster) → DHIS2/eIDSR summaries → National Data Warehouse "
         "when ready → CHW/VHT lists that can align with iHRIS. A bridge — not a "
         "replacement. Built to follow NDHA / OpenHIE sharing rules.",
         "pharmacy"),
        ("IX", "Ownership Over Time",
         "By project close, write down how staff roles and running costs move to "
         "government ownership.\n\n"
         "Train Ugandan developers, CHW supervisors, and data staff.\n\n"
         "The clinic and outreach keep running as FairBanks work. The grant helps "
         "connect the data and prepare the handoff.",
         None),
        ("X", "Ready for Phase 2", CLOSING, None),
    ]

    for i, (roman, title, body, img_key) in enumerate(sections, start=2):
        s = prs.slides.add_slide(blank)
        header(s, roman, title)
        if img_key:
            pic(s, img_key, Inches(0.45), Inches(1.05), Inches(5.8), Inches(5.9))
            tb(s, Inches(6.5), Inches(1.15), Inches(6.4), Inches(5.6), body, size=17, color=SLATE)
        else:
            tb(s, Inches(0.55), Inches(1.15), Inches(12.2), Inches(5.8), body, size=18, color=SLATE)
        footer(s, i)

    OUT.mkdir(parents=True, exist_ok=True)
    for _sl in prs.slides:
        _add_entrance_anims(_sl)
    prs.save(str(OUT_PPT))
    print(f"PPTX: {OUT_PPT}")


if __name__ == "__main__":
    build_docx()
    build_pdf()
    build_pptx()
    print("Done.", OUT)
