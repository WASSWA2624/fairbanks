#!/usr/bin/env python3
"""
U.S. Dept of State — Uganda Health System MOU (DFOP0017890) — Statement of Interest.

Win-win Step 1 SOI under Addendum E (Uganda), focused on Objective 4 digital health.
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

CALL_URL = (
    "https://opportunitiesforyouth.org/2026/06/15/"
    "u-s-department-of-state-launches-up-to-60-million-funding-opportunity-to-strengthen-"
    "ugandas-health-system-through-the-health-foreign-assistance-mou-implementation-plan/"
)
GRANTS_URL = "https://simpler.grants.gov/opportunity/44a573b8-d46a-4cec-8c84-f319cd1cc1b6"

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

PROGRAMME = "U.S. Department of State / GHSD — Advancing Global Health (DFOP0017890)"
DOC_TITLE = "Statement of Interest — Step 1"
SUBTITLE = (
    "Addendum E: Health Foreign Assistance MOU Implementation in Uganda  |  "
    "Objective 4: Strengthen and Scale Digital Health Systems"
)
ORG_LINE = "FairBanks Medical Centre  ·  FairBanks Community Reach  ·  FCHIP"

# ---------------------------------------------------------------------------
# Narrative content — win-win SOI for Step 1
# ---------------------------------------------------------------------------

PROBLEM = [
    "Uganda's five-year health transition (April 2026-December 2030) asks facilities, "
    "districts, and communities to keep HIV, TB, malaria, maternal and child health, "
    "immunisation, and outbreak services strong while shifting toward government-led, "
    "performance-based programming. That shift only works if decision-makers see what is "
    "happening at the last mile - not weeks later in paper registers.",
    "Community Health Workers and Village Health Teams already walk households and "
    "refer patients. Too often their work never reaches eCHIS, DHIS2, facility EMRs, or "
    "district dashboards in a usable form. National digital systems then miss early "
    "fever clusters, missed antenatal visits, immunisation gaps, and stock pressure. "
    "The result is reactive care, weak accountability, and higher risk that transition "
    "gains reverse.",
]

STATEMENT = [
    "FairBanks Medical Centre and FairBanks Community Reach respectfully submit this "
    "Statement of Interest under Track 2, Objective 4 of the Uganda MOU Implementation "
    "Addendum: strengthen and scale digital health systems with full interoperability, "
    "government ownership, and sustainable operations.",
    "We propose to pilot and then transfer a community-to-facility intelligence layer - "
    "the FairBanks Community Health Intelligence Platform (FCHIP) - that captures "
    "CHW/VHT household and visit data offline, syncs securely, maps indicators to "
    "national reporting elements, and feeds facility and district action. FCHIP is not "
    "a parallel silo. It is a bridge that makes Uganda's existing digital stack "
    "(eCHIS, DHIS2, facility EMRs, and related national platforms) more complete at "
    "community level and more useful for prevention.",
]

WIN_WIN = [
    [
        "Uganda / Ministry of Health",
        "Richer last-mile data into eCHIS and DHIS2; clearer referral closure; "
        "documented handoff pathways toward government ownership and reduced "
        "dependence on parallel partner tools.",
    ],
    [
        "U.S. Department of State / GHSD",
        "A Ugandan implementer with a live clinic and community network that "
        "protects essential services while proving interoperable digital scale-up "
        "under the MOU - accountable, measurable, and transferable.",
    ],
    [
        "Districts & facilities",
        "Dashboards for outreach targeting, referral follow-up, and programme "
        "monitoring; GIS hotspot views that support planning without waiting for "
        "outbreaks to arrive at the ward.",
    ],
    [
        "Communities & CHWs/VHTs",
        "Simple offline tools, faster support from the medical centre, and care "
        "that reaches mothers, children, older persons, and people with chronic "
        "conditions before crises.",
    ],
    [
        "FairBanks / FCHIP",
        "Room to refine and document a production-grade interoperability model "
        "inside a real catchment - then share specifications, training packages, "
        "and transfer plans with MoH-aligned partners.",
    ],
]

OBJECTIVES = [
    "O1. Deploy offline-capable CHW/VHT capture aligned with eCHIS household and "
    "visit concepts across FairBanks Community Reach catchments.",
    "O2. Establish a secure sync and role-based access pipeline linking community "
    "records to FairBanks Medical Centre workflows and authorised district viewers.",
    "O3. Map community indicators to DHIS2 reporting elements and demonstrate "
    "aggregate submission pathways for MNCH, immunisation, NCD screening, and "
    "disease surveillance signals.",
    "O4. Deliver facility and programme dashboards plus GIS hotspot views that "
    "trigger outreach, referrals, and supply planning actions.",
    "O5. Produce interoperability specifications, data-governance protocols, "
    "training curricula, and a staged government-ownership transition plan suitable "
    "for full-proposal detailed design.",
]

APPROACH = [
    "Anchor in a live cascade: community members identify needs; CHWs/VHTs collect "
    "and educate; FairBanks Community Reach runs outreach and home visits; "
    "FairBanks Medical Centre provides clinical care, diagnostics, pharmacy, and "
    "referrals; FCHIP is the intelligence layer that closes the data-and-feedback loop.",
    "Build API-first, standards-aware exports - not a closed proprietary vault. "
    "Prioritise interoperability with eCHIS (household/visit concepts), DHIS2 "
    "(aggregate indicators), facility EMR encounter/referral handoff, and readiness "
    "for national warehouse feeds when schemas are confirmed.",
    "Protect essential services during transition: structure community data around "
    "HIV/TB/malaria flags where relevant, maternal and child pathways, immunisation "
    "follow-up, and outbreak early warning so digital work reinforces - not distracts "
    "from - service continuity.",
    "Co-design forms and alerts with CHWs, facility staff, and local leaders; train "
    "Ugandan data stewards and developers; document every interface for MoH review.",
    "Start where FairBanks already operates (Bukoto, Kyebando, Kisaasi, Kamwokya, "
    "Kikaaya and surrounding communities), then design for district replication "
    "using Uganda's existing CHW/VHT architecture.",
]

OUTCOMES = [
    [
        "Digital completeness",
        "Increase structured community visit capture flowing into facility and "
        "aggregate reporting pathways within the pilot catchment.",
    ],
    [
        "Interoperability proof",
        "Working mappings and demo exports for eCHIS-aligned records and DHIS2 "
        "aggregates, with open documentation for MoH and partners.",
    ],
    [
        "Service continuity signals",
        "Earlier alerts for missed ANC, immunisation gaps, fever clusters, and "
        "referral drop-offs - linked to CHW follow-up and clinic action.",
    ],
    [
        "Accountability",
        "Role-based audit trails, data-quality checks, and programme dashboards "
        "usable for performance-oriented district conversations.",
    ],
    [
        "Ownership readiness",
        "Handoff package: source documentation, SOPs, training, hosting options, "
        "and a costed path to transfer recurrent operations toward domestic systems.",
    ],
]

CAPACITY = [
    "Ugandan social enterprise with a functioning medical centre - clinical care, "
    "diagnostics, pharmacy, and referral pathways already in daily use.",
    "Active FairBanks Community Reach programmes: maternal and child health, "
    "Gericare, chronic disease screening, school and corporate health outreach.",
    "Established CHW/VHT relationships in Kampala-area communities named above - "
    "real users, not a greenfield pilot.",
    "Digital health records foundation and pharmacy dispensing data already "
    "supporting facility operations.",
    "Research orientation and partnership practice - evidence, ethics, and skills "
    "development sit inside the FairBanks community health model.",
]

INTEROP_ROWS = [
    ["eCHIS", "Household registration, visit logs, referral flags via structured mobile capture"],
    ["DHIS2", "Aggregate exports for immunisation, ANC, NCD screening, disease surveillance"],
    ["Facility EMR", "Referral handoff and encounter linkage at FairBanks Medical Centre"],
    ["National warehouse / related systems", "Standardised JSON/API feeds when national schema is confirmed"],
]

PRINCIPLES = [
    ["Sustainability & country ownership", "Design for MoH alignment and staged transfer of operations"],
    ["Protect essential services", "Digital workflows reinforce HIV, TB, malaria, MNCH, immunisation"],
    ["Data-driven accountability", "Quality data for decisions - not parallel reporting burden"],
    ["Integrated, people-centred care", "One community-to-facility loop; less fragmentation"],
    ["Local partnerships", "Government, CHWs/VHTs, CBOs, academic and NGO partners"],
    ["Evidence-based programming", "Pilot in live catchments; measure; refine before scale"],
]

GEO = [
    "Primary pilot: FairBanks catchment - Kampala-area communities with live CHW operations.",
    "Scale path: architecture designed for district replication wherever CHW/VHT systems operate.",
    "MOU fit: Ugandan organisation implementing under national health priorities and digital stack.",
]

SUSTAIN = [
    "Core clinic and outreach continue as FairBanks operations - the grant accelerates "
    "interoperable intelligence and transfer, not basic organisational survival.",
    "Diversified sustainability pathways after transfer: facility use, district "
    "deployments, NGO programme monitoring, partner APIs, and training packages.",
    "Local capacity: train Ugandan developers, CHW supervisors, and data stewards; "
    "document systems for government handoff.",
    "Community ownership: CHWs and leaders co-design forms and alerts - technology "
    "serves the cascade, not the other way around.",
]

PARTNERS = [
    "Ministry of Health digital health and community health strategy alignment",
    "District health teams for aggregate reporting validation and readiness dialogue",
    "CHWs, VHTs, and community leaders as primary data partners",
    "Academic partners for evaluation and interoperability testing",
    "NGOs and CBOs already collaborating with FairBanks Community Reach",
]

CLOSING = (
    "This Statement of Interest confirms FairBanks' intent to advance Objective 4 "
    "with a community-rooted, interoperable, government-aligned digital health "
    "proposal. We seek invitation to Step 2 (full proposal), where we will detail "
    "workplans, budgets, M&E, partnership letters, and a clear transition of "
    "staffing and recurrent costs toward sustainable Ugandan ownership."
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
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    from PIL import Image as PILImage

    def rgb(c):
        return RGBColor.from_string(c)

    def font(run, size=11, bold=False, color=SLATE, italic=False):
        run.font.name = "Calibri"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = rgb(color)

    def shade(cell, color):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))

    def border(cell, color=LINE):
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="8" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="8" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="8" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="8" w:color="{color}"/></w:tcBorders>'))

    def para(text, **kw):
        defaults = dict(size=11, bold=False, color=SLATE, after=8, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False)
        defaults.update(kw)
        p = doc.add_paragraph()
        p.alignment = defaults["align"]
        p.paragraph_format.space_after = Pt(defaults["after"])
        font(p.add_run(text), size=defaults["size"], bold=defaults["bold"],
             color=defaults["color"], italic=defaults["italic"])

    def heading(text, level=1):
        sizes, colors = {1: 16, 2: 13, 3: 12}, {1: NAVY, 2: TEAL, 3: SLATE}
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
        p.paragraph_format.space_after = Pt(5)
        font(p.add_run(text), size=sizes[level], bold=True, color=colors[level])

    def bullets(items):
        for it in items:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            p.clear()
            font(p.add_run(it), size=10)

    def table(headers, rows, widths=None):
        t = doc.add_table(rows=1 + len(rows), cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ""
            font(c.paragraphs[0].add_run(h), size=9, bold=True, color="FFFFFF")
            shade(c, TEAL)
            border(c, TEAL)
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                c = t.rows[ri + 1].cells[ci]
                c.text = ""
                font(c.paragraphs[0].add_run(str(val)), size=9)
                if ri % 2:
                    shade(c, CREAM)
                border(c)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Inches(w)
        doc.add_paragraph()

    def image(key, width_in=5.4, caption=None):
        path = embed(key)
        with PILImage.open(path) as im:
            iw, ih = im.size
        w = min(width_in, 3.2 * iw / ih)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(w))
        if caption:
            para(caption, size=8, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=8)

    doc = Document()
    for m in doc.sections:
        m.left_margin = m.right_margin = Inches(0.8)
        m.top_margin = m.bottom_margin = Inches(0.7)

    # Cover — no application-details table
    para(PROGRAMME, size=11, bold=True, color=TEAL, align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
    para(DOC_TITLE, size=20, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
    para(ORG_LINE, size=12, bold=True, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=3)
    para(SUBTITLE, size=10, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=4)
    para(SLOGAN, size=12, bold=True, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=8)
    image("cover", caption="FairBanks — community health ecosystem rooted in Uganda")
    para(
        "SOI deadline: 31 July 2026, 5:00 PM EDT  |  Opportunity: DFOP0017890  |  "
        "Addendum E — Uganda Health MOU",
        size=9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, after=4,
    )
    doc.add_page_break()

    heading("I. The Problem We Address")
    for p in PROBLEM:
        para(p, size=10, after=6)

    heading("II. Statement of Interest")
    for p in STATEMENT:
        para(p, size=10, after=6)

    heading("III. Win-Win Value Proposition")
    para(
        "This partnership is designed so each party gains something durable - "
        "not a one-sided software purchase.",
        size=10, after=6, italic=True,
    )
    table(["Who gains", "What they gain"], WIN_WIN, widths=[2.1, 4.5])

    heading("IV. Proposed Objectives (Objective 4)")
    bullets(OBJECTIVES)

    heading("V. Project Approach")
    for p in APPROACH:
        para(p, size=10, after=5)
    image("architecture", caption="Community cascade → FCHIP → national digital health systems")

    heading("VI. Expected Outcomes")
    table(["Outcome", "What success looks like"], OUTCOMES, widths=[1.7, 4.9])

    heading("VII. Organizational Capacity")
    image("facility", width_in=5.0, caption="FairBanks Medical Centre — clinical anchor for community data")
    bullets(CAPACITY)

    heading("VIII. Interoperability Approach")
    table(["Platform", "Integration pathway"], INTEROP_ROWS, widths=[2.0, 4.6])
    image("dashboard", width_in=5.0, caption="Facility and programme dashboards for action")

    heading("IX. Alignment with MOU Guiding Principles")
    table(["Principle", "How FairBanks responds"], PRINCIPLES, widths=[2.2, 4.4])

    heading("X. Geographic Focus")
    bullets(GEO)
    image("outreach", width_in=5.0, caption="Live community outreach — pilot ground truth")

    heading("XI. Sustainability & Local Ownership")
    bullets(SUSTAIN)

    heading("XII. Partnership Landscape")
    bullets(PARTNERS)

    heading("XIII. Readiness for Full Proposal")
    para(CLOSING, size=10, after=8)
    para(
        "Official references: DFOP0017890 Addendum E (Uganda). Step 1 SOI. "
        "Narrative kept concise for Step 1 review.",
        size=8, color=MUTED, italic=True, after=2,
    )
    para(CALL_URL, size=7, color=TEAL, after=2)
    para(GRANTS_URL, size=7, color=TEAL, after=8)
    para(SLOGAN, size=12, bold=True, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)

    OUT.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_DOC))
    print(f"DOCX: {OUT_DOC}")


def build_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from reportlab.lib.colors import HexColor, white
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether,
    )
    from PIL import Image as PILImage

    navy, teal, accent = HexColor("#" + NAVY), HexColor("#" + TEAL), HexColor("#" + ACCENT)
    slate, muted, cream, line = (
        HexColor("#" + SLATE), HexColor("#" + MUTED), HexColor("#" + CREAM), HexColor("#" + LINE),
    )
    st = getSampleStyleSheet()
    for name, kw in [
        ("CoverTitle", dict(fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=navy, alignment=TA_CENTER, spaceAfter=5)),
        ("H1", dict(fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=navy, spaceBefore=9, spaceAfter=4)),
        ("Body", dict(fontName="Helvetica", fontSize=9, leading=12, textColor=slate, alignment=TA_JUSTIFY, spaceAfter=4)),
        ("Meta", dict(fontName="Helvetica", fontSize=8, leading=10, textColor=muted, alignment=TA_CENTER)),
        ("FBullet", dict(fontName="Helvetica", fontSize=9, leading=11, textColor=slate, leftIndent=10, spaceAfter=2)),
        ("CellHead", dict(fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=white)),
        ("CellBody", dict(fontName="Helvetica", fontSize=7.5, leading=10, textColor=slate)),
        ("AccentMeta", dict(fontName="Helvetica-BoldOblique", fontSize=10, leading=12, textColor=accent, alignment=TA_CENTER, spaceAfter=3)),
    ]:
        st.add(ParagraphStyle(name, **kw))

    pw = A4[0] - 1.5 * inch
    story = []

    def img(key, w=pw * 0.82, cap=None, max_h=2.2 * inch):
        path = embed(key)
        with PILImage.open(path) as pi:
            iw, ih = pi.size
        aspect = ih / iw
        h = min(w * aspect, max_h)
        w = h / aspect if h == max_h else w
        block = [Image(path, width=w, height=h)]
        if cap:
            block.append(Paragraph(cap, st["Meta"]))
        story.append(KeepTogether(block))
        story.append(Spacer(1, 4))

    def tbl(headers, rows, widths=None):
        data = [[Paragraph(h, st["CellHead"]) for h in headers]]
        data += [[Paragraph(str(c), st["CellBody"]) for c in row] for row in rows]
        t = Table(data, colWidths=widths or [pw / len(headers)] * len(headers), repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), teal),
            ("GRID", (0, 0), (-1, -1), 0.4, line),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, cream]),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 6))

    # Cover
    story.append(Paragraph(PROGRAMME, st["Meta"]))
    story.append(Spacer(1, 4))
    story.append(Paragraph(DOC_TITLE, st["CoverTitle"]))
    story.append(Paragraph(ORG_LINE, st["AccentMeta"]))
    story.append(Paragraph(f'<i><font color="#{MUTED}">{SUBTITLE}</font></i>', st["Meta"]))
    story.append(Paragraph(f'<b><i><font color="#{ACCENT}">{SLOGAN}</font></i></b>', st["Meta"]))
    story.append(Spacer(1, 8))
    img("cover", cap="FairBanks — community health ecosystem rooted in Uganda", max_h=2.8 * inch)
    story.append(Paragraph(
        "SOI deadline: 31 July 2026, 5:00 PM EDT  |  DFOP0017890  |  Addendum E — Uganda",
        st["Meta"],
    ))
    story.append(PageBreak())

    story.append(Paragraph("I. The Problem We Address", st["H1"]))
    for p in PROBLEM:
        story.append(Paragraph(p, st["Body"]))

    story.append(Paragraph("II. Statement of Interest", st["H1"]))
    for p in STATEMENT:
        story.append(Paragraph(p, st["Body"]))

    story.append(Paragraph("III. Win-Win Value Proposition", st["H1"]))
    story.append(Paragraph(
        "This partnership is designed so each party gains something durable - "
        "not a one-sided software purchase.",
        st["Body"],
    ))
    tbl(["Who gains", "What they gain"], WIN_WIN, [pw * 0.28, pw * 0.72])

    story.append(Paragraph("IV. Proposed Objectives (Objective 4)", st["H1"]))
    for o in OBJECTIVES:
        story.append(Paragraph(f"• {o}", st["FBullet"]))

    story.append(Paragraph("V. Project Approach", st["H1"]))
    for p in APPROACH[:3]:
        story.append(Paragraph(p, st["Body"]))
    for p in APPROACH[3:]:
        story.append(Paragraph(f"• {p}", st["FBullet"]))
    img(
        "architecture",
        w=pw * 0.78,
        cap="Community cascade → FCHIP → national digital health systems",
        max_h=1.65 * inch,
    )

    story.append(Paragraph("VI. Expected Outcomes", st["H1"]))
    tbl(["Outcome", "What success looks like"], OUTCOMES, [pw * 0.26, pw * 0.74])

    story.append(Paragraph("VII. Organizational Capacity", st["H1"]))
    for c in CAPACITY:
        story.append(Paragraph(f"• {c}", st["FBullet"]))

    story.append(PageBreak())
    story.append(Paragraph("VIII. Interoperability Approach", st["H1"]))
    tbl(["Platform", "Integration pathway"], INTEROP_ROWS, [pw * 0.28, pw * 0.72])

    story.append(Paragraph("IX. Alignment with MOU Guiding Principles", st["H1"]))
    tbl(["Principle", "How FairBanks responds"], PRINCIPLES, [pw * 0.32, pw * 0.68])

    story.append(Paragraph("X. Geographic Focus", st["H1"]))
    for g in GEO:
        story.append(Paragraph(f"• {g}", st["FBullet"]))

    story.append(Paragraph("XI. Sustainability & Local Ownership", st["H1"]))
    for s in SUSTAIN:
        story.append(Paragraph(f"• {s}", st["FBullet"]))

    story.append(Paragraph("XII. Partnership Landscape", st["H1"]))
    for p in PARTNERS:
        story.append(Paragraph(f"• {p}", st["FBullet"]))

    story.append(Paragraph("XIII. Readiness for Full Proposal", st["H1"]))
    story.append(Paragraph(CLOSING, st["Body"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f'<b><i><font color="#{ACCENT}">{SLOGAN}</font></i></b>', st["Meta"]))

    OUT.mkdir(parents=True, exist_ok=True)
    SimpleDocTemplate(
        str(OUT_PDF), pagesize=A4,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch,
        topMargin=0.65 * inch, bottomMargin=0.65 * inch,
    ).build(story)
    print(f"PDF: {OUT_PDF}")


def build_pptx():
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
       "Objective 4 — Community-to-facility intelligence interoperable with eCHIS and DHIS2",
       size=12, italic=True, color="D0E8E8")
    tb(s, Inches(0.55), SH - Inches(1.1), Inches(12), Inches(0.3),
       "Win-win Step 1 Statement of Interest  |  Deadline 31 July 2026, 5:00 PM EDT",
       size=12, color="FFFFFF")
    tb(s, Inches(0.55), SH - Inches(0.7), Inches(12), Inches(0.3), SLOGAN, size=13, bold=True, color="FFFFFF")

    win_body = "\n".join(
        f"• {row[0]}: {row[1][:90]}..." if len(row[1]) > 90 else f"• {row[0]}: {row[1]}"
        for row in WIN_WIN[:4]
    )
    obj_body = "\n".join(f"• {o[:110]}..." if len(o) > 110 else f"• {o}" for o in OBJECTIVES[:4])
    out_body = "\n".join(f"• {r[0]} — {r[1][:80]}" for r in OUTCOMES[:4])

    sections = [
        ("I", "The Problem", PROBLEM[0], "outreach"),
        ("II", "Statement of Interest", STATEMENT[1], "architecture"),
        ("III", "Win-Win Value", win_body, None),
        ("IV", "Proposed Objectives", obj_body, "mobile"),
        ("V", "Project Approach",
         "Cascade-rooted capture → secure sync → eCHIS/DHIS2/EMR interoperability → "
         "dashboards & GIS → MoH-aligned transfer package.", "dashboard"),
        ("VI", "Expected Outcomes", out_body, "gis"),
        ("VII", "Organizational Capacity",
         "\n".join(f"• {c}" for c in CAPACITY[:4]), "facility"),
        ("VIII", "Interoperability",
         "eCHIS household capture → DHIS2 aggregates → facility EMR referrals → "
         "national warehouse readiness. Bridge layer — not a replacement silo.", "pharmacy"),
        ("IX", "Sustainability & Ownership",
         "\n".join(f"• {s}" for s in SUSTAIN[:3]), None),
        ("X", "Readiness for Full Proposal",
         CLOSING, None),
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
