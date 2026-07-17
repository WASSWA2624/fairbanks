#!/usr/bin/env python3
"""
U.S. Dept of State — Uganda Health System MOU (DFOP0017890) — Statement of Interest.

Generates synced Word, PDF, and PowerPoint aligned with Objective 4 digital health.
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

PROGRAMME = "U.S.–Uganda Health MOU — Funding Opportunity DFOP0017890"
DOC_TITLE = "Statement of Interest — Objective 4 Digital Health Scale-Up"
SUBTITLE = "Community-to-facility intelligence interoperable with eCHIS and DHIS2"

STATEMENT = [
    "FairBanks Medical Centre and FairBanks Community Reach respectfully submit this "
    "Statement of Interest under Objective 4: scaling digital health systems in Uganda. "
    "We propose to extend and interoperate community-generated health data — captured "
    "by Community Health Workers and Village Health Teams — with national platforms "
    "(eCHIS, DHIS2, facility EMRs) through FairBanks Community Health Intelligence "
    "Platform (FCHIP).",
    "FairBanks is a Ugandan social enterprise with a functioning medical centre, active "
    "outreach in Kampala communities, and established CHW/VHT relationships. We are "
    "not proposing a standalone app silo. We are proposing a bridge layer that strengthens "
    "Uganda's digital health architecture while keeping communities at the centre of care.",
]

CAPACITY = [
    "Operating medical centre with pharmacy, diagnostics, and referral pathways",
    "Community Reach programmes: maternal & child health, Gericare, NCD screening, school health",
    "Active CHW/VHT networks in Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya",
    "Digital health records foundation and pharmacy dispensing data already in use",
    "Research partnerships and ethical data governance experience",
]

OBJ4_FIT = [
    ["Objective 4 intent", "Scale EMRs, eCHIS, DHIS2, iHRIS, and national data warehouse capacity"],
    ["FairBanks response", "FCHIP extends last-mile capture and feeds structured community data upstream"],
    ["Interoperability", "API-first exports mapped to DHIS2 indicators; eCHIS-aligned household records"],
    ["Not a replacement", "Complements MoH systems — adds predictive intelligence and community loop closure"],
]

PROPOSED = [
    "Deploy offline-capable CHW capture aligned with eCHIS household and visit concepts",
    "Build secure sync pipeline with role-based access for CHWs, facilities, and district viewers",
    "Map community indicators to DHIS2 reporting elements for aggregate submission",
    "Provide facility dashboards for referrals, alerts, and programme monitoring",
    "Pilot GIS hotspot views to support targeted outreach and supply planning",
    "Document interoperability specifications for full proposal stage",
]

INTEROP_ROWS = [
    ["eCHIS", "Household registration, visit logs, referral flags — structured mobile capture"],
    ["DHIS2", "Aggregate exports for immunisation, ANC, NCD screening, disease surveillance"],
    ["Facility EMR", "Referral handoff and encounter linkage at FairBanks Medical Centre"],
    ["National warehouse", "Standardised JSON/API feeds when national schema is confirmed"],
]

GEO = [
    "Primary pilot: FairBanks catchment — Kampala-area communities with live CHW operations",
    "Scalable model: CHW/VHT systems exist nationwide; architecture designed for district replication",
    "Alignment with MOU geography: Ugandan organisations implementing under national health priorities",
]

SUSTAIN = [
    "FairBanks operates the medical centre and outreach without grant dependency for core services",
    "Revenue pathways: clinic subscriptions, district deployments, NGO programme monitoring, partner APIs",
    "Local capacity: train Ugandan developers and data stewards; document systems for MoH handoff",
    "Community ownership: CHWs and leaders co-design forms and alerts — not top-down IT",
]

PARTNERS = [
    "Ministry of Health community health strategy alignment",
    "District health teams for aggregate reporting validation",
    "Academic partners for evaluation and interoperability testing",
    "NGOs and CBOs already working with FairBanks outreach programmes",
]


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
        sizes, colors = {1: 20, 2: 14, 3: 12}, {1: NAVY, 2: TEAL, 3: SLATE}
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14 if level == 1 else 8)
        p.paragraph_format.space_after = Pt(6)
        font(p.add_run(text), size=sizes[level], bold=True, color=colors[level])

    def bullets(items):
        for it in items:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(3)
            p.clear()
            font(p.add_run(it))

    def table(headers, rows, widths=None):
        t = doc.add_table(rows=1 + len(rows), cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ""
            font(c.paragraphs[0].add_run(h), size=10, bold=True, color="FFFFFF")
            shade(c, TEAL)
            border(c, TEAL)
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                c = t.rows[ri + 1].cells[ci]
                c.text = ""
                font(c.paragraphs[0].add_run(str(val)), size=10)
                if ri % 2:
                    shade(c, CREAM)
                border(c)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Inches(w)
        doc.add_paragraph()

    def image(key, width_in=6.0, caption=None):
        path = embed(key)
        with PILImage.open(path) as im:
            iw, ih = im.size
        w = min(width_in, 3.4 * iw / ih)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(w))
        if caption:
            para(caption, size=9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=10)

    doc = Document()
    for m in doc.sections:
        m.left_margin = m.right_margin = Inches(0.85)
        m.top_margin = m.bottom_margin = Inches(0.75)

    para(PROGRAMME, size=12, bold=True, color=TEAL, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    para(DOC_TITLE, size=22, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    para(SUBTITLE, size=13, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=6)
    para(SLOGAN, size=12, bold=True, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, after=12)
    image("cover", caption="FairBanks — community health ecosystem rooted in Uganda")
    doc.add_page_break()

    heading("I. Statement of Interest")
    for p in STATEMENT:
        para(p)

    heading("II. Organizational Capacity")
    image("facility", width_in=5.6, caption="FairBanks Medical Centre — clinical anchor for community data")
    bullets(CAPACITY)

    heading("III. Alignment with Objective 4")
    table(["Element", "FairBanks / FCHIP response"], OBJ4_FIT, widths=[2.0, 4.4])

    heading("IV. Proposed Digital Health Activity")
    bullets(PROPOSED)
    image("architecture", caption="Community cascade → FCHIP → national digital health systems")

    heading("V. Interoperability Approach")
    table(["Platform", "Integration pathway"], INTEROP_ROWS, widths=[1.6, 4.8])
    image("dashboard", width_in=5.4, caption="Facility and programme dashboards for action, not data hoarding")

    heading("VI. Geographic Focus")
    bullets(GEO)
    image("outreach", width_in=5.6, caption="Live community outreach — pilot ground truth")

    heading("VII. Sustainability & Local Ownership")
    bullets(SUSTAIN)

    heading("VIII. Partnership Landscape")
    bullets(PARTNERS)
    image("gis", width_in=5.4, caption="GIS views support district planning and targeted outreach")

    heading("IX. Readiness for Full Proposal")
    para("This Statement of Interest confirms FairBanks' intent to respond to Objective 4 with a "
         "community-rooted, interoperable digital health proposal. We welcome dialogue with "
         "implementing partners and MoH-aligned reviewers before full application submission.")
    para("Official opportunity: DFOP0017890. Statement of Interest deadline per programme page.",
         size=9, color=MUTED, italic=True)
    para(CALL_URL, size=8, color=TEAL)
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
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
    from PIL import Image as PILImage

    navy, teal, accent = HexColor("#" + NAVY), HexColor("#" + TEAL), HexColor("#" + ACCENT)
    slate, muted, cream, line = HexColor("#" + SLATE), HexColor("#" + MUTED), HexColor("#" + CREAM), HexColor("#" + LINE)
    st = getSampleStyleSheet()
    for name, kw in [
        ("CoverTitle", dict(fontName="Helvetica-Bold", fontSize=19, leading=23, textColor=navy, alignment=TA_CENTER, spaceAfter=6)),
        ("H1", dict(fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=navy, spaceBefore=12, spaceAfter=6)),
        ("Body", dict(fontName="Helvetica", fontSize=10, leading=14, textColor=slate, alignment=TA_JUSTIFY, spaceAfter=6)),
        ("Meta", dict(fontName="Helvetica", fontSize=9, leading=12, textColor=muted, alignment=TA_CENTER)),
        ("FBullet", dict(fontName="Helvetica", fontSize=10, leading=13, textColor=slate, leftIndent=12, spaceAfter=3)),
        ("CellHead", dict(fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=white)),
        ("CellBody", dict(fontName="Helvetica", fontSize=8.5, leading=11, textColor=slate)),
    ]:
        st.add(ParagraphStyle(name, **kw))

    pw = A4[0] - 1.6 * inch
    story = []

    def img(key, w=pw * 0.88, cap=None, max_h=2.6 * inch):
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

    def tbl(headers, rows, widths=None):
        data = [[Paragraph(h, st["CellHead"]) for h in headers]]
        data += [[Paragraph(str(c), st["CellBody"]) for c in row] for row in rows]
        t = Table(data, colWidths=widths or [pw / len(headers)] * len(headers), repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), teal), ("GRID", (0, 0), (-1, -1), 0.4, line),
            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, cream]),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    story.append(Paragraph(PROGRAMME, st["Meta"]))
    story.append(Paragraph(DOC_TITLE, st["CoverTitle"]))
    story.append(Paragraph(f'<i><font color="#{ACCENT}">{SUBTITLE}</font></i>', st["Meta"]))
    story.append(Paragraph(f'<b><i><font color="#{ACCENT}">{SLOGAN}</font></i></b>', st["Meta"]))
    story.append(Spacer(1, 10))
    story.append(img("cover", cap="FairBanks — community health ecosystem rooted in Uganda"))
    story.append(PageBreak())

    story.append(Paragraph("I. Statement of Interest", st["H1"]))
    for p in STATEMENT:
        story.append(Paragraph(p, st["Body"]))

    story.append(Paragraph("II. Organizational Capacity", st["H1"]))
    for c in CAPACITY:
        story.append(Paragraph(f"• {c}", st["FBullet"]))

    story.append(Paragraph("III. Alignment with Objective 4", st["H1"]))
    tbl(["Element", "Response"], OBJ4_FIT, [pw * 0.28, pw * 0.72])

    story.append(Paragraph("IV. Proposed Digital Health Activity", st["H1"]))
    for p in PROPOSED:
        story.append(Paragraph(f"• {p}", st["FBullet"]))
    story.append(img("architecture", cap="Community cascade → FCHIP → national systems"))

    story.append(PageBreak())
    story.append(Paragraph("V. Interoperability Approach", st["H1"]))
    tbl(["Platform", "Integration pathway"], INTEROP_ROWS, [pw * 0.22, pw * 0.78])

    story.append(Paragraph("VI. Geographic Focus", st["H1"]))
    for g in GEO:
        story.append(Paragraph(f"• {g}", st["FBullet"]))

    story.append(Paragraph("VII. Sustainability & Local Ownership", st["H1"]))
    for s in SUSTAIN:
        story.append(Paragraph(f"• {s}", st["FBullet"]))

    story.append(Paragraph("VIII. Partnership Landscape", st["H1"]))
    for p in PARTNERS:
        story.append(Paragraph(f"• {p}", st["FBullet"]))

    story.append(Paragraph("IX. Readiness for Full Proposal", st["H1"]))
    story.append(Paragraph(
        "FairBanks confirms intent to respond to Objective 4 with a community-rooted, "
        "interoperable digital health proposal.", st["Body"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f'<b><i><font color="#{ACCENT}">{SLOGAN}</font></i></b>', st["Meta"]))

    OUT.mkdir(parents=True, exist_ok=True)
    SimpleDocTemplate(str(OUT_PDF), pagesize=A4, leftMargin=0.8 * inch, rightMargin=0.8 * inch,
                      topMargin=0.7 * inch, bottomMargin=0.7 * inch).build(story)
    print(f"PDF: {OUT_PDF}")


def build_pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    from PIL import Image as PILImage

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
        for i, ln in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = C(color)
            r.font.name = "Calibri"

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
        tb(sl, SW - Inches(0.8), SH - Inches(0.26), Inches(0.5), Inches(0.22), str(n), size=9, color="FFFFFF", align=PP_ALIGN.RIGHT)

    # Cover — formal SOI, no fact table
    s = prs.slides.add_slide(blank)
    pic(s, "cover", 0, 0, SW, SH)
    rect(s, 0, SH - Inches(3.5), SW, Inches(3.5), NAVY)
    tb(s, Inches(0.55), SH - Inches(3.2), Inches(12), Inches(0.35), PROGRAMME, size=12, bold=True, color=TEAL)
    tb(s, Inches(0.55), SH - Inches(2.65), Inches(12), Inches(0.85), DOC_TITLE, size=24, bold=True, color="FFFFFF")
    tb(s, Inches(0.55), SH - Inches(1.55), Inches(12), Inches(0.45), SUBTITLE, size=14, italic=True, color="F2C79B")
    tb(s, Inches(0.55), SH - Inches(0.95), Inches(12), Inches(0.35), SLOGAN, size=13, bold=True, color="FFFFFF")

    sections = [
        ("I", "Statement of Interest", STATEMENT[0], None),
        ("II", "Organizational Capacity", "\n".join(f"• {c}" for c in CAPACITY[:4]), "facility"),
        ("III", "Objective 4 Alignment", OBJ4_FIT[1][1], "architecture"),
        ("IV", "Interoperability", "eCHIS household capture → DHIS2 aggregates → facility EMR referrals", "mobile"),
        ("V", "Geographic Focus", GEO[0], "outreach"),
        ("VI", "Sustainability", SUSTAIN[0], "pharmacy"),
        ("VII", "Next Step", "Ready to develop full proposal with partners and MoH-aligned reviewers.", "gis"),
    ]

    for i, (roman, title, body, img_key) in enumerate(sections, start=2):
        s = prs.slides.add_slide(blank)
        header(s, roman, title)
        if img_key:
            pic(s, img_key, Inches(0.45), Inches(1.05), Inches(5.8), Inches(5.9))
            tb(s, Inches(6.5), Inches(1.15), Inches(6.4), Inches(5.6), body, size=15, color=SLATE)
        else:
            tb(s, Inches(0.55), Inches(1.15), Inches(12.2), Inches(5.8), body, size=17, color=SLATE)
        footer(s, i)

    OUT.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT_PPT))
    print(f"PPTX: {OUT_PPT}")


if __name__ == "__main__":
    build_docx()
    build_pdf()
    build_pptx()
    print("Done.", OUT)
