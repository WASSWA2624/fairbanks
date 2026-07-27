#!/usr/bin/env python3
"""Build five brief, clear CV designs (PDF + Word) for Racheal Nabukeera Sekagiri."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageDraw
from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    FrameBreak,
    KeepTogether,
    NextPageTemplate,
    PageTemplate,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    Image as RLImage,
)

ROOT = Path(__file__).resolve().parent
PHOTO_SRC = ROOT.parent / "WhatsApp Image 2026-07-27 at 14.23.53.jpeg"
PHOTO_DIR = ROOT / "_assets"
PHOTO_CIRCLE = PHOTO_DIR / "photo_circle.png"
PHOTO_SQUARE = PHOTO_DIR / "photo_square.png"
PREVIEW = ROOT / "_preview"

# Shared brief content
NAME = "RACHEAL NABUKEERA SEKAGIRI"
TITLE = "Senior Human Resource Executive | Human Capital Strategist"
LOCATION = "Kampala, Uganda"
MOBILE = "+256 772 849258 | +256 701 849258"
EMAIL = "info@fairbanksmedicalcentre.org"
WEB = "www.fairbanksmedicalcentre.org"

PROFILE_SHORT = (
    "Strategic HR executive with nearly <b>30 years</b> of experience and over "
    "<b>18 years</b> in executive HR leadership. Aligns people strategy with "
    "organisational performance across complex institutions."
)
PROFILE_MED = (
    "Founder and Managing Director, FairBanks Medical Centre Ltd. Formerly led "
    "Group HR for <b>2,000+ employees</b> across <b>50+ departments</b>. "
    "PhD Candidate (Management) researching AI prediction of occupational burnout."
)

HIGHLIGHTS = [
    "30 years HR and executive management experience",
    "18+ years executive Human Resource leadership",
    "HR leadership for 2,000+ employees across 50+ units",
    "Founder & MD, FairBanks Medical Centre Ltd",
    "Social enterprise reach: 10,000+ beneficiaries",
]

SKILLS = [
    "HR Strategy",
    "Organisational Development",
    "Talent Management",
    "Workforce Planning",
    "Change Management",
    "Performance Management",
    "Employee Relations",
    "Corporate Governance",
    "Leadership Development",
    "HR Analytics / HRIS",
]

ROLES = [
    {
        "title": "Founder, Managing Director & Executive HR Leader",
        "org": "FairBanks Medical Centre Ltd",
        "dates": "2025 - Present",
        "points": [
            "Lead strategy, governance, HR and institutional growth.",
            "Founded social enterprise initiative (10,000+ beneficiaries).",
        ],
    },
    {
        "title": "Group Human Resource & Administration Manager",
        "org": "Norvik Group",
        "dates": "Jul 2016 - Feb 2026",
        "points": [
            "Directed Group HR for 2,000+ employees across 50+ units.",
            "Led workforce planning, restructuring and leadership development.",
        ],
    },
    {
        "title": "Human Resource & Administration Manager",
        "org": "Norvik Hospital Ltd",
        "dates": "Sep 2013 - 2016",
        "points": [
            "Led recruitment, employee relations and performance systems.",
        ],
    },
    {
        "title": "Human Resource Manager",
        "org": "St. Catherine's Hospital",
        "dates": "2007 - 2010",
        "points": [
            "Managed HR operations, welfare and labour compliance.",
        ],
        "earlier": "Earlier: Health Systems Administrator, St. Catherine's (2010-2013); "
        "HR Assistant, E Power Limited (1997-2000).",
    },
]

EDUCATION = [
    ("PhD in Management (Ongoing)", "Uganda Christian University",
     "AI/ML research on occupational burnout in medical practice"),
    ("Master of Social Sector Planning and Management", "Makerere University", None),
    ("Bachelor of Arts in Social Sciences", "Makerere University", None),
]

# Colors
NAVY = HexColor("#0B3D5C")
TEAL = HexColor("#1A6B7A")
GOLD = HexColor("#C9A227")
INK = HexColor("#243B53")
MUTED = HexColor("#627D98")
LINE = HexColor("#D9E2EC")
SOFT = HexColor("#F0F4F8")
SIDEBAR = HexColor("#0B3D5C")
SIDEBAR2 = HexColor("#102A43")
CHARCOAL = HexColor("#1A202C")
WARM = HexColor("#F7F3EE")
FOREST = HexColor("#234E52")
SLATE = HexColor("#334E68")

DOCX_NAVY = RGBColor(0x0B, 0x3D, 0x5C)
DOCX_TEAL = RGBColor(0x1A, 0x6B, 0x7A)
DOCX_INK = RGBColor(0x24, 0x3B, 0x53)
DOCX_MUTED = RGBColor(0x62, 0x7D, 0x98)
DOCX_GOLD = RGBColor(0xC9, 0xA2, 0x27)


def ensure_dirs():
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW.mkdir(parents=True, exist_ok=True)
    for slug in VERSION_SLUGS:
        (ROOT / slug).mkdir(parents=True, exist_ok=True)


VERSION_SLUGS = [
    "01_classic_airy",
    "02_sidebar_modern",
    "03_one_page",
    "04_timeline",
    "05_editorial_split",
]


def prepare_photos():
    img = Image.open(PHOTO_SRC).convert("RGBA")
    w, h = img.size
    side = min(w, h)
    left = (w - side) // 2
    top = max(0, (h - side) // 2 - int(side * 0.08))
    crop = img.crop((left, top, left + side, top + side))

    # Circle
    size = 560
    c = crop.resize((size, size), Image.Resampling.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, size - 1, size - 1), fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(c, (0, 0))
    out.putalpha(mask)
    ring = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(ring).ellipse((3, 3, size - 4, size - 4), outline=(11, 61, 92, 255), width=10)
    Image.alpha_composite(out, ring).save(PHOTO_CIRCLE)

    # Soft rounded square
    size2 = 520
    s = crop.resize((size2, size2), Image.Resampling.LANCZOS)
    rad = 36
    mask2 = Image.new("L", (size2, size2), 0)
    ImageDraw.Draw(mask2).rounded_rectangle((0, 0, size2 - 1, size2 - 1), radius=rad, fill=255)
    out2 = Image.new("RGBA", (size2, size2), (0, 0, 0, 0))
    out2.paste(s, (0, 0))
    out2.putalpha(mask2)
    out2.save(PHOTO_SQUARE)


def styles_base():
    s = getSampleStyleSheet()
    # Clear defaults we will replace by unique names
    return s


def add_style(styles, name, **kw):
    if name in styles.byName:
        return styles[name]
    styles.add(ParagraphStyle(name=name, **kw))
    return styles[name]


def bullet(text, style):
    return Paragraph(f"- {text}", style)


def footer_fn(name_text="Racheal Nabukeera Sekagiri | CV"):
    def _draw(canv, doc):
        canv.saveState()
        canv.setStrokeColor(LINE)
        canv.setLineWidth(0.4)
        y = 10 * mm
        canv.line(doc.leftMargin, y, A4[0] - doc.rightMargin, y)
        canv.setFont("Helvetica", 7.5)
        canv.setFillColor(MUTED)
        canv.drawString(doc.leftMargin, 6 * mm, name_text)
        canv.drawRightString(A4[0] - doc.rightMargin, 6 * mm, f"Page {doc.page}")
        canv.restoreState()

    return _draw


# ---------------------------------------------------------------------------
# VERSION 01 - Classic Airy (2 pages, open spacing)
# ---------------------------------------------------------------------------
def build_v01_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "V1Name", fontName="Helvetica-Bold", fontSize=18, leading=21,
              textColor=NAVY, spaceAfter=3)
    add_style(styles, "V1Title", fontName="Helvetica", fontSize=9.5, leading=12,
              textColor=TEAL, spaceAfter=6)
    add_style(styles, "V1Contact", fontName="Helvetica", fontSize=8.5, leading=11,
              textColor=MUTED, spaceAfter=1)
    add_style(styles, "V1H", fontName="Helvetica-Bold", fontSize=10, leading=12,
              textColor=NAVY, spaceBefore=9, spaceAfter=4)
    add_style(styles, "V1Body", fontName="Helvetica", fontSize=9.2, leading=12.5,
              textColor=INK, alignment=TA_JUSTIFY, spaceAfter=5)
    add_style(styles, "V1Bullet", fontName="Helvetica", fontSize=9, leading=12,
              textColor=INK, leftIndent=8, spaceAfter=2)
    add_style(styles, "V1Job", fontName="Helvetica-Bold", fontSize=9.5, leading=12,
              textColor=NAVY, spaceBefore=6, spaceAfter=0)
    add_style(styles, "V1Org", fontName="Helvetica", fontSize=8.8, leading=11,
              textColor=TEAL, spaceAfter=2)
    add_style(styles, "V1Edu", fontName="Helvetica-Bold", fontSize=9.2, leading=11.5,
              textColor=NAVY, spaceBefore=2, spaceAfter=0)
    add_style(styles, "V1Meta", fontName="Helvetica", fontSize=8.5, leading=11,
              textColor=MUTED, spaceAfter=3)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=16 * mm, rightMargin=16 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"{NAME} - Classic Airy CV",
    )
    w = A4[0] - 32 * mm
    story = []

    photo = RLImage(str(PHOTO_CIRCLE), width=26 * mm, height=26 * mm)
    head = Table(
        [[
            [
                Paragraph(NAME, styles["V1Name"]),
                Paragraph(TITLE, styles["V1Title"]),
                Paragraph(f"{LOCATION}<br/>{MOBILE}<br/>{EMAIL}<br/>{WEB}", styles["V1Contact"]),
            ],
            photo,
        ]],
        colWidths=[w - 36 * mm, 36 * mm],
    )
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story += [head, Spacer(1, 6),
              HRFlowable(width="100%", thickness=2, color=NAVY, spaceAfter=1),
              HRFlowable(width="40%", thickness=1.2, color=GOLD, spaceAfter=10)]

    story += [
        Paragraph("PROFILE", styles["V1H"]),
        HRFlowable(width=40, thickness=1.5, color=GOLD, spaceAfter=8),
        Paragraph(PROFILE_SHORT, styles["V1Body"]),
        Paragraph(PROFILE_MED, styles["V1Body"]),
        Paragraph("CORE COMPETENCIES", styles["V1H"]),
        HRFlowable(width=40, thickness=1.5, color=GOLD, spaceAfter=8),
        Paragraph("  ·  ".join(SKILLS), styles["V1Body"]),
        Paragraph("EXPERIENCE", styles["V1H"]),
        HRFlowable(width=40, thickness=1.5, color=GOLD, spaceAfter=4),
    ]
    for role in ROLES:
        block = [
            Paragraph(role["title"], styles["V1Job"]),
            Paragraph(f"{role['org']}  |  {role['dates']}", styles["V1Org"]),
        ]
        for p in role["points"]:
            block.append(bullet(p, styles["V1Bullet"]))
        if role.get("earlier"):
            block.append(Paragraph(role["earlier"], styles["V1Meta"]))
        story.append(KeepTogether(block))

    story += [
        Paragraph("EDUCATION", styles["V1H"]),
        HRFlowable(width=40, thickness=1.5, color=GOLD, spaceAfter=8),
    ]
    for deg, uni, note in EDUCATION:
        story.append(Paragraph(deg, styles["V1Edu"]))
        story.append(Paragraph(uni if not note else f"{uni} - {note}", styles["V1Meta"]))

    story += [
        Paragraph("ADDITIONAL", styles["V1H"]),
        HRFlowable(width=40, thickness=1.5, color=GOLD, spaceAfter=8),
        Paragraph(
            "Member, Federation of Uganda Employers (FUE)  ·  "
            "English (Fluent), Luganda (Fluent)  ·  "
            "References available on request",
            styles["V1Body"],
        ),
    ]
    doc.build(story, onFirstPage=footer_fn("Classic Airy"), onLaterPages=footer_fn("Classic Airy"))


def build_v01_docx(out: Path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.6)
        sec.bottom_margin = Cm(1.6)
        sec.left_margin = Cm(1.8)
        sec.right_margin = Cm(1.8)

    t = doc.add_table(rows=1, cols=2)
    left, right = t.rows[0].cells
    p = left.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 18, True, DOCX_NAVY)
    p = left.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 10, False, DOCX_TEAL)
    for line in (LOCATION, f"Mobile: {MOBILE}", f"Email: {EMAIL}", f"Web: {WEB}"):
        p = left.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        _run(r, 9, False, DOCX_MUTED)
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.paragraphs[0].add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.15))

    _h(doc, "PROFILE")
    _p(doc, PROFILE_SHORT.replace("<b>", "").replace("</b>", ""))
    _p(doc, PROFILE_MED.replace("<b>", "").replace("</b>", ""))
    _h(doc, "CORE COMPETENCIES")
    _p(doc, "  ·  ".join(SKILLS))
    _h(doc, "EXPERIENCE")
    for role in ROLES:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        r = p.add_run(role["title"])
        _run(r, 11, True, DOCX_NAVY)
        p = doc.add_paragraph()
        r = p.add_run(f"{role['org']}  |  {role['dates']}")
        _run(r, 10, False, DOCX_TEAL)
        for pt in role["points"]:
            _b(doc, pt)
        if role.get("earlier"):
            _p(doc, role["earlier"], muted=True)
    _h(doc, "EDUCATION")
    for deg, uni, note in EDUCATION:
        p = doc.add_paragraph()
        r = p.add_run(deg)
        _run(r, 10, True, DOCX_NAVY)
        _p(doc, uni if not note else f"{uni} - {note}", muted=True)
    _h(doc, "ADDITIONAL")
    _p(doc, "Member, FUE  ·  English & Luganda (Fluent)  ·  References on request")
    doc.save(str(out))


# ---------------------------------------------------------------------------
# VERSION 02 - Sidebar Modern
# ---------------------------------------------------------------------------
def build_v02_pdf(out: Path):
    width, height = A4
    side_w = 68 * mm
    margin = 11 * mm
    c = pdfcanvas.Canvas(str(out), pagesize=A4)

    def draw_sidebar(page_num: int):
        c.setFillColor(SIDEBAR)
        c.rect(0, 0, side_w, height, fill=1, stroke=0)
        c.setFillColor(HexColor("#0A3048"))
        c.rect(0, height - 56 * mm, side_w, 56 * mm, fill=1, stroke=0)

        c.drawImage(
            str(PHOTO_CIRCLE),
            (side_w - 32 * mm) / 2,
            height - 48 * mm,
            width=32 * mm,
            height=32 * mm,
            mask="auto",
        )

        y = height - 62 * mm
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(7 * mm, y, "CONTACT")
        y -= 4 * mm
        c.setStrokeColor(GOLD)
        c.setLineWidth(1)
        c.line(7 * mm, y + 1.5 * mm, side_w - 7 * mm, y + 1.5 * mm)
        y -= 4 * mm
        c.setFillColor(white)
        c.setFont("Helvetica", 7.4)
        contact_lines = [
            LOCATION,
            "+256 772 849258",
            "+256 701 849258",
            "info@fairbanks",
            "medicalcentre.org",
            "www.fairbanks",
            "medicalcentre.org",
        ]
        for line in contact_lines:
            c.drawString(7 * mm, y, line)
            y -= 3.8 * mm

        y -= 3 * mm
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(7 * mm, y, "STRENGTHS")
        y -= 4 * mm
        c.setStrokeColor(GOLD)
        c.line(7 * mm, y + 1.5 * mm, side_w - 7 * mm, y + 1.5 * mm)
        y -= 3 * mm
        c.setFillColor(white)
        c.setFont("Helvetica", 7.6)
        for sk in SKILLS:
            c.drawString(7 * mm, y, f"- {sk}")
            y -= 4.0 * mm

        y -= 3 * mm
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(7 * mm, y, "LANGUAGES")
        y -= 4 * mm
        c.setStrokeColor(GOLD)
        c.line(7 * mm, y + 1.5 * mm, side_w - 7 * mm, y + 1.5 * mm)
        y -= 3 * mm
        c.setFillColor(white)
        c.setFont("Helvetica", 7.6)
        c.drawString(7 * mm, y, "English - Fluent")
        y -= 3.8 * mm
        c.drawString(7 * mm, y, "Luganda - Fluent")

        y -= 6 * mm
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(7 * mm, y, "MEMBERSHIP")
        y -= 4 * mm
        c.setStrokeColor(GOLD)
        c.line(7 * mm, y + 1.5 * mm, side_w - 7 * mm, y + 1.5 * mm)
        y -= 3 * mm
        c.setFillColor(white)
        c.setFont("Helvetica", 7.6)
        c.drawString(7 * mm, y, "FUE Member")

        c.setFillColor(HexColor("#8DA2B5"))
        c.setFont("Helvetica", 7)
        c.drawString(7 * mm, 8 * mm, f"Sidebar Modern  ·  {page_num}")

    draw_sidebar(1)
    x = side_w + margin
    max_w = width - side_w - 2 * margin
    y = height - 16 * mm

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 13.5)
    for line in _wrap_canvas(c, NAME, max_w, "Helvetica-Bold", 13.5):
        c.drawString(x, y, line)
        y -= 5.5 * mm
    c.setFillColor(TEAL)
    c.setFont("Helvetica", 8.5)
    for line in _wrap_canvas(c, TITLE, max_w, "Helvetica", 8.5):
        c.drawString(x, y, line)
        y -= 4.2 * mm

    y -= 3 * mm
    y = _section(c, x, y, max_w, "PROFILE")
    y = _para(c, x, y, max_w, PROFILE_SHORT.replace("<b>", "").replace("</b>", ""), 9, 12)
    y = _para(c, x, y, max_w, PROFILE_MED.replace("<b>", "").replace("</b>", ""), 9, 12)

    y -= 2 * mm
    y = _section(c, x, y, max_w, "EXPERIENCE")

    for i, role in enumerate(ROLES):
        if y < 40 * mm:
            c.showPage()
            draw_sidebar(2)
            y = height - 16 * mm
            y = _section(c, x, y, max_w, "EXPERIENCE (CONTINUED)")
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 9)
        for line in _wrap_canvas(c, role["title"], max_w, "Helvetica-Bold", 9):
            c.drawString(x, y, line)
            y -= 3.8 * mm
        c.setFillColor(TEAL)
        c.setFont("Helvetica", 8)
        c.drawString(x, y, f"{role['org']}  |  {role['dates']}")
        y -= 4 * mm
        for pt in role["points"]:
            y = _para(c, x, y, max_w, f"- {pt}", 8.5, 11)
        if role.get("earlier"):
            y = _para(c, x, y, max_w, role["earlier"], 7.8, 10, MUTED)
        y -= 2.5 * mm

    if y < 50 * mm:
        c.showPage()
        draw_sidebar(2)
        y = height - 16 * mm
    y = _section(c, x, y, max_w, "EDUCATION")
    for deg, uni, note in EDUCATION:
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(x, y, deg)
        y -= 3.5 * mm
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 7.8)
        c.drawString(x, y, uni)
        y -= 3.2 * mm
        if note:
            y = _para(c, x, y, max_w, note, 7.6, 10, MUTED)
        y -= 1.2 * mm

    c.setFillColor(MUTED)
    c.setFont("Helvetica-Oblique", 8)
    c.drawString(x, max(14 * mm, y - 2 * mm), "References available on request.")
    c.save()

def _wrap_canvas(c, text, max_w, font, size):
    c.setFont(font, size)
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def _section(c, x, y, max_w, title):
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x, y, title)
    y -= 2.5 * mm
    c.setStrokeColor(GOLD)
    c.setLineWidth(1.4)
    c.line(x, y, x + 28 * mm, y)
    return y - 4.5 * mm


def _para(c, x, y, max_w, text, size, leading, color=INK):
    c.setFillColor(color)
    c.setFont("Helvetica", size)
    for line in _wrap_canvas(c, text, max_w, "Helvetica", size):
        c.drawString(x, y, line)
        y -= leading
    return y - 2


def build_v02_docx(out: Path):
    # Approximate sidebar feel with 2-col table
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.0)
        sec.bottom_margin = Cm(1.0)
        sec.left_margin = Cm(1.0)
        sec.right_margin = Cm(1.0)

    table = doc.add_table(rows=1, cols=2)
    side, main = table.rows[0].cells
    side.width = Cm(5.8)
    main.width = Cm(13.5)
    _shade(side, "0B3D5C")

    sp = side.paragraphs[0]
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.2))

    def side_h(text):
        p = side.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        r = p.add_run(text)
        _run(r, 9, True, DOCX_GOLD)

    def side_t(text):
        p = side.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(text)
        _run(r, 8.5, False, RGBColor(255, 255, 255))

    side_h("CONTACT")
    for line in [LOCATION, MOBILE, EMAIL, WEB]:
        side_t(line)
    side_h("STRENGTHS")
    for sk in SKILLS:
        side_t(f"- {sk}")
    side_h("LANGUAGES")
    side_t("English - Fluent")
    side_t("Luganda - Fluent")
    side_h("MEMBERSHIP")
    side_t("FUE Member")

    p = main.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 15, True, DOCX_NAVY)
    p = main.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9.5, False, DOCX_TEAL)
    _h(doc, "PROFILE", cell=main)
    _p_in(main, PROFILE_SHORT.replace("<b>", "").replace("</b>", ""))
    _p_in(main, PROFILE_MED.replace("<b>", "").replace("</b>", ""))
    _h(doc, "HIGHLIGHTS", cell=main)
    for h in HIGHLIGHTS:
        _b_in(main, h)
    _h(doc, "EXPERIENCE", cell=main)
    for role in ROLES:
        p = main.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        r = p.add_run(role["title"])
        _run(r, 10, True, DOCX_NAVY)
        p = main.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 9, False, DOCX_TEAL)
        for pt in role["points"]:
            _b_in(main, pt)
    _h(doc, "EDUCATION", cell=main)
    for deg, uni, note in EDUCATION:
        p = main.add_paragraph()
        r = p.add_run(deg)
        _run(r, 9.5, True, DOCX_NAVY)
        _p_in(main, uni if not note else f"{uni} - {note}", muted=True)
    _p_in(main, "References available on request.", muted=True)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# VERSION 03 - One Page
# ---------------------------------------------------------------------------
def build_v03_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "V3Name", fontName="Helvetica-Bold", fontSize=15, leading=17,
              textColor=CHARCOAL, spaceAfter=1)
    add_style(styles, "V3Title", fontName="Helvetica", fontSize=8.5, leading=11,
              textColor=FOREST, spaceAfter=2)
    add_style(styles, "V3Contact", fontName="Helvetica", fontSize=7.8, leading=10,
              textColor=MUTED, alignment=TA_LEFT)
    add_style(styles, "V3H", fontName="Helvetica-Bold", fontSize=8.5, leading=10,
              textColor=FOREST, spaceBefore=7, spaceAfter=3)
    add_style(styles, "V3Body", fontName="Helvetica", fontSize=8.2, leading=11,
              textColor=INK, alignment=TA_JUSTIFY, spaceAfter=3)
    add_style(styles, "V3Bullet", fontName="Helvetica", fontSize=8, leading=10.5,
              textColor=INK, leftIndent=6, spaceAfter=1.2)
    add_style(styles, "V3Job", fontName="Helvetica-Bold", fontSize=8.3, leading=10,
              textColor=CHARCOAL, spaceBefore=4, spaceAfter=0)
    add_style(styles, "V3Org", fontName="Helvetica", fontSize=7.8, leading=10,
              textColor=FOREST, spaceAfter=1.5)
    add_style(styles, "V3Small", fontName="Helvetica", fontSize=7.6, leading=10,
              textColor=MUTED, spaceAfter=2)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=11 * mm, bottomMargin=11 * mm,
        title=f"{NAME} - One Page CV",
    )
    w = A4[0] - 28 * mm
    story = []
    photo = RLImage(str(PHOTO_SQUARE), width=22 * mm, height=22 * mm)
    head = Table([[
        [
            Paragraph(NAME, styles["V3Name"]),
            Paragraph(TITLE, styles["V3Title"]),
            Paragraph(f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}", styles["V3Contact"]),
        ],
        photo,
    ]], colWidths=[w - 26 * mm, 26 * mm])
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story += [head, Spacer(1, 3),
              HRFlowable(width="100%", thickness=1.5, color=FOREST, spaceAfter=4)]

    story += [
        Paragraph("PROFILE", styles["V3H"]),
        Paragraph(PROFILE_SHORT + " " + PROFILE_MED, styles["V3Body"]),
    ]

    # Highlights in 2 cols
    left_h = [bullet(h, styles["V3Bullet"]) for h in HIGHLIGHTS[:3]]
    right_h = [bullet(h, styles["V3Bullet"]) for h in HIGHLIGHTS[3:]]
    story += [
        Paragraph("HIGHLIGHTS", styles["V3H"]),
        Table([[left_h, right_h]], colWidths=[w / 2, w / 2],
              style=TableStyle([
                  ("VALIGN", (0, 0), (-1, -1), "TOP"),
                  ("LEFTPADDING", (0, 0), (-1, -1), 0),
                  ("RIGHTPADDING", (0, 0), (-1, -1), 3),
              ])),
        Paragraph("SKILLS", styles["V3H"]),
        Paragraph("  ·  ".join(SKILLS), styles["V3Body"]),
        Paragraph("EXPERIENCE", styles["V3H"]),
    ]

    # Only top 3 roles + earlier note for one-pager
    for role in ROLES[:3]:
        story.append(Paragraph(role["title"], styles["V3Job"]))
        story.append(Paragraph(f"{role['org']}  |  {role['dates']}", styles["V3Org"]))
        for pt in role["points"][:2]:
            story.append(bullet(pt, styles["V3Bullet"]))
    story.append(Paragraph(
        "Earlier: HR Manager & Systems Administrator, St. Catherine's (2007-2013); "
        "HR Assistant, E Power Limited (1997-2000).",
        styles["V3Small"],
    ))

    edu_line = (
        "<b>PhD Management (Ongoing)</b>, Uganda Christian University  ·  "
        "<b>MSSPM</b>, Makerere  ·  <b>BA Social Sciences</b>, Makerere"
    )
    story += [
        Paragraph("EDUCATION & MORE", styles["V3H"]),
        Paragraph(edu_line, styles["V3Body"]),
        Paragraph(
            "FUE Member  ·  English & Luganda (Fluent)  ·  References on request",
            styles["V3Small"],
        ),
    ]
    doc.build(story)


def build_v03_docx(out: Path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.1)
        sec.bottom_margin = Cm(1.1)
        sec.left_margin = Cm(1.4)
        sec.right_margin = Cm(1.4)

    t = doc.add_table(rows=1, cols=2)
    left, right = t.rows[0].cells
    p = left.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 14, True, RGBColor(0x1A, 0x20, 0x2C))
    p = left.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9, False, RGBColor(0x23, 0x4E, 0x52))
    p = left.add_paragraph()
    r = p.add_run(f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}")
    _run(r, 8, False, DOCX_MUTED)
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.paragraphs[0].add_run().add_picture(str(PHOTO_SQUARE), width=Inches(0.9))

    _h(doc, "PROFILE", size=9)
    _p(doc, (PROFILE_SHORT + " " + PROFILE_MED).replace("<b>", "").replace("</b>", ""), size=9)
    _h(doc, "HIGHLIGHTS", size=9)
    for h in HIGHLIGHTS:
        _b(doc, h, size=8.5)
    _h(doc, "SKILLS", size=9)
    _p(doc, "  ·  ".join(SKILLS), size=9)
    _h(doc, "EXPERIENCE", size=9)
    for role in ROLES[:3]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        r = p.add_run(role["title"])
        _run(r, 9, True, DOCX_INK)
        p = doc.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 8.5, False, RGBColor(0x23, 0x4E, 0x52))
        for pt in role["points"][:2]:
            _b(doc, pt, size=8.5)
    _p(doc,
       "Earlier: HR Manager & Systems Administrator, St. Catherine's (2007-2013); "
       "HR Assistant, E Power Limited (1997-2000).",
       muted=True, size=8)
    _h(doc, "EDUCATION & MORE", size=9)
    _p(doc,
       "PhD Management (Ongoing), Uganda Christian University · "
       "MSSPM, Makerere · BA Social Sciences, Makerere",
       size=9)
    _p(doc, "FUE Member · English & Luganda (Fluent) · References on request",
       muted=True, size=8)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# VERSION 04 - Timeline
# ---------------------------------------------------------------------------
def build_v04_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "V4Name", fontName="Helvetica-Bold", fontSize=18, leading=22,
              textColor=SLATE, alignment=TA_CENTER, spaceAfter=3)
    add_style(styles, "V4Title", fontName="Helvetica", fontSize=9.5, leading=12,
              textColor=TEAL, alignment=TA_CENTER, spaceAfter=4)
    add_style(styles, "V4Contact", fontName="Helvetica", fontSize=8.5, leading=11,
              textColor=MUTED, alignment=TA_CENTER, spaceAfter=8)
    add_style(styles, "V4H", fontName="Helvetica-Bold", fontSize=10, leading=12,
              textColor=SLATE, alignment=TA_CENTER, spaceBefore=10, spaceAfter=6)
    add_style(styles, "V4Body", fontName="Helvetica", fontSize=9.5, leading=13.5,
              textColor=INK, alignment=TA_LEFT, spaceAfter=6)
    add_style(styles, "V4Date", fontName="Helvetica-Bold", fontSize=8.5, leading=11,
              textColor=TEAL, alignment=TA_RIGHT)
    add_style(styles, "V4Job", fontName="Helvetica-Bold", fontSize=9.5, leading=12,
              textColor=SLATE)
    add_style(styles, "V4Org", fontName="Helvetica", fontSize=8.5, leading=11,
              textColor=MUTED, spaceAfter=2)
    add_style(styles, "V4Bullet", fontName="Helvetica", fontSize=8.8, leading=12,
              textColor=INK, leftIndent=4, spaceAfter=1.5)
    add_style(styles, "V4Edu", fontName="Helvetica", fontSize=9, leading=12,
              textColor=INK, alignment=TA_CENTER, spaceAfter=3)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=16 * mm, rightMargin=16 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
        title=f"{NAME} - Timeline CV",
    )
    w = A4[0] - 32 * mm
    story = [
        Paragraph(NAME, styles["V4Name"]),
        Paragraph(TITLE, styles["V4Title"]),
        Paragraph(f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}", styles["V4Contact"]),
    ]
    # Centered photo
    photo = RLImage(str(PHOTO_CIRCLE), width=28 * mm, height=28 * mm)
    ph = Table([[photo]], colWidths=[w])
    ph.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    story += [ph, Spacer(1, 6),
              HRFlowable(width="30%", thickness=1.2, color=GOLD, spaceBefore=2, spaceAfter=8)]

    story += [
        Paragraph("PROFILE", styles["V4H"]),
        Paragraph(PROFILE_SHORT.replace("<b>", "").replace("</b>", ""), styles["V4Body"]),
        Paragraph(PROFILE_MED.replace("<b>", "").replace("</b>", ""), styles["V4Body"]),
        Paragraph("CAREER TIMELINE", styles["V4H"]),
        HRFlowable(width="20%", thickness=1, color=GOLD, spaceAfter=8),
    ]

    for role in ROLES:
        left = Paragraph(role["dates"].replace(" - ", "<br/>"), styles["V4Date"])
        right = [
            Paragraph(role["title"], styles["V4Job"]),
            Paragraph(role["org"], styles["V4Org"]),
        ]
        for pt in role["points"]:
            right.append(bullet(pt, styles["V4Bullet"]))
        if role.get("earlier"):
            right.append(Paragraph(role["earlier"], styles["V4Org"]))
        row = Table([[left, right]], colWidths=[28 * mm, w - 28 * mm])
        row.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("RIGHTPADDING", (0, 0), (0, 0), 8),
            ("LEFTPADDING", (1, 0), (1, 0), 10),
            ("LINEBEFORE", (1, 0), (1, 0), 1.5, TEAL),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (1, 0), (1, 0), SOFT),
        ]))
        story.append(row)

    story += [
        Paragraph("EDUCATION", styles["V4H"]),
        Paragraph(
            "PhD Management (Ongoing) - Uganda Christian University<br/>"
            "Master of Social Sector Planning and Management - Makerere University<br/>"
            "Bachelor of Arts in Social Sciences - Makerere University",
            styles["V4Edu"],
        ),
        Paragraph(
            "Skills: " + " · ".join(SKILLS[:8]) + "<br/>"
            "FUE Member · English & Luganda (Fluent) · References on request",
            styles["V4Edu"],
        ),
    ]
    doc.build(story, onFirstPage=footer_fn("Timeline"), onLaterPages=footer_fn("Timeline"))


def build_v04_docx(out: Path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.4)
        sec.bottom_margin = Cm(1.4)
        sec.left_margin = Cm(1.6)
        sec.right_margin = Cm(1.6)

    for text, size, bold, color, center in [
        (NAME, 16, True, RGBColor(0x33, 0x4E, 0x68), True),
        (TITLE, 10, False, DOCX_TEAL, True),
        (f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}", 9, False, DOCX_MUTED, True),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        _run(r, size, bold, color)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.05))

    _h(doc, "PROFILE", center=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(PROFILE_SHORT.replace("<b>", "").replace("</b>", ""))
    _run(r, 9.5, False, DOCX_INK)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(PROFILE_MED.replace("<b>", "").replace("</b>", ""))
    _run(r, 9.5, False, DOCX_INK)

    _h(doc, "CAREER TIMELINE", center=True)
    for role in ROLES:
        t = doc.add_table(rows=1, cols=2)
        d, body = t.rows[0].cells
        d.width = Cm(3.2)
        body.width = Cm(14)
        p = d.paragraphs[0]
        r = p.add_run(role["dates"])
        _run(r, 8.5, True, DOCX_TEAL)
        p = body.paragraphs[0]
        r = p.add_run(role["title"])
        _run(r, 10, True, RGBColor(0x33, 0x4E, 0x68))
        p = body.add_paragraph()
        r = p.add_run(role["org"])
        _run(r, 9, False, DOCX_MUTED)
        for pt in role["points"]:
            _b_in(body, pt)
        doc.add_paragraph()

    _h(doc, "EDUCATION", center=True)
    for deg, uni, _ in EDUCATION:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"{deg} - {uni}")
        _run(r, 9, False, DOCX_INK)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("FUE Member · English & Luganda (Fluent) · References on request")
    _run(r, 9, False, DOCX_MUTED)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# VERSION 05 - Editorial Split
# ---------------------------------------------------------------------------
def build_v05_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "V5Name", fontName="Helvetica-Bold", fontSize=22, leading=26,
              textColor=NAVY, spaceAfter=2)
    add_style(styles, "V5Title", fontName="Helvetica", fontSize=10, leading=13,
              textColor=TEAL, spaceAfter=6)
    add_style(styles, "V5Tag", fontName="Helvetica-Bold", fontSize=8, leading=10,
              textColor=GOLD, spaceAfter=8)
    add_style(styles, "V5H", fontName="Helvetica-Bold", fontSize=9.5, leading=12,
              textColor=NAVY, spaceBefore=2, spaceAfter=5)
    add_style(styles, "V5Body", fontName="Helvetica", fontSize=9, leading=12.5,
              textColor=INK, alignment=TA_JUSTIFY, spaceAfter=5)
    add_style(styles, "V5Bullet", fontName="Helvetica", fontSize=8.6, leading=11.5,
              textColor=INK, leftIndent=4, spaceAfter=2)
    add_style(styles, "V5Job", fontName="Helvetica-Bold", fontSize=9.2, leading=11.5,
              textColor=NAVY, spaceBefore=5, spaceAfter=0)
    add_style(styles, "V5Org", fontName="Helvetica", fontSize=8.2, leading=10.5,
              textColor=TEAL, spaceAfter=2)
    add_style(styles, "V5Small", fontName="Helvetica", fontSize=8, leading=10.5,
              textColor=MUTED, spaceAfter=3)
    add_style(styles, "V5Skill", fontName="Helvetica", fontSize=8.3, leading=12,
              textColor=INK, spaceAfter=1)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
        title=f"{NAME} - Editorial Split CV",
    )
    w = A4[0] - 30 * mm
    story = []

    # Banner header on warm background via table
    photo = RLImage(str(PHOTO_CIRCLE), width=32 * mm, height=32 * mm)
    banner = Table([[
        [
            Paragraph("CURRICULUM VITAE", styles["V5Tag"]),
            Paragraph(NAME, styles["V5Name"]),
            Paragraph(TITLE, styles["V5Title"]),
            Paragraph(f"{LOCATION}  ·  {MOBILE}<br/>{EMAIL}  ·  {WEB}", styles["V5Small"]),
        ],
        photo,
    ]], colWidths=[w - 38 * mm, 38 * mm])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WARM),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story += [banner, Spacer(1, 10)]

    # Two columns: left narrow (highlights/skills/edu), right (profile + experience)
    left_bits = [
        Paragraph("AT A GLANCE", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=5),
    ]
    for h in HIGHLIGHTS:
        left_bits.append(bullet(h, styles["V5Bullet"]))
    left_bits += [
        Spacer(1, 8),
        Paragraph("SKILLS", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=5),
    ]
    for sk in SKILLS:
        left_bits.append(Paragraph(f"- {sk}", styles["V5Skill"]))
    left_bits += [
        Spacer(1, 8),
        Paragraph("EDUCATION", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=5),
        Paragraph("<b>PhD Management</b> (Ongoing)<br/>Uganda Christian University", styles["V5Small"]),
        Paragraph("<b>MSSPM</b><br/>Makerere University", styles["V5Small"]),
        Paragraph("<b>BA Social Sciences</b><br/>Makerere University", styles["V5Small"]),
        Spacer(1, 6),
        Paragraph("MORE", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=5),
        Paragraph("FUE Member<br/>English & Luganda - Fluent<br/>References on request", styles["V5Small"]),
    ]

    right_bits = [
        Paragraph("PROFILE", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=5),
        Paragraph(PROFILE_SHORT, styles["V5Body"]),
        Paragraph(PROFILE_MED, styles["V5Body"]),
        Spacer(1, 4),
        Paragraph("EXPERIENCE", styles["V5H"]),
        HRFlowable(width="100%", thickness=0.8, color=GOLD, spaceAfter=4),
    ]
    for role in ROLES:
        right_bits.append(Paragraph(role["title"], styles["V5Job"]))
        right_bits.append(Paragraph(f"{role['org']}  |  {role['dates']}", styles["V5Org"]))
        for pt in role["points"]:
            right_bits.append(bullet(pt, styles["V5Bullet"]))
        if role.get("earlier"):
            right_bits.append(Paragraph(role["earlier"], styles["V5Small"]))

    split = Table([[left_bits, right_bits]], colWidths=[w * 0.34, w * 0.66])
    split.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (0, 0), SOFT),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
        ("RIGHTPADDING", (0, 0), (0, 0), 8),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, 0), (0, 0), 8),
        ("LEFTPADDING", (1, 0), (1, 0), 12),
        ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (1, 0), (1, 0), 0),
        ("LINEAFTER", (0, 0), (0, 0), 0.6, LINE),
    ]))
    story.append(split)
    doc.build(story, onFirstPage=footer_fn("Editorial Split"), onLaterPages=footer_fn("Editorial Split"))


def build_v05_docx(out: Path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.3)
        sec.bottom_margin = Cm(1.3)
        sec.left_margin = Cm(1.4)
        sec.right_margin = Cm(1.4)

    head = doc.add_table(rows=1, cols=2)
    left, right = head.rows[0].cells
    _shade(left, "F7F3EE")
    _shade(right, "F7F3EE")
    p = left.paragraphs[0]
    r = p.add_run("CURRICULUM VITAE")
    _run(r, 8, True, DOCX_GOLD)
    p = left.add_paragraph()
    r = p.add_run(NAME)
    _run(r, 16, True, DOCX_NAVY)
    p = left.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9.5, False, DOCX_TEAL)
    p = left.add_paragraph()
    r = p.add_run(f"{LOCATION} · {MOBILE} · {EMAIL}")
    _run(r, 8.5, False, DOCX_MUTED)
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.paragraphs[0].add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.15))

    doc.add_paragraph()
    body = doc.add_table(rows=1, cols=2)
    L, R = body.rows[0].cells
    L.width = Cm(6)
    R.width = Cm(12)
    _shade(L, "F0F4F8")

    def lh(cell, text):
        p = cell.add_paragraph() if cell.paragraphs[0].text else cell.paragraphs[0]
        if cell.paragraphs[0].text:
            p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        r = p.add_run(text)
        _run(r, 9, True, DOCX_NAVY)

    # Clear first empty and build left
    L.paragraphs[0].clear()
    p = L.paragraphs[0]
    r = p.add_run("AT A GLANCE")
    _run(r, 9, True, DOCX_NAVY)
    for h in HIGHLIGHTS:
        _b_in(L, h, size=8.5)
    p = L.add_paragraph()
    r = p.add_run("SKILLS")
    _run(r, 9, True, DOCX_NAVY)
    for sk in SKILLS:
        _b_in(L, sk, size=8.5)
    p = L.add_paragraph()
    r = p.add_run("EDUCATION")
    _run(r, 9, True, DOCX_NAVY)
    for deg, uni, _ in EDUCATION:
        _p_in(L, f"{deg} - {uni}", size=8.5)
    p = L.add_paragraph()
    r = p.add_run("MORE")
    _run(r, 9, True, DOCX_NAVY)
    _p_in(L, "FUE Member · English & Luganda · References on request", size=8.5)

    R.paragraphs[0].clear()
    p = R.paragraphs[0]
    r = p.add_run("PROFILE")
    _run(r, 9, True, DOCX_NAVY)
    _p_in(R, PROFILE_SHORT.replace("<b>", "").replace("</b>", ""))
    _p_in(R, PROFILE_MED.replace("<b>", "").replace("</b>", ""))
    p = R.add_paragraph()
    r = p.add_run("EXPERIENCE")
    _run(r, 9, True, DOCX_NAVY)
    for role in ROLES:
        p = R.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        r = p.add_run(role["title"])
        _run(r, 9.5, True, DOCX_NAVY)
        p = R.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 8.5, False, DOCX_TEAL)
        for pt in role["points"]:
            _b_in(R, pt, size=8.5)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# DOCX helpers
# ---------------------------------------------------------------------------
def _run(run, size=10, bold=False, color=None):
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _shade(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _h(doc, text, size=11, center=False, cell=None):
    target = cell if cell is not None else doc
    p = target.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    _run(r, size, True, DOCX_NAVY)
    # gold bottom border
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "10")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "C9A227")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _p(doc, text, muted=False, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    _run(r, size, False, DOCX_MUTED if muted else DOCX_INK)


def _b(doc, text, size=9.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    _run(r, size, False, DOCX_INK)


def _p_in(cell, text, muted=False, size=9.5):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    _run(r, size, False, DOCX_MUTED if muted else DOCX_INK)


def _b_in(cell, text, size=9):
    p = cell.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    _run(r, size, False, DOCX_INK)


# ---------------------------------------------------------------------------
# Preview + main
# ---------------------------------------------------------------------------
def render_previews():
    import fitz

    for slug in VERSION_SLUGS:
        pdf = ROOT / slug / "racheal_cv.pdf"
        if not pdf.exists():
            continue
        out_dir = PREVIEW / slug
        out_dir.mkdir(parents=True, exist_ok=True)
        for old in out_dir.glob("*.png"):
            old.unlink()
        doc = fitz.open(str(pdf))
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
            pix.save(str(out_dir / f"page_{i}.png"))
        print(f"  preview {slug}: {doc.page_count} page(s)")
        doc.close()


def main():
    if not PHOTO_SRC.exists():
        raise SystemExit(f"Missing photo: {PHOTO_SRC}")
    ensure_dirs()
    prepare_photos()

    builders = [
        ("01_classic_airy", build_v01_pdf, build_v01_docx),
        ("02_sidebar_modern", build_v02_pdf, build_v02_docx),
        ("03_one_page", build_v03_pdf, build_v03_docx),
        ("04_timeline", build_v04_pdf, build_v04_docx),
        ("05_editorial_split", build_v05_pdf, build_v05_docx),
    ]
    for slug, pdf_fn, docx_fn in builders:
        folder = ROOT / slug
        pdf_path = folder / "racheal_cv.pdf"
        docx_path = folder / "racheal_cv.docx"
        print(f"Building {slug}...")
        pdf_fn(pdf_path)
        docx_fn(docx_path)
        print(f"  -> {pdf_path.name}, {docx_path.name}")

    print("Rendering previews...")
    render_previews()
    print("Done.")


if __name__ == "__main__":
    main()
