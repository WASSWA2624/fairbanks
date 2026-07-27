#!/usr/bin/env python3
"""Build five elegant, balanced, one-page CV designs for Racheal Nabukeera Sekagiri."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFilter
from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    Image as RLImage,
    Flowable,
)

ROOT = Path(__file__).resolve().parent
PHOTO_SRC = ROOT.parent / "WhatsApp Image 2026-07-27 at 14.23.53.jpeg"
PHOTO_DIR = ROOT / "_assets"
PHOTO_CIRCLE = PHOTO_DIR / "photo_circle.png"
PHOTO_SQUARE = PHOTO_DIR / "photo_square.png"
PREVIEW = ROOT / "_preview"

NAME = "RACHEAL NABUKEERA SEKAGIRI"
TITLE = "Senior Human Resource Executive  ·  Human Capital Strategist"
LOCATION = "Kampala, Uganda"
MOBILE = "+256 772 849258  ·  +256 701 849258"
EMAIL = "info@fairbanksmedicalcentre.org"
WEB = "www.fairbanksmedicalcentre.org"

PROFILE = (
    "Strategic HR executive with nearly <b>30 years</b> of experience and over "
    "<b>18 years</b> in executive leadership. Aligns people strategy with "
    "organisational performance. Founder & Managing Director, FairBanks Medical "
    "Centre Ltd. Formerly led Group HR for <b>2,000+ employees</b> across "
    "<b>50+ units</b>. PhD Candidate researching AI prediction of occupational burnout."
)

HIGHLIGHTS = [
    "30 years progressive HR & executive management",
    "18+ years executive Human Resource leadership",
    "HR leadership for 2,000+ staff across 50+ units",
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
        "line": "Strategy, governance, HR and growth; founded social enterprise (10,000+ beneficiaries).",
    },
    {
        "title": "Group Human Resource & Administration Manager",
        "org": "Norvik Group",
        "dates": "Jul 2016 - Feb 2026",
        "line": "Directed Group HR for 2,000+ employees; led workforce planning and leadership development.",
    },
    {
        "title": "Human Resource & Administration Manager",
        "org": "Norvik Hospital Ltd",
        "dates": "Sep 2013 - 2016",
        "line": "Led recruitment, employee relations, performance systems and HR policy.",
    },
    {
        "title": "Human Resource Manager",
        "org": "St. Catherine's Hospital",
        "dates": "2007 - 2010",
        "line": "Managed HR operations, welfare and labour compliance.",
    },
]
EARLIER = (
    "Earlier: Health Systems Administrator, St. Catherine's (2010-2013); "
    "HR Assistant, E Power Limited (1997-2000)."
)

EDU_LINE = (
    "<b>PhD Management</b> (Ongoing), Uganda Christian University  ·  "
    "<b>MSSPM</b>, Makerere University  ·  "
    "<b>BA Social Sciences</b>, Makerere University"
)
MORE_LINE = (
    "Member, Federation of Uganda Employers (FUE)  ·  "
    "English & Luganda (Fluent)  ·  References on request"
)

# Elegant palette
NAVY = HexColor("#0A3A52")
TEAL = HexColor("#1F6F78")
GOLD = HexColor("#B8953E")
INK = HexColor("#1F2A37")
MUTED = HexColor("#5B6B7C")
LINE = HexColor("#E4E9EF")
SOFT = HexColor("#F4F7FA")
CREAM = HexColor("#F8F5F0")
SIDEBAR = HexColor("#0A3A52")
CHARCOAL = HexColor("#18212B")
FOREST = HexColor("#1E4D52")
SLATE = HexColor("#2F4458")

DOCX_NAVY = RGBColor(0x0A, 0x3A, 0x52)
DOCX_TEAL = RGBColor(0x1F, 0x6F, 0x78)
DOCX_INK = RGBColor(0x1F, 0x2A, 0x37)
DOCX_MUTED = RGBColor(0x5B, 0x6B, 0x7C)
DOCX_GOLD = RGBColor(0xB8, 0x95, 0x3E)

VERSION_SLUGS = [
    "01_classic_airy",
    "02_sidebar_modern",
    "03_one_page",
    "04_timeline",
    "05_editorial_split",
]


class GoldRule(Flowable):
    """Short elegant gold accent under section titles."""

    def __init__(self, width=36 * mm, thickness=1.3):
        super().__init__()
        self._w = width
        self._t = thickness
        self.height = 6

    def wrap(self, availWidth, availHeight):
        return availWidth, self.height

    def draw(self):
        self.canv.setStrokeColor(GOLD)
        self.canv.setLineWidth(self._t)
        self.canv.setLineCap(1)
        self.canv.line(0, 3, self._w, 3)


def ensure_dirs():
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW.mkdir(parents=True, exist_ok=True)
    for slug in VERSION_SLUGS:
        (ROOT / slug).mkdir(parents=True, exist_ok=True)


def face_crop(img: Image.Image) -> Image.Image:
    """Full head-and-shoulders crop: complete face, full hair, upper blazer.

    Matches the original portrait framing (not an extreme close-up).
    Source is a tall waist-up shot (approx 2:3).
    """
    w, h = img.size
    # Use nearly full width so face and shoulders stay complete
    side = min(w, int(h * 0.68))
    # Start near the top so hair is never clipped; keep gentle headroom
    top = max(0, int(h * 0.02))
    if top + side > h:
        top = h - side
    left = max(0, (w - side) // 2)
    return img.crop((left, top, left + side, top + side))


def prepare_photos():
    """Build circle + rounded-square assets with full face and elegant rings."""
    img = Image.open(PHOTO_SRC).convert("RGBA")
    crop = face_crop(img)

    # --- Circle: photo inset so rings never cover hair/chin ---
    size = 720
    inset = int(size * 0.07)  # generous ring band; full face stays clear
    inner = size - 2 * inset
    photo = crop.resize((inner, inner), Image.Resampling.LANCZOS)
    mask = Image.new("L", (inner, inner), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, inner - 1, inner - 1), fill=255)
    photo.putalpha(mask)

    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(photo, (inset, inset), photo)

    ring = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    rd = ImageDraw.Draw(ring)
    # Outer soft navy
    rd.ellipse((4, 4, size - 5, size - 5), outline=(10, 58, 82, 255), width=6)
    # Inner gold accent (sits in the inset band)
    rd.ellipse(
        (inset - 2, inset - 2, size - inset + 1, size - inset + 1),
        outline=(184, 149, 62, 255),
        width=2,
    )
    Image.alpha_composite(out, ring).save(PHOTO_CIRCLE)

    # --- Rounded square: same generous crop ---
    size2 = 680
    inset2 = 10
    inner2 = size2 - 2 * inset2
    rad = 48
    s = crop.resize((inner2, inner2), Image.Resampling.LANCZOS)
    mask2 = Image.new("L", (inner2, inner2), 0)
    ImageDraw.Draw(mask2).rounded_rectangle(
        (0, 0, inner2 - 1, inner2 - 1), radius=rad, fill=255
    )
    s.putalpha(mask2)
    out2 = Image.new("RGBA", (size2, size2), (0, 0, 0, 0))
    out2.paste(s, (inset2, inset2), s)
    edge = Image.new("RGBA", (size2, size2), (0, 0, 0, 0))
    ed = ImageDraw.Draw(edge)
    ed.rounded_rectangle(
        (2, 2, size2 - 3, size2 - 3),
        radius=rad + 4,
        outline=(10, 58, 82, 255),
        width=6,
    )
    ed.rounded_rectangle(
        (inset2 - 1, inset2 - 1, size2 - inset2, size2 - inset2),
        radius=rad,
        outline=(184, 149, 62, 230),
        width=2,
    )
    Image.alpha_composite(out2, edge).save(PHOTO_SQUARE)

    # Also keep a clean rectangular portrait for reference (full face)
    portrait = PHOTO_DIR / "photo_portrait.png"
    # Keep original aspect, trim only tiny margins
    pw, ph = img.size
    margin_x = int(pw * 0.04)
    margin_top = int(ph * 0.015)
    margin_bot = int(ph * 0.28)  # drop lower torso; keep head+shoulders+arms
    rect = img.crop((margin_x, margin_top, pw - margin_x, ph - margin_bot))
    rect = rect.resize((520, int(520 * rect.size[1] / rect.size[0])), Image.Resampling.LANCZOS)
    rect.save(portrait)

    for old in PHOTO_DIR.glob("cand_*.png"):
        old.unlink(missing_ok=True)


def styles_base():
    return getSampleStyleSheet()


def add_style(styles, name, **kw):
    if name in styles.byName:
        # overwrite safely
        styles.byName[name] = ParagraphStyle(name, **kw)
        return styles.byName[name]
    styles.add(ParagraphStyle(name, **kw))
    return styles[name]


def bullet(text, style):
    return Paragraph(f"•  {text}", style)


def footer_fn(label: str):
    def _draw(canv, doc):
        canv.saveState()
        canv.setStrokeColor(LINE)
        canv.setLineWidth(0.5)
        y = 9 * mm
        canv.line(doc.leftMargin, y, A4[0] - doc.rightMargin, y)
        canv.setFillColor(GOLD)
        canv.circle(doc.leftMargin + 1.2 * mm, y, 1.1, fill=1, stroke=0)
        canv.setFillColor(MUTED)
        canv.setFont("Helvetica", 7.2)
        canv.drawString(doc.leftMargin + 5 * mm, 5.2 * mm, label)
        canv.drawRightString(A4[0] - doc.rightMargin, 5.2 * mm, "One page")
        canv.restoreState()

    return _draw


def skill_chips_table(skills, width, cols=5):
    styles = styles_base()
    add_style(
        styles,
        "Chip",
        fontName="Helvetica",
        fontSize=7.6,
        leading=9,
        textColor=NAVY,
        alignment=TA_CENTER,
    )
    cells = [Paragraph(s, styles["Chip"]) for s in skills]
    while len(cells) % cols:
        cells.append(Paragraph("", styles["Chip"]))
    rows = [cells[i : i + cols] for i in range(0, len(cells), cols)]
    t = Table(rows, colWidths=[width / cols] * cols)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), SOFT),
                ("BOX", (0, 0), (-1, -1), 0.4, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.3, LINE),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return t


# ---------------------------------------------------------------------------
# 01 Classic Airy
# ---------------------------------------------------------------------------
def build_v01_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "N", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=NAVY, spaceAfter=2)
    add_style(styles, "T", fontName="Helvetica", fontSize=9, leading=12, textColor=TEAL, spaceAfter=5)
    add_style(styles, "C", fontName="Helvetica", fontSize=8.2, leading=11, textColor=MUTED)
    add_style(styles, "H", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=NAVY, spaceBefore=8, spaceAfter=1)
    add_style(styles, "B", fontName="Helvetica", fontSize=8.8, leading=12.2, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=4)
    add_style(styles, "J", fontName="Helvetica-Bold", fontSize=9, leading=11.5, textColor=NAVY, spaceBefore=5, spaceAfter=0)
    add_style(styles, "O", fontName="Helvetica", fontSize=8.2, leading=10.5, textColor=TEAL, spaceAfter=1)
    add_style(styles, "L", fontName="Helvetica", fontSize=8.4, leading=11.2, textColor=INK, spaceAfter=1)
    add_style(styles, "M", fontName="Helvetica", fontSize=7.8, leading=10.5, textColor=MUTED, spaceAfter=2)
    add_style(styles, "Bu", fontName="Helvetica", fontSize=8.3, leading=11.2, textColor=INK, leftIndent=2, spaceAfter=1.5)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=11 * mm, bottomMargin=12 * mm,
        title=f"{NAME} - Classic Airy",
    )
    w = A4[0] - 30 * mm
    photo = RLImage(str(PHOTO_CIRCLE), width=32 * mm, height=32 * mm)
    head = Table(
        [[
            [Paragraph(NAME, styles["N"]), Paragraph(TITLE, styles["T"]),
             Paragraph(f"{LOCATION}<br/>{MOBILE}<br/>{EMAIL}  ·  {WEB}", styles["C"])],
            photo,
        ]],
        colWidths=[w - 38 * mm, 38 * mm],
    )
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BACKGROUND", (0, 0), (-1, -1), CREAM),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
        ("RIGHTPADDING", (1, 0), (1, 0), 8),
    ]))

    story = [
        head,
        Spacer(1, 4),
        HRFlowable(width="100%", thickness=1.8, color=NAVY, spaceAfter=0.5),
        HRFlowable(width="28%", thickness=1.1, color=GOLD, spaceAfter=6),
        Paragraph("PROFILE", styles["H"]),
        GoldRule(32 * mm),
        Paragraph(PROFILE, styles["B"]),
        Paragraph("HIGHLIGHTS", styles["H"]),
        GoldRule(32 * mm),
    ]
    left_h = [bullet(h, styles["Bu"]) for h in HIGHLIGHTS[:3]]
    right_h = [bullet(h, styles["Bu"]) for h in HIGHLIGHTS[3:]]
    story.append(Table([[left_h, right_h]], colWidths=[w / 2, w / 2], style=TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ])))
    story += [
        Spacer(1, 3),
        Paragraph("CORE COMPETENCIES", styles["H"]),
        GoldRule(32 * mm),
        skill_chips_table(SKILLS, w, cols=5),
        Spacer(1, 3),
        Paragraph("EXPERIENCE", styles["H"]),
        GoldRule(32 * mm),
    ]
    for role in ROLES:
        story.append(KeepTogether([
            Paragraph(role["title"], styles["J"]),
            Paragraph(f"{role['org']}   |   {role['dates']}", styles["O"]),
            Paragraph(role["line"], styles["L"]),
        ]))
    story += [
        Paragraph(EARLIER, styles["M"]),
        Spacer(1, 2),
        Paragraph("EDUCATION", styles["H"]),
        GoldRule(32 * mm),
        Paragraph(EDU_LINE, styles["L"]),
        Spacer(1, 4),
        Paragraph("ADDITIONAL", styles["H"]),
        GoldRule(32 * mm),
        Paragraph(MORE_LINE, styles["M"]),
    ]
    doc.build(story, onFirstPage=footer_fn("Classic Airy"), onLaterPages=footer_fn("Classic Airy"))


def build_v01_docx(out: Path):
    doc = Document()
    _page(doc, 1.4, 1.4, 1.5, 1.5)
    t = doc.add_table(rows=1, cols=2)
    L, R = t.rows[0].cells
    _shade(L, "F8F5F0")
    _shade(R, "F8F5F0")
    p = L.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 15, True, DOCX_NAVY)
    p = L.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9, False, DOCX_TEAL)
    for line in (LOCATION, MOBILE, f"{EMAIL}  ·  {WEB}"):
        p = L.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        _run(r, 8.5, False, DOCX_MUTED)
    R.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    R.paragraphs[0].add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.25))
    _h(doc, "PROFILE")
    _p(doc, PROFILE.replace("<b>", "").replace("</b>", ""))
    _h(doc, "HIGHLIGHTS")
    for h in HIGHLIGHTS:
        _b(doc, h)
    _h(doc, "CORE COMPETENCIES")
    _p(doc, "  ·  ".join(SKILLS))
    _h(doc, "EXPERIENCE")
    for role in ROLES:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        r = p.add_run(role["title"])
        _run(r, 9.5, True, DOCX_NAVY)
        p = doc.add_paragraph()
        r = p.add_run(f"{role['org']}  |  {role['dates']}")
        _run(r, 8.5, False, DOCX_TEAL)
        _p(doc, role["line"], size=9)
    _p(doc, EARLIER, muted=True, size=8)
    _h(doc, "EDUCATION")
    _p(doc, EDU_LINE.replace("<b>", "").replace("</b>", ""))
    _h(doc, "ADDITIONAL")
    _p(doc, MORE_LINE, muted=True, size=8.5)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# 02 Sidebar Modern
# ---------------------------------------------------------------------------
def build_v02_pdf(out: Path):
    width, height = A4
    side_w = 64 * mm
    margin = 10 * mm
    c = pdfcanvas.Canvas(str(out), pagesize=A4)

    # Sidebar
    c.setFillColor(SIDEBAR)
    c.rect(0, 0, side_w, height, fill=1, stroke=0)
    c.setFillColor(HexColor("#082F43"))
    c.rect(0, height - 68 * mm, side_w, 68 * mm, fill=1, stroke=0)

    photo_s = 40 * mm
    c.drawImage(
        str(PHOTO_CIRCLE),
        (side_w - photo_s) / 2,
        height - 56 * mm,
        width=photo_s,
        height=photo_s,
        mask="auto",
    )

    y = height - 72 * mm

    def side_h(label, y0):
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 7.8)
        c.drawString(8 * mm, y0, label)
        c.setStrokeColor(GOLD)
        c.setLineWidth(0.9)
        c.line(8 * mm, y0 - 2 * mm, side_w - 8 * mm, y0 - 2 * mm)
        return y0 - 6 * mm

    y = side_h("CONTACT", y)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.3)
    for line in [LOCATION, "+256 772 849258", "+256 701 849258", EMAIL, WEB]:
        # wrap long lines
        if c.stringWidth(line, "Helvetica", 7.3) > side_w - 16 * mm:
            # split at @ or .
            if "@" in line:
                a, b = line.split("@", 1)
                c.drawString(8 * mm, y, a + "@")
                y -= 3.4 * mm
                c.drawString(8 * mm, y, b)
            else:
                mid = len(line) // 2
                c.drawString(8 * mm, y, line[:mid])
                y -= 3.4 * mm
                c.drawString(8 * mm, y, line[mid:])
        else:
            c.drawString(8 * mm, y, line)
        y -= 3.8 * mm

    y -= 3 * mm
    y = side_h("STRENGTHS", y)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.4)
    for sk in SKILLS:
        c.drawString(8 * mm, y, f"•  {sk}")
        y -= 4.0 * mm

    y -= 3 * mm
    y = side_h("LANGUAGES", y)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.4)
    c.drawString(8 * mm, y, "English  —  Fluent")
    y -= 3.8 * mm
    c.drawString(8 * mm, y, "Luganda  —  Fluent")

    y -= 5 * mm
    y = side_h("MEMBERSHIP", y)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.4)
    c.drawString(8 * mm, y, "FUE Member")

    c.setFillColor(HexColor("#8FA3B5"))
    c.setFont("Helvetica", 6.8)
    c.drawString(8 * mm, 8 * mm, "Sidebar Modern  ·  One page")

    # Main
    x = side_w + margin
    max_w = width - side_w - 2 * margin
    y = height - 16 * mm

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 13)
    for line in _wrap(c, NAME, max_w, "Helvetica-Bold", 13):
        c.drawString(x, y, line)
        y -= 5.2 * mm
    c.setFillColor(TEAL)
    c.setFont("Helvetica", 8.3)
    for line in _wrap(c, TITLE.replace("  ·  ", " | "), max_w, "Helvetica", 8.3):
        c.drawString(x, y, line)
        y -= 4 * mm

    y -= 3 * mm
    y = _sec(c, x, y, "PROFILE")
    y = _text(c, x, y, max_w, PROFILE.replace("<b>", "").replace("</b>", ""), 8.5, 11.2)
    y -= 2 * mm
    y = _sec(c, x, y, "EXPERIENCE")
    for role in ROLES:
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 8.5)
        for line in _wrap(c, role["title"], max_w, "Helvetica-Bold", 8.5):
            c.drawString(x, y, line)
            y -= 3.5 * mm
        c.setFillColor(TEAL)
        c.setFont("Helvetica", 7.8)
        c.drawString(x, y, f"{role['org']}  |  {role['dates']}")
        y -= 3.6 * mm
        y = _text(c, x, y, max_w, role["line"], 8, 10.5)
        y -= 2.2 * mm
    y = _text(c, x, y, max_w, EARLIER, 7.4, 9.8, MUTED)
    y -= 3 * mm
    y = _sec(c, x, y, "EDUCATION")
    y = _text(c, x, y, max_w, EDU_LINE.replace("<b>", "").replace("</b>", ""), 8, 10.5)
    y -= 3 * mm
    y = _sec(c, x, y, "ADDITIONAL")
    y = _text(c, x, y, max_w, MORE_LINE, 7.6, 10, MUTED)
    c.save()


def build_v02_docx(out: Path):
    doc = Document()
    _page(doc, 1.0, 1.0, 1.0, 1.0)
    table = doc.add_table(rows=1, cols=2)
    side, main = table.rows[0].cells
    side.width = Cm(5.6)
    main.width = Cm(13.8)
    _shade(side, "0A3A52")
    sp = side.paragraphs[0]
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.25))

    def sh(text):
        p = side.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        r = p.add_run(text)
        _run(r, 8.5, True, DOCX_GOLD)

    def st(text):
        p = side.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(text)
        _run(r, 8, False, RGBColor(255, 255, 255))

    sh("CONTACT")
    for line in [LOCATION, MOBILE, EMAIL, WEB]:
        st(line)
    sh("STRENGTHS")
    for sk in SKILLS:
        st(f"• {sk}")
    sh("LANGUAGES")
    st("English — Fluent")
    st("Luganda — Fluent")
    sh("MEMBERSHIP")
    st("FUE Member")

    p = main.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 13, True, DOCX_NAVY)
    p = main.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9, False, DOCX_TEAL)
    _h(doc, "PROFILE", cell=main, size=9)
    _p_in(main, PROFILE.replace("<b>", "").replace("</b>", ""))
    _h(doc, "HIGHLIGHTS", cell=main, size=9)
    for h in HIGHLIGHTS:
        _b_in(main, h, size=8.5)
    _h(doc, "EXPERIENCE", cell=main, size=9)
    for role in ROLES:
        p = main.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        r = p.add_run(role["title"])
        _run(r, 9, True, DOCX_NAVY)
        p = main.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 8, False, DOCX_TEAL)
        _p_in(main, role["line"], size=8.5)
    _p_in(main, EARLIER, muted=True, size=8)
    _h(doc, "EDUCATION", cell=main, size=9)
    _p_in(main, EDU_LINE.replace("<b>", "").replace("</b>", ""), size=8.5)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# 03 One Page Compact Elegant
# ---------------------------------------------------------------------------
def build_v03_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "N", fontName="Helvetica-Bold", fontSize=14.5, leading=17, textColor=CHARCOAL, spaceAfter=1)
    add_style(styles, "T", fontName="Helvetica", fontSize=8.3, leading=11, textColor=FOREST, spaceAfter=2)
    add_style(styles, "C", fontName="Helvetica", fontSize=7.6, leading=10, textColor=MUTED)
    add_style(styles, "H", fontName="Helvetica-Bold", fontSize=8.2, leading=10, textColor=FOREST, spaceBefore=6, spaceAfter=1)
    add_style(styles, "B", fontName="Helvetica", fontSize=8, leading=10.8, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=3)
    add_style(styles, "J", fontName="Helvetica-Bold", fontSize=8.2, leading=10.5, textColor=CHARCOAL, spaceBefore=3, spaceAfter=0)
    add_style(styles, "O", fontName="Helvetica", fontSize=7.6, leading=9.8, textColor=FOREST, spaceAfter=0.5)
    add_style(styles, "L", fontName="Helvetica", fontSize=7.8, leading=10.2, textColor=INK, spaceAfter=1)
    add_style(styles, "M", fontName="Helvetica", fontSize=7.3, leading=9.5, textColor=MUTED, spaceAfter=1)
    add_style(styles, "Bu", fontName="Helvetica", fontSize=7.6, leading=10, textColor=INK, spaceAfter=1)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=13 * mm, rightMargin=13 * mm,
        topMargin=10 * mm, bottomMargin=10 * mm,
        title=f"{NAME} - One Page",
    )
    w = A4[0] - 26 * mm
    photo = RLImage(str(PHOTO_SQUARE), width=28 * mm, height=28 * mm)
    head = Table([[
        [Paragraph(NAME, styles["N"]), Paragraph(TITLE, styles["T"]),
         Paragraph(f"{LOCATION}   ·   {MOBILE}   ·   {EMAIL}", styles["C"])],
        photo,
    ]], colWidths=[w - 32 * mm, 32 * mm])
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story = [
        head,
        Spacer(1, 3),
        HRFlowable(width="100%", thickness=1.4, color=FOREST, spaceAfter=1),
        HRFlowable(width="22%", thickness=0.9, color=GOLD, spaceAfter=4),
        Paragraph("PROFILE", styles["H"]),
        GoldRule(28 * mm, 1.1),
        Paragraph(PROFILE, styles["B"]),
        Paragraph("HIGHLIGHTS & SKILLS", styles["H"]),
        GoldRule(28 * mm, 1.1),
    ]
    left = [bullet(h, styles["Bu"]) for h in HIGHLIGHTS]
    right = [Paragraph(f"•  {s}", styles["Bu"]) for s in SKILLS]
    story.append(Table([[left, right]], colWidths=[w * 0.52, w * 0.48], style=TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, -1), SOFT),
        ("BOX", (0, 0), (-1, -1), 0.35, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])))
    story += [Spacer(1, 3), Paragraph("EXPERIENCE", styles["H"]), GoldRule(28 * mm, 1.1)]
    for role in ROLES:
        story.append(Paragraph(role["title"], styles["J"]))
        story.append(Paragraph(f"{role['org']}  |  {role['dates']}", styles["O"]))
        story.append(Paragraph(role["line"], styles["L"]))
    story += [
        Paragraph(EARLIER, styles["M"]),
        Spacer(1, 2),
        Paragraph("EDUCATION & MORE", styles["H"]),
        GoldRule(28 * mm, 1.1),
        Paragraph(EDU_LINE, styles["L"]),
        Paragraph(MORE_LINE, styles["M"]),
    ]
    doc.build(story, onFirstPage=footer_fn("One Page"), onLaterPages=footer_fn("One Page"))


def build_v03_docx(out: Path):
    doc = Document()
    _page(doc, 1.1, 1.1, 1.3, 1.3)
    t = doc.add_table(rows=1, cols=2)
    L, R = t.rows[0].cells
    p = L.paragraphs[0]
    r = p.add_run(NAME)
    _run(r, 13, True, RGBColor(0x18, 0x21, 0x2B))
    p = L.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 8.5, False, RGBColor(0x1E, 0x4D, 0x52))
    p = L.add_paragraph()
    r = p.add_run(f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}")
    _run(r, 8, False, DOCX_MUTED)
    R.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    R.paragraphs[0].add_run().add_picture(str(PHOTO_SQUARE), width=Inches(1.1))
    _h(doc, "PROFILE", size=9)
    _p(doc, PROFILE.replace("<b>", "").replace("</b>", ""), size=9)
    _h(doc, "HIGHLIGHTS", size=9)
    for h in HIGHLIGHTS:
        _b(doc, h, size=8.5)
    _h(doc, "SKILLS", size=9)
    _p(doc, "  ·  ".join(SKILLS), size=8.5)
    _h(doc, "EXPERIENCE", size=9)
    for role in ROLES:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run(role["title"])
        _run(r, 9, True, DOCX_INK)
        p = doc.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 8, False, RGBColor(0x1E, 0x4D, 0x52))
        _p(doc, role["line"], size=8.5)
    _p(doc, EARLIER, muted=True, size=8)
    _h(doc, "EDUCATION & MORE", size=9)
    _p(doc, EDU_LINE.replace("<b>", "").replace("</b>", ""), size=8.5)
    _p(doc, MORE_LINE, muted=True, size=8)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# 04 Timeline
# ---------------------------------------------------------------------------
def build_v04_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "N", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=SLATE, alignment=TA_CENTER, spaceAfter=3)
    add_style(styles, "T", fontName="Helvetica", fontSize=9.5, leading=12, textColor=TEAL, alignment=TA_CENTER, spaceAfter=3)
    add_style(styles, "C", fontName="Helvetica", fontSize=8.2, leading=11, textColor=MUTED, alignment=TA_CENTER, spaceAfter=6)
    add_style(styles, "H", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=SLATE, alignment=TA_CENTER, spaceBefore=8, spaceAfter=3)
    add_style(styles, "B", fontName="Helvetica", fontSize=9, leading=12.5, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=6)
    add_style(styles, "D", fontName="Helvetica-Bold", fontSize=8, leading=10.5, textColor=TEAL, alignment=TA_RIGHT)
    add_style(styles, "J", fontName="Helvetica-Bold", fontSize=9, leading=11.5, textColor=SLATE)
    add_style(styles, "O", fontName="Helvetica", fontSize=8.2, leading=10.5, textColor=MUTED, spaceAfter=2)
    add_style(styles, "L", fontName="Helvetica", fontSize=8.4, leading=11.5, textColor=INK, spaceAfter=1)
    add_style(styles, "M", fontName="Helvetica", fontSize=8, leading=11, textColor=MUTED, alignment=TA_CENTER, spaceAfter=3)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=11 * mm, bottomMargin=12 * mm,
        title=f"{NAME} - Timeline",
    )
    w = A4[0] - 30 * mm
    photo = RLImage(str(PHOTO_CIRCLE), width=36 * mm, height=36 * mm)
    ph = Table([[photo]], colWidths=[w])
    ph.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))

    story = [
        Paragraph(NAME, styles["N"]),
        Paragraph(TITLE, styles["T"]),
        Paragraph(f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}", styles["C"]),
        ph,
        Spacer(1, 5),
        HRFlowable(width="26%", thickness=1.2, color=GOLD, spaceAfter=6),
        Paragraph("PROFILE", styles["H"]),
        GoldRule(30 * mm),
        Paragraph(PROFILE, styles["B"]),
        Paragraph("CAREER TIMELINE", styles["H"]),
        GoldRule(30 * mm),
        Spacer(1, 2),
    ]
    for role in ROLES:
        left = Paragraph(role["dates"].replace(" - ", "<br/>"), styles["D"])
        right = [
            Paragraph(role["title"], styles["J"]),
            Paragraph(role["org"], styles["O"]),
            Paragraph(role["line"], styles["L"]),
        ]
        row = Table([[left, right]], colWidths=[28 * mm, w - 28 * mm])
        row.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (1, 0), (1, 0), SOFT),
            ("LINEBEFORE", (1, 0), (1, 0), 2, TEAL),
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("RIGHTPADDING", (0, 0), (0, 0), 8),
            ("LEFTPADDING", (1, 0), (1, 0), 10),
            ("RIGHTPADDING", (1, 0), (1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        story.append(row)
        story.append(Spacer(1, 3.5))
    story += [
        Paragraph(EARLIER, styles["M"]),
        Spacer(1, 6),
        Paragraph("SKILLS", styles["H"]),
        GoldRule(30 * mm),
        skill_chips_table(SKILLS, w, cols=5),
        Spacer(1, 8),
        Paragraph("EDUCATION & MORE", styles["H"]),
        GoldRule(30 * mm),
    ]
    edu_box = Table(
        [[
            Paragraph(EDU_LINE, styles["M"]),
            Paragraph(MORE_LINE, styles["M"]),
        ]],
        colWidths=[w],
    )
    # Single cell cream panel for balance at page end
    edu_panel = Table(
        [[
            Paragraph(EDU_LINE + "<br/><br/>" + MORE_LINE, styles["M"]),
        ]],
        colWidths=[w],
    )
    edu_panel.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CREAM),
        ("BOX", (0, 0), (-1, -1), 0.4, LINE),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(edu_panel)
    doc.build(story, onFirstPage=footer_fn("Timeline"), onLaterPages=footer_fn("Timeline"))


def build_v04_docx(out: Path):
    doc = Document()
    _page(doc, 1.2, 1.2, 1.4, 1.4)
    for text, size, bold, color in [
        (NAME, 14, True, RGBColor(0x2F, 0x44, 0x58)),
        (TITLE, 9, False, DOCX_TEAL),
        (f"{LOCATION}  ·  {MOBILE}  ·  {EMAIL}", 8, False, DOCX_MUTED),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(text)
        _run(r, size, bold, color)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.2))
    _h(doc, "PROFILE", center=True, size=9)
    _p(doc, PROFILE.replace("<b>", "").replace("</b>", ""), size=9)
    _h(doc, "CAREER TIMELINE", center=True, size=9)
    for role in ROLES:
        t = doc.add_table(rows=1, cols=2)
        d, body = t.rows[0].cells
        d.width = Cm(3)
        p = d.paragraphs[0]
        r = p.add_run(role["dates"])
        _run(r, 8, True, DOCX_TEAL)
        p = body.paragraphs[0]
        r = p.add_run(role["title"])
        _run(r, 9, True, RGBColor(0x2F, 0x44, 0x58))
        p = body.add_paragraph()
        r = p.add_run(f"{role['org']} — {role['line']}")
        _run(r, 8.5, False, DOCX_INK)
        doc.add_paragraph()
    _p(doc, EARLIER, muted=True, size=8)
    _h(doc, "EDUCATION & MORE", center=True, size=9)
    _p(doc, EDU_LINE.replace("<b>", "").replace("</b>", ""), size=8.5)
    _p(doc, MORE_LINE, muted=True, size=8)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# 05 Editorial Split
# ---------------------------------------------------------------------------
def build_v05_pdf(out: Path):
    styles = styles_base()
    add_style(styles, "Tag", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=GOLD, spaceAfter=3)
    add_style(styles, "N", fontName="Helvetica-Bold", fontSize=16.5, leading=19, textColor=NAVY, spaceAfter=2)
    add_style(styles, "T", fontName="Helvetica", fontSize=8.5, leading=11, textColor=TEAL, spaceAfter=3)
    add_style(styles, "C", fontName="Helvetica", fontSize=7.6, leading=10, textColor=MUTED)
    add_style(styles, "H", fontName="Helvetica-Bold", fontSize=8.5, leading=10.5, textColor=NAVY, spaceBefore=1, spaceAfter=2)
    add_style(styles, "B", fontName="Helvetica", fontSize=8.2, leading=11, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=3)
    add_style(styles, "J", fontName="Helvetica-Bold", fontSize=8.3, leading=10.5, textColor=NAVY, spaceBefore=3, spaceAfter=0)
    add_style(styles, "O", fontName="Helvetica", fontSize=7.5, leading=9.5, textColor=TEAL, spaceAfter=0.5)
    add_style(styles, "L", fontName="Helvetica", fontSize=7.8, leading=10.2, textColor=INK, spaceAfter=1)
    add_style(styles, "S", fontName="Helvetica", fontSize=7.5, leading=10, textColor=INK, spaceAfter=1)
    add_style(styles, "M", fontName="Helvetica", fontSize=7.3, leading=9.5, textColor=MUTED, spaceAfter=2)
    add_style(styles, "Bu", fontName="Helvetica", fontSize=7.5, leading=10, textColor=INK, spaceAfter=1.2)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=10 * mm, bottomMargin=11 * mm,
        title=f"{NAME} - Editorial Split",
    )
    w = A4[0] - 24 * mm
    photo = RLImage(str(PHOTO_CIRCLE), width=34 * mm, height=34 * mm)
    banner = Table([[
        [
            Paragraph("CURRICULUM VITAE", styles["Tag"]),
            Paragraph(NAME, styles["N"]),
            Paragraph(TITLE, styles["T"]),
            Paragraph(f"{LOCATION}  ·  {MOBILE}<br/>{EMAIL}  ·  {WEB}", styles["C"]),
        ],
        photo,
    ]], colWidths=[w - 40 * mm, 40 * mm])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CREAM),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))

    left = [
        Paragraph("AT A GLANCE", styles["H"]),
        GoldRule(100),
    ]
    for h in HIGHLIGHTS:
        left.append(bullet(h, styles["Bu"]))
    left += [
        Spacer(1, 6),
        Paragraph("SKILLS", styles["H"]),
        GoldRule(100),
    ]
    for sk in SKILLS:
        left.append(Paragraph(f"•  {sk}", styles["S"]))
    left += [
        Spacer(1, 6),
        Paragraph("EDUCATION", styles["H"]),
        GoldRule(100),
        Paragraph("<b>PhD Management</b> (Ongoing)<br/>Uganda Christian University", styles["M"]),
        Paragraph("<b>MSSPM</b><br/>Makerere University", styles["M"]),
        Paragraph("<b>BA Social Sciences</b><br/>Makerere University", styles["M"]),
        Spacer(1, 5),
        Paragraph("MORE", styles["H"]),
        GoldRule(100),
        Paragraph("FUE Member<br/>English & Luganda — Fluent<br/>References on request", styles["M"]),
    ]

    right = [
        Paragraph("PROFILE", styles["H"]),
        GoldRule(100),
        Paragraph(PROFILE, styles["B"]),
        Spacer(1, 3),
        Paragraph("EXPERIENCE", styles["H"]),
        GoldRule(100),
    ]
    for role in ROLES:
        right += [
            Paragraph(role["title"], styles["J"]),
            Paragraph(f"{role['org']}  |  {role['dates']}", styles["O"]),
            Paragraph(role["line"], styles["L"]),
        ]
    right.append(Paragraph(EARLIER, styles["M"]))

    split = Table([[left, right]], colWidths=[w * 0.34, w * 0.66])
    split.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (0, 0), SOFT),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
        ("RIGHTPADDING", (0, 0), (0, 0), 8),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, 0), (0, 0), 8),
        ("LEFTPADDING", (1, 0), (1, 0), 10),
        ("RIGHTPADDING", (1, 0), (1, 0), 2),
        ("TOPPADDING", (1, 0), (1, 0), 2),
        ("LINEAFTER", (0, 0), (0, 0), 0.6, LINE),
    ]))

    story = [banner, Spacer(1, 7), split]
    doc.build(story, onFirstPage=footer_fn("Editorial Split"), onLaterPages=footer_fn("Editorial Split"))


def build_v05_docx(out: Path):
    doc = Document()
    _page(doc, 1.1, 1.1, 1.2, 1.2)
    head = doc.add_table(rows=1, cols=2)
    L, R = head.rows[0].cells
    _shade(L, "F8F5F0")
    _shade(R, "F8F5F0")
    p = L.paragraphs[0]
    r = p.add_run("CURRICULUM VITAE")
    _run(r, 8, True, DOCX_GOLD)
    p = L.add_paragraph()
    r = p.add_run(NAME)
    _run(r, 14, True, DOCX_NAVY)
    p = L.add_paragraph()
    r = p.add_run(TITLE)
    _run(r, 9, False, DOCX_TEAL)
    p = L.add_paragraph()
    r = p.add_run(f"{LOCATION} · {MOBILE} · {EMAIL}")
    _run(r, 8, False, DOCX_MUTED)
    R.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    R.paragraphs[0].add_run().add_picture(str(PHOTO_CIRCLE), width=Inches(1.25))
    doc.add_paragraph()
    body = doc.add_table(rows=1, cols=2)
    left, right = body.rows[0].cells
    left.width = Cm(6)
    right.width = Cm(12.5)
    _shade(left, "F4F7FA")
    left.paragraphs[0].clear()
    p = left.paragraphs[0]
    r = p.add_run("AT A GLANCE")
    _run(r, 9, True, DOCX_NAVY)
    for h in HIGHLIGHTS:
        _b_in(left, h, size=8)
    p = left.add_paragraph()
    r = p.add_run("SKILLS")
    _run(r, 9, True, DOCX_NAVY)
    for sk in SKILLS:
        _b_in(left, sk, size=8)
    p = left.add_paragraph()
    r = p.add_run("EDUCATION")
    _run(r, 9, True, DOCX_NAVY)
    _p_in(left, "PhD Management (Ongoing) - Uganda Christian University", size=8)
    _p_in(left, "MSSPM - Makerere University", size=8)
    _p_in(left, "BA Social Sciences - Makerere University", size=8)
    p = left.add_paragraph()
    r = p.add_run("MORE")
    _run(r, 9, True, DOCX_NAVY)
    _p_in(left, "FUE Member · English & Luganda · References on request", size=8)

    right.paragraphs[0].clear()
    p = right.paragraphs[0]
    r = p.add_run("PROFILE")
    _run(r, 9, True, DOCX_NAVY)
    _p_in(right, PROFILE.replace("<b>", "").replace("</b>", ""))
    p = right.add_paragraph()
    r = p.add_run("EXPERIENCE")
    _run(r, 9, True, DOCX_NAVY)
    for role in ROLES:
        p = right.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run(role["title"])
        _run(r, 9, True, DOCX_NAVY)
        p = right.add_paragraph()
        r = p.add_run(f"{role['org']} | {role['dates']}")
        _run(r, 8, False, DOCX_TEAL)
        _p_in(right, role["line"], size=8.5)
    _p_in(right, EARLIER, muted=True, size=8)
    doc.save(str(out))


# ---------------------------------------------------------------------------
# Canvas helpers
# ---------------------------------------------------------------------------
def _wrap(c, text, max_w, font, size):
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


def _sec(c, x, y, title):
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x, y, title)
    y -= 2.2 * mm
    c.setStrokeColor(GOLD)
    c.setLineWidth(1.2)
    c.line(x, y, x + 26 * mm, y)
    return y - 4 * mm


def _text(c, x, y, max_w, text, size, leading, color=INK):
    c.setFillColor(color)
    c.setFont("Helvetica", size)
    for line in _wrap(c, text, max_w, "Helvetica", size):
        c.drawString(x, y, line)
        y -= leading
    return y - 1.5


# ---------------------------------------------------------------------------
# DOCX helpers
# ---------------------------------------------------------------------------
def _page(doc, top, bottom, left, right):
    for sec in doc.sections:
        sec.top_margin = Cm(top)
        sec.bottom_margin = Cm(bottom)
        sec.left_margin = Cm(left)
        sec.right_margin = Cm(right)


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


def _h(doc, text, size=10, center=False, cell=None):
    target = cell if cell is not None else doc
    p = target.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    _run(r, size, True, DOCX_NAVY)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "10")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "B8953E")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _p(doc, text, muted=False, size=9.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    _run(r, size, False, DOCX_MUTED if muted else DOCX_INK)


def _b(doc, text, size=9):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    _run(r, size, False, DOCX_INK)


def _p_in(cell, text, muted=False, size=9):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    _run(r, size, False, DOCX_MUTED if muted else DOCX_INK)


def _b_in(cell, text, size=8.5):
    p = cell.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    _run(r, size, False, DOCX_INK)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def assert_one_page(pdf_path: Path):
    import fitz
    doc = fitz.open(str(pdf_path))
    n = doc.page_count
    doc.close()
    if n != 1:
        raise RuntimeError(f"{pdf_path.name} has {n} pages; expected 1")


def render_previews():
    import fitz
    for slug in VERSION_SLUGS:
        pdf = ROOT / slug / "racheal_cv.pdf"
        out_dir = PREVIEW / slug
        out_dir.mkdir(parents=True, exist_ok=True)
        for old in out_dir.glob("*.png"):
            old.unlink()
        doc = fitz.open(str(pdf))
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pix.save(str(out_dir / "page_1.png"))
        print(f"  {slug}: {doc.page_count} page(s)")
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
        assert_one_page(pdf_path)
        print(f"  OK -> pdf + docx (1 page)")
    print("Rendering previews...")
    render_previews()
    print("Done.")


if __name__ == "__main__":
    main()
