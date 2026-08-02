"""
Build Fravent Limited company profile PDF.
Brand colors: Deep Blue #071B36, Lemon Green #A3C324, White #FFFFFF
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image as PILImage
from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
PHOTOS = ROOT
OUT_PDF = ROOT / "profile-2026.pdf"
OUT_DOCX = ROOT / "profile-2026.docx"
OUT = OUT_PDF  # backward-compatible alias
TMP = ROOT.parent / "tmp" / "pdfs"
TMP.mkdir(parents=True, exist_ok=True)

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm

DEEP_BLUE = HexColor("#071B36")
DEEP_BLUE_MID = HexColor("#0C2E52")
LEMON = HexColor("#A3C324")
LEMON_DARK = HexColor("#8AAA1A")
LIGHT_GRAY = HexColor("#F4F6F8")
SOFT_BLUE = HexColor("#E8EEF4")
TEXT = HexColor("#1A2330")
MUTED = HexColor("#5A6570")

TAGLINE = "Your toolkit for real solutions"
SLOGAN = "Safe. Smart. Sustainable."


def p(name: str) -> Path:
    return PHOTOS / name


def a(name: str) -> Path:
    return ASSETS / name


def draw_logo_mark(c: canvas.Canvas, x: float, y: float, size: float = 12 * mm, light: bool = True):
    """Simple Fravent F + lemon arc mark."""
    f_color = white if light else DEEP_BLUE
    # green arc
    c.setStrokeColor(LEMON)
    c.setLineWidth(max(1.8, size * 0.08))
    c.arc(x, y, x + size, y + size, 40, 160)
    # stylized F
    c.setFillColor(f_color)
    stem_w = size * 0.18
    c.rect(x + size * 0.28, y + size * 0.18, stem_w, size * 0.62, fill=1, stroke=0)
    c.rect(x + size * 0.28, y + size * 0.68, size * 0.42, size * 0.12, fill=1, stroke=0)
    c.rect(x + size * 0.28, y + size * 0.42, size * 0.32, size * 0.10, fill=1, stroke=0)


def draw_header(c: canvas.Canvas, title: str, subtitle: str | None = None):
    c.setFillColor(DEEP_BLUE)
    c.rect(0, PAGE_H - 22 * mm, PAGE_W, 22 * mm, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.rect(0, PAGE_H - 23.5 * mm, PAGE_W, 1.5 * mm, fill=1, stroke=0)

    draw_logo_mark(c, MARGIN, PAGE_H - 19 * mm, size=11 * mm, light=True)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(MARGIN + 13 * mm, PAGE_H - 12 * mm, "FRAVENT LIMITED")
    c.setFont("Helvetica", 8)
    c.setFillColor(LEMON)
    c.drawString(MARGIN + 13 * mm, PAGE_H - 16.5 * mm, TAGLINE)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 14)
    c.drawRightString(PAGE_W - MARGIN, PAGE_H - 13 * mm, title)
    if subtitle:
        c.setFont("Helvetica", 8)
        c.setFillColor(LEMON)
        c.drawRightString(PAGE_W - MARGIN, PAGE_H - 17.5 * mm, subtitle)


def draw_footer(c: canvas.Canvas, page_no: int, total: int):
    y = 10 * mm
    c.setFillColor(DEEP_BLUE)
    c.rect(0, 0, PAGE_W, 12 * mm, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.rect(0, 12 * mm, PAGE_W, 1.2 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.5)
    c.drawString(MARGIN, y - 1 * mm, "fraventlimited@gmail.com  |  +256 778 455 042  |  Kampala, Uganda")
    c.drawRightString(PAGE_W - MARGIN, y - 1 * mm, f"{page_no} / {total}")


def fit_image(path: Path, max_w: float, max_h: float) -> tuple[float, float]:
    with PILImage.open(path) as im:
        w, h = im.size
    scale = min(max_w / w, max_h / h)
    return w * scale, h * scale


def draw_image(c: canvas.Canvas, path: Path, x: float, y: float, w: float, h: float, cover: bool = True):
    if not path.exists():
        c.setFillColor(SOFT_BLUE)
        c.rect(x, y, w, h, fill=1, stroke=0)
        return
    if cover:
        with PILImage.open(path) as im:
            im = im.convert("RGB")
            src_w, src_h = im.size
            target_ratio = w / h
            src_ratio = src_w / src_h
            if src_ratio > target_ratio:
                new_w = int(src_h * target_ratio)
                left = (src_w - new_w) // 2
                im = im.crop((left, 0, left + new_w, src_h))
            else:
                new_h = int(src_w / target_ratio)
                top = (src_h - new_h) // 2
                im = im.crop((0, top, src_w, top + new_h))
            tmp = TMP / f"crop_{path.stem}_{int(w)}_{int(h)}.jpg"
            im.save(tmp, quality=90)
            c.drawImage(str(tmp), x, y, width=w, height=h, mask="auto")
    else:
        iw, ih = fit_image(path, w, h)
        c.drawImage(str(path), x + (w - iw) / 2, y + (h - ih) / 2, width=iw, height=ih, mask="auto", preserveAspectRatio=True)


def draw_caption(c: canvas.Canvas, text: str, x: float, y: float, w: float):
    """Caption bar under a photo. y is the bottom of the photo."""
    bar_h = 9 * mm
    c.setFillColor(DEEP_BLUE)
    c.rect(x, y - bar_h, w, bar_h, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.rect(x, y - 1.2, w, 1.2, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica", 6.5)
    lines = wrap_text(c, text, "Helvetica", 6.5, w - 4 * mm)
    ty = y - 4.2 * mm if len(lines) == 1 else y - 3.5 * mm
    for line in lines[:2]:
        c.drawString(x + 2 * mm, ty, line)
        ty -= 7


def draw_captioned_image(
    c: canvas.Canvas,
    path: Path,
    x: float,
    y: float,
    w: float,
    h: float,
    caption: str,
    cover: bool = True,
):
    """Draw image with caption bar. (x,y) is bottom-left of full block including caption."""
    bar_h = 9 * mm
    img_h = h - bar_h
    draw_image(c, path, x, y + bar_h, w, img_h, cover=cover)
    draw_caption(c, caption, x, y + bar_h, w)


def circle_clip_image(c: canvas.Canvas, path: Path, cx: float, cy: float, r: float):
    if not path.exists():
        c.setFillColor(SOFT_BLUE)
        c.circle(cx, cy, r, fill=1, stroke=0)
        return
    # Prefer a pre-squared cover crop saved for stable circular clipping
    with PILImage.open(path) as im:
        im = im.convert("RGB")
        src_w, src_h = im.size
        side = min(src_w, src_h)
        left = (src_w - side) // 2
        top = (src_h - side) // 2
        im = im.crop((left, top, left + side, top + side))
        tmp = TMP / f"circle_{path.stem}_{int(r)}.jpg"
        im.save(tmp, quality=92)
    c.saveState()
    pth = c.beginPath()
    pth.circle(cx, cy, r)
    c.clipPath(pth, stroke=0)
    c.drawImage(str(tmp), cx - r, cy - r, width=2 * r, height=2 * r, mask="auto")
    c.restoreState()
    c.setStrokeColor(LEMON)
    c.setLineWidth(2.2)
    c.circle(cx, cy, r, fill=0, stroke=1)


def wrap_text(c: canvas.Canvas, text: str, font: str, size: float, max_w: float) -> list[str]:
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_paragraph(c: canvas.Canvas, text: str, x: float, y: float, max_w: float, size: float = 9.5, leading: float = 13, color=TEXT, font="Helvetica") -> float:
    c.setFont(font, size)
    c.setFillColor(color)
    for line in wrap_text(c, text, font, size, max_w):
        c.drawString(x, y, line)
        y -= leading
    return y


def section_label(c: canvas.Canvas, text: str, x: float, y: float):
    c.setFillColor(LEMON)
    c.roundRect(x, y - 2, 3 * mm, 8, 1, fill=1, stroke=0)
    c.setFillColor(DEEP_BLUE)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x + 5 * mm, y, text)


def meta_row(c: canvas.Canvas, label: str, value: str, x: float, y: float, label_w: float = 38 * mm, max_w: float = 120 * mm) -> float:
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(DEEP_BLUE)
    c.drawString(x, y, label)
    c.setFont("Helvetica", 9)
    c.setFillColor(TEXT)
    lines = wrap_text(c, value, "Helvetica", 9, max_w)
    for i, line in enumerate(lines):
        c.drawString(x + label_w, y - i * 12, line)
    return y - max(12, len(lines) * 12) - 4


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

def page_cover(c: canvas.Canvas):
    # Full-bleed hero
    hero = p("nakivale-solar-hybrid-ground-array-installation.jpeg")
    draw_image(c, hero, 0, 0, PAGE_W, PAGE_H, cover=True)

    # Dark overlay panels
    c.setFillColor(Color(0.03, 0.11, 0.21, alpha=0.72))
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # Left accent bar
    c.setFillColor(LEMON)
    c.rect(0, 0, 6 * mm, PAGE_H, fill=1, stroke=0)

    # Brand block
    draw_logo_mark(c, MARGIN + 4 * mm, PAGE_H - 42 * mm, size=18 * mm, light=True)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(MARGIN + 26 * mm, PAGE_H - 30 * mm, "FRAVENT LIMITED")
    c.setFillColor(LEMON)
    c.rect(MARGIN + 26 * mm, PAGE_H - 33 * mm, 42 * mm, 1.6, fill=1, stroke=0)
    c.setFont("Helvetica", 11)
    c.setFillColor(white)
    c.drawString(MARGIN + 26 * mm, PAGE_H - 40 * mm, "Electrical & Solar Solutions")

    c.setFont("Helvetica-Bold", 36)
    c.setFillColor(white)
    c.drawString(MARGIN + 4 * mm, PAGE_H / 2 + 20 * mm, "COMPANY")
    c.setFillColor(LEMON)
    c.drawString(MARGIN + 4 * mm, PAGE_H / 2 + 4 * mm, "PROFILE")

    c.setFillColor(white)
    c.setFont("Helvetica", 11)
    intro = "An introduction to who we are, what we do, and how we create lasting value for communities and clients across Uganda."
    y = PAGE_H / 2 - 10 * mm
    for line in wrap_text(c, intro, "Helvetica", 11, PAGE_W - 2 * MARGIN - 10 * mm):
        c.drawString(MARGIN + 4 * mm, y, line)
        y -= 15

    # Quote
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Oblique", 12)
    c.drawString(MARGIN + 4 * mm, 48 * mm, f'"{TAGLINE}"')

    c.setFillColor(white)
    c.setFont("Helvetica", 9)
    c.drawString(MARGIN + 4 * mm, 32 * mm, SLOGAN)
    c.drawString(MARGIN + 4 * mm, 24 * mm, "www.fraventlimited.com")
    c.drawString(MARGIN + 4 * mm, 16 * mm, "fraventlimited@gmail.com  ·  +256 778 455 042")


def page_about(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "About Us", "Our Story")
    y = PAGE_H - 36 * mm

    section_label(c, "Who We Are", MARGIN, y)
    y -= 10 * mm

    story = [
        "Fravent Limited is a fully registered company with the Registrar of Companies (Registration Number: 80020003614101). We exist to help clients realise their goals in electrical engineering, solar and renewable energy, property management, consultancy and advisory services, import and export, logistics and transport, land brokerage, and business intelligence.",
        "We employ experienced, exposed, and professional staff for every project we deliver. Our clients include private individuals, schools, private and public companies, organisations, local and international investors, and government ministries.",
        "A steady flow of orders from Fravent's customers is an encouraging sign and a true reflection of their trust and confidence in our management's rich heritage and diversified experience. This recognition rests on Fravent's persistent track record in delivering successful projects year after year.",
    ]
    for para in story:
        y = draw_paragraph(c, para, MARGIN, y, PAGE_W - 2 * MARGIN, size=10, leading=14)
        y -= 6

    y -= 4 * mm
    # Photo strip
    photos = [
        p("nakivale-solar-hybrid-rooftop-panel-installation.jpeg"),
        a("cover_field_team_main.png"),
        a("cover_pole_climber.png"),
    ]
    gap = 4 * mm
    card_w = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    card_h = 52 * mm
    for i, ph in enumerate(photos):
        x = MARGIN + i * (card_w + gap)
        draw_image(c, ph, x, y - card_h, card_w, card_h)
        c.setStrokeColor(LEMON)
        c.setLineWidth(1.5)
        c.rect(x, y - card_h, card_w, card_h, fill=0, stroke=1)

    y = y - card_h - 10 * mm
    # Stats / highlights
    boxes = [
        ("Registered", "Company No.\n80020003614101"),
        ("Focus", "Electrical &\nSolar Solutions"),
        ("Base", "Kampala,\nUganda"),
        ("Promise", "Quality,\nAffordable, On-time"),
    ]
    bw = (PAGE_W - 2 * MARGIN - 3 * gap) / 4
    bh = 28 * mm
    for i, (t, s) in enumerate(boxes):
        x = MARGIN + i * (bw + gap)
        c.setFillColor(DEEP_BLUE if i % 2 == 0 else LEMON)
        c.roundRect(x, y - bh, bw, bh, 3, fill=1, stroke=0)
        c.setFillColor(white if i % 2 == 0 else DEEP_BLUE)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 3 * mm, y - 8 * mm, t)
        c.setFont("Helvetica", 7.5)
        ty = y - 14 * mm
        for line in s.split("\n"):
            c.drawString(x + 3 * mm, ty, line)
            ty -= 9

    draw_footer(c, page_no, total)


def page_mission(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Purpose", "Mission · Vision · Values")

    # Mission card
    y = PAGE_H - 40 * mm
    c.setFillColor(DEEP_BLUE)
    c.roundRect(MARGIN, y - 55 * mm, PAGE_W - 2 * MARGIN, 55 * mm, 4, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.rect(MARGIN, y - 55 * mm, 4 * mm, 55 * mm, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(MARGIN + 10 * mm, y - 10 * mm, "OUR MISSION")
    c.setFillColor(white)
    mission = "Fravent Limited exists to establish a strong and reliable working bond with its clients, delivering quality, affordable and up-to-standard services for all our stakeholders."
    draw_paragraph(c, mission, MARGIN + 10 * mm, y - 20 * mm, PAGE_W - 2 * MARGIN - 16 * mm, size=11, leading=15, color=white)

    # Vision card
    y = y - 65 * mm
    c.setFillColor(SOFT_BLUE)
    c.roundRect(MARGIN, y - 45 * mm, PAGE_W - 2 * MARGIN, 45 * mm, 4, fill=1, stroke=0)
    c.setFillColor(DEEP_BLUE)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(MARGIN + 10 * mm, y - 10 * mm, "OUR VISION")
    vision = "Our company envisions evolving into the leading company providing unsurpassed services to its participants across the globe."
    draw_paragraph(c, vision, MARGIN + 10 * mm, y - 22 * mm, PAGE_W - 2 * MARGIN - 16 * mm, size=11, leading=15, color=TEXT)

    # Values
    y = y - 58 * mm
    section_label(c, "What Guides Us", MARGIN, y)
    y -= 12 * mm
    values = [
        ("Reliable", "We deliver systems that last and support that clients can trust."),
        ("Efficient", "We plan carefully, execute cleanly, and finish on time."),
        ("Sustainable", "We build solar and electrical solutions that power a brighter tomorrow."),
        ("Client-first", "We do not just install systems. We deliver lasting solutions."),
    ]
    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - gap) / 2
    ch = 28 * mm
    for i, (t, d) in enumerate(values):
        col = i % 2
        row = i // 2
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + gap)
        c.setFillColor(white)
        c.setStrokeColor(LEMON)
        c.setLineWidth(1.2)
        c.roundRect(x, yy - ch, cw, ch, 3, fill=1, stroke=1)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x + 4 * mm, yy - 8 * mm, t)
        draw_paragraph(c, d, x + 4 * mm, yy - 15 * mm, cw - 8 * mm, size=8.5, leading=11, color=MUTED)

    # CEO quote strip
    c.setFillColor(DEEP_BLUE)
    c.rect(0, 18 * mm, PAGE_W, 28 * mm, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Oblique", 13)
    c.drawCentredString(PAGE_W / 2, 32 * mm, f'"{TAGLINE}"')
    c.setFillColor(white)
    c.setFont("Helvetica", 8)
    c.drawCentredString(PAGE_W / 2, 23 * mm, "Message from the leadership · Fravent Limited")

    draw_footer(c, page_no, total)


def draw_people_grid(c: canvas.Canvas, people: list[tuple[str, str, Path]], start_y: float, cols: int = 3):
    gap = 6 * mm
    usable = PAGE_W - 2 * MARGIN
    card_w = (usable - (cols - 1) * gap) / cols
    card_h = 52 * mm
    for i, (name, role, photo) in enumerate(people):
        col = i % cols
        row = i // cols
        x = MARGIN + col * (card_w + gap)
        y = start_y - row * (card_h + 8 * mm)
        c.setFillColor(LIGHT_GRAY)
        c.roundRect(x, y - card_h, card_w, card_h, 4, fill=1, stroke=0)
        r = 14 * mm
        circle_clip_image(c, photo, x + card_w / 2, y - 18 * mm, r)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(x + card_w / 2, y - 38 * mm, name)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 7.5)
        for j, line in enumerate(wrap_text(c, role, "Helvetica", 7.5, card_w - 6 * mm)):
            c.drawCentredString(x + card_w / 2, y - 44 * mm - j * 9, line)
        c.setStrokeColor(LEMON)
        c.setLineWidth(1)
        c.line(x + card_w / 2 - 10 * mm, y - card_h + 5 * mm, x + card_w / 2 + 10 * mm, y - card_h + 5 * mm)


def page_management(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Leadership", "Management Team")
    y = PAGE_H - 38 * mm
    section_label(c, "Management Team", MARGIN, y)
    people = [
        ("Faray Joseph", "Managing Director", a("portrait_faray_joseph.png")),
        ("Olowo Bwenge Norman", "Secretary", a("portrait_olowo_bwenge_norman.png")),
        ("Ekapuloni Francis", "Nominee Director", a("portrait_ekapuloni_francis.png")),
        ("Mutebi Alvin", "Partner", a("portrait_mutebi_alvin.png")),
        ("Tumwesigye Henry", "Alternate Director", a("portrait_tumwesigye_henry.png")),
    ]
    draw_people_grid(c, people, y - 8 * mm, cols=3)

    # Bottom project visual
    y = 55 * mm
    draw_image(c, p("nakivale-solar-hybrid-ground-array-installation.jpeg"), MARGIN, 18 * mm, PAGE_W - 2 * MARGIN, y - 18 * mm)
    c.setFillColor(Color(0.03, 0.11, 0.21, alpha=0.45))
    c.rect(MARGIN, 18 * mm, PAGE_W - 2 * MARGIN, y - 18 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(MARGIN + 5 * mm, 28 * mm, "Hands-on leadership · Field-proven delivery")

    draw_footer(c, page_no, total)


def page_advisors(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Governance", "Board of Advisors")
    y = PAGE_H - 38 * mm
    section_label(c, "Board of Advisors", MARGIN, y)
    people = [
        ("Mr. Mukaaya Eddie", "Advisor", a("portrait_mukaaya_eddie.png")),
        ("Mr. Mawanda Hamzah", "Advisor", a("portrait_mawanda_hamzah.png")),
        ("Mr. Kibirige David", "Advisor", a("portrait_kibirige_david.png")),
        ("Mr. Joshua S. Mugabi", "Advisor", a("portrait_joshua_s_mugabi.png")),
    ]
    draw_people_grid(c, people, y - 8 * mm, cols=4)

    y = 95 * mm
    section_label(c, "In the Field", MARGIN, y)
    photos = [
        a("cover_field_team_main.png"),
        a("cover_field_team_bg.png"),
        p("nakivale-solar-hybrid-control-panel-installation.jpeg"),
    ]
    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = 48 * mm
    for i, ph in enumerate(photos):
        x = MARGIN + i * (cw + gap)
        draw_image(c, ph, x, 22 * mm, cw, ch)

    draw_footer(c, page_no, total)


def page_technical(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Our People", "Technical & Operations")
    y = PAGE_H - 36 * mm
    section_label(c, "Electrical Technical Leadership", MARGIN, y)
    y -= 8 * mm

    leaders = [
        ("Ekapuloni Francis", "Project & Technical Manager"),
        ("Tumwesigye Henry", "Project & Technical Advisor"),
        ("John Odutu", "Senior Technical Supervisor"),
        ("Osinde Mark", "Supervisor"),
        ("Richard Olec", "Supervisor"),
        ("James Kayeny", "Supervisor"),
        ("Melissa Johnson", "Supervisor"),
        ("Nicholas Wakatama", "Supervisor"),
        ("Egesa Deo", "Solar Sales Expert"),
    ]

    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = 18 * mm
    for i, (name, role) in enumerate(leaders):
        col = i % 3
        row = i // 3
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + 3 * mm)
        c.setFillColor(LIGHT_GRAY if row % 2 == 0 else SOFT_BLUE)
        c.roundRect(x, yy - ch, cw, ch, 2, fill=1, stroke=0)
        c.setFillColor(LEMON)
        c.circle(x + 5 * mm, yy - ch / 2, 2.2 * mm, fill=1, stroke=0)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 10 * mm, yy - 7 * mm, name)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 7)
        c.drawString(x + 10 * mm, yy - 13 * mm, role)

    y = y - 3 * (ch + 3 * mm) - 10 * mm
    section_label(c, "Electrical Technical & Operations Team", MARGIN, y)
    y -= 8 * mm
    ops = [
        "Sebufu Moses — Technician & Foreman",
        "Oburu David — Technician & Foreman",
        "Ongala Denish — Technician & Foreman",
        "Omella Thomas — Technician & Linesman",
        "Segawa Marvin — Technician",
        "Katushabe Bernard — Technician",
        "Mbazira Ronald — Technician",
        "Sendagire Derrick — Technician",
        "Mwesigwa Joseph — Linesman",
    ]
    for i, item in enumerate(ops):
        col = i % 2
        row = i // 2
        x = MARGIN + col * ((PAGE_W - 2 * MARGIN) / 2)
        yy = y - row * 11 * mm
        c.setFillColor(DEEP_BLUE)
        c.circle(x + 2 * mm, yy + 2, 1.5, fill=1, stroke=0)
        c.setFont("Helvetica", 8.5)
        c.setFillColor(TEXT)
        c.drawString(x + 6 * mm, yy, item)

    y = y - 5 * 11 * mm - 8 * mm
    note = "Our team is skilled and experienced, with nationals of Ugandan origin. We are dedicated to timely project delivery and passionate about seeing Fravent Limited excel among companies in Uganda."
    draw_paragraph(c, note, MARGIN, y, PAGE_W - 2 * MARGIN, size=9, leading=12, color=MUTED)

    y -= 18 * mm
    section_label(c, "Property Team", MARGIN, y)
    y -= 8 * mm
    prop = [
        ("Mutesi Jamirah", "Property Administrator"),
        ("Odoch Derrick", "Sales & Marketing Executive"),
        ("Odida Asiimwe Gerald", "Sales & Marketing Executive"),
        ("Mugaya Edrine", "Marketing Executive"),
    ]
    cw = (PAGE_W - 2 * MARGIN - 3 * 3 * mm) / 4
    for i, (name, role) in enumerate(prop):
        x = MARGIN + i * (cw + 3 * mm)
        c.setFillColor(DEEP_BLUE)
        c.roundRect(x, y - 22 * mm, cw, 22 * mm, 3, fill=1, stroke=0)
        c.setFillColor(LEMON)
        c.setFont("Helvetica-Bold", 7.5)
        for j, line in enumerate(wrap_text(c, name, "Helvetica-Bold", 7.5, cw - 4 * mm)):
            c.drawCentredString(x + cw / 2, y - 8 * mm - j * 9, line)
        c.setFillColor(white)
        c.setFont("Helvetica", 6.5)
        for j, line in enumerate(wrap_text(c, role, "Helvetica", 6.5, cw - 4 * mm)):
            c.drawCentredString(x + cw / 2, y - 16 * mm - j * 8, line)

    draw_footer(c, page_no, total)


def page_services_electrical(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Services", "Electrical & Solar")
    y = PAGE_H - 36 * mm
    section_label(c, "Electrical Engineering & Renewable Energy", MARGIN, y)
    y -= 8 * mm

    left = [
        "General electrical works (domestic & industrial wiring)",
        "Power line construction",
        "Transformer installation",
        "Solar backup systems",
        "Mini-grid installation",
        "Grid-tied solar systems",
        "Solar water irrigation installation",
        "Renewable energy solutions",
        "Earthing works",
    ]
    right = [
        "Facilities maintenance & engineering",
        "Power generators",
        "Air conditioners",
        "Solar PV systems",
        "Security cameras",
        "Low voltage systems",
        "LV and MV installations",
        "UPS systems & capacity batteries",
        "Switchgear, protection & control relays",
    ]

    col_w = (PAGE_W - 2 * MARGIN - 6 * mm) / 2
    for i, item in enumerate(left):
        yy = y - i * 11 * mm
        c.setFillColor(LEMON)
        c.circle(MARGIN + 2 * mm, yy + 2, 1.8, fill=1, stroke=0)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 8.5)
        c.drawString(MARGIN + 7 * mm, yy, item)
    for i, item in enumerate(right):
        yy = y - i * 11 * mm
        c.setFillColor(LEMON)
        c.circle(MARGIN + col_w + 8 * mm, yy + 2, 1.8, fill=1, stroke=0)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 8.5)
        c.drawString(MARGIN + col_w + 13 * mm, yy, item)

    y = y - 9 * 11 * mm - 8 * mm
    c.setFillColor(DEEP_BLUE)
    c.roundRect(MARGIN, y - 42 * mm, PAGE_W - 2 * MARGIN, 42 * mm, 4, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(MARGIN + 6 * mm, y - 10 * mm, "Turnkey renewable energy delivery")
    body = "From load assessment and system design to construction, installation, testing, commissioning and monitoring — we deliver solar power plants and solar mini-grids for both grid-tied and off-grid needs."
    draw_paragraph(c, body, MARGIN + 6 * mm, y - 20 * mm, PAGE_W - 2 * MARGIN - 12 * mm, size=9, leading=12, color=white)

    y = y - 50 * mm
    photos = [
        p("nakivale-solar-hybrid-electrical-panel-work.jpeg"),
        p("nakivale-solar-hybrid-rooftop-panel-installation.jpeg"),
        p("nakivale-solar-hybrid-completed-inverter-batteries-team.jpeg"),
    ]
    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = 42 * mm
    for i, ph in enumerate(photos):
        draw_image(c, ph, MARGIN + i * (cw + gap), 18 * mm, cw, ch)

    draw_footer(c, page_no, total)


def page_services_property(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Services", "Property Management")
    y = PAGE_H - 36 * mm
    section_label(c, "Property Management Services", MARGIN, y)
    y -= 10 * mm

    services = [
        ("Property Management", "Day-to-day management that keeps properties productive and well maintained."),
        ("Property Sales Agency", "Support for buying and selling with clear market guidance."),
        ("Tenant Search & Placement", "Finding, screening, and placing the right tenants."),
        ("Preparing Property for Rental", "Getting units ready for the market, including repair and maintenance."),
        ("Retrofitting & Consulting", "Practical advice and upgrades that improve performance and value."),
        ("Site Inspections & Reporting", "Regular inspections, reporting, and follow-up action."),
        ("Sub-contractor Supervision", "Hiring and supervising trusted sub-contractors on site."),
        ("Compliance & Remittances", "Submission of governmental remittances and related paperwork."),
        ("Tenant Administration", "Up-to-date tenant lists, notices, leasing, and communication."),
    ]

    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = 38 * mm
    for i, (t, d) in enumerate(services):
        col = i % 3
        row = i // 3
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + gap)
        c.setFillColor(white)
        c.setStrokeColor(SOFT_BLUE)
        c.setLineWidth(1)
        c.roundRect(x, yy - ch, cw, ch, 3, fill=1, stroke=1)
        c.setFillColor(LEMON)
        c.rect(x, yy - ch, 2.5 * mm, ch, fill=1, stroke=0)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 8)
        for j, line in enumerate(wrap_text(c, t, "Helvetica-Bold", 8, cw - 8 * mm)):
            c.drawString(x + 5 * mm, yy - 8 * mm - j * 10, line)
        draw_paragraph(c, d, x + 5 * mm, yy - 20 * mm, cw - 8 * mm, size=7.5, leading=10, color=MUTED)

    # Bottom banner
    c.setFillColor(DEEP_BLUE)
    c.rect(0, 18 * mm, PAGE_W, 24 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica", 9)
    c.drawCentredString(PAGE_W / 2, 28 * mm, "Electrical excellence · Property care · Lasting client partnerships")

    draw_footer(c, page_no, total)


def page_capabilities(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Capabilities", "How We Deliver")
    y = PAGE_H - 36 * mm

    blocks = [
        ("Power line construction", "We serve the power-line and energy industry with pride and integrity. Our teams deliver projects ranging from 11kV to 33kV, with continuous training across a wide scope of construction work."),
        ("Substation development", "For industrial, commercial, or residential power sub-stations, we assemble resources into a turnkey package — from ground-breaking to finished projects."),
        ("LV and MV installations", "Design, supply and construction of LV and MV installations, mainline and non-mainline compact secondary sub-stations, switchgear, protection and control relays, transformer installations, capacity batteries and filters, and UPS systems."),
        ("Renewable energy", "Load assessment, system design, construction, installation, testing, commissioning and monitoring of solar power plants and solar mini-grids — both grid-tied and off-grid."),
    ]
    for title, body in blocks:
        c.setFillColor(LIGHT_GRAY)
        c.roundRect(MARGIN, y - 36 * mm, PAGE_W - 2 * MARGIN, 36 * mm, 3, fill=1, stroke=0)
        c.setFillColor(LEMON)
        c.rect(MARGIN, y - 36 * mm, 3 * mm, 36 * mm, fill=1, stroke=0)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(MARGIN + 7 * mm, y - 9 * mm, title)
        draw_paragraph(c, body, MARGIN + 7 * mm, y - 18 * mm, PAGE_W - 2 * MARGIN - 12 * mm, size=8.5, leading=11, color=TEXT)
        y -= 40 * mm

    draw_footer(c, page_no, total)


def page_featured_overview(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Featured Project", "Solar Hybrid · Off Grid")

    draw_captioned_image(
        c,
        p("ugafode-nakivale-branch-site.jpeg"),
        MARGIN,
        PAGE_H - 98 * mm,
        PAGE_W - 2 * MARGIN,
        58 * mm,
        "UGAFODE Microfinance Limited (MDI) Nakivale Branch — project site.",
    )

    y = PAGE_H - 108 * mm
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(MARGIN, y, "Solar Hybrid System — OFF Grid")
    y -= 8 * mm
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8.5)
    c.drawString(MARGIN, y, "Completed and Commissioned · Nakivale Refugee Settlement")

    y -= 11 * mm
    y = meta_row(c, "Description:", "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System.", MARGIN, y, max_w=PAGE_W - 2 * MARGIN - 40 * mm)
    y = meta_row(c, "Client:", "UGAFODE Microfinance Limited - MDI", MARGIN, y)
    y = meta_row(c, "Location:", "Nakivale Refugee Settlement - Isingiro District, Western Uganda", MARGIN, y)
    y = meta_row(c, "System Size:", "5.58 kWh", MARGIN, y)
    y = meta_row(c, "Status:", "Completed and Commissioned", MARGIN, y)

    y -= 2 * mm
    photos = [
        (p("ugafode-nakivale-branch-directional-sign.jpeg"), "UGAFODE Nakivale Branch directional signage."),
        (p("nakivale-solar-hybrid-ground-array-installation.jpeg"), "Ground-mounted solar array installation at Nakivale."),
        (p("nakivale-solar-hybrid-rooftop-panel-installation.jpeg"), "Rooftop solar panel mounting by Fravent technicians."),
    ]
    gap = 3 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = 40 * mm
    for i, (ph, cap) in enumerate(photos):
        draw_captioned_image(c, ph, MARGIN + i * (cw + gap), 16 * mm, cw, ch, cap)

    draw_footer(c, page_no, total)


def page_featured_install(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Featured Project", "Installation in Progress")
    y = PAGE_H - 34 * mm
    section_label(c, "Field Installation — Nakivale (5.58 kWh)", MARGIN, y)
    y -= 6 * mm

    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-ground-array-installation.jpeg"),
        MARGIN,
        y - 92 * mm,
        (PAGE_W - 2 * MARGIN) * 0.58,
        92 * mm,
        "Technician securing ground-array mounts at Nakivale.",
    )
    rw = (PAGE_W - 2 * MARGIN) * 0.40
    rx = PAGE_W - MARGIN - rw
    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-technician-array-rear.jpeg"),
        rx,
        y - 45 * mm,
        rw,
        45 * mm,
        "Fravent branded PPE during array works.",
    )
    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-electrical-panel-work.jpeg"),
        rx,
        y - 92 * mm,
        rw,
        45 * mm,
        "Indoor electrical panel integration.",
    )

    y = y - 100 * mm
    c.setFillColor(DEEP_BLUE)
    c.roundRect(MARGIN, y - 28 * mm, PAGE_W - 2 * MARGIN, 28 * mm, 3, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(MARGIN + 5 * mm, y - 9 * mm, "Quality workmanship on site")
    body = "Solar array mounting, rooftop installation, electrical panel integration, and commissioning for reliable UGAFODE Nakivale branch operations."
    draw_paragraph(c, body, MARGIN + 5 * mm, y - 17 * mm, PAGE_W - 2 * MARGIN - 10 * mm, size=8, leading=10, color=white)

    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-control-panel-installation.jpeg"),
        MARGIN,
        16 * mm,
        PAGE_W - 2 * MARGIN,
        y - 36 * mm - 16 * mm,
        "Control panel wiring and commissioning — Fravent Limited technician.",
    )

    draw_footer(c, page_no, total)


def page_featured_completed(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Featured Project", "Completed System")
    y = PAGE_H - 34 * mm
    section_label(c, "Commissioned Off-Grid Hybrid System — 5.58 kWh", MARGIN, y)
    y -= 6 * mm

    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-completed-deye-system.jpeg"),
        MARGIN,
        y - 78 * mm,
        (PAGE_W - 2 * MARGIN) * 0.48,
        78 * mm,
        "Completed Deye hybrid inverter and battery installation.",
    )
    draw_captioned_image(
        c,
        p("nakivale-solar-hybrid-completed-inverter-batteries-team.jpeg"),
        MARGIN + (PAGE_W - 2 * MARGIN) * 0.50,
        y - 78 * mm,
        (PAGE_W - 2 * MARGIN) * 0.50,
        78 * mm,
        "Commissioned wall-mounted inverter and battery bank.",
    )

    y = y - 86 * mm
    cards = [
        ("Hybrid Inverter", "Deye hybrid inverter commissioned for off-grid branch power."),
        ("Battery Storage", "5.58 kWh lithium storage for overnight and backup supply."),
        ("Clean Integration", "Neat trunking, distribution boards, and isolators."),
        ("Client Outcome", "Stable power for UGAFODE MDI at Nakivale."),
    ]
    gap = 3 * mm
    cw = (PAGE_W - 2 * MARGIN - 3 * gap) / 4
    ch = 34 * mm
    for i, (t, d) in enumerate(cards):
        x = MARGIN + i * (cw + gap)
        c.setFillColor(DEEP_BLUE if i % 2 == 0 else LEMON)
        c.roundRect(x, y - ch, cw, ch, 3, fill=1, stroke=0)
        c.setFillColor(white if i % 2 == 0 else DEEP_BLUE)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(x + 2.5 * mm, y - 7 * mm, t)
        draw_paragraph(c, d, x + 2.5 * mm, y - 14 * mm, cw - 5 * mm, size=6.5, leading=8.5, color=white if i % 2 == 0 else DEEP_BLUE)

    y = y - ch - 6 * mm
    photos = [
        (p("ugafode-branch-completed-deye-customer-area.jpeg"), "Completed system viewed from the UGAFODE customer area."),
        (p("ugafode-solar-hybrid-changeover-control-panel.jpeg"), "Changeover panel — mains, generator, and inverter paths."),
    ]
    gap = 3 * mm
    cw = (PAGE_W - 2 * MARGIN - gap) / 2
    ch = y - 16 * mm
    for i, (ph, cap) in enumerate(photos):
        draw_captioned_image(c, ph, MARGIN + i * (cw + gap), 16 * mm, cw, ch, cap)

    draw_footer(c, page_no, total)


def page_ugafode_portfolio(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "More Projects", "UGAFODE Solar Hybrid Sites")
    y = PAGE_H - 34 * mm
    section_label(c, "Completed Solar Hybrid Systems — UGAFODE Microfinance Limited (MDI)", MARGIN, y)
    y -= 8 * mm

    projects = [
        {
            "title": "Rushere — Western Uganda",
            "size": "6.0 kWh",
            "status": "Completed and Commissioned",
            "note": "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System.",
        },
        {
            "title": "Ntungamo — Western Uganda",
            "size": "6.0 kWh",
            "status": "Completed and Commissioned",
            "note": "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System.",
        },
        {
            "title": "Kyaka Sales Branch",
            "size": "5.0 kWp DC",
            "status": "Completed",
            "note": "Procurement, Delivery, Installation, Testing and Commissioning of a Hybrid Solar System.",
        },
        {
            "title": "Rwamwanja Sales Branch",
            "size": "6.25 kWp DC",
            "status": "Completed",
            "note": "Procurement, Delivery, Installation, Testing and Commissioning of a Hybrid Solar System.",
        },
    ]

    gap = 4 * mm
    cw = (PAGE_W - 2 * MARGIN - gap) / 2
    ch = 42 * mm
    for i, proj in enumerate(projects):
        col = i % 2
        row = i // 2
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + gap)
        c.setFillColor(LIGHT_GRAY)
        c.roundRect(x, yy - ch, cw, ch, 3, fill=1, stroke=0)
        c.setFillColor(LEMON)
        c.rect(x, yy - ch, 3 * mm, ch, fill=1, stroke=0)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 6 * mm, yy - 9 * mm, proj["title"])
        c.setFont("Helvetica", 7.5)
        c.setFillColor(TEXT)
        c.drawString(x + 6 * mm, yy - 17 * mm, f"Client: UGAFODE Microfinance Limited - MDI")
        c.drawString(x + 6 * mm, yy - 24 * mm, f"System Size: {proj['size']}  ·  Status: {proj['status']}")
        draw_paragraph(c, proj["note"], x + 6 * mm, yy - 32 * mm, cw - 10 * mm, size=7, leading=9, color=MUTED)

    y = y - 2 * (ch + gap) - 4 * mm
    section_label(c, "Branch installation gallery", MARGIN, y)
    y -= 4 * mm
    photos = [
        (p("ugafode-solar-hybrid-transfer-switch-installation.jpeg"), "Transfer switch and Deye battery works at a UGAFODE branch."),
        (p("ugafode-branch-completed-deye-system-bw.jpeg"), "Completed Deye hybrid inverter and battery wall mount."),
        (p("ugafode-solar-hybrid-changeover-control-panel.jpeg"), "Manual changeover controls for mains, generator, and inverter."),
    ]
    gap = 3 * mm
    cw = (PAGE_W - 2 * MARGIN - 2 * gap) / 3
    ch = y - 16 * mm
    for i, (ph, cap) in enumerate(photos):
        draw_captioned_image(c, ph, MARGIN + i * (cw + gap), 16 * mm, cw, ch, cap)

    draw_footer(c, page_no, total)


def page_kiwoko_project(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "On-going Project", "Kiwoko Hospital")
    y = PAGE_H - 34 * mm
    section_label(c, "Solar Hybrid System — Kiwoko Hospital", MARGIN, y)
    y -= 8 * mm

    y = meta_row(c, "Description:", "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System.", MARGIN, y, max_w=PAGE_W - 2 * MARGIN - 40 * mm)
    y = meta_row(c, "Location:", "Kiwoko Hospital - Nakaseke District", MARGIN, y)
    y = meta_row(c, "System Size:", "10 kWh", MARGIN, y)
    y = meta_row(c, "Client:", "Adara Group", MARGIN, y)
    y = meta_row(c, "Status:", "On-going", MARGIN, y)

    y -= 4 * mm
    draw_captioned_image(
        c,
        p("kiwoko-hospital-adara-team-onsite.jpeg"),
        MARGIN,
        y - 58 * mm,
        PAGE_W - 2 * MARGIN,
        58 * mm,
        "Fravent team on site at Adara Development Uganda — Kiwoko, Nakaseke District.",
    )
    y = y - 64 * mm

    photos = [
        (p("kiwoko-jinko-hybrid-battery-inverter-room.jpeg"), "Jinko hybrid inverters and battery racks during installation."),
        (p("kiwoko-jinko-hybrid-commissioning-work.jpeg"), "Commissioning checks on the 10 kWh hybrid storage system."),
    ]
    gap = 3 * mm
    cw = (PAGE_W - 2 * MARGIN - gap) / 2
    ch = y - 16 * mm
    for i, (ph, cap) in enumerate(photos):
        draw_captioned_image(c, ph, MARGIN + i * (cw + gap), 16 * mm, cw, ch, cap)

    draw_footer(c, page_no, total)


def page_project_block(
    c: canvas.Canvas,
    page_no: int,
    total: int,
    title: str,
    meta: list[tuple[str, str]],
    photos: list[Path] | list[tuple[Path, str]],
    subtitle: str = "Previous Projects",
):
    draw_header(c, "Portfolio", subtitle)
    y = PAGE_H - 34 * mm
    c.setFillColor(DEEP_BLUE)
    c.setFont("Helvetica-Bold", 12)
    for line in wrap_text(c, title, "Helvetica-Bold", 12, PAGE_W - 2 * MARGIN):
        c.drawString(MARGIN, y, line)
        y -= 14
    y -= 4
    for label, value in meta:
        y = meta_row(c, label, value, MARGIN, y, max_w=PAGE_W - 2 * MARGIN - 42 * mm)

    y -= 4 * mm
    items: list[tuple[Path, str]] = []
    for item in photos:
        if isinstance(item, tuple):
            items.append(item)
        else:
            items.append((item, item.stem.replace("-", " ").replace("_", " ").title()))

    n = len(items)
    if n == 0:
        draw_footer(c, page_no, total)
        return
    cols = 2 if n <= 4 else 3
    rows = (n + cols - 1) // cols
    gap = 3 * mm
    avail_h = y - 16 * mm
    cw = (PAGE_W - 2 * MARGIN - (cols - 1) * gap) / cols
    ch = min(52 * mm, (avail_h - (rows - 1) * gap) / rows)
    for i, (ph, cap) in enumerate(items):
        col = i % cols
        row = i // cols
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + gap) - ch
        draw_captioned_image(c, ph, x, yy, cw, ch, cap)
    draw_footer(c, page_no, total)


def page_other_projects(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Portfolio", "Other Projects")
    y = PAGE_H - 36 * mm
    projects = [
        {
            "name": "Hybrid Solar System — Ernest Cook University",
            "location": "Balintuma Road, Mengo",
            "client": "Ernest Cook University",
            "status": "Completed",
            "desc": "Procurement, Delivery, Installation, Testing and Commissioning of a Hybrid Solar System.",
        },
        {
            "name": "Electrical Upgrade and Installation of Automatic Change Over System",
            "location": "Uganda Museum",
            "client": "Equatorial Power",
            "status": "Complete",
        },
        {
            "name": "Residential Apartment Wiring",
            "location": "Sentema Road, Mengo",
            "purpose": "Domestic Wiring",
            "status": "Complete",
        },
        {
            "name": "Maintenance of Lighting System",
            "location": "Luthuli Avenue, Bugolobi",
            "client": "Madam Ivy",
            "status": "Complete",
        },
    ]
    for proj in projects:
        c.setFillColor(LIGHT_GRAY)
        c.roundRect(MARGIN, y - 36 * mm, PAGE_W - 2 * MARGIN, 36 * mm, 3, fill=1, stroke=0)
        c.setFillColor(LEMON)
        c.rect(MARGIN, y - 36 * mm, 3 * mm, 36 * mm, fill=1, stroke=0)
        c.setFillColor(DEEP_BLUE)
        c.setFont("Helvetica-Bold", 9)
        ty = y - 9 * mm
        for line in wrap_text(c, proj["name"], "Helvetica-Bold", 9, PAGE_W - 2 * MARGIN - 12 * mm):
            c.drawString(MARGIN + 7 * mm, ty, line)
            ty -= 11
        c.setFont("Helvetica", 7.5)
        c.setFillColor(TEXT)
        details = []
        for key in ("location", "client", "purpose", "status"):
            if key in proj:
                details.append(f"{key.title()}: {proj[key]}")
        c.drawString(MARGIN + 7 * mm, y - 28 * mm, "  ·  ".join(details))
        if "desc" in proj:
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 7)
            c.drawString(MARGIN + 7 * mm, y - 33 * mm, proj["desc"][:110])
        y -= 40 * mm

    c.setFillColor(DEEP_BLUE)
    c.roundRect(MARGIN, 22 * mm, PAGE_W - 2 * MARGIN, 32 * mm, 3, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(MARGIN + 6 * mm, 42 * mm, "Ready for the next assignment")
    c.setFillColor(white)
    c.setFont("Helvetica", 8)
    c.drawString(MARGIN + 6 * mm, 32 * mm, "From UGAFODE branch hybrids and hospital systems to mini-grids and facility upgrades —")
    c.drawString(MARGIN + 6 * mm, 24 * mm, "Fravent Limited delivers practical electrical and solar solutions across Uganda.")

    draw_footer(c, page_no, total)


def page_partners(c: canvas.Canvas, page_no: int, total: int):
    draw_header(c, "Network", "Business Partners")
    y = PAGE_H - 36 * mm
    section_label(c, "Organisations We Work With", MARGIN, y)
    y -= 8 * mm

    partner_files = sorted(ASSETS.glob("partner_*.png"))
    # Prefer larger logos
    usable = [p for p in partner_files if PILImage.open(p).size[0] >= 100][:8]
    if not usable:
        usable = partner_files[:8]

    gap = 5 * mm
    cols = 4
    cw = (PAGE_W - 2 * MARGIN - (cols - 1) * gap) / cols
    ch = 28 * mm
    for i, ph in enumerate(usable):
        col = i % cols
        row = i // cols
        x = MARGIN + col * (cw + gap)
        yy = y - row * (ch + gap) - ch
        c.setFillColor(white)
        c.setStrokeColor(SOFT_BLUE)
        c.roundRect(x, yy, cw, ch, 3, fill=1, stroke=1)
        draw_image(c, ph, x + 3 * mm, yy + 3 * mm, cw - 6 * mm, ch - 6 * mm, cover=False)

    y = y - 2 * (ch + gap) - 10 * mm
    names = [
        "Rescue a Child Uganda (RACU)",
        "Hear His Voice Uganda",
        "Forbes Coffee Training Center",
        "Nirvana Wellness Spa",
        "Hope & Glory Foundation",
        "ESO J Technology Solutions Point",
        "RIA ROUND Creations",
        "MEGA Property Services",
    ]
    for i, name in enumerate(names):
        col = i % 2
        row = i // 2
        x = MARGIN + col * ((PAGE_W - 2 * MARGIN) / 2)
        yy = y - row * 10 * mm
        c.setFillColor(LEMON)
        c.circle(x + 2 * mm, yy + 2, 1.6, fill=1, stroke=0)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 8.5)
        c.drawString(x + 6 * mm, yy, name)

    y = y - 5 * 10 * mm - 6 * mm
    draw_image(c, a("cover_field_team_bg.png"), MARGIN, 18 * mm, PAGE_W - 2 * MARGIN, max(30 * mm, y - 18 * mm))

    draw_footer(c, page_no, total)


def page_contact(c: canvas.Canvas, page_no: int, total: int):
    # Full branded closing
    c.setFillColor(DEEP_BLUE)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(LEMON)
    c.rect(0, 0, 8 * mm, PAGE_H, fill=1, stroke=0)

    draw_logo_mark(c, MARGIN + 4 * mm, PAGE_H - 48 * mm, size=20 * mm, light=True)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 26)
    c.drawString(MARGIN + 28 * mm, PAGE_H - 32 * mm, "FRAVENT LIMITED")
    c.setFillColor(LEMON)
    c.setFont("Helvetica", 11)
    c.drawString(MARGIN + 28 * mm, PAGE_H - 40 * mm, TAGLINE)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(MARGIN + 4 * mm, PAGE_H - 70 * mm, "Let's power")
    c.setFillColor(LEMON)
    c.drawString(MARGIN + 4 * mm, PAGE_H - 82 * mm, "your next project.")

    y = PAGE_H - 105 * mm
    details = [
        ("Address", "P. O. Box 108053, Kampala (U)\nRofra House, 04th Floor,\nKansanga – Ggaba Road"),
        ("Phone", "+256 778 455 042\n+256 703 305 262"),
        ("Email", "fraventlimited@gmail.com"),
        ("Website", "www.fraventlimited.com"),
    ]
    for label, value in details:
        c.setFillColor(LEMON)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(MARGIN + 4 * mm, y, label.upper())
        c.setFillColor(white)
        c.setFont("Helvetica", 10)
        y -= 12
        for line in value.split("\n"):
            c.drawString(MARGIN + 4 * mm, y, line)
            y -= 12
        y -= 6

    # Photo strip
    photos = [
        p("nakivale-solar-hybrid-ground-array-installation.jpeg"),
        p("nakivale-solar-hybrid-completed-deye-system.jpeg"),
        p("fravent-brand-colors-flyer.jpeg"),
    ]
    gap = 3 * mm
    cw = (PAGE_W - MARGIN - 12 * mm - 2 * gap) / 3
    ch = 40 * mm
    for i, ph in enumerate(photos):
        draw_image(c, ph, MARGIN + 4 * mm + i * (cw + gap), 28 * mm, cw, ch)

    c.setFillColor(LEMON)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(PAGE_W / 2, 16 * mm, SLOGAN)
    c.setFillColor(white)
    c.setFont("Helvetica", 8)
    c.drawRightString(PAGE_W - MARGIN, 16 * mm, f"{page_no} / {total}")


def build_pdf():
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    pages = [
        page_cover,
        page_about,
        page_mission,
        page_management,
        page_advisors,
        page_technical,
        page_services_electrical,
        page_services_property,
        page_capabilities,
        page_featured_overview,
        page_featured_install,
        page_featured_completed,
        page_ugafode_portfolio,
        page_kiwoko_project,
        lambda cv, n, t: page_project_block(
            cv,
            n,
            t,
            "PROMOTION OF MINI-GRID FOR RURAL ELECTRIFICATION (Pro Mini Grids)",
            [
                ("Location:", "Lamwo District - Uganda (Paloga, Potika, Pangira, Labayango, Lapidyeni, Pawena, Apwoyo, Apyetta Town Council & Apyeta West, Kapeta, Muddu, Agoro & Lollimebeng)"),
                ("Duration:", "Three (3) Months"),
                ("Purpose:", "Last Mile Connection / Home Connection / House wiring"),
                ("Funders:", "GIZ, REA, Government of Uganda"),
                ("Contractor:", "Sagemcom Limited"),
                ("Client:", "WINCH Energy"),
            ],
            [
                (a(f"project_prominigrid_{i}.png"), f"Pro Mini Grids field works — Lamwo District ({i}/4).")
                for i in range(1, 5)
            ],
        ),
        lambda cv, n, t: page_project_block(
            cv,
            n,
            t,
            "SHARED COMMUNITY SOLAR PROJECT - Pro Mini Grid",
            [
                ("Location:", "Kalyamawolu Village, Kyabagamba Sub-county, Maddu Parish, Gomba District"),
                ("Client:", "Ministry of Local Government"),
                ("Status:", "Completed & Commissioned"),
            ],
            [
                (a(f"project_gomba_{i}.png"), f"Shared community solar project — Gomba District ({i}/4).")
                for i in range(1, 5)
            ],
        ),
        lambda cv, n, t: page_project_block(
            cv,
            n,
            t,
            "MAZIMA COMMUNITY SOLAR SYSTEM",
            [
                ("Location:", "Buwenge, Jinja District - Uganda"),
                ("Client:", "Mazima Community Development Center"),
                ("Status:", "Completed and Commissioned"),
            ],
            [
                (a(f"project_mazima_{i}.png"), f"Mazima Community Solar System — Jinja District ({i}/4).")
                for i in range(1, 5)
            ],
        ),
        page_other_projects,
        page_partners,
        page_contact,
    ]

    total = len(pages)
    for i, fn in enumerate(pages, start=1):
        if fn is page_cover:
            fn(c)
        else:
            fn(c, i, total)
        c.showPage()

    c.save()
    print(f"Wrote {OUT_PDF} ({total} pages)")
    return total


def build_docx():
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)

    def set_run_color(run, hex_color: str):
        run.font.color.rgb = RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

    def add_heading_bar(text: str, sub: str = ""):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = True
        run.font.size = Pt(16)
        set_run_color(run, "071B36")
        if sub:
            p2 = doc.add_paragraph()
            r2 = p2.add_run(sub)
            r2.font.size = Pt(10)
            set_run_color(r2, "A3C324")

    def add_caption(text: str):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.italic = True
        run.font.size = Pt(8)
        set_run_color(run, "5A6570")

    def add_picture(path: Path, width: float = 5.8, caption: str = ""):
        if path.exists():
            doc.add_picture(str(path), width=Inches(width))
            if caption:
                add_caption(caption)

    # Cover
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("FRAVENT LIMITED")
    r.bold = True
    r.font.size = Pt(28)
    set_run_color(r, "071B36")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("COMPANY PROFILE 2026")
    r.bold = True
    r.font.size = Pt(18)
    set_run_color(r, "A3C324")

    tag = doc.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tag.add_run(f'"{TAGLINE}"')
    r.italic = True
    r.font.size = Pt(12)
    set_run_color(r, "071B36")

    add_picture(
        p("nakivale-solar-hybrid-ground-array-installation.jpeg"),
        5.8,
        "Cover photo: Solar Hybrid System installation for UGAFODE at Nakivale Refugee Settlement.",
    )
    doc.add_paragraph("Electrical & Solar Solutions  ·  Safe. Smart. Sustainable.")
    doc.add_paragraph("www.fraventlimited.com  ·  fraventlimited@gmail.com  ·  +256 778 455 042")
    doc.add_page_break()

    # About
    add_heading_bar("About Us", "Our Story")
    for para in [
        "Fravent Limited is a fully registered company with the Registrar of Companies (Registration Number: 80020003614101). We help clients realise their goals in electrical engineering, solar and renewable energy, property management, consultancy and advisory services, import and export, logistics and transport, land brokerage, and business intelligence.",
        "We employ experienced professional staff for every project. Our clients include private individuals, schools, private and public companies, organisations, local and international investors, and government ministries.",
        "A steady flow of orders reflects trust in Fravent management's heritage and diversified experience, built on successful project delivery year after year.",
    ]:
        doc.add_paragraph(para)
    doc.add_page_break()

    # Mission
    add_heading_bar("Purpose", "Mission · Vision · Values")
    para = doc.add_paragraph()
    r = para.add_run("Mission: ")
    r.bold = True
    para.add_run("Establish a strong and reliable working bond with clients, delivering quality, affordable and up-to-standard services for all stakeholders.")
    para = doc.add_paragraph()
    r = para.add_run("Vision: ")
    r.bold = True
    para.add_run("Evolve into the leading company providing unsurpassed services to participants across the globe.")
    for v in ["Reliable", "Efficient", "Sustainable", "Client-first"]:
        doc.add_paragraph(v, style="List Bullet")
    doc.add_page_break()

    # Featured Nakivale
    add_heading_bar("Featured Project", "Solar Hybrid System — OFF Grid")
    meta = [
        ("Description", "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System."),
        ("Client", "UGAFODE Microfinance Limited - MDI"),
        ("Location", "Nakivale Refugee Settlement - Isingiro District, Western Uganda"),
        ("System Size", "5.58 kWh"),
        ("Status", "Completed and Commissioned"),
    ]
    for k, v in meta:
        para = doc.add_paragraph()
        r = para.add_run(f"{k}: ")
        r.bold = True
        para.add_run(v)

    nakivale_photos = [
        (p("ugafode-nakivale-branch-site.jpeg"), "UGAFODE Microfinance Limited (MDI) Nakivale Branch — project site."),
        (p("ugafode-nakivale-branch-directional-sign.jpeg"), "UGAFODE Nakivale Branch directional signage."),
        (p("nakivale-solar-hybrid-ground-array-installation.jpeg"), "Ground-mounted solar array installation at Nakivale."),
        (p("nakivale-solar-hybrid-rooftop-panel-installation.jpeg"), "Rooftop solar panel mounting by Fravent technicians."),
        (p("nakivale-solar-hybrid-technician-array-rear.jpeg"), "Fravent branded PPE during array works."),
        (p("nakivale-solar-hybrid-electrical-panel-work.jpeg"), "Indoor electrical panel integration."),
        (p("nakivale-solar-hybrid-control-panel-installation.jpeg"), "Control panel wiring and commissioning."),
        (p("nakivale-solar-hybrid-completed-deye-system.jpeg"), "Completed Deye hybrid inverter and battery installation."),
        (p("nakivale-solar-hybrid-completed-inverter-batteries-team.jpeg"), "Commissioned wall-mounted inverter and battery bank."),
        (p("ugafode-branch-completed-deye-customer-area.jpeg"), "Completed system viewed from the UGAFODE customer area."),
        (p("ugafode-solar-hybrid-changeover-control-panel.jpeg"), "Changeover panel — mains, generator, and inverter paths."),
    ]
    for path, cap in nakivale_photos:
        add_picture(path, 5.5, cap)
    doc.add_page_break()

    # UGAFODE portfolio
    add_heading_bar("More Projects", "UGAFODE Solar Hybrid Sites")
    ugafode_sites = [
        ("Rushere — Western Uganda", "6.0 kWh", "Completed and Commissioned"),
        ("Ntungamo — Western Uganda", "6.0 kWh", "Completed and Commissioned"),
        ("Nakivale Refugee Settlement — Isingiro District", "5.58 kWh", "Completed and Commissioned"),
        ("Kyaka Sales Branch", "5.0 kWp DC", "Completed"),
        ("Rwamwanja Sales Branch", "6.25 kWp DC", "Completed"),
    ]
    for title, size, status in ugafode_sites:
        para = doc.add_paragraph()
        r = para.add_run(title)
        r.bold = True
        doc.add_paragraph("Description: Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System.")
        doc.add_paragraph(f"Client: UGAFODE Microfinance Limited - MDI")
        doc.add_paragraph(f"System Size: {size}")
        doc.add_paragraph(f"Status: {status}")
        doc.add_paragraph("")

    for path, cap in [
        (p("ugafode-solar-hybrid-transfer-switch-installation.jpeg"), "Transfer switch and Deye battery works at a UGAFODE branch."),
        (p("ugafode-branch-completed-deye-system-bw.jpeg"), "Completed Deye hybrid inverter and battery wall mount."),
        (p("ugafode-solar-hybrid-changeover-control-panel.jpeg"), "Manual changeover controls for mains, generator, and inverter."),
    ]:
        add_picture(path, 5.2, cap)
    doc.add_page_break()

    # Kiwoko
    add_heading_bar("On-going Project", "Kiwoko Hospital — Adara Group")
    for k, v in [
        ("Description", "Procurement, Delivery, Installation, Testing and Commissioning of a Solar Hybrid System."),
        ("Location", "Kiwoko Hospital - Nakaseke District"),
        ("System Size", "10 kWh"),
        ("Client", "Adara Group"),
        ("Status", "On-going"),
    ]:
        para = doc.add_paragraph()
        r = para.add_run(f"{k}: ")
        r.bold = True
        para.add_run(v)
    for path, cap in [
        (p("kiwoko-hospital-adara-team-onsite.jpeg"), "Fravent team on site at Adara Development Uganda — Kiwoko, Nakaseke District."),
        (p("kiwoko-jinko-hybrid-battery-inverter-room.jpeg"), "Jinko hybrid inverters and battery racks during installation."),
        (p("kiwoko-jinko-hybrid-commissioning-work.jpeg"), "Commissioning checks on the 10 kWh hybrid storage system."),
    ]:
        add_picture(path, 5.5, cap)
    doc.add_page_break()

    # Earlier portfolio
    add_heading_bar("Earlier Portfolio", "Community Solar & Mini-Grids")
    doc.add_paragraph("Pro Mini Grids — Lamwo District (WINCH Energy / Sagemcom / GIZ, REA, Government of Uganda).")
    for i in range(1, 5):
        add_picture(a(f"project_prominigrid_{i}.png"), 4.8, f"Pro Mini Grids field works — Lamwo District ({i}/4).")
    doc.add_paragraph("Shared Community Solar — Gomba District (Ministry of Local Government).")
    for i in range(1, 5):
        add_picture(a(f"project_gomba_{i}.png"), 4.8, f"Shared community solar project — Gomba District ({i}/4).")
    doc.add_paragraph("Mazima Community Solar System — Buwenge, Jinja District.")
    for i in range(1, 5):
        add_picture(a(f"project_mazima_{i}.png"), 4.8, f"Mazima Community Solar System — Jinja District ({i}/4).")
    doc.add_page_break()

    # Other
    add_heading_bar("Other Projects")
    for name, detail in [
        ("Ernest Cook University — Balintuma Road, Mengo", "Hybrid Solar System. Status: Completed."),
        ("Uganda Museum", "Electrical Upgrade and Automatic Change Over — Equatorial Power. Status: Complete."),
        ("Sentema Road, Mengo", "Residential Apartment Wiring. Status: Complete."),
        ("Luthuli Avenue, Bugolobi", "Maintenance of Lighting System — Madam Ivy. Status: Complete."),
    ]:
        para = doc.add_paragraph()
        r = para.add_run(name)
        r.bold = True
        doc.add_paragraph(detail)

    doc.add_page_break()
    add_heading_bar("Contact")
    doc.add_paragraph("Address: P. O. Box 108053, Kampala (U), Rofra House, 04th Floor, Kansanga - Ggaba Road")
    doc.add_paragraph("Phone: +256 778 455 042  ·  +256 703 305 262")
    doc.add_paragraph("Email: fraventlimited@gmail.com")
    doc.add_paragraph("Website: www.fraventlimited.com")
    doc.add_paragraph(SLOGAN)

    doc.save(str(OUT_DOCX))
    print(f"Wrote {OUT_DOCX}")


def build():
    build_pdf()
    build_docx()


if __name__ == "__main__":
    build()
