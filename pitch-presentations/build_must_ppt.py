#!/usr/bin/env python3
"""
MUST partnership pitch deck — FairBanks Medical Centre & FCHIP.
10 slides max; large readable fonts; no Community Reach diagram image.
"""

from pathlib import Path

from lxml import etree
from PIL import Image as PILImage
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

REPO = Path(__file__).resolve().parents[1]
ASSETS = REPO / "assets"
LOGO = ASSETS / "fairbanks_logo.jpeg"
OUT = Path(__file__).resolve().parent / "documents"
OUT.mkdir(parents=True, exist_ok=True)
OUT_PPT = OUT / "must_ppt.pptx"

NAVY, TEAL, GREEN = "0A1F2E", "0D6E6E", "2D7A55"
ORANGE, GOLD, CREAM = "C45C26", "D99A2B", "F7F5F0"
WHITE = "FFFFFF"
SLATE, MUTED, LINE, LIGHT = "1E2F38", "52636C", "CED9D8", "D4E8DC"

ORG = "FairBanks Medical Centre"
SLOGAN = "Your health, our mission."
FOOTER = f"{ORG}  ·  {SLOGAN}"
FCHIP = "FairBanks Community Health Intelligence Platform"
MUST = "Mbarara University of Science and Technology"

PHOTOS = {
    "cover": ASSETS / "facility_exterior_branded_entrance_01.jpeg",
    "facility": ASSETS / "facility_exterior_entrance_01.jpg",
    "facility_wide": ASSETS / "facility_exterior_branded_entrance_02.jpeg",
    "outreach": ASSETS / "outreach_bp_screening.jpeg",
    "dashboard": ASSETS / "dashboard_demo.png",
    "camp": ASSETS / "outreach_medical_camp_01.jpg",
}

ML, MR = 0.55, 0.55
SW, SH = 13.333, 7.5
CW = SW - ML - MR
TOP_BAR = 0.06
LOGO_H = 0.40
LOGO_W = LOGO_H * (269 / 101)
LOGO_Y = 0.24
FOOTER_Y = 7.02
CONTENT_BOTTOM = 6.62
HEADER_Y = 1.05
HEADER_Y_SUB = 1.42


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def set_run(run, text, size, color, bold=False, font="Calibri"):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)


def add_fade(slide):
    sld = slide._element
    for child in list(sld):
        if child.tag == qn("p:transition"):
            sld.remove(child)
    transition = etree.Element(qn("p:transition"), {"spd": "med"})
    etree.SubElement(transition, qn("p:fade"))
    cSld = sld.find(qn("p:cSld"))
    if cSld is not None:
        cSld.addnext(transition)
    else:
        sld.insert(0, transition)


def build():
    missing = [p for p in [LOGO, *PHOTOS.values()] if not p.exists()]
    if missing:
        raise FileNotFoundError("Missing assets:\n" + "\n".join(map(str, missing)))

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    blank = prs.slide_layouts[6]

    def rect(slide, x, y, w, h, fill, line=None, rounded=False):
        kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
        if line:
            shape.line.color.rgb = rgb(line)
        else:
            shape.line.fill.background()
        if rounded:
            try:
                shape.adjustments[0] = 0.09
            except Exception:
                pass
        try:
            spPr = shape._element.spPr
            for child in list(spPr):
                if "effectLst" in child.tag or "effectDag" in child.tag:
                    spPr.remove(child)
            etree.SubElement(spPr, qn("a:effectLst"))
        except Exception:
            pass
        return shape

    def textbox(
        slide,
        value,
        x,
        y,
        w,
        h,
        size=22,
        color=SLATE,
        bold=False,
        align=PP_ALIGN.LEFT,
        valign=MSO_ANCHOR.TOP,
        line_spacing=1.5,
        para_gap=12,
    ):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(22000)
        tf.margin_top = tf.margin_bottom = Emu(8000)
        tf.vertical_anchor = valign
        lines = value.split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.space_before = Pt(0)
            if len(lines) > 1:
                p.space_after = Pt(22 if line.strip() == "" else para_gap)
            else:
                p.space_after = Pt(0)
            p.line_spacing = line_spacing
            r = p.add_run()
            set_run(r, line, size, color, bold)
        return box

    def bullets(slide, items, x, y, w, h, size=22, color=SLATE, space=18):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(22000)
        tf.margin_top = tf.margin_bottom = Emu(6000)
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(4)
            p.space_after = Pt(space)
            p.line_spacing = 1.5
            r = p.add_run()
            set_run(r, "•  " + item, size, color)
        return box

    def add_logo(slide, x, y, height=LOGO_H):
        pad = 0.08
        py = max(y - pad, TOP_BAR + 0.05)
        rect(slide, x - pad, py, LOGO_W + 2 * pad, height + (y - py) + pad, WHITE, LINE, True)
        return slide.shapes.add_picture(str(LOGO), Inches(x), Inches(y), height=Inches(height))

    def crop_photo(slide, path, x, y, w, h):
        with PILImage.open(path) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
        fr, ir = w / h, iw / ih
        if ir > fr:
            a = (1 - fr / ir) / 2
            pic.crop_left = pic.crop_right = a
        else:
            a = (1 - ir / fr) / 2
            pic.crop_top = pic.crop_bottom = a
        return pic

    def new_slide(bg=CREAM):
        s = prs.slides.add_slide(blank)
        rect(s, 0, 0, SW, SH, bg)
        add_fade(s)
        return s

    def header(s, kicker, title, subtitle=""):
        rect(s, 0, 0, SW, TOP_BAR, TEAL)
        rect(s, 0, TOP_BAR, 0.08, SH - TOP_BAR, ORANGE)
        logo_x = SW - MR - LOGO_W
        add_logo(s, logo_x, LOGO_Y)
        tw = logo_x - ML - 0.40
        if kicker.strip():
            # One line: "Kicker: Title" — orange kicker, navy title
            box = s.shapes.add_textbox(Inches(ML), Inches(0.28), Inches(tw), Inches(0.62))
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = Emu(18000)
            tf.margin_top = tf.margin_bottom = Emu(4000)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            p.line_spacing = 1.15
            r1 = p.add_run()
            set_run(r1, kicker.strip() + ": ", 28, ORANGE, True)
            r2 = p.add_run()
            set_run(r2, title, 28, NAVY, True)
            y = 1.00
        else:
            textbox(s, title, ML, 0.24, tw, 0.62, 34, NAVY, True)
            y = 0.98
        if subtitle:
            textbox(s, subtitle, ML, y, CW, 0.42, 20, MUTED)
            return HEADER_Y_SUB
        return HEADER_Y

    def footer(s, number, total):
        rect(s, ML, FOOTER_Y, CW, 0.012, LINE)
        textbox(s, FOOTER, ML, FOOTER_Y + 0.08, 9.8, 0.28, 14, MUTED)
        textbox(
            s,
            f"{number}  /  {total}",
            SW - MR - 1.3,
            FOOTER_Y + 0.08,
            1.3,
            0.28,
            14,
            MUTED,
            align=PP_ALIGN.RIGHT,
        )

    slides = []

    # ------------------------------------------------------------------
    # 1 Cover
    # ------------------------------------------------------------------
    s = new_slide(NAVY)
    panel = 7.85
    crop_photo(s, PHOTOS["cover"], 0, 0, SW, SH)
    rect(s, 0, 0, panel, SH, NAVY)
    rect(s, panel, 0, 0.12, SH, GOLD)
    add_logo(s, 0.55, 0.28, 0.50)

    textbox(s, "UNIVERSITY PARTNERSHIP PROPOSAL", 0.55, 1.00, 7.0, 0.36, 18, GOLD, True)
    textbox(s, ORG, 0.55, 1.42, 7.0, 0.42, 24, TEAL, True)
    textbox(s, "Working with MUST", 0.55, 2.05, 7.0, 0.58, 38, WHITE, True)
    textbox(s, "for community health,", 0.55, 2.65, 7.0, 0.58, 38, WHITE, True)
    textbox(s, "research, and grants", 0.55, 3.25, 7.0, 0.58, 38, GOLD, True)
    rect(s, 0.55, 3.95, 2.80, 0.09, GOLD)

    textbox(s, "Introducing FCHIP", 0.55, 4.25, 7.0, 0.42, 24, LIGHT)
    textbox(s, FCHIP, 0.55, 4.72, 7.0, 0.55, 19, GOLD, True)
    textbox(s, SLOGAN, 0.55, 5.35, 7.0, 0.40, 24, GOLD, True)

    rect(s, 0, 6.25, panel, 1.25, TEAL)
    textbox(s, f"Prepared for {MUST}", 0.55, 6.42, 7.0, 0.38, 19, WHITE, True)
    textbox(s, "Kyebando–Kisalosalo, Kampala  ·  fairbanksmedicalcentre.org", 0.55, 6.88, 7.0, 0.35, 17, LIGHT)
    slides.append(s)

    # ------------------------------------------------------------------
    # 2 Why we are here
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Purpose", "Why we are here")
    textbox(
        s,
        "FairBanks Medical Centre seeks a practical partnership with MUST —\n"
        "linking a working clinic and community programmes with university teaching,\n"
        "research, innovation, and grant writing.",
        ML,
        y0,
        CW,
        1.20,
        22,
        SLATE,
        para_gap=10,
    )
    points = [
        ("Share a living field site", "A licensed medical centre and active community work in Kampala peri-urban areas.", TEAL),
        ("Grow FCHIP together", "Build and test the Community Health Intelligence Platform with MUST scholars and students.", ORANGE),
        ("Win stronger grants", "Joint proposals that combine academic depth with real community and facility data.", GREEN),
    ]
    start = y0 + 1.35
    gap = 0.14
    row_h = (CONTENT_BOTTOM - start - 2 * gap) / 3
    for i, (t, body, accent) in enumerate(points):
        y = start + i * (row_h + gap)
        rect(s, ML, y, CW, row_h, WHITE, LINE, True)
        rect(s, ML + 0.18, y + 0.14, 0.10, row_h - 0.28, accent)
        textbox(s, t, ML + 0.45, y + 0.12, CW - 0.75, 0.42, 24, NAVY, True)
        textbox(s, body, ML + 0.45, y + 0.58, CW - 0.75, row_h - 0.72, 21, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 3 About FairBanks
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Who we are", "About FairBanks Medical Centre")
    left_w, gap = 7.15, 0.26
    right_w = CW - left_w - gap
    box_h = CONTENT_BOTTOM - y0
    rect(s, ML, y0, left_w, box_h, WHITE, LINE, True)
    textbox(s, "A community medical centre in Kampala", ML + 0.35, y0 + 0.28, left_w - 0.70, 0.42, 23, NAVY, True)
    textbox(s, SLOGAN, ML + 0.35, y0 + 0.75, left_w - 0.70, 0.34, 20, ORANGE, True)
    bullets(
        s,
        [
            "Located in Kyebando–Kisalosalo, Kampala",
            "Outpatient & inpatient care, emergency, diagnostics, pharmacy",
            "Maternal and child health, chronic care, prevention",
            "Community Reach with CHWs / VHTs nearby",
            "Building FCHIP as our health intelligence layer",
        ],
        ML + 0.35,
        y0 + 1.25,
        left_w - 0.70,
        box_h - 1.50,
        20,
        SLATE,
        18,
    )
    rx = ML + left_w + gap
    photo_h = box_h * 0.58
    crop_photo(s, PHOTOS["facility"], rx, y0, right_w, photo_h)
    cap_y = y0 + photo_h + 0.14
    cap_h = CONTENT_BOTTOM - cap_y
    rect(s, rx, cap_y, right_w, cap_h, TEAL, rounded=True)
    textbox(s, ORG, rx + 0.22, cap_y + 0.20, right_w - 0.44, 0.36, 18, GOLD, True)
    textbox(
        s,
        "Care close to where families live —\nand a base for learning and research.",
        rx + 0.22,
        cap_y + 0.60,
        right_w - 0.44,
        cap_h - 0.80,
        18,
        WHITE,
        para_gap=8,
    )
    slides.append(s)

    # ------------------------------------------------------------------
    # 4 Community Reach (no model diagram)
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s,
        "Operating model",
        "FairBanks Community Reach",
        "How care, learning, and empowerment connect. FCHIP supports Data & Feedback.",
    )
    steps = [
        ("1", "Community members", "Communities name needs and own solutions"),
        ("2", "CHWs / VHTs", "The bridge — outreach, referrals, and data"),
        ("3", "Community programmes", "Screening, education, maternal & child care"),
        ("4", "Medical Centre", "Clinical care, diagnostics, pharmacy, follow-up"),
        ("5", "Research & skills", "Evidence, partners, training, and learning"),
        ("6", "Empowerment & CHIS", "Livelihoods and affordable shared protection"),
    ]
    gap_x, gap_y = 0.24, 0.18
    cell_w = (CW - gap_x) / 2
    cell_h = (CONTENT_BOTTOM - y0 - 2 * gap_y) / 3
    accents = [TEAL, ORANGE, GREEN, TEAL, ORANGE, GREEN]
    for i, (num, title, desc) in enumerate(steps):
        col, row = i % 2, i // 2
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x, y, 0.85, cell_h, accents[i])
        textbox(s, num, x, y, 0.85, cell_h, 30, WHITE, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        textbox(s, title, x + 1.05, y + 0.26, cell_w - 1.25, 0.48, 23, NAVY, True)
        textbox(s, desc, x + 1.05, y + 0.80, cell_w - 1.25, cell_h - 1.00, 21, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 5 What is FCHIP
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "What is FCHIP?")
    left_w = CW * 0.48
    gap = 0.24
    right_w = CW - left_w - gap
    h = CONTENT_BOTTOM - y0
    rect(s, ML, y0, left_w, h, TEAL, rounded=True)
    textbox(s, "FCHIP stands for", ML + 0.35, y0 + 0.28, left_w - 0.70, 0.38, 18, GOLD, True)
    textbox(s, FCHIP, ML + 0.35, y0 + 0.72, left_w - 0.70, 0.90, 26, WHITE, True)
    textbox(
        s,
        "FCHIP brings community and facility health data together.\n"
        "It uses AI, GIS maps, and climate signals to spot risks early —\n"
        "so CHWs, clinics, and partners can act sooner.",
        ML + 0.35,
        y0 + 1.80,
        left_w - 0.70,
        2.40,
        21,
        LIGHT,
        para_gap=12,
    )
    crop_photo(s, PHOTOS["dashboard"], ML + left_w + gap, y0, right_w, h * 0.55)
    focus = [
        "Disease early warning",
        "Maternal health risk",
        "NCD & child health",
        "CHW mobile tools",
        "GIS & climate links",
        "Facility dashboards",
    ]
    fx = ML + left_w + gap
    fy = y0 + h * 0.55 + 0.16
    fh = CONTENT_BOTTOM - fy
    cols = 2
    gap_x, gap_y = 0.14, 0.12
    cell_w = (right_w - gap_x) / 2
    cell_h = (fh - 2 * gap_y) / 3
    for i, item in enumerate(focus):
        col, row = i % cols, i // cols
        x = fx + col * (cell_w + gap_x)
        y = fy + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        textbox(s, item, x + 0.12, y + 0.08, cell_w - 0.24, cell_h - 0.16, 19, NAVY, True, valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 6 Why MUST
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s,
        "Fit with MUST",
        "Why MUST is a strong partner",
        "Drawn from MUST’s published faculties, programmes, and innovation units.",
    )
    reasons = [
        ("Faculty of Health Sciences", "Doctors, nurses, pharmacists, lab scientists, and physiotherapists — with COBERS community learning since 1989.", TEAL),
        ("Department of Community Health", "MPH, epidemiology, maternal and child health, NCDs, and health systems research.", ORANGE),
        ("Computing & innovation", "Faculty of Computing and Informatics, CITT, and CAMTech Uganda — a natural home for FCHIP.", GREEN),
        ("Grant and research culture", "Long record of partnered research with MoH, districts, and global universities.", GOLD),
    ]
    gap = 0.18
    cell_w = (CW - gap) / 2
    cell_h = (CONTENT_BOTTOM - y0 - gap) / 2
    for i, (t, body, accent) in enumerate(reasons):
        col, row = i % 2, i // 2
        x = ML + col * (cell_w + gap)
        y = y0 + row * (cell_h + gap)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + 0.18, y + 0.22, 0.10, cell_h - 0.44, accent)
        textbox(s, t, x + 0.45, y + 0.28, cell_w - 0.70, 0.55, 22, NAVY, True)
        textbox(s, body, x + 0.45, y + 0.95, cell_w - 0.70, cell_h - 1.25, 20, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 7 How we can work together
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "How we can work together")
    areas = [
        ("A", "Teaching & placements", "COBERS-style learning, clinical attachments, and supervised field work", TEAL),
        ("B", "Joint research", "Community health, digital health, climate–health, MCH, and NCDs", ORANGE),
        ("C", "Community programmes", "Camps, school health, screening, and VHT-linked outreach", GREEN),
        ("D", "FCHIP co-development", "AI, GIS, mobile tools, and dashboards with Computing & CITT", GOLD),
        ("E", "Joint grants", "Shared concept notes, MoUs, and competitive funding bids", TEAL),
    ]
    badge, inset, gap_y = 0.62, 0.22, 0.14
    top = y0 + 0.08
    row_h = (CONTENT_BOTTOM - top - 4 * gap_y) / 5
    for i, (letter, title, desc, accent) in enumerate(areas):
        y = top + i * (row_h + gap_y)
        rect(s, ML, y, CW, row_h, WHITE, LINE, True)
        by = y + (row_h - badge) / 2
        rect(s, ML + inset, by, badge, badge, accent, rounded=True)
        textbox(s, letter, ML + inset, by, badge, badge, 22, WHITE, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        tx = ML + inset + badge + 0.28
        tw = CW - (inset + badge + 0.28) - 0.28
        textbox(s, title, tx, y + (row_h - 0.78) / 2, tw, 0.38, 22, NAVY, True)
        textbox(s, desc, tx, y + (row_h - 0.78) / 2 + 0.38, tw, 0.38, 20, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 8 Joint research & grants
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s,
        "Funding & science",
        "Joint research and grants",
        "Together we write stronger proposals — clinic + community + university science.",
    )
    left_w = (CW - 0.24) * 0.52
    right_w = CW - left_w - 0.24
    h = CONTENT_BOTTOM - y0
    rect(s, ML, y0, left_w, h, WHITE, LINE, True)
    textbox(s, "What we can pursue", ML + 0.32, y0 + 0.28, left_w - 0.60, 0.42, 22, NAVY, True)
    bullets(
        s,
        [
            "Community health intelligence research",
            "Climate–health and early warning pilots",
            "Maternal, child, and NCD studies",
            "Digital health / AI innovation awards",
            "Partner M&E and capacity-building grants",
        ],
        ML + 0.32,
        y0 + 0.90,
        left_w - 0.60,
        h - 1.15,
        21,
        SLATE,
        20,
    )
    rx = ML + left_w + 0.24
    rect(s, rx, y0, right_w, h, TEAL, rounded=True)
    textbox(s, "Why funders listen", rx + 0.32, y0 + 0.35, right_w - 0.64, 0.42, 22, GOLD, True)
    textbox(
        s,
        "MUST brings research quality, ethics, and student power.\n\n"
        "FairBanks brings a licensed facility, Community Reach, and a working FCHIP MVP.\n\n"
        "That mix lowers risk and raises impact for grant reviewers.",
        rx + 0.32,
        y0 + 1.00,
        right_w - 0.64,
        h - 1.40,
        20,
        WHITE,
        para_gap=10,
    )
    slides.append(s)

    # ------------------------------------------------------------------
    # 9 Roadmap + ask
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Next steps", "Roadmap and invitation")
    phases = [
        ("I", "Meet & plan", "Briefing · shared priorities · working group", TEAL),
        ("II", "Formalise", "MoU · 2–3 pilots · ethics & data rules", ORANGE),
        ("III", "Deliver", "Placements · research · first grant bids", GREEN),
        ("IV", "Grow", "Review · publish · scale what works", NAVY),
    ]
    gap = 0.18
    card_w = (CW - 3 * gap) / 4
    card_h = 2.35
    for i, (num, label, detail, accent) in enumerate(phases):
        x = ML + i * (card_w + gap)
        rect(s, x, y0, card_w, card_h, WHITE, LINE, True)
        rect(s, x, y0, card_w, 0.95, accent)
        textbox(s, f"Phase {num}", x + 0.14, y0 + 0.18, card_w - 0.28, 0.32, 18, WHITE, True)
        textbox(s, label, x + 0.14, y0 + 0.50, card_w - 0.28, 0.36, 22, WHITE, True)
        textbox(s, detail, x + 0.14, y0 + 1.15, card_w - 0.28, 1.00, 18, SLATE)
    asks_y = y0 + card_h + 0.22
    asks_h = CONTENT_BOTTOM - asks_y
    asks = [
        "Host a planning meeting with Health Sciences, Community Health, and Computing / CITT",
        "Agree a partnership MoU with clear roles for teaching, research, and FCHIP",
        "Pick first pilots: one teaching track, one research track, one grant concept",
    ]
    rect(s, ML, asks_y, CW, asks_h, WHITE, LINE, True)
    textbox(s, "What we invite MUST to do", ML + 0.35, asks_y + 0.16, CW - 0.70, 0.40, 22, NAVY, True)
    bullets(s, asks, ML + 0.35, asks_y + 0.60, CW - 0.70, asks_h - 0.75, 21, SLATE, 16)
    slides.append(s)

    # ------------------------------------------------------------------
    # 10 Closing
    # ------------------------------------------------------------------
    s = new_slide(NAVY)
    panel = 7.85
    crop_photo(s, PHOTOS["facility_wide"], 0, 0, SW, SH)
    rect(s, 0, 0, panel, SH, NAVY)
    rect(s, panel, 0, 0.12, SH, GOLD)
    add_logo(s, 0.55, 0.28, 0.50)

    textbox(s, "INVITATION TO PARTNER", 0.55, 1.00, 7.0, 0.36, 18, GOLD, True)
    textbox(s, ORG, 0.55, 1.42, 7.0, 0.40, 22, TEAL, True)
    textbox(s, "Let us build healthier", 0.55, 2.05, 7.0, 0.52, 34, WHITE, True)
    textbox(s, "communities — and stronger", 0.55, 2.60, 7.0, 0.52, 34, WHITE, True)
    textbox(s, "science — together.", 0.55, 3.15, 7.0, 0.52, 34, GOLD, True)
    rect(s, 0.55, 3.75, 2.80, 0.09, GOLD)

    textbox(
        s,
        f"FairBanks Medical Centre invites {MUST}\nto partner on Community Reach, FCHIP, and joint grants.",
        0.55,
        4.05,
        7.0,
        0.80,
        19,
        LIGHT,
        para_gap=8,
    )
    textbox(s, SLOGAN, 0.55, 4.95, 7.0, 0.40, 24, GOLD, True)

    rect(s, 0.55, 5.45, 6.90, 1.55, TEAL, rounded=True)
    textbox(s, "Institutional contact", 0.80, 5.58, 6.4, 0.34, 17, GOLD, True)
    textbox(
        s,
        "Kyebando–Kisalosalo, Tirupati Road, Kampala\n"
        "info@fairbanksmedicalcentre.org  ·  +256 748 319 052\n"
        "fairbanksmedicalcentre.org",
        0.80,
        5.98,
        6.4,
        0.90,
        18,
        WHITE,
        para_gap=8,
    )
    slides.append(s)

    total = len(slides)
    assert total <= 10, f"Expected ≤10 slides, got {total}"
    for i, s in enumerate(slides):
        if i not in (0, total - 1):
            footer(s, i + 1, total)

    prs.save(OUT_PPT)
    print(f"Wrote {OUT_PPT} ({total} slides)")


if __name__ == "__main__":
    build()
