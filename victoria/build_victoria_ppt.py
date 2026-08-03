#!/usr/bin/env python3
"""
Victoria University partnership proposal deck.
Balanced layout, explicit font sizes, clear spacing.
FCHIP = FairBanks Community Health Intelligence Platform.
No personal names or phone numbers.
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
OUT_PPT = OUT / "victoria_ppt.pptx"

NAVY, TEAL, GREEN = "0A1F2E", "0D6E6E", "2D7A55"
ORANGE, GOLD, CREAM = "C45C26", "D99A2B", "F7F5F0"
PALE_TEAL, WHITE = "E8F3F2", "FFFFFF"
SLATE, MUTED, LINE, LIGHT = "1E2F38", "52636C", "CED9D8", "D4E8DC"

ORG = "FairBanks Medical Centre"
SLOGAN = "Your health, our mission."
FOOTER = f"{ORG}  ·  {SLOGAN}"
FCHIP_FULL = "FairBanks Community Health Intelligence Platform"

PHOTOS = {
    "cover": ASSETS / "facility_exterior_branded_entrance_01.jpeg",
    "facility": ASSETS / "facility_exterior_entrance_01.jpg",
    "facility_wide": ASSETS / "facility_exterior_branded_entrance_02.jpeg",
    "outreach": ASSETS / "outreach_bp_screening.jpeg",
    "camp": ASSETS / "outreach_medical_camp_01.jpg",
    "mission": ASSETS / "reception_mission_wall.jpeg",
    "audience": ASSETS / "outreach_audience_full_group_01.jpg",
}

# Layout grid
ML, MR = 0.65, 0.60
SW, SH = 13.333, 7.5
CW = SW - ML - MR
TOP_BAR = 0.06
LOGO_H = 0.38
LOGO_W = LOGO_H * (269 / 101)
LOGO_Y = 0.28
FOOTER_Y = 7.08
CONTENT_BOTTOM = 6.70
HEADER_Y = 1.20
HEADER_Y_SUB = 1.55


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
        size=16,
        color=SLATE,
        bold=False,
        align=PP_ALIGN.LEFT,
        valign=MSO_ANCHOR.TOP,
    ):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(18000)
        tf.margin_top = tf.margin_bottom = Emu(6000)
        tf.vertical_anchor = valign
        lines = value.split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.space_before = Pt(0)
            p.space_after = Pt(4) if len(lines) > 1 else Pt(0)
            r = p.add_run()
            set_run(r, line, size, color, bold)
        return box

    def bullets(slide, items, x, y, w, h, size=15, color=SLATE, space=8):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(18000)
        tf.margin_top = tf.margin_bottom = Emu(4000)
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(space)
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
        tw = logo_x - ML - 0.45
        if kicker.strip():
            textbox(s, kicker.upper(), ML, 0.22, tw, 0.26, 13, ORANGE, True)
            textbox(s, title, ML, 0.48, tw, 0.50, 24, NAVY, True)
            y = 1.05
        else:
            textbox(s, title, ML, 0.30, tw, 0.55, 26, NAVY, True)
            y = 0.95
        if subtitle:
            textbox(s, subtitle, ML, y + 0.05, CW, 0.35, 15, MUTED)
            return HEADER_Y_SUB
        return HEADER_Y

    def footer(s, number, total):
        rect(s, ML, FOOTER_Y, CW, 0.012, LINE)
        textbox(s, FOOTER, ML, FOOTER_Y + 0.08, 9.8, 0.28, 11, MUTED)
        textbox(s, f"{number}  /  {total}", SW - MR - 1.3, FOOTER_Y + 0.08, 1.3, 0.28, 11, MUTED, align=PP_ALIGN.RIGHT)

    def card(s, x, y, w, h, title, items, accent=TEAL):
        rect(s, x, y, w, h, WHITE, LINE, True)
        rect(s, x + 0.14, y + 0.22, 0.08, h - 0.44, accent)
        textbox(s, title, x + 0.38, y + 0.28, w - 0.58, 0.60, 17, NAVY, True)
        bullets(s, items, x + 0.32, y + 1.05, w - 0.52, h - 1.25, 16, SLATE, space=12)

    slides = []

    # ------------------------------------------------------------------
    # 1 Cover
    # ------------------------------------------------------------------
    s = new_slide(NAVY)
    panel = 7.85
    crop_photo(s, PHOTOS["cover"], 0, 0, SW, SH)
    rect(s, 0, 0, panel, SH, NAVY)
    rect(s, panel, 0, 0.12, SH, GOLD)
    add_logo(s, 0.55, 0.28, 0.48)

    textbox(s, "STRATEGIC INSTITUTIONAL PROPOSAL", 0.55, 1.05, 7.0, 0.32, 15, GOLD, True)
    textbox(s, ORG, 0.55, 1.42, 7.0, 0.40, 20, TEAL, True)
    textbox(s, "University–Community Health", 0.55, 2.05, 7.0, 0.65, 32, WHITE, True)
    textbox(s, "Partnership Proposal", 0.55, 2.75, 7.0, 0.65, 32, GOLD, True)
    rect(s, 0.55, 3.50, 2.80, 0.09, GOLD)

    textbox(s, f"Building the {FCHIP_FULL}", 0.55, 3.80, 7.0, 0.40, 16, LIGHT)
    textbox(s, "(FCHIP) with Victoria University", 0.55, 4.30, 7.0, 0.42, 20, GOLD, True)
    textbox(s, SLOGAN, 0.55, 4.95, 7.0, 0.42, 22, GOLD, True)

    rect(s, 0, 6.30, panel, 1.20, TEAL)
    textbox(s, "Prepared for Victoria University leadership", 0.55, 6.50, 7.0, 0.38, 18, WHITE, True)
    textbox(s, "Kyebando–Kisalosalo, Kampala  ·  Institutional briefing", 0.55, 6.95, 7.0, 0.35, 15, LIGHT)
    slides.append(s)

    # ------------------------------------------------------------------
    # 2 Slide Contents
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "Slide Contents")
    items = [
        ("01", "Executive summary"),
        ("02", "About FairBanks Medical Centre"),
        ("03", "What FCHIP is"),
        ("04", "How we work together"),
        ("05", "Why Victoria University"),
        ("06", "Partnership network and outcomes"),
        ("07", "Roadmap and long-term vision"),
        ("08", "Invitation to partner"),
    ]
    gap_x, gap_y = 0.30, 0.22
    card_w = (CW - gap_x) / 2
    avail = CONTENT_BOTTOM - y0 - 0.10
    card_h = (avail - 3 * gap_y) / 4
    for i, (num, label) in enumerate(items):
        col, row = i % 2, i // 2
        x = ML + col * (card_w + gap_x)
        y = y0 + 0.10 + row * (card_h + gap_y)
        rect(s, x, y, card_w, card_h, WHITE, LINE, True)
        rect(s, x + 0.20, y + 0.25, 0.08, card_h - 0.50, ORANGE if col == 0 else TEAL)
        mid = y + (card_h - 0.42) / 2
        textbox(s, num, x + 0.42, mid, 0.70, 0.42, 20, ORANGE, True, valign=MSO_ANCHOR.MIDDLE)
        textbox(s, label, x + 1.20, mid, card_w - 1.45, 0.42, 17, NAVY, True, valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 3 Executive summary
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Executive summary", "A strategic partnership for health, learning, and impact")
    textbox(
        s,
        "FairBanks Medical Centre proposes a strategic partnership with Victoria University "
        "to integrate healthcare service delivery, academic excellence, research, innovation, "
        "and community engagement.",
        ML, y0, CW, 0.70, 16, SLATE,
    )
    points = [
        ("FCHIP at the centre",
         f"The {FCHIP_FULL} is the deep-tech platform for predictive community health — "
         "using community and facility data to support prevention, outreach, research, "
         "skills development, and partnerships.", TEAL),
        ("Shared strengths",
         "Victoria University brings education, innovation, and research. FairBanks brings "
         "a practical healthcare platform and community presence.", ORANGE),
        ("Shared goal",
         "Build a replicable University–Community Health Partnership Model that advances "
         "healthcare, education, innovation, and sustainable development in Uganda.", GREEN),
    ]
    start = y0 + 0.85
    gap = 0.18
    row_h = (CONTENT_BOTTOM - start - 2 * gap) / 3
    for i, (t, body, accent) in enumerate(points):
        y = start + i * (row_h + gap)
        rect(s, ML, y, CW, row_h, WHITE, LINE, True)
        rect(s, ML + 0.18, y + 0.20, 0.08, row_h - 0.40, accent)
        textbox(s, t, ML + 0.42, y + 0.18, CW - 0.70, 0.35, 17, NAVY, True)
        textbox(s, body, ML + 0.42, y + 0.55, CW - 0.70, row_h - 0.70, 15, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 4 About FairBanks
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Who we are", "About FairBanks Medical Centre")
    left_w, gap = 7.00, 0.28
    right_w = CW - left_w - gap
    box_h = CONTENT_BOTTOM - y0
    rect(s, ML, y0, left_w, box_h, WHITE, LINE, True)
    textbox(s, "Community-based primary healthcare", ML + 0.35, y0 + 0.28, left_w - 0.70, 0.38, 18, NAVY, True)
    textbox(s, SLOGAN, ML + 0.35, y0 + 0.70, left_w - 0.70, 0.32, 15, ORANGE, True)
    bullets(
        s,
        [
            "Located in Kyebando–Kisalosalo, Kampala",
            "Accessible, affordable, quality healthcare",
            "Strong focus on preventive health and community well-being",
            "Committed to social determinants of health through partnerships",
            "Active in innovation, research, and community engagement",
        ],
        ML + 0.35, y0 + 1.20, left_w - 0.70, box_h - 1.50, 15, SLATE, 11,
    )
    rx = ML + left_w + gap
    photo_h = box_h * 0.58
    crop_photo(s, PHOTOS["facility"], rx, y0, right_w, photo_h)
    cap_y = y0 + photo_h + 0.16
    cap_h = CONTENT_BOTTOM - cap_y
    rect(s, rx, cap_y, right_w, cap_h, TEAL, rounded=True)
    textbox(s, ORG, rx + 0.25, cap_y + 0.28, right_w - 0.50, 0.35, 15, GOLD, True)
    textbox(
        s,
        "A trusted community health home — care close to where families live.",
        rx + 0.25, cap_y + 0.70, right_w - 0.50, cap_h - 0.95, 14, WHITE,
    )
    slides.append(s)

    # ------------------------------------------------------------------
    # 5 What is FCHIP
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "What is FCHIP?")
    def_h = 2.50
    rect(s, ML, y0, CW, def_h, TEAL, rounded=True)
    textbox(s, "FCHIP stands for", ML + 0.40, y0 + 0.22, CW - 0.80, 0.32, 15, GOLD, True)
    textbox(s, FCHIP_FULL, ML + 0.40, y0 + 0.58, CW - 0.80, 0.45, 22, WHITE, True)
    textbox(
        s,
        "FCHIP is a deep-technology platform that turns community and facility health data into "
        "useful predictions. It connects community members, CHWs and VHTs, schools, outreach "
        "programmes, clinics, and hospitals — and uses AI, GIS mapping, and climate signals to "
        "spot risks early so partners can act before crises escalate.",
        ML + 0.40, y0 + 1.15, CW - 0.80, 1.15, 15, LIGHT,
    )
    textbox(s, "What FCHIP focuses on", ML, y0 + def_h + 0.18, CW, 0.32, 16, NAVY, True)
    focus = [
        "Disease early warning", "Maternal health risk", "NCD hotspots",
        "Child health signals", "Medicine demand forecasts", "CHW / VHT mobile tools",
        "GIS risk mapping", "Climate–health fusion", "EMR / HMS data APIs", "Facility dashboards",
    ]
    cols, rows = 5, 2
    gap_x, gap_y = 0.20, 0.16
    grid_top = y0 + def_h + 0.55
    cell_w = (CW - (cols - 1) * gap_x) / cols
    cell_h = (CONTENT_BOTTOM - grid_top - (rows - 1) * gap_y) / rows
    accents = [TEAL, ORANGE, GREEN, GOLD, TEAL]
    for i, item in enumerate(focus):
        col, row = i % cols, i // cols
        x = ML + col * (cell_w + gap_x)
        y = grid_top + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + 0.14, y + 0.14, cell_w - 0.28, 0.07, accents[col])
        textbox(s, item, x + 0.10, y + 0.30, cell_w - 0.20, cell_h - 0.42, 14, NAVY, True,
                align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 6 How we work together
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "How we work together")
    textbox(
        s,
        "Five strategic areas of collaboration between Victoria University and FairBanks Medical Centre.",
        ML, y0, CW, 0.38, 15, MUTED,
    )
    areas = [
        ("A", "Student Experiential Learning", "Structured placements across health and related fields", TEAL),
        ("B", "Research and Innovation", "Joint studies, publications, and evidence for policy", ORANGE),
        ("C", "Community Engagement", "Camps, schools, screening, and volunteer programmes", GREEN),
        ("D", "Digital Health & Innovation", "AI, EMRs, telemedicine, and predictive analytics", GOLD),
        ("E", "Resource Mobilisation", "Shared pursuit of grants and development partnerships", TEAL),
    ]
    badge, inset, gap_y = 0.55, 0.22, 0.16
    top = y0 + 0.48
    row_h = (CONTENT_BOTTOM - top - 4 * gap_y) / 5
    for i, (letter, title, desc, accent) in enumerate(areas):
        y = top + i * (row_h + gap_y)
        rect(s, ML, y, CW, row_h, WHITE, LINE, True)
        by = y + (row_h - badge) / 2
        rect(s, ML + inset, by, badge, badge, accent, rounded=True)
        textbox(s, letter, ML + inset, by, badge, badge, 18, WHITE, True,
                align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        tx = ML + inset + badge + 0.28
        tw = CW - (inset + badge + 0.28) - 0.28
        textbox(s, title, tx, y + (row_h - 0.62) / 2, tw, 0.30, 17, NAVY, True)
        textbox(s, desc, tx, y + (row_h - 0.62) / 2 + 0.30, tw, 0.30, 15, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 7 Why Victoria University
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "", "Why Victoria University?")
    textbox(
        s,
        "Victoria University is committed to practical learning, innovation, entrepreneurship, "
        "and community transformation — values that closely align with FairBanks Medical Centre.",
        ML, y0, CW, 0.60, 15, SLATE,
    )
    roles = [
        ("What the University contributes",
         ["Academic expertise", "Research capacity", "Student engagement", "Innovation culture"], TEAL),
        ("What the University gains",
         ["Real-world experiential learning", "Applied research environment",
          "Community impact pathways", "Visible institutional leadership"], ORANGE),
        ("Proposed role",
         ["Founding academic partner for FCHIP", "Co-shaper of a national partnership model",
          "Bridge between campus and community", "Partner in grants and innovation"], GREEN),
    ]
    gap = 0.25
    card_w = (CW - 2 * gap) / 3
    card_top = y0 + 0.75
    card_h = CONTENT_BOTTOM - card_top
    for i, (title, items, accent) in enumerate(roles):
        card(s, ML + i * (card_w + gap), card_top, card_w, card_h, title, items, accent)
    slides.append(s)

    # ------------------------------------------------------------------
    # 8 Student learning
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Collaboration A", "Student experiential learning",
        "Students gain practical experience while contributing to community healthcare.",
    )
    disciplines = [
        "Nursing", "Public Health", "Human Resource Management", "Business Administration",
        "Information Technology", "Social Work", "Community Development", "Marketing and Communications",
    ]
    cols, rows = 4, 2
    gap_x, gap_y = 0.24, 0.24
    cell_w = (CW - 3 * gap_x) / 4
    cell_h = (CONTENT_BOTTOM - y0 - gap_y) / 2
    for i, d in enumerate(disciplines):
        col, row = i % 4, i // 4
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + (cell_w - 0.70) / 2, y + 0.45, 0.70, 0.07, ORANGE)
        textbox(s, d, x + 0.15, y + 0.70, cell_w - 0.30, cell_h - 1.00, 15, NAVY, True,
                align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 9 Research
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Collaboration B", "Research and innovation",
        "Joint programmes, publications, conferences, and evidence-based policy development.",
    )
    themes = [
        "Community Health", "Health Systems Strengthening", "Artificial Intelligence in Healthcare",
        "Digital Health", "Occupational Burnout", "Maternal and Child Health",
        "Non-Communicable Diseases", "Preventive Healthcare", "Healthcare Quality Improvement",
    ]
    cols, rows = 3, 3
    gap_x, gap_y = 0.24, 0.18
    cell_w = (CW - 2 * gap_x) / 3
    cell_h = (CONTENT_BOTTOM - y0 - 2 * gap_y) / 3
    row_accent = [TEAL, ORANGE, GREEN]
    for i, t in enumerate(themes):
        col, row = i % 3, i // 3
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + 0.18, y + 0.22, 0.08, cell_h - 0.44, row_accent[row])
        textbox(s, t, x + 0.40, y + 0.22, cell_w - 0.60, cell_h - 0.44, 15, NAVY, True,
                valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 10 Community engagement
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Collaboration C", "Community engagement through FCHIP",
        "Measurable social impact with meaningful practical learning.",
    )
    photo_h = 1.55
    gap_p = 0.22
    pw = (CW - 2 * gap_p) / 3
    crop_photo(s, PHOTOS["outreach"], ML, y0, pw, photo_h)
    crop_photo(s, PHOTOS["camp"], ML + pw + gap_p, y0, pw, photo_h)
    crop_photo(s, PHOTOS["audience"], ML + 2 * (pw + gap_p), y0, pw, photo_h)
    activities = [
        "Community medical camps", "School health programmes", "Corporate wellness initiatives",
        "Community disease screening", "Health promotion campaigns", "Maternal and child health outreach",
        "Elderly care programmes", "Student volunteer programmes", "Community-based research",
    ]
    cols, rows = 3, 3
    gap_x, gap_y = 0.22, 0.14
    grid_top = y0 + photo_h + 0.22
    cell_w = (CW - 2 * gap_x) / 3
    cell_h = (CONTENT_BOTTOM - grid_top - 2 * gap_y) / 3
    for i, a in enumerate(activities):
        col, row = i % 3, i // 3
        x = ML + col * (cell_w + gap_x)
        y = grid_top + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, PALE_TEAL if row % 2 == 0 else WHITE, LINE, True)
        textbox(s, a, x + 0.18, y + 0.10, cell_w - 0.36, cell_h - 0.20, 14, NAVY, True,
                valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 11 Digital health
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Collaboration D", "Digital health and innovation",
        "Developing practical healthcare solutions together.",
    )
    digi = [
        ("Artificial Intelligence", "Applications that support smarter clinical and community decisions"),
        ("Electronic Medical Records", "Stronger continuity of care and better clinical information"),
        ("Telemedicine", "Care and advice that reach people beyond the facility walls"),
        ("Community Health Information Systems", "Better visibility of community health activity and needs"),
        ("Predictive analytics", "Earlier insight for planning and prevention"),
        ("Digital patient engagement", "Clearer communication and follow-up with patients and families"),
    ]
    cols, rows = 3, 2
    gap_x, gap_y = 0.24, 0.22
    cell_w = (CW - 2 * gap_x) / 3
    cell_h = (CONTENT_BOTTOM - y0 - gap_y) / 2
    accents = [TEAL, ORANGE, GREEN, GOLD, TEAL, ORANGE]
    for i, (t, d) in enumerate(digi):
        col, row = i % 3, i // 3
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + 0.22, y + 0.22, cell_w - 0.44, 0.08, accents[i])
        textbox(s, t, x + 0.25, y + 0.45, cell_w - 0.50, 0.55, 15, NAVY, True)
        textbox(s, d, x + 0.25, y + 1.10, cell_w - 0.50, cell_h - 1.30, 14, MUTED)
    slides.append(s)

    # ------------------------------------------------------------------
    # 12 Resource mobilisation
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Collaboration E", "Resource mobilisation",
        "Collaborative proposals improve competitiveness for local and international funding.",
    )
    funds = [
        "Research grants", "Development funding", "Innovation grants",
        "Corporate Social Responsibility partnerships",
        "International development partnerships", "Capacity-building programmes",
    ]
    cols, rows = 3, 2
    gap_x, gap_y = 0.24, 0.22
    cell_w = (CW - 2 * gap_x) / 3
    cell_h = (CONTENT_BOTTOM - y0 - gap_y) / 2
    for i, f in enumerate(funds):
        col, row = i % 3, i // 3
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        textbox(s, f"{i + 1:02}", x + 0.30, y + 0.40, cell_w - 0.60, 0.40, 20, ORANGE, True)
        textbox(s, f, x + 0.30, y + 1.00, cell_w - 0.60, cell_h - 1.30, 15, NAVY, True)
    slides.append(s)

    # ------------------------------------------------------------------
    # 13 Network
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(
        s, "Ecosystem", "Strategic partnerships and collaborative network",
        "Victoria University as the flagship academic and research partner.",
    )
    partners = [
        "Government ministries and public health institutions",
        "Local government leadership and Village Health Teams",
        "Health insurance providers",
        "Specialist healthcare professionals",
        "Schools and educational institutions",
        "Faith-based and community organisations",
        "Corporate organisations",
        "Media partners",
        "National and international development partners",
        "Philanthropic foundations and research institutions",
    ]
    cols, rows = 2, 5
    gap_x, gap_y = 0.24, 0.12
    cell_w = (CW - gap_x) / 2
    cell_h = (CONTENT_BOTTOM - y0 - 4 * gap_y) / 5
    for i, p in enumerate(partners):
        col, row = i % 2, i // 2
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        rect(s, x + 0.16, y + 0.16, 0.08, cell_h - 0.32, TEAL if col == 0 else ORANGE)
        textbox(s, p, x + 0.38, y + 0.10, cell_w - 0.55, cell_h - 0.20, 14, SLATE, True,
                valign=MSO_ANCHOR.MIDDLE)
    slides.append(s)

    # ------------------------------------------------------------------
    # 14 Outcomes
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Impact", "Expected outcomes of the partnership")
    outcomes = [
        "Improve access to quality primary healthcare",
        "Strengthen preventive healthcare programmes",
        "Expand experiential learning opportunities",
        "Increase community-based research",
        "Promote healthcare innovation",
        "Strengthen graduate employability",
        "Enhance institutional visibility",
        "Mobilise sustainable development partnerships",
        "Improve health outcomes in surrounding communities",
    ]
    cols, rows = 3, 3
    gap_x, gap_y = 0.24, 0.16
    cell_w = (CW - 2 * gap_x) / 3
    cell_h = (CONTENT_BOTTOM - y0 - 2 * gap_y) / 3
    for i, o in enumerate(outcomes):
        col, row = i % 3, i // 3
        x = ML + col * (cell_w + gap_x)
        y = y0 + row * (cell_h + gap_y)
        rect(s, x, y, cell_w, cell_h, WHITE, LINE, True)
        textbox(s, f"{i + 1:02}", x + 0.28, y + 0.22, cell_w - 0.56, 0.32, 14, ORANGE, True)
        textbox(s, o, x + 0.28, y + 0.60, cell_w - 0.56, cell_h - 0.80, 14, NAVY, True)
    slides.append(s)

    # ------------------------------------------------------------------
    # 15 Roadmap
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Implementation", "Proposed roadmap")
    phases = [
        ("Phase I", "Foundation",
         ["Introductory planning meeting", "Joint Technical Working Group"], TEAL),
        ("Phase II", "Formalisation",
         ["Sign Memorandum of Understanding", "Identify priority pilot projects"], ORANGE),
        ("Phase III", "Delivery",
         ["Student placements", "Joint research and outreach", "Innovation and grant development"], GREEN),
        ("Phase IV", "Learning and scale",
         ["Monitoring and evaluation", "Document lessons", "Expand and replicate"], NAVY),
    ]
    gap = 0.22
    card_w = (CW - 3 * gap) / 4
    card_h = CONTENT_BOTTOM - y0
    for i, (phase, label, items, accent) in enumerate(phases):
        x = ML + i * (card_w + gap)
        rect(s, x, y0, card_w, card_h, WHITE, LINE, True)
        rect(s, x, y0, card_w, 1.05, accent)
        textbox(s, phase, x + 0.18, y0 + 0.18, card_w - 0.36, 0.30, 14, WHITE, True)
        textbox(s, label, x + 0.18, y0 + 0.52, card_w - 0.36, 0.38, 16, WHITE, True)
        bullets(s, items, x + 0.16, y0 + 1.30, card_w - 0.32, card_h - 1.50, 14, SLATE, 10)
    slides.append(s)

    # ------------------------------------------------------------------
    # 16 Vision
    # ------------------------------------------------------------------
    s = new_slide()
    y0 = header(s, "Looking ahead", "Long-term vision")
    hero_h = 2.20
    left_w = CW * 0.58
    gap = 0.25
    right_w = CW - left_w - gap
    rect(s, ML, y0, left_w, hero_h, TEAL, rounded=True)
    textbox(
        s,
        "Establish FCHIP as Uganda's leading University–Community Health Partnership Model — "
        "showing how academia, healthcare providers, government, development partners, and "
        "communities can work together to improve lives.",
        ML + 0.38, y0 + 0.35, left_w - 0.76, hero_h - 0.70, 16, WHITE, True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    crop_photo(s, PHOTOS["mission"], ML + left_w + gap, y0, right_w, hero_h)
    half_gap = 0.25
    half = (CW - half_gap) / 2
    pair_y = y0 + hero_h + 0.25
    pair_h = CONTENT_BOTTOM - pair_y
    rect(s, ML, pair_y, half, pair_h, WHITE, LINE, True)
    textbox(s, "For FairBanks Medical Centre", ML + 0.30, pair_y + 0.28, half - 0.55, 0.35, 15, ORANGE, True)
    textbox(
        s,
        "A living laboratory for education, research, innovation, and community service.",
        ML + 0.30, pair_y + 0.75, half - 0.55, pair_h - 1.00, 15, NAVY, True,
    )
    rx = ML + half + half_gap
    rect(s, rx, pair_y, half, pair_h, WHITE, LINE, True)
    textbox(s, "For Victoria University", rx + 0.30, pair_y + 0.28, half - 0.55, 0.35, 15, ORANGE, True)
    textbox(
        s,
        "Stronger leadership in experiential learning, applied research, and social impact — "
        "with lessons that can be replicated across Uganda.",
        rx + 0.30, pair_y + 0.75, half - 0.55, pair_h - 1.00, 15, NAVY, True,
    )
    slides.append(s)

    # ------------------------------------------------------------------
    # 17 Closing
    # ------------------------------------------------------------------
    s = new_slide(NAVY)
    panel = 7.85
    crop_photo(s, PHOTOS["facility_wide"], 0, 0, SW, SH)
    rect(s, 0, 0, panel, SH, NAVY)
    rect(s, panel, 0, 0.12, SH, GOLD)
    add_logo(s, 0.55, 0.28, 0.48)

    textbox(s, "INVITATION TO PARTNER", 0.55, 1.05, 7.0, 0.32, 15, GOLD, True)
    textbox(s, ORG, 0.55, 1.42, 7.0, 0.38, 18, TEAL, True)
    textbox(s, "Together, we can build", 0.55, 2.00, 7.0, 0.48, 28, WHITE, True)
    textbox(s, "healthier communities", 0.55, 2.52, 7.0, 0.48, 28, WHITE, True)
    textbox(s, "and stronger professionals.", 0.55, 3.04, 7.0, 0.48, 28, GOLD, True)
    rect(s, 0.55, 3.60, 2.80, 0.09, GOLD)

    textbox(
        s,
        "FairBanks Medical Centre invites Victoria University to join as a\n"
        "strategic partner in advancing FCHIP.",
        0.55, 3.90, 7.0, 0.70, 16, LIGHT,
    )
    textbox(s, SLOGAN, 0.55, 4.70, 7.0, 0.38, 20, GOLD, True)

    rect(s, 0.55, 5.30, 6.90, 1.70, TEAL, rounded=True)
    textbox(s, "Institutional contact", 0.80, 5.50, 6.4, 0.32, 14, GOLD, True)
    textbox(
        s,
        "Kyebando–Kisalosalo, Tirupati Road, Kampala\n"
        "info@fairbanksmedicalcentre.org\n"
        "fairbanksmedicalcentre.org",
        0.80, 5.90, 6.4, 0.95, 15, WHITE,
    )
    slides.append(s)

    total = len(slides)
    for i, s in enumerate(slides):
        if i not in (0, total - 1):
            footer(s, i + 1, total)

    prs.save(OUT_PPT)
    print(f"Wrote {OUT_PPT} ({total} slides)")


if __name__ == "__main__":
    build()
