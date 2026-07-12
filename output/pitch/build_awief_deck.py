"""
AWIEF Pitch n Grow 2026 — FairBanks / FCIN / CHIP investor pitch deck (12 slides).
Image-rich widescreen deck with large, high-visibility typography.
"""

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- Brand palette (HealthTech: deep teal + navy) ---
NAVY = RGBColor(0x0A, 0x1F, 0x2E)
TEAL = RGBColor(0x0D, 0x6E, 0x6E)
TEAL_LIGHT = RGBColor(0x14, 0xA3, 0xA3)
CREAM = RGBColor(0xF7, 0xF5, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SLATE = RGBColor(0x1E, 0x2F, 0x38)       # darker for readability
MUTED = RGBColor(0x3A, 0x4A, 0x54)       # darker secondary text
ACCENT = RGBColor(0xC4, 0x5C, 0x26)
LINE = RGBColor(0xD0, 0xDC, 0xDC)
SOFT_TEAL = RGBColor(0xE4, 0xF2, 0xF2)
LIGHT = RGBColor(0xE8, 0xF0, 0xF0)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

ASSETS = Path(__file__).resolve().parents[1] / "assets"

# Visibility scale (pt)
H1 = 34
H2 = 30
H3 = 20
BODY = 17
BODY_SM = 15
CAPTION = 13
FOOTER = 12


def img(name: str) -> str:
    path = ASSETS / name
    if not path.exists():
        raise FileNotFoundError(f"Missing asset: {path}")
    return str(path)


def set_run(run, size=18, bold=False, color=SLATE, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    fill_solid(shape, color)
    return shape


def add_round_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    fill_solid(shape, color)
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def textbox(slide, left, top, width, height, text, size=18, bold=False, color=SLATE,
            align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def bullet_box(slide, left, top, width, height, items, size=BODY_SM, color=SLATE, bullet="•"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(6 if i else 0)
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = f"{bullet}  {item}"
        set_run(run, size=size, bold=False, color=color)
    return box


def footer(slide, page, total=12):
    add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.35), NAVY)
    textbox(
        slide, Inches(0.45), Inches(7.17), Inches(9.5), Inches(0.3),
        "AWIEF Pitch n Grow 2026  |  Deep Roots. Digital Futures.  |  HealthTech · AI/ML",
        size=FOOTER, color=LIGHT,
    )
    textbox(
        slide, Inches(11.0), Inches(7.17), Inches(1.8), Inches(0.3),
        f"{page} / {total}", size=FOOTER, bold=True, color=LIGHT, align=PP_ALIGN.RIGHT,
    )


def add_picture_cover(slide, path, left, top, width, height):
    pic = slide.shapes.add_picture(path, left, top, width=width)
    nat_w, nat_h = pic.image.size
    target_ratio = width / height
    img_ratio = nat_w / nat_h
    if img_ratio > target_ratio:
        new_w = height * img_ratio
        pic.height = height
        pic.width = int(new_w)
        crop_frac = (pic.width - width) / (2 * pic.width)
        pic.crop_left = crop_frac
        pic.crop_right = crop_frac
        pic.left = left
        pic.top = top
        pic.width = width
        pic.height = height
    else:
        new_h = width / img_ratio
        pic.width = width
        pic.height = int(new_h)
        crop_frac = (pic.height - height) / (2 * pic.height)
        pic.crop_top = crop_frac
        pic.crop_bottom = crop_frac
        pic.left = left
        pic.top = top
        pic.width = width
        pic.height = height
    return pic


def image_card(slide, path, left, top, width, height, caption=None):
    frame = add_round_rect(slide, left, top, width, height, WHITE)
    frame.line.color.rgb = LINE
    frame.line.width = Pt(1)
    inset = Inches(0.06)
    cap_h = Inches(0.42) if caption else 0
    add_picture_cover(
        slide, path,
        left + inset, top + inset,
        width - 2 * inset,
        height - (cap_h + inset if caption else 2 * inset),
    )
    if caption:
        textbox(
            slide, left + Inches(0.1), top + height - Inches(0.4),
            width - Inches(0.2), Inches(0.36),
            caption, size=CAPTION, bold=True, color=SLATE, align=PP_ALIGN.CENTER,
        )


def section_band(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), TEAL)
    add_rect(slide, 0, Inches(0.12), SLIDE_W, Inches(7.03), CREAM)
    textbox(slide, Inches(0.45), Inches(0.25), Inches(12.4), Inches(0.5),
            title, size=H2, bold=True, color=NAVY)
    if subtitle:
        textbox(slide, Inches(0.45), Inches(0.75), Inches(12.4), Inches(0.4),
                subtitle, size=BODY, color=MUTED)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # ========== 1. TITLE ==========
    s = prs.slides.add_slide(blank)
    add_picture_cover(s, img("hero_chw_mobile.png"), 0, 0, SLIDE_W, SLIDE_H)
    add_rect(s, 0, 0, Inches(7.4), SLIDE_H, NAVY)
    add_rect(s, Inches(7.25), 0, Inches(0.14), SLIDE_H, TEAL_LIGHT)
    textbox(s, Inches(0.5), Inches(0.85), Inches(6.5), Inches(0.4),
            "AWIEF Pitch n Grow 2026  ·  Startup Track  ·  HealthTech",
            size=BODY_SM, bold=True, color=TEAL_LIGHT)
    textbox(s, Inches(0.5), Inches(1.4), Inches(6.6), Inches(1.2),
            "FairBanks Community\nIntelligence Network",
            size=H1, bold=True, color=WHITE)
    textbox(s, Inches(0.5), Inches(2.75), Inches(6.5), Inches(0.45),
            "Community Health Intelligence Platform (CHIP)",
            size=H3, bold=True, color=TEAL_LIGHT)
    textbox(s, Inches(0.5), Inches(3.25), Inches(6.5), Inches(0.35),
            "Your health, our mission.",
            size=BODY, bold=True, color=ACCENT)
    textbox(s, Inches(0.5), Inches(3.7), Inches(6.5), Inches(1.15),
            "AI-powered community health intelligence that turns last-mile data into "
            "predictions — so African primary care prevents crises instead of waiting for them.",
            size=BODY, color=LIGHT)
    textbox(s, Inches(0.5), Inches(5.05), Inches(6.5), Inches(0.4),
            "Theme: Deep Roots. Digital Futures.",
            size=BODY, bold=True, color=ACCENT)
    textbox(s, Inches(0.5), Inches(5.5), Inches(6.5), Inches(0.4),
            "Uganda pilot  →  East Africa  →  Pan-African scale",
            size=BODY_SM, bold=True, color=LIGHT)
    textbox(s, Inches(0.5), Inches(6.35), Inches(6.5), Inches(0.4),
            "Woman-led deep-tech  |  Live medical centre + CHW/VHT foundation",
            size=BODY_SM, color=RGBColor(0xB0, 0xC4, 0xC8))

    # ========== 2. PROBLEM ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "The problem", "Primary healthcare waits for sickness — intelligence arrives too late.")
    image_card(s, img("reactive_clinic.png"), Inches(0.4), Inches(1.3), Inches(5.2), Inches(5.55),
               "Reactive clinic — patients arrive after the crisis starts")
    problems = [
        ("Facilities wait for sickness", "Outbreaks detected too late"),
        ("Care ends at the gate", "No village-level visibility"),
        ("Outreach data siloed", "Districts lack live intel"),
        ("Guesswork procurement", "Seasonal stock-outs"),
        ("Late high-risk detection", "Preventable maternal & NCD harm"),
    ]
    for i, (title, desc) in enumerate(problems):
        top = Inches(1.3) + Inches(i * 1.08)
        card = add_round_rect(s, Inches(5.9), top, Inches(6.95), Inches(0.98), WHITE)
        card.line.color.rgb = LINE
        add_rect(s, Inches(5.9), top, Inches(0.14), Inches(0.98), ACCENT if i == 0 else TEAL)
        textbox(s, Inches(6.3), top + Inches(0.12), Inches(6.3), Inches(0.4),
                title, size=BODY, bold=True, color=NAVY)
        textbox(s, Inches(6.3), top + Inches(0.52), Inches(6.3), Inches(0.35),
                desc, size=BODY_SM, color=MUTED)
    footer(s, 2)

    # ========== 3. SOLUTION ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "The solution: FCIN / CHIP",
                 "From community medical practice to Africa's Community Health Intelligence Company.")
    textbox(s, Inches(0.45), Inches(1.25), Inches(6.4), Inches(1.0),
            "CHIP ingests community-generated health data, runs AI/ML and GIS analytics, "
            "and delivers predictive alerts to CHWs, facilities, districts, and partners — before crises escalate.",
            size=BODY, color=SLATE)
    pillars = [
        ("FCIN", "Venture brand connecting communities, providers, governments & partners"),
        ("CHIP", "Deep-tech platform: capture → intelligence → action"),
        ("FairBanks edge", "Live centre + CHW/VHT networks to pilot & validate in the field"),
    ]
    for i, (t, b) in enumerate(pillars):
        top = Inches(2.4) + Inches(i * 1.4)
        card = add_round_rect(s, Inches(0.45), top, Inches(6.4), Inches(1.25), WHITE)
        card.line.color.rgb = LINE
        add_rect(s, Inches(0.45), top, Inches(0.14), Inches(1.25), TEAL)
        textbox(s, Inches(0.85), top + Inches(0.18), Inches(5.75), Inches(0.4),
                t, size=H3, bold=True, color=TEAL)
        textbox(s, Inches(0.85), top + Inches(0.65), Inches(5.75), Inches(0.5),
                b, size=BODY_SM, color=SLATE)
    image_card(s, img("dashboard_demo.png"), Inches(7.1), Inches(1.25), Inches(5.75), Inches(5.55),
               "Facility & district intelligence dashboard (concept)")
    footer(s, 3)

    # ========== 4. DEEP TECH ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Deep technology core",
                 "Substantial engineering — not a lightweight booking app.")
    image_card(s, img("deep_tech_collage.png"), Inches(0.4), Inches(1.3), Inches(5.7), Inches(5.55),
               "AI · ML · GIS · mobile edge · cloud · NLP")
    techs = [
        ("Artificial Intelligence", "Outbreak early warning; maternal & NCD risk scoring"),
        ("Machine Learning", "Seasonal patterns from community & outreach data"),
        ("GIS Mapping", "Hotspots, disease distribution, resource gaps"),
        ("Mobile Data Collection", "Offline CHW/VHT apps at household level"),
        ("Cloud Computing", "Secure sync across facilities & partners"),
        ("Analytics + NLP", "Live dashboards; local-language symptom capture"),
    ]
    for i, (t, b) in enumerate(techs):
        top = Inches(1.3) + Inches(i * 0.9)
        card = add_round_rect(s, Inches(6.35), top, Inches(6.5), Inches(0.82), WHITE)
        card.line.color.rgb = LINE
        textbox(s, Inches(6.6), top + Inches(0.08), Inches(6.0), Inches(0.35),
                t, size=BODY, bold=True, color=TEAL)
        textbox(s, Inches(6.6), top + Inches(0.42), Inches(6.0), Inches(0.32),
                b, size=BODY_SM, color=MUTED)
    footer(s, 4)

    # ========== 5. HOW IT WORKS ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "How it works", "Last-mile signals become predictions — then action.")
    image_card(s, img("data_flow_iso.png"), Inches(0.35), Inches(1.25), Inches(7.2), Inches(5.6),
               "Community → Capture → CHIP intelligence → Decision-makers")
    layers = [
        ("1. Sources", "CHWs, clinics, pharmacies, schools, ANC/PNC, labs, HMIS"),
        ("2. Capture", "Offline mobile & facility apps · cloud sync"),
        ("3. CHIP core", "AI/ML · Predictive analytics · GIS · CDS"),
        ("4. Action", "Alerts · dashboards · referrals · stock · outreach"),
    ]
    for i, (t, b) in enumerate(layers):
        top = Inches(1.3) + Inches(i * 1.35)
        card = add_round_rect(s, Inches(7.8), top, Inches(5.05), Inches(1.2), WHITE)
        card.line.color.rgb = LINE
        add_rect(s, Inches(7.8), top, Inches(5.05), Inches(0.12), TEAL if i < 3 else ACCENT)
        textbox(s, Inches(8.05), top + Inches(0.25), Inches(4.55), Inches(0.35),
                t, size=BODY, bold=True, color=NAVY)
        textbox(s, Inches(8.05), top + Inches(0.62), Inches(4.55), Inches(0.45),
                b, size=BODY_SM, color=MUTED)
    footer(s, 5)

    # ========== 6. USE CASES ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Predictive use cases", "Proof-of-concept scenarios in FairBanks' live catchment.")
    image_card(s, img("maternal_visit.png"), Inches(0.35), Inches(1.25), Inches(4.15), Inches(2.55),
               "Maternal risk — home visits")
    image_card(s, img("gis_hotspots.png"), Inches(4.6), Inches(1.25), Inches(4.15), Inches(2.55),
               "Surveillance — GIS early warning")
    image_card(s, img("pharmacy_stock.png"), Inches(8.85), Inches(1.25), Inches(4.05), Inches(2.55),
               "Medicine demand forecasting")

    cases = [
        ("Disease surveillance", "Fever clusters → malaria risk in ~14 days → testing & bed-nets"),
        ("Maternal health", "BP / Hb / ANC gaps → high-risk flags → CHW referral"),
        ("NCDs", "Community BP & glucose → hotspots → targeted screening"),
        ("Child health", "Growth & immunisation → coverage gaps → drives"),
        ("Medicine demand", "Trends + seasonality → buy before stock-outs"),
    ]
    for i, (t, b) in enumerate(cases):
        left = Inches(0.35) + Inches(i * 2.58)
        card = add_round_rect(s, left, Inches(4.0), Inches(2.48), Inches(2.85), WHITE)
        card.line.color.rgb = LINE
        add_rect(s, left, Inches(4.0), Inches(2.48), Inches(0.12), TEAL if i % 2 == 0 else ACCENT)
        textbox(s, left + Inches(0.12), Inches(4.25), Inches(2.24), Inches(0.7),
                t, size=BODY_SM, bold=True, color=NAVY)
        textbox(s, left + Inches(0.12), Inches(5.05), Inches(2.24), Inches(1.55),
                b, size=CAPTION + 1, color=MUTED)
    footer(s, 6)

    # ========== 7. MARKET ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Market & customers", "B2B / B2G buyers across Africa's primary-care stack.")
    image_card(s, img("mobile_capture.png"), Inches(0.35), Inches(1.3), Inches(5.2), Inches(5.55),
               "CHW/VHT mobile capture — frontline product")
    customers = [
        ("Medical centres & clinics", "Outreach planning & population visibility"),
        ("District health offices", "Disease intelligence & early warning"),
        ("NGOs & partners", "Real-time M&E and impact evidence"),
        ("CHWs / VHTs", "Tools, workflows, decision support"),
        ("Ministries of health", "Planning & outbreak preparedness"),
        ("Insurers & research", "Prevention insights; ethical datasets"),
    ]
    for i, (t, b) in enumerate(customers):
        col = i % 2
        row = i // 2
        left = Inches(5.8) + Inches(col * 3.6)
        top = Inches(1.3) + Inches(row * 1.8)
        card = add_round_rect(s, left, top, Inches(3.45), Inches(1.65), WHITE)
        card.line.color.rgb = LINE
        textbox(s, left + Inches(0.18), top + Inches(0.25), Inches(3.1), Inches(0.55),
                t, size=BODY, bold=True, color=TEAL)
        textbox(s, left + Inches(0.18), top + Inches(0.9), Inches(3.1), Inches(0.55),
                b, size=BODY_SM, color=MUTED)
    footer(s, 7)

    # ========== 8. BUSINESS MODEL ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Business model", "Diversified HealthTech revenue — B2B, B2G, and partner channels.")
    image_card(s, img("dashboard_demo.png"), Inches(9.15), Inches(1.3), Inches(3.75), Inches(5.55),
               "SaaS dashboards & analytics")
    streams = [
        "Subscription licences for clinics and hospitals",
        "District health office deployments (SaaS + implementation)",
        "NGO programme monitoring contracts",
        "Ministry of Health national / sub-national implementations",
        "Custom analytics and reporting for partners",
        "Research collaborations with universities",
        "API integrations for digital health partners",
        "CHW training and certification on platform use",
    ]
    for i, item in enumerate(streams):
        top = Inches(1.3) + Inches(i * 0.68)
        card = add_round_rect(s, Inches(0.4), top, Inches(8.5), Inches(0.6), WHITE)
        card.line.color.rgb = LINE
        add_rect(s, Inches(0.4), top, Inches(0.12), Inches(0.6), TEAL)
        textbox(s, Inches(0.75), top + Inches(0.1), Inches(7.95), Inches(0.4),
                f"{i + 1}.  {item}", size=BODY_SM, bold=True, color=SLATE)
    footer(s, 8)

    # ========== 9. TRACTION ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Traction & unfair advantage",
                 "Not a concept detached from the field — a live operating foundation.")
    image_card(s, img("hero_chw_mobile.png"), Inches(0.35), Inches(1.25), Inches(5.4), Inches(3.2),
               "Active CHW/VHT engagement across Kampala communities")
    image_card(s, img("maternal_visit.png"), Inches(0.35), Inches(4.55), Inches(2.6), Inches(2.3),
               "Maternal & child")
    image_card(s, img("pharmacy_stock.png"), Inches(3.1), Inches(4.55), Inches(2.65), Inches(2.3),
               "Pharmacy & NCD")

    assets = [
        "Functioning FairBanks medical centre",
        "Community Reach Programme (live)",
        "CHW/VHT: Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya+",
        "Maternal & child health initiatives",
        "Gericare + chronic disease screening",
        "Corporate & school health programmes",
        "Digital health records foundation",
        "Research & community partnerships",
    ]
    right = add_round_rect(s, Inches(6.0), Inches(1.25), Inches(6.85), Inches(3.4), WHITE)
    right.line.color.rgb = LINE
    textbox(s, Inches(6.25), Inches(1.4), Inches(6.4), Inches(0.4),
            "Existing FairBanks ecosystem", size=H3, bold=True, color=TEAL)
    bullet_box(s, Inches(6.25), Inches(1.95), Inches(6.4), Inches(2.5), assets, size=BODY_SM)

    win = add_round_rect(s, Inches(6.0), Inches(4.8), Inches(6.85), Inches(2.05), NAVY)
    textbox(s, Inches(6.25), Inches(4.95), Inches(6.4), Inches(0.35),
            "Why FairBanks wins", size=BODY, bold=True, color=TEAL_LIGHT)
    textbox(s, Inches(6.25), Inches(5.4), Inches(6.4), Inches(1.3),
            "Technology + field access in one venture.\n"
            "Context-rooted AI from African community realities.\n"
            "Design → pilot → validate → refine on a live site.\n"
            "Recommended track: Startup.",
            size=BODY_SM, color=WHITE)
    footer(s, 9)

    # ========== 10. MVP & ROADMAP ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "MVP & roadmap", "Prove predictive intelligence in FairBanks' catchment first.")
    image_card(s, img("mobile_capture.png"), Inches(0.35), Inches(1.25), Inches(4.05), Inches(2.45),
               "MVP: CHW/VHT offline app")
    image_card(s, img("gis_hotspots.png"), Inches(4.55), Inches(1.25), Inches(4.05), Inches(2.45),
               "MVP: GIS risk layer v1")
    image_card(s, img("dashboard_demo.png"), Inches(8.75), Inches(1.25), Inches(4.15), Inches(2.45),
               "MVP: Facility analytics dashboard")

    mvp_bits = [
        "Cloud sync & validated data pipeline",
        "AI Module v1 — fever clusters, maternal flags, BP trends",
        "Pilot integration with FairBanks records & outreach",
    ]
    for i, t in enumerate(mvp_bits):
        left = Inches(0.35) + Inches(i * 4.3)
        card = add_round_rect(s, left, Inches(3.9), Inches(4.15), Inches(0.75), SOFT_TEAL)
        textbox(s, left + Inches(0.15), Inches(4.05), Inches(3.85), Inches(0.5),
                t, size=CAPTION + 1, bold=True, color=NAVY)

    phases = [
        ("Phase 1 · 0–12 mo", "Pilot MVP; validate 3 use cases"),
        ("Phase 2 · 12–24 mo", "District scale; NGO M&E; APIs"),
        ("Phase 3 · 24–36 mo", "Multi-district UG + East Africa"),
        ("Phase 4 · Year 3+", "CDS, local NLP, research modules"),
    ]
    for i, (t, b) in enumerate(phases):
        left = Inches(0.35) + Inches(i * 3.22)
        add_round_rect(s, left, Inches(4.85), Inches(3.1), Inches(2.0), NAVY)
        textbox(s, left + Inches(0.15), Inches(5.05), Inches(2.8), Inches(0.45),
                t, size=BODY_SM, bold=True, color=TEAL_LIGHT)
        textbox(s, left + Inches(0.15), Inches(5.6), Inches(2.8), Inches(0.95),
                b, size=BODY_SM, color=WHITE)
    footer(s, 10)

    # ========== 11. IMPACT ==========
    s = prs.slides.add_slide(blank)
    section_band(s, "Impact & SDG alignment",
                 "From sick-care to predictive, community-centred prevention.")
    image_card(s, img("maternal_visit.png"), Inches(0.35), Inches(1.3), Inches(5.2), Inches(5.55),
               "Earlier intervention at household & village level")
    outcomes = [
        "Earlier outbreak detection; lower communicable morbidity",
        "Fewer maternal & neonatal complications",
        "Reduced NCD burden via hotspot screening",
        "Improved child nutrition & immunisation coverage",
        "Fewer medicine stock-outs; smarter allocation",
        "Evidence base for NGOs, donors & governments",
    ]
    card = add_round_rect(s, Inches(5.8), Inches(1.3), Inches(7.05), Inches(3.45), WHITE)
    card.line.color.rgb = LINE
    textbox(s, Inches(6.1), Inches(1.45), Inches(6.5), Inches(0.4),
            "Expected outcomes", size=H3, bold=True, color=TEAL)
    bullet_box(s, Inches(6.1), Inches(2.0), Inches(6.5), Inches(2.5), outcomes, size=BODY_SM)

    sdgs = [("SDG 3", "Health"), ("SDG 5", "Gender"), ("SDG 9", "Innovation"),
            ("SDG 10", "Equality"), ("SDG 17", "Partnerships")]
    for i, (code, name) in enumerate(sdgs):
        left = Inches(5.8) + Inches(i * 1.42)
        add_round_rect(s, left, Inches(4.95), Inches(1.35), Inches(1.9), NAVY)
        textbox(s, left + Inches(0.05), Inches(5.3), Inches(1.25), Inches(0.4),
                code, size=BODY_SM, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        textbox(s, left + Inches(0.05), Inches(5.75), Inches(1.25), Inches(0.5),
                name, size=CAPTION + 1, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(s, 11)

    # ========== 12. ASK ==========
    s = prs.slides.add_slide(blank)
    add_picture_cover(s, img("gis_hotspots.png"), Inches(5.8), 0, Inches(7.533), SLIDE_H)
    add_rect(s, 0, 0, Inches(6.3), SLIDE_H, NAVY)
    add_rect(s, Inches(6.15), 0, Inches(0.15), SLIDE_H, TEAL_LIGHT)

    textbox(s, Inches(0.45), Inches(0.45), Inches(5.5), Inches(0.55),
            "The ask & next steps", size=H2, bold=True, color=WHITE)
    textbox(s, Inches(0.45), Inches(1.1), Inches(5.5), Inches(1.35),
            "Support to pilot the CHIP MVP — mobile CHW/VHT capture, cloud sync, "
            "predictive analytics, and facility/district dashboards — validated in "
            "FairBanks' live ecosystem before district and regional scale-up.",
            size=BODY, color=LIGHT)

    steps = [
        "Finalise MVP scope & AI priorities",
        "Submit AWIEF application + PDF deck",
        "Record 3-minute technology demo",
        "Launch CHW/VHT pilot in catchment",
    ]
    for i, step in enumerate(steps):
        top = Inches(2.65) + Inches(i * 0.7)
        add_round_rect(s, Inches(0.45), top, Inches(5.5), Inches(0.6), RGBColor(0x12, 0x32, 0x42))
        textbox(s, Inches(0.65), top + Inches(0.12), Inches(5.1), Inches(0.4),
                f"{i + 1}.  {step}", size=BODY_SM, bold=True, color=WHITE)

    textbox(s, Inches(0.45), Inches(5.4), Inches(5.5), Inches(0.35),
            "Your health, our mission.",
            size=BODY_SM, bold=True, color=ACCENT)
    textbox(s, Inches(0.45), Inches(5.8), Inches(5.5), Inches(0.4),
            "Vision: Africa's leading Community Health Intelligence Platform.",
            size=BODY_SM, bold=True, color=TEAL_LIGHT)
    textbox(s, Inches(0.45), Inches(6.25), Inches(5.5), Inches(0.3),
            "FairBanks Community Intelligence Network (FCIN) · CHIP",
            size=CAPTION + 1, color=LIGHT)
    textbox(s, Inches(0.45), Inches(6.55), Inches(5.5), Inches(0.25),
            "Contact: [founder · email · phone · web]",
            size=CAPTION + 1, color=RGBColor(0xB0, 0xC4, 0xC8))
    textbox(s, Inches(0.45), Inches(6.85), Inches(5.5), Inches(0.25),
            "awief.untap.us/pitch-n-grow2026  ·  Cape Town 10–11 Nov 2026",
            size=CAPTION, bold=True, color=TEAL_LIGHT)

    pitch_dir = Path(__file__).resolve().parent
    out = pitch_dir / "ppt_version.pptx"
    try:
        prs.save(str(out))
        print(out)
    except PermissionError:
        alt = out.with_name(out.stem + "_unlocked" + out.suffix)
        prs.save(str(alt))
        print(f"File locked; saved as: {alt}")
        return str(alt)
    return str(out)


if __name__ == "__main__":
    build()
