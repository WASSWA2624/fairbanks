"""
AWIEF Pitch n Grow 2026 — FairBanks / FCIN / CHIP investor pitch deck (12 slides).
Strictly capped at 10–12 slides per programme requirements.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- Brand palette (HealthTech: deep teal + navy; avoid purple AI cliche) ---
NAVY = RGBColor(0x0A, 0x1F, 0x2E)
TEAL = RGBColor(0x0D, 0x6E, 0x6E)
TEAL_LIGHT = RGBColor(0x14, 0xA3, 0xA3)
CREAM = RGBColor(0xF7, 0xF5, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SLATE = RGBColor(0x2C, 0x3E, 0x4A)
MUTED = RGBColor(0x5A, 0x6B, 0x75)
ACCENT = RGBColor(0xC4, 0x5C, 0x26)  # warm terracotta accent sparingly
CARD = RGBColor(0xEE, 0xF4, 0xF4)
LINE = RGBColor(0xD0, 0xDC, 0xDC)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


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
    # Soften corners
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def textbox(slide, left, top, width, height, text, size=18, bold=False, color=SLATE,
            align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def add_paragraph(tf, text, size=16, bold=False, color=SLATE, space_before=6, space_after=0,
                  align=PP_ALIGN.LEFT, font="Calibri"):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return p


def bullet_box(slide, left, top, width, height, items, size=15, color=SLATE, bullet="•"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4 if i else 0)
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = f"{bullet}  {item}"
        set_run(run, size=size, bold=False, color=color)
    return box


def footer(slide, page, total=12):
    add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.35), NAVY)
    textbox(slide, Inches(0.5), Inches(7.18), Inches(9), Inches(0.28),
            "AWIEF Pitch n Grow 2026  |  Deep Roots. Digital Futures.  |  HealthTech · AI/ML",
            size=10, color=RGBColor(0xA8, 0xC4, 0xC4), align=PP_ALIGN.LEFT)
    textbox(slide, Inches(11.2), Inches(7.18), Inches(1.6), Inches(0.28),
            f"{page} / {total}", size=10, color=RGBColor(0xA8, 0xC4, 0xC4), align=PP_ALIGN.RIGHT)


def section_header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), TEAL)
    add_rect(slide, 0, Inches(0.12), SLIDE_W, Inches(7.03), CREAM)
    textbox(slide, Inches(0.55), Inches(0.35), Inches(12), Inches(0.45),
            title, size=28, bold=True, color=NAVY, font="Calibri")
    if subtitle:
        textbox(slide, Inches(0.55), Inches(0.82), Inches(12), Inches(0.35),
                subtitle, size=14, color=MUTED)


def card(slide, left, top, width, height, title, body_lines, title_size=14, body_size=12):
    shape = add_round_rect(slide, left, top, width, height, WHITE)
    shape.line.color.rgb = LINE
    shape.line.width = Pt(1)
    textbox(slide, left + Inches(0.18), top + Inches(0.15), width - Inches(0.3), Inches(0.35),
            title, size=title_size, bold=True, color=TEAL)
    bullet_box(slide, left + Inches(0.15), top + Inches(0.5), width - Inches(0.3), height - Inches(0.6),
               body_lines, size=body_size, color=SLATE)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]  # blank

    # ========== 1. TITLE ==========
    s = prs.slides.add_slide(blank_layout)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, TEAL_LIGHT)
    textbox(s, Inches(0.7), Inches(1.1), Inches(12), Inches(0.35),
            "AWIEF Pitch n Grow 2026  ·  Startup Track  ·  HealthTech",
            size=13, color=TEAL_LIGHT)
    textbox(s, Inches(0.7), Inches(1.7), Inches(12), Inches(0.7),
            "FairBanks Community Intelligence Network",
            size=34, bold=True, color=WHITE, font="Calibri")
    textbox(s, Inches(0.7), Inches(2.4), Inches(12), Inches(0.4),
            "Community Health Intelligence Platform (CHIP)",
            size=22, color=TEAL_LIGHT)
    textbox(s, Inches(0.7), Inches(3.2), Inches(11.5), Inches(1.1),
            "AI-powered community health intelligence that turns last-mile data into "
            "predictions — so African primary care prevents crises instead of waiting for them.",
            size=18, color=RGBColor(0xD0, 0xDC, 0xE0))
    textbox(s, Inches(0.7), Inches(4.6), Inches(12), Inches(0.35),
            "Theme alignment: Deep Roots. Digital Futures.",
            size=14, bold=True, color=ACCENT)
    textbox(s, Inches(0.7), Inches(5.2), Inches(12), Inches(0.35),
            "Uganda pilot  →  East Africa  →  Pan-African scale",
            size=14, color=RGBColor(0xA8, 0xC4, 0xC4))
    textbox(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.3),
            "Woman-led deep-tech venture  |  Live medical centre + community outreach foundation",
            size=12, color=MUTED)

    # ========== 2. PROBLEM ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "The problem", "Primary healthcare in underserved communities is reactive — not intelligent.")
    problems = [
        ("Facilities wait for sickness", "Outbreaks and complications detected too late"),
        ("Care ends at the gate", "No continuous view of village-level health trends"),
        ("Outreach data stays siloed", "Districts and NGOs lack real-time intelligence"),
        ("Guesswork procurement", "Medicine stock-outs during seasonal surges"),
        ("Late high-risk detection", "Preventable maternal and NCD harm"),
    ]
    for i, (title, desc) in enumerate(problems):
        col = i % 3
        row = i // 3
        left = Inches(0.55) + Inches(col * 4.15)
        top = Inches(1.45) + Inches(row * 2.2)
        shape = add_round_rect(s, left, top, Inches(3.95), Inches(1.95), WHITE)
        shape.line.color.rgb = LINE
        add_rect(s, left, top, Inches(0.12), Inches(1.95), TEAL if i < 4 else ACCENT)
        textbox(s, left + Inches(0.3), top + Inches(0.35), Inches(3.4), Inches(0.5),
                title, size=16, bold=True, color=NAVY)
        textbox(s, left + Inches(0.3), top + Inches(0.95), Inches(3.4), Inches(0.7),
                desc, size=13, color=MUTED)
    # fifth card spans nicer - already placed
    footer(s, 2)

    # ========== 3. SOLUTION ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "The solution: FCIN / CHIP",
                   "From a community medical centre to Africa's Community Health Intelligence Company.")
    textbox(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.7),
            "CHIP continuously ingests community-generated health data, runs AI/ML and GIS analytics, "
            "and delivers predictive alerts to CHWs, facilities, districts, and partners — before crises escalate.",
            size=16, color=SLATE)

    pillars = [
        ("FCIN", "Venture brand connecting communities, providers, governments, and development partners"),
        ("CHIP", "Deep-tech platform: capture → intelligence → action across the health ecosystem"),
        ("FairBanks edge", "Live medical centre + CHW/VHT networks to design, pilot, and validate in the field"),
    ]
    for i, (t, b) in enumerate(pillars):
        left = Inches(0.55) + Inches(i * 4.15)
        shape = add_round_rect(s, left, Inches(2.3), Inches(3.95), Inches(2.6), WHITE)
        shape.line.color.rgb = LINE
        add_rect(s, left, Inches(2.3), Inches(3.95), Inches(0.12), TEAL)
        textbox(s, left + Inches(0.25), Inches(2.6), Inches(3.45), Inches(0.4),
                t, size=20, bold=True, color=TEAL)
        textbox(s, left + Inches(0.25), Inches(3.2), Inches(3.45), Inches(1.4),
                b, size=14, color=SLATE)

    textbox(s, Inches(0.55), Inches(5.3), Inches(12.2), Inches(0.9),
            "Positioning: Not \"another clinic app\" — a Community Health Intelligence Company "
            "rooted in African community realities (Deep Roots) and powered by deep technology (Digital Futures).",
            size=14, bold=True, color=NAVY)
    footer(s, 3)

    # ========== 4. DEEP TECH ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Deep technology core",
                   "AWIEF 2026 requires substantial engineering — not lightweight digital services.")
    techs = [
        ("Artificial Intelligence", "Disease-risk prediction, outbreak early warning, maternal & NCD risk scoring"),
        ("Machine Learning", "Patterns from community data, seasonal trends, outreach outcomes"),
        ("GIS Mapping", "Geospatial disease distribution, hotspots, resource gaps"),
        ("Mobile Data Collection", "Offline-capable CHW/VHT apps for household & community capture"),
        ("Cloud Computing", "Secure sync across facilities, partners, administrative levels"),
        ("Analytics + NLP", "Real-time dashboards; local-language symptom reporting where appropriate"),
    ]
    for i, (t, b) in enumerate(techs):
        col = i % 3
        row = i // 3
        left = Inches(0.55) + Inches(col * 4.15)
        top = Inches(1.4) + Inches(row * 2.45)
        shape = add_round_rect(s, left, top, Inches(3.95), Inches(2.2), WHITE)
        shape.line.color.rgb = LINE
        textbox(s, left + Inches(0.25), top + Inches(0.3), Inches(3.45), Inches(0.45),
                t, size=15, bold=True, color=TEAL)
        textbox(s, left + Inches(0.25), top + Inches(0.9), Inches(3.45), Inches(1.0),
                b, size=13, color=SLATE)
    footer(s, 4)

    # ========== 5. HOW IT WORKS ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "How it works", "Data from the last mile → intelligence → action in the community.")

    layers = [
        ("1. Sources", "CHWs/VHTs, patients, clinics, pharmacies, schools/faith outreach, ANC/PNC, labs, corporate wellness, HMIS/DHIS2, seasonality signals"),
        ("2. Capture", "Offline mobile & facility apps · structured forms · secure cloud sync"),
        ("3. CHIP core", "AI/ML engine · Predictive analytics · GIS maps · Clinical decision support"),
        ("4. Action", "CHW alerts · Facility dashboards · District / NGO / partner intelligence → referrals, outreach, stock, surveillance"),
    ]
    for i, (t, b) in enumerate(layers):
        top = Inches(1.35) + Inches(i * 1.25)
        add_round_rect(s, Inches(0.55), top, Inches(12.2), Inches(1.1), WHITE).line.color.rgb = LINE
        add_rect(s, Inches(0.55), top, Inches(2.4), Inches(1.1), TEAL if i < 3 else ACCENT)
        textbox(s, Inches(0.7), top + Inches(0.35), Inches(2.1), Inches(0.45),
                t, size=16, bold=True, color=WHITE)
        textbox(s, Inches(3.2), top + Inches(0.25), Inches(9.2), Inches(0.7),
                b, size=14, color=SLATE)
    footer(s, 5)

    # ========== 6. USE CASES ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Predictive use cases", "Proof-of-concept scenarios validated in FairBanks' live catchment.")
    cases = [
        ("Disease surveillance", "Fever clusters via VHTs → malaria outbreak risk in ~14 days → targeted testing & bed-nets"),
        ("Maternal health", "Home-visit BP/Hb/ANC adherence → high-risk pregnancy flags → CHW referral before crisis"),
        ("NCDs", "Community & workplace BP/glucose → hypertension/diabetes hotspots → targeted screening"),
        ("Child health", "Growth, immunisation, diarrhoea → malnutrition & coverage gaps → nutrition & immunisation drives"),
        ("Medicine demand", "Trends + seasonality + history → facility demand forecast → procurement before stock-outs"),
    ]
    for i, (t, b) in enumerate(cases):
        top = Inches(1.3) + Inches(i * 1.05)
        add_rect(s, Inches(0.55), top, Inches(0.12), Inches(0.9), TEAL if i % 2 == 0 else ACCENT)
        textbox(s, Inches(0.9), top + Inches(0.1), Inches(3.2), Inches(0.7),
                t, size=15, bold=True, color=NAVY)
        textbox(s, Inches(4.2), top + Inches(0.1), Inches(8.5), Inches(0.7),
                b, size=13, color=SLATE)
    footer(s, 6)

    # ========== 7. MARKET ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Market & customers", "B2B / B2G health intelligence buyers across the African primary-care stack.")
    customers = [
        ("Medical centres & clinics", "Follow-up, outreach planning, population health visibility"),
        ("District health offices", "Disease intelligence and early warning"),
        ("NGOs & partners", "Real-time M&E, impact evidence, programme optimisation"),
        ("CHWs / VHTs", "Mobile tools, structured workflows, decision support"),
        ("Ministries of health", "Sub-national planning and outbreak preparedness"),
        ("Insurers & research", "Prevention insights; ethical anonymised research datasets"),
    ]
    for i, (t, b) in enumerate(customers):
        col = i % 3
        row = i // 3
        left = Inches(0.55) + Inches(col * 4.15)
        top = Inches(1.4) + Inches(row * 2.35)
        shape = add_round_rect(s, left, top, Inches(3.95), Inches(2.1), WHITE)
        shape.line.color.rgb = LINE
        textbox(s, left + Inches(0.25), top + Inches(0.35), Inches(3.45), Inches(0.5),
                t, size=15, bold=True, color=TEAL)
        textbox(s, left + Inches(0.25), top + Inches(1.0), Inches(3.45), Inches(0.8),
                b, size=13, color=SLATE)
    textbox(s, Inches(0.55), Inches(6.35), Inches(12.2), Inches(0.4),
            "Phase 1: Uganda (FairBanks catchment + Kampala metro)  ·  Phase 2–3: East Africa → continental scale",
            size=13, bold=True, color=NAVY)
    footer(s, 7)

    # ========== 8. BUSINESS MODEL ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Business model", "Diversified HealthTech revenue — B2B, B2G, and partner channels.")
    streams = [
        "Subscription licences for clinics and hospitals",
        "District health office deployments (SaaS + implementation)",
        "NGO programme monitoring contracts",
        "Ministry of Health national / sub-national implementations",
        "Custom analytics and reporting for partners",
        "Research collaborations with universities",
        "API integrations for digital health ecosystem partners",
        "CHW training and certification on platform use",
    ]
    for i, item in enumerate(streams):
        col = i % 2
        row = i // 2
        left = Inches(0.55) + Inches(col * 6.3)
        top = Inches(1.4) + Inches(row * 1.15)
        shape = add_round_rect(s, left, top, Inches(6.05), Inches(1.0), WHITE)
        shape.line.color.rgb = LINE
        add_rect(s, left, top, Inches(0.12), Inches(1.0), TEAL)
        textbox(s, left + Inches(0.35), top + Inches(0.3), Inches(5.4), Inches(0.45),
                item, size=14, color=SLATE)
    footer(s, 8)

    # ========== 9. TRACTION ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Traction & unfair advantage",
                   "Not a concept detached from the field — a live operating foundation.")
    assets = [
        "Functioning FairBanks medical centre",
        "Community Reach Programme (live)",
        "CHW/VHT networks: Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya+",
        "Maternal & child health initiatives",
        "Gericare (geriatric care)",
        "Chronic disease screening",
        "Corporate & school health programmes",
        "Digital health records foundation",
        "Research & community partnerships",
    ]
    left_box = add_round_rect(s, Inches(0.55), Inches(1.35), Inches(6.5), Inches(5.3), WHITE)
    left_box.line.color.rgb = LINE
    textbox(s, Inches(0.8), Inches(1.55), Inches(6), Inches(0.4),
            "Existing FairBanks ecosystem", size=16, bold=True, color=TEAL)
    bullet_box(s, Inches(0.8), Inches(2.1), Inches(5.9), Inches(4.3), assets, size=13)

    right = add_round_rect(s, Inches(7.3), Inches(1.35), Inches(5.45), Inches(5.3), NAVY)
    textbox(s, Inches(7.55), Inches(1.6), Inches(5), Inches(0.4),
            "Why FairBanks wins", size=16, bold=True, color=TEAL_LIGHT)
    wins = [
        "Technology + field access in one venture",
        "Context-rooted AI from African community realities",
        "Design → pilot → validate → refine on a live site",
        "De-risks product-market fit before district/national scale",
        "Recommended track: Startup (operating centre + deep-tech build)",
    ]
    for i, w in enumerate(wins):
        textbox(s, Inches(7.55), Inches(2.2) + Inches(i * 0.75), Inches(5), Inches(0.65),
                f"→  {w}", size=13, color=WHITE)
    footer(s, 9)

    # ========== 10. MVP & ROADMAP ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "MVP & roadmap", "Prove predictive community health intelligence in FairBanks' catchment first.")

    textbox(s, Inches(0.55), Inches(1.3), Inches(12), Inches(0.35),
            "MVP components", size=15, bold=True, color=TEAL)
    mvp = [
        ("CHW/VHT mobile app", "Offline capture: symptoms, vitals, maternal/child, household visits"),
        ("Cloud pipeline", "Secure ingestion, validation, storage"),
        ("Facility dashboard", "Trends, alerts, outreach priorities"),
        ("AI Module v1", "Rule-based + ML risk scoring (fever clusters, maternal flags, BP trends)"),
        ("GIS layer v1", "Cases & risk zones by village/parish"),
        ("Pilot integration", "Link FairBanks records & outreach workflows"),
    ]
    for i, (t, b) in enumerate(mvp):
        col = i % 3
        row = i // 3
        left = Inches(0.55) + Inches(col * 4.15)
        top = Inches(1.7) + Inches(row * 1.35)
        shape = add_round_rect(s, left, top, Inches(3.95), Inches(1.2), WHITE)
        shape.line.color.rgb = LINE
        textbox(s, left + Inches(0.2), top + Inches(0.15), Inches(3.55), Inches(0.3),
                t, size=12, bold=True, color=NAVY)
        textbox(s, left + Inches(0.2), top + Inches(0.5), Inches(3.55), Inches(0.55),
                b, size=11, color=MUTED)

    textbox(s, Inches(0.55), Inches(4.55), Inches(12), Inches(0.3),
            "Roadmap", size=15, bold=True, color=TEAL)
    phases = [
        ("Phase 1 · 0–12 mo", "Pilot MVP with FairBanks CHWs; validate 3 use cases"),
        ("Phase 2 · 12–24 mo", "District scale; NGO M&E; partner APIs"),
        ("Phase 3 · 24–36 mo", "Multi-district UG + East Africa; advanced ML & demand forecasting"),
        ("Phase 4 · Year 3+", "CDS, local-language NLP, research modules"),
    ]
    for i, (t, b) in enumerate(phases):
        left = Inches(0.55) + Inches(i * 3.15)
        add_round_rect(s, left, Inches(4.95), Inches(3.0), Inches(1.55), NAVY)
        textbox(s, left + Inches(0.15), Inches(5.1), Inches(2.7), Inches(0.4),
                t, size=12, bold=True, color=TEAL_LIGHT)
        textbox(s, left + Inches(0.15), Inches(5.55), Inches(2.7), Inches(0.8),
                b, size=11, color=WHITE)
    footer(s, 10)

    # ========== 11. IMPACT ==========
    s = prs.slides.add_slide(blank_layout)
    section_header(s, "Impact & SDG alignment",
                   "From sick-care to predictive, community-centred prevention.")
    outcomes = [
        "Earlier outbreak detection; lower communicable-disease morbidity",
        "Fewer maternal and neonatal complications via risk-based alerting",
        "Reduced NCD burden through hotspot-targeted screening",
        "Improved child nutrition and immunisation coverage",
        "Fewer medicine stock-outs; smarter public-health resource allocation",
        "Evidence base for NGOs, donors, and governments",
    ]
    left = add_round_rect(s, Inches(0.55), Inches(1.35), Inches(7.0), Inches(5.3), WHITE)
    left.line.color.rgb = LINE
    textbox(s, Inches(0.8), Inches(1.55), Inches(6.5), Inches(0.4),
            "Expected outcomes", size=16, bold=True, color=TEAL)
    bullet_box(s, Inches(0.8), Inches(2.15), Inches(6.4), Inches(4.2), outcomes, size=14)

    sdgs = [
        ("SDG 3", "Good Health & Well-Being"),
        ("SDG 5", "Gender Equality (women-led; maternal focus)"),
        ("SDG 9", "Industry, Innovation & Infrastructure"),
        ("SDG 10", "Reduced Inequalities"),
        ("SDG 17", "Partnerships for the Goals"),
    ]
    right = add_round_rect(s, Inches(7.8), Inches(1.35), Inches(4.95), Inches(5.3), NAVY)
    textbox(s, Inches(8.05), Inches(1.55), Inches(4.5), Inches(0.4),
            "SDG alignment", size=16, bold=True, color=TEAL_LIGHT)
    for i, (code, name) in enumerate(sdgs):
        textbox(s, Inches(8.05), Inches(2.2) + Inches(i * 0.75), Inches(4.5), Inches(0.35),
                code, size=13, bold=True, color=ACCENT)
        textbox(s, Inches(8.05), Inches(2.5) + Inches(i * 0.75), Inches(4.5), Inches(0.35),
                name, size=12, color=WHITE)
    footer(s, 11)

    # ========== 12. ASK ==========
    s = prs.slides.add_slide(blank_layout)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, TEAL_LIGHT)
    textbox(s, Inches(0.7), Inches(0.55), Inches(12), Inches(0.4),
            "The ask & next steps", size=28, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(1.15), Inches(12), Inches(0.7),
            "Support to develop and pilot the CHIP MVP — mobile CHW/VHT capture, cloud sync, "
            "predictive analytics, and facility/district dashboards — validated in FairBanks' live ecosystem "
            "before district and regional scale-up.",
            size=16, color=RGBColor(0xD0, 0xDC, 0xE0))

    next_steps = [
        "Finalise MVP scope & AI priorities (surveillance, maternal risk, NCD hotspots)",
        "Submit AWIEF application + this 12-slide deck (export PDF for portal)",
        "Record 3-minute technology demo video",
        "Launch FairBanks pilot with CHW/VHT cohort in primary catchment communities",
    ]
    for i, step in enumerate(next_steps):
        top = Inches(2.15) + Inches(i * 0.7)
        add_round_rect(s, Inches(0.7), top, Inches(11.8), Inches(0.6), RGBColor(0x12, 0x32, 0x42))
        textbox(s, Inches(0.95), top + Inches(0.12), Inches(11.3), Inches(0.4),
                f"{i + 1}.  {step}", size=14, color=WHITE)

    textbox(s, Inches(0.7), Inches(5.15), Inches(12), Inches(0.35),
            "Vision: Africa's leading Community Health Intelligence Platform.",
            size=16, bold=True, color=TEAL_LIGHT)
    textbox(s, Inches(0.7), Inches(5.7), Inches(12), Inches(0.35),
            "FairBanks Community Intelligence Network (FCIN)  ·  CHIP",
            size=14, color=RGBColor(0xA8, 0xC4, 0xC4))
    textbox(s, Inches(0.7), Inches(6.3), Inches(12), Inches(0.35),
            "Contact: [add founder name · email · phone · website]   |   Finals: Cape Town, 10–11 Nov 2026",
            size=13, color=MUTED)
    textbox(s, Inches(0.7), Inches(6.85), Inches(12), Inches(0.3),
            "Apply: https://awief.untap.us/pitch-n-grow2026",
            size=12, color=TEAL_LIGHT)

    out = r"d:\coding\apps\flutter\fairbanks\output\pitch\FairBanks_FCIN_CHIP_AWIEF_Pitch_n_Grow_2026.pptx"
    prs.save(out)
    print(out)
    return out


if __name__ == "__main__":
    build()
