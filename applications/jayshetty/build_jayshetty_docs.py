#!/usr/bin/env python3
"""Build the Jay Shetty strategic partnership pack for FairBanks.

Creates one synchronized set:
  documents/jayshetty_word.docx
  documents/jayshetty_pdf.pdf
  documents/jayshetty_ppt.pptx

Narrative: strategic partnership (not a donation ask). Three pillars + FCHIP.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ASSETS = REPO / "assets"
OUT = HERE / "documents"
DOCX = OUT / "jayshetty_word.docx"
PDF = OUT / "jayshetty_pdf.pdf"
PPTX = OUT / "jayshetty_ppt.pptx"
TMP = REPO / "tmp" / "jayshetty"

NAVY = "0A1F2E"
TEAL = "0D6E6E"
GREEN = "2D7A55"
ORANGE = "C45C26"
GOLD = "D99A2B"
CREAM = "F7F5F0"
PALE_TEAL = "E8F3F2"
PALE_ORANGE = "FBEDE6"
PALE_GREEN = "E9F2EC"
SLATE = "1E2F38"
MUTED = "52636C"
LINE = "CED9D8"
WHITE = "FFFFFF"

SLOGAN = "Your health, our mission."
PROGRAMME = "Strategic Partnership Brief"
TITLE = "FairBanks x Jay Shetty Ecosystem"
SUBTITLE = "From inspiration to measurable community health outcomes"
ORG = "FairBanks Medical Centre & Pharmacy"
CONTACT_NAME = "Racheal Nabukeera Sekagiri"
CONTACT_TITLE = "Managing Director"
WEBSITE = "https://fairbanksmedicalcentre.org"
EMAIL = "info@fairbanksmedicalcentre.org"
PHONE = "+256 777 462 398"
LOCATION = "Kampala, Uganda"

PHOTOS = {
    "cover": "outreach_facilitator_canopy_01.jpg",
    "logo": "fairbanks_logo.jpeg",
    "facility": "facility_exterior_entrance_01.jpg",
    "facility_branded": "facility_exterior_branded_entrance_01.jpeg",
    "facility_street": "facility_exterior_street_view_01.jpg",
    "pharmacy": "pharmacy_storefront_01.jpeg",
    "pharmacy_alt": "pharmacy_exterior_01.jpg",
    "outreach": "outreach_bp_screening.jpeg",
    "outreach_camp": "outreach_medical_camp_01.jpg",
    "community": "outreach_audience_full_group_01.jpg",
    "audience": "outreach_audience_attentive_01.jpg",
    "mobile": "outreach_mobile_phone_demo_01.jpg",
    "dashboard": "dashboard_demo.png",
    "architecture": "data_flow_iso_labeled.png",
    "gis": "gis_hotspots.png",
    "maternal": "bloom_maternal_health_participant_01.jpg",
    "gericare": "gericare_wheelchair_assist_02.jpg",
    "doctor_hands": "clinic_doctor_patient_hands_01.jpg",
    "lab": "clinic_lab_fingerprick_01.jpg",
    "compassion": "clinic_consult_compassion_01.jpg",
    "reception": "reception_wheelchair_checkin_01.jpg",
    "reception_staff": "reception_staff_documents_01.jpg",
    "team": "staff_team_reception.jpeg",
    "training": "indoor_training_audience_02.jpg",
    "mission": "reception_mission_wall.jpeg",
    "mothers": "waiting_room_mothers_01.jpeg",
    "canopy": "outreach_facilitator_group_01.jpg",
    "close": "cover_hero_cinematic.jpg",
}

CONCEPT = REPO / ".cursor" / "concept_improved.jpeg"

EMAIL_BODY = f"""Subject: Partnership Pitch: FairBanks Community Health x Jay Shetty Ecosystem

Dear Partnerships Team,

I hope this message finds you well.

I am writing on behalf of {ORG}, a community health social enterprise in {LOCATION}. We are building a scalable model that turns compassion and wellbeing into practical primary care, community outreach, and intelligent prevention for underserved families.

The Opportunity
At FairBanks we integrate three pillars: our medical centre and pharmacy, FairBanks Community Reach (outreach, education, prevention, CHIS, and livelihoods), and FCHIP - the FairBanks Community Health Intelligence Platform. FCHIP is our AI-enabled layer that helps communities and care teams see risk earlier and act before crises escalate. We already operate a live clinical and community environment in Kampala peri-urban communities, and we are validating a working FCHIP MVP for the next phase of scale.

Why Jay Shetty?
Jay Shetty has consistently championed wellbeing, compassion, purposeful living, and stronger communities. FairBanks shares that commitment - and converts it into tangible healthcare access, preventive education, and community health intelligence on the ground.

The Partnership
We would welcome a strategic partnership with the Jay Shetty ecosystem in areas such as:
- Community health and wellness initiatives
- Preventive healthcare and mental wellbeing programmes
- AI for social impact and digital health innovation
- Global storytelling and advocacy for equitable healthcare access
- Social enterprise growth support and introductions to aligned impact partners

What We Bring
- An operational community-centred medical facility serving as a real-world innovation hub
- An active Community Reach cascade with CHWs/VHTs, school and community programmes, and livelihood pathways including CHIS
- A working FCHIP MVP ready for structured validation and scale
- Deep local relationships and a leadership team committed to measurable, sustainable impact

I have attached a concise executive partnership brief and deck outlining our vision, model, innovation roadmap, and partnership options.

We would be honoured to schedule a brief 10-minute introductory conversation with your partnerships or impact team to explore how we might work together to improve community health and wellbeing.

Thank you for your time and consideration. We deeply appreciate the meaningful work Jay Shetty and his team continue to do in inspiring positive change around the world.

Warm regards,

{CONTACT_NAME}
{CONTACT_TITLE}
{ORG}
{LOCATION}

Website: {WEBSITE}
Email: {EMAIL}
Tel: {PHONE}
"""


def photo(key: str) -> Path:
    path = ASSETS / PHOTOS[key]
    if not path.exists():
        raise FileNotFoundError(path)
    return path


def _pil_size(path: Path):
    from PIL import Image as PILImage

    with PILImage.open(path) as im:
        return im.size


def fit_width(path: Path, max_width_in: float, max_height_in: float = 3.4) -> float:
    iw, ih = _pil_size(path)
    aspect = ih / float(iw)
    w = max_width_in
    h = w * aspect
    if h > max_height_in:
        w = max_height_in / aspect
    return w


def build_docx() -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Inches, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = Inches(0.6)
    sec.bottom_margin = Inches(0.7)
    sec.left_margin = sec.right_margin = Inches(0.7)

    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(11)
    styles["Normal"].font.color.rgb = RGBColor.from_string(SLATE)
    styles["Normal"].paragraph_format.space_after = Pt(8)
    styles["Normal"].paragraph_format.line_spacing = 1.15

    for name, size, color in (
        ("Title", 26, NAVY),
        ("Heading 1", 16, NAVY),
        ("Heading 2", 13, TEAL),
    ):
        style = styles[name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)

    def shade_cell(cell, hex_color: str) -> None:
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), hex_color)
        shd.set(qn("w:val"), "clear")
        tc_pr.append(shd)

    def add_para(text, *, style=None, bold=False, size=11, color=SLATE, align=None, space_after=8):
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
        run.font.name = "Calibri"
        return p

    def add_bullets(items):
        for item in items:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(item)
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor.from_string(SLATE)
            run.font.name = "Calibri"

    def set_cell_border(cell, color=LINE, sz="4") -> None:
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="nil"/>'
            f'<w:left w:val="nil"/>'
            f'<w:bottom w:val="nil"/>'
            f'<w:right w:val="nil"/>'
            f"</w:tcBorders>"
        )
        tc_pr.append(borders)

    def add_caption(text: str) -> None:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(10)
        run = p.add_run(text)
        run.italic = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(MUTED)
        run.font.name = "Calibri"

    def add_image(path: Path, width: float = 6.6, height_cap: float = 3.2, caption: str = ""):
        w = fit_width(path, width, height_cap)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2 if caption else 10)
        p.add_run().add_picture(str(path), width=Inches(w))
        if caption:
            add_caption(caption)

    def add_photo_row(
        left_key: str,
        right_key: str,
        left_caption: str,
        right_caption: str,
        *,
        width: float = 3.2,
        height_cap: float = 2.35,
    ) -> None:
        """Two related photos side by side, each with its own caption."""
        left_path, right_path = photo(left_key), photo(right_key)
        table = doc.add_table(rows=2, cols=2)
        table.autofit = True
        for col, (path, cap) in enumerate(
            [(left_path, left_caption), (right_path, right_caption)]
        ):
            img_cell = table.rows[0].cells[col]
            cap_cell = table.rows[1].cells[col]
            set_cell_border(img_cell)
            set_cell_border(cap_cell)
            img_cell.text = ""
            p = img_cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            w = fit_width(path, width, height_cap)
            p.add_run().add_picture(str(path), width=Inches(w))
            cap_cell.text = ""
            cp = cap_cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(2)
            cp.paragraph_format.space_after = Pt(6)
            run = cp.add_run(cap)
            run.italic = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string(MUTED)
            run.font.name = "Calibri"
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(6)

    # ---- Cover ----
    add_para(PROGRAMME.upper(), bold=True, size=11, color=ORANGE, space_after=4)
    add_para(TITLE, style="Title", bold=True, size=26, color=NAVY, space_after=6)
    add_para(SUBTITLE, size=13, color=MUTED, space_after=4)
    add_para(SLOGAN, bold=True, size=12, color=TEAL, space_after=10)
    add_photo_row(
        "cover",
        "facility_branded",
        "Community Reach under the canopy - education and care close to home",
        "FairBanks Medical Centre - Your health, our mission.",
    )

    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    facts = [
        ("Organisation", ORG),
        ("Location", LOCATION),
        ("Contact", f"{CONTACT_NAME}, {CONTACT_TITLE}"),
        ("Ask type", "Strategic partnership / discovery call (10 minutes)"),
    ]
    for i, (k, v) in enumerate(facts):
        table.rows[i].cells[0].text = k
        table.rows[i].cells[1].text = v
        shade_cell(table.rows[i].cells[0], PALE_TEAL)
        for cell in table.rows[i].cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.font.name = "Calibri"
                    r.font.color.rgb = RGBColor.from_string(SLATE)

    doc.add_paragraph()
    add_para(
        "This brief invites the Jay Shetty ecosystem to explore a strategic collaboration "
        "with FairBanks - a Ugandan community health enterprise combining compassionate care, "
        "Community Reach, and FCHIP community health intelligence.",
        size=11,
        color=SLATE,
    )

    # ---- Vision ----
    add_para("1. Building the future of community health", style="Heading 1", bold=True, size=16, color=NAVY)
    add_photo_row(
        "facility_street",
        "pharmacy",
        "Clinical anchor: FairBanks Medical Centre in Kampala",
        "Pharmacy storefront - affordable medicines beside the clinic",
    )
    add_para(
        "FairBanks is a community health social enterprise dedicated to transforming family "
        "and community health through accessible primary care, preventive medicine, community "
        "outreach, and technology-driven innovation. We believe quality healthcare should not "
        "depend on geography or income.",
    )
    add_para(
        "Our medical centre is the real-world laboratory where clinical care, Community Reach, "
        "and digital innovation come together - so solutions are designed, tested, and refined "
        "with real families before wider deployment.",
    )

    add_para("Three complementary pillars", style="Heading 2", bold=True, size=13, color=TEAL)
    pillars = [
        (
            "FairBanks Medical Centre & Pharmacy",
            "Quality family and community primary healthcare, diagnostics, pharmacy, and referrals.",
        ),
        (
            "FairBanks Community Reach",
            "Outreach, health education, screening, maternal and child health, Gericare, school health, CHIS, and livelihoods.",
        ),
        (
            "FCHIP (Community Health Intelligence Platform)",
            "AI-enabled intelligence that fuses community, facility, GIS, and climate signals to support earlier, smarter action.",
        ),
    ]
    for title, body in pillars:
        add_para(title, bold=True, size=11, color=NAVY, space_after=2)
        add_para(body, size=11, color=SLATE, space_after=8)

    add_photo_row(
        "reception",
        "gericare",
        "Reception and triage - dignity at every first contact",
        "Gericare in action - compassionate support for older patients",
    )

    # ---- How we work ----
    add_para("2. How FairBanks works - Community Reach", style="Heading 1", bold=True, size=16, color=NAVY)
    add_para(
        "FairBanks operates a six-step cascade with a continuous feedback loop. FCHIP sits on "
        "the Data & Feedback layer - the digital nervous system that helps the model learn and improve.",
    )
    if CONCEPT.exists():
        add_image(
            CONCEPT,
            5.0,
            4.8,
            caption="Community Reach cascade - communities, CHWs/VHTs, programmes, medical centre, research, empowerment",
        )
    add_bullets(
        [
            "Community members identify needs and own solutions",
            "CHWs / VHTs bridge homes, schools, and care",
            "Community Reach programmes deliver education, screening, and home visits",
            "FairBanks Medical Centre anchors clinical care and quality",
            "Research, partnerships, and skills strengthen evidence and capacity",
            "Economic empowerment (including CHIS and IGAs) supports resilient families",
        ]
    )
    add_photo_row(
        "audience",
        "community",
        "Community members listening at an outreach session",
        "Full Community Reach gathering - participation at the centre",
    )

    # ---- Problem / FCHIP ----
    add_para("3. The gap - and the FCHIP answer", style="Heading 1", bold=True, size=16, color=NAVY)
    add_photo_row(
        "outreach",
        "outreach_camp",
        "Blood-pressure screening during community outreach",
        "Medical camp consultations - care beyond clinic walls",
    )
    add_para(
        "Primary healthcare in underserved communities is still largely reactive. Facilities "
        "often see people only after illness appears. Data from CHWs, schools, outreach, clinics, "
        "and pharmacies stays fragmented. Climate signals rarely join the picture. Clinical records "
        "often remain locked inside facility EMR/HMS systems - so the continuous feedback loop "
        "cannot learn fast enough to prevent outbreaks, maternal complications, or stock-outs.",
    )
    add_para(
        "FCHIP is FairBanks' deep-technology component. It connects community members, CHWs/VHTs, "
        "schools, communities, hospitals, and clinics; safely exposes data APIs to existing EMR/HMS "
        "systems for real-time clinical ingest; and fuses signals with GIS mapping and climate APIs "
        "to surface explainable early warnings and dashboards.",
        space_after=8,
    )
    add_image(
        photo("architecture"),
        6.4,
        2.9,
        caption="FCHIP flow: community and facility capture - intelligence - action for facilities and districts",
    )
    add_para("Deep technology core", style="Heading 2", bold=True, size=13, color=TEAL)
    add_bullets(
        [
            "Artificial intelligence and machine learning for risk scoring and early warning",
            "GIS mapping of disease and resource geography",
            "Climate API integration for climate-sensitive health risk",
            "Secure EMR/HMS data APIs (no rip-and-replace for clinics)",
            "Offline-capable mobile capture for CHWs/VHTs",
            "Cloud sync and analytics dashboards for facilities and partners",
        ]
    )
    add_photo_row(
        "mobile",
        "dashboard",
        "Mobile tools for last-mile capture with CHWs and VHTs",
        "Facility and programme dashboards for timely decisions",
    )

    # ---- Traction / roadmap ----
    add_para("4. Traction and innovation roadmap", style="Heading 1", bold=True, size=16, color=NAVY)
    add_photo_row(
        "training",
        "maternal",
        "Community health education and training in progress",
        "Maternal and family health - earlier support for mothers",
    )
    add_para("What is already live", style="Heading 2", bold=True, size=13, color=TEAL)
    add_bullets(
        [
            "Operational FairBanks Medical Centre & Pharmacy in Kampala",
            "Active Community Reach across Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya, and surrounding communities",
            "CHW/VHT engagement, maternal and child health, Gericare, chronic-disease screening, school and corporate health",
            "CHIS and livelihood pathways linked to affordable access",
            "Working FCHIP MVP ready for structured field validation and evidence building",
        ]
    )
    add_photo_row(
        "doctor_hands",
        "lab",
        "Compassionate clinical care - listening before treating",
        "Point-of-care testing that keeps diagnosis close to the patient",
    )

    add_para("Innovation roadmap", style="Heading 2", bold=True, size=13, color=TEAL)
    roadmap = doc.add_table(rows=7, cols=2)
    roadmap.style = "Table Grid"
    rows = [
        ("Stage", "Status"),
        ("Community healthcare model established", "Completed"),
        ("FairBanks Medical Centre operational", "Completed"),
        ("Community Reach platform operational", "Completed"),
        ("FCHIP working MVP", "Completed - validating"),
        ("Structured pilot evidence & partner onboarding", "Next phase"),
        ("District and regional expansion", "Future vision"),
    ]
    for i, (a, b) in enumerate(rows):
        roadmap.rows[i].cells[0].text = a
        roadmap.rows[i].cells[1].text = b
        if i == 0:
            shade_cell(roadmap.rows[i].cells[0], TEAL)
            shade_cell(roadmap.rows[i].cells[1], TEAL)
            for cell in roadmap.rows[i].cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.bold = True
                        r.font.color.rgb = RGBColor.from_string(WHITE)
                        r.font.size = Pt(10)
        else:
            for cell in roadmap.rows[i].cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(10)
                        r.font.name = "Calibri"
                        r.font.color.rgb = RGBColor.from_string(SLATE)

    doc.add_paragraph()
    add_para(
        "Long-term ambition: build one of Africa's leading community health intelligence "
        "ecosystems - combining primary care delivery with intelligent digital technologies "
        "that empower communities, providers, and partners.",
    )

    # ---- Partnership ----
    add_para("5. Partnership opportunity", style="Heading 1", bold=True, size=16, color=NAVY)
    add_photo_row(
        "compassion",
        "mothers",
        "Clinical compassion - the human face of FairBanks care",
        "Mothers and families waiting with trust in the clinic",
    )
    add_para(
        "Jay Shetty inspires millions to live healthier, more purposeful lives through "
        "compassion, mindfulness, and service. FairBanks converts that inspiration into "
        "measurable community health outcomes - care close to home, earlier prevention, "
        "and intelligent tools for last-mile Africa.",
    )
    add_para("Potential collaboration areas", style="Heading 2", bold=True, size=13, color=TEAL)
    add_bullets(
        [
            "Community health and wellness programmes",
            "Preventive healthcare and health promotion",
            "Mental health and holistic wellbeing initiatives",
            "Artificial intelligence for social impact",
            "Digital health innovation and research exchange",
            "Global advocacy and storytelling",
            "Impact investment and philanthropic introductions",
            "Social enterprise growth and ecosystem mentorship",
        ]
    )
    add_para("What FairBanks brings", style="Heading 2", bold=True, size=13, color=TEAL)
    add_bullets(
        [
            "An operational community facility as a live innovation and implementation hub",
            "A scalable Community Reach model rooted in local needs",
            "FCHIP - a working MVP moving into structured validation and scale",
            "Experienced healthcare leadership and trusted community relationships",
            "A commitment to measurable, sustainable, and scalable impact",
        ]
    )

    add_para("Our invitation", style="Heading 2", bold=True, size=13, color=TEAL)
    add_para(
        "We would be honoured to explore how the Jay Shetty ecosystem and FairBanks can "
        "work together to improve health, wellbeing, and opportunity for underserved "
        "communities. As FairBanks validates and scales FCHIP, we seek strategic "
        "collaborators who can help accelerate innovation, expand reach, and maximise impact.",
    )
    add_para(
        "Together, we can transform community healthcare into an intelligent, preventive, "
        "and compassionate system that empowers every family and every community to thrive.",
        bold=True,
        color=NAVY,
    )
    add_para(
        "Next step: a brief 10-minute introductory conversation with your partnerships or impact team.",
        size=11,
        color=ORANGE,
        bold=True,
    )

    add_photo_row(
        "team",
        "reception_staff",
        "FairBanks team - ready to serve and partner",
        "Reception team documenting care with care",
    )
    add_para(f"{CONTACT_NAME}  |  {CONTACT_TITLE}", bold=True, size=11, color=NAVY, space_after=2)
    add_para(f"{ORG}  |  {LOCATION}", size=10, color=MUTED, space_after=2)
    add_para(f"{WEBSITE}  |  {EMAIL}  |  {PHONE}", size=10, color=MUTED, space_after=12)

    # ---- Email appendix ----
    doc.add_page_break()
    add_para("Appendix - send-ready outreach email", style="Heading 1", bold=True, size=16, color=NAVY)
    add_para(
        "Copy the message below into your email client. Attach jayshetty_pdf.pdf and "
        "jayshetty_ppt.pptx. Keep the subject line. Prefer a short personal note only if "
        "you have a warm introduction.",
        size=10,
        color=MUTED,
    )
    for line in EMAIL_BODY.strip().split("\n"):
        add_para(line if line else " ", size=10, color=SLATE, space_after=2)

    OUT.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX)
    print(f"DOCX: {DOCX}")


def convert_pdf() -> None:
    import shutil

    import win32com.client

    TMP.mkdir(parents=True, exist_ok=True)
    tmp_pdf = TMP / "jayshetty_pdf_build.pdf"
    if tmp_pdf.exists():
        tmp_pdf.unlink()

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        document.ExportAsFixedFormat(str(tmp_pdf.resolve()), 17)
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()

    if PDF.exists():
        try:
            PDF.unlink()
        except PermissionError as exc:
            raise RuntimeError(f"Close {PDF.name} in any PDF viewer, then re-run.") from exc
    shutil.move(str(tmp_pdf), str(PDF))
    print(f"PDF:  {PDF}")


def _transition(slide) -> None:
    from lxml import etree

    ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    sld = slide._element
    for old in sld.findall(f"{{{ns}}}transition"):
        sld.remove(old)
    transition = etree.SubElement(sld, f"{{{ns}}}transition")
    transition.set("spd", "med")
    etree.SubElement(transition, f"{{{ns}}}fade")


def _add_entrance_anims(slide) -> None:
    from lxml import etree

    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    spids = []
    for shape in slide.shapes:
        try:
            st = int(shape.shape_type) if shape.shape_type is not None else -1
        except Exception:
            st = -1
        has_text = bool(getattr(shape, "has_text_frame", False) and shape.has_text_frame)
        if st == 13 or has_text:
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


def build_pptx() -> None:
    from PIL import Image as PILImage
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    slides_for_anim = []

    def rgb(value):
        return RGBColor.from_string(value)

    def rect(slide, x, y, w, h, fill, line=None, rounded=False):
        kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
        shape.line.color.rgb = rgb(line or fill)
        return shape

    def text(
        slide,
        value,
        x,
        y,
        w,
        h,
        size=18,
        color=SLATE,
        bold=False,
        align=PP_ALIGN.LEFT,
        font="Calibri",
        valign=MSO_ANCHOR.TOP,
    ):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.04)
        tf.vertical_anchor = valign
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = value
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        return box

    def bullets(slide, items, x, y, w, h, size=15, color=SLATE):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "- " + item
            p.space_after = Pt(8)
            p.font.name = "Calibri"
            p.font.size = Pt(size)
            p.font.color.rgb = rgb(color)
        return box

    def crop(slide, path, x, y, w, h):
        with PILImage.open(path) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(
            str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h)
        )
        frame_ratio, image_ratio = w / h, iw / ih
        if image_ratio > frame_ratio:
            amount = (1 - frame_ratio / image_ratio) / 2
            pic.crop_left = pic.crop_right = amount
        else:
            amount = (1 - image_ratio / frame_ratio) / 2
            pic.crop_top = pic.crop_bottom = amount
        return pic

    def fit(slide, path, x, y, max_w, max_h):
        with PILImage.open(path) as im:
            iw, ih = im.size
        image_ratio = iw / ih
        frame_ratio = max_w / max_h
        if image_ratio > frame_ratio:
            pw, ph = max_w, max_w / image_ratio
            px, py = x, y + (max_h - ph) / 2
        else:
            ph, pw = max_h, max_h * image_ratio
            px, py = x + (max_w - pw) / 2, y
        return slide.shapes.add_picture(
            str(path), Inches(px), Inches(py), width=Inches(pw), height=Inches(ph)
        )

    def slide():
        s = prs.slides.add_slide(blank)
        rect(s, 0, 0, 13.333, 7.5, CREAM)
        _transition(s)
        slides_for_anim.append(s)
        return s

    def band(s, kicker, title_value, subtitle=""):
        rect(s, 0, 0, 13.333, 0.12, TEAL)
        text(s, kicker.upper(), 0.55, 0.28, 6.5, 0.28, 10, ORANGE, True)
        text(s, title_value, 0.55, 0.55, 12.1, 0.55, 26, NAVY, True)
        if subtitle:
            text(s, subtitle, 0.58, 1.15, 11.8, 0.32, 13, MUTED)

    def footer(s, number):
        text(s, "FairBanks | Strategic Partnership Brief", 0.55, 7.18, 6.0, 0.18, 9, MUTED)
        text(s, f"{number:02}", 12.2, 7.16, 0.5, 0.18, 9, MUTED, align=PP_ALIGN.RIGHT)

    def photo_pair(s, left_key, right_key, left_cap, right_cap, x=0.55, y=1.7, w=6.0, h=4.5):
        """Two related photos on one row, each with a caption underneath."""
        gap = 0.25
        pw = (w - gap) / 2
        ph = h - 0.55
        crop(s, photo(left_key), x, y, pw, ph)
        crop(s, photo(right_key), x + pw + gap, y, pw, ph)
        text(s, left_cap, x, y + ph + 0.08, pw, 0.4, 11, MUTED, align=PP_ALIGN.CENTER)
        text(
            s,
            right_cap,
            x + pw + gap,
            y + ph + 0.08,
            pw,
            0.4,
            11,
            MUTED,
            align=PP_ALIGN.CENTER,
        )

    # 1 Cover
    s = slide()
    crop(s, photo("cover"), 0, 0, 13.333, 7.5)
    rect(s, 0, 0, 7.4, 7.5, NAVY)
    try:
        s.shapes.add_picture(str(photo("logo")), Inches(0.7), Inches(0.45), height=Inches(0.55))
    except Exception:
        pass
    text(s, "STRATEGIC PARTNERSHIP BRIEF", 0.7, 1.35, 6.0, 0.35, 12, GOLD, True)
    text(s, "FairBanks x\nJay Shetty Ecosystem", 0.7, 1.85, 6.2, 1.6, 32, WHITE, True)
    text(s, SUBTITLE, 0.72, 3.7, 5.9, 0.9, 16, WHITE)
    text(s, SLOGAN, 0.72, 5.85, 4.5, 0.35, 14, GOLD, True)
    text(s, f"{LOCATION}  |  Community Health Intelligence", 0.72, 6.4, 5.5, 0.3, 11, WHITE)

    # 2 Who we are
    s = slide()
    band(s, "Who we are", "A community health enterprise with a scalable model", "Care + Community Reach + Intelligence")
    photo_pair(
        s,
        "facility_street",
        "pharmacy",
        "Medical Centre - clinical anchor",
        "Pharmacy - care beside the clinic",
        x=0.5,
        y=1.7,
        w=7.0,
        h=5.0,
    )
    bullets(
        s,
        [
            "Ugandan social enterprise in Kampala peri-urban communities",
            "Medical centre and pharmacy as the clinical anchor",
            "Community Reach for prevention, education, livelihoods",
            "FCHIP closes the Data & Feedback loop",
            "Uganda first - designed for African scale",
        ],
        7.8,
        2.0,
        5.0,
        4.5,
        14,
    )
    footer(s, 2)

    # 3 Three pillars
    s = slide()
    band(s, "The FairBanks model", "Three pillars. One mission.", SLOGAN)
    cards = [
        ("Medical Centre", "Family and community primary care, diagnostics, pharmacy, referrals", photo("facility_branded")),
        ("Community Reach", "CHWs/VHTs, outreach, school health, CHIS, livelihoods", photo("canopy")),
        ("FCHIP", "AI, GIS, climate fusion, secure EMR/HMS APIs, early warning", photo("dashboard")),
    ]
    for i, (title, body, img) in enumerate(cards):
        x = 0.55 + i * 4.2
        rect(s, x, 1.75, 3.95, 5.0, WHITE, LINE, True)
        crop(s, img, x + 0.15, 1.95, 3.65, 2.2)
        text(s, title, x + 0.25, 4.35, 3.4, 0.4, 18, NAVY, True)
        text(s, body, x + 0.25, 4.85, 3.4, 1.4, 13, MUTED)
    footer(s, 3)

    # 4 Cascade
    s = slide()
    band(
        s,
        "How we care",
        "Community Reach cascade + continuous feedback",
        "FCHIP powers the Data & Feedback loop - learn, improve, serve again.",
    )
    cascade_h = 5.15
    cascade_w = cascade_h * 0.728
    cascade_x, cascade_y = 0.45, 1.6
    rect(s, cascade_x, cascade_y, cascade_w + 0.16, cascade_h + 0.16, WHITE, LINE, True)
    if CONCEPT.exists():
        fit(s, CONCEPT, cascade_x + 0.08, cascade_y + 0.08, cascade_w, cascade_h)
    bullets(
        s,
        [
            "Communities identify needs and own solutions",
            "CHWs / VHTs bridge homes, schools, and care",
            "Programmes deliver screening, education, home visits",
            "Medical centre anchors clinical quality",
            "Research and partners strengthen evidence",
            "CHIS and livelihoods build resilient families",
        ],
        cascade_x + cascade_w + 0.45,
        2.0,
        13.0 - (cascade_x + cascade_w + 0.7),
        4.6,
        15,
    )
    footer(s, 4)

    # 5 Problem
    s = slide()
    band(s, "The gap", "Healthcare still waits for people to get sick", "Fragmented data. Late detection. Missed prevention.")
    photo_pair(
        s,
        "outreach",
        "outreach_camp",
        "Community BP screening",
        "Medical camp consultations",
        x=0.5,
        y=1.7,
        w=6.8,
        h=5.0,
    )
    bullets(
        s,
        [
            "Facilities react after illness appears",
            "CHW, school, clinic, pharmacy signals stay siloed",
            "Climate risk rarely joins health decisions",
            "Clinical records often lock inside EMR/HMS",
            "The feedback loop cannot learn in real time",
        ],
        7.6,
        2.0,
        5.2,
        4.5,
        14,
    )
    footer(s, 5)

    # 6 FCHIP solution
    s = slide()
    band(
        s,
        "The intelligence layer",
        "FCHIP - FairBanks Community Health Intelligence Platform",
        "From reactive treatment to proactive, climate-aware prevention.",
    )
    photo_pair(
        s,
        "mobile",
        "dashboard",
        "Mobile capture for CHWs / VHTs",
        "Dashboards for timely action",
        x=0.5,
        y=1.7,
        w=6.5,
        h=5.0,
    )
    for i, (a, b) in enumerate(
        [
            ("Capture", "Field, school, clinic signals"),
            ("Connect", "Secure EMR/HMS APIs"),
            ("Fuse", "GIS + climate + community"),
            ("Act", "Alerts, referrals, outreach"),
        ]
    ):
        y = 1.85 + i * 1.15
        rect(s, 7.3, y, 5.4, 1.0, WHITE, LINE, True)
        rect(s, 7.3, y, 0.14, 1.0, TEAL)
        text(s, a, 7.6, y + 0.18, 1.5, 0.3, 15, NAVY, True)
        text(s, b, 9.2, y + 0.18, 3.2, 0.55, 13, MUTED)
    footer(s, 6)

    # 7 Deep tech
    s = slide()
    band(s, "Deep technology", "Built for last-mile Africa - not a simple app", "AI - ML - GIS - Climate - Secure APIs - Mobile - Cloud")
    crop(s, photo("architecture"), 6.4, 1.8, 6.3, 4.55)
    text(
        s,
        "FCHIP architecture: capture to intelligence to action",
        6.4,
        6.45,
        6.3,
        0.35,
        11,
        MUTED,
        align=PP_ALIGN.CENTER,
    )
    bullets(
        s,
        [
            "AI/ML risk scoring for outbreaks, maternal risk, NCDs",
            "GIS hotspot mapping by village and catchment",
            "Climate API fusion for rainfall, heat, extremes",
            "Secure EMR/HMS data APIs for real-time clinical ingest",
            "Offline mobile tools for CHWs and VHTs",
            "Facility and partner dashboards for action",
        ],
        0.65,
        1.95,
        5.4,
        4.8,
        14,
    )
    footer(s, 7)

    # 8 Traction
    s = slide()
    band(s, "Traction", "A live ecosystem ready to validate and scale", "Technology with field access - field access with technology.")
    photo_pair(
        s,
        "community",
        "gericare",
        "Community Reach gatherings",
        "Gericare - dignity in ageing",
        x=0.5,
        y=1.7,
        w=8.0,
        h=5.0,
    )
    bullets(
        s,
        [
            "Live medical centre and pharmacy",
            "Named Kampala peri-urban communities",
            "CHW/VHT, MCH, school health",
            "CHIS and livelihood pathways",
            "Working FCHIP MVP in validation",
        ],
        8.8,
        2.0,
        4.0,
        4.5,
        13,
    )
    footer(s, 8)

    # 9 Roadmap
    s = slide()
    band(s, "Roadmap", "From working MVP to continental community health intelligence", "Validate -> district scale -> regional expansion")
    stages = [
        ("Done", "Model + centre", "Community Reach and medical centre live"),
        ("Done", "FCHIP MVP", "Working platform ready for evidence"),
        ("Next", "Validate", "Structured CHW/facility cohorts + proof cases"),
        ("Next", "Partner", "Storytelling, mentorship, impact introductions"),
        ("Future", "District", "Partner clinics and district structures"),
        ("Future", "Regional", "East Africa and pan-African pathways"),
    ]
    for i, (tag, title, body) in enumerate(stages):
        col, row = i % 3, i // 3
        x, y = 0.55 + col * 4.2, 1.85 + row * 2.35
        rect(s, x, y, 3.95, 2.05, WHITE, LINE, True)
        fill = TEAL if tag == "Done" else (ORANGE if tag == "Next" else GREEN)
        rect(s, x, y, 3.95, 0.42, fill)
        text(s, tag.upper(), x + 0.2, y + 0.08, 2.0, 0.28, 11, WHITE, True)
        text(s, title, x + 0.2, y + 0.6, 3.5, 0.35, 16, NAVY, True)
        text(s, body, x + 0.2, y + 1.05, 3.5, 0.7, 13, MUTED)
    footer(s, 9)

    # 10 Why Jay Shetty
    s = slide()
    band(
        s,
        "Why Jay Shetty",
        "Inspiration becomes impact when communities can act",
        "Wellbeing - compassion - purposeful living - stronger communities",
    )
    photo_pair(
        s,
        "doctor_hands",
        "maternal",
        "Compassion in the consulting room",
        "Maternal and family health support",
        x=0.5,
        y=1.7,
        w=7.0,
        h=5.0,
    )
    bullets(
        s,
        [
            "Jay Shetty champions healthier, purposeful lives",
            "FairBanks turns vision into access and prevention",
            "FCHIP helps communities see risk earlier",
            "Global storytelling + local health delivery",
            "A partnership of mission - not a one-way ask",
        ],
        7.8,
        2.0,
        5.0,
        4.5,
        14,
    )
    footer(s, 10)

    # 11 Partnership menu
    s = slide()
    band(s, "Partnership menu", "What we invite - and what we bring", "Two-way value. Clear exchange.")
    rect(s, 0.55, 1.65, 5.9, 3.35, WHITE, LINE, True)
    text(s, "Collaboration areas", 0.8, 1.8, 5.3, 0.35, 16, TEAL, True)
    bullets(
        s,
        [
            "Community wellness and preventive campaigns",
            "Mental wellbeing and holistic health programmes",
            "AI for social impact storytelling",
            "Social enterprise growth support",
            "Introductions to aligned impact partners",
        ],
        0.8,
        2.25,
        5.3,
        2.5,
        13,
    )
    rect(s, 6.8, 1.65, 5.9, 3.35, PALE_TEAL, TEAL, True)
    text(s, "FairBanks brings", 7.05, 1.8, 5.3, 0.35, 16, NAVY, True)
    bullets(
        s,
        [
            "Live clinical and community laboratory",
            "Trusted CHW/VHT and community relationships",
            "Working FCHIP MVP ready to validate",
            "Scalable East Africa community health model",
            "Leadership committed to measurable impact",
        ],
        7.05,
        2.25,
        5.3,
        2.5,
        13,
    )
    photo_pair(
        s,
        "compassion",
        "reception",
        "Human-centred clinical care",
        "Welcoming reception and triage",
        x=0.55,
        y=5.15,
        w=12.2,
        h=1.85,
    )
    footer(s, 11)

    # 12 Close
    s = slide()
    crop(s, photo("close"), 0, 0, 13.333, 7.5)
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    text(s, "OUR INVITATION", 0.8, 1.1, 4.0, 0.35, 13, GOLD, True)
    text(
        s,
        "Partner with us to turn compassion into earlier care for underserved communities.",
        0.8,
        1.7,
        10.5,
        1.5,
        28,
        WHITE,
        True,
    )
    text(
        s,
        "Next step: a brief 10-minute conversation with your partnerships or impact team.",
        0.82,
        3.5,
        10.0,
        0.6,
        16,
        WHITE,
    )
    text(s, SLOGAN, 0.82, 5.5, 4.0, 0.35, 14, GOLD, True)
    text(
        s,
        f"{CONTACT_NAME}  -  {CONTACT_TITLE}\n{EMAIL}  -  {PHONE}  -  {WEBSITE}",
        0.82,
        6.05,
        8.0,
        0.7,
        12,
        WHITE,
    )

    for sl in slides_for_anim:
        _add_entrance_anims(sl)

    OUT.mkdir(parents=True, exist_ok=True)
    prs.save(PPTX)
    print(f"PPTX: {PPTX}")

def validate() -> None:
    from zipfile import BadZipFile, ZipFile

    from docx import Document
    from pptx import Presentation

    for path in (DOCX, PDF, PPTX):
        if not path.exists() or path.stat().st_size < 20_000:
            raise RuntimeError(f"Missing or unexpectedly small output: {path}")
    for path in (DOCX, PPTX):
        try:
            with ZipFile(path) as zf:
                bad = zf.testzip()
                if bad:
                    raise RuntimeError(f"Corrupt archive member: {bad}")
        except BadZipFile as exc:
            raise RuntimeError(f"Corrupt Office file: {path}") from exc

    doc = Document(DOCX)
    content = "\n".join(p.text for p in doc.paragraphs)
    for phrase in (
        "FCHIP",
        "Jay Shetty",
        "Community Reach",
        "send-ready outreach email",
        CONTACT_NAME,
    ):
        if phrase not in content:
            raise RuntimeError(f"DOCX validation failed: missing {phrase}")

    deck = Presentation(PPTX)
    if len(deck.slides) != 12:
        raise RuntimeError(f"Expected 12 slides, found {len(deck.slides)}")
    print(f"Validated: DOCX + PDF + {len(deck.slides)} PPT slides")


def main() -> None:
    print(f"Building Jay Shetty partnership pack ({date.today().isoformat()})")
    OUT.mkdir(parents=True, exist_ok=True)
    build_docx()
    convert_pdf()
    build_pptx()
    validate()
    print("Partnership pack complete.")
    print("\n--- Send-ready email (also in Word/PDF appendix) ---\n")
    print(EMAIL_BODY)


if __name__ == "__main__":
    main()
