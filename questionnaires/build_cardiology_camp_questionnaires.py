"""
Build FairBanks Community Cardiology Camp evaluation questionnaires.

Outputs (in this folder):
  - consultant_sho_evaluation.docx
  - consultant_sho_evaluation_fillable.pdf
  - internal_staff_evaluation.docx
  - internal_staff_evaluation_fillable.pdf
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

OUT_DIR = Path(__file__).resolve().parent
BRAND = HexColor("#0B6E6B")
BRAND_DARK = HexColor("#084E4C")
ACCENT = HexColor("#1A8A86")
LIGHT = HexColor("#E8F5F4")
MUTED = HexColor("#5A6A6A")
LINE = HexColor("#C5D5D4")
INK = HexColor("#1C2B2A")
WARN = HexColor("#FFF8E7")

PAGE_W, PAGE_H = A4
MARGIN_L = 16 * mm
MARGIN_R = 16 * mm
MARGIN_T = 14 * mm
MARGIN_B = 16 * mm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R

SUBMIT_EMAIL = "info@fairbanksmedicalcentre.org"
SLOGAN = "Your health, our mission."


# ---------------------------------------------------------------------------
# Shared content
# ---------------------------------------------------------------------------

CONSULTANT_RATINGS = [
    "Overall organisation",
    "Communication before the camp",
    "Coordination during the camp",
    "Patient registration process",
    "Patient flow / triage",
    "Availability of investigations",
    "Availability of medicines",
    "Support from organising team",
    "Teamwork among health workers",
    "Quality of clinical workspace",
    "Overall experience",
]

STAFF_RATINGS = [
    "Pre-camp briefing and role clarity",
    "Communication during the camp",
    "Registration and crowd control",
    "Patient flow and triage",
    "Station setup and workspace",
    "Availability of supplies / consumables",
    "Lab / investigation support",
    "Pharmacy / medicine support",
    "Staffing levels for your station",
    "Breaks, meals and welfare support",
    "Safety and infection prevention",
    "Overall organisation",
]

CONSULTANT_DESIGNATIONS = [
    ("Consultant Physician", "des_consult"),
    ("Cardiologist", "des_cardio"),
    ("Senior House Officer (SHO)", "des_sho"),
    ("Medical Officer", "des_mo"),
    ("Other", "des_other"),
]

STAFF_ROLES = [
    ("Nursing", "role_nursing"),
    ("Clinical / Medical Officer", "role_clinical"),
    ("Pharmacy", "role_pharmacy"),
    ("Laboratory", "role_lab"),
    ("Registration / Records", "role_reg"),
    ("CHW / VHT", "role_chw"),
    ("Admin / Logistics", "role_admin"),
    ("Volunteer", "role_vol"),
    ("Other", "role_other"),
]

PRIORITY_SERVICES = [
    ("Cardiology", "svc_cardio"),
    ("Hypertension", "svc_htn"),
    ("Diabetes", "svc_dm"),
    ("Women's Health", "svc_wh"),
    ("Men's Health", "svc_mh"),
    ("Children's Health", "svc_ch"),
    ("Eye Care", "svc_eye"),
    ("Mental Health", "svc_mental"),
    ("Cancer Screening", "svc_cancer"),
    ("Elderly Care / Gericare", "svc_elderly"),
]

COLLAB_CAPACITY = [
    ("Clinical services", "cap_clinical"),
    ("Training and mentorship", "cap_train"),
    ("Research", "cap_research"),
    ("Community outreach", "cap_outreach"),
    ("Health education", "cap_educ"),
    ("Programme planning", "cap_plan"),
]


# ---------------------------------------------------------------------------
# Word helpers
# ---------------------------------------------------------------------------

def _set_cell_shading(cell, hex_color: str) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def _set_run_font(run, size=11, bold=False, color=None, name="Calibri") -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def _add_heading_block(doc: Document, title: str, subtitle: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("FAIRBANKS MEDICAL CENTRE")
    _set_run_font(r, size=16, bold=True, color=(11, 110, 107))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Community Cardiology Camp")
    _set_run_font(r, size=13, bold=True, color=(28, 43, 42))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    _set_run_font(r, size=12, bold=True, color=(8, 78, 76))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    _set_run_font(r, size=10, color=(90, 106, 106))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(SLOGAN)
    _set_run_font(r, size=9, color=(11, 110, 107))


def _add_para(doc: Document, text: str, *, size=10.5, bold=False, color=(28, 43, 42),
              space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    _set_run_font(r, size=size, bold=bold, color=color)


def _add_section(doc: Document, code: str, title: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"{code}: {title}")
    _set_run_font(r, size=11.5, bold=True, color=(11, 110, 107))


def _add_field_line(doc: Document, label: str, width_hint: str = "_" * 48) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"{label} ")
    _set_run_font(r, size=10.5, bold=True)
    r2 = p.add_run(width_hint)
    _set_run_font(r2, size=10.5, color=(90, 106, 106))


def _add_check_options(doc: Document, options: list[str], per_row: int = 2) -> None:
    row = []
    for i, opt in enumerate(options, 1):
        row.append(f"☐  {opt}")
        if i % per_row == 0 or i == len(options):
            _add_para(doc, "     ".join(row), size=10, space_after=3)
            row = []


def _add_open_prompt(doc: Document, number: str, prompt: str, lines: int = 3) -> None:
    _add_para(doc, f"{number}  {prompt}", size=10.5, bold=True, space_after=3)
    for _ in range(lines):
        _add_para(doc, "_" * 92, size=9, color=(165, 180, 179), space_after=2)


def _add_rating_table(doc: Document, items: list[str]) -> None:
    headers = ["Area", "Excellent\n(5)", "Very Good\n(4)", "Good\n(3)", "Fair\n(2)", "Poor\n(1)"]
    table = doc.add_table(rows=1 + len(items), cols=6)
    table.style = "Table Grid"
    table.autofit = True

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        _set_run_font(r, size=8.5, bold=True, color=(255, 255, 255))
        _set_cell_shading(cell, "0B6E6B")

    for r_idx, item in enumerate(items, 1):
        row = table.rows[r_idx]
        row.cells[0].text = ""
        p = row.cells[0].paragraphs[0]
        r = p.add_run(item)
        _set_run_font(r, size=9)
        if r_idx % 2 == 0:
            _set_cell_shading(row.cells[0], "E8F5F4")
        for c in range(1, 6):
            row.cells[c].text = ""
            p = row.cells[c].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = p.add_run("☐")
            _set_run_font(rr, size=11)
            if r_idx % 2 == 0:
                _set_cell_shading(row.cells[c], "E8F5F4")

    doc.add_paragraph()


def _add_footer_block(doc: Document) -> None:
    _add_section(doc, "SUBMISSION", "How to return this form")
    _add_para(
        doc,
        f"Please email your completed questionnaire to {SUBMIT_EMAIL}. "
        "You may type responses in this Word file, or print, fill by hand, and scan. "
        "Your feedback is confidential and used only for programme improvement, "
        "reporting, and partnership development.",
        size=10,
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run("Thank you for strengthening community cardiovascular care with FairBanks.")
    _set_run_font(r, size=10, bold=True, color=(11, 110, 107))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"FairBanks Medical Centre  ·  {SLOGAN}")
    _set_run_font(r, size=9, color=(90, 106, 106))


def build_consultant_docx(path: Path) -> None:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.6)
        section.bottom_margin = Cm(1.6)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    _add_heading_block(
        doc,
        "Consultant & Senior House Officer (SHO) Post-Camp Evaluation",
        "FairBanks Reach Programme  ·  Confidential feedback  ·  ~10–12 minutes",
    )

    _add_para(
        doc,
        "Dear Consultant / SHO,",
        size=10.5,
        bold=True,
        space_after=4,
    )
    _add_para(
        doc,
        "On behalf of the Board, Management and FairBanks Medical Centre team, thank you for "
        "your contribution to our Community Cardiology Camp. Your expertise helped community "
        "members access specialist cardiovascular screening, diagnosis and advice.",
        size=10.5,
    )
    _add_para(
        doc,
        "As we plan future outreaches under the FairBanks Reach Programme, please share honest "
        "feedback. Your answers will help improve camp quality and strengthen partnership and "
        "sponsorship proposals for long-term community cardiovascular care.",
        size=10.5,
    )
    _add_para(
        doc,
        "Tip: You can edit this Word file on a computer, or print and fill by hand. "
        "A fillable PDF version is also available for soft-copy completion.",
        size=9.5,
        color=(8, 78, 76),
    )

    _add_section(doc, "SECTION A", "General information")
    _add_field_line(doc, "Name (optional):")
    _add_field_line(doc, "Institution / hospital:")
    _add_field_line(doc, "Camp date(s):", "_" * 28)
    _add_field_line(doc, "Camp venue / community:", "_" * 36)
    _add_para(doc, "Designation (tick one):", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, [d[0] for d in CONSULTANT_DESIGNATIONS], per_row=2)
    _add_field_line(doc, "If Other, please specify:")
    _add_field_line(doc, "Years of practice (optional):", "_" * 20)

    _add_section(doc, "SECTION B", "Overall rating")
    _add_para(
        doc,
        "Please rate each area. Scale: 5 = Excellent · 4 = Very Good · 3 = Good · 2 = Fair · 1 = Poor. "
        "Tick one box per row.",
        size=9.5,
        color=(90, 106, 106),
    )
    _add_rating_table(doc, CONSULTANT_RATINGS)

    _add_section(doc, "SECTION C", "Professional feedback")
    _add_open_prompt(doc, "1.", "How would you describe your overall experience at this camp?")
    _add_open_prompt(doc, "2.", "What impressed you most (clinical quality, organisation, teamwork, community response)?")
    _add_open_prompt(doc, "3.", "What challenges did you face (patient volume, investigations, medicines, space, flow)?")
    _add_open_prompt(doc, "4.", "What should we change to improve the next cardiology camp?")
    _add_open_prompt(
        doc,
        "5.",
        "Were there any clinical patterns or high-need findings we should plan for next time?",
        lines=2,
    )

    _add_section(doc, "SECTION D", "Patient follow-up and referral")
    _add_para(
        doc,
        "Many patients need ongoing cardiovascular care after the camp. Please advise how "
        "FairBanks Medical Centre can:",
        size=10,
    )
    _add_open_prompt(doc, "a)", "Improve post-camp follow-up of patients identified with CVD or risk factors.")
    _add_open_prompt(doc, "b)", "Strengthen referral pathways for complex cardiac cases.")
    _add_open_prompt(doc, "c)", "Improve medication adherence and continuity of care.")
    _add_open_prompt(
        doc,
        "d)",
        "Use CHWs, VHTs and digital tools to monitor patients in the community after the camp.",
    )

    _add_section(doc, "SECTION E", "Programme sustainability")
    _add_open_prompt(
        doc,
        "1.",
        "Which strengths of this programme would appeal most to partners and sponsors?",
    )
    _add_para(doc, "2. Which services should FairBanks prioritise in future community outreaches? (tick all that apply)", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, [s[0] for s in PRIORITY_SERVICES] + ["Other"], per_row=3)
    _add_field_line(doc, "If Other, please specify:")
    _add_open_prompt(
        doc,
        "3.",
        "What would make a community cardiovascular programme sustainable over 2–3 years?",
    )

    _add_section(doc, "SECTION F", "Future collaboration")
    _add_para(doc, "Would you be willing to join future FairBanks Reach Programmes?", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, ["Yes", "No", "Maybe"], per_row=3)
    _add_para(doc, "If yes or maybe, in what capacity? (tick all that apply)", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, [c[0] for c in COLLAB_CAPACITY] + ["Other"], per_row=2)
    _add_field_line(doc, "If Other, please specify:")
    _add_field_line(doc, "Preferred contact (phone or email, optional):")

    _add_section(doc, "SECTION G", "Closing remarks")
    _add_open_prompt(
        doc,
        "",
        "Any additional comments, lessons learned or suggestions for the Board and Management:",
        lines=4,
    )

    _add_footer_block(doc)
    doc.save(path)


def build_staff_docx(path: Path) -> None:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.6)
        section.bottom_margin = Cm(1.6)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    _add_heading_block(
        doc,
        "Internal Staff Post-Camp Evaluation",
        "FairBanks Reach Programme  ·  Confidential internal feedback  ·  ~8–10 minutes",
    )

    _add_para(doc, "Dear Colleague,", size=10.5, bold=True, space_after=4)
    _add_para(
        doc,
        "Thank you for serving at the Community Cardiology Camp. Your work at registration, "
        "clinical stations, pharmacy, laboratory, logistics and community support made the "
        "outreach possible.",
        size=10.5,
    )
    _add_para(
        doc,
        "This internal form collects practical feedback so we can improve the next FairBanks "
        "Reach Programme camp. Please be frank — responses are for internal learning and "
        "planning, not individual performance appraisal.",
        size=10.5,
    )
    _add_para(
        doc,
        "Tip: Edit this Word file on a computer, or print and fill by hand. "
        "A fillable PDF version is also available for soft-copy completion.",
        size=9.5,
        color=(8, 78, 76),
    )

    _add_section(doc, "SECTION A", "About you and your role")
    _add_field_line(doc, "Name (optional):")
    _add_field_line(doc, "Department / unit:")
    _add_field_line(doc, "Camp date(s):", "_" * 28)
    _add_field_line(doc, "Camp venue / community:", "_" * 36)
    _add_para(doc, "Main role during the camp (tick one):", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, [r[0] for r in STAFF_ROLES], per_row=2)
    _add_field_line(doc, "If Other, please specify:")
    _add_para(doc, "Was this your first FairBanks community camp?", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, ["Yes", "No"], per_row=2)

    _add_section(doc, "SECTION B", "Operations rating")
    _add_para(
        doc,
        "Rate each area from your station's point of view. "
        "Scale: 5 = Excellent · 4 = Very Good · 3 = Good · 2 = Fair · 1 = Poor. "
        "Tick one box per row.",
        size=9.5,
        color=(90, 106, 106),
    )
    _add_rating_table(doc, STAFF_RATINGS)

    _add_section(doc, "SECTION C", "What worked and what did not")
    _add_open_prompt(doc, "1.", "What worked well at your station or team?")
    _add_open_prompt(doc, "2.", "What caused delays, confusion or bottlenecks?")
    _add_open_prompt(doc, "3.", "Were supplies, forms, devices or medicines missing or short? Please list.")
    _add_open_prompt(doc, "4.", "How clear was your briefing and task assignment before and during the camp?")
    _add_open_prompt(
        doc,
        "5.",
        "How would you improve patient registration, triage and flow next time?",
    )

    _add_section(doc, "SECTION D", "Teamwork, welfare and safety")
    _add_open_prompt(doc, "1.", "How was teamwork between clinical, support and community teams?")
    _add_open_prompt(doc, "2.", "Did you have enough staff at your station? If not, what support was missing?")
    _add_open_prompt(doc, "3.", "Any safety, infection-prevention or patient-privacy concerns?")
    _add_open_prompt(doc, "4.", "Comments on breaks, meals, transport or staff welfare:")

    _add_section(doc, "SECTION E", "Community and continuity")
    _add_open_prompt(
        doc,
        "1.",
        "What community needs or patient patterns stood out (e.g. hypertension, diabetes, late presentation)?",
    )
    _add_open_prompt(
        doc,
        "2.",
        "How can CHWs / VHTs and FairBanks better support follow-up after the camp?",
    )
    _add_para(doc, "3. Which services should future camps prioritise? (tick all that apply)", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, [s[0] for s in PRIORITY_SERVICES] + ["Other"], per_row=3)
    _add_field_line(doc, "If Other, please specify:")

    _add_section(doc, "SECTION F", "Future camps")
    _add_para(doc, "Would you like to serve again at future FairBanks Reach camps?", size=10.5, bold=True, space_after=3)
    _add_check_options(doc, ["Yes", "No", "Maybe"], per_row=3)
    _add_open_prompt(doc, "1.", "What training or tools would help you perform better at the next camp?")
    _add_open_prompt(doc, "2.", "One change that would most improve the next camp:")

    _add_section(doc, "SECTION G", "Closing remarks")
    _add_open_prompt(
        doc,
        "",
        "Any other comments for Management (logistics, partnerships, recognition, ideas):",
        lines=4,
    )

    _add_footer_block(doc)
    doc.save(path)


# ---------------------------------------------------------------------------
# PDF helpers (fillable)
# ---------------------------------------------------------------------------

class FormPDF:
    def __init__(self, path: Path, doc_title: str):
        self.path = path
        self.c = canvas.Canvas(str(path), pagesize=A4)
        self.c.setTitle(doc_title)
        self.c.setAuthor("FairBanks Medical Centre")
        self.y = PAGE_H - MARGIN_T
        self.page = 1
        self._field_n = 0

    def new_name(self, prefix: str) -> str:
        self._field_n += 1
        return f"{prefix}_{self._field_n}"

    def ensure(self, need: float) -> None:
        if self.y - need < MARGIN_B + 10 * mm:
            self.footer()
            self.c.showPage()
            self.page += 1
            self.y = PAGE_H - MARGIN_T
            self.header_mini()

    def footer(self) -> None:
        self.c.setStrokeColor(LINE)
        self.c.setLineWidth(0.4)
        self.c.line(MARGIN_L, 11 * mm, PAGE_W - MARGIN_R, 11 * mm)
        self.c.setFillColor(MUTED)
        self.c.setFont("Helvetica", 7.5)
        self.c.drawString(MARGIN_L, 6.5 * mm, "FairBanks Medical Centre  ·  Confidential")
        self.c.drawRightString(PAGE_W - MARGIN_R, 6.5 * mm, f"Page {self.page}")

    def header_mini(self) -> None:
        self.c.setFillColor(BRAND)
        self.c.rect(0, PAGE_H - 8 * mm, PAGE_W, 8 * mm, fill=1, stroke=0)
        self.c.setFillColor(white)
        self.c.setFont("Helvetica-Bold", 8)
        self.c.drawString(MARGIN_L, PAGE_H - 5.5 * mm, "FairBanks  ·  Community Cardiology Camp Evaluation")
        self.y = PAGE_H - 14 * mm

    def header_full(self, title: str, audience: str) -> None:
        self.c.setFillColor(BRAND)
        self.c.rect(0, PAGE_H - 28 * mm, PAGE_W, 28 * mm, fill=1, stroke=0)
        self.c.setFillColor(white)
        self.c.setFont("Helvetica-Bold", 14)
        self.c.drawCentredString(PAGE_W / 2, PAGE_H - 10 * mm, "FAIRBANKS MEDICAL CENTRE")
        self.c.setFont("Helvetica", 10)
        self.c.drawCentredString(PAGE_W / 2, PAGE_H - 16 * mm, "Community Cardiology Camp")
        self.c.setFont("Helvetica-Oblique", 8)
        self.c.drawCentredString(PAGE_W / 2, PAGE_H - 22 * mm, SLOGAN)

        self.y = PAGE_H - 34 * mm
        self.c.setFillColor(BRAND_DARK)
        self.c.setFont("Helvetica-Bold", 11)
        self.c.drawCentredString(PAGE_W / 2, self.y, title)
        self.y -= 5 * mm
        self.c.setFillColor(MUTED)
        self.c.setFont("Helvetica", 8.5)
        self.c.drawCentredString(PAGE_W / 2, self.y, audience)
        self.y -= 7 * mm

        # Soft-copy tip
        self.ensure(14 * mm)
        box_h = 12 * mm
        self.c.setFillColor(WARN)
        self.c.setStrokeColor(HexColor("#E0C070"))
        self.c.roundRect(MARGIN_L, self.y - box_h + 2 * mm, CONTENT_W, box_h, 3, fill=1, stroke=1)
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Bold", 8)
        self.c.drawString(MARGIN_L + 3 * mm, self.y - 2 * mm, "Soft-copy friendly (fillable PDF)")
        self.c.setFont("Helvetica", 7.5)
        self.c.drawString(
            MARGIN_L + 3 * mm,
            self.y - 6 * mm,
            "Click any field to type. Use checkboxes and rating circles to select. Save a copy before emailing.",
        )
        self.c.drawString(
            MARGIN_L + 3 * mm,
            self.y - 9.5 * mm,
            f"Submit to: {SUBMIT_EMAIL}   ·   Responses are confidential.",
        )
        self.y -= box_h + 3 * mm

    def intro(self, paragraphs: list[str]) -> None:
        for text in paragraphs:
            self.ensure(12 * mm)
            self.c.setFillColor(INK)
            self.c.setFont("Helvetica", 8.5)
            self._wrap(text, MARGIN_L, CONTENT_W, leading=10)
            self.y -= 2 * mm

    def _wrap(self, text: str, x: float, width: float, leading: float = 10, font="Helvetica", size=8.5) -> None:
        self.c.setFont(font, size)
        words = text.split()
        line = ""
        for w in words:
            trial = f"{line} {w}".strip()
            if self.c.stringWidth(trial, font, size) <= width:
                line = trial
            else:
                self.ensure(leading + 1)
                self.c.drawString(x, self.y, line)
                self.y -= leading
                line = w
        if line:
            self.ensure(leading + 1)
            self.c.drawString(x, self.y, line)
            self.y -= leading

    def section(self, code: str, title: str) -> None:
        self.ensure(12 * mm)
        self.y -= 2 * mm
        self.c.setFillColor(LIGHT)
        self.c.roundRect(MARGIN_L, self.y - 4 * mm, CONTENT_W, 7 * mm, 2, fill=1, stroke=0)
        self.c.setFillColor(BRAND)
        self.c.setFont("Helvetica-Bold", 9.5)
        self.c.drawString(MARGIN_L + 2 * mm, self.y - 1.5 * mm, f"{code}  ·  {title}")
        self.y -= 8 * mm

    def label(self, text: str, *, bold: bool = True, size: float = 8.5) -> None:
        self.ensure(6 * mm)
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        self.c.drawString(MARGIN_L, self.y, text)
        self.y -= 4.5 * mm

    def text_line(self, label: str, name: str, field_w: float | None = None) -> None:
        self.ensure(8 * mm)
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica", 8.5)
        self.c.drawString(MARGIN_L, self.y, label)
        lw = self.c.stringWidth(label, "Helvetica", 8.5) + 3 * mm
        fw = field_w or (CONTENT_W - lw)
        self.c.acroForm.textfield(
            name=name,
            x=MARGIN_L + lw,
            y=self.y - 1.5 * mm,
            width=fw,
            height=5.5 * mm,
            borderWidth=0.6,
            borderColor=LINE,
            fillColor=white,
            textColor=INK,
            forceBorder=True,
            fontSize=8,
        )
        self.y -= 8 * mm

    def multiline(self, prompt: str, name: str, height: float = 18 * mm) -> None:
        self.ensure(height + 8 * mm)
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Bold", 8.5)
        self._wrap(prompt, MARGIN_L, CONTENT_W, leading=10, font="Helvetica-Bold", size=8.5)
        self.y -= 1 * mm
        self.ensure(height + 2 * mm)
        self.c.acroForm.textfield(
            name=name,
            tooltip=prompt[:120],
            x=MARGIN_L,
            y=self.y - height,
            width=CONTENT_W,
            height=height,
            borderWidth=0.6,
            borderColor=LINE,
            fillColor=HexColor("#FAFCFC"),
            textColor=INK,
            forceBorder=True,
            fontSize=8,
            fieldFlags="multiline",
        )
        self.y -= height + 4 * mm

    def check_row(self, options: list[tuple[str, str]], cols: int = 2) -> None:
        col_w = CONTENT_W / cols
        row_h = 6.5 * mm
        for i, (label, name) in enumerate(options):
            if i % cols == 0:
                self.ensure(row_h + 1 * mm)
                row_y = self.y
            col = i % cols
            x = MARGIN_L + col * col_w
            self.c.acroForm.checkbox(
                name=name,
                x=x,
                y=row_y - 1 * mm,
                size=9,
                buttonStyle="check",
                borderColor=BRAND,
                fillColor=white,
                textColor=BRAND,
                forceBorder=True,
                checked=False,
                fieldFlags="",
            )
            self.c.setFillColor(INK)
            self.c.setFont("Helvetica", 8)
            self.c.drawString(x + 4.5 * mm, row_y, label)
            if col == cols - 1 or i == len(options) - 1:
                self.y = row_y - row_h

    def radio_yes_no_maybe(self, group: str) -> None:
        opts = [("Yes", "Yes"), ("No", "No"), ("Maybe", "Maybe")]
        self.ensure(8 * mm)
        x = MARGIN_L
        for label, val in opts:
            self.c.acroForm.radio(
                name=group,
                value=val,
                x=x,
                y=self.y - 1 * mm,
                size=9,
                selected=False,
                buttonStyle="circle",
                borderColor=BRAND,
                fillColor=white,
                textColor=BRAND,
                forceBorder=True,
                fieldFlags="radio",
            )
            self.c.setFillColor(INK)
            self.c.setFont("Helvetica", 8.5)
            self.c.drawString(x + 4.5 * mm, self.y, label)
            x += 32 * mm
        self.y -= 8 * mm

    def rating_table(self, items: list[str], prefix: str) -> None:
        headers = ["Area", "5", "4", "3", "2", "1"]
        col_area = CONTENT_W - 5 * 11 * mm
        col_w = 11 * mm
        row_h = 6.2 * mm
        head_h = 7 * mm

        self.ensure(head_h + row_h * min(3, len(items)) + 4 * mm)
        self.c.setFillColor(MUTED)
        self.c.setFont("Helvetica", 7)
        self.c.drawString(MARGIN_L, self.y, "Scale: 5 Excellent · 4 Very Good · 3 Good · 2 Fair · 1 Poor  (select one per row)")
        self.y -= 4 * mm

        # header
        self.ensure(head_h + 2 * mm)
        self.c.setFillColor(BRAND)
        self.c.rect(MARGIN_L, self.y - head_h + 2 * mm, CONTENT_W, head_h, fill=1, stroke=0)
        self.c.setFillColor(white)
        self.c.setFont("Helvetica-Bold", 7.5)
        self.c.drawString(MARGIN_L + 2 * mm, self.y - 2.5 * mm, headers[0])
        for i, h in enumerate(headers[1:]):
            cx = MARGIN_L + col_area + i * col_w + col_w / 2
            self.c.drawCentredString(cx, self.y - 2.5 * mm, h)
        self.y -= head_h

        for idx, item in enumerate(items):
            self.ensure(row_h + 1 * mm)
            if idx % 2 == 0:
                self.c.setFillColor(LIGHT)
                self.c.rect(MARGIN_L, self.y - row_h + 2 * mm, CONTENT_W, row_h, fill=1, stroke=0)
            self.c.setFillColor(INK)
            self.c.setFont("Helvetica", 7.5)
            self.c.drawString(MARGIN_L + 2 * mm, self.y - 1.5 * mm, item)
            group = f"{prefix}_r{idx}"
            for score in range(5, 0, -1):
                i = 5 - score
                cx = MARGIN_L + col_area + i * col_w + col_w / 2 - 4
                self.c.acroForm.radio(
                    name=group,
                    value=str(score),
                    x=cx,
                    y=self.y - 2.2 * mm,
                    size=8,
                    selected=False,
                    buttonStyle="circle",
                    borderColor=BRAND,
                    fillColor=white,
                    textColor=BRAND,
                    forceBorder=True,
                    fieldFlags="radio",
                )
            self.y -= row_h
        self.y -= 2 * mm

    def save(self) -> None:
        self.footer()
        self.c.save()


def build_consultant_pdf(path: Path) -> None:
    pdf = FormPDF(path, "FairBanks Consultant/SHO Cardiology Camp Evaluation")
    pdf.header_full(
        "Consultant & SHO Post-Camp Evaluation Questionnaire",
        "FairBanks Reach Programme  ·  Confidential  ·  About 10–12 minutes",
    )
    pdf.intro(
        [
            "Dear Consultant / SHO, thank you for supporting our Community Cardiology Camp. "
            "Your expertise helped community members access specialist cardiovascular screening, "
            "diagnosis and medical advice.",
            "Please share honest feedback to improve future FairBanks Reach Programme camps and "
            "strengthen partnership and sponsorship proposals for long-term community cardiovascular care.",
        ]
    )

    pdf.section("SECTION A", "General information")
    pdf.text_line("Name (optional):", "c_name")
    pdf.text_line("Institution / hospital:", "c_institution")
    pdf.text_line("Camp date(s):", "c_date", field_w=55 * mm)
    pdf.text_line("Camp venue / community:", "c_venue")
    pdf.label("Designation (select all that apply):")
    pdf.check_row(CONSULTANT_DESIGNATIONS, cols=2)
    pdf.text_line("If Other, specify:", "c_des_other")
    pdf.text_line("Years of practice (optional):", "c_years", field_w=30 * mm)

    pdf.section("SECTION B", "Overall rating")
    pdf.rating_table(CONSULTANT_RATINGS, "c")

    pdf.section("SECTION C", "Professional feedback")
    pdf.multiline("1. How would you describe your overall experience at this camp?", "c_exp", height=16 * mm)
    pdf.multiline("2. What impressed you most?", "c_highlights", height=16 * mm)
    pdf.multiline("3. What challenges did you face?", "c_challenges", height=16 * mm)
    pdf.multiline("4. What should we change for the next cardiology camp?", "c_recommend", height=16 * mm)
    pdf.multiline(
        "5. Any clinical patterns or high-need findings to plan for next time?",
        "c_patterns",
        height=14 * mm,
    )

    pdf.section("SECTION D", "Patient follow-up and referral")
    pdf.multiline("a) How can FairBanks improve post-camp follow-up for CVD / risk patients?", "c_fu_a", height=16 * mm)
    pdf.multiline("b) How can we strengthen referral pathways for complex cardiac cases?", "c_fu_b", height=16 * mm)
    pdf.multiline("c) How can we improve medication adherence and continuity of care?", "c_fu_c", height=16 * mm)
    pdf.multiline(
        "d) How can CHWs, VHTs and digital tools help monitor patients after the camp?",
        "c_fu_d",
        height=16 * mm,
    )

    pdf.section("SECTION E", "Programme sustainability")
    pdf.multiline("1. Which strengths would appeal most to partners and sponsors?", "c_strengths", height=16 * mm)
    pdf.label("2. Priority services for future outreaches (select all that apply):")
    pdf.check_row(PRIORITY_SERVICES + [("Other", "svc_other")], cols=3)
    pdf.text_line("If Other, specify:", "c_svc_other_txt")
    pdf.multiline(
        "3. What would make a community cardiovascular programme sustainable over 2–3 years?",
        "c_sustain",
        height=16 * mm,
    )

    pdf.section("SECTION F", "Future collaboration")
    pdf.label("Would you join future FairBanks Reach Programmes?")
    pdf.radio_yes_no_maybe("c_future")
    pdf.label("If yes or maybe, in what capacity? (select all that apply)")
    pdf.check_row(COLLAB_CAPACITY + [("Other", "cap_other")], cols=2)
    pdf.text_line("If Other, specify:", "c_cap_other_txt")
    pdf.text_line("Preferred contact (optional):", "c_contact")

    pdf.section("SECTION G", "Closing remarks")
    pdf.multiline(
        "Additional comments, lessons learned or suggestions for the Board and Management:",
        "c_closing",
        height=22 * mm,
    )

    # Keep thank-you on the same page as closing remarks when possible
    if pdf.y < MARGIN_B + 16 * mm:
        pdf.footer()
        pdf.c.showPage()
        pdf.page += 1
        pdf.y = PAGE_H - MARGIN_T
        pdf.header_mini()
    pdf.c.setFillColor(BRAND)
    pdf.c.setFont("Helvetica-Bold", 9)
    pdf.c.drawCentredString(PAGE_W / 2, pdf.y, "Thank you for strengthening community cardiovascular care.")
    pdf.y -= 5 * mm
    pdf.c.setFillColor(MUTED)
    pdf.c.setFont("Helvetica", 8)
    pdf.c.drawCentredString(PAGE_W / 2, pdf.y, f"Submit to {SUBMIT_EMAIL}  ·  {SLOGAN}")

    pdf.save()


def build_staff_pdf(path: Path) -> None:
    pdf = FormPDF(path, "FairBanks Internal Staff Cardiology Camp Evaluation")
    pdf.header_full(
        "Internal Staff Post-Camp Evaluation Questionnaire",
        "FairBanks Reach Programme  ·  Confidential internal use  ·  About 8–10 minutes",
    )
    pdf.intro(
        [
            "Dear Colleague, thank you for serving at the Community Cardiology Camp. Your work at "
            "registration, clinical stations, pharmacy, laboratory, logistics and community support "
            "made the outreach possible.",
            "This internal form is for learning and planning — not individual performance appraisal. "
            "Please be frank so we can improve the next FairBanks Reach Programme camp.",
        ]
    )

    pdf.section("SECTION A", "About you and your role")
    pdf.text_line("Name (optional):", "s_name")
    pdf.text_line("Department / unit:", "s_dept")
    pdf.text_line("Camp date(s):", "s_date", field_w=55 * mm)
    pdf.text_line("Camp venue / community:", "s_venue")
    pdf.label("Main role during the camp (select one or more):")
    pdf.check_row(STAFF_ROLES, cols=2)
    pdf.text_line("If Other, specify:", "s_role_other")
    pdf.label("Was this your first FairBanks community camp?")
    pdf.ensure(8 * mm)
    x = MARGIN_L
    for label, val in [("Yes", "Yes"), ("No", "No")]:
        pdf.c.acroForm.radio(
            name="s_first",
            value=val,
            x=x,
            y=pdf.y - 1 * mm,
            size=9,
            selected=False,
            buttonStyle="circle",
            borderColor=BRAND,
            fillColor=white,
            textColor=BRAND,
            forceBorder=True,
            fieldFlags="radio",
        )
        pdf.c.setFillColor(INK)
        pdf.c.setFont("Helvetica", 8.5)
        pdf.c.drawString(x + 4.5 * mm, pdf.y, label)
        x += 28 * mm
    pdf.y -= 8 * mm

    pdf.section("SECTION B", "Operations rating")
    pdf.rating_table(STAFF_RATINGS, "s")

    pdf.section("SECTION C", "What worked and what did not")
    pdf.multiline("1. What worked well at your station or team?", "s_worked")
    pdf.multiline("2. What caused delays, confusion or bottlenecks?", "s_delays")
    pdf.multiline("3. Were supplies, forms, devices or medicines missing or short? List them.", "s_supplies")
    pdf.multiline("4. How clear was briefing and task assignment before and during the camp?", "s_briefing")
    pdf.multiline("5. How would you improve registration, triage and patient flow next time?", "s_flow")

    pdf.section("SECTION D", "Teamwork, welfare and safety")
    pdf.multiline("1. How was teamwork between clinical, support and community teams?", "s_team")
    pdf.multiline("2. Did you have enough staff at your station? If not, what was missing?", "s_staffing")
    pdf.multiline("3. Any safety, infection-prevention or patient-privacy concerns?", "s_safety")
    pdf.multiline("4. Comments on breaks, meals, transport or staff welfare:", "s_welfare")

    pdf.section("SECTION E", "Community and continuity")
    pdf.multiline(
        "1. What community needs or patient patterns stood out?",
        "s_patterns",
    )
    pdf.multiline(
        "2. How can CHWs / VHTs and FairBanks better support follow-up after the camp?",
        "s_followup",
    )
    pdf.label("3. Priority services for future camps (select all that apply):")
    pdf.check_row(PRIORITY_SERVICES + [("Other", "s_svc_other")], cols=3)
    pdf.text_line("If Other, specify:", "s_svc_other_txt")

    pdf.section("SECTION F", "Future camps")
    pdf.label("Would you like to serve again at future FairBanks Reach camps?")
    pdf.radio_yes_no_maybe("s_future")
    pdf.multiline("1. What training or tools would help you at the next camp?", "s_training")
    pdf.multiline("2. One change that would most improve the next camp:", "s_one_change", height=18 * mm)

    pdf.section("SECTION G", "Closing remarks")
    pdf.multiline(
        "Any other comments for Management (logistics, partnerships, recognition, ideas):",
        "s_closing",
        height=22 * mm,
    )

    if pdf.y < MARGIN_B + 16 * mm:
        pdf.footer()
        pdf.c.showPage()
        pdf.page += 1
        pdf.y = PAGE_H - MARGIN_T
        pdf.header_mini()
    pdf.c.setFillColor(BRAND)
    pdf.c.setFont("Helvetica-Bold", 9)
    pdf.c.drawCentredString(PAGE_W / 2, pdf.y, "Thank you - your feedback improves every FairBanks Reach camp.")
    pdf.y -= 5 * mm
    pdf.c.setFillColor(MUTED)
    pdf.c.setFont("Helvetica", 8)
    pdf.c.drawCentredString(PAGE_W / 2, pdf.y, f"Submit to {SUBMIT_EMAIL}  ·  {SLOGAN}")

    pdf.save()


def main() -> None:
    # Clean leftover from earlier typo helper if present — rebuild all outputs
    files = {
        "consultant_sho_evaluation.docx": build_consultant_docx,
        "internal_staff_evaluation.docx": build_staff_docx,
        "consultant_sho_evaluation_fillable.pdf": build_consultant_pdf,
        "internal_staff_evaluation_fillable.pdf": build_staff_pdf,
    }
    for name, builder in files.items():
        out = OUT_DIR / name
        builder(out)
        print(f"Wrote {out}")


if __name__ == "__main__":
    main()
