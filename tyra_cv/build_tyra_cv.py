"""Generate professional CV for Tyra Rebecca Nalukwago (DOCX + PDF).
Leaves the original TYRA REBECCA NALUKWAGO.docx untouched.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    ListFlowable,
    ListItem,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
DOCX_OUT = ROOT / "TYRA_REBECCA_NALUKWAGO_CV.docx"
PDF_OUT = ROOT / "TYRA_REBECCA_NALUKWAGO_CV.pdf"

NAVY = HexColor("#1B3A4B")
ACCENT = HexColor("#2C5F6E")
MUTED = HexColor("#4A5568")
RULE = HexColor("#CBD5E0")
LIGHT_BG = HexColor("#F0F4F6")

NAVY_RGB = RGBColor(0x1B, 0x3A, 0x4B)
ACCENT_RGB = RGBColor(0x2C, 0x5F, 0x6E)
MUTED_RGB = RGBColor(0x4A, 0x55, 0x68)

PROFILE = (
    "Accomplished Land Economist and Graduate Valuation Surveyor with over five years "
    "of professional experience in property valuation, land acquisition, compensation "
    "assessment, Resettlement Action Plans (RAPs), and property management. Demonstrated "
    "expertise in supporting large-scale infrastructure, energy, mining, water, and "
    "transport projects across Uganda for government agencies, financial institutions, "
    "and private-sector clients."
)
PROFILE_2 = (
    "Highly experienced in valuation methodologies, stakeholder engagement, community "
    "sensitisation, mortgage valuations, market analysis, and preparation of comprehensive "
    "valuation reports. A Graduate Member of the Institute of Surveyors of Uganda (ISU), "
    "committed to delivering professional, ethical, and data-driven valuation services "
    "that support sustainable development and investment."
)

HIGHLIGHTS = [
    "5+ years of professional valuation and land economics experience",
    "Graduate Member, Institute of Surveyors of Uganda (ISU)",
    "Extensive experience in Resettlement Action Plans (RAPs)",
    "Property valuation for mortgage, compensation and acquisition",
    "Worked on major national infrastructure, energy and mining projects",
    "Valuation assignments for leading commercial banks in Uganda",
    "Strong stakeholder engagement and community mobilisation skills",
]

EXPERIENCE = [
    {
        "title": "Valuation Surveyor",
        "org": "COVAE (U) Ltd",
        "dates": "Aug 2023 - Present",
        "bullets": [
            "Lead valuation assignments for mortgage, compensation and acquisition purposes.",
            "Prepare valuation methodologies and comprehensive valuation reports.",
            "Conduct market research and neighbourhood analysis.",
            "Support Resettlement Action Plan implementation for public infrastructure projects.",
        ],
    },
    {
        "title": "Valuer",
        "org": "GMT Consults Ltd",
        "dates": "Sept 2021 - Jul 2023",
        "bullets": [
            "Delivered RAP valuation assignments for road, water and electricity transmission projects.",
            "Coordinated stakeholder consultations, community engagement and grievance management.",
            "Managed valuation data collection, verification and reporting.",
        ],
    },
    {
        "title": "Valuer",
        "org": "Ministry of Energy & Mineral Development",
        "dates": "May 2022 - Jul 2022",
        "bullets": [
            "Negotiated wayleaves for rural electrification projects.",
            "Managed compensation assessments and stakeholder engagement.",
            "Supported electricity transmission corridor development.",
        ],
    },
    {
        "title": "Wayleave Liaison Officer",
        "org": "INTEC-GOPA International Energy Consultants",
        "dates": "May 2021 - Sept 2021",
        "bullets": [
            "Coordinated landowner negotiations and compensation processes.",
            "Facilitated access agreements and community liaison for power distribution projects.",
        ],
    },
    {
        "title": "Valuation Surveyor",
        "org": "COVAE (U) Ltd",
        "dates": "May 2018 - Apr 2020",
        "bullets": [
            "Conducted mortgage, compensation and property valuations.",
            "Prepared valuation reports for commercial banks, private clients and government projects.",
        ],
    },
]

EARLIER = [
    "Intern Valuer - Kampala Capital City Authority",
    "Intern Valuer - Peak Partners Property Valuers & Surveyors",
]

PROJECTS = [
    "Bagkara Mining Land Acquisition Project",
    "Buliisa Service Station Compensation Valuation",
    "Singo Military Training School Compensation Project",
    "Small Hydro Power Dam Resettlement Projects",
    "Isingiro Water Works Project",
    "Muhanga-Kisizi-Rwashamaire Road Project",
    "Masaka-Mbarara 400KV Transmission Line",
    "Rural Electrification Projects across Uganda",
    "Mortgage valuation assignments for major commercial banks",
]

TECHNICAL = [
    "Property Valuation",
    "Resettlement Action Plans (RAP)",
    "Land Acquisition",
    "Compensation Assessment",
    "Property Management",
    "Mortgage Valuation",
    "Property Market Analysis",
    "Community Engagement",
    "Geographic Information Systems (GIS)",
    "AutoCAD",
    "ArchiCAD",
    "REDAH Informatics",
    "VEKTA",
    "Microsoft Office Suite",
]

COMPETENCIES = [
    "Property & Asset Valuation",
    "Land Acquisition & Compensation",
    "Resettlement Action Planning (RAP)",
    "Stakeholder Engagement",
    "Community Mobilisation",
    "Property Market Research",
    "Negotiation & Conflict Resolution",
    "Valuation Reporting",
    "Project Coordination",
    "Data Collection & Analysis",
]


# ---------------------------------------------------------------------------
# DOCX helpers
# ---------------------------------------------------------------------------

def _set_run_font(run, name="Calibri", size=10, bold=False, color=None, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _para_spacing(p, before=0, after=6, line=1.08):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def _add_bottom_border(paragraph, color="2C5F6E", size="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _section_heading(doc, text):
    p = doc.add_paragraph()
    _para_spacing(p, before=12, after=4, line=1.0)
    run = p.add_run(text.upper())
    _set_run_font(run, size=11, bold=True, color=NAVY_RGB)
    _add_bottom_border(p)
    return p


def _bullet(doc, text, size=9.5):
    p = doc.add_paragraph(style="List Bullet")
    _para_spacing(p, before=0, after=2, line=1.05)
    p.clear()
    run = p.add_run(text)
    _set_run_font(run, size=size, color=MUTED_RGB)
    # tighten left indent slightly
    p.paragraph_format.left_indent = Cm(0.55)
    return p


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    # Name
    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_spacing(name, before=0, after=2, line=1.0)
    r = name.add_run("TYRA REBECCA NALUKWAGO")
    _set_run_font(r, size=20, bold=True, color=NAVY_RGB)

    # Tagline
    tag = doc.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_spacing(tag, before=0, after=2, line=1.0)
    r = tag.add_run(
        "Senior Valuation Surveyor  |  Land Economist  |  Property & Resettlement Specialist"
    )
    _set_run_font(r, size=10, color=ACCENT_RGB)

    # Accent rule under header
    rule = doc.add_paragraph()
    _para_spacing(rule, before=2, after=8, line=1.0)
    _add_bottom_border(rule, color="1B3A4B", size="18")

    # Profile
    _section_heading(doc, "Professional Profile")
    for text in (PROFILE, PROFILE_2):
        p = doc.add_paragraph()
        _para_spacing(p, before=2, after=4, line=1.1)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(text)
        _set_run_font(r, size=9.5, color=MUTED_RGB)

    # Highlights - two columns via table
    _section_heading(doc, "Career Highlights")
    mid = (len(HIGHLIGHTS) + 1) // 2
    left, right = HIGHLIGHTS[:mid], HIGHLIGHTS[mid:]
    table = doc.add_table(rows=max(len(left), len(right)), cols=2)
    table.autofit = True
    for i in range(max(len(left), len(right))):
        for col, items in enumerate((left, right)):
            cell = table.rows[i].cells[col]
            cell.text = ""
            p = cell.paragraphs[0]
            _para_spacing(p, before=0, after=1, line=1.05)
            if i < len(items):
                r = p.add_run(f"•  {items[i]}")
                _set_run_font(r, size=9, color=MUTED_RGB)
            # remove cell margins slightly
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement("w:tcMar")
            for m, v in (("top", "20"), ("left", "40"), ("bottom", "20"), ("right", "80")):
                node = OxmlElement(f"w:{m}")
                node.set(qn("w:w"), v)
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
            tcPr.append(tcMar)

    # Experience
    _section_heading(doc, "Professional Experience")
    for job in EXPERIENCE:
        head = doc.add_paragraph()
        _para_spacing(head, before=6, after=0, line=1.05)
        r = head.add_run(job["title"])
        _set_run_font(r, size=10.5, bold=True, color=NAVY_RGB)

        meta = doc.add_paragraph()
        _para_spacing(meta, before=0, after=2, line=1.0)
        r = meta.add_run(f"{job['org']}  |  {job['dates']}")
        _set_run_font(r, size=9, italic=True, color=ACCENT_RGB)

        for b in job["bullets"]:
            _bullet(doc, b)

    # Earlier experience
    early_h = doc.add_paragraph()
    _para_spacing(early_h, before=6, after=2, line=1.0)
    r = early_h.add_run("Earlier Experience")
    _set_run_font(r, size=10, bold=True, color=NAVY_RGB)
    for item in EARLIER:
        _bullet(doc, item)

    # Education + Membership side by side
    _section_heading(doc, "Education & Professional Membership")
    em = doc.add_table(rows=2, cols=2)
    em.rows[0].cells[0].text = ""
    em.rows[0].cells[1].text = ""
    em.rows[1].cells[0].text = ""
    em.rows[1].cells[1].text = ""

    p = em.rows[0].cells[0].paragraphs[0]
    _para_spacing(p, before=2, after=0, line=1.05)
    r = p.add_run("Bachelor of Science (Hons) in Land Economics")
    _set_run_font(r, size=10, bold=True, color=NAVY_RGB)
    p = em.rows[1].cells[0].paragraphs[0]
    _para_spacing(p, before=0, after=2, line=1.0)
    r = p.add_run("Makerere University")
    _set_run_font(r, size=9.5, italic=True, color=ACCENT_RGB)

    p = em.rows[0].cells[1].paragraphs[0]
    _para_spacing(p, before=2, after=0, line=1.05)
    r = p.add_run("Graduate Member")
    _set_run_font(r, size=10, bold=True, color=NAVY_RGB)
    p = em.rows[1].cells[1].paragraphs[0]
    _para_spacing(p, before=0, after=2, line=1.0)
    r = p.add_run("Institute of Surveyors of Uganda (ISU)")
    _set_run_font(r, size=9.5, italic=True, color=ACCENT_RGB)

    # Key Projects - two columns
    _section_heading(doc, "Key Projects")
    mid = (len(PROJECTS) + 1) // 2
    left, right = PROJECTS[:mid], PROJECTS[mid:]
    table = doc.add_table(rows=max(len(left), len(right)), cols=2)
    for i in range(max(len(left), len(right))):
        for col, items in enumerate((left, right)):
            cell = table.rows[i].cells[col]
            cell.text = ""
            p = cell.paragraphs[0]
            _para_spacing(p, before=0, after=1, line=1.05)
            if i < len(items):
                r = p.add_run(f"•  {items[i]}")
                _set_run_font(r, size=9, color=MUTED_RGB)

    # Technical expertise
    _section_heading(doc, "Technical Expertise")
    tech = doc.add_paragraph()
    _para_spacing(tech, before=2, after=4, line=1.15)
    tech.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = tech.add_run("  ·  ".join(TECHNICAL))
    _set_run_font(r, size=9, color=MUTED_RGB)

    # Competencies - two columns
    _section_heading(doc, "Core Competencies")
    mid = (len(COMPETENCIES) + 1) // 2
    left, right = COMPETENCIES[:mid], COMPETENCIES[mid:]
    table = doc.add_table(rows=max(len(left), len(right)), cols=2)
    for i in range(max(len(left), len(right))):
        for col, items in enumerate((left, right)):
            cell = table.rows[i].cells[col]
            cell.text = ""
            p = cell.paragraphs[0]
            _para_spacing(p, before=0, after=1, line=1.05)
            if i < len(items):
                r = p.add_run(f"•  {items[i]}")
                _set_run_font(r, size=9, color=MUTED_RGB)

    # Languages
    _section_heading(doc, "Languages")
    lang = doc.add_paragraph()
    _para_spacing(lang, before=2, after=0, line=1.05)
    r = lang.add_run("English  ·  Fluent")
    _set_run_font(r, size=9.5, color=MUTED_RGB)
    r = lang.add_run("          ")
    _set_run_font(r, size=9.5)
    r = lang.add_run("Luganda  ·  Native")
    _set_run_font(r, size=9.5, color=MUTED_RGB)

    doc.save(DOCX_OUT)
    print(f"Wrote {DOCX_OUT}")


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def _styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CVName",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVTag",
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=ACCENT,
            alignment=TA_CENTER,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVSection",
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=NAVY,
            spaceBefore=8,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVBody",
            fontName="Helvetica",
            fontSize=8.6,
            leading=11,
            textColor=MUTED,
            alignment=TA_JUSTIFY,
            spaceAfter=3,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVJobTitle",
            fontName="Helvetica-Bold",
            fontSize=9.2,
            leading=11.5,
            textColor=NAVY,
            spaceBefore=4,
            spaceAfter=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVMeta",
            fontName="Helvetica-Oblique",
            fontSize=8.2,
            leading=10.5,
            textColor=ACCENT,
            spaceAfter=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVBullet",
            fontName="Helvetica",
            fontSize=8.3,
            leading=10.5,
            textColor=MUTED,
            leftIndent=10,
            bulletIndent=0,
            spaceAfter=0.5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVCell",
            fontName="Helvetica",
            fontSize=8.3,
            leading=11,
            textColor=MUTED,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVSubhead",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=NAVY,
            spaceBefore=5,
            spaceAfter=2,
        )
    )
    return styles


def _section_block(title, styles):
    """Return plain flowables (no nested KeepTogether)."""
    return [
        Paragraph(title.upper(), styles["CVSection"]),
        HRFlowable(width="100%", thickness=1, color=ACCENT, spaceBefore=0, spaceAfter=4),
    ]


def _two_col_bullets(items, styles, col_widths=None):
    mid = (len(items) + 1) // 2
    left, right = items[:mid], items[mid:]
    rows = []
    for i in range(max(len(left), len(right))):
        l = f"•  {left[i]}" if i < len(left) else ""
        r = f"•  {right[i]}" if i < len(right) else ""
        rows.append(
            [
                Paragraph(l, styles["CVCell"]),
                Paragraph(r, styles["CVCell"]),
            ]
        )
    if col_widths is None:
        col_widths = [8.7 * cm, 8.7 * cm]
    t = Table(rows, colWidths=col_widths)
    t.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    return t


def build_pdf():
    styles = _styles()
    doc = SimpleDocTemplate(
        str(PDF_OUT),
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1.3 * cm,
    )
    story = []

    story.append(Paragraph("TYRA REBECCA NALUKWAGO", styles["CVName"]))
    story.append(
        Paragraph(
            "Senior Valuation Surveyor  |  Land Economist  |  Property &amp; Resettlement Specialist",
            styles["CVTag"],
        )
    )
    story.append(
        HRFlowable(width="100%", thickness=1.6, color=NAVY, spaceBefore=0, spaceAfter=6)
    )

    story.extend(_section_block("Professional Profile", styles))
    story.append(Paragraph(PROFILE, styles["CVBody"]))
    story.append(Paragraph(PROFILE_2, styles["CVBody"]))

    story.extend(_section_block("Career Highlights", styles))
    story.append(_two_col_bullets(HIGHLIGHTS, styles))

    story.extend(_section_block("Professional Experience", styles))
    for job in EXPERIENCE:
        block = [
            Paragraph(job["title"], styles["CVJobTitle"]),
            Paragraph(f"{job['org']}  |  {job['dates']}", styles["CVMeta"]),
        ]
        for b in job["bullets"]:
            block.append(Paragraph(f"•  {b}", styles["CVBullet"]))
        story.append(KeepTogether(block))

    story.append(Paragraph("Earlier Experience", styles["CVSubhead"]))
    for item in EARLIER:
        story.append(Paragraph(f"•  {item}", styles["CVBullet"]))

    story.extend(_section_block("Education", styles))
    story.append(
        Paragraph(
            "Bachelor of Science (Hons) in Land Economics", styles["CVJobTitle"]
        )
    )
    story.append(Paragraph("Makerere University", styles["CVMeta"]))

    story.extend(_section_block("Professional Membership", styles))
    story.append(Paragraph("Graduate Member", styles["CVJobTitle"]))
    story.append(
        Paragraph("Institute of Surveyors of Uganda (ISU)", styles["CVMeta"])
    )

    # Keep projects block intact so the 2-col table does not split mid-row
    story.append(
        KeepTogether(
            _section_block("Key Projects", styles)
            + [_two_col_bullets(PROJECTS, styles)]
        )
    )

    story.extend(_section_block("Technical Expertise", styles))
    story.append(Paragraph("  ·  ".join(TECHNICAL), styles["CVBody"]))

    story.append(
        KeepTogether(
            _section_block("Core Competencies", styles)
            + [_two_col_bullets(COMPETENCIES, styles)]
        )
    )

    story.extend(_section_block("Languages", styles))
    story.append(
        Paragraph(
            "English  ·  Fluent&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Luganda  ·  Native",
            styles["CVBody"],
        )
    )

    doc.build(story)
    print(f"Wrote {PDF_OUT}")


def render_preview():
    """Render PDF pages to PNG for visual QA."""
    import pypdfium2 as pdfium

    preview_dir = ROOT / "tmp_preview"
    preview_dir.mkdir(exist_ok=True)
    for old in preview_dir.glob("*.png"):
        old.unlink()
    pdf = pdfium.PdfDocument(str(PDF_OUT))
    paths = []
    print(f"PDF pages: {len(pdf)}")
    for i in range(len(pdf)):
        page = pdf[i]
        bitmap = page.render(scale=2)
        pil = bitmap.to_pil()
        out = preview_dir / f"cv_page_{i + 1}.png"
        pil.save(out)
        paths.append(out)
        print(f"Preview: {out}")
    return paths


if __name__ == "__main__":
    build_docx()
    build_pdf()
    render_preview()
