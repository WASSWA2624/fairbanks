#!/usr/bin/env python3
"""Build five detailed CV designs from racheal_cv.md.

The Markdown file is the sole content source. All PDF and Word deliverables
are written into one detailed_versions folder.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Flowable,
    HRFlowable,
    Image as RLImage,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "racheal_cv.md"
PHOTO = ROOT / "brief_versions" / "_assets" / "photo_portrait.png"
OUTPUT = ROOT / "detailed_versions"
PREVIEW = ROOT.parent / "tmp" / "detailed_cv_previews"


@dataclass(frozen=True)
class Theme:
    slug: str
    label: str
    primary: str
    secondary: str
    accent: str
    soft: str
    header_bg: str
    header_light: bool
    body_font: str
    heading_font: str
    layout: str


THEMES = [
    Theme(
        "01_classic_executive",
        "Classic Executive",
        "#0A3A52",
        "#1F6F78",
        "#B8953E",
        "#F5F7FA",
        "#F8F5F0",
        True,
        "Helvetica",
        "Helvetica-Bold",
        "classic",
    ),
    Theme(
        "02_modern_navy",
        "Modern Navy",
        "#0B2942",
        "#2F7185",
        "#D2AA52",
        "#EEF4F7",
        "#0B2942",
        False,
        "Helvetica",
        "Helvetica-Bold",
        "modern",
    ),
    Theme(
        "03_academic_leadership",
        "Academic Leadership",
        "#5B2434",
        "#7A4B58",
        "#C39B58",
        "#F8F3F4",
        "#F8F3F4",
        True,
        "Times-Roman",
        "Times-Bold",
        "academic",
    ),
    Theme(
        "04_career_timeline",
        "Career Timeline",
        "#234E52",
        "#2C7A7B",
        "#C39B58",
        "#EFF6F6",
        "#EFF6F6",
        True,
        "Helvetica",
        "Helvetica-Bold",
        "timeline",
    ),
    Theme(
        "05_editorial_profile",
        "Editorial Profile",
        "#252B34",
        "#4B6070",
        "#B8953E",
        "#F3F4F6",
        "#F7F3ED",
        True,
        "Helvetica",
        "Helvetica-Bold",
        "editorial",
    ),
]


def clean_inline(text: str) -> str:
    """Convert Markdown inline markup to clean printable text."""
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = text.replace("**", "").replace("`", "")
    text = text.replace("*(Ongoing)*", "(Ongoing)")
    text = text.replace("*", "")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u00a0", " ")
    return text.strip()


def parse_source() -> tuple[dict[str, str], list[tuple[str, str]]]:
    """Read header details and ordered body blocks directly from Markdown."""
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    header = {
        "name": clean_inline(lines[0].lstrip("#").strip()),
        "title": clean_inline(lines[2]),
        "location": clean_inline(lines[4]),
        "mobile": clean_inline(lines[6]).replace("Mobile:", "").strip(),
        "email": clean_inline(lines[8]).replace("Email:", "").strip(),
        "website": clean_inline(lines[10]).replace("Website:", "").strip(),
    }

    blocks: list[tuple[str, str]] = []
    for raw in lines[12:]:
        line = raw.strip()
        if not line:
            continue
        if line == "---":
            blocks.append(("separator", ""))
        elif line.startswith("# "):
            blocks.append(("section", clean_inline(line[2:])))
        elif line.startswith("## "):
            blocks.append(("role", clean_inline(line[3:])))
        elif line.startswith("### "):
            blocks.append(("subheading", clean_inline(line[4:])))
        elif line.startswith("* "):
            blocks.append(("bullet", clean_inline(line[2:])))
        else:
            blocks.append(("paragraph", clean_inline(line)))
    return header, blocks


class SectionBand(Flowable):
    def __init__(self, text: str, theme: Theme):
        super().__init__()
        self.text = text
        self.theme = theme
        self.height = 22

    def wrap(self, avail_width, avail_height):
        self.width = avail_width
        return avail_width, self.height

    def draw(self):
        primary = HexColor(self.theme.primary)
        accent = HexColor(self.theme.accent)
        if self.theme.layout in {"modern", "editorial"}:
            self.canv.setFillColor(primary)
            self.canv.roundRect(0, 1, self.width, 19, 3, fill=1, stroke=0)
            self.canv.setFillColor(white)
            self.canv.setFont(self.theme.heading_font, 10)
            self.canv.drawString(8, 7, self.text.upper())
        else:
            self.canv.setFillColor(primary)
            self.canv.setFont(self.theme.heading_font, 11)
            self.canv.drawString(0, 8, self.text.upper())
            self.canv.setStrokeColor(accent)
            self.canv.setLineWidth(1.6)
            self.canv.line(0, 3, min(self.width, 105), 3)


def pdf_styles(theme: Theme):
    styles = getSampleStyleSheet()
    primary = HexColor(theme.primary)
    secondary = HexColor(theme.secondary)
    ink = HexColor("#25313D")
    muted = HexColor("#607080")

    def add(name, **kwargs):
        styles.add(ParagraphStyle(name=name, **kwargs))

    add(
        "CvName",
        fontName=theme.heading_font,
        fontSize=19,
        leading=22,
        textColor=primary,
        spaceAfter=3,
    )
    add(
        "CvNameWhite",
        fontName=theme.heading_font,
        fontSize=19,
        leading=22,
        textColor=white,
        spaceAfter=3,
    )
    add(
        "CvTitle",
        fontName=theme.body_font,
        fontSize=9.5,
        leading=12,
        textColor=secondary,
        spaceAfter=5,
    )
    add(
        "CvTitleWhite",
        fontName=theme.body_font,
        fontSize=9.5,
        leading=12,
        textColor=HexColor("#DCEAF0"),
        spaceAfter=5,
    )
    add(
        "CvContact",
        fontName=theme.body_font,
        fontSize=8.5,
        leading=11,
        textColor=muted,
    )
    add(
        "CvContactWhite",
        fontName=theme.body_font,
        fontSize=8.5,
        leading=11,
        textColor=white,
    )
    add(
        "CvBody",
        fontName=theme.body_font,
        fontSize=9.4,
        leading=13.2,
        textColor=ink,
        alignment=TA_JUSTIFY,
        spaceAfter=6,
    )
    add(
        "CvBullet",
        fontName=theme.body_font,
        fontSize=9.2,
        leading=12.5,
        textColor=ink,
        leftIndent=13,
        firstLineIndent=-8,
        spaceAfter=2.5,
    )
    add(
        "CvRole",
        fontName=theme.heading_font,
        fontSize=11,
        leading=13.5,
        textColor=primary,
        spaceBefore=7,
        spaceAfter=2,
    )
    add(
        "CvSub",
        fontName=theme.heading_font,
        fontSize=9.5,
        leading=12,
        textColor=secondary,
        spaceBefore=4,
        spaceAfter=3,
    )
    add(
        "CvSmall",
        fontName=theme.body_font,
        fontSize=8,
        leading=10,
        textColor=muted,
    )
    return styles


def pdf_header(header: dict[str, str], theme: Theme, styles):
    header_light = theme.header_light
    name_style = styles["CvName"] if header_light else styles["CvNameWhite"]
    title_style = styles["CvTitle"] if header_light else styles["CvTitleWhite"]
    contact_style = styles["CvContact"] if header_light else styles["CvContactWhite"]

    photo = RLImage(str(PHOTO), width=31 * mm, height=47 * mm)
    text = [
        Paragraph(header["name"].upper(), name_style),
        Paragraph(header["title"], title_style),
        Paragraph(
            f'{header["location"]}<br/>'
            f'Mobile: {header["mobile"]}<br/>'
            f'Email: {header["email"]}<br/>'
            f'Website: {header["website"]}',
            contact_style,
        ),
    ]
    width = A4[0] - 28 * mm
    table = Table([[text, photo]], colWidths=[width - 38 * mm, 38 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), HexColor(theme.header_bg)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                ("BOX", (0, 0), (-1, -1), 0.5, HexColor(theme.accent)),
            ]
        )
    )
    return table


def build_pdf(theme: Theme, header: dict[str, str], blocks: list[tuple[str, str]]):
    output = OUTPUT / f"{theme.slug}.pdf"
    styles = pdf_styles(theme)
    primary = HexColor(theme.primary)
    soft = HexColor(theme.soft)
    accent = HexColor(theme.accent)

    doc = SimpleDocTemplate(
        str(output),
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=11 * mm,
        bottomMargin=14 * mm,
        title=f'{header["name"]} - {theme.label} CV',
        author=header["name"],
    )
    story = [
        pdf_header(header, theme, styles),
        Spacer(1, 8),
        HRFlowable(width="100%", thickness=1.5, color=primary, spaceAfter=5),
    ]

    in_experience = False
    pending_section: list[Flowable] | None = None

    def append_content(flowable):
        nonlocal pending_section
        if pending_section:
            story.append(KeepTogether([*pending_section, flowable]))
            pending_section = None
        else:
            story.append(flowable)

    for kind, text in blocks:
        if kind == "separator":
            if pending_section is None:
                story.append(Spacer(1, 3))
        elif kind == "section":
            in_experience = text.upper() == "PROFESSIONAL EXPERIENCE"
            pending_section = [Spacer(1, 6), SectionBand(text, theme), Spacer(1, 3)]
        elif kind == "role":
            if theme.layout == "timeline" and in_experience:
                role = Table(
                    [[Paragraph(text, styles["CvRole"])]],
                    colWidths=[doc.width],
                )
                role.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, -1), soft),
                            ("LINEBEFORE", (0, 0), (0, 0), 3, primary),
                            ("LEFTPADDING", (0, 0), (-1, -1), 9),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                            ("TOPPADDING", (0, 0), (-1, -1), 2),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ]
                    )
                )
                append_content(role)
            else:
                append_content(Paragraph(text, styles["CvRole"]))
        elif kind == "subheading":
            append_content(Paragraph(text, styles["CvSub"]))
        elif kind == "bullet":
            append_content(Paragraph(f"- {text}", styles["CvBullet"]))
        elif kind == "paragraph":
            # Degree/institution lines are naturally compact; narrative stays justified.
            append_content(Paragraph(text, styles["CvBody"]))

    def footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(HexColor(theme.primary))
        canvas.setLineWidth(0.5)
        canvas.line(14 * mm, 9 * mm, A4[0] - 14 * mm, 9 * mm)
        canvas.setFillColor(HexColor("#6B7785"))
        canvas.setFont("Helvetica", 7.3)
        canvas.drawString(14 * mm, 5.2 * mm, f'{header["name"]} | {theme.label}')
        canvas.drawRightString(A4[0] - 14 * mm, 5.2 * mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return output


def rgb(hex_value: str) -> RGBColor:
    value = hex_value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def shade_cell(cell, color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color.lstrip("#"))
    tc_pr.append(shd)


def set_cell_margins(cell, top=90, start=90, bottom=90, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def add_page_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])


def set_run(run, font: str, size: float, color: RGBColor, bold=False):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.bold = bold


def build_docx(theme: Theme, header: dict[str, str], blocks: list[tuple[str, str]]):
    output = OUTPUT / f"{theme.slug}.docx"
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.25)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    primary = rgb(theme.primary)
    secondary = rgb(theme.secondary)
    ink = RGBColor(0x25, 0x31, 0x3D)
    muted = RGBColor(0x60, 0x70, 0x80)
    header_color = primary if theme.header_light else RGBColor(255, 255, 255)
    header_secondary = secondary if theme.header_light else RGBColor(0xDC, 0xEA, 0xF0)

    table = doc.add_table(rows=1, cols=2)
    left, right = table.rows[0].cells
    shade_cell(left, theme.header_bg)
    shade_cell(right, theme.header_bg)
    set_cell_margins(left, 140, 140, 140, 140)
    set_cell_margins(right, 140, 70, 140, 100)
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    set_run(p.add_run(header["name"].upper()), theme.heading_font.replace("Helvetica-", "Arial "), 18, header_color, True)
    p = left.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    set_run(p.add_run(header["title"]), "Arial", 9.5, header_secondary)
    for line in (
        header["location"],
        f'Mobile: {header["mobile"]}',
        f'Email: {header["email"]}',
        f'Website: {header["website"]}',
    ):
        p = left.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(line), "Arial", 8.5, header_color if not theme.header_light else muted)

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run().add_picture(str(PHOTO), width=Inches(1.2))

    for style_name, font, size, color, bold in (
        ("Normal", "Arial" if theme.body_font == "Helvetica" else "Times New Roman", 9.5, ink, False),
        ("Title", "Arial", 18, primary, True),
        ("Heading 1", "Arial", 12, primary, True),
        ("Heading 2", "Arial", 11, primary, True),
        ("Heading 3", "Arial", 9.5, secondary, True),
    ):
        style = doc.styles[style_name]
        style.font.name = font
        style._element.rPr.rFonts.set(qn("w:eastAsia"), font)
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = bold

    body_font = "Times New Roman" if theme.body_font == "Times-Roman" else "Arial"
    for kind, text in blocks:
        if kind == "separator":
            continue
        if kind == "section":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(5)
            if theme.layout in {"modern", "editorial"}:
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), theme.primary.lstrip("#"))
                p._p.get_or_add_pPr().append(shade)
                set_run(p.add_run(text.upper()), "Arial", 10.5, RGBColor(255, 255, 255), True)
            else:
                set_run(p.add_run(text.upper()), "Arial", 11.5, primary, True)
                p_bdr = OxmlElement("w:pBdr")
                bottom = OxmlElement("w:bottom")
                bottom.set(qn("w:val"), "single")
                bottom.set(qn("w:sz"), "12")
                bottom.set(qn("w:color"), theme.accent.lstrip("#"))
                p_bdr.append(bottom)
                p._p.get_or_add_pPr().append(p_bdr)
        elif kind == "role":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(7)
            p.paragraph_format.space_after = Pt(2)
            set_run(p.add_run(text), "Arial", 11, primary, True)
        elif kind == "subheading":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(3)
            set_run(p.add_run(text), "Arial", 9.5, secondary, True)
        elif kind == "bullet":
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.05
            set_run(p.add_run(text), body_font, 9.3, ink)
        elif kind == "paragraph":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.08
            set_run(p.add_run(text), body_font, 9.5, ink)

    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(f'{header["name"]} | {theme.label} | Page '), "Arial", 8, muted)
    add_page_field(p)

    doc.save(str(output))
    return output


def render_previews(pdf_paths: list[Path]):
    import fitz

    PREVIEW.mkdir(parents=True, exist_ok=True)
    for old in PREVIEW.glob("*.png"):
        old.unlink()
    for pdf_path in pdf_paths:
        pdf = fitz.open(str(pdf_path))
        for index, page in enumerate(pdf, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
            pix.save(str(PREVIEW / f"{pdf_path.stem}_page_{index}.png"))
        print(f"{pdf_path.name}: {pdf.page_count} pages")
        pdf.close()


def verify_source_coverage(files: list[Path], blocks: list[tuple[str, str]]):
    """Check high-value source phrases exist in every generated deliverable."""
    import fitz

    required = [
        "2,000",
        "50+",
        "FairBanks Social Enterprise Initiative",
        "Machine Learning-Based Prediction",
        "Federation of Uganda Employers",
        "Human Resource Assistant",
        "Health Management Systems Administrator",
    ]
    for path in files:
        if path.suffix == ".pdf":
            document = fitz.open(str(path))
            text = "\n".join(page.get_text() for page in document)
            document.close()
        else:
            document = Document(str(path))
            text = "\n".join(p.text for p in document.paragraphs)
            text += "\n" + "\n".join(
                cell.text
                for table in document.tables
                for row in table.rows
                for cell in row.cells
            )
        missing = [phrase for phrase in required if phrase not in text]
        if missing:
            raise RuntimeError(f"{path.name} missing source content: {missing}")


def main():
    if not SOURCE.exists():
        raise SystemExit(f"Missing source of truth: {SOURCE}")
    if not PHOTO.exists():
        raise SystemExit(f"Missing portrait: {PHOTO}")
    OUTPUT.mkdir(parents=True, exist_ok=True)
    for old in OUTPUT.iterdir():
        if old.is_file():
            old.unlink()

    header, blocks = parse_source()
    generated: list[Path] = []
    pdf_paths: list[Path] = []
    for theme in THEMES:
        pdf = build_pdf(theme, header, blocks)
        docx = build_docx(theme, header, blocks)
        generated.extend([pdf, docx])
        pdf_paths.append(pdf)
    verify_source_coverage(generated, blocks)
    render_previews(pdf_paths)
    print(f"Generated {len(generated)} detailed CV files in {OUTPUT}")


if __name__ == "__main__":
    main()
