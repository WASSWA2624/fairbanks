# -*- coding: utf-8 -*-
"""
Build the general-purpose CV for Wasswa Wilson in both Word (.docx) and PDF.

    python build_wilson_cv.py

Content lives in cv_content.py so the two documents cannot drift apart.
House palette matches the other FairBanks / Wilson application documents:
navy #0F2C4C, teal #1A6B5C.
"""

import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import docx
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether,
                                PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

import cv_content as C

HERE = os.path.dirname(os.path.abspath(__file__))
DOCX_OUT = os.path.join(HERE, "Wasswa_Wilson_CV_2026.docx")
PDF_OUT = os.path.join(HERE, "Wasswa_Wilson_CV_2026.pdf")

# ---------------------------------------------------------------- palette --
NAVY_HEX = "0F2C4C"
TEAL_HEX = "1A6B5C"
INK_HEX = "23272B"
MUTED_HEX = "5A6472"
RULE_HEX = "D7DEE6"
BAND_HEX = "F2F5F8"

NAVY_RGB = RGBColor(0x0F, 0x2C, 0x4C)
TEAL_RGB = RGBColor(0x1A, 0x6B, 0x5C)
INK_RGB = RGBColor(0x23, 0x27, 0x2B)
MUTED_RGB = RGBColor(0x5A, 0x64, 0x72)

NAVY = colors.HexColor("#" + NAVY_HEX)
TEAL = colors.HexColor("#" + TEAL_HEX)
INK = colors.HexColor("#" + INK_HEX)
MUTED = colors.HexColor("#" + MUTED_HEX)
RULE = colors.HexColor("#" + RULE_HEX)
BAND = colors.HexColor("#" + BAND_HEX)

BODY_FONT = "Calibri"
BULLET = "•"


# ==========================================================================
#  WORD
# ==========================================================================

def _shade(el, hex_fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    el.append(shd)


def cell_shade(cell, hex_fill):
    _shade(cell._tc.get_or_add_tcPr(), hex_fill)


def para_border(par, edge="bottom", hex_color=TEAL_HEX, size=8, space=2):
    pPr = par._p.get_or_add_pPr()
    borders = pPr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        pPr.append(borders)
    e = OxmlElement("w:" + edge)
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(size))
    e.set(qn("w:space"), str(space))
    e.set(qn("w:color"), hex_color)
    borders.append(e)


def set_cell_width(cell, inches):
    cell.width = Inches(inches)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(inches * 1440)))
    tcW.set(qn("w:type"), "dxa")


def no_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        borders.append(e)
    tblPr.append(borders)


def keep_next(par):
    par.paragraph_format.keep_with_next = True


def run(par, text, size=10, bold=False, italic=False, color=INK_RGB,
        font=BODY_FONT, caps=False, spacing=None):
    r = par.add_run(text.upper() if caps else text)
    r.font.name = font
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color
    rPr = r._element.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), font)
    rf.set(qn("w:hAnsi"), font)
    rf.set(qn("w:cs"), font)
    rPr.append(rf)
    if spacing:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(spacing))
        rPr.append(sp)
    return r


def para(container, space_before=0, space_after=4, align=None, line=1.06,
         left=0.0, hanging=0.0):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    if left:
        pf.left_indent = Inches(left)
    if hanging:
        pf.first_line_indent = Inches(-hanging)
    return p


def section(doc, title):
    p = para(doc, space_before=11, space_after=5)
    keep_next(p)
    run(p, title, size=10.5, bold=True, color=NAVY_RGB, caps=True, spacing=24)
    para_border(p, "bottom", TEAL_HEX, size=8, space=3)
    return p


def bullet(doc, text, size=9.8, left=0.20, color=INK_RGB):
    p = para(doc, space_after=2.5, left=left, hanging=0.16, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, BULLET + "   ", size=size, color=TEAL_RGB, bold=True)
    run(p, text, size=size, color=color)
    return p


def label_table(doc, rows, label_w=1.62, value_w=5.58, band=True, size=9.6):
    """Two-column label/value table used for skills, competencies, equipment."""
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    no_table_borders(t)
    for i, (label, value) in enumerate(rows):
        cells = t.add_row().cells
        set_cell_width(cells[0], label_w)
        set_cell_width(cells[1], value_w)
        if band and i % 2 == 0:
            cell_shade(cells[0], BAND_HEX)
            cell_shade(cells[1], BAND_HEX)

        c0 = cells[0].paragraphs[0]
        c0.paragraph_format.space_before = Pt(3)
        c0.paragraph_format.space_after = Pt(3)
        c0.paragraph_format.line_spacing = 1.02
        run(c0, label, size=size, bold=True, color=NAVY_RGB)

        c1 = cells[1].paragraphs[0]
        c1.paragraph_format.space_before = Pt(3)
        c1.paragraph_format.space_after = Pt(3)
        c1.paragraph_format.line_spacing = 1.06
        run(c1, value, size=size, color=INK_RGB)
    return t


def add_page_footer(section_obj):
    footer = section_obj.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    run(p, "Wasswa Wilson  ·  Curriculum Vitae  ·  ", size=8, color=MUTED_RGB)
    r = p.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    r._r.append(fld)
    r2 = p.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    r2._r.append(instr)
    r3 = p.add_run()
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    r3._r.append(fld2)
    for rr in (r, r2, r3):
        rr.font.size = Pt(8)
        rr.font.name = BODY_FONT
        rr.font.color.rgb = MUTED_RGB


def build_docx():
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = BODY_FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

    sec = doc.sections[0]
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.52)
    sec.bottom_margin = Inches(0.52)
    sec.left_margin = Inches(0.58)
    sec.right_margin = Inches(0.58)
    add_page_footer(sec)

    # ---- header ----------------------------------------------------------
    p = para(doc, space_after=1, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, C.NAME, size=23, bold=True, color=NAVY_RGB, spacing=36)

    p = para(doc, space_after=5, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, C.TAGLINE, size=10, bold=True, color=TEAL_RGB)

    for i, line in enumerate(C.CONTACT_LINES):
        p = para(doc, space_after=1.5 if i < len(C.CONTACT_LINES) - 1 else 0,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
        run(p, line, size=9, color=MUTED_RGB)
    para_border(p, "bottom", NAVY_HEX, size=12, space=6)

    # ---- profile ---------------------------------------------------------
    section(doc, "Professional profile")
    p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, C.PROFILE, size=9.9)

    # ---- highlights ------------------------------------------------------
    section(doc, "Signature achievements")
    for h in C.HIGHLIGHTS:
        bullet(doc, h)

    # ---- skills ----------------------------------------------------------
    section(doc, "Technical skills")
    label_table(doc, C.SKILLS)

    # ---- experience ------------------------------------------------------
    section(doc, "Professional experience")
    for i, job in enumerate(C.EXPERIENCE):
        head = para(doc, space_before=7 if i else 2, space_after=0)
        keep_next(head)
        head.paragraph_format.tab_stops.add_tab_stop(
            Inches(7.11), WD_TAB_ALIGNMENT.RIGHT)
        run(head, job["role"], size=10.4, bold=True, color=NAVY_RGB)
        run(head, "\t" + job["dates"], size=9.3, bold=True, color=TEAL_RGB)

        sub = para(doc, space_after=3)
        keep_next(sub)
        run(sub, job["org"], size=9.4, italic=True, color=MUTED_RGB)

        for b in job["bullets"]:
            bullet(doc, b)

    # ---- projects --------------------------------------------------------
    section(doc, "Selected projects and systems delivered")
    for i, (name, role, org, period, detail) in enumerate(C.PROJECTS):
        head = para(doc, space_before=6 if i else 2, space_after=0)
        keep_next(head)
        head.paragraph_format.tab_stops.add_tab_stop(
            Inches(7.11), WD_TAB_ALIGNMENT.RIGHT)
        run(head, name, size=10, bold=True, color=NAVY_RGB)
        run(head, "\t" + period, size=9.3, bold=True, color=TEAL_RGB)

        sub = para(doc, space_after=2)
        keep_next(sub)
        run(sub, role + "  —  " + org, size=9.2, italic=True, color=MUTED_RGB)

        p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        run(p, detail, size=9.7)

    # ---- competencies ----------------------------------------------------
    section(doc, "Professional competencies")
    label_table(doc, C.COMPETENCIES, label_w=2.05, value_w=5.15)

    # ---- equipment -------------------------------------------------------
    section(doc, "Medical equipment and clinical technology experience")
    label_table(doc, C.EQUIPMENT, label_w=1.72, value_w=5.48)

    # ---- education -------------------------------------------------------
    section(doc, "Education")
    for i, (award, inst, year, note) in enumerate(C.EDUCATION):
        head = para(doc, space_before=6 if i else 2, space_after=0)
        keep_next(head)
        head.paragraph_format.tab_stops.add_tab_stop(
            Inches(7.11), WD_TAB_ALIGNMENT.RIGHT)
        run(head, award, size=10, bold=True, color=NAVY_RGB)
        run(head, "\t" + year, size=9.3, bold=True, color=TEAL_RGB)

        sub = para(doc, space_after=2 if note else 0)
        keep_next(sub)
        run(sub, inst, size=9.3, italic=True, color=MUTED_RGB)

        if note:
            p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            run(p, note, size=9.4, color=INK_RGB)

    # ---- training --------------------------------------------------------
    section(doc, "Professional training and certification")
    for course, provider, year in C.TRAINING:
        p = para(doc, space_after=3, left=0.20, hanging=0.16)
        run(p, BULLET + "   ", size=9.7, color=TEAL_RGB, bold=True)
        run(p, course, size=9.7, bold=True, color=INK_RGB)
        run(p, "  —  " + provider + ", " + year, size=9.7, color=MUTED_RGB)

    # ---- interests -------------------------------------------------------
    section(doc, "Interests")
    p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, C.INTERESTS, size=9.7)

    # ---- referees --------------------------------------------------------
    section(doc, "Referees")
    t = doc.add_table(rows=1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    no_table_borders(t)
    for idx, (name, role, contact) in enumerate(C.REFEREES):
        cell = t.rows[0].cells[idx]
        set_cell_width(cell, 2.40)
        p0 = cell.paragraphs[0]
        p0.paragraph_format.space_after = Pt(1)
        p0.paragraph_format.line_spacing = 1.05
        run(p0, name, size=9.6, bold=True, color=NAVY_RGB)
        p1 = cell.add_paragraph()
        p1.paragraph_format.space_after = Pt(1)
        p1.paragraph_format.line_spacing = 1.05
        run(p1, role, size=9, italic=True, color=MUTED_RGB)
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        run(p2, contact, size=9, color=INK_RGB)

    doc.save(DOCX_OUT)
    return DOCX_OUT


# ==========================================================================
#  PDF
# ==========================================================================

PW, PH = A4
LM = RM = 0.58 * inch
TM = 0.52 * inch
BM = 0.58 * inch
CW = PW - LM - RM


def pdf_styles():
    s = {}
    s["name"] = ParagraphStyle("name", fontName="Helvetica-Bold", fontSize=22.5,
                               leading=25, textColor=NAVY, alignment=1,
                               spaceAfter=2)
    s["tagline"] = ParagraphStyle("tagline", fontName="Helvetica-Bold", fontSize=9.8,
                                  leading=12.5, textColor=TEAL, alignment=1,
                                  spaceAfter=5)
    s["contact"] = ParagraphStyle("contact", fontName="Helvetica", fontSize=8.8,
                                  leading=11.6, textColor=MUTED, alignment=1)
    s["section"] = ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=10.2,
                                  leading=12, textColor=NAVY, spaceBefore=0,
                                  spaceAfter=0)
    s["body"] = ParagraphStyle("body", fontName="Helvetica", fontSize=9.4,
                               leading=12.4, textColor=INK, alignment=TA_JUSTIFY)
    s["bullet"] = ParagraphStyle("bullet", fontName="Helvetica", fontSize=9.3,
                                 leading=12.2, textColor=INK, alignment=TA_JUSTIFY,
                                 leftIndent=13, bulletIndent=1, spaceAfter=3.2,
                                 bulletFontName="Helvetica-Bold", bulletFontSize=9.3,
                                 bulletColor=TEAL)
    s["role"] = ParagraphStyle("role", fontName="Helvetica-Bold", fontSize=10,
                               leading=12.4, textColor=NAVY)
    s["dates"] = ParagraphStyle("dates", fontName="Helvetica-Bold", fontSize=8.9,
                                leading=12.4, textColor=TEAL, alignment=TA_RIGHT)
    s["org"] = ParagraphStyle("org", fontName="Helvetica-Oblique", fontSize=8.9,
                              leading=11.2, textColor=MUTED, spaceAfter=2.5)
    s["cellL"] = ParagraphStyle("cellL", fontName="Helvetica-Bold", fontSize=9.1,
                                leading=11.6, textColor=NAVY)
    s["cellR"] = ParagraphStyle("cellR", fontName="Helvetica", fontSize=9.1,
                                leading=11.8, textColor=INK)
    s["refname"] = ParagraphStyle("refname", fontName="Helvetica-Bold", fontSize=9.2,
                                  leading=11.4, textColor=NAVY)
    s["refrole"] = ParagraphStyle("refrole", fontName="Helvetica-Oblique", fontSize=8.6,
                                  leading=10.8, textColor=MUTED)
    s["refcon"] = ParagraphStyle("refcon", fontName="Helvetica", fontSize=8.6,
                                 leading=10.8, textColor=INK)
    return s


class Rule(Spacer):
    """A horizontal rule flowable."""

    def __init__(self, width, thickness=0.6, color=RULE, above=0, below=0):
        Spacer.__init__(self, width, thickness + above + below)
        self.rw = width
        self.thickness = thickness
        self.color = color
        self.above = above
        self.below = below

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        y = self.below
        self.canv.line(0, y, self.rw, y)


def pdf_section(story, s, title):
    """Heading + underline as ONE atomic flowable that cannot be split from the
    content that follows it."""
    t = Table([[Paragraph(_spaced_caps(title), s["section"])]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
        ("LINEBELOW", (0, 0), (-1, -1), 0.9, TEAL),
    ]))
    t.keepWithNext = 1
    story.append(Spacer(1, 9))
    story.append(t)
    story.append(Spacer(1, 5))


def _spaced_caps(text):
    """Letter-spaced small caps look, done by hand (reportlab has no tracking)."""
    return "&nbsp;".join(text.upper())


def head_row(s, left_text, right_text, left_style="role"):
    t = Table([[Paragraph(left_text, s[left_style]),
                Paragraph(right_text, s["dates"])]],
              colWidths=[CW - 1.55 * inch, 1.55 * inch])
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
    ]))
    return t


def pdf_label_table(s, rows, label_w=1.62):
    data = [[Paragraph(a, s["cellL"]), Paragraph(b, s["cellR"])] for a, b in rows]
    t = Table(data, colWidths=[label_w * inch, CW - label_w * inch], hAlign="LEFT")
    style = [
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    for i in range(len(rows)):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), BAND))
    t.setStyle(TableStyle(style))
    return t


def on_page(canv, doc):
    canv.saveState()
    canv.setFont("Helvetica", 7.8)
    canv.setFillColor(MUTED)
    canv.drawCentredString(PW / 2.0, 0.34 * inch,
                           "Wasswa Wilson  ·  Curriculum Vitae  ·  %d" % canv.getPageNumber())
    canv.restoreState()


def build_pdf():
    s = pdf_styles()
    doc = BaseDocTemplate(PDF_OUT, pagesize=A4,
                          leftMargin=LM, rightMargin=RM,
                          topMargin=TM, bottomMargin=BM,
                          title="Wasswa Wilson - Curriculum Vitae",
                          author="Wasswa Wilson",
                          subject="Curriculum Vitae")
    frame = Frame(LM, BM, CW, PH - TM - BM, id="body",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])

    story = []

    # ---- header ----------------------------------------------------------
    story.append(Paragraph(_spaced_caps(C.NAME), s["name"]))
    story.append(Paragraph(C.TAGLINE, s["tagline"]))
    for line in C.CONTACT_LINES:
        story.append(Paragraph(line, s["contact"]))
    story.append(Spacer(1, 6))
    story.append(Rule(CW, thickness=1.4, color=NAVY))

    # ---- profile ---------------------------------------------------------
    pdf_section(story, s, "Professional profile")
    story.append(Paragraph(C.PROFILE, s["body"]))

    # ---- highlights ------------------------------------------------------
    pdf_section(story, s, "Signature achievements")
    for h in C.HIGHLIGHTS:
        story.append(Paragraph(h, s["bullet"],
                               bulletText=BULLET))

    # ---- skills ----------------------------------------------------------
    pdf_section(story, s, "Technical skills")
    story.append(pdf_label_table(s, C.SKILLS, label_w=1.62))

    # ---- experience ------------------------------------------------------
    pdf_section(story, s, "Professional experience")
    for i, job in enumerate(C.EXPERIENCE):
        block = [head_row(s, job["role"], job["dates"]),
                 Spacer(1, 1.5),
                 Paragraph(job["org"], s["org"]),
                 Paragraph(job["bullets"][0], s["bullet"],
                           bulletText=BULLET)]
        story.append(Spacer(1, 6 if i else 0))
        story.append(KeepTogether(block))
        for b in job["bullets"][1:]:
            story.append(Paragraph(b, s["bullet"],
                                   bulletText=BULLET))

    # ---- projects --------------------------------------------------------
    pdf_section(story, s, "Selected projects and systems delivered")
    for i, (name, role, org, period, detail) in enumerate(C.PROJECTS):
        block = [head_row(s, name, period),
                 Spacer(1, 1.5),
                 Paragraph(role + "  —  " + org, s["org"]),
                 Paragraph(detail, s["body"])]
        story.append(Spacer(1, 6 if i else 0))
        story.append(KeepTogether(block))

    # ---- competencies ----------------------------------------------------
    pdf_section(story, s, "Professional competencies")
    story.append(pdf_label_table(s, C.COMPETENCIES, label_w=2.05))

    # ---- equipment -------------------------------------------------------
    pdf_section(story, s, "Medical equipment and clinical technology experience")
    story.append(pdf_label_table(s, C.EQUIPMENT, label_w=1.72))

    # ---- education -------------------------------------------------------
    pdf_section(story, s, "Education")
    for i, (award, inst, year, note) in enumerate(C.EDUCATION):
        block = [head_row(s, award, year), Spacer(1, 1.5),
                 Paragraph(inst, s["org"])]
        if note:
            block.append(Paragraph(note, s["body"]))
        story.append(Spacer(1, 6 if i else 0))
        story.append(KeepTogether(block))

    # ---- training --------------------------------------------------------
    pdf_section(story, s, "Professional training and certification")
    for course, provider, year in C.TRAINING:
        story.append(Paragraph(
            "<b>%s</b>  <font color='#%s'>—  %s, %s</font>" % (course, MUTED_HEX, provider, year),
            s["bullet"],
            bulletText=BULLET))

    # ---- interests -------------------------------------------------------
    pdf_section(story, s, "Interests")
    story.append(Paragraph(C.INTERESTS, s["body"]))

    # ---- referees --------------------------------------------------------
    pdf_section(story, s, "Referees")
    cells = []
    for name, role, contact in C.REFEREES:
        cells.append([Paragraph(name, s["refname"]),
                      Paragraph(role, s["refrole"]),
                      Paragraph(contact, s["refcon"])])
    inner = []
    for c in cells:
        t = Table([[x] for x in c], colWidths=[CW / 3.0 - 10])
        t.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0.6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        inner.append(t)
    outer = Table([inner], colWidths=[CW / 3.0] * 3)
    outer.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("LEFTPADDING", (1, 0), (-1, 0), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(KeepTogether(outer))

    doc.build(story)
    return PDF_OUT


if __name__ == "__main__":
    d = build_docx()
    print("DOCX ->", d)
    p = build_pdf()
    print("PDF  ->", p)
