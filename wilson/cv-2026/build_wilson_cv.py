# -*- coding: utf-8 -*-
"""
Build the general-purpose CV for Wasswa Wilson in both Word (.docx) and PDF.

    python build_wilson_cv.py

Content lives in cv_content.py so the two documents cannot drift apart.

Design notes
------------
* Navy masthead with the name reversed out in white, capped by a teal bar.
* Section headings carry a teal left accent bar and a hairline underline.
* Career highlights sit in a tinted panel so they read first.
* Referees are laid out as three tinted cards.
* No em dashes anywhere, in the chrome or the content.
"""

import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_RIGHT
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
PALE_HEX = "9FD6C8"      # tagline on navy
INK_HEX = "23272B"
MUTED_HEX = "5A6472"
HAIR_HEX = "CFD8E3"      # hairline under section headings
BAND_HEX = "F3F6F9"      # zebra rows and panels

NAVY_RGB = RGBColor(0x0F, 0x2C, 0x4C)
TEAL_RGB = RGBColor(0x1A, 0x6B, 0x5C)
PALE_RGB = RGBColor(0x9F, 0xD6, 0xC8)
INK_RGB = RGBColor(0x23, 0x27, 0x2B)
MUTED_RGB = RGBColor(0x5A, 0x64, 0x72)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)

NAVY = colors.HexColor("#" + NAVY_HEX)
TEAL = colors.HexColor("#" + TEAL_HEX)
PALE = colors.HexColor("#" + PALE_HEX)
INK = colors.HexColor("#" + INK_HEX)
MUTED = colors.HexColor("#" + MUTED_HEX)
HAIR = colors.HexColor("#" + HAIR_HEX)
BAND = colors.HexColor("#" + BAND_HEX)

BODY_FONT = "Calibri"
BULLET = "•"

CONTENT_W = 7.11  # inches, A4 minus margins


# ==========================================================================
#  WORD helpers
# ==========================================================================

def cell_shade(cell, hex_fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    cell._tc.get_or_add_tcPr().append(shd)


def cell_border(cell, edge, hex_color, size):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    e = OxmlElement("w:" + edge)
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(size))
    e.set(qn("w:space"), "0")
    e.set(qn("w:color"), hex_color)
    borders.append(e)


def para_border(par, edge, hex_color, size=8, space=2):
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


def set_cell_margins(table, top=0, left=0, bottom=0, right=0):
    """Margins in twips (1440 per inch)."""
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement("w:" + tag)
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    table._tbl.tblPr.append(mar)


def no_table_borders(table):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        borders.append(e)
    table._tbl.tblPr.append(borders)


def drop_first_para(cell):
    p = cell.paragraphs[0]
    p._element.getparent().remove(p._element)


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


def para(container, space_before=0, space_after=4, align=None, line=1.02,
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
    """Teal left accent bar + navy caps + hairline underline."""
    p = para(doc, space_before=9, space_after=4.5, left=0.11)
    keep_next(p)
    run(p, title, size=10.5, bold=True, color=NAVY_RGB, caps=True, spacing=26)
    para_border(p, "left", TEAL_HEX, size=22, space=6)
    para_border(p, "bottom", HAIR_HEX, size=4, space=4)
    return p


def bullet(container, text, size=9.6, left=0.20, color=INK_RGB, after=2.2):
    p = para(container, space_after=after, left=left, hanging=0.16,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, BULLET + "   ", size=size, color=TEAL_RGB, bold=True)
    run(p, text, size=size, color=color)
    return p


def label_table(doc, rows, label_w=1.66, size=9.5):
    value_w = CONTENT_W - label_w
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    no_table_borders(t)
    set_cell_margins(t, top=42, left=85, bottom=42, right=85)
    for i, (label, value) in enumerate(rows):
        cells = t.add_row().cells
        set_cell_width(cells[0], label_w)
        set_cell_width(cells[1], value_w)
        if i % 2 == 0:
            cell_shade(cells[0], BAND_HEX)
            cell_shade(cells[1], BAND_HEX)
        c0 = cells[0].paragraphs[0]
        c0.paragraph_format.space_after = Pt(0)
        c0.paragraph_format.line_spacing = 1.04
        run(c0, label, size=size, bold=True, color=NAVY_RGB)
        c1 = cells[1].paragraphs[0]
        c1.paragraph_format.space_after = Pt(0)
        c1.paragraph_format.line_spacing = 1.08
        run(c1, value, size=size, color=INK_RGB)
    return t


def entry_head(doc, title, right, first=False, title_size=10.2):
    p = para(doc, space_before=0 if first else 6, space_after=0)
    keep_next(p)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    run(p, title, size=title_size, bold=True, color=NAVY_RGB)
    run(p, "\t" + right, size=9.2, bold=True, color=TEAL_RGB)
    return p


def entry_sub(doc, text, after=3.0):
    p = para(doc, space_after=after)
    keep_next(p)
    run(p, text, size=9.2, italic=True, color=MUTED_RGB)
    return p


def add_page_footer(sec):
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, "Wasswa Wilson  ·  Curriculum Vitae  ·  ", size=8, color=MUTED_RGB)
    parts = []
    for kind, payload in (("begin", None), ("instr", "PAGE"), ("end", None)):
        r = p.add_run()
        if kind == "instr":
            e = OxmlElement("w:instrText")
            e.set(qn("xml:space"), "preserve")
            e.text = payload
        else:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), kind)
        r._r.append(e)
        parts.append(r)
    for r in parts:
        r.font.size = Pt(8)
        r.font.name = BODY_FONT
        r.font.color.rgb = MUTED_RGB


# ==========================================================================
#  WORD document
# ==========================================================================

def build_docx():
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = BODY_FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

    sec = doc.sections[0]
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.58)
    sec.right_margin = Inches(0.58)
    add_page_footer(sec)

    # ---- masthead --------------------------------------------------------
    mast = doc.add_table(rows=1, cols=1)
    mast.alignment = WD_TABLE_ALIGNMENT.LEFT
    mast.autofit = False
    no_table_borders(mast)
    set_cell_margins(mast, top=170, left=120, bottom=170, right=120)
    cell = mast.rows[0].cells[0]
    set_cell_width(cell, CONTENT_W)
    cell_shade(cell, NAVY_HEX)
    cell_border(cell, "bottom", TEAL_HEX, 26)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.0
    run(p, C.NAME, size=23, bold=True, color=WHITE_RGB, spacing=48)

    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run(p, C.TAGLINE, size=9.5, bold=True, color=PALE_RGB)

    for i, line in enumerate(C.CONTACT_LINES):
        p = para(doc, space_before=6 if i == 0 else 0,
                 space_after=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        run(p, line, size=9.2, color=MUTED_RGB)

    # ---- profile ---------------------------------------------------------
    section(doc, "Professional profile")
    p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, C.PROFILE, size=9.8)

    # ---- career highlights (tinted panel) --------------------------------
    section(doc, "Career highlights")
    panel = doc.add_table(rows=1, cols=1)
    panel.alignment = WD_TABLE_ALIGNMENT.LEFT
    panel.autofit = False
    no_table_borders(panel)
    set_cell_margins(panel, top=110, left=150, bottom=110, right=130)
    pc = panel.rows[0].cells[0]
    set_cell_width(pc, CONTENT_W)
    cell_shade(pc, BAND_HEX)
    cell_border(pc, "left", TEAL_HEX, 26)
    for i, h in enumerate(C.HIGHLIGHTS):
        bullet(pc, h, size=9.6, left=0.18,
               after=0 if i == len(C.HIGHLIGHTS) - 1 else 3.4)
    drop_first_para(pc)

    # ---- skills ----------------------------------------------------------
    section(doc, "Technical skills")
    label_table(doc, C.SKILLS, label_w=1.66)

    # ---- experience ------------------------------------------------------
    section(doc, "Professional experience")
    for i, job in enumerate(C.EXPERIENCE):
        entry_head(doc, job["role"], job["dates"], first=(i == 0), title_size=10.4)
        entry_sub(doc, job["org"])
        for b in job["bullets"]:
            bullet(doc, b)

    # ---- projects --------------------------------------------------------
    section(doc, "Selected projects")
    for i, (name, role, org, period, detail) in enumerate(C.PROJECTS):
        entry_head(doc, name, period, first=(i == 0), title_size=10)
        entry_sub(doc, role + "  ·  " + org, after=2.0)
        p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        run(p, detail, size=9.6)

    # ---- competencies ----------------------------------------------------
    section(doc, "Professional competencies")
    label_table(doc, C.COMPETENCIES, label_w=2.05)

    # ---- equipment -------------------------------------------------------
    section(doc, "Medical equipment experience")
    label_table(doc, C.EQUIPMENT, label_w=1.76)

    # ---- education -------------------------------------------------------
    section(doc, "Education")
    for i, (award, inst, year, note) in enumerate(C.EDUCATION):
        entry_head(doc, award, year, first=(i == 0), title_size=10)
        entry_sub(doc, inst, after=2.0 if note else 0.0)
        if note:
            p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            run(p, note, size=9.4)

    # ---- training --------------------------------------------------------
    section(doc, "Training and certification")
    for course, provider, year in C.TRAINING:
        p = para(doc, space_after=3.4, left=0.20, hanging=0.16)
        run(p, BULLET + "   ", size=9.7, color=TEAL_RGB, bold=True)
        run(p, course, size=9.7, bold=True, color=INK_RGB)
        run(p, ".  " + provider + ", " + year + ".", size=9.7, color=MUTED_RGB)

    # ---- interests -------------------------------------------------------
    section(doc, "Interests")
    p = para(doc, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    run(p, C.INTERESTS, size=9.7)

    # ---- referees (three tinted cards) -----------------------------------
    section(doc, "Referees")
    card_w, gap_w = 2.27, 0.15
    t = doc.add_table(rows=1, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    no_table_borders(t)
    set_cell_margins(t, top=90, left=110, bottom=90, right=110)
    cells = t.rows[0].cells
    for idx, (name, role, contact) in enumerate(C.REFEREES):
        if idx:
            set_cell_width(cells[idx * 2 - 1], gap_w)
        cell = cells[idx * 2]
        set_cell_width(cell, card_w)
        cell_shade(cell, BAND_HEX)
        cell_border(cell, "top", TEAL_HEX, 18)
        p0 = cell.paragraphs[0]
        p0.paragraph_format.space_after = Pt(1.5)
        p0.paragraph_format.line_spacing = 1.05
        run(p0, name, size=9.5, bold=True, color=NAVY_RGB)
        p1 = cell.add_paragraph()
        p1.paragraph_format.space_after = Pt(1.5)
        p1.paragraph_format.line_spacing = 1.05
        run(p1, role, size=8.8, italic=True, color=MUTED_RGB)
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        run(p2, contact, size=8.8, color=INK_RGB)

    doc.save(DOCX_OUT)
    return DOCX_OUT


# ==========================================================================
#  PDF
# ==========================================================================

PW, PH = A4
LM = RM = 0.58 * inch
TM = 0.5 * inch
BM = 0.56 * inch
CW = PW - LM - RM

NOPAD = [("LEFTPADDING", (0, 0), (-1, -1), 0),
         ("RIGHTPADDING", (0, 0), (-1, -1), 0),
         ("TOPPADDING", (0, 0), (-1, -1), 0),
         ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]


def pdf_styles():
    s = {}
    s["name"] = ParagraphStyle("name", fontName="Helvetica-Bold", fontSize=23,
                               leading=26, textColor=colors.white, alignment=1,
                               spaceAfter=3)
    s["tagline"] = ParagraphStyle("tagline", fontName="Helvetica-Bold", fontSize=9.4,
                                  leading=12, textColor=PALE, alignment=1)
    s["contact"] = ParagraphStyle("contact", fontName="Helvetica", fontSize=9,
                                  leading=12, textColor=MUTED, alignment=1)
    s["section"] = ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=10.2,
                                  leading=12.4, textColor=NAVY)
    s["body"] = ParagraphStyle("body", fontName="Helvetica", fontSize=9.25,
                               leading=11.6, textColor=INK, alignment=TA_JUSTIFY)
    s["bullet"] = ParagraphStyle("bullet", fontName="Helvetica", fontSize=9.15,
                                 leading=11.5, textColor=INK, alignment=TA_JUSTIFY,
                                 leftIndent=13, bulletIndent=1, spaceAfter=2.4,
                                 bulletFontName="Helvetica-Bold", bulletFontSize=9.2,
                                 bulletColor=TEAL)
    s["panel"] = ParagraphStyle("panel", parent=s["bullet"], spaceAfter=3.6)
    s["role"] = ParagraphStyle("role", fontName="Helvetica-Bold", fontSize=10,
                               leading=12.4, textColor=NAVY)
    s["dates"] = ParagraphStyle("dates", fontName="Helvetica-Bold", fontSize=8.8,
                                leading=12.4, textColor=TEAL, alignment=TA_RIGHT)
    s["org"] = ParagraphStyle("org", fontName="Helvetica-Oblique", fontSize=8.8,
                              leading=11, textColor=MUTED, spaceAfter=3)
    s["cellL"] = ParagraphStyle("cellL", fontName="Helvetica-Bold", fontSize=9,
                                leading=11.5, textColor=NAVY)
    s["cellR"] = ParagraphStyle("cellR", fontName="Helvetica", fontSize=9,
                                leading=11.7, textColor=INK)
    s["refname"] = ParagraphStyle("refname", fontName="Helvetica-Bold", fontSize=9.1,
                                  leading=11.2, textColor=NAVY, spaceAfter=1.5)
    s["refrole"] = ParagraphStyle("refrole", fontName="Helvetica-Oblique", fontSize=8.4,
                                  leading=10.4, textColor=MUTED, spaceAfter=1.5)
    s["refcon"] = ParagraphStyle("refcon", fontName="Helvetica", fontSize=8.4,
                                 leading=10.4, textColor=INK)
    return s


def _spaced_caps(text):
    """Letter-spaced caps, done by hand since reportlab has no tracking."""
    return "&nbsp;".join(text.upper())


def pdf_section(story, s, title):
    """Teal left bar + navy caps + hairline underline, as one atomic flowable
    that stays glued to whatever follows it."""
    t = Table([[Paragraph(_spaced_caps(title), s["section"])]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBEFORE", (0, 0), (0, 0), 2.6, TEAL),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, HAIR),
    ]))
    t.keepWithNext = 1
    story.append(Spacer(1, 7))
    story.append(t)
    story.append(Spacer(1, 4))


def head_row(s, left_text, right_text):
    t = Table([[Paragraph(left_text, s["role"]), Paragraph(right_text, s["dates"])]],
              colWidths=[CW - 1.5 * inch, 1.5 * inch])
    t.setStyle(TableStyle(NOPAD + [("VALIGN", (0, 0), (-1, -1), "BOTTOM")]))
    return t


def pdf_label_table(s, rows, label_w):
    data = [[Paragraph(a, s["cellL"]), Paragraph(b, s["cellR"])] for a, b in rows]
    t = Table(data, colWidths=[label_w * inch, CW - label_w * inch], hAlign="LEFT")
    style = [("LEFTPADDING", (0, 0), (-1, -1), 6),
             ("RIGHTPADDING", (0, 0), (-1, -1), 6),
             ("TOPPADDING", (0, 0), (-1, -1), 3.3),
             ("BOTTOMPADDING", (0, 0), (-1, -1), 3.3),
             ("VALIGN", (0, 0), (-1, -1), "TOP")]
    for i in range(len(rows)):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), BAND))
    t.setStyle(TableStyle(style))
    return t


def on_page(canv, doc):
    canv.saveState()
    canv.setFont("Helvetica", 7.8)
    canv.setFillColor(MUTED)
    canv.drawCentredString(
        PW / 2.0, 0.32 * inch,
        "Wasswa Wilson  ·  Curriculum Vitae  ·  %d" % canv.getPageNumber())
    canv.restoreState()


def build_pdf():
    s = pdf_styles()
    doc = BaseDocTemplate(PDF_OUT, pagesize=A4,
                          leftMargin=LM, rightMargin=RM,
                          topMargin=TM, bottomMargin=BM,
                          title="Wasswa Wilson, Curriculum Vitae",
                          author="Wasswa Wilson", subject="Curriculum Vitae")
    doc.addPageTemplates([PageTemplate(
        id="all",
        frames=[Frame(LM, BM, CW, PH - TM - BM, id="body",
                      leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)],
        onPage=on_page)])

    story = []

    # ---- masthead --------------------------------------------------------
    mast = Table([[[Paragraph(_spaced_caps(C.NAME), s["name"]),
                    Paragraph(C.TAGLINE, s["tagline"])]]], colWidths=[CW])
    mast.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LINEBELOW", (0, 0), (-1, -1), 3.2, TEAL),
    ]))
    story.append(mast)
    story.append(Spacer(1, 7))
    for line in C.CONTACT_LINES:
        story.append(Paragraph(line, s["contact"]))

    # ---- profile ---------------------------------------------------------
    pdf_section(story, s, "Professional profile")
    story.append(Paragraph(C.PROFILE, s["body"]))

    # ---- career highlights ----------------------------------------------
    pdf_section(story, s, "Career highlights")
    inner = [Paragraph(h, s["panel"], bulletText=BULLET) for h in C.HIGHLIGHTS]
    panel = Table([[inner]], colWidths=[CW])
    panel.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BAND),
        ("LEFTPADDING", (0, 0), (-1, -1), 11),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LINEBEFORE", (0, 0), (0, 0), 2.8, TEAL),
    ]))
    story.append(panel)

    # ---- skills ----------------------------------------------------------
    pdf_section(story, s, "Technical skills")
    story.append(pdf_label_table(s, C.SKILLS, 1.66))

    # ---- experience ------------------------------------------------------
    pdf_section(story, s, "Professional experience")
    for i, job in enumerate(C.EXPERIENCE):
        story.append(Spacer(1, 5 if i else 0))
        story.append(KeepTogether([
            head_row(s, job["role"], job["dates"]),
            Spacer(1, 1.5),
            Paragraph(job["org"], s["org"]),
            Paragraph(job["bullets"][0], s["bullet"], bulletText=BULLET)]))
        for b in job["bullets"][1:]:
            story.append(Paragraph(b, s["bullet"], bulletText=BULLET))

    # ---- projects --------------------------------------------------------
    pdf_section(story, s, "Selected projects")
    for i, (name, role, org, period, detail) in enumerate(C.PROJECTS):
        story.append(Spacer(1, 5 if i else 0))
        story.append(KeepTogether([
            head_row(s, name, period),
            Spacer(1, 1.5),
            Paragraph(role + "  ·  " + org, s["org"]),
            Paragraph(detail, s["body"])]))

    # ---- competencies ----------------------------------------------------
    pdf_section(story, s, "Professional competencies")
    story.append(pdf_label_table(s, C.COMPETENCIES, 2.05))

    # ---- equipment -------------------------------------------------------
    pdf_section(story, s, "Medical equipment experience")
    story.append(pdf_label_table(s, C.EQUIPMENT, 1.76))

    # ---- education -------------------------------------------------------
    pdf_section(story, s, "Education")
    for i, (award, inst, year, note) in enumerate(C.EDUCATION):
        block = [head_row(s, award, year), Spacer(1, 1.5), Paragraph(inst, s["org"])]
        if note:
            block.append(Paragraph(note, s["body"]))
        story.append(Spacer(1, 5 if i else 0))
        story.append(KeepTogether(block))

    # ---- training --------------------------------------------------------
    pdf_section(story, s, "Training and certification")
    for course, provider, year in C.TRAINING:
        story.append(Paragraph(
            "<b>%s</b>.  <font color='#%s'>%s, %s.</font>" % (course, MUTED_HEX, provider, year),
            s["bullet"], bulletText=BULLET))

    # ---- interests -------------------------------------------------------
    pdf_section(story, s, "Interests")
    story.append(Paragraph(C.INTERESTS, s["body"]))

    # ---- referees --------------------------------------------------------
    pdf_section(story, s, "Referees")
    gap = 10
    card_w = (CW - 2 * gap) / 3.0
    row_cells, style = [], list(NOPAD) + [("VALIGN", (0, 0), (-1, -1), "TOP")]
    for i, (name, role, contact) in enumerate(C.REFEREES):
        col = i * 2
        row_cells.append([Paragraph(name, s["refname"]),
                          Paragraph(role, s["refrole"]),
                          Paragraph(contact, s["refcon"])])
        style += [("BACKGROUND", (col, 0), (col, 0), BAND),
                  ("LINEABOVE", (col, 0), (col, 0), 1.8, TEAL),
                  ("LEFTPADDING", (col, 0), (col, 0), 8),
                  ("RIGHTPADDING", (col, 0), (col, 0), 8),
                  ("TOPPADDING", (col, 0), (col, 0), 7),
                  ("BOTTOMPADDING", (col, 0), (col, 0), 7)]
        if i < len(C.REFEREES) - 1:
            row_cells.append("")
    t = Table([row_cells], colWidths=[card_w, gap, card_w, gap, card_w], hAlign="LEFT")
    t.setStyle(TableStyle(style))
    story.append(KeepTogether(t))

    doc.build(story)
    return PDF_OUT


if __name__ == "__main__":
    print("DOCX ->", build_docx())
    print("PDF  ->", build_pdf())
