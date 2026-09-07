"""
FairBanks house style for the new-grants investor pack.

Every document in this folder is generated through this module so that the
letterhead, palette, typography, tables and page furniture are identical
across the pack. Nothing here is document-specific.

Palette is sampled from assets/fairbanks_logo.jpeg:

    logo green   #07BA06   the mark itself
    logo orange  #F08400   the cross in the mark

The vivid logo green is kept for rules and small accents, where it reads as
the brand. Headings use a darkened version of the same hue, because vivid
green is uncomfortable to read as text at 12pt.

Page furniture: US Letter, 0.65in top/bottom, 0.7in sides. Page one carries
the full letterhead; continuation pages carry a compact running header and a
"Page X of Y" footer.
"""

from pathlib import Path

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
LOGO = REPO / "assets" / "fairbanks_logo.jpeg"


def logo_path():
    """The FairBanks mark, or a clear failure.

    assets/ is gitignored, so on a fresh clone the logo is absent. These are
    branded investor documents: silently emitting them without the mark is
    worse than not building them, so this raises rather than degrading.
    """
    if not LOGO.exists():
        raise FileNotFoundError(
            "FairBanks logo not found at " + str(LOGO) + ". The assets/ "
            "directory is gitignored, so restore it before building the "
            "investor pack."
        )
    return str(LOGO)


# -- palette ---------------------------------------------------------------
BRAND = "07BA06"          # logo green, for rules and accents
DEEP = "12591A"           # darkened logo green, for heading text
ORANGE = "E07B00"         # logo orange, used sparingly
MID = "D9EDD9"            # table total rows, callout fills
PALE = "EFF7EF"           # table banding
LINE = "C6DCC6"           # hairlines

BRAND_RGB = RGBColor(0x07, 0xBA, 0x06)
DEEP_RGB = RGBColor(0x12, 0x59, 0x1A)
ORANGE_RGB = RGBColor(0xE0, 0x7B, 0x00)
INK = RGBColor(0x22, 0x22, 0x22)
MUTED_RGB = RGBColor(0x5A, 0x66, 0x5A)
PALE_RGB = RGBColor(0xD9, 0xED, 0xD9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# -- typography ------------------------------------------------------------
BODY_FONT = "Aptos"
DISPLAY_FONT = "Aptos Display"
BODY_PT = 10.5

# -- page geometry (twips) -------------------------------------------------
PAGE_W, PAGE_H = 12240, 15840
MARGIN_TB, MARGIN_LR = 936, 1008
CONTENT_W = PAGE_W - 2 * MARGIN_LR          # 10224

# -- organisation ----------------------------------------------------------
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
ORG_MIXED = "FairBanks Medical Centre Limited"
TAGLINE = "Your Health, Our Mission"
ADDRESS = ("Kyebando-Kisalosalo, Tirupati Road (opposite the Northern Bypass "
           "roundabout), Kampala, Uganda")
ADDRESS_SHORT = "Kyebando-Kisalosalo, Tirupati Road, Kampala"
PHONES = "+256 748 319 052 / +256 777 462 398"
EMAIL = "info@fairbanksmedicalcentre.org"
WEB = "www.fairbanksmedicalcentre.org"
DOT = "   ·   "        # separator used across the pack
CONTACTS = EMAIL + DOT + PHONES + DOT + WEB
TIN = "1053370026"
CONFIDENTIAL = "Private and confidential. Prepared for the named recipient"


# --------------------------------------------------------------------------
# low-level XML helpers
# --------------------------------------------------------------------------

def _set_rfonts(rpr, font):
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font)


def shade(cell, fill):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:color"), "auto")
    el.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(el)


def cell_margins(table, top=90, bottom=90, left=130, right=130):
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement("w:" + tag)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    table._tbl.tblPr.append(mar)


def hairline_borders(table, color=LINE):
    """Horizontal rules only - the cleanest look for a financial table."""
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement("w:" + edge)
        if edge in ("insideH", "top", "bottom"):
            node.set(qn("w:val"), "single")
            node.set(qn("w:sz"), "4")
            node.set(qn("w:color"), color)
        else:
            node.set(qn("w:val"), "none")
            node.set(qn("w:sz"), "0")
            node.set(qn("w:color"), "auto")
        borders.append(node)
    table._tbl.tblPr.append(borders)


def no_borders(table):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "none")
        node.set(qn("w:sz"), "0")
        node.set(qn("w:color"), "auto")
        borders.append(node)
    table._tbl.tblPr.append(borders)


def rule(par, *, color=BRAND, sz=6, edge="top", space=4):
    ppr = par.paragraph_format.element.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.append(pbdr)
    node = OxmlElement("w:" + edge)
    node.set(qn("w:val"), "single")
    node.set(qn("w:sz"), str(sz))
    node.set(qn("w:space"), str(space))
    node.set(qn("w:color"), color)
    pbdr.append(node)


def left_accent(par, *, color=BRAND, sz=18, space=8):
    """A thick coloured bar down the left edge - used on callouts."""
    rule(par, color=color, sz=sz, edge="left", space=space)


def field(par, instr, *, size=8, color="5A665A", bold=False):
    """Insert a Word field (PAGE, NUMPAGES) styled like surrounding text."""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    _set_rfonts(rpr, BODY_FONT)
    if bold:
        rpr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), color)
    rpr.append(col)
    r.append(rpr)
    fld.append(r)
    par._p.append(fld)


# --------------------------------------------------------------------------
# text helpers
# --------------------------------------------------------------------------

def run(par, text, *, bold=False, italic=False, size=BODY_PT, color=INK,
        caps=False, font=BODY_FONT, spacing=None):
    r = par.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.all_caps = caps
    rpr = r._element.get_or_add_rPr()
    _set_rfonts(rpr, font)
    if spacing:                        # letter-spacing, twentieths of a point
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(spacing))
        rpr.append(sp)
    return r


def para(container, *, before=0, after=6, align=None, spacing=1.15,
         keep_with_next=False, keep_lines=False):
    """A styled paragraph.

    keep_lines is off by default: forcing a long justified paragraph to stay
    whole leaves half-empty pages. Widow/orphan control on the Normal style
    already stops single stranded lines, which is the effect actually wanted.
    """
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = spacing
    pf.keep_together = keep_lines
    pf.keep_with_next = keep_with_next
    if align is not None:
        pf.alignment = align
    return p


def body(doc, text, *, after=7, before=0, size=BODY_PT,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, spacing=1.15, lead=None):
    """Justified body copy. `lead` is a bold run-in opener."""
    p = para(doc, before=before, after=after, align=align, spacing=spacing)
    if lead:
        run(p, lead, bold=True, size=size, color=DEEP_RGB)
        run(p, "  ", size=size)
    run(p, text, size=size)
    return p


def heading(doc, number, text, *, before=11):
    p = para(doc, before=before, after=7, keep_with_next=True)
    if number:
        run(p, str(number) + "   ", bold=True, size=12, color=BRAND_RGB,
            font=DISPLAY_FONT)
    run(p, text, bold=True, size=12, color=DEEP_RGB, font=DISPLAY_FONT)
    rule(p, sz=8)
    return p


def subhead(doc, text, *, before=6, after=4):
    p = para(doc, before=before, after=after, keep_with_next=True)
    run(p, text, bold=True, size=9, color=DEEP_RGB, caps=True, spacing=12)
    return p


def bullet(doc, lead, text, *, after=4, indent=230):
    """List items are set ragged-right.

    Justifying a hanging-indent item with a bold run-in lead stretches the
    word spaces into rivers, because there is less line to work with than in
    a full-measure paragraph. Body copy stays justified; lists do not.
    """
    p = para(doc, after=after, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.left_indent = Twips(indent)
    p.paragraph_format.first_line_indent = Twips(-indent)
    run(p, "•   ", color=BRAND_RGB, bold=True)
    if lead:
        run(p, lead, bold=True, color=DEEP_RGB)
        run(p, "  ")
    run(p, text)
    return p


def numbered(doc, index, text, *, after=4, indent=250):
    p = para(doc, after=after, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.left_indent = Twips(indent)
    p.paragraph_format.first_line_indent = Twips(-indent)
    run(p, str(index) + ".   ", bold=True, color=BRAND_RGB)
    run(p, text)
    return p


def mark_size(par, pt):
    """Shrink a paragraph's own paragraph mark.

    An "empty" paragraph is never empty to Word: its height comes from the
    paragraph mark, which inherits Normal at 10.5pt. Without this a 2pt
    spacer really occupies about 14pt, which is why gaps drifted.
    """
    ppr = par.paragraph_format.element.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.insert(0, rpr)
    for tag in ("w:sz", "w:szCs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            rpr.append(node)
        node.set(qn("w:val"), str(int(pt * 2)))
    return par


def spacer(doc, pts=4):
    p = para(doc, after=pts, spacing=1.0)
    run(p, "", size=2)
    mark_size(p, 2)
    return p


# --------------------------------------------------------------------------
# tables
# --------------------------------------------------------------------------

def make_table(doc, widths, *, borders="hairline"):
    t = doc.add_table(rows=0, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    if borders == "hairline":
        hairline_borders(t)
    else:
        no_borders(t)
    cell_margins(t)
    grid = t._tbl.find(qn("w:tblGrid"))
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))
    return t


def fill_row(table, widths, values, *, fill=None, bold=False, size=9,
             color=INK, aligns=None, caps=False, header=False, spacing=None):
    row = table.add_row()
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))
    if header:
        trpr.append(OxmlElement("w:tblHeader"))
    for i, (cell, w, val) in enumerate(zip(row.cells, widths, values)):
        cell.width = Twips(w)
        if fill:
            shade(cell, fill)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.08
        if aligns:
            p.paragraph_format.alignment = aligns[i]
        if header:
            # stop a header row stranding itself at the foot of a page
            p.paragraph_format.keep_with_next = True
        run(p, val, bold=bold, size=size, color=color, caps=caps, spacing=spacing)
    return row


def head_row(table, widths, labels, aligns=None):
    return fill_row(table, widths, labels, fill=DEEP, bold=True, size=8.5,
                    color=WHITE, aligns=aligns, caps=True, header=True,
                    spacing=10)


def total_row(table, widths, values, aligns=None):
    return fill_row(table, widths, values, fill=MID, bold=True, size=9,
                    color=DEEP_RGB, aligns=aligns)


def banded(table, widths, rows, aligns=None, size=9):
    for i, r in enumerate(rows):
        fill_row(table, widths, r, fill=(PALE if i % 2 == 0 else "FFFFFF"),
                 aligns=aligns, size=size)


def no_split(row):
    """Keep a one-row banner (callout, figure strip, panel) on one page."""
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return row


def keep_together(table):
    """Stop Word splitting a short table across a page break."""
    for row in table.rows[:-1]:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True


def gap(doc, pts=7):
    """A hairline paragraph between two tables.

    Word silently merges tables that touch, so every table in this pack is
    followed by a paragraph. Anything visible (a subhead, body copy) does the
    job; this is for the places where nothing else separates them.
    """
    p = para(doc, before=0, after=pts, spacing=1.0)
    run(p, "", size=2)
    mark_size(p, 2)
    return p


def callout(doc, title, text, *, fill=PALE, accent=BRAND, after=7):
    t = make_table(doc, [CONTENT_W], borders="none")
    cell = no_split(t.add_row()).cells[0]
    cell.width = Twips(CONTENT_W)
    shade(cell, fill)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    left_accent(p, color=accent)
    if title:
        run(p, title, bold=True, size=8.5, color=DEEP_RGB, caps=True, spacing=12)
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.line_spacing = 1.15
        p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_accent(p2, color=accent)
        run(p2, text, size=9.5)
    else:
        run(p, text, size=9.5)
    gap(doc, after)
    return t


def statement(doc, text, *, fill=MID, accent=ORANGE, after=8, size=13):
    """A closing line set as a banner - the one place a slogan is allowed."""
    t = make_table(doc, [CONTENT_W], borders="none")
    cell = no_split(t.add_row()).cells[0]
    cell.width = Twips(CONTENT_W)
    shade(cell, fill)
    p = cell.paragraphs[0]
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.05
    left_accent(p, color=accent)
    run(p, text, bold=True, size=size, color=DEEP_RGB, font=DISPLAY_FONT,
        spacing=16)
    gap(doc, after)
    return t


def page_break(doc):
    """Start the next block on a fresh page - used for appendices."""
    p = para(doc, before=0, after=0, spacing=1.0)
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r = OxmlElement("w:r")
    r.append(br)
    p._p.append(r)
    return p


def label_panels(doc, items, *, fill=PALE, after=7, center=False, size=9):
    """A row of tinted panels, each a small green label over optional text.

    items: [(label, text_or_None), ...]. Used for at-a-glance strips such as
    "what we run / what we cannot do" or a list of reporting fields.
    """
    w = CONTENT_W // len(items)
    widths = [w] * len(items)
    t = make_table(doc, widths, borders="none")
    row = no_split(t.add_row())
    align = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    for cell, (label, text) in zip(row.cells, items):
        cell.width = Twips(w)
        shade(cell, fill)
        p = cell.paragraphs[0]
        p.paragraph_format.alignment = align
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1 if text else 3)
        p.paragraph_format.line_spacing = 1.05
        run(p, label, bold=True, size=8, color=DEEP_RGB, caps=True, spacing=12)
        if text:
            p2 = cell.add_paragraph()
            p2.paragraph_format.alignment = align
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(3)
            p2.paragraph_format.line_spacing = 1.1
            run(p2, text, size=size)
    gap(doc, after)
    return t


def figure_strip(doc, items, *, fill=DEEP, after=7):
    """A band of headline numbers: [(value, label), ...]."""
    w = CONTENT_W // len(items)
    widths = [w] * len(items)
    t = make_table(doc, widths, borders="none")
    row = no_split(t.add_row())
    for cell, (value, label) in zip(row.cells, items):
        cell.width = Twips(w)
        shade(cell, fill)
        p = cell.paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run(p, value, bold=True, size=13, color=WHITE, font=DISPLAY_FONT)
        p2 = cell.add_paragraph()
        p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(6)
        p2.paragraph_format.line_spacing = 1.0
        run(p2, label, size=7.5, color=PALE_RGB, caps=True, spacing=14)
    gap(doc, after)
    return t


# --------------------------------------------------------------------------
# document furniture
# --------------------------------------------------------------------------

def _page_setup(doc):
    s = doc.sections[0]
    s.page_width = Twips(PAGE_W)
    s.page_height = Twips(PAGE_H)
    s.top_margin = Twips(MARGIN_TB)
    s.bottom_margin = Twips(MARGIN_TB)
    s.left_margin = Twips(MARGIN_LR)
    s.right_margin = Twips(MARGIN_LR)
    s.header_distance = Twips(432)
    s.footer_distance = Twips(432)
    s.different_first_page_header_footer = True
    return s


def _normal_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(BODY_PT)
    normal.font.color.rgb = INK
    rpr = normal.element.get_or_add_rPr()
    _set_rfonts(rpr, BODY_FONT)
    ppr = normal.element.get_or_add_pPr()
    for tag, val in (("w:widowControl", "1"), ("w:suppressAutoHyphens", "1")):
        node = OxmlElement(tag)
        node.set(qn("w:val"), val)
        ppr.append(node)


def _metadata(doc, title, subject, author, keywords, comments):
    cp = doc.core_properties
    cp.title = title
    cp.subject = subject
    cp.author = author
    cp.last_modified_by = author
    cp.category = "Investment proposal"
    cp.comments = comments
    cp.keywords = keywords


def new_document(*, title, subject, author=ORG_MIXED, keywords="",
                 comments=CONFIDENTIAL):
    """A blank FairBanks-styled document: page geometry, Normal style, metadata."""
    from docx import Document
    doc = Document()
    _page_setup(doc)
    _normal_style(doc)
    _metadata(doc, title, subject, author, keywords, comments)
    return doc


def _tabbed(par, right_at=CONTENT_W):
    par.paragraph_format.tab_stops.add_tab_stop(
        Twips(right_at), WD_TAB_ALIGNMENT.RIGHT)
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.line_spacing = 1.0
    return par


def brand_rule(doc_or_cell, *, before=0, after=8):
    """The house double rule: a heavy brand-green bar over a fine deep line.

    Both rules live on one paragraph (top border and bottom border) whose
    line height is pinned to 2pt, so they read as a single device rather
    than as two rules that happen to be near each other.
    """
    p = para(doc_or_cell, before=before, after=after, spacing=None)
    mark_size(p, 2)
    p.paragraph_format.line_spacing = Pt(2)
    rule(p, color=BRAND, sz=14, space=0, edge="top")
    rule(p, color=DEEP, sz=4, space=0, edge="bottom")
    run(p, "", size=1)
    return p


def letterhead(doc, *, logo_width=1.95, with_address=True, with_contacts=False):
    """Page-one masthead: logo, legal name, tagline, address, house rule."""
    p = para(doc, after=3, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    p.add_run().add_picture(logo_path(), width=Inches(logo_width))
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    run(p, ORG, bold=True, size=14, color=DEEP_RGB, font=DISPLAY_FONT,
        spacing=24)
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    run(p, TAGLINE, italic=True, size=9.5, color=BRAND_RGB)
    if with_address:
        p = para(doc, before=2, after=0, align=WD_ALIGN_PARAGRAPH.CENTER,
                 spacing=1.0)
        run(p, ADDRESS, size=8.5, color=MUTED_RGB)
    if with_contacts:
        p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
        run(p, CONTACTS, size=8.5, color=MUTED_RGB)
    brand_rule(doc, before=6, after=10)


def running_header(doc, text, *, logo_width=1.05):
    """Compact masthead repeated on continuation pages."""
    header = doc.sections[0].header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    _tabbed(p)
    p.add_run().add_picture(logo_path(), width=Inches(logo_width))
    run(p, "	")
    run(p, text, size=8, color=MUTED_RGB, caps=True, spacing=14)
    rule(p, color=LINE, sz=4, edge="bottom", space=3)
    # page one carries the full letterhead instead, so leave its header empty
    doc.sections[0].first_page_header.is_linked_to_previous = False
    return p


def _footer_line(par, note):
    _tabbed(par)
    rule(par, color=LINE, sz=4, edge="top", space=5)
    run(par, note, size=7.5, color=MUTED_RGB)
    run(par, "	")
    run(par, "Page ", size=8, color=MUTED_RGB)
    field(par, "PAGE", size=8, color="12591A", bold=True)
    run(par, " of ", size=8, color=MUTED_RGB)
    field(par, "NUMPAGES", size=8, color="12591A", bold=True)
    return par


def page_footer(doc, note=None):
    """Confidentiality note left, live 'Page X of Y' right, on every page."""
    note = note or (ORG_MIXED + "  ·  " + CONFIDENTIAL)
    section = doc.sections[0]
    for footer in (section.footer, section.first_page_footer):
        footer.is_linked_to_previous = False
        _footer_line(footer.paragraphs[0], note)


def reference_line(doc, ref, date, *, after=10):
    """Document control block: reference number left, date right."""
    p = para(doc, after=after, spacing=1.0)
    _tabbed(p)
    p.paragraph_format.space_after = Pt(after)
    run(p, "Ref: " + ref, size=9, color=DEEP_RGB, bold=True)
    run(p, "	")
    run(p, date, size=9, color=MUTED_RGB)
    return p


def signoff_block(doc, name, role, *, with_contacts=True):
    p = para(doc, before=2, after=0, spacing=1.0)
    run(p, name, bold=True, size=11, color=DEEP_RGB)
    p = para(doc, after=0, spacing=1.0)
    run(p, role, size=9.5)
    p = para(doc, after=0, spacing=1.0)
    run(p, ORG_MIXED, size=9.5)
    if with_contacts:
        p = para(doc, before=3, after=0, spacing=1.0)
        run(p, CONTACTS, size=8.5, color=MUTED_RGB)
    return p


def closing_block(doc, *, before=9):
    """Foot-of-document identity block, centred under the house rule.

    Chained with keep-with-next so the rule, the name and the two contact
    lines can never be split by a page break, which would leave the document
    ending on half an address.
    """
    r = brand_rule(doc, before=before, after=2)
    r.paragraph_format.keep_with_next = True
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0,
             keep_with_next=True)
    run(p, ORG_MIXED, bold=True, size=9, color=DEEP_RGB)
    run(p, "   ·   ", size=9, color=BRAND_RGB, bold=True)
    run(p, TAGLINE, italic=True, size=9, color=BRAND_RGB)
    p = para(doc, before=2, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0,
             keep_with_next=True)
    run(p, ADDRESS_SHORT + DOT + "TIN " + TIN, size=8, color=MUTED_RGB)
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    run(p, CONTACTS, size=8, color=MUTED_RGB)
    return p
