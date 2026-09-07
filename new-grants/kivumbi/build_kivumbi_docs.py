"""
Builds the FairBanks maternity investment pack for Mr. Alex Kivumbi.

Two outputs, written next to this script:
  FairBanks_Cover_Letter_Alex_Kivumbi.docx
  FairBanks_Maternity_Investment_Proposal_Alex_Kivumbi.docx

House styling matches the rest of the new-grants pack: US Letter, Aptos,
FairBanks green 165A43 with DDEBE4 / EAF4EF tints.

Yellow-highlighted text marks the blanks Racheal must fill before sending.
"""

import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

HERE = os.path.dirname(os.path.abspath(__file__))

GREEN = "165A43"
MID = "DDEBE4"
PALE = "EAF4EF"
GREEN_RGB = RGBColor(0x16, 0x5A, 0x43)
INK = RGBColor(0x22, 0x22, 0x22)

CONTENT_W = 10224  # dxa, US Letter less 1008 twip side margins

ADDRESS = "Kyebando-Kisalosalo, Northern Bypass, Kampala"
CONTACTS = ("info@fairbanksmedicalcentre.org  |  +256 748 319 052 / +256 777 462 398  |  "
            "www.fairbanksmedicalcentre.org")


# --------------------------------------------------------------------------
# low-level helpers
# --------------------------------------------------------------------------

def shade(cell, fill):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:color"), "auto")
    el.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(el)


def cell_margins(table, top=80, bottom=80, left=120, right=120):
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    table._tbl.tblPr.append(mar)


def no_borders(table):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:color"), "FFFFFF")
        borders.append(node)
    table._tbl.tblPr.append(borders)


def top_rule(par, color=GREEN, sz=6):
    pbdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), str(sz))
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), color)
    pbdr.append(top)
    par.paragraph_format.element.get_or_add_pPr().append(pbdr)


def run(par, text, *, bold=False, italic=False, size=10.5, color=INK,
        caps=False, highlight=False, font="Aptos"):
    r = par.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.all_caps = caps
    if highlight:
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), font)
    return r


def para(doc, *, before=0, after=6, align=None, spacing=1.15):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = spacing
    if align is not None:
        pf.alignment = align
    return p


def body(doc, text, *, after=8, before=0, size=10.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         spacing=1.15):
    p = para(doc, before=before, after=after, align=align, spacing=spacing)
    run(p, text, size=size)
    return p


def heading(doc, text):
    p = para(doc, before=14, after=6)
    p.paragraph_format.keep_with_next = True
    run(p, text, bold=True, size=12, color=GREEN_RGB, font="Aptos Display")
    top_rule(p)
    return p


def subhead(doc, text, *, before=4):
    p = para(doc, before=before, after=3)
    p.paragraph_format.keep_with_next = True
    run(p, text, bold=True, size=9.5, color=GREEN_RGB)
    return p


def bullet(doc, lead, text):
    p = para(doc, after=5, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.left_indent = Twips(220)
    p.paragraph_format.first_line_indent = Twips(-220)
    run(p, "•   ")
    if lead:
        run(p, lead, bold=True)
        run(p, "  ")
    run(p, text)
    return p


def make_table(doc, widths):
    t = doc.add_table(rows=0, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    no_borders(t)
    cell_margins(t)
    grid = t._tbl.find(qn("w:tblGrid"))
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))
    return t


def fill_row(table, widths, values, *, fill, bold=False, size=9,
             color=INK, aligns=None, caps=False, header=False):
    row = table.add_row()
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))
    if header:
        trpr.append(OxmlElement("w:tblHeader"))
    for i, (cell, w, val) in enumerate(zip(row.cells, widths, values)):
        cell.width = Twips(w)
        shade(cell, fill)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        if aligns:
            p.paragraph_format.alignment = aligns[i]
        run(p, val, bold=bold, size=size, color=color, caps=caps)
    return row


def spacer(doc, pts=4):
    p = para(doc, after=pts)
    run(p, "")
    return p


def callout(doc, title, text, *, fill=PALE):
    t = make_table(doc, [CONTENT_W])
    row = t.add_row()
    cell = row.cells[0]
    cell.width = Twips(CONTENT_W)
    shade(cell, fill)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    if title:
        run(p, title, bold=True, size=9, color=GREEN_RGB, caps=True)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(1)
    p2.paragraph_format.line_spacing = 1.15
    run(p2, text, size=9.5)
    return t


def page_setup(doc):
    s = doc.sections[0]
    s.page_width = Twips(12240)
    s.page_height = Twips(15840)
    s.top_margin = Twips(936)
    s.bottom_margin = Twips(936)
    s.left_margin = Twips(1008)
    s.right_margin = Twips(1008)
    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = INK
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Aptos")


def letterhead(doc, *, with_address=True):
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, "FAIRBANKS MEDICAL CENTRE LIMITED", bold=True, size=17,
        color=GREEN_RGB, font="Aptos Display")
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, "Your Health, Our Mission", italic=True, size=10, color=GREEN_RGB)
    if with_address:
        p = para(doc, after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
        run(p, ADDRESS, size=9)
    p = para(doc, after=7)
    top_rule(p, sz=12)
    run(p, "")


def signoff_block(doc):
    p = para(doc, before=2, after=0)
    run(p, "Racheal Nabukeera", bold=True, size=11)
    p = para(doc, after=0)
    run(p, "Managing Director", size=10)
    p = para(doc, after=0)
    run(p, "FairBanks Medical Centre Limited", size=10)
    p = para(doc, before=2, after=0)
    run(p, CONTACTS, size=8.5, color=GREEN_RGB)


# --------------------------------------------------------------------------
# document 1 — cover letter
# --------------------------------------------------------------------------

def build_cover_letter(path):
    doc = Document()
    page_setup(doc)
    letterhead(doc)

    p = para(doc, after=0)
    run(p, "Ref: FBMC/INV/2026/002", size=9.5, color=GREEN_RGB, bold=True)
    p = para(doc, after=8)
    run(p, "7 September 2026", size=10)

    for line in ("Mr. Alex Kivumbi", "Kampala"):
        p = para(doc, after=0)
        run(p, line, size=10.5, bold=(line.startswith("Mr.")))

    p = para(doc, before=8, after=7)
    run(p, "Dear Mr. Kivumbi,", size=10.5)

    p = para(doc, after=8)
    run(p, "RE:  PROPOSAL FOR INVESTMENT IN THE FAIRBANKS MATERNITY AND INPATIENT UNIT",
        bold=True, size=10.5, color=GREEN_RGB)

    p = para(doc, after=5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, spacing=1.07)
    run(p, "Thank you for the time you gave us on ")
    run(p, "[date of our meeting]", highlight=True)
    run(p, " and for the interest you showed in what we are building at Kyebando. "
           "The full proposal is enclosed.")

    def para_(text, after=5):
        return body(doc, text, after=after, spacing=1.07)

    para_("Before you read it, one thing has changed since we spoke, and it affects the "
          "figure.")

    para_("We had described this project as needing between 75 and 100 million shillings. "
          "That is still right for the ward itself: the delivery room, the inpatient beds, "
          "the scanner and the staff to run them. It is not right for a maternity unit that "
          "can also operate. Having gone through it properly with our clinical team, we do "
          "not think we should build the ward and leave the theatre as a vague intention for "
          "later. A maternity unit without a theatre sends every obstructed labour and every "
          "emergency caesarean out of the gate in an ambulance, at the worst possible moment "
          "for the mother and for our name.")

    para_("So the enclosed proposal prices the complete unit, ward and theatre together, at "
          "between 190 and 265 million shillings, split into two phases. Phase one is the 75 "
          "to 100 million we discussed and can begin as soon as funds are released. Phase "
          "two is the theatre, and it can be committed separately once phase one is running "
          "and you have seen the numbers it actually produces.")

    para_("You should also know how firm these costs are. One line is a real quotation: the "
          "obstetric scanner, at 11 million shillings from Canon Healthcare Systems on 13 "
          "April this year. The rest are our own estimates, made before going to suppliers, "
          "and each will be quoted and returned to you as a signed one-page budget before "
          "any money moves.")

    para_("What we are asking for now is a conversation, not a commitment. Tell us the level "
          "you would be comfortable considering and how you would want a return structured, "
          "and we will build the priced budget around that. May I also suggest you come and "
          "see it? The space is there and needs finishing. Twenty minutes walking through it "
          "will tell you more than this document will, and I can arrange for you to meet the "
          "clinical team. I am available any weekday.")

    para_("Thank you again for your time, and for taking this seriously.", after=10)

    p = para(doc, after=0)
    run(p, "Yours sincerely,", size=10.5)
    spacer(doc, 10)
    signoff_block(doc)

    spacer(doc, 5)
    encl = ("Investment Proposal — Maternity and Inpatient Unit, 7 September 2026",
            "Proforma invoice 1896, Canon Healthcare Systems Ltd, 13 April 2026")
    for i, item in enumerate(encl):
        p = para(doc, after=0)
        p.paragraph_format.left_indent = Twips(520)
        p.paragraph_format.first_line_indent = Twips(-520)
        p.paragraph_format.tab_stops.add_tab_stop(Twips(520))
        if i == 0:
            top_rule(p, color=MID, sz=6)
            run(p, "Enc.", bold=True, size=9, color=GREEN_RGB)
            run(p, "\t", size=9)
        else:
            run(p, "\t", size=9)
        run(p, item, size=9)

    doc.save(path)
    return path


# --------------------------------------------------------------------------
# document 2 — proposal
# --------------------------------------------------------------------------

def build_proposal(path):
    doc = Document()
    page_setup(doc)
    letterhead(doc, with_address=False)

    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, "INVESTMENT PROPOSAL", bold=True, size=15, color=GREEN_RGB,
        font="Aptos Display")
    p = para(doc, after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, "Maternity and Inpatient Unit with Obstetric Theatre", size=12,
        color=GREEN_RGB)
    p = para(doc, after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, "Prepared for Mr. Alex Kivumbi   |   7 September 2026", size=9.5)

    callout(doc, "In short",
            "FairBanks is asking for UGX 190–265M to complete and commission a maternity "
            "and inpatient unit with its own obstetric theatre, built in two phases. Phase "
            "one costs UGX 75–100M and can begin immediately. Phase two, the theatre, can "
            "be committed once phase one is open and reporting.")
    spacer(doc, 6)

    # 1 --------------------------------------------------------------------
    heading(doc, "1.  Why maternity, and why now")

    body(doc,
         "FairBanks Medical Centre operates in Kyebando-Kisalosalo, on the Northern Bypass. "
         "The centre is licensed and running: outpatient consultations, laboratory, pharmacy, "
         "X-ray, specialist clinics, physiotherapy and a 24-hour emergency service, together "
         "producing between UGX 8M and UGX 12M a month. We accept CIC and AAR, with Jubilee "
         "and Mutual in process.")

    body(doc,
         "Obstetrics and gynaecology is already one of our clinics. What we do not have is "
         "anywhere to admit the women who use it. An expectant mother can have her antenatal "
         "care with us, her scans with us and her medicines from our pharmacy, and then "
         "deliver somewhere else, because at the point where she needs a bed we have nothing "
         "to offer her. Every one of those deliveries is revenue we have already paid to "
         "acquire and then handed to another facility.")

    body(doc,
         "The maternity and inpatient unit closes that gap. It is also the only expansion "
         "open to us that creates a genuinely new service line rather than adding capacity to "
         "one we already run.")

    # 2 --------------------------------------------------------------------
    heading(doc, "2.  What we are asking you to fund")

    body(doc,
         "The work splits cleanly into two phases. Phase one produces a functioning maternity "
         "and inpatient ward. Phase two adds the operating theatre and neonatal capacity that "
         "let us keep emergencies in the building.")

    w3 = [2500, 5324, 2400]
    heads = ["ITEM", "WHAT IT COVERS", "ESTIMATE"]
    right = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]

    subhead(doc, "PHASE ONE  —  MATERNITY AND INPATIENT WARD")

    t = make_table(doc, w3)
    fill_row(t, w3, heads, fill=GREEN, bold=True, size=8.5,
             color=RGBColor(0xFF, 0xFF, 0xFF), aligns=right, header=True)
    phase1 = [
        ("Ward completion and fit-out",
         "Partitioning, flooring, plumbing, sluice room, washrooms, lighting and power "
         "points for the delivery room, postnatal room and inpatient bays",
         "UGX 22–30M"),
        ("Delivery and ward equipment",
         "Two delivery beds, ten patient beds and lockers, delivery and episiotomy sets, "
         "examination couches, vital-signs monitors, suction units, a resuscitaire and "
         "oxygen concentrators",
         "UGX 28–36M"),
        ("Obstetric ultrasound",
         "Mindray DP-10 with printer and trolley. Quoted at UGX 11,000,000 by Canon "
         "Healthcare Systems Ltd, proforma 1896 of 13 April 2026",
         "UGX 11M"),
        ("Staffing and launch",
         "Recruitment and first-quarter cost of UNMC-registered midwives and ward nursing "
         "cover, orientation, licence variation and inspection",
         "UGX 10–14M"),
        ("Opening working capital",
         "Oxytocin and essential obstetric medicines, sutures, gloves, cord clamps, linen, "
         "laundry and utilities for the first three months",
         "UGX 6–9M"),
    ]
    for i, row in enumerate(phase1):
        fill_row(t, w3, row, fill=(PALE if i % 2 == 0 else "FFFFFF"), aligns=right)
    fill_row(t, w3, ("Phase one total", "", "UGX 77–100M"), fill=MID,
             bold=True, size=9, aligns=right)

    subhead(doc, "PHASE TWO  —  OBSTETRIC THEATRE AND NEONATAL CARE", before=12)

    t = make_table(doc, w3)
    fill_row(t, w3, heads, fill=GREEN, bold=True, size=8.5,
             color=RGBColor(0xFF, 0xFF, 0xFF), aligns=right, header=True)
    phase2 = [
        ("Theatre construction and services",
         "Theatre, scrub area, sterile store and sluice; sealed walls and flooring; "
         "ventilation; theatre lighting; dedicated power with UPS and generator backup",
         "UGX 30–45M"),
        ("Anaesthesia and surgical equipment",
         "Anaesthetic machine with ventilator, operating table, theatre lights, diathermy "
         "unit, theatre and recovery monitors, theatre suction",
         "UGX 45–60M"),
        ("Instruments and sterilisation",
         "Caesarean and gynaecological instrument sets, autoclave, instrument trolleys and "
         "sterile storage",
         "UGX 14–20M"),
        ("Neonatal capacity",
         "Incubator, phototherapy unit and neonatal resuscitation equipment",
         "UGX 12–18M"),
        ("Theatre staffing and commissioning",
         "Obstetric and anaesthetic cover, theatre nurse, commissioning, clinical protocols, "
         "blood-supply arrangement and licensing for surgical services",
         "UGX 14–22M"),
    ]
    for i, row in enumerate(phase2):
        fill_row(t, w3, row, fill=(PALE if i % 2 == 0 else "FFFFFF"), aligns=right)
    fill_row(t, w3, ("Phase two total", "", "UGX 115–165M"), fill=MID,
             bold=True, size=9, aligns=right)

    spacer(doc, 6)
    callout(doc, "Complete programme",
            "UGX 190–265M. Every figure above except the scanner is our own estimate, "
            "prepared before going to suppliers. They are ranges because that is honestly "
            "what we know today. No funds should be released against an estimate: each line "
            "will be quoted and returned to you as a signed one-page budget first.",
            fill=MID)

    # 3 --------------------------------------------------------------------
    heading(doc, "3.  What the unit should earn")

    body(doc,
         "The figures below are FairBanks' planning assumptions, not results. They are set "
         "out in full so you can argue with them, which is the point of showing the "
         "arithmetic rather than the conclusion.")

    subhead(doc, "PHASE ONE AT STEADY UTILISATION")

    w4 = [3624, 2000, 2200, 2400]
    a4 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT,
          WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    t = make_table(doc, w4)
    fill_row(t, w4, ["SERVICE", "PER MONTH", "PRICE (UGX)", "REVENUE (UGX)"],
             fill=GREEN, bold=True, size=8.5, color=RGBColor(0xFF, 0xFF, 0xFF),
             aligns=a4, header=True)
    rev = [
        ("Antenatal visits", "100", "25,000", "2,500,000"),
        ("Normal deliveries", "18", "450,000", "8,100,000"),
        ("Obstetric ultrasound scans", "70", "50,000", "3,500,000"),
        ("Inpatient bed-days", "70", "70,000", "4,900,000"),
    ]
    for i, row in enumerate(rev):
        fill_row(t, w4, row, fill=(PALE if i % 2 == 0 else "FFFFFF"), aligns=a4)
    fill_row(t, w4, ("Monthly revenue", "", "", "19,000,000"), fill=MID,
             bold=True, size=9, aligns=a4)

    subhead(doc, "MONTHLY RUNNING COSTS", before=10)

    w2 = [7824, 2400]
    a2 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]
    t = make_table(doc, w2)
    fill_row(t, w2, ["COST", "UGX"], fill=GREEN, bold=True, size=8.5,
             color=RGBColor(0xFF, 0xFF, 0xFF), aligns=a2, header=True)
    costs = [
        ("Midwifery and ward nursing — four midwives, a ward in-charge, two support staff",
         "5,500,000"),
        ("Medicines, sutures and clinical consumables", "2,800,000"),
        ("Utilities, standby power, water and clinical waste disposal", "1,200,000"),
        ("Laundry, cleaning, maintenance and equipment servicing", "700,000"),
    ]
    for i, row in enumerate(costs):
        fill_row(t, w2, row, fill=(PALE if i % 2 == 0 else "FFFFFF"), aligns=a2)
    fill_row(t, w2, ("Monthly running costs", "10,200,000"), fill=MID,
             bold=True, size=9, aligns=a2)

    spacer(doc, 6)
    callout(doc, "Phase one monthly contribution",
            "UGX 19,000,000 less UGX 10,200,000 = UGX 8,800,000 a month, once utilisation "
            "settles.", fill=MID)

    subhead(doc, "WHAT THE THEATRE ADDS", before=12)

    t = make_table(doc, w2)
    fill_row(t, w2, ["PHASE TWO", "UGX"], fill=GREEN, bold=True, size=8.5,
             color=RGBColor(0xFF, 0xFF, 0xFF), aligns=a2, header=True)
    theatre = [
        ("Caesarean sections — 5 a month at UGX 1,700,000", "8,500,000"),
        ("Less obstetric and anaesthetic session fees, 5 cases", "(2,500,000)"),
        ("Less theatre nurse and sterilisation cover", "(1,400,000)"),
        ("Less theatre consumables, drugs and blood handling", "(1,300,000)"),
        ("Less theatre maintenance and equipment servicing", "(600,000)"),
    ]
    for i, row in enumerate(theatre):
        fill_row(t, w2, row, fill=(PALE if i % 2 == 0 else "FFFFFF"), aligns=a2)
    fill_row(t, w2, ("Additional monthly contribution", "2,700,000"), fill=MID,
             bold=True, size=9, aligns=a2)

    body(doc,
         "That takes the combined contribution to roughly UGX 11,500,000 a month. We have "
         "not modelled gynaecological use of the theatre, which in practice will carry a "
         "share of the load and improve the figure.", before=8)

    body(doc,
         "At those volumes phase one returns its own capital in about ten months of steady "
         "operation. Ten months of steady operation is not ten months from the day you "
         "release funds. Construction, recruitment, licensing and the time an antenatal book "
         "takes to fill mean the first six months will run below the table above. Our working "
         "expectation is capital recovery on phase one in fourteen to eighteen months, and on "
         "the full programme including the theatre in two and a half to three years.")

    spacer(doc, 4)
    callout(doc, "Five numbers to replace before this is final",
            "Current monthly antenatal attendance · the delivery fee we intend to charge "
            "· the bed count we are actually building to · our intended bed-day rate "
            "· the caesarean fee. Those five drive everything else on this page, and "
            "FairBanks' own figures should replace the assumptions above before any agreement "
            "is signed.")

    # 4 --------------------------------------------------------------------
    heading(doc, "4.  What could go wrong")

    body(doc, "We would rather set these out ourselves than have you find them.")

    for lead, text in [
        ("Utilisation.",
         "A delivery room fills from the antenatal book, not from advertising. If antenatal "
         "attendance does not grow, the delivery numbers above do not happen. That is exactly "
         "why phase one puts the ward and the scanner in first: you get six months of real "
         "figures before the larger commitment."),
        ("Clinical risk.",
         "Obstetrics carries the highest liability of anything we do. The centre already "
         "holds professional indemnity and all-risks cover, and both would be reviewed and "
         "increased before the unit opens. Protocols, partograph discipline, emergency drills "
         "and genuine 24-hour cover are conditions of opening, not refinements to add later."),
        ("The gap before the theatre.",
         "Until phase two is commissioned, caesareans and complications have to go out. That "
         "needs a written referral arrangement with a partner hospital and a transport plan "
         "that works at three in the morning, both in place before the first delivery."),
        ("Regulation.",
         "A maternity unit with a theatre needs the facility licence varied to cover that "
         "level of care, UNMC-registered midwives, UMDPC-registered clinicians and compliant "
         "clinical waste handling. We would not spend phase two money before the licence "
         "position is confirmed in writing."),
        ("Cost.",
         "Every figure in this document except the scanner is an estimate made before "
         "quotation. Firm prices could come back above these ranges, most likely on "
         "anaesthesia equipment, which is imported and moves with the exchange rate."),
        ("Staffing.",
         "Anaesthetic officers and obstetricians are scarce and they move. We propose "
         "sessional arrangements in the first year rather than full-time salaries, which "
         "keeps the cost variable and the downside smaller."),
    ]:
        bullet(doc, lead, text)

    # 5 --------------------------------------------------------------------
    heading(doc, "5.  How your money would be protected")

    for lead, text in [
        ("Separate books.",
         "The unit is accounted for apart from the rest of the centre, with its own revenue "
         "and cost ledger from the first day."),
        ("Release against quotations.",
         "Funds move in tranches against signed supplier quotations and delivery notes, never "
         "against the estimates in this document."),
        ("A one-page report every month.",
         "Capital deployed, antenatal visits, deliveries, theatre cases, revenue, direct "
         "costs, monthly contribution and capital recovered to date."),
        ("Your own eyes.",
         "You may nominate someone to inspect the books, the store and the equipment at any "
         "time, without notice."),
        ("Registered assets.",
         "Capital equipment is asset-tagged, insured and listed, and can be named in the "
         "investment agreement as security."),
    ]:
        bullet(doc, lead, text)

    body(doc,
         "FairBanks Medical Centre Limited is a registered Ugandan company, TIN 1053370026. "
         "The centre holds a current operating licence, a National Drug Authority licence for "
         "the pharmacy, a certificate of suitability of premises, NSSF registration, and "
         "professional indemnity and all-risks insurance. Copies of all of these are "
         "available for your review before you commit anything.", before=6)

    # 6 --------------------------------------------------------------------
    heading(doc, "6.  How the investment could be structured")

    for lead, text in [
        ("A project loan.",
         "You lend a fixed amount for a fixed term at an agreed rate, repaid from the unit's "
         "contribution. Simplest to document, and your return does not depend on how well we "
         "run the place."),
        ("Revenue share until recovery, plus a premium.",
         "An agreed share of the unit's monthly contribution comes to you until your capital "
         "and an agreed premium are repaid, after which the arrangement ends. Your return "
         "tracks performance, and repayment only starts once there is something to share."),
        ("Equity.",
         "A shareholding in the unit or in the company, giving you a permanent share of the "
         "upside and a voice in the decisions."),
    ]:
        bullet(doc, lead, text)

    body(doc,
         "Our own preference is the second. It ties what you earn to how well we run the "
         "unit, it does not saddle the centre with fixed repayments during the months when "
         "utilisation is still building, and it has a defined end. That said, we are open on "
         "this. The structure should be the one that suits you.", before=6)

    # 7 --------------------------------------------------------------------
    heading(doc, "7.  What we would like to agree")

    steps = [
        "You tell us the level of investment you would be comfortable considering, and which "
        "structure you would prefer.",
        "We put every line in section 2 to suppliers and come back within three weeks with a "
        "priced one-page budget, a tranche schedule and a cash-flow projection.",
        "We agree the return formula, the reporting format and the release conditions in "
        "writing.",
        "The first tranche is released and phase one begins.",
    ]
    for i, s in enumerate(steps, 1):
        p = para(doc, after=5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        p.paragraph_format.left_indent = Twips(220)
        p.paragraph_format.first_line_indent = Twips(-220)
        run(p, f"{i}.   ", bold=True, color=GREEN_RGB)
        run(p, s)

    body(doc,
         "We would be grateful for a meeting at the centre within the next two weeks, so you "
         "can see the space before you decide anything.", before=6)

    spacer(doc, 8)
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    top_rule(p, sz=12)
    run(p, "FairBanks Medical Centre Limited   ·   " + ADDRESS, size=9,
        color=GREEN_RGB, bold=True)
    p = para(doc, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run(p, CONTACTS, size=8.5)

    doc.save(path)
    return path


if __name__ == "__main__":
    a = build_cover_letter(os.path.join(HERE, "FairBanks_Cover_Letter_Alex_Kivumbi.docx"))
    b = build_proposal(os.path.join(
        HERE, "FairBanks_Maternity_Investment_Proposal_Alex_Kivumbi.docx"))
    for f in (a, b):
        print("wrote", f, os.path.getsize(f), "bytes")
