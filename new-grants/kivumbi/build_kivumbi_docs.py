"""
Builds the FairBanks maternity investment pack for Mr. Alex Kivumbi.

Two outputs, written next to this script:

    FairBanks_Maternity_Cover_Letter_Kivumbi.docx
    FairBanks_Maternity_Proposal_Kivumbi.docx

All styling comes from ../fairbanks_brand.py, which is shared with the other
packs in new-grants so the letterhead, palette and page furniture match.

No fill-in blanks remain; the assumptions still to be replaced with FairBanks'
own figures are named explicitly in section 3 of the proposal.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx.enum.text import WD_ALIGN_PARAGRAPH as ALIGN

import fairbanks_brand as fb

HERE = Path(__file__).resolve().parent

RECIPIENT = "Mr. Alex Kivumbi"
DATE = "7 September 2026"
LETTER_REF = "FBMC/INV/2026/002-A"
PROPOSAL_REF = "FBMC/INV/2026/002"
SUBJECT = "Maternity and Inpatient Unit with Obstetric Theatre"
SIGNATORY = ("Racheal Nabukeera", "Managing Director")

L, R = ALIGN.LEFT, ALIGN.RIGHT


# --------------------------------------------------------------------------
# document 1 - cover letter
# --------------------------------------------------------------------------

def build_cover_letter(path):
    doc = fb.new_document(
        title="Cover letter - " + SUBJECT,
        subject="Investment proposal cover letter prepared for " + RECIPIENT,
        keywords="FairBanks, maternity, investment, Kivumbi",
    )
    fb.running_header(doc, "Cover letter:  " + SUBJECT)
    fb.page_footer(doc)
    fb.letterhead(doc, logo_width=1.75)

    fb.reference_line(doc, LETTER_REF, DATE, after=10)

    for line in (RECIPIENT, "Kampala"):
        p = fb.para(doc, after=0, spacing=1.0)
        fb.run(p, line, bold=(line == RECIPIENT))

    p = fb.para(doc, before=8, after=7)
    fb.run(p, "Dear Mr. Kivumbi,")

    p = fb.para(doc, after=8)
    fb.run(p, "RE:   PROPOSAL FOR INVESTMENT IN THE FAIRBANKS MATERNITY AND "
              "INPATIENT UNIT", bold=True, size=10.5, color=fb.DEEP_RGB)
    fb.rule(p, color=fb.BRAND, sz=6, edge="bottom", space=3)

    def letter_para(text, after=5):
        return fb.body(doc, text, after=after, spacing=1.06)

    letter_para("Thank you for the time you gave us when we last spoke, and for the "
                "interest you showed in what we are building at Kyebando. The full "
                "proposal is enclosed.")

    letter_para("Before you read it, one thing has changed since we spoke, and it "
                "affects the figure.")

    letter_para("We had described this project as needing between 75 and 100 million "
                "shillings. That is still right for the ward itself: the delivery room, "
                "the inpatient beds, the scanner and the staff to run them. It is not "
                "right for a maternity unit that can also operate. Having gone through "
                "it properly with our clinical team, we do not think we should build "
                "the ward and leave the theatre as a vague intention for later. A "
                "maternity unit without a theatre sends every obstructed labour and "
                "every emergency caesarean out of the gate in an ambulance, at the "
                "worst possible moment for the mother and for our name.")

    letter_para("So the enclosed proposal prices the complete unit, ward and theatre "
                "together, at between 192 and 265 million shillings, split into two "
                "phases. Phase one is the 75 to 100 million we discussed and can begin "
                "as soon as funds are released. Phase two is the theatre, and it can be "
                "committed separately once phase one is running and you have seen the "
                "numbers it actually produces.")

    letter_para("You should also know how firm these costs are. One line is a real "
                "quotation: the obstetric scanner, at 11 million shillings from Canon "
                "Healthcare Systems on 13 April this year. The rest are our own "
                "estimates, made before going to suppliers, and each will be quoted and "
                "returned to you as a signed one-page budget before any money moves.")

    letter_para("What we are asking for now is a conversation, not a commitment. Tell "
                "us the level you would be comfortable considering and how you would "
                "want a return structured, and we will build the priced budget around "
                "that. May I also suggest you come and see it? The space is there and "
                "needs finishing. Twenty minutes walking through it will tell you more "
                "than this document will, and I can arrange for you to meet the "
                "clinical team. I am available any weekday.")

    letter_para("Thank you again for your time, and for taking this seriously.", after=10)

    p = fb.para(doc, after=0)
    fb.run(p, "Yours sincerely,")
    fb.spacer(doc, 11)          # space for a wet signature
    fb.signoff_block(doc, *SIGNATORY)

    # enclosures
    fb.spacer(doc, 4)
    enclosures = (
        "Investment Proposal, Maternity and Inpatient Unit, " + DATE,
        "Proforma invoice 1896, Canon Healthcare Systems Ltd, 13 April 2026",
    )
    for i, item in enumerate(enclosures):
        p = fb.para(doc, after=1, spacing=1.0)
        p.paragraph_format.left_indent = fb.Twips(560)
        p.paragraph_format.first_line_indent = fb.Twips(-560)
        p.paragraph_format.tab_stops.add_tab_stop(fb.Twips(560))
        if i == 0:
            fb.rule(p, color=fb.LINE, sz=6, space=6)
            p.paragraph_format.space_before = fb.Pt(6)
            fb.run(p, "Enc.", bold=True, size=9, color=fb.DEEP_RGB)
        fb.run(p, "\t", size=9)
        fb.run(p, item, size=9)

    doc.save(path)
    return path


# --------------------------------------------------------------------------
# document 2 - proposal
# --------------------------------------------------------------------------

def build_proposal(path):
    doc = fb.new_document(
        title="Investment Proposal - " + SUBJECT,
        subject="Investment proposal prepared for " + RECIPIENT,
        keywords="FairBanks, maternity, obstetric theatre, investment, Kivumbi",
    )
    fb.running_header(doc, "Investment proposal:  " + SUBJECT)
    fb.page_footer(doc)
    fb.letterhead(doc)

    p = fb.para(doc, after=1, align=ALIGN.CENTER, spacing=1.0)
    fb.run(p, "INVESTMENT PROPOSAL", bold=True, size=17, color=fb.DEEP_RGB,
           font=fb.DISPLAY_FONT, spacing=18)
    p = fb.para(doc, after=3, align=ALIGN.CENTER, spacing=1.0)
    fb.run(p, SUBJECT, size=12, color=fb.BRAND_RGB)
    p = fb.para(doc, after=10, align=ALIGN.CENTER, spacing=1.0)
    fb.run(p, "Prepared for " + RECIPIENT + fb.DOT + DATE, size=9.5,
           color=fb.MUTED_RGB)

    fb.reference_line(doc, PROPOSAL_REF, DATE, after=8)

    fb.callout(doc, "In short",
               "FairBanks is asking for UGX 192-265M to complete and commission a "
               "maternity and inpatient unit with its own obstetric theatre, built in "
               "two phases. Phase one costs UGX 75-100M and can begin immediately. "
               "Phase two, the theatre, can be committed once phase one is open and "
               "reporting.", after=8)

    fb.figure_strip(doc, [
        ("UGX 192-265M", "complete programme"),
        ("UGX 75-100M", "phase one, starts now"),
        ("14-18 months", "phase one capital recovery"),
    ], after=4)

    # 1 --------------------------------------------------------------------
    fb.heading(doc, "1.", "Why maternity, and why now", before=12)

    fb.body(doc,
            "FairBanks Medical Centre operates in Kyebando-Kisalosalo, on the Northern "
            "Bypass. The centre is licensed and running: outpatient consultations, "
            "laboratory, pharmacy, X-ray, specialist clinics, physiotherapy and a "
            "24-hour emergency service, together producing between UGX 8M and UGX 12M a "
            "month. We accept CIC and AAR, with Jubilee and Mutual in process.")

    fb.body(doc,
            "Obstetrics and gynaecology is already one of our clinics. What we do not "
            "have is anywhere to admit the women who use it. An expectant mother can "
            "have her antenatal care with us, her scans with us and her medicines from "
            "our pharmacy, and then deliver somewhere else, because at the point where "
            "she needs a bed we have nothing to offer her. Every one of those "
            "deliveries is revenue we have already paid to acquire and then handed to "
            "another facility.")

    fb.body(doc,
            "The maternity and inpatient unit closes that gap. It is also the only "
            "expansion open to us that creates a genuinely new service line rather than "
            "adding capacity to one we already run.")

    # 2 --------------------------------------------------------------------
    fb.heading(doc, "2.", "What we are asking you to fund")

    fb.body(doc,
            "The work splits cleanly into two phases. Phase one produces a functioning "
            "maternity and inpatient ward. Phase two adds the operating theatre and "
            "neonatal capacity that let us keep emergencies in the building.")

    w3 = [2500, 5324, 2400]
    heads3 = ["Item", "What it covers", "Estimate"]
    a3 = [L, L, R]

    fb.subhead(doc, "Phase one: maternity and inpatient ward", before=8)
    t = fb.make_table(doc, w3)
    fb.head_row(t, w3, heads3, aligns=a3)
    fb.banded(t, w3, [
        ("Ward completion and fit-out",
         "Partitioning, flooring, plumbing, sluice room, washrooms, lighting and power "
         "points for the delivery room, postnatal room and inpatient bays",
         "UGX 22-30M"),
        ("Delivery and ward equipment",
         "Two delivery beds, ten patient beds and lockers, delivery and episiotomy "
         "sets, examination couches, vital-signs monitors, suction units, a "
         "resuscitaire and oxygen concentrators",
         "UGX 28-36M"),
        ("Obstetric ultrasound",
         "Mindray DP-10 with printer and trolley. Quoted at UGX 11,000,000 by Canon "
         "Healthcare Systems Ltd, proforma 1896 of 13 April 2026",
         "UGX 11M"),
        ("Staffing and launch",
         "Recruitment and first-quarter cost of UNMC-registered midwives and ward "
         "nursing cover, orientation, licence variation and inspection",
         "UGX 10-14M"),
        ("Opening working capital",
         "Oxytocin and essential obstetric medicines, sutures, gloves, cord clamps, "
         "linen, laundry and utilities for the first three months",
         "UGX 6-9M"),
    ], aligns=a3)
    fb.total_row(t, w3, ("Phase one total", "", "UGX 77-100M"), aligns=a3)

    fb.subhead(doc, "Phase two: obstetric theatre and neonatal care", before=12)
    t = fb.make_table(doc, w3)
    fb.head_row(t, w3, heads3, aligns=a3)
    fb.banded(t, w3, [
        ("Theatre construction and services",
         "Theatre, scrub area, sterile store and sluice; sealed walls and flooring; "
         "ventilation; theatre lighting; dedicated power with UPS and generator backup",
         "UGX 30-45M"),
        ("Anaesthesia and surgical equipment",
         "Anaesthetic machine with ventilator, operating table, theatre lights, "
         "diathermy unit, theatre and recovery monitors, theatre suction",
         "UGX 45-60M"),
        ("Instruments and sterilisation",
         "Caesarean and gynaecological instrument sets, autoclave, instrument trolleys "
         "and sterile storage",
         "UGX 14-20M"),
        ("Neonatal capacity",
         "Incubator, phototherapy unit and neonatal resuscitation equipment",
         "UGX 12-18M"),
        ("Theatre staffing and commissioning",
         "Obstetric and anaesthetic cover, theatre nurse, commissioning, clinical "
         "protocols, blood-supply arrangement and licensing for surgical services",
         "UGX 14-22M"),
    ], aligns=a3)
    fb.total_row(t, w3, ("Phase two total", "", "UGX 115-165M"), aligns=a3)
    fb.gap(doc, 8)

    fb.callout(doc, "Complete programme",
               "UGX 192-265M. Every figure above except the scanner is our own "
               "estimate, prepared before going to suppliers. They are ranges because "
               "that is honestly what we know today. No funds should be released "
               "against an estimate: each line will be quoted and returned to you as a "
               "signed one-page budget first.", fill=fb.MID)

    # 3 --------------------------------------------------------------------
    fb.heading(doc, "3.", "What the unit should earn")

    fb.body(doc,
            "The figures below are FairBanks' planning assumptions, not results. They "
            "are set out in full so you can argue with them, which is the point of "
            "showing the arithmetic rather than the conclusion.")

    fb.subhead(doc, "Phase one at steady utilisation", before=8)
    w4 = [3624, 2000, 2200, 2400]
    a4 = [L, R, R, R]
    t = fb.make_table(doc, w4)
    fb.head_row(t, w4, ["Service", "Per month", "Price (UGX)", "Revenue (UGX)"],
                aligns=a4)
    fb.banded(t, w4, [
        ("Antenatal visits", "100", "25,000", "2,500,000"),
        ("Normal deliveries", "18", "450,000", "8,100,000"),
        ("Obstetric ultrasound scans", "70", "50,000", "3,500,000"),
        ("Inpatient bed-days", "70", "70,000", "4,900,000"),
    ], aligns=a4)
    fb.total_row(t, w4, ("Monthly revenue", "", "", "19,000,000"), aligns=a4)

    fb.subhead(doc, "Monthly running costs", before=12)
    w2 = [7824, 2400]
    a2 = [L, R]
    t = fb.make_table(doc, w2)
    fb.head_row(t, w2, ["Cost", "UGX"], aligns=a2)
    fb.banded(t, w2, [
        ("Midwifery and ward nursing: four midwives, a ward in-charge, two support "
         "staff", "5,500,000"),
        ("Medicines, sutures and clinical consumables", "2,800,000"),
        ("Utilities, standby power, water and clinical waste disposal", "1,200,000"),
        ("Laundry, cleaning, maintenance and equipment servicing", "700,000"),
    ], aligns=a2)
    fb.total_row(t, w2, ("Monthly running costs", "10,200,000"), aligns=a2)
    fb.gap(doc, 8)

    fb.callout(doc, "Phase one monthly contribution",
               "UGX 19,000,000 less UGX 10,200,000 = UGX 8,800,000 a month, once "
               "utilisation settles.", fill=fb.MID)

    fb.subhead(doc, "What the theatre adds", before=8)
    t = fb.make_table(doc, w2)
    fb.head_row(t, w2, ["Phase two", "UGX"], aligns=a2)
    fb.banded(t, w2, [
        ("Caesarean sections, 5 a month at UGX 1,700,000", "8,500,000"),
        ("Less obstetric and anaesthetic session fees, 5 cases", "(2,500,000)"),
        ("Less theatre nurse and sterilisation cover", "(1,400,000)"),
        ("Less theatre consumables, drugs and blood handling", "(1,300,000)"),
        ("Less theatre maintenance and equipment servicing", "(600,000)"),
    ], aligns=a2)
    fb.total_row(t, w2, ("Additional monthly contribution", "2,700,000"), aligns=a2)
    fb.gap(doc, 4)

    fb.body(doc,
            "That takes the combined contribution to UGX 11,500,000 a month. We have "
            "not modelled gynaecological use of the theatre, which in practice will "
            "carry a share of the load and improve the figure.", before=6)

    fb.body(doc,
            "At those volumes phase one returns its own capital in about ten months of "
            "steady operation. Ten months of steady operation is not ten months from "
            "the day you release funds. Construction, recruitment, licensing and the "
            "time an antenatal book takes to fill mean the first six months will run "
            "below the table above. Our working expectation is capital recovery on "
            "phase one in fourteen to eighteen months, and on the full programme "
            "including the theatre in two and a half to three years.")

    fb.callout(doc, "Five numbers to replace before this is final",
               "Current monthly antenatal attendance · the delivery fee we intend to "
               "charge · the bed count we are actually building to · our intended "
               "bed-day rate · the caesarean fee. Those five drive everything else on "
               "this page, and FairBanks' own figures should replace the assumptions "
               "above before any agreement is signed.", accent=fb.ORANGE)

    # 4 --------------------------------------------------------------------
    fb.heading(doc, "4.", "What could go wrong")

    fb.body(doc, "We would rather set these out ourselves than have you find them.")

    for lead, text in [
        ("Utilisation.",
         "A delivery room fills from the antenatal book, not from advertising. If "
         "antenatal attendance does not grow, the delivery numbers above do not happen. "
         "That is exactly why phase one puts the ward and the scanner in first: you get "
         "six months of real figures before the larger commitment."),
        ("Clinical risk.",
         "Obstetrics carries the highest liability of anything we do. The centre "
         "already holds professional indemnity and all-risks cover, and both would be "
         "reviewed and increased before the unit opens. Protocols, partograph "
         "discipline, emergency drills and genuine 24-hour cover are conditions of "
         "opening, not refinements to add later."),
        ("The gap before the theatre.",
         "Until phase two is commissioned, caesareans and complications have to go out. "
         "That needs a written referral arrangement with a partner hospital and a "
         "transport plan that works at three in the morning, both in place before the "
         "first delivery."),
        ("Regulation.",
         "A maternity unit with a theatre needs the facility licence varied to cover "
         "that level of care, UNMC-registered midwives, UMDPC-registered clinicians and "
         "compliant clinical waste handling. We would not spend phase two money before "
         "the licence position is confirmed in writing."),
        ("Cost.",
         "Every figure in this document except the scanner is an estimate made before "
         "quotation. Firm prices could come back above these ranges, most likely on "
         "anaesthesia equipment, which is imported and moves with the exchange rate."),
        ("Staffing.",
         "Anaesthetic officers and obstetricians are scarce and they move. We propose "
         "sessional arrangements in the first year rather than full-time salaries, "
         "which keeps the cost variable and the downside smaller."),
    ]:
        fb.bullet(doc, lead, text)

    # 5 --------------------------------------------------------------------
    fb.heading(doc, "5.", "How your money would be protected")

    for lead, text in [
        ("Separate books.",
         "The unit is accounted for apart from the rest of the centre, with its own "
         "revenue and cost ledger from the first day."),
        ("Release against quotations.",
         "Funds move in tranches against signed supplier quotations and delivery notes, "
         "never against the estimates in this document."),
        ("A one-page report every month.",
         "Capital deployed, antenatal visits, deliveries, theatre cases, revenue, "
         "direct costs, monthly contribution and capital recovered to date."),
        ("Your own eyes.",
         "You may nominate someone to inspect the books, the store and the equipment at "
         "any time, without notice."),
        ("Registered assets.",
         "Capital equipment is asset-tagged, insured and listed, and can be named in "
         "the investment agreement as security."),
    ]:
        fb.bullet(doc, lead, text)

    fb.body(doc,
            "FairBanks Medical Centre Limited is a registered Ugandan company, TIN "
            + fb.TIN + ". The centre holds a current operating licence, a National Drug "
            "Authority licence for the pharmacy, a certificate of suitability of "
            "premises, NSSF registration, and professional indemnity and all-risks "
            "insurance. Copies of all of these are available for your review before you "
            "commit anything.", before=6)

    # 6 --------------------------------------------------------------------
    fb.heading(doc, "6.", "How the investment could be structured")

    for lead, text in [
        ("A project loan.",
         "You lend a fixed amount for a fixed term at an agreed rate, repaid from the "
         "unit's contribution. Simplest to document, and your return does not depend on "
         "how well we run the place."),
        ("Revenue share until recovery, plus a premium.",
         "An agreed share of the unit's monthly contribution comes to you until your "
         "capital and an agreed premium are repaid, after which the arrangement ends. "
         "Your return tracks performance, and repayment only starts once there is "
         "something to share."),
        ("Equity.",
         "A shareholding in the unit or in the company, giving you a permanent share of "
         "the upside and a voice in the decisions."),
    ]:
        fb.bullet(doc, lead, text)

    fb.body(doc,
            "Our own preference is the second. It ties what you earn to how well we run "
            "the unit, it does not saddle the centre with fixed repayments during the "
            "months when utilisation is still building, and it has a defined end. That "
            "said, we are open on this. The structure should be the one that suits you.",
            before=6)

    # 7 --------------------------------------------------------------------
    fb.heading(doc, "7.", "What we would like to agree")

    for i, step in enumerate([
        "You tell us the level of investment you would be comfortable considering, and "
        "which structure you would prefer.",
        "We put every line in section 2 to suppliers and come back within three weeks "
        "with a priced one-page budget, a tranche schedule and a cash-flow projection.",
        "We agree the return formula, the reporting format and the release conditions "
        "in writing.",
        "The first tranche is released and phase one begins.",
    ], 1):
        fb.numbered(doc, i, step)

    fb.body(doc,
            "We would be grateful for a meeting at the centre within the next two "
            "weeks, so you can see the space before you decide anything.", before=4)

    fb.closing_block(doc)

    doc.save(path)
    return path


if __name__ == "__main__":
    outputs = [
        build_cover_letter(HERE / "FairBanks_Maternity_Cover_Letter_Kivumbi.docx"),
        build_proposal(HERE / "FairBanks_Maternity_Proposal_Kivumbi.docx"),
    ]
    for f in outputs:
        print("wrote", f.name, f.stat().st_size, "bytes")
