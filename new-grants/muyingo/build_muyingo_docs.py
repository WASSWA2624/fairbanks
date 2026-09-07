"""
Builds the FairBanks diagnostics and pharmacy investment proposal for
Dr. Martin Muyingo.

One output, written next to this script:

    FairBanks_Diagnostics_Proposal_Muyingo.docx

All styling comes from ../fairbanks_brand.py, shared with the Kivumbi pack so
the letterhead, palette and page furniture match across new-grants.

The wording is carried over verbatim from the original hand-built document;
this script exists so the proposal can be regenerated and stays branded
consistently with the rest of the pack.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx.enum.text import WD_ALIGN_PARAGRAPH as ALIGN

import fairbanks_brand as fb

HERE = Path(__file__).resolve().parent

RECIPIENT = "Dr. Martin Muyingo"
DATE = "7 September 2026"
REF = "FBMC/INV/2026/001"
SUBJECT = "Ultrasound, Laboratory and Pharmacy"
SIGNATORY = ("Racheal Nabukeera", "Managing Director and Co-founder")

L, R, C = ALIGN.LEFT, ALIGN.RIGHT, ALIGN.CENTER


def build_proposal(path):
    doc = fb.new_document(
        title="Investment Proposal - " + SUBJECT,
        subject="Investment proposal prepared for " + RECIPIENT,
        keywords="FairBanks, ultrasound, laboratory, pharmacy, investment, Muyingo",
    )
    fb.running_header(doc, "Investment proposal:  " + SUBJECT)
    fb.page_footer(doc)
    fb.letterhead(doc)

    p = fb.para(doc, after=1, align=C, spacing=1.0)
    fb.run(p, "INVESTMENT PROPOSAL", bold=True, size=17, color=fb.DEEP_RGB,
           font=fb.DISPLAY_FONT, spacing=18)
    p = fb.para(doc, after=3, align=C, spacing=1.0)
    fb.run(p, "Ultrasound, Laboratory & Pharmacy", size=12, color=fb.BRAND_RGB)
    p = fb.para(doc, after=10, align=C, spacing=1.0)
    fb.run(p, "Prepared for " + RECIPIENT + fb.DOT + DATE, size=9.5,
           color=fb.MUTED_RGB)

    fb.reference_line(doc, REF, DATE, after=8)

    fb.callout(doc, "The ask in one paragraph",
               "UGX 83.4 million to add ultrasound imaging, automate our laboratory "
               "bench and put real depth into the pharmacy. Of that, UGX 29.9 million "
               "is already priced on a supplier proforma for the equipment itself. Once "
               "the three lines are running we expect them to contribute UGX 6.3-13.3 "
               "million a month against a facility that currently turns over UGX 8-12 "
               "million in total.", after=8)

    fb.figure_strip(doc, [
        ("UGX 83.4M", "total ask"),
        ("UGX 29.9M", "already quoted"),
        ("7-14 months", "capital recovery at run rate"),
    ], after=4)

    # 1 --------------------------------------------------------------------
    fb.heading(doc, "1.", "Why we are bringing this to you", before=12)

    fb.body(doc,
            "Dr. Muyingo, you understand the referral problem in Kampala better than "
            "most of the people we could have taken this to. A patient walks into a "
            "clinic on the Northern Bypass, needs a scan or a liver panel, and walks "
            "out holding a piece of paper that sends them somewhere else. Some of them "
            "come back to us. A good number do not, and we lose both the revenue and "
            "the thread of their care.")

    fb.body(doc,
            "That is what we are asking you to help us fix. Not in the abstract, but "
            "through three specific purchases we have already started pricing, on a "
            "defined budget, with monthly reporting you can hold us to.")

    # 2 --------------------------------------------------------------------
    fb.heading(doc, "2.", "Where FairBanks stands today")

    fb.body(doc,
            "FairBanks is an operating community-based medical centre in "
            "Kyebando-Kisalosalo, on the Northern Bypass corridor. We run general "
            "outpatient and inpatient care, a 24-hour emergency and trauma service, a "
            "laboratory, a licensed retail pharmacy, and specialist clinics covering "
            "internal medicine, paediatrics, obstetrics and gynaecology, ENT, "
            "ophthalmology, urology and physiotherapy. We are on the CIC and AAR "
            "insurance panels, with Jubilee and Mutual in progress. Alongside the "
            "clinical service we run GeriCare elderly care, community screenings and "
            "health awareness work.")

    fb.body(doc,
            "The facility currently turns over UGX 8-12 million a month. That is a "
            "working business with a patient base, a licence and a payroll, not a "
            "start-up asking you to fund an idea. What it cannot yet do is finish a "
            "diagnosis under its own roof.")

    fb.label_panels(doc, [
        ("Current monthly revenue", "UGX 8-12M"),
        ("What we already run",
         "OPD, inpatient, 24-hour emergency, basic laboratory, pharmacy, seven "
         "specialist clinics"),
        ("What we cannot do",
         "Any imaging on site, automated chemistry and immunoassay, consistent "
         "pharmacy stock"),
    ], after=4)

    # 3 --------------------------------------------------------------------
    fb.heading(doc, "3.", "The three gaps")

    fb.body(doc, lead="Imaging.", text=(
        "We have none. Every patient needing a scan, whether an antenatal check, "
        "abdominal pain or pelvic and soft-tissue work, leaves the building for it, and "
        "we refer out work our own consultation rooms generate. Ultrasound is the right "
        "place to start: it covers the largest share of what our clinicians actually "
        "order, and unlike X-ray it needs no radiation licence, no shielded room and no "
        "radiographer. A Mindray DP-10 and a trolley close that gap for UGX 11 million."))

    fb.body(doc, lead="Laboratory.", text=(
        "Our bench copes with routine work. Anything needing chemistry or immunoassay, "
        "such as liver and renal function, lipid profiles, thyroid, cardiac markers and "
        "hormonal panels, goes out. That costs the patient a day and costs us the test. "
        "A dry chemistry analyser and a Fincare immunoassay unit bring the work in-house "
        "and cut turnaround from days to the same visit."))

    fb.body(doc, lead="Pharmacy.", text=(
        "The pharmacy is licensed and trading, but thin stock means we regularly write a "
        "prescription and watch it filled two doors down. This is a working-capital "
        "problem rather than an equipment problem, and it is the cheapest of the three "
        "to solve."))

    # 4 --------------------------------------------------------------------
    fb.heading(doc, "4.", "What the money buys")

    w4 = [2350, 4374, 2100, 1400]
    a4 = [L, L, L, R]
    t = fb.make_table(doc, w4)
    fb.head_row(t, w4, ["Item", "Detail", "Status", "UGX"], aligns=a4)
    fb.banded(t, w4, [
        ("Diagnostic equipment",
         "Mindray DP-10 ultrasound with printer, ultrasound trolley, dry chemistry "
         "analyser, Fincare immunoassay analyser, 2000W and 1000W stabilisers",
         "Quoted: Canon Healthcare Systems, proforma 1896", "29,900,000"),
        ("Pharmacy stock depth",
         "Opening working capital to carry a fuller formulary and stop losing filled "
         "prescriptions to neighbouring outlets",
         "Estimated from our own dispensing records", "18,000,000"),
        ("Laboratory reagents and consumables",
         "Opening three-month fill for the chemistry and immunoassay analysers, plus "
         "controls and quality-control material",
         "To be quoted with the analysers", "9,200,000"),
        ("Staffing",
         "Sonographer and laboratory technologist: recruitment and first two months of "
         "salary",
         "Estimated at current Kampala rates", "7,400,000"),
        ("Ultrasound room",
         "Couch, screening, cabinetry, lighting, dedicated power point, minor works, "
         "installation and commissioning",
         "Estimated: local contractor", "6,500,000"),
        ("Power protection and cold chain",
         "UPS and inverter cover for the analysers, and a dedicated reagent refrigerator",
         "To be quoted", "3,800,000"),
        ("Licensing, training and launch",
         "Practitioner registration, NDA and premises compliance, operator training, "
         "opening publicity",
         "Estimated", "2,600,000"),
        ("Contingency (approx. 8%)",
         "Exchange movement, freight and scope corrections", "-", "6,000,000"),
    ], aligns=a4)
    fb.total_row(t, w4, ("Total", "", "", "83,400,000"), aligns=a4)
    fb.gap(doc, 6)

    fb.body(doc, lead="About the quotation.", text=(
        "The equipment line is not a guess. Canon Healthcare Systems Ltd of "
        "Kireka-Namugongo Road issued us proforma invoice 1896 on 13 April 2026 covering "
        "all six items at UGX 29,900,000, with the equipment held in stock and payment "
        "due on delivery against a local purchase order. That quotation carried 60 days "
        "of validity and has therefore lapsed, so we would re-confirm pricing before "
        "placing any order, which is part of what the contingency line is for. It is "
        "reproduced in full in the appendix. Every other figure in the table is our own "
        "estimate and should be read as one until it has been quoted."))

    # 5 --------------------------------------------------------------------
    fb.heading(doc, "5.", "What it earns")

    fb.body(doc,
            "We have set out the assumptions rather than a single headline number, "
            "because the assumptions are the part worth arguing about.")

    fb.body(doc, lead="Ultrasound.", text=(
        "120 to 180 scans a month at UGX 30,000-40,000. That is six or seven scans a "
        "working day, which is modest for a centre already running antenatal and general "
        "outpatient clinics."))

    fb.body(doc, lead="Laboratory.", text=(
        "350 to 500 additional in-house tests a month at an average of UGX "
        "18,000-25,000. This is work we already generate and currently refer out for "
        "nothing."))

    fb.body(doc, lead="Pharmacy.", text=(
        "UGX 4-6 million of additional monthly sales at roughly 30% gross margin. "
        "Pharmacy is a margin business, so its contribution looks small beside its "
        "revenue. That is normal, and we would rather show it plainly than dress it up."))

    w5 = [2300, 1900, 3924, 2100]
    a5 = [L, R, L, R]
    t = fb.make_table(doc, w5)
    fb.head_row(t, w5, ["Service line", "Monthly revenue", "Direct costs",
                        "Monthly contribution"], aligns=a5)
    fb.banded(t, w5, [
        ("Ultrasound", "UGX 3.6-7.2M",
         "UGX 1.2-1.6M: sonographer, consumables, power, service", "UGX 2.4-5.6M"),
        ("Laboratory (incremental)", "UGX 6.3-12.5M",
         "UGX 3.6-6.6M: reagents at 40-45%, technologist, QC, service", "UGX 2.7-5.9M"),
        ("Pharmacy (incremental)", "UGX 4.0-6.0M",
         "UGX 2.8-4.2M: cost of goods at roughly 70%", "UGX 1.2-1.8M"),
    ], aligns=a5)
    fb.total_row(t, w5, ("Combined", "UGX 13.9-25.7M", "UGX 7.6-12.4M",
                         "UGX 6.3-13.3M"), aligns=a5)
    fb.keep_together(t)
    fb.gap(doc, 7)

    fb.figure_strip(doc, [
        ("UGX 83.4M", "invested"),
        ("UGX 6.3-13.3M", "contributed each month"),
        ("7-14 months", "capital recovered at run rate"),
    ], after=6)

    fb.body(doc,
            "Two caveats we would rather state than bury. First, none of this happens in "
            "month one. We would expect two to three months of ramp-up before the lines "
            "reach run rate, so measured from the first patient the conservative case is "
            "closer to sixteen months than thirteen. Second, these are our own numbers, "
            "built from our dispensing and referral records rather than from published "
            "benchmarks. We are happy to open those records to you before you commit "
            "anything.")

    # 6 --------------------------------------------------------------------
    fb.heading(doc, "6.", "What could go wrong")

    fb.body(doc,
            "A proposal that lists no risks is not being straight with you. These are "
            "the six we actually worry about.")

    w6 = [2500, 3862, 3862]
    a6 = [L, L, L]
    t = fb.make_table(doc, w6)
    fb.head_row(t, w6, ["Risk", "Why it matters", "How we would handle it"], aligns=a6)
    fb.banded(t, w6, [
        ("Utilisation below assumption",
         "The whole case rests on scan and test volumes we have estimated rather than "
         "measured",
         "Release the funds in two tranches, and make the second conditional on three "
         "months of actual volumes from the first"),
        ("Reagent lock-in",
         "Closed-system analysers tie us to one supplier's reagents, and a price rise "
         "there comes straight out of the laboratory margin",
         "Agree reagent pricing and lead times with Canon in writing before the purchase "
         "order is issued, not after the machines are installed"),
        ("Equipment price movement",
         "The proforma has lapsed and the shilling moves against the dollar",
         "Re-confirm pricing immediately before ordering. The 8% contingency is carried "
         "in the budget for exactly this"),
        ("Losing the sonographer",
         "Sonographers are scarce in Kampala and move for small increases",
         "Written notice terms in the contract, and a second operator cross-trained from "
         "existing clinical staff so the service does not stop"),
        ("Power interruption",
         "Analysers are intolerant of unstable or interrupted supply, and a spoiled run "
         "wastes reagent",
         "Stabilisers are already in the equipment order, and UPS and inverter cover "
         "sits as a separate budget line"),
        ("Regulatory delay",
         "Practitioner registration and premises compliance take time and can stall a "
         "launch after the money is spent",
         "Racheal handles these relationships directly through UMDPC, the Allied Health "
         "Council and NDA, and would start them in parallel with procurement"),
    ], aligns=a6)
    fb.gap(doc, 4)

    # 7 --------------------------------------------------------------------
    fb.heading(doc, "7.", "How your investment could be structured")

    fb.body(doc,
            "We have deliberately not assumed a structure. Three would work for us, and "
            "we are open to a fourth if you would prefer something else.")

    for lead, text in [
        ("A project loan with an agreed return.",
         "You advance the capital and we repay it out of the monthly contribution over "
         "an agreed period at an agreed premium. Simplest to document, and cleanest for "
         "you to exit."),
        ("A revenue share on the funded lines.",
         "You take an agreed percentage of the ultrasound, laboratory and pharmacy "
         "contribution until capital and return are recovered, after which the share "
         "steps down or ends. Your return then tracks performance in both directions."),
        ("Equity in the expansion.",
         "A shareholding negotiated on valuation, if you would rather hold a long "
         "position in FairBanks than simply be repaid."),
    ]:
        fb.bullet(doc, lead, text)

    fb.body(doc,
            "Whichever we settle on, the funded lines would be reported separately from "
            "the rest of the facility, every month, on a single page: capital drawn, "
            "procedures performed, revenue, direct costs, contribution and cumulative "
            "recovery. You would be looking at the same numbers we are, on the same day.",
            before=4)

    fb.label_panels(doc, [
        ("Capital drawn", None),
        ("Procedures done", None),
        ("Revenue & costs", None),
        ("Cumulative recovery", None),
    ], center=True, after=4)

    # 8 --------------------------------------------------------------------
    fb.heading(doc, "8.", "What we are asking for")

    for text in [
        "One conversation at the centre, so you can see the room, the pharmacy and the "
        "laboratory bench before you form a view.",
        "An indication of the capital you would be comfortable placing. If UGX 83.4M is "
        "more than you want to commit at once, the equipment order stands perfectly well "
        "on its own at UGX 29.9M, and reagents and pharmacy stock can follow out of what "
        "it earns.",
        "If the conversation goes well, we turn this into a one-page deployment budget "
        "and an agreed structure within two weeks of your indication.",
    ]:
        fb.bullet(doc, None, text)

    fb.body(doc,
            "We are not asking for an open-ended injection into the business. We are "
            "asking you to fund a defined list of equipment and stock, and then to hold "
            "us to the numbers it produces.", before=4)

    fb.statement(doc, "FINISH THE DIAGNOSIS UNDER ONE ROOF.", after=10)

    fb.signoff_block(doc, *SIGNATORY)

    # appendix ------------------------------------------------------------
    # The quotation is evidence, not argument: it gets its own page rather
    # than trailing two lines onto the end of the proposal.
    fb.page_break(doc)
    fb.heading(doc, None, "Appendix: supplier quotation as received", before=0)

    fb.body(doc,
            "Canon Healthcare Systems Ltd, SK House, Kireka-Namugongo Road, Kampala. "
            "Proforma invoice 1896, dated 13 April 2026.", after=7)

    wa = [700, 5024, 900, 1800, 1800]
    aa = [C, L, C, R, R]
    t = fb.make_table(doc, wa)
    fb.head_row(t, wa, ["#", "Item", "Qty", "Unit price (UGX)", "Total (UGX)"],
                aligns=aa)
    fb.banded(t, wa, [
        ("1", "Mindray DP-10 ultrasound machine with printer", "1",
         "10,000,000", "10,000,000"),
        ("2", "Ultrasound trolley", "1", "1,000,000", "1,000,000"),
        ("3", "Dry chemistry analyser", "1", "14,000,000", "14,000,000"),
        ("4", "Fincare immunoassay analyser", "1", "4,000,000", "4,000,000"),
        ("5", "Stabiliser, 2000W", "1", "600,000", "600,000"),
        ("6", "Stabiliser, 1000W", "1", "300,000", "300,000"),
    ], aligns=aa)
    fb.total_row(t, wa, ("", "Total", "", "", "29,900,000"), aligns=aa)
    fb.keep_together(t)
    fb.gap(doc, 6)

    fb.body(doc,
            "Terms as quoted: prices valid for 60 days from the quotation date, now "
            "lapsed; local purchase order to confirm the order; 100% cash or "
            "current-dated cheque on delivery; items stated as available in stock.",
            size=9.5)

    fb.closing_block(doc)

    doc.save(path)
    return path


if __name__ == "__main__":
    f = build_proposal(HERE / "FairBanks_Diagnostics_Proposal_Muyingo.docx")
    print("wrote", f.name, f.stat().st_size, "bytes")
