#!/usr/bin/env python3
"""Fill CDC-RFA-JG-26-0054 official Grants.gov application documents.

The shells in official-application-docs/ are Adobe XFA forms (need Adobe Reader
or Grants.gov webforms). This script builds filled, upload-ready companion PDFs
plus a field map for webform paste.

Output: official-application-docs/filled/
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

HERE = Path(__file__).resolve().parent
OUT = HERE / "official-application-docs" / "filled"
ANSWERS = HERE / "application_answers.md"

# --- Applicant facts ---
ORG = "FAIRBANKS MEDICAL CENTRE LIMITED"
ORG_SHORT = "FairBanks Medical Centre"
UEI = "[CONFIRM - paste UEI from SAM.gov]"
EIN = "N/A - foreign entity (Uganda TIN 1053370026)"
COMPANY_NO = "80020003843337"
TIN = "1053370026"
NSSF = "NS043295"
STREET = "Plot 1423 and 1425 Tirupati Road, Fairbanks Medical Centre"
CITY = "Kampala"
COUNTY = "Kampala Central Division, Kololo IV, Lugogo"
STATE = "N/A"
PROVINCE = "Kampala"
COUNTRY = "UGANDA"
ZIP = "N/A"
PHONE = "+256772849258"
EMAIL = "info@fairbanksmedicalcentre.org"
WEBSITE = "https://www.fairbanksmedicalcentre.org/"
PD_NAME = "Racheal Nabukeera"
PD_TITLE = "Managing Director and Co-founder"
OPPORTUNITY = "CDC-RFA-JG-26-0054"
CFDA = "93.318"
CFDA_TITLE = "Protecting and Improving Health Globally: Building and Strengthening Public Health Impact, Systems, Capacity, and Security"
TITLE = (
    "Last-mile Global Health Security for Uganda: FairBanks will help CHWs, VHTs, "
    "clinics, and districts put timely outbreak and priority-disease signals into "
    "Ministry of Health systems (including NISS-linked pathways), train frontline "
    "workers, and practise 7-1-7 style detect-notify-respond loops so threats are "
    "contained closer to source."
)
START = "09/30/2026"
END = "09/29/2031"
TODAY = date.today().strftime("%m/%d/%Y")
SLOGAN = "Your health, our mission."

# Year 1 component asks (USD) - CONFIRM before submit
C1, C2, C3, C4, C5 = 2_450_000, 1_800_000, 2_750_000, 1_500_000, 1_500_000
TOTAL = C1 + C2 + C3 + C4 + C5  # 10,000,000

# Component 1 category split
C1_CAT = {
    "Personnel": 720_000,
    "Fringe Benefits": 144_000,
    "Travel": 85_000,
    "Equipment": 60_000,
    "Supplies": 95_000,
    "Contractual": 780_000,
    "Construction": 0,
    "Other": 280_000,
    "Indirect Charges": 286_000,
}


def money(n: int) -> str:
    return f"${n:,.0f}"


def styles():
    st = getSampleStyleSheet()
    specs = [
        ("Cover", dict(fontName="Times-Bold", fontSize=14, leading=17, alignment=TA_CENTER, spaceAfter=8)),
        ("H1", dict(fontName="Times-Bold", fontSize=12, leading=15, spaceBefore=10, spaceAfter=6)),
        ("H2", dict(fontName="Times-Bold", fontSize=11, leading=14, spaceBefore=8, spaceAfter=4)),
        ("Body", dict(fontName="Times-Roman", fontSize=11, leading=14, alignment=TA_JUSTIFY, spaceAfter=6)),
        ("Small", dict(fontName="Times-Roman", fontSize=9, leading=11, spaceAfter=4)),
        ("Meta", dict(fontName="Times-Italic", fontSize=9, leading=11, textColor=colors.HexColor("#444444"), spaceAfter=6)),
        ("Cell", dict(fontName="Times-Roman", fontSize=8, leading=10)),
        ("CellB", dict(fontName="Times-Bold", fontSize=8, leading=10)),
        ("Center", dict(fontName="Times-Roman", fontSize=11, leading=14, alignment=TA_CENTER, spaceAfter=4)),
    ]
    for name, kw in specs:
        st.add(ParagraphStyle(name=name, **kw))
    return st


def doc_template(path: Path, title: str):
    path.parent.mkdir(parents=True, exist_ok=True)

    def _page(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Roman", 8)
        canvas.drawString(0.75 * inch, 0.45 * inch, f"{ORG_SHORT} | {OPPORTUNITY} | {title}")
        canvas.drawRightString(letter[0] - 0.75 * inch, 0.45 * inch, f"Page {doc.page}")
        canvas.restoreState()

    return SimpleDocTemplate(
        str(path),
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title=title,
        author=ORG,
    ), _page


def tbl(st, headers, rows, widths=None):
    data = [[Paragraph(h, st["CellB"]) for h in headers]]
    for row in rows:
        data.append([Paragraph(str(c), st["Cell"]) for c in row])
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEEE")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    return t


def bullets(st, items):
    return ListFlowable(
        [ListItem(Paragraph(i, st["Body"]), leftIndent=8, value="•") for i in items],
        bulletType="bullet",
        start="•",
    )


ABSTRACT = (
    f"{ORG} ({ORG_SHORT}) is a Uganda-registered Family & Community Health organisation "
    f"in Kampala, guided by The FairBanks Blueprint: Building the Future of Family & "
    f"Community Health. We run a licensed medical centre and FairBanks Community Reach with "
    f"CHWs and VHTs in Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya and nearby peri-urban "
    f"communities — extending care beyond consultation rooms into homes, schools, and "
    f"community spaces. Slogan: {SLOGAN} "
    "Strong health systems are built on strong local institutions. Uganda's 2023 Joint "
    "External Evaluation showed progress — and still named hard gaps: community and facility "
    "signals that reach decision-makers too late; weak information sharing across sectors; "
    "limited surge-ready workforce at subnational level; and incomplete links into national "
    "surveillance and emergency structures, including work toward a functional National "
    "Integrated Surveillance System (NISS). When those gaps stay open, outbreaks grow before "
    "anyone can act. That puts Ugandans at risk first — and Americans at risk later. "
    "Under Ministry of Health leadership, this project puts the Blueprint into practice for "
    "Global Health Security: we will not rebuild Uganda's national architecture; we will make "
    "community and clinic signals usable inside it. "
    "Year 1 focus (Component 1 — Core GHS) has four practical pillars aligned to FairBanks' "
    "Community Health Intelligence Platform (FCHIP) and Community Reach model: (1) community "
    "and facility surveillance that contributes to MoH/NISS pathways; (2) faster "
    "community-to-district detection and response using 7-1-7 style timing; (3) CHW/VHT and "
    "frontline workforce skills for One Health-aware surveillance and surge readiness; and "
    "(4) stronger community-facility links for priority disease programmes (including HIV, "
    "TB, malaria, cholera, viral hemorrhagic fevers, mpox, and immunisation) during routine "
    "work and outbreaks. Laboratory and border-health work will be done with MoH-aligned "
    "partners rather than as a FairBanks-only national lead. "
    "Components 2-5 set out contingency surge plans for small and large outbreaks, emerging "
    "infectious threats, and humanitarian emergencies. They may be approved but unfunded "
    "until CDC releases emergency funds. "
    "Expected shared results: shorter detection and response times; more complete community "
    "signals in national channels; a better-prepared frontline workforce; and clear handoff "
    "of tools and data to government systems — advancing healthier families and communities "
    f"while containing threats closer to source. Contact: {PD_NAME}, {EMAIL}, {PHONE}."
)


def build_readme():
    text = f"""# Filled application documents — CDC-RFA-JG-26-0054

Generated: {date.today().isoformat()}
Applicant: {ORG}

## Important

The original Grants.gov files in the parent folder are **Adobe XFA shells**.
They only open/fill properly in **Adobe Reader** or the **Grants.gov online webform**.
Python tools cannot reliably write into those shells.

This `filled/` folder has **complete companion PDFs** with FairBanks answers ready to:
1. Copy into Grants.gov webforms / Adobe Reader fields, or
2. Upload as attachments where the shell asks for a file (Project Narrative, Budget Narrative).

## Files (names match Grants.gov package shells)

| File | Use |
|---|---|
| `WS01739425-SF424_4_0-V4.0.pdf` | Copy into SF-424 (V4.0) |
| `WS01739425-SF424A-V1.0.pdf` | Copy into SF-424A (Components 1-4 + note for Component 5 / second form) |
| `WS01739425-SFLLL_2_0-V2.0.pdf` | Copy into SF-LLL (no lobbying activities) |
| `WS01739425-Project_AbstractSummary_2_0-V2.0.pdf` | Paste/upload into Project Abstract Summary |
| `WS01739425-ProjectNarrativeAttachments_1_2-V1.2.pdf` | **Attach** to Project Narrative Attachments form |
| `WS01739425-BudgetNarrativeAttachments_1_2-V1.2.pdf` | **Attach** to Budget Narrative Attachments form |
| `Local_Partner_Preference_Letter.pdf` | Attach with company evidence |
| `FIELD_MAP.txt` | Box-by-box paste guide |

Also mirrors `WS01739423-SF424_4_0-V4.0.pdf` (same SF-424 content as the WS01739425 SF-424).

## CONFIRM before submit

- Paste live **UEI** from SAM.gov everywhere marked CONFIRM
- Ownership / staff / board citizenship % for local partner letter
- Final budget numbers (draft Year 1 Component 1 = {money(C1)})
- Authorised official signature / date in Grants.gov

## Do not edit

`PKG00293651-instructions.pdf` — NOFO instructions only (not a form to fill).
"""
    (OUT / "README.md").write_text(text, encoding="utf-8")


def build_field_map():
    lines = [
        f"CDC-RFA-JG-26-0054 FIELD MAP — {ORG}",
        f"Generated {date.today().isoformat()}",
        "",
        "=== SF-424 ===",
        "1. Type of Submission: Application",
        "2. Type of Application: New",
        f"5a. Federal Entity Identifier: {OPPORTUNITY}",
        f"5b. Federal Award Identifier: (leave blank for new)",
        f"8a. Legal Name: {ORG}",
        f"8b. Employer/Taxpayer ID: {EIN}",
        f"8c. UEI: {UEI}",
        f"8d. Address: {STREET}; {CITY}; {COUNTY}; {PROVINCE}; {COUNTRY}; {ZIP}",
        "8e. Organizational Unit: FairBanks Community Reach / GHS Programme",
        f"8f. Name and contact: {PD_NAME}; {PD_TITLE}; {PHONE}; {EMAIL}",
        "9. Type of Applicant: X. For-Profit Organization (Other than Small Business) OR foreign entity as listed",
        "   Also applicable: Foreign Organization / Non-domestic (non-U.S.) Entity",
        f"10. Name of Federal Agency: Centers for Disease Control and Prevention",
        f"11. CFDA/Assistance Listing: {CFDA} — {CFDA_TITLE}",
        f"12. Funding Opportunity Number: {OPPORTUNITY}",
        "    Title: Strengthening global health security through local partnerships in Uganda",
        "13. Competition ID: (if shown on package — leave blank if N/A)",
        "14. Areas Affected: Uganda (Kampala peri-urban catchments; expand under MoH guidance)",
        f"15. Descriptive Title: {TITLE}",
        f"16. Congressional Districts: Applicant N/A (foreign); Program N/A (Uganda)",
        f"17. Proposed Project: Start {START}  Ending {END}",
        "18. Estimated Funding:",
        f"    a. Federal {money(TOTAL)}",
        "    b. Applicant $0",
        "    c. State $0",
        "    d. Local $0",
        "    e. Other $0",
        "    f. Program Income $0",
        f"    g. TOTAL {money(TOTAL)}",
        "19. Is Application Subject to Review by State Under EO 12372? c. Program is not covered by EO 12372",
        "20. Delinquent on Federal Debt? No",
        f"21. Authorized Representative: {PD_NAME}; {PD_TITLE}; {PHONE}; {EMAIL}; Date {TODAY}; Signature: sign in Grants.gov",
        "",
        "=== SF-424A Year 1 ===",
        f"Section A — Component 1 Core GHS federal {money(C1)}",
        f"Section A — Component 2 Small-scale response federal {money(C2)}",
        f"Section A — Component 3 Large-scale response federal {money(C3)}",
        f"Section A — Component 4 Emerging threats federal {money(C4)}",
        f"Component 5 Humanitarian {money(C5)} — put on second SF-424A if form only has 4 columns",
        f"TOTAL FEDERAL Year 1 {money(TOTAL)}",
        "Non-federal / match: $0 (none required, none proposed)",
        "",
        "=== SF-LLL ===",
        f"Reporting Entity: {ORG}",
        f"Street/City/State/Zip: {STREET}, {CITY}, {COUNTRY}",
        "Program: CDC Global Health Security — Uganda local partnerships",
        f"Federal Agency: CDC / Opportunity {OPPORTUNITY}",
        "Lobbying registrant: NONE — no lobbying activities; check 'No' / not applicable as form allows",
        f"Signature: {PD_NAME}, {PD_TITLE}, {TODAY}",
        "",
        "=== Project Abstract ===",
        "See WS01739425-Project_AbstractSummary_2_0-V2.0.pdf",
        "",
        "=== Attachments ===",
        "Project Narrative -> attach WS01739425-ProjectNarrativeAttachments_1_2-V1.2.pdf",
        "Budget Narrative -> attach WS01739425-BudgetNarrativeAttachments_1_2-V1.2.pdf",
        "Local partner letter -> attach Local_Partner_Preference_Letter.pdf + company-docs/",
    ]
    (OUT / "FIELD_MAP.txt").write_text("\n".join(lines), encoding="utf-8")


def build_sf424():
    st = styles()
    # Same content for both workspace SF-424 shells in the package download
    paths = [
        OUT / "WS01739425-SF424_4_0-V4.0.pdf",
        OUT / "WS01739423-SF424_4_0-V4.0.pdf",
    ]
    path = paths[0]
    template, on_page = doc_template(path, "SF-424")
    story = [
        Paragraph("SF-424 Application for Federal Assistance — FILLED VALUES", st["Cover"]),
        Paragraph(f"{ORG} | {OPPORTUNITY} | Generated {date.today().isoformat()}", st["Center"]),
        Paragraph(
            "Copy these values into the Grants.gov SF-424 (V4.0) Adobe/web form. "
            f"UEI field remains CONFIRM until you paste the live SAM.gov UEI. {SLOGAN}",
            st["Meta"],
        ),
        Paragraph("1-7. Application identifiers", st["H1"]),
        tbl(
            st,
            ["Box", "Value"],
            [
                ["1. Type of Submission", "Application"],
                ["2. Type of Application", "New"],
                ["3. Date Received", "(Grants.gov stamps on submit)"],
                ["4. Applicant Identifier", COMPANY_NO],
                ["5a. Federal Entity Identifier", OPPORTUNITY],
                ["5b. Federal Award Identifier", "N/A (new application)"],
                ["6. Date Received by State", "N/A"],
                ["7. State Application Identifier", "N/A"],
            ],
            [2.2 * inch, 5.0 * inch],
        ),
        Paragraph("8. Applicant information", st["H1"]),
        tbl(
            st,
            ["Box", "Value"],
            [
                ["8a. Legal Name", ORG],
                ["8b. EIN / TIN", EIN],
                ["8c. UEI", UEI],
                ["8d. Address", f"{STREET}"],
                ["City / County", f"{CITY} / {COUNTY}"],
                ["State / Province / Country / ZIP", f"{STATE} / {PROVINCE} / {COUNTRY} / {ZIP}"],
                ["8e. Organizational Unit", "FairBanks Community Reach — Global Health Security Programme"],
                ["8f. Name", PD_NAME],
                ["Title", PD_TITLE],
                ["Organizational Affiliation", ORG_SHORT],
                ["Telephone / Email", f"{PHONE} / {EMAIL}"],
            ],
            [2.2 * inch, 5.0 * inch],
        ),
        Paragraph("9-16. Agency, CFDA, opportunity, geography", st["H1"]),
        tbl(
            st,
            ["Box", "Value"],
            [
                ["9. Type of Applicant", "For-Profit Organization (Other than Small Business); Foreign / non-U.S. entity"],
                ["10. Federal Agency", "Centers for Disease Control and Prevention (CDC)"],
                ["11. CFDA Number / Title", f"{CFDA} — {CFDA_TITLE}"],
                ["12. Funding Opportunity Number", OPPORTUNITY],
                ["12. Title", "Strengthening global health security through local partnerships in Uganda"],
                ["13. Competition Identification", "N/A / as shown on package"],
                ["14. Areas Affected by Project", "Uganda — Kampala peri-urban catchments (Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya and nearby); expand under MoH/CDC guidance"],
                ["15. Descriptive Title of Project", TITLE],
                ["16. Congressional Districts", "Applicant: N/A (foreign). Program: N/A (performed in Uganda)."],
            ],
            [2.2 * inch, 5.0 * inch],
        ),
        PageBreak(),
        Paragraph("17-21. Period, funding, certifications", st["H1"]),
        tbl(
            st,
            ["Box", "Value"],
            [
                ["17a. Proposed Start Date", START],
                ["17b. Proposed End Date", END],
                ["18a. Federal", money(TOTAL)],
                ["18b. Applicant", "$0"],
                ["18c. State", "$0"],
                ["18d. Local", "$0"],
                ["18e. Other", "$0"],
                ["18f. Program Income", "$0"],
                ["18g. TOTAL", money(TOTAL)],
                ["19. EO 12372 Review", "c. Program is not covered by E.O. 12372"],
                ["20. Federal Debt Delinquency", "No"],
                ["21. Authorized Representative", f"{PD_NAME}"],
                ["Title", PD_TITLE],
                ["Telephone / Email", f"{PHONE} / {EMAIL}"],
                ["Signature / Date", f"Sign in Grants.gov / Adobe Reader — Date {TODAY}"],
            ],
            [2.2 * inch, 5.0 * inch],
        ),
        Spacer(1, 8),
        Paragraph(
            "Note: Year 1 federal total above equals Components 1-5 draft asks "
            f"({money(C1)} + {money(C2)} + {money(C3)} + {money(C4)} + {money(C5)}). "
            "Only Component 1 is expected to be funded initially; Components 2-5 may be "
            "approved but unfunded until CDC activates emergency funding. CONFIRM final figures.",
            st["Small"],
        ),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("SF-424:", path)
    # Duplicate for the second SF-424 shell name in the package
    import shutil

    for extra in paths[1:]:
        shutil.copy2(path, extra)
        print("SF-424 copy:", extra)


def build_sf424a():
    st = styles()
    path = OUT / "WS01739425-SF424A-V1.0.pdf"
    template, on_page = doc_template(path, "SF-424A")
    pw = letter[0] - 1.5 * inch

    # Approximate category splits for C2-C5 (surge-weighted)
    def split(total, weights):
        # weights sum ~1
        raw = {k: int(total * w) for k, w in weights.items()}
        # fix rounding on Other
        gap = total - sum(raw.values())
        raw["Other"] = raw.get("Other", 0) + gap
        return raw

    w_core = {
        "Personnel": 720_000 / C1,
        "Fringe Benefits": 144_000 / C1,
        "Travel": 85_000 / C1,
        "Equipment": 60_000 / C1,
        "Supplies": 95_000 / C1,
        "Contractual": 780_000 / C1,
        "Other": 280_000 / C1,
        "Indirect Charges": 286_000 / C1,
    }
    w_surge = {
        "Personnel": 0.28,
        "Fringe Benefits": 0.05,
        "Travel": 0.08,
        "Equipment": 0.03,
        "Supplies": 0.12,
        "Contractual": 0.28,
        "Other": 0.10,
        "Indirect Charges": 0.06,
    }
    comps = {
        "1 Core GHS": C1_CAT,
        "2 Small-scale": split(C2, w_surge),
        "3 Large-scale": split(C3, w_surge),
        "4 Emerging": split(C4, w_surge),
        "5 Humanitarian": split(C5, w_surge),
    }
    cats = [
        "Personnel",
        "Fringe Benefits",
        "Travel",
        "Equipment",
        "Supplies",
        "Contractual",
        "Construction",
        "Other",
        "Indirect Charges",
    ]

    story = [
        Paragraph("SF-424A Budget Information — Non-Construction Programs", st["Cover"]),
        Paragraph(f"{ORG} | Year 1 draft (CONFIRM) | {date.today().isoformat()}", st["Center"]),
        Paragraph(
            "Enter these amounts on SF-424A. Forms often show only 4 grant-program columns — "
            "put Components 1-4 on the first SF-424A and Component 5 on a second SF-424A if required. "
            "All figures USD. Cost share: none.",
            st["Meta"],
        ),
        Paragraph("Section A — Budget Summary (Federal)", st["H1"]),
        tbl(
            st,
            ["Grant Program Function", "CFDA", "Federal"],
            [
                ["Component 1 — Core GHS priorities", CFDA, money(C1)],
                ["Component 2 — Small-scale outbreak / PHE response", CFDA, money(C2)],
                ["Component 3 — Large-scale outbreak / PHE response", CFDA, money(C3)],
                ["Component 4 — Emerging infectious disease threats", CFDA, money(C4)],
                ["Component 5 — Humanitarian emergency", CFDA, money(C5)],
                ["TOTAL", "", money(TOTAL)],
            ],
            [pw * 0.55, pw * 0.15, pw * 0.30],
        ),
        Paragraph("Non-Federal funds: $0 for all components. Total (Federal + Non-Federal) = Federal.", st["Small"]),
        Paragraph("Section B — Budget Categories (Year 1 Federal by component)", st["H1"]),
    ]

    header = ["Object Class"] + list(comps.keys()) + ["Total"]
    rows = []
    totals = {k: 0 for k in comps}
    grand = 0
    for cat in cats:
        row = [cat]
        line = 0
        for name, bud in comps.items():
            v = bud.get(cat, 0)
            row.append(money(v) if v else "$0")
            totals[name] += v
            line += v
        row.append(money(line))
        grand += line
        rows.append(row)
    rows.append(["TOTAL"] + [money(totals[k]) for k in comps] + [money(grand)])
    story.append(tbl(st, header, rows, [1.15 * inch] + [0.95 * inch] * 5 + [0.9 * inch]))
    story.append(
        Paragraph(
            "Ceilings (do not exceed): C1 $5M; C2 $10M; C3 $15M; C4 $15M; C5 $20M. "
            "Indirect: foreign organisation 8% MTDC estimated (CONFIRM). "
            "Component 1 expected initial funding; Components 2-5 contingency.",
            st["Small"],
        )
    )
    story.append(Paragraph("Section C — Non-Federal Resources", st["H1"]))
    story.append(Paragraph("All zeros — no cost sharing or matching proposed.", st["Body"]))
    story.append(Paragraph("Section D — Forecasted Cash Needs", st["H1"]))
    story.append(
        Paragraph(
            f"Year 1 total federal need {money(TOTAL)}. If only Component 1 is funded initially, "
            f"cash need is {money(C1)}. Quarterly distribution for Component 1 (illustrative): "
            f"Q1 {money(C1 // 4)}; Q2 {money(C1 // 4)}; Q3 {money(C1 // 4)}; Q4 {money(C1 - 3 * (C1 // 4))}.",
            st["Body"],
        )
    )
    story.append(Paragraph("Section E — Budget Estimates of Federal Funds Needed for Balance of the Project", st["H1"]))
    story.append(
        Paragraph(
            "Years 2-5 amounts to be set at continuation based on progress and available funds. "
            "Illustrative planning level for Core GHS (Component 1) similar to Year 1 order of magnitude "
            f"({money(C1)}/year) subject to CDC continuation decisions. Contingency components remain "
            "event-driven.",
            st["Body"],
        )
    )
    story.append(Paragraph("Section F — Other Budget Information", st["H1"]))
    story.append(
        Paragraph(
            "Indirect charges: estimated 8% of MTDC for foreign organisation (exclusive of equipment "
            "and subawards over $25,000), unless a negotiated rate applies — CONFIRM before submit. "
            "Remarks: Award funds systems strengthening, surveillance, training, emergency readiness, "
            "and community-facility linkages. Routine clinical care is not charged to this award. "
            "No research proposed.",
            st["Body"],
        )
    )
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("SF-424A:", path)


def build_sflll():
    st = styles()
    path = OUT / "WS01739425-SFLLL_2_0-V2.0.pdf"
    template, on_page = doc_template(path, "SF-LLL")
    story = [
        Paragraph("SF-LLL Disclosure of Lobbying Activities — FILLED VALUES", st["Cover"]),
        Paragraph(f"{ORG} | {OPPORTUNITY}", st["Center"]),
        Paragraph(
            "FairBanks is not engaging a lobbyist for this application. Enter as No lobbying / not applicable "
            "per form instructions. If the form still requires entity fields, use the values below.",
            st["Meta"],
        ),
        tbl(
            st,
            ["Field", "Value"],
            [
                ["1. Type of Federal Action", "b. Grant / Cooperative agreement"],
                ["2. Status of Federal Action", "a. Bid/offer/application"],
                ["3. Report Type", "a. Initial filing"],
                ["4. Name and Address of Reporting Entity", f"{ORG}; {STREET}; {CITY}; {COUNTRY}"],
                ["Congressional District", "N/A (foreign entity)"],
                ["5. If Reporting Entity in No. 4 is Subawardee", "N/A — prime applicant"],
                ["6. Federal Department/Agency", "Department of Health and Human Services — CDC"],
                ["7. Federal Program Name / CFDA", f"Global Health Security / {CFDA}"],
                ["8. Federal Action Number", OPPORTUNITY],
                ["9. Award Amount", f"Requested Year 1 federal {money(TOTAL)} (draft)"],
                ["10a. Name and Address of Lobbying Registrant", "NONE — no lobbying registrant engaged"],
                ["10b. Individuals Performing Services", "NONE"],
                ["11. Information requested through this form is authorized...", "Yes — authorised official certifies"],
                ["Signature", f"{PD_NAME}, {PD_TITLE}"],
                ["Telephone / Email / Date", f"{PHONE} / {EMAIL} / {TODAY}"],
            ],
            [2.4 * inch, 4.8 * inch],
        ),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("SF-LLL:", path)


def build_abstract():
    st = styles()
    path = OUT / "WS01739425-Project_AbstractSummary_2_0-V2.0.pdf"
    template, on_page = doc_template(path, "Project Abstract")
    story = [
        Paragraph("Project Abstract Summary", st["Cover"]),
        Paragraph(f"{ORG} | {OPPORTUNITY} | {SLOGAN}", st["Center"]),
        Paragraph(
            "Guided by The FairBanks Blueprint: Building the Future of Family & Community Health. "
            "Paste this text into the Grants.gov Project Abstract Summary form (trim if character limit requires).",
            st["Meta"],
        ),
        Paragraph(
            "Project Title: FairBanks Blueprint in action: strengthening Uganda last-mile Global Health "
            "Security through Community Reach, CHW/VHT surveillance, and FCHIP intelligence feeding "
            "MoH/NISS pathways",
            st["Body"],
        ),
        Paragraph(f"Applicant: {ORG}", st["Body"]),
        Paragraph(f"Project Director: {PD_NAME}, {PD_TITLE}", st["Body"]),
        Paragraph(f"Period: {START} to {END} | Funding instrument: Cooperative agreement | CFDA: {CFDA}", st["Body"]),
        Paragraph(f"Year 1 federal request (draft, all components): {money(TOTAL)} | Component 1 core: {money(C1)}", st["Body"]),
        Paragraph("Abstract", st["H1"]),
        Paragraph(ABSTRACT, st["Body"]),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("Abstract:", path)


def build_narrative():
    st = styles()
    path = OUT / "WS01739425-ProjectNarrativeAttachments_1_2-V1.2.pdf"
    template, on_page = doc_template(path, "Project Narrative")
    story = [
        Paragraph("Project Narrative", st["Cover"]),
        Paragraph(
            f"CDC-RFA-JG-26-0054 — Strengthening global health security through local partnerships in Uganda",
            st["Center"],
        ),
        Paragraph(f"{ORG} ({ORG_SHORT}) | {SLOGAN}", st["Center"]),
        Paragraph(
            f"Attach this PDF to the Grants.gov Project Narrative Attachments form. "
            f"Format: English, Times 11-12 pt, ~1-inch margins. Generated {date.today().isoformat()}. "
            f"UEI: {UEI}",
            st["Meta"],
        ),
        Paragraph("Table of contents", st["H1"]),
        bullets(
            st,
            [
                "1. Background and approach",
                "2. Evaluation and performance measurement plan (EPMP)",
                "3. Organizational capacity and collaborations",
                "4. Work plan snapshots (Components 1-5)",
            ],
        ),
        Paragraph("1. Background and approach", st["H1"]),
        Paragraph("1.1 Shared win", st["H2"]),
        Paragraph(
            "CDC's America First Global Health Strategy asks partners to help make America safer, "
            "stronger, and more prosperous by containing threats at source, building resilient systems, "
            "and using data for real response. Uganda's NAPHS II and the 2023 JEE ask for the same "
            "practical things on the ground: better surveillance links, ready workers, and emergency "
            "systems that reach districts and communities.",
            st["Body"],
        ),
        Paragraph(
            "FairBanks sits where those two needs meet. We already walk the streets and homes of "
            "peri-urban Kampala. We already see patients at FairBanks Medical Centre. We already "
            "collect community data through CHWs/VHTs and a working FCHIP MVP. What CDC and MoH "
            "still need is a local partner that can turn that daily work into timely, government-usable "
            "signals — without inventing a second national system.",
            st["Body"],
        ),
        Paragraph(
            "If funded, CDC gets a Uganda partner with skin in the game. MoH gets last-mile feed into "
            "NISS-aligned pathways and practised 7-1-7 loops. Communities get earlier alerts and clearer "
            "referrals. FairBanks gets to prove its tools serve the public system.",
            st["Body"],
        ),
        Paragraph("1.2 Problem", st["H2"]),
        Paragraph(
            "Uganda sits in a region where endemic, new, and returning infectious diseases keep causing "
            "outbreaks. Fast urban growth, crowded peri-urban settlements, and high border and travel "
            "movement raise the chance that a local cluster becomes a national — and cross-border — "
            "threat. Americans and Ugandans are safer when outbreaks are found and contained early, "
            "close to where people live.",
            st["Body"],
        ),
        Paragraph(
            "In Kampala peri-urban communities where FairBanks already works — Bukoto, Kyebando, "
            "Kisaasi, Kamwokya, Kikaaya and nearby areas — families often reach care late. CHWs and "
            "VHTs visit homes and schools, but much of what they see still sits in paper books or "
            "separate tools. Clinic records, community reports, lab results, and weather or place "
            "signals rarely join one picture.",
            st["Body"],
        ),
        Paragraph(
            "Uganda's 2023 Joint External Evaluation showed real progress on emergency response, labs, "
            "surveillance, and workforce — and still named clear gaps: weak information sharing across "
            "sectors; limited internet access for surveillance data; incomplete surge plans and staffing "
            "at subnational level; uneven use of data to forecast risk; and the need for stronger links "
            "from communities and facilities into national emergency and surveillance structures, "
            "including work toward a functional National Integrated Surveillance System (NISS).",
            st["Body"],
        ),
        Paragraph(
            "FairBanks proposes to close the last-mile gap under MoH leadership: help community and "
            "clinic signals reach MoH systems faster, train the people who collect those signals, and "
            "practise response with district structures so 7-1-7 style timing becomes real on the ground.",
            st["Body"],
        ),
        Paragraph("1.3 Goal, theory of action, and outcomes", st["H2"]),
        Paragraph(
            "Goal: Strengthen Uganda's prevention, detection, and response capacities for priority "
            "human and zoonotic threats by linking FairBanks Community Reach, FairBanks Medical Centre, "
            "and FCHIP tools into Ministry of Health-led surveillance and emergency pathways — so "
            "outbreaks are found earlier and contained closer to source.",
            st["Body"],
        ),
        Paragraph(
            "Theory of action: Train and equip frontline workers -> capture structured community and "
            "facility signals -> share into NISS-aligned / MoH channels -> trigger district and regional "
            "response -> review timing against 7-1-7 -> improve and hand tools to government use.",
            st["Body"],
        ),
        Paragraph("Asterisked outcomes we will substantially contribute to:", st["Body"]),
        bullets(
            st,
            [
                "Strategy 1: Improved health systems resilience for emergencies linked to districts and communities.",
                "Strategy 2: Increased capacity, coordination, and response readiness of the public health workforce.",
                "Strategy 4: Resilient integrated surveillance infrastructure capable of real-time detection and response; shorter detection and response time guided by 7-1-7 style benchmarks.",
                "Strategy 6: Better coordination between public health systems, clinical care, and communities; culturally responsive frontline workforce supporting public health operations.",
            ],
        ),
        Paragraph(
            "For Strategies 3 (national laboratory systems) and 5 (border health), FairBanks will support "
            "MoH and specialised partners with referral pathways, specimen logistics coordination, "
            "private-facility incident reporting support where asked, and community risk communication "
            "near travel corridors — without claiming sole national lead.",
            st["Body"],
        ),
        Paragraph(
            "Geographic start: Kampala Capital City Authority catchments linked to FairBanks Community "
            "Reach, with Year 1-2 expansion to selected neighbouring districts agreed with MoH/CDC.",
            st["Body"],
        ),
        PageBreak(),
        Paragraph("1.4 Year 1 Component 1 activities by strategy", st["H2"]),
        Paragraph(
            "We build on IDSR/eIDSR, community-based surveillance, CHW/VHT structures, PHEOC incident "
            "management, 7-1-7 timing, One Health awareness, and RCCE. Tools will work offline first and "
            "respect Uganda data-protection and consent norms.",
            st["Body"],
        ),
        Paragraph("Strategy 1 — Health emergency management", st["H2"]),
        bullets(
            st,
            [
                "Map FairBanks catchments to district/regional PHEOC and agree notification SOPs.",
                "Run table-top and field drills with CHWs, facility staff, and district teams using 7-1-7 timing.",
                "Practise incident management for small events from community alert to district activation.",
                "Support district-level response coordination for events in FairBanks catchments.",
                "Document after-action reviews and corrective actions.",
            ],
        ),
        Paragraph("Strategy 2 — Human resources / workforce development", st["H2"]),
        bullets(
            st,
            [
                "Train and coach CHWs/VHTs and selected facility staff on community-based surveillance, event notification, and safe specimen/referral pathways.",
                "Support One Health-aware briefings with local animal/environmental focal points where MoH/districts arrange joint sessions.",
                "Build a surge roster of FairBanks-linked workers for community investigation and RCCE.",
                "Align training with MoH competency expectations; share materials with MoH and CDC.",
            ],
        ),
        Paragraph("Strategy 3 — National laboratory systems (support role)", st["H2"]),
        bullets(
            st,
            [
                "Strengthen sample referral and transport links from FairBanks Medical Centre and outreach points into MoH-approved laboratory pathways.",
                "Help MoH track private-facility coverage and routine incident reporting for priority outbreak diseases in our catchments.",
                "Support biosafety basics at facility sample handling points (SOPs, PPE, documentation).",
            ],
        ),
        Paragraph("Strategy 4 — Surveillance", st["H2"]),
        bullets(
            st,
            [
                "Deploy and validate FCHIP mobile capture for CHW/VHT and facility sentinel signals.",
                "Configure exports and dashboards that contribute to NISS/MoH pathways under MoH rules (not a parallel silo).",
                "Train staff on data quality, analysis, and simple visualisation for catchment and district review.",
                "Use GIS and approved climate/weather feeds where they help early warning, with clear limits on model claims.",
            ],
        ),
        Paragraph("Strategy 5 — Border health security (support role)", st["H2"]),
        bullets(
            st,
            [
                "Coordinate with MoH/border health partners on RCCE and community surveillance near high-mobility corridors serving Kampala catchments.",
                "Share community mobility and event signals that support POPCAB-style risk pictures when districts request them.",
                "Do not claim to run national Points of Entry programmes.",
            ],
        ),
        Paragraph("Strategy 6 — Public health programmes and service delivery links", st["H2"]),
        bullets(
            st,
            [
                "Strengthen community-facility linkages for priority diseases (GHS, HIV, TB, malaria, cholera, Ebola, Marburg, mpox, immunisation) during routine and outbreak periods.",
                "Support public health campaigns with community mobilisation and data feedback — not award-funded routine clinical care.",
                "Write SOPs for data use in service and public health decisions; train teams to use them.",
                "Provide MoH, CDC, and partners with copies of or access to tools, training materials, and systems developed under the award.",
            ],
        ),
        Paragraph("1.5 Complementing the national programme and timelines", st["H2"]),
        Paragraph(
            "This proposal continues the direction of prior CDC GHS investment in Uganda (including "
            "work under CDC-RFA-GH20-2124 as described in the NOFO). FairBanks adds the missing "
            "community last mile through the Community Reach cascade and FCHIP on the Data & Feedback "
            "loop, configured to feed government systems.",
            st["Body"],
        ),
        tbl(
            st,
            ["Quarter", "Year 1 milestones"],
            [
                ["Q1", "Staffing; MoH/district kick-off; SOP drafts; CHW roster; FCHIP form pack; EPMP/DMP drafts started"],
                ["Q2", "Train first cohorts; routine community event reporting; first joint drill; sample referral pathway live"],
                ["Q3", "Expand sites; mid-year data quality review; second drill; MoH data-sharing test exports"],
                ["Q4", "Annual performance package; Year 2 work plan; tool handoff package; surge roster exercise"],
            ],
            [1.0 * inch, 6.2 * inch],
        ),
        PageBreak(),
        Paragraph("2. Evaluation and performance measurement plan (EPMP)", st["H1"]),
        Paragraph(
            "About 5-10% of project funds will support monitoring, reporting, and evaluation. Final "
            "indicators will be agreed with CDC after award from the DGHP partner-level list.",
            st["Body"],
        ),
        Paragraph("2.1 Priority indicator areas", st["H2"]),
        bullets(
            st,
            [
                "Surveillance / community mitigation: staff trained in investigation/contact support; tailored risk mitigation strategies; RCCE knowledge gains.",
                "Emergency operations: trainings on RRT/RCCE/PHEM skills; SOPs for notification/IMS/RCCE; timely notification toward district activation (7-1-7 style).",
                "Laboratory (support): sample referral completeness and turnaround from FairBanks-linked sites.",
                "IPC: IPC focal person and guideline-based improvements at FairBanks Medical Centre and linked sites.",
            ],
        ),
        Paragraph("2.2 Project-specific measures (Year 1 plan)", st["H2"]),
        tbl(
            st,
            ["Measure", "Baseline plan", "Year 1 target plan"],
            [
                ["Median days community signal to district notification", "Establish Q1-Q2", "Move toward 7-1-7; publish quarterly trend"],
                ["CHWs/VHTs with complete weekly reports", "Roster at kick-off", "At least 80% of active roster by Q4"],
                ["Priority referrals with documented outcome", "Baseline audit Q1", "At least 70% documented by Q4"],
                ["Successful MoH/NISS-aligned data package tests", "0", "At least 2 by Q4"],
                ["After-action reviews after drills/events", "0", "At least 2 by Q4"],
            ],
            [2.6 * inch, 2.2 * inch, 2.4 * inch],
        ),
        Paragraph("2.3 Methods, reporting, DMP summary", st["H2"]),
        bullets(
            st,
            [
                "Collect: mobile forms (offline), facility registers, training pre/post tests, drill logs, sample referral logs, dashboard extracts.",
                "Frequency: weekly ops checks; monthly internal review; quarterly packs to CDC; Annual Performance Report.",
                "Quality: form validation; supervisor spot checks; quarterly data-quality audits.",
                "Use: monthly improvement meetings; share with district/MoH; adjust SOPs and training.",
                "Evaluation: Year 2 process evaluation; optional Year 4 outcome evaluation if funds allow.",
                "DMP: role-based access; Uganda Data Protection and Privacy Act; minimise PII; de-identify analytics; detailed DMP within six months of award.",
            ],
        ),
        Paragraph("3. Organizational capacity and collaborations", st["H1"]),
        Paragraph(
            f"{ORG} (Company No. {COMPANY_NO}; TIN {TIN}; NSSF {NSSF}) is a Uganda company with "
            f"principal operations in Kampala ({STREET}, {COUNTY}). We run: (1) FairBanks Medical "
            "Centre — outpatient care, diagnostics, pharmacy, referrals, and related services; "
            "(2) FairBanks Community Reach — CHWs/VHTs, outreach, MCH support, GeriCare, school "
            "health, CHIS, and livelihood pathways; (3) FCHIP — working MVP for mobile capture, "
            "sync, dashboards, GIS, and secure data APIs.",
            st["Body"],
        ),
        Paragraph(
            f"Project Director: {PD_NAME}, {PD_TITLE} — 15+ years in Uganda private healthcare "
            "leadership and HR; MA Social Sector Planning and Management (Makerere); PhD in "
            f"Management in progress; Uganda Healthcare Federation links. Contact: {EMAIL}, {PHONE}. "
            "Website: " + WEBSITE + ".",
            st["Body"],
        ),
        Paragraph(
            "Honest limit: we have not previously managed a multi-million-dollar U.S. federal "
            "cooperative agreement. We will strengthen segregated ledgers, procurement, timesheets, "
            "and audits for this award. We ask CDC to judge us on local presence, technical fit for "
            "community surveillance, documented company standing, and willingness to work under MoH leadership.",
            st["Body"],
        ),
        Paragraph("Collaborations (if funded):", st["H2"]),
        bullets(
            st,
            [
                "Ministry of Health (surveillance, community health, emergency operations, digital health).",
                "District / KCCA health teams covering our catchments.",
                "Other CDC-funded GHS, HIV/TB, immunisation, and emerging infection partners.",
                "Community leaders, CHWs/VHTs, schools, and local CBOs in Community Reach.",
            ],
        ),
        Paragraph("Key personnel", st["H2"]),
        tbl(
            st,
            ["Role", "Person", "Notes"],
            [
                ["Project Director / Authorized Official", PD_NAME, "Overall accountability"],
                ["GHS Programme Manager", "To hire / CONFIRM", "Day-to-day delivery"],
                ["M&E Lead", "To hire / CONFIRM", "EPMP, indicators, APR"],
                ["Data / FCHIP Lead", "To hire / CONFIRM", "Mobile tools, NISS-aligned exports"],
                ["CHW Supervisor", "Existing + surge", "Field quality and coaching"],
                ["Clinical / IPC Focal", "Existing Medical Centre role", "Sample SOPs/IPC — award-allowed time only"],
            ],
            [2.2 * inch, 2.0 * inch, 3.0 * inch],
        ),
        PageBreak(),
        Paragraph("4. Work plan snapshots", st["H1"]),
        Paragraph("4.1 Component 1 — Core GHS (Year 1)", st["H2"]),
        tbl(
            st,
            ["Strategy", "Activity", "Measure examples", "By"],
            [
                ["1", "District notification SOPs + 2 drills", "Drill reports; time metrics", "Q2, Q4"],
                ["2", "Train CHW/VHT cohorts (target CONFIRM e.g. 80)", "# trained; pre/post scores", "Q2-Q3"],
                ["3", "Sample referral pathway SOPs live", "% complete referrals", "Q2"],
                ["4", "FCHIP surveillance live + MoH export tests", "Weekly reports; export tests", "Q2-Q4"],
                ["5", "Corridor RCCE + signal sharing", "# sessions; signals shared", "Q3-Q4"],
                ["6", "Priority-disease community-facility linkage", "Completed referral outcomes", "Ongoing"],
                ["M&E", "EPMP/DMP detailed; quarterly reviews", "Reports submitted", "Month 6+"],
            ],
            [0.8 * inch, 2.6 * inch, 2.4 * inch, 0.9 * inch],
        ),
        Paragraph("4.2 Components 2-5 — Contingency (approved but unfunded until CDC activates)", st["H2"]),
        bullets(
            st,
            [
                f"Component 2 ({money(C2)} draft): surge roster; community investigation/contact support; RCCE; temporary dashboards; district IMS support.",
                f"Component 3 ({money(C3)} draft): expand surge staffing; multi-district mobilisation under MoH; sample referral volume logistics; extended RCCE; recovery support.",
                f"Component 4 ({money(C4)} draft): rapid form updates for new pathogens; sentinel intensification; partner lab referral; special training modules.",
                f"Component 5 ({money(C5)} draft): community health surveillance and RCCE in crisis-affected groups within MoH tasking; maintain essential programme links.",
            ],
        ),
        Paragraph("Closing statement", st["H1"]),
        Paragraph(
            "FairBanks is ready to serve as a documented local partner for CDC-RFA-JG-26-0054: "
            "closing Uganda's last-mile Global Health Security gap under MoH leadership so threats "
            "are detected and contained closer to source — safer for Uganda and for the United States.",
            st["Body"],
        ),
        Paragraph(f"{PD_NAME} | {PD_TITLE} | {EMAIL} | {PHONE} | {SLOGAN}", st["Center"]),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("Narrative:", path)


def build_budget_narrative():
    st = styles()
    path = OUT / "WS01739425-BudgetNarrativeAttachments_1_2-V1.2.pdf"
    template, on_page = doc_template(path, "Budget Narrative")
    story = [
        Paragraph("Budget Narrative (Year 1) — All Components", st["Cover"]),
        Paragraph(f"{ORG} | {OPPORTUNITY} | All figures U.S. dollars | Generated {date.today().isoformat()}", st["Center"]),
        Paragraph(
            "Attach this PDF to the Grants.gov Budget Narrative Attachments form. "
            "CONFIRM final figures before submit. Cost share: none. No research. "
            "Routine clinical care not charged to this award.",
            st["Meta"],
        ),
        Paragraph("Summary of Year 1 federal requests", st["H1"]),
        tbl(
            st,
            ["Component", "Draft ask", "Ceiling", "Status"],
            [
                ["1 Core GHS", money(C1), "$5,000,000", "Expected initial funding"],
                ["2 Small-scale response", money(C2), "$10,000,000", "Contingency"],
                ["3 Large-scale response", money(C3), "$15,000,000", "Contingency"],
                ["4 Emerging threats", money(C4), "$15,000,000", "Contingency"],
                ["5 Humanitarian", money(C5), "$20,000,000", "Contingency"],
                ["TOTAL", money(TOTAL), "", ""],
            ],
            [2.4 * inch, 1.4 * inch, 1.4 * inch, 2.0 * inch],
        ),
        Paragraph(
            "Indirect costs: foreign organisation rate 8% of MTDC (exclusive of equipment and "
            "subawards over $25,000), unless a negotiated rate applies — CONFIRM method before submit.",
            st["Body"],
        ),
        Paragraph("Component 1 — Core GHS detailed justification", st["H1"]),
        tbl(
            st,
            ["Category", "Amount", "Justification"],
            [
                ["Salaries and wages", money(720_000), "Project Director portion; Programme Manager; M&E; Data/FCHIP; CHW supervisors; finance/admin charged to award"],
                ["Fringe benefits", money(144_000), "Statutory and organisational benefits on award-charged staff"],
                ["Travel", money(85_000), "In-country field supervision, district meetings, limited regional workshops"],
                ["Equipment", money(60_000), "Field devices and approved IT — each item justified; furniture generally not allowed"],
                ["Supplies", money(95_000), "Training materials, PPE for drills/sample handling, connectivity, form printing"],
                ["Contractual", money(780_000), "MoH-aligned technical partners for lab/border support modules; software hosting; external audit"],
                ["Other", money(280_000), "Short-term trainers, translation, community meeting costs, evaluation support"],
                ["Direct subtotal", money(2_164_000), ""],
                ["Indirect (8% MTDC est.)", money(286_000), "CONFIRM MTDC base"],
                ["Component 1 TOTAL", money(C1), "Within $5,000,000 ceiling"],
            ],
            [1.6 * inch, 1.2 * inch, 4.4 * inch],
        ),
        Paragraph("Components 2-5 — Contingency narratives", st["H1"]),
        Paragraph(
            f"Component 2 ({money(C2)}): Surge staffing and fringe; rapid travel; emergency supplies/PPE; "
            "contractual surge support for investigation/RCCE/data dashboards; other community meeting "
            "costs; indirect. Activate only if CDC funds small-scale outbreak/PHE response.",
            st["Body"],
        ),
        Paragraph(
            f"Component 3 ({money(C3)}): Larger surge staffing; multi-district travel/logistics; expanded "
            "supplies; contractual partners for mobilisation and sample referral volume; recovery support; "
            "indirect. For large-scale outbreak/PHE response when activated.",
            st["Body"],
        ),
        Paragraph(
            f"Component 4 ({money(C4)}): Rapid form/tool updates; sentinel intensification staffing; "
            "special training; partner lab referral contractual support; travel/supplies; indirect. "
            "For emerging infectious disease threats when activated.",
            st["Body"],
        ),
        Paragraph(
            f"Component 5 ({money(C5)}): Community surveillance and RCCE in humanitarian settings under "
            "MoH tasking; surge staffing/travel/supplies; contractual support; indirect. For humanitarian "
            "emergency response when activated.",
            st["Body"],
        ),
        Paragraph("Unallowable costs (we will not charge)", st["H1"]),
        bullets(
            st,
            [
                "Research as defined by this NOFO",
                "Lobbying",
                "Routine clinical care not allowed by law / not the purpose of the award",
                "Pre-award costs without written approval",
                "Budgets in any currency other than U.S. dollars",
            ],
        ),
        Paragraph(
            f"Authorised official for budget: {PD_NAME}, {PD_TITLE}, {EMAIL}, {PHONE}.",
            st["Body"],
        ),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("Budget narrative:", path)


def build_local_partner_letter():
    st = styles()
    path = OUT / "Local_Partner_Preference_Letter.pdf"
    template, on_page = doc_template(path, "Local Partner Letter")
    story = [
        Paragraph(ORG, st["Cover"]),
        Paragraph(f"{STREET}", st["Center"]),
        Paragraph(f"{CITY}, {COUNTRY}", st["Center"]),
        Paragraph(f"{EMAIL} | {PHONE} | {WEBSITE}", st["Center"]),
        Spacer(1, 12),
        Paragraph(TODAY, st["Body"]),
        Paragraph("Broderick Yoerg / DGHP NOFO Review", st["Body"]),
        Paragraph("Centers for Disease Control and Prevention", st["Body"]),
        Paragraph("Email: DGHPNOFOs@cdc.gov", st["Body"]),
        Spacer(1, 8),
        Paragraph(f"Re: {OPPORTUNITY} — Local partner funding preference", st["Body"]),
        Paragraph(
            f"I am the authorised official of {ORG}. We ask to be considered for the local partner "
            "funding preference.",
            st["Body"],
        ),
        Paragraph("We meet the entity definition as follows:", st["Body"]),
        bullets(
            st,
            [
                f"We are incorporated under the laws of Uganda (Company No. {COMPANY_NO}) with our principal place of business in Kampala, Uganda ({STREET}, {COUNTY}).",
                "CONFIRM before signing: (a) at least 75% beneficial ownership by Uganda citizens or permanent residents; and/or (b) at least 75% of senior, mid-level, or support staff are Uganda citizens or permanent residents. Supporting lists are attached.",
                "CONFIRM board: at least 51% of board members are Uganda citizens or permanent residents — OR state if the company has no board of directors and explain governance.",
            ],
        ),
        Paragraph(
            "Attached (as available): Certificate of Incorporation / company registration evidence; "
            f"TIN certificate ({TIN}); NSSF employer certificate ({NSSF}); officers/owners list with "
            "citizenship; staff residency summary; board list with citizenship (or statement of no board); "
            "facility licence extracts.",
            st["Body"],
        ),
        Spacer(1, 16),
        Paragraph("Sincerely,", st["Body"]),
        Spacer(1, 28),
        Paragraph(f"{PD_NAME}", st["Body"]),
        Paragraph(PD_TITLE, st["Body"]),
        Paragraph(ORG, st["Body"]),
        Paragraph(f"{EMAIL} · {PHONE}", st["Body"]),
        Paragraph(
            "NOTE: Do not submit until CONFIRM citizenship/ownership percentages are verified from "
            "company records and filled into the letter.",
            st["Meta"],
        ),
    ]
    template.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print("Local partner letter:", path)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    build_readme()
    build_field_map()
    build_sf424()
    build_sf424a()
    build_sflll()
    build_abstract()
    build_narrative()
    build_budget_narrative()
    build_local_partner_letter()
    print("Done ->", OUT)
    print("Note: Original Adobe XFA shells cannot be binary-filled here; use filled/ companions + Grants.gov webforms.")


if __name__ == "__main__":
    main()
