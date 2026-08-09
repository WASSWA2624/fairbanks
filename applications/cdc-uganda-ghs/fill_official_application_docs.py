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

# Year 1 federal ask on SF-424A = $3,300,000 across exactly 4 activity rows
# (matches official SF-424A V1.0 / Grants.gov Section A limit).
# NOFO Component 5 (humanitarian) contingency plan stays in the narrative;
# SF-424A dollars are entered on Components 1-4 only.
C1, C2, C3, C4 = 2_100_000, 400_000, 500_000, 300_000
TOTAL = C1 + C2 + C3 + C4  # 3,300,000
assert TOTAL == 3_300_000

# Component 1 category split (sums to C1)
C1_CAT = {
    "Personnel": 610_000,
    "Fringe Benefits": 122_000,
    "Travel": 75_000,
    "Equipment": 52_000,
    "Supplies": 80_000,
    "Contractual": 670_000,
    "Construction": 0,
    "Other": 235_000,
    "Indirect Charges": 256_000,
}
assert sum(C1_CAT.values()) == C1


def money(n: int) -> str:
    return f"${n:,.0f}"


def money_cents(n: int) -> str:
    """Match Grants.gov SF-424A style ($X,XXX.00)."""
    return f"${n:,.0f}.00"


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
    f"{ORG} (FairBanks) is a health organisation in Kampala, Uganda. "
    "We run FairBanks Medical Centre and FairBanks Community Reach. "
    "Our CHWs and VHTs already work in Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya "
    "and nearby communities. We meet people in homes, schools, and local gatherings, "
    f"not only at the clinic. Our slogan is simple: {SLOGAN} "
    "Uganda has made progress on health security, but the 2023 Joint External Evaluation "
    "still pointed to real gaps. Community and clinic reports often reach decision makers "
    "too late. Information does not always move well between sectors. Subnational teams "
    "need more workers ready for surge response. Links into national surveillance, "
    "including the National Integrated Surveillance System (NISS), are still incomplete. "
    "When that happens, outbreaks grow before anyone can act. Ugandans feel the harm first. "
    "Americans are safer when threats are stopped closer to where they start. "
    "This project helps close that last mile gap under Ministry of Health leadership. "
    "We are not trying to replace national systems. We want community and facility signals "
    "to reach MoH pathways faster and in a form government can use. "
    "In Year 1 (Component 1, Core GHS) we will focus on four practical things: "
    "(1) community and facility surveillance that feeds MoH and NISS pathways using our FCHIP "
    "tools; (2) faster detect, notify, and respond timing with districts, using CDC 717 style "
    "benchmarks; (3) training CHWs, VHTs, and frontline staff for surveillance and surge "
    "readiness, including One Health awareness; and (4) stronger links between communities "
    "and facilities for priority diseases such as HIV, TB, malaria, cholera, viral "
    "hemorrhagic fevers, mpox, and immunisation work. For laboratory and border health, "
    "we will support MoH aligned partners rather than claim a national lead role alone. "
    "Components 2 to 5 are contingency plans for small and large outbreaks, emerging "
    "threats, and humanitarian emergencies. CDC may approve them but leave them unfunded "
    "until emergency money is released. "
    "If we succeed, alerts should move faster, community reports should be more complete, "
    "frontline workers should be better prepared, and tools and data should be handed to "
    f"government for continued use. Contact: {PD_NAME}, {EMAIL}, {PHONE}."
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
| `WS01739425-SF424A-V1.0.pdf` | Copy into SF-424A (exactly 4 activity rows totaling $3.3M) |
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
        "=== SF-424A Year 1 (exactly 4 Section A rows = official form) ===",
        f"APPLICATION FEDERAL TOTAL (SF-424 Box 18a / 18g): {money(TOTAL)}",
        "",
        "Section A — four activity rows only (line 5 Totals = $3,300,000):",
        f"  1. Component 1 - Core GHS priorities | CFDA {CFDA} | (e) {money_cents(C1)}",
        f"  2. Component 2 - Small-scale outbreak / PHE response | (e) {money_cents(C2)}",
        f"  3. Component 3 - Large-scale outbreak / PHE response | (e) {money_cents(C3)}",
        f"  4. Component 4 - Emerging infectious disease threats | (e) {money_cents(C4)}",
        f"  5. Totals (e)/(g): {money_cents(TOTAL)}",
        "  Columns (c)/(d)/(f) = $0.00 on every row (new application; no cost share).",
        "",
        "Section B — object classes 6a-6k; columns (1)-(4) = Section A activities 1-4; (5) Total = $3,300,000.",
        "Section D line 13 Federal Year 1 = $3,300,000.",
        "Component 5 humanitarian contingency plan: described in Project/Budget Narrative (NOFO);",
        "  not a fifth SF-424A dollar row (form limit is 4). Activation would use CDC emergency funding.",
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
            "Note: Year 1 federal total above equals the four SF-424A activity rows "
            f"({money(C1)} + {money(C2)} + {money(C3)} + {money(C4)} = {money(TOTAL)}). "
            "Only Component 1 is expected to be funded initially; Components 2-4 may be "
            "approved but unfunded until CDC activates emergency funding. "
            "Component 5 humanitarian contingency plan is in the narrative (NOFO) and is not "
            "a fifth SF-424A dollar row. CONFIRM final figures.",
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
    """Companion PDF matching official SF-424A V1.0: exactly 4 Section A rows totaling $3.3M."""
    st = styles()
    path = OUT / "WS01739425-SF424A-V1.0.pdf"
    template, on_page = doc_template(path, "SF-424A")
    pw = letter[0] - 1.5 * inch
    z = money_cents(0)

    def split(total, weights):
        raw = {k: int(total * w) for k, w in weights.items()}
        gap = total - sum(raw.values())
        raw["Other"] = raw.get("Other", 0) + gap
        return raw

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
    # Exactly 4 columns — matches official form / Grants.gov webform
    col_budgets = [C1_CAT, split(C2, w_surge), split(C3, w_surge), split(C4, w_surge)]
    col_totals = [C1, C2, C3, C4]
    assert sum(col_totals) == TOTAL == 3_300_000

    activities = [
        "Component 1 - Core GHS priorities",
        "Component 2 - Small-scale outbreak / PHE response",
        "Component 3 - Large-scale outbreak / PHE response",
        "Component 4 - Emerging infectious disease threats",
    ]
    amounts = [C1, C2, C3, C4]

    object_lines = [
        ("6a. Personnel", "Personnel"),
        ("6b. Fringe Benefits", "Fringe Benefits"),
        ("6c. Travel", "Travel"),
        ("6d. Equipment", "Equipment"),
        ("6e. Supplies", "Supplies"),
        ("6f. Contractual", "Contractual"),
        ("6g. Construction", "Construction"),
        ("6h. Other", "Other"),
    ]

    def section_b_rows(budgets, totals):
        rows = []
        direct_by_col = [0] * len(budgets)
        for label, key in object_lines:
            vals = [b.get(key, 0) for b in budgets]
            for i, v in enumerate(vals):
                direct_by_col[i] += v
            rows.append([label] + [money_cents(v) for v in vals] + [money_cents(sum(vals))])
        rows.append(
            ["6i. Total Direct Charges (sum of 6a-6h)"]
            + [money_cents(v) for v in direct_by_col]
            + [money_cents(sum(direct_by_col))]
        )
        indirect = [b.get("Indirect Charges", 0) for b in budgets]
        rows.append(
            ["6j. Indirect Charges"]
            + [money_cents(v) for v in indirect]
            + [money_cents(sum(indirect))]
        )
        computed = [d + i for d, i in zip(direct_by_col, indirect)]
        assert computed == totals, (computed, totals)
        assert sum(totals) == TOTAL
        rows.append(
            ["6k. TOTALS (sum of 6i and 6j)"]
            + [money_cents(v) for v in totals]
            + [money_cents(sum(totals))]
        )
        return rows

    sec_a_widths = [pw * 0.34, pw * 0.10, pw * 0.11, pw * 0.11, pw * 0.12, pw * 0.11, pw * 0.11]
    sec_b_widths = [1.35 * inch] + [0.95 * inch] * 4 + [1.0 * inch]

    story = [
        Paragraph("BUDGET INFORMATION - Non-Construction Programs", st["Cover"]),
        Paragraph("OMB Number: 4040-0006 | Expiration Date: 02/28/2025 | SF-424A (Rev. 7-97)", st["Center"]),
        Paragraph(f"{ORG} | {OPPORTUNITY} | Year 1 draft | {date.today().isoformat()}", st["Center"]),
        Paragraph(
            f"Year 1 federal ask: {money_cents(TOTAL)}. "
            "Official SF-424A V1.0 has four Section A activity rows — this companion uses the same "
            "four rows (Components 1-4). Section A line 5 Totals and Section B line 6k Total both "
            f"equal {money_cents(TOTAL)}. All figures USD. Cost share: none. "
            "Component 5 humanitarian contingency plan is in the Project/Budget Narrative (NOFO) "
            "and is not a fifth SF-424A dollar row.",
            st["Body"],
        ),
        Paragraph("SECTION A - BUDGET SUMMARY", st["H2"]),
        tbl(
            st,
            [
                "(a) Grant Program Function or Activity",
                "(b) Catalog of Federal Domestic Assistance Number",
                "(c) Federal Unobligated",
                "(d) Non-Federal Unobligated",
                "(e) Federal New or Revised",
                "(f) Non-Federal New or Revised",
                "(g) Total",
            ],
            [
                [act, CFDA, z, z, money_cents(amt), z, money_cents(amt)]
                for act, amt in zip(activities, amounts)
            ]
            + [
                [
                    "5. Totals",
                    "",
                    z,
                    z,
                    money_cents(TOTAL),
                    z,
                    money_cents(TOTAL),
                ]
            ],
            sec_a_widths,
        ),
        Paragraph(
            f"Exactly four activity rows. Line 5 Totals = {money_cents(TOTAL)}. "
            "Must match Section B line 6k column (5). Unobligated (c)/(d) and Non-Federal (f) = $0.00.",
            st["Small"],
        ),
        Paragraph("SECTION B - BUDGET CATEGORIES", st["H2"]),
        Paragraph(
            "6. Object Class Categories — columns (1)-(4) match Section A activities 1-4; "
            f"column (5) Total = {money_cents(TOTAL)}.",
            st["Small"],
        ),
        tbl(
            st,
            [
                "6. Object Class Categories",
                "(1) Comp 1",
                "(2) Comp 2",
                "(3) Comp 3",
                "(4) Comp 4",
                "(5) Total",
            ],
            section_b_rows(col_budgets, col_totals),
            sec_b_widths,
        ),
        Paragraph(
            "Ceilings (do not exceed): C1 $5M; C2 $10M; C3 $15M; C4 $15M. "
            "Component 1 expected initial funding; Components 2-4 contingency. "
            "Component 5 ceiling $20M applies if CDC later activates humanitarian emergency funding.",
            st["Small"],
        ),
        Paragraph("SECTION C - NON-FEDERAL RESOURCES", st["H2"]),
        tbl(
            st,
            ["(a) Grant Program", "(b) Applicant", "(c) State", "(d) Other Sources", "(e) TOTALS"],
            [
                [f"8. {activities[0]}", z, z, z, z],
                [f"9. {activities[1]}", z, z, z, z],
                [f"10. {activities[2]}", z, z, z, z],
                [f"11. {activities[3]}", z, z, z, z],
                ["12. TOTAL (sum of lines 8-11)", z, z, z, z],
            ],
            [pw * 0.40, pw * 0.15, pw * 0.15, pw * 0.15, pw * 0.15],
        ),
        Paragraph("SECTION D - FORECASTED CASH NEEDS", st["H2"]),
        tbl(
            st,
            ["", "Total for 1st Year", "1st Quarter", "2nd Quarter", "3rd Quarter", "4th Quarter"],
            [
                [
                    "13. Federal",
                    money_cents(TOTAL),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL - 3 * (TOTAL // 4)),
                ],
                ["14. Non-Federal", z, z, z, z, z],
                [
                    "15. TOTAL (sum of lines 13 and 14)",
                    money_cents(TOTAL),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL // 4),
                    money_cents(TOTAL - 3 * (TOTAL // 4)),
                ],
            ],
            [1.4 * inch] + [0.95 * inch] * 5,
        ),
        Paragraph(
            f"If only Component 1 is funded initially, Year 1 federal cash need is {money_cents(C1)} "
            f"(illustrative quarters {money_cents(C1 // 4)} each).",
            st["Small"],
        ),
        Paragraph("SECTION E - BUDGET ESTIMATES OF FEDERAL FUNDS NEEDED FOR BALANCE OF THE PROJECT", st["H2"]),
        Paragraph(
            "Future years (2-5) set at continuation based on progress and available funds. "
            f"Illustrative Core GHS planning level about {money(C1)}/year subject to CDC decisions. "
            "Contingency components remain event-driven.",
            st["Body"],
        ),
        Paragraph("SECTION F - OTHER BUDGET INFORMATION", st["H2"]),
        Paragraph(
            f"21. Direct Charges: see Section B 6i. 22. Indirect Charges: estimated 8% MTDC "
            f"(foreign organisation; exclusive of equipment and subawards over $25,000) — "
            f"Year 1 indirect across Components 1-4 = "
            f"{money_cents(sum(b.get('Indirect Charges', 0) for b in col_budgets))} (CONFIRM). "
            "23. Remarks: Award funds systems strengthening, surveillance, training, emergency "
            "readiness, and community-facility linkages. Routine clinical care is not charged to "
            "this award. No research proposed. "
            f"SF-424A federal total (four rows): {money_cents(TOTAL)}. "
            "Component 5 humanitarian contingency plan is described in attachments; activation "
            "would use CDC emergency funding under the NOFO Component 5 ceiling.",
            st["Body"],
        ),
        Paragraph(
            f"CHECK: Section A line 5 = Section B line 6k Total = Section D line 13 = "
            f"SF-424 Box 18a / 18g = {money_cents(TOTAL)}.",
            st["Body"],
        ),
    ]
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
    title = (
        "Helping Uganda find and stop health threats earlier through FairBanks Community "
        "Reach, CHW and VHT reporting, and FCHIP tools that feed Ministry of Health and NISS pathways"
    )
    story = [
        Paragraph("Project Abstract Summary", st["Cover"]),
        Paragraph(f"{ORG}", st["Center"]),
        Paragraph(f"{OPPORTUNITY}", st["Center"]),
        Paragraph(SLOGAN, st["Center"]),
        Paragraph(
            "Copy this into the Grants.gov Project Abstract form. Shorten it if the portal has a character limit.",
            st["Meta"],
        ),
        Paragraph(f"Project title: {title}", st["Body"]),
        Paragraph(f"Applicant: {ORG}", st["Body"]),
        Paragraph(f"Project director: {PD_NAME}, {PD_TITLE}", st["Body"]),
        Paragraph(f"Period: {START} to {END}", st["Body"]),
        Paragraph(f"Funding type: Cooperative agreement | CFDA: {CFDA}", st["Body"]),
        Paragraph(
            f"Year 1 ask (draft): {money(TOTAL)} total across components; "
            f"Component 1 (core) {money(C1)}. CONFIRM before submit.",
            st["Body"],
        ),
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
        Paragraph("4.2 Components 2-4 — Contingency (approved but unfunded until CDC activates)", st["H2"]),
        bullets(
            st,
            [
                f"Component 2 ({money(C2)} draft): surge roster; community investigation/contact support; RCCE; temporary dashboards; district IMS support.",
                f"Component 3 ({money(C3)} draft): expand surge staffing; multi-district mobilisation under MoH; sample referral volume logistics; extended RCCE; recovery support.",
                f"Component 4 ({money(C4)} draft): rapid form updates for new pathogens; sentinel intensification; partner lab referral; special training modules.",
                "Component 5 (humanitarian emergency): contingency plan in this narrative for NOFO completeness. "
                "SF-424A has only four activity rows, so Year 1 federal dollars are on Components 1-4 "
                f"totaling {money(TOTAL)}. Component 5 would activate with CDC emergency funding under the $20M ceiling.",
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
                ["TOTAL (SF-424A 4 rows)", money(TOTAL), "", "Equals Box 18a"],
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
                ["Salaries and wages", money(C1_CAT["Personnel"]), "Project Director portion; Programme Manager; M&E; Data/FCHIP; CHW supervisors; finance/admin charged to award"],
                ["Fringe benefits", money(C1_CAT["Fringe Benefits"]), "Statutory and organisational benefits on award-charged staff"],
                ["Travel", money(C1_CAT["Travel"]), "In-country field supervision, district meetings, limited regional workshops"],
                ["Equipment", money(C1_CAT["Equipment"]), "Field devices and approved IT — each item justified; furniture generally not allowed"],
                ["Supplies", money(C1_CAT["Supplies"]), "Training materials, PPE for drills/sample handling, connectivity, form printing"],
                ["Contractual", money(C1_CAT["Contractual"]), "MoH-aligned technical partners for lab/border support modules; software hosting; external audit"],
                ["Other", money(C1_CAT["Other"]), "Short-term trainers, translation, community meeting costs, evaluation support"],
                [
                    "Direct subtotal",
                    money(sum(C1_CAT[k] for k in C1_CAT if k != "Indirect Charges")),
                    "",
                ],
                ["Indirect (8% MTDC est.)", money(C1_CAT["Indirect Charges"]), "CONFIRM MTDC base"],
                ["Component 1 TOTAL", money(C1), "Within $5,000,000 ceiling"],
            ],
            [1.6 * inch, 1.2 * inch, 4.4 * inch],
        ),
        Paragraph("Components 2-4 — Contingency narratives (SF-424A dollar rows)", st["H1"]),
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
        Paragraph("Component 5 — Humanitarian emergency (narrative plan; not a fifth SF-424A row)", st["H1"]),
        Paragraph(
            "The NOFO requires a Component 5 humanitarian contingency plan. Because official SF-424A "
            "V1.0 / Grants.gov allows only four Section A activity rows, Year 1 federal dollars are "
            f"entered on Components 1-4 only and total {money(TOTAL)}. "
            "Component 5 readiness: community surveillance and RCCE in crisis-affected groups under "
            "MoH tasking; surge staffing/travel/supplies; contractual support when CDC activates "
            "humanitarian emergency funding under the $20,000,000 ceiling.",
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
