#!/usr/bin/env python3
"""Build opportunities/new_opportunities.xlsx from the July 2026
www2.fundsforngos.org scan for FairBanks / FCHIP.

Same Excel layout as opportunities/opportunities.xlsx
(per opportunities/rules/source_of_truth.mdc).
"""

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "new_opportunities.xlsx"

GENDER_BASED = "Gender-based (women/girls)"
MULTI_GENDER = "Multi-gender (all genders)"

# Only still-open, Uganda-eligible, FairBanks-applicable calls found on
# https://www2.fundsforngos.org/ (verified 20 July 2026). Prefer official
# apply URLs when known; otherwise the fundsforNGOs listing page.
# Excludes anything already tracked in opportunities/opportunities.xlsx.
OPPORTUNITIES = [
    {
        "title": "ISS African Futures & Innovation — Young Changemakers 2026",
        "url": "https://www2.fundsforngos.org/science/call-for-submissions-young-changemakers-2026/",
        "gender": MULTI_GENDER,
        "description": (
            "Institute for Security Studies call for young Africans "
            "(18–35) to share innovative ideas about Africa’s future. "
            "Visibility and networking for youth-led community health "
            "intelligence concepts. Uganda eligible via Africa-wide "
            "scope. Confirm age and submission format on the official "
            "ISS / African Futures page before applying."
        ),
        "deadline": "26 July 2026",
        "deadline_sort": "2026-07-26",
        "fit": "Visibility / youth innovation",
    },
    {
        "title": "Feminist Leadership Accelerator for Change Fellowship 2026",
        "url": "https://www2.fundsforngos.org/leadership/feminist-leadership-accelerator-for-change-fellowship-2026/",
        "gender": GENDER_BASED,
        "description": (
            "Structured leadership and learning fellowship for young "
            "women across Africa. Builds advocacy, organising, and "
            "leadership capacity useful for a women-led FairBanks "
            "community health venture. Not a direct FCHIP product grant "
            "but strong founder/team development. Confirm age and "
            "eligibility on the official programme page."
        ),
        "deadline": "27 July 2026",
        "deadline_sort": "2026-07-27",
        "fit": "Women founder capacity",
        "highlight": "92D050",
    },
    {
        "title": "DPI Safeguards Accelerator (Digital Public Infrastructure)",
        "url": "https://www2.fundsforngos.org/construction/safeguards-accelerator-supporting-safe-and-inclusive-digital-public-infrastructure/",
        "gender": MULTI_GENDER,
        "description": (
            "Global cohort programme helping NGOs, CSOs, CBOs, and UN "
            "entities build safe, inclusive safeguards into Digital "
            "Public Infrastructure. Catalytic grants up to about USD "
            "70,000 plus tools and peer learning. Relevant if FCHIP "
            "aligns with national digital health / DPI pathways and "
            "privacy safeguards. Uganda tagged; confirm active DPI "
            "project requirement on the official accelerator site."
        ),
        "deadline": "30 July 2026",
        "deadline_sort": "2026-07-30",
        "fit": "Digital health safeguards / DPI",
        "highlight": "00B0F0",
    },
    {
        "title": "UNDP African Youth Co-Creators Council 2026",
        "url": "https://www.undp.org/africa/blog/call-applications-african-youth-co-creators-council",
        "gender": MULTI_GENDER,
        "description": (
            "UNDP selects 11 young African leaders (about 18–34, Africa "
            "or diaspora) to advise on youth policy, programmes, and "
            "partnerships. High visibility for community health and "
            "digital innovation voices. Not a project grant. Ugandans "
            "eligible. Apply via the official UNDP Africa call."
        ),
        "deadline": "31 July 2026",
        "deadline_sort": "2026-07-31",
        "fit": "Youth policy advisory visibility",
    },
    {
        "title": "Uganda Public Health Fellowship Program (UPHFP) — Field Epidemiology Track",
        "url": "https://www2.fundsforngos.org/leadership/advanced-field-epidemiology-fellowship-program-uganda/",
        "gender": MULTI_GENDER,
        "description": (
            "Two-year learning-through-service fellowship strengthening "
            "Uganda’s field epidemiology workforce (outbreaks, "
            "surveillance, investigations). Directly relevant to FCHIP "
            "disease-surveillance capacity and MoH links. Individual "
            "fellowship for Ugandan public-health professionals — "
            "useful for FairBanks staff/partners. Confirm degree, "
            "employer, and MoH/CDC application steps on the official "
            "UPHFP site."
        ),
        "deadline": "31 July 2026",
        "deadline_sort": "2026-07-31",
        "fit": "Uganda epidemiology capacity (strong)",
        "highlight": "00B0F0",
    },
    {
        "title": "Uganda Public Health Fellowship Program (UPHFP) — Laboratory Leadership Track",
        "url": "https://www2.fundsforngos.org/leadership/laboratory-leadership-fellowship-program-uganda/",
        "gender": MULTI_GENDER,
        "description": (
            "Two-year UPHFP fellowship to strengthen Uganda’s public "
            "health laboratory leadership. Supports lab quality, "
            "networks, and outbreak readiness that complement FairBanks "
            "community testing and surveillance pathways. Individual "
            "Ugandan applicants; verify lab experience and official "
            "UPHFP application requirements before preparing a package."
        ),
        "deadline": "31 July 2026",
        "deadline_sort": "2026-07-31",
        "fit": "Uganda lab leadership capacity",
    },
    {
        "title": "Barry & Marie Lipman Family Prize 2027 (University of Pennsylvania)",
        "url": "https://www2.fundsforngos.org/cat/submission-open-for-lipman-family-prize-2027/",
        "gender": MULTI_GENDER,
        "description": (
            "Prize recognising nonprofits with exceptional leadership, "
            "innovation, and measurable social impact on global "
            "challenges. Possible organisational recognition path for "
            "FairBanks community health work. Highly selective; confirm "
            "nomination/application rules and geography on the official "
            "Penn Lipman Prize page."
        ),
        "deadline": "31 July 2026",
        "deadline_sort": "2026-07-31",
        "fit": "Nonprofit recognition (selective)",
    },
    {
        "title": "Climate Tech Fellowship 2026 (The New York Climate Exchange)",
        "url": "https://www2.fundsforngos.org/innovation/apply-now-for-climate-tech-fellowship-program-2026/",
        "gender": MULTI_GENDER,
        "description": (
            "Six-month hybrid fellowship for early-stage climate "
            "technology innovators (about USD 10,000 stipend). Global "
            "eligibility. Partial FairBanks fit only if framed around "
            "climate-adaptation tech tied to health resilience (heat, "
            "floods). Energy-focused teams may fit better. Verify "
            "technology scope on the official Climate Exchange page."
        ),
        "deadline": "1 August 2026",
        "deadline_sort": "2026-08-01",
        "fit": "Partial — climate tech only",
    },
    {
        "title": "UEFA Foundation for Children — Call for Projects 2026",
        "url": "https://uefafoundation.org/general-information/call-for-projects-2/",
        "gender": MULTI_GENDER,
        "description": (
            "Global grants for NGOs improving vulnerable children’s "
            "lives through sport, education, health, and inclusion. "
            "Needs 3+ years registration and audited accounts. FairBanks "
            "child-health / school-health programmes may fit if sport or "
            "healthy-lifestyle components are clear. Uganda eligible. "
            "Apply on the official UEFA Foundation portal."
        ),
        "deadline": "3 August 2026",
        "deadline_sort": "2026-08-03",
        "fit": "Child health via sport/education",
    },
    {
        "title": "Purpose Earth Grant and Mentorship Program 2027",
        "url": "https://www.purposeearth.org/apply-for-a-grant",
        "gender": MULTI_GENDER,
        "description": (
            "Grants up to about USD 10,000 plus mentorship for "
            "purpose-driven leaders and organisations on environment, "
            "climate, community empowerment, and cultural impact. Does "
            "not cover salaries/ops. Possible small support for "
            "community climate–health outreach. Global; Uganda tagged. "
            "Requires a professional reference letter."
        ),
        "deadline": "10 August 2026",
        "deadline_sort": "2026-08-10",
        "fit": "Small community/climate grant",
    },
    {
        "title": "African Futures Conference 2026 — Open Call (ISS AFI)",
        "url": "https://www2.fundsforngos.org/innovation/open-call-for-african-futures-conference-2026/",
        "gender": MULTI_GENDER,
        "description": (
            "ISS African Futures & Innovation conference (13–14 October "
            "2026, hybrid) seeking proposals on data-driven African "
            "futures. Good platform to present community health "
            "intelligence / predictive analytics ideas. Uganda eligible. "
            "Confirm proposal themes and format on the official ISS "
            "conference call."
        ),
        "deadline": "10 August 2026",
        "deadline_sort": "2026-08-10",
        "fit": "Conference / data futures visibility",
    },
    {
        "title": "UNESCO Youth Hackathon 2026 (Media & Information Literacy)",
        "url": "https://www2.fundsforngos.org/innovation/entries-open-unesco-youth-hackathon-2026/",
        "gender": MULTI_GENDER,
        "description": (
            "Team hackathon for ages 18–30 building innovative MIL "
            "solutions, including AI tools. Indirect FairBanks fit for "
            "youth digital skills and health-misinformation work. "
            "Ugandans eligible if age/team rules met. Verify themes and "
            "registration on the official UNESCO hackathon page."
        ),
        "deadline": "16 August 2026",
        "deadline_sort": "2026-08-16",
        "fit": "Youth digital / MIL (indirect)",
    },
    {
        "title": "Science Advice Skills Development Program (SASDP) 5th Cohort 2026",
        "url": "https://www2.fundsforngos.org/innovation/2026-science-advice-skills-development-program-sasdp-5th-cohort/",
        "gender": MULTI_GENDER,
        "description": (
            "Six-month mentorship for early/mid-career African "
            "researchers at the science–policy interface. Useful "
            "capacity for FairBanks evidence-to-district-policy work "
            "around community health intelligence. Not a product grant. "
            "Confirm eligibility on the official SASDP host site."
        ),
        "deadline": "17 August 2026",
        "deadline_sort": "2026-08-17",
        "fit": "Science–policy capacity",
    },
    {
        "title": "Joke Waller-Hunter Initiative — Small Grants for Young CSO Leaders",
        "url": "https://www2.fundsforngos.org/cat/rfas-small-grants-for-young-civil-society-leaders-in-developing-countries/",
        "gender": MULTI_GENDER,
        "description": (
            "Small grants helping young environmental professionals in "
            "developing countries strengthen leadership through "
            "training, internships, and conferences. Indirect fit for "
            "climate–health staff development. Confirm age, environment "
            "focus, and country eligibility on the official JWHI page."
        ),
        "deadline": "31 August 2026",
        "deadline_sort": "2026-08-31",
        "fit": "Young leader training (indirect)",
    },
    {
        "title": "Novo Nordisk Foundation — Global Science Summit Programme 2026",
        "url": "https://novonordiskfonden.dk/en/grant/global-science-summit-programme/",
        "gender": MULTI_GENDER,
        "description": (
            "Interdisciplinary research call for technology-enabled, "
            "scalable solutions on cardiometabolic disease prevention, "
            "detection, diagnosis, treatment, and long-term management. "
            "Aligns with FairBanks NCD screening and community data. "
            "Needs international collaborative teams. Verify PI rules "
            "and budget on the Novo Nordisk Foundation grant page."
        ),
        "deadline": "1 September 2026",
        "deadline_sort": "2026-09-01",
        "fit": "NCD / cardiometabolic research",
        "highlight": "00B0F0",
    },
    {
        "title": "Neurotorium Clinical Education Grants (Psychiatry & Neurology)",
        "url": "https://www2.fundsforngos.org/individuals/cfps-neurotorium-clinical-education-grants-within-psychiatry-and-neurology/",
        "gender": MULTI_GENDER,
        "description": (
            "Up to about DKK 100,000 (~USD 15,000) for nonprofit "
            "clinical education projects that raise psychiatry and "
            "neurology skills among health professionals. Narrower fit "
            "— possible if FairBanks builds clinician training linked "
            "to Gericare / brain-health outreach. Confirm nonprofit "
            "rules on the official Neurotorium site."
        ),
        "deadline": "2 September 2026",
        "deadline_sort": "2026-09-02",
        "fit": "Clinical education (narrow)",
        "highlight": "FF0000",
    },
    {
        "title": "Dementia, Mental Health and Brain Ageing School (Africa) — Nairobi",
        "url": "https://www2.fundsforngos.org/individuals/applications-open-for-dementia-mental-health-and-brain-ageing-school-africa/",
        "gender": MULTI_GENDER,
        "description": (
            "In-person professional development school in Nairobi for "
            "graduate students, early-career researchers, fellows, and "
            "medical professionals across Africa. Supports Gericare / "
            "elderly-care capacity around brain ageing and mental "
            "health. Not a project grant. Ugandans eligible. Confirm "
            "fees, travel, and selection on the official school page."
        ),
        "deadline": "30 September 2026",
        "deadline_sort": "2026-09-30",
        "fit": "Gericare / brain-health training",
    },
    {
        "title": "Encephalitis International Seed Funding for Research Projects 2026 (Africa)",
        "url": "https://www2.fundsforngos.org/individuals/encephalitis-international-seed-funding-for-research-projects-2026/",
        "gender": MULTI_GENDER,
        "description": (
            "Seed grants up to £10,000 for pilot encephalitis research "
            "based in Africa. Helps early-career researchers generate "
            "preliminary evidence. Narrow clinical-research fit unless "
            "FairBanks partners with academic investigators. Confirm "
            "PI location and ethics requirements on the official "
            "Encephalitis International page."
        ),
        "deadline": "30 September 2026",
        "deadline_sort": "2026-09-30",
        "fit": "Narrow Africa research seed",
        "highlight": "FF0000",
    },
    {
        "title": "Oneness Revival Team (ORT) SEED Grant Program 2026",
        "url": "https://onenessrevivalteam.info/grant-application/",
        "gender": MULTI_GENDER,
        "description": (
            "USD 25,000–50,000 one-year grants for NGOs worldwide with "
            "annual budgets under USD 100,000 for grassroots community "
            "impact. Possible organisational support for FairBanks "
            "community reach if budget/eligibility matches. Verify "
            "budget ceiling, allowable costs, and authenticity of the "
            "ORT portal before investing major proposal time."
        ),
        "deadline": "20 December 2026",
        "deadline_sort": "2026-12-20",
        "fit": "Small NGO general support",
    },
]


def word_count(text: str) -> int:
    return len(text.split())


def main() -> None:
    for row in OPPORTUNITIES:
        wc = word_count(row["description"])
        if wc > 100:
            raise SystemExit(f"Description too long ({wc} words): {row['title']}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Opportunities"

    headers = [
        "Project Title",
        "URL",
        "Brief Description",
        "Application Deadline",
        "Gender Category",
        "Application Status",
        "Submission Status",
        "Submission Date",
        "Application Folder",
    ]
    header_fill = PatternFill("solid", fgColor="0B3D2E")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    alt_fill = PatternFill("solid", fgColor="F3F8F5")
    gender_fill = PatternFill("solid", fgColor="F6E7F2")
    multi_fill = PatternFill("solid", fgColor="E7F0FA")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal="left")

    rows = sorted(OPPORTUNITIES, key=lambda r: (r["deadline_sort"], r["title"]))

    for i, row in enumerate(rows, 2):
        highlight = row.get("highlight")
        values = [
            row["title"],
            row["url"],
            row["description"],
            row["deadline"],
            row["gender"],
            "Not started",
            "unknown",
            "",
            "",
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(i, col, val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin
            cell.font = Font(name="Calibri", size=10)
            if highlight:
                cell.fill = PatternFill("solid", fgColor=highlight)
            elif i % 2 == 0:
                cell.fill = alt_fill
        gcell = ws.cell(i, 5)
        if not highlight:
            gcell.fill = gender_fill if row["gender"] == GENDER_BASED else multi_fill
        gcell.font = Font(name="Calibri", size=10, bold=True)
        ws.row_dimensions[i].height = 96

    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 74
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["I"].width = 20
    ws.row_dimensions[1].height = 26
    ws.auto_filter.ref = f"A1:I{len(rows) + 1}"
    ws.freeze_panes = "A2"

    meta = wb.create_sheet("Scan Notes")
    gb = sum(1 for r in rows if r["gender"] == GENDER_BASED)
    mg = len(rows) - gb
    priority = [
        r["title"] for r in rows if r.get("highlight") in {"00B0F0", "92D050"}
    ][:8]
    notes = [
        ("Scanned on", datetime.now().strftime("%Y-%m-%d %H:%M")),
        (
            "Document type",
            "NEW scan file: new_opportunities.xlsx from www2.fundsforngos.org "
            "only. Same column layout as opportunities.xlsx. Does not replace "
            "the canonical tracker. Deduped against opportunities.xlsx.",
        ),
        (
            "Applicant context",
            "FairBanks / FCHIP — Ugandan community health intelligence "
            "venture (CHW/VHT cascade, maternal/child, Gericare, NCDs, "
            "GIS + climate early warning).",
        ),
        (
            "Source",
            "https://www2.fundsforngos.org/ — homepage categories, "
            "latest-funds-for-ngos, health/innovation/HIV/children/"
            "leadership/community-development, Uganda tag pages 1–3, "
            "and July 2026 leadership listing round-up. Detail pages "
            "verified for deadline and Uganda/Africa/global eligibility.",
        ),
        (
            "Approach",
            "1) Crawl listing/tag/category pages for Deadline + titles. "
            "2) Keep deadlines on/after 20 July 2026 (or rolling). "
            "3) Filter for health, digital, AI, climate, women, community, "
            "Uganda/Africa. 4) Open detail pages; drop closed, "
            "country-blocked, or weak-sector calls. 5) Prefer official "
            "apply URLs when listed on the FFN page.",
        ),
        (
            "Eligibility filter",
            "Still open + Ugandans can apply (Uganda-specific, East Africa, "
            "Africa-wide, or global). Excluded US/UK/AU/CA-only, "
            "West-Africa-only cooking funds, Ghana/Kenya-only when Uganda "
            "excluded, and pure arts/climbing/journalism calls.",
        ),
        ("Gender categories", f"Gender-based (women/girls): {gb}. Multi-gender: {mg}. Total: {len(rows)}."),
        (
            "Highlight legend",
            "Green = women/gender priority. Blue = strong FCHIP / digital / "
            "Uganda health capacity fit. Red = narrower research or "
            "indirect fit.",
        ),
        (
            "Priority soon",
            "; ".join(priority)
            if priority
            else "See blue/green highlighted rows sorted by deadline.",
        ),
        (
            "Best FairBanks / FCHIP fits from this scan",
            "1) UPHFP Field Epidemiology 2) DPI Safeguards Accelerator "
            "3) Feminist Leadership Accelerator 4) Global Science Summit "
            "(NCD) 5) SEED Grant if budget under USD 100k 6) Purpose Earth "
            "small climate–community grant 7) UEFA Children if sport/"
            "child-health framing fits.",
        ),
        (
            "Excluded (already in opportunities.xlsx)",
            "AWIEF Pitch n Grow; Nexa Climate and Health; UN Women SIARP "
            "2.0; Brandtech AI for Good; IDRC ANeSA Cohort 2; Wellcome "
            "Snakebite Innovation Prize.",
        ),
        (
            "Excluded (closed on FFN detail check)",
            "SAFEStart+ Uganda (30 Jun); UNICEF Uganda Health & HIV EOI "
            "(20 Mar); UNICEF Nutrition EOI (31 Mar); Evidence for AI in "
            "Health (1 Apr); Google.org AI for Science (17 Apr); POWER "
            "Young Women Accelerator Cohort V (15 May); la Caixa Child "
            "Survival (26 May); Uganda Gov digital systems showcase "
            "(1 Jun); Japan GGP Uganda (15 Feb / 15 Jun variants); Mama "
            "Hope LDA (10 Jan); EAGYFF (28 Feb); MARS Awards (30 Jun); "
            "AI for Climate Action Award (3 Jul); CREATIVE Fellowship "
            "(29 May); WSA Young Innovators page still showed 2025 "
            "deadline.",
        ),
        (
            "Excluded (open but unfit / ineligible)",
            "Pan-Orthoebolavirus diagnostics (specialised BSL pathogen "
            "diagnostics); country-only calls (Philippines, Botswana "
            "Humphrey, Vietnam Humphrey, Ghana SPARK, Kenya Rise Up, "
            "Canada/UK/AU community grants, Ireland AI awards, West "
            "Africa clean cooking); climbing grants; ocean youth "
            "council; fashion/film/photography; ICSSR India fellowships; "
            "Digital Rights Exchange (rights advocacy, not health "
            "product); many AUD community arts grants.",
        ),
        (
            "Important",
            "fundsforNGOs republishes calls — ALWAYS confirm deadline, "
            "eligibility, and apply link on the official funder page "
            "before submitting. This file is a scan snapshot, not legal "
            "advice.",
        ),
        ("Build script", "opportunities/build_new_opportunities.py"),
        ("Related canonical tracker", "opportunities/opportunities.xlsx"),
    ]
    meta["A1"] = "Field"
    meta["B1"] = "Detail"
    meta["A1"].font = Font(bold=True, color="FFFFFF")
    meta["B1"].font = Font(bold=True, color="FFFFFF")
    meta["A1"].fill = header_fill
    meta["B1"].fill = header_fill
    for i, (field, detail) in enumerate(notes, 2):
        meta.cell(i, 1, field).alignment = Alignment(vertical="top", wrap_text=True)
        meta.cell(i, 2, detail).alignment = Alignment(vertical="top", wrap_text=True)
        meta.row_dimensions[i].height = 70
    meta.column_dimensions["A"].width = 28
    meta.column_dimensions["B"].width = 96

    # Fit summary sheet for quick triage
    fit = wb.create_sheet("Fit Summary")
    fit_headers = ["Deadline sort", "Project Title", "FairBanks fit note", "Gender"]
    for col, h in enumerate(fit_headers, 1):
        cell = fit.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
    for i, row in enumerate(rows, 2):
        fit.cell(i, 1, row["deadline_sort"])
        fit.cell(i, 2, row["title"])
        fit.cell(i, 3, row.get("fit", ""))
        fit.cell(i, 4, row["gender"])
        for c in range(1, 5):
            fit.cell(i, c).alignment = Alignment(wrap_text=True, vertical="top")
            fit.cell(i, c).border = thin
        fit.row_dimensions[i].height = 36
    fit.column_dimensions["A"].width = 14
    fit.column_dimensions["B"].width = 55
    fit.column_dimensions["C"].width = 36
    fit.column_dimensions["D"].width = 28

    wb.save(OUT)
    print(f"Wrote {OUT} with {len(rows)} opportunities")
    print(f"Gender-based: {gb}, Multi-gender: {mg}")


if __name__ == "__main__":
    main()
