#!/usr/bin/env python3
"""Build the Nexa Climate and Health 2026 FairBanks application pack.

Creates one synchronized set:
  documents/nexa-climate-health_word.docx
  documents/nexa-climate-health_pdf.pdf
  documents/nexa-climate-health_ppt.pptx

The Word/PDF files are a Proof-of-Concept answer bank and submission workbook.
The PowerPoint is a 12-slide visual proposal summary.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ASSETS = REPO / "assets"
OUT = HERE / "documents"
DOCX = OUT / "nexa-climate-health_word.docx"
PDF = OUT / "nexa-climate-health_pdf.pdf"
PPTX = OUT / "nexa-climate-health_ppt.pptx"

OFFICIAL_PAGE = "https://www.grandchallenges.ca/rfp-nexa/"
OFFICIAL_RFP = "https://www.grandchallenges.ca/wp-content/uploads/2026/06/EN-Nexa-RFP-Jun-18-2026.pdf"
OFFICIAL_POC = "https://www.grandchallenges.ca/wp-content/uploads/2026/06/Nexa-POC-application_for-webpage.pdf"
OFFICIAL_TTS = "https://www.grandchallenges.ca/wp-content/uploads/2026/06/Nexa-TTS-Stage-1-application_for-webpage.pdf"
FLUXX = "https://gcc.fluxx.io"

NAVY = "0A1F2E"
TEAL = "0D6E6E"
GREEN = "2D7A55"
ORANGE = "C45C26"
GOLD = "D99A2B"
CREAM = "F7F5F0"
PALE_TEAL = "E8F3F2"
PALE_ORANGE = "FBEDE6"
PALE_GREEN = "E9F2EC"
SLATE = "1E2F38"
MUTED = "52636C"
LINE = "CED9D8"
WHITE = "FFFFFF"
RED = "A3312D"

PROGRAMME = "Nexa Climate and Health Initiative 2026"
TITLE = "FCHIP Climate Health Early Action"
SUBTITLE = "Turning rainfall, heat, and community health signals into timely care in Uganda"
SLOGAN = "Your health, our mission."
DEADLINE = "22 July 2026, 2:00 p.m. ET / 6:00 p.m. UTC"
TRACK = "Proof of Concept - climate-informed early warning and monitoring systems"
REQUEST = "Up to USD 200,000 for 18, 21, or 24 months"
CONFIRM = "[CONFIRM BEFORE SUBMISSION]"

PHOTOS = {
    "cover": "cover_hero_cinematic.jpg",
    "logo": "fairbanks_logo.jpeg",
    "outreach": "outreach_bp_screening.jpeg",
    "mobile": "outreach_mobile_phone_demo_01.jpg",
    "dashboard": "dashboard_demo.png",
    "gis": "gis_hotspots.png",
    "architecture": "data_flow_iso_labeled.png",
    "maternal": "bloom_maternal_health_participant_01.jpg",
    "elderly": "gericare_wheelchair_assist.jpeg",
    "community": "outreach_audience_full_group_01.jpg",
    "facility": "facility_exterior_entrance_01.jpg",
    "team": "staff_team_reception.jpeg",
}

CONCEPT = {
    "cascade": REPO / ".cursor" / "concept_improved.jpeg",
    "simple": REPO / ".cursor" / "concept_simple.jpeg",
    "classic": REPO / ".cursor" / "concept.jpeg",
}


def photo(key: str) -> Path:
    path = ASSETS / PHOTOS[key]
    if not path.exists():
        raise FileNotFoundError(path)
    return path


def concept(key: str = "cascade") -> Path:
    path = CONCEPT[key]
    if not path.exists():
        raise FileNotFoundError(path)
    return path


CALL_FACTS = [
    ("Call", "Nexa - Climate and Health Innovation: Africa, Latin America and the Caribbean"),
    ("Co-leads", "Grand Challenges Canada and the Science for Africa Foundation"),
    ("Deadline", DEADLINE),
    ("Submission", "Fluxx only; complete application in English or French"),
    ("Recommended stage", "Proof of Concept; current evidence does not support Transition to Scale"),
    ("Amount", "Up to USD 200,000"),
    ("Duration", "18, 21, or 24 months"),
    ("Applicant", "Incorporated or equivalent in Africa or Brazil; Uganda is eligible"),
    ("Implementation", "Uganda is an eligible Proof-of-Concept implementation country"),
    ("Community threshold", "At least community linked"),
    ("Climate bar", "Must be climate-integrated: climate shapes design, function, and deployment"),
    ("Innovation Screen", "Only seven Overview answers reviewed first; typically over 80% declined there"),
]

READINESS = [
    ("Applicant legal name and entity type", CONFIRM),
    ("Incorporated, active, and in good standing in Uganda", CONFIRM),
    ("Project Lead identity, role, email, and one-application confirmation", CONFIRM),
    ("Requested amount and 18/21/24-month duration", CONFIRM),
    ("Collaborators and written roles", CONFIRM),
    ("Climate/weather data source and permission or licence", CONFIRM),
    ("District, facility, community, ethics, and data approvals", CONFIRM),
    ("Baseline values, sample, targets, and budget", CONFIRM),
    ("Prior GCC applications/awards and third-party IP", CONFIRM),
]

PROBLEM = (
    "Climate change is making malaria risk less predictable in Kampala and surrounding "
    "communities. Shifting rainfall, temperature, and humidity can change mosquito breeding "
    "and transmission, while extreme heat adds risk for pregnant women and people living "
    "with cardiovascular disease or diabetes. Yet local health actors mainly see separate, "
    "late signals: CHW and VHT reports, outreach screening records, facility visits, medicine "
    "use, and weather information. These signals are rarely combined into simple triggers "
    "that tell a CHW, clinic, or programme when and where to act. Underserved households, "
    "including pregnant women, children, older people, and residents of informal urban "
    "communities, therefore face delayed prevention, testing, referral, follow-up, and stock "
    "preparation. FairBanks needs a practical local system that turns climate-driven risk "
    "signals into timely health service action, not another dashboard that stops at an alert."
)

INNOVATION = (
    "FCHIP Climate Health Early Action is a climate-integrated FairBanks platform component "
    "on the Community Reach Data and Feedback loop. It will combine approved rainfall, "
    "temperature, humidity, and seasonal data with structured CHW/VHT reports, outreach "
    "screening, facility trends, and medicine-use signals. An offline-capable mobile "
    "workflow will help frontline workers collect and review essential data. Explainable "
    "risk rules will first identify time-and-place malaria risk and heat-sensitive risk for "
    "pregnancy, cardiovascular disease, and diabetes. A GIS action board will show risk, "
    "data quality, and agreed response triggers. Each trigger will connect to a named "
    "protocol: targeted health messages, household follow-up, malaria testing/referral, "
    "outreach scheduling, clinician review, or medicine and supply preparation. The Proof "
    "of Concept will test whether this end-to-end loop is feasible, accepted, timely, and "
    "able to improve access to climate-responsive care. Models will not diagnose or replace "
    "clinical judgement."
)

POPULATION = (
    "The pilot will focus on climate-vulnerable households in FairBanks Community Reach "
    "areas in and around Kampala, with priority attention to pregnant women, children, older "
    "people, and people living with cardiovascular disease or diabetes. These groups may be "
    "more exposed to mosquito-borne infection, heat stress, disrupted care, and financial "
    "barriers. CHWs and VHTs will help identify needs, test language and workflows, explain "
    "consent, collect only necessary data, and connect households to services. Community "
    "members will join design sessions, usability tests, feedback meetings, and review of "
    "what actions followed alerts. Access will not depend on owning a smartphone: frontline "
    "workers will use offline-capable tools and existing outreach channels. Participation "
    "targets, locations, safeguarding steps, disability access, and gender-responsive "
    "engagement must be confirmed with communities before submission and implementation."
)

INNOVATIVE = (
    "Current approaches often keep weather information, community reports, and clinical "
    "records in separate systems. Some early warning tools stop at forecasting and leave "
    "local teams to decide what to do. FCHIP is designed around the full early-action chain: "
    "community signal, climate signal, explainable trigger, named actor, service response, "
    "and follow-up result. It combines technological innovation (offline capture, secure data "
    "integration, interpretable analytics, and GIS), social innovation (CHW/VHT-led design "
    "and community feedback), and a delivery model rooted in a functioning medical centre "
    "and Community Reach programmes. It is a better fit for a lower-resource setting because "
    "it starts with simple validated rules, works with intermittent connectivity, and keeps "
    "clinicians responsible for care decisions. The pilot will compare timeliness, completion, "
    "acceptability, and cost with current manual planning before adding any machine learning."
)

OUTCOMES = (
    "The primary aim is improved and more timely access to prevention, testing, referral, "
    "monitoring, and treatment for climate-sensitive conditions. For malaria, rainfall and "
    "community symptom signals should trigger earlier targeted communication, testing, "
    "outreach, and supply preparation. For heat-sensitive pregnancy, cardiovascular disease, "
    "and diabetes, heat thresholds should trigger follow-up, hydration and risk messages, "
    "clinician review, and continuity-of-care checks. The pilot will test whether more "
    "at-risk people receive a documented service action after a climate-informed alert, "
    "whether response time falls, and whether referral and follow-up completion improve. "
    "It will also track false alerts, missed events, equity of reach, service continuity, and "
    "unintended harm. Final baseline values and realistic targets must be set from verified "
    "FairBanks and partner data; technical alert accuracy alone will not count as health impact."
)

CLIMATE_DESIGN = (
    "FCHIP began as a community health intelligence concept and is being fundamentally "
    "re-engineered for this pilot as a climate-integrated health innovation under Nexa's "
    "definition. Climate factors will shape its design through rainfall, temperature, "
    "humidity, seasonality, and agreed hazard thresholds; its functioning through "
    "climate-sensitive risk rules and alert timing; and its deployment through risk-based "
    "outreach, staffing, follow-up, and supply preparation. The innovation will be tested "
    "across changing climate-risk periods rather than as a static health register. The team "
    "will document whether performance and service response change during heavy rainfall, "
    "heat, poor connectivity, or service pressure. Exact climate data provider, geographic "
    "resolution, historical coverage, update frequency, licences, and locally approved "
    "thresholds: " + CONFIRM + "."
)

TEAM_CONNECTION = (
    "FairBanks operates a medical centre and Community Reach cascade serving Kampala-area "
    "communities including Bukoto, Kyebando, Kisaasi, Kamwokya, and Kikaaya. The cascade "
    "links community members, CHWs/VHTs, outreach programmes, clinical care, research and "
    "skills, and economic empowerment including CHIS where relevant. This gives the project "
    "a local service presence and a practical route for community design, referral, and "
    "follow-up. The application must still name the Project Lead, technical and clinical "
    "leads, CHW/VHT representatives, MEL support, data protection responsibility, and "
    "confirmed collaborators, with training, lived experience, location, role, and prior "
    "partnership outcomes. Do not claim community-owned or community-led status until "
    "leadership and residence evidence confirms it."
)

OVERVIEW_QUESTIONS = [
    ("1. Specific climate-driven health gap [1,500]", PROBLEM, 1500),
    ("2. Proposed innovation [2,000]", INNOVATION, 2000),
    ("3. Population design and access [1,500]", POPULATION, 1500),
    ("4. Innovation in the target setting [2,000]", INNOVATIVE, 2000),
    ("5. Potential to improve priority outcomes [2,000]", OUTCOMES, 2000),
    ("6. Climate influence on design and deployment [1,500]", CLIMATE_DESIGN, 1500),
    ("7. Team and community connection [1,500]", TEAM_CONNECTION, 1500),
]

EXECUTION = [
    (
        "1. Execution plan [3,000]",
        "Months 1-3: confirm governance, communities, users, climate data, baseline, protocols, "
        "ethics and safeguarding; co-design workflows. Months 4-7: build and test offline forms, "
        "data pipeline, explainable rules, GIS action board, and audit trail. Months 8-10: train "
        "users and run a small usability and data-quality test. Months 11-18: pilot agreed malaria "
        "and heat-sensitive care workflows across selected sites and seasons, with monthly safety "
        "and learning reviews. Months 19-21: analyse quantitative and qualitative results, refine "
        "thresholds and workflows, and assess cost and sustainability. Months 22-24: validate "
        "findings with communities and partners, publish a responsible evidence package, and make "
        "a scale/no-scale decision. Exact sites, sample, milestones, owners, and duration: " + CONFIRM,
    ),
    (
        "2. Community engagement [2,000]",
        "Use paid or otherwise fairly supported CHW/VHT and community representation in design, "
        "testing, interpretation, and governance. Hold separate listening sessions where needed "
        "for pregnant women, older people, people with disabilities, and low-income residents. "
        "Test language, consent, alert burden, referral barriers, and preferred communication. "
        "Publish a feedback route and show what changed. Track participation by sex, age, "
        "disability where appropriate, and location without collecting unnecessary personal data. "
        "Community representatives, compensation, safeguarding route, and engagement targets: " + CONFIRM,
    ),
    (
        "3. Risks and mitigation [2,000]",
        "Main risks are weak or biased data, false or missed alerts, alert fatigue, privacy harm, "
        "poor connectivity, low adoption, exclusion, unsafe referrals, partner delay, climate-data "
        "licensing, corruption, and supply constraints after an alert. Mitigation includes data "
        "minimisation, consent, role access, encryption, audit logs, clinician oversight, threshold "
        "validation, offline workflows, escalation protocols, safeguarding, complaints handling, "
        "anti-fraud controls, partner agreements, and pause criteria. Applicable Uganda approvals, "
        "ethics review, incident response, insurance, and organisational policies: " + CONFIRM,
    ),
    (
        "4. Expected impact [2,000]",
        "Expected results are improved capacity to interpret climate-health signals; faster action "
        "after risk triggers; more at-risk people reached with prevention, testing, referral, or "
        "continuity-of-care support; improved referral completion; and better preparation of "
        "outreach and supplies. The pilot will report reach and outcomes by vulnerable group and "
        "location. It will not promise reductions in illness or hospitalisation without a powered "
        "evaluation. Verified baseline, sample size, effect assumptions, targets, and attribution "
        "approach: " + CONFIRM,
    ),
    (
        "5. Proof of Concept and objectives [2,000]",
        "Establish that an end-to-end climate-informed workflow is technically feasible, acceptable, "
        "safe, affordable enough to continue, and able to improve timely access to service actions. "
        "Proposed objectives: (1) achieve an agreed completeness and timeliness level for core data; "
        "(2) validate malaria and heat-risk triggers against observed data; (3) reduce median time "
        "from trigger to documented action; (4) improve referral/follow-up completion; (5) reach "
        "priority groups equitably without serious safety or privacy events; and (6) secure a "
        "costed continuation decision from FairBanks and confirmed partners. Numeric targets: " + CONFIRM,
    ),
    (
        "6. Monitoring and evaluation [2,500]",
        "Use a prospective mixed-methods pilot with a pre-defined theory of change. Compare baseline "
        "and implementation periods and, if feasible, matched workflows or sites. Measure data "
        "quality, alert performance, response time, service actions, referrals, continuity of care, "
        "reach, equity, acceptability, adoption, cost per person reached, and safety events. Sources "
        "include system logs, CHW forms, facility records, referral registers, climate data, surveys, "
        "interviews, and focus groups. Use confidence intervals and transparent missing-data analysis; "
        "report false positives and negatives. An independent statistician or research partner, final "
        "design, power/sample calculation, tools, approvals, and baseline: " + CONFIRM,
    ),
    (
        "7. Sustainability [2,000]",
        "Build on existing FairBanks outreach and clinical workflows, use low-cost offline tools, "
        "train local users, document protocols, and test willingness and ability to pay. Future "
        "routes may include clinic subscriptions, programme implementation contracts, district or "
        "NGO deployments, and approved integrations. Sustainability depends on evidence, data cost, "
        "maintenance, workforce time, public-system fit, and reliable service capacity after alerts. "
        "The pilot will produce total-cost-of-ownership and financing scenarios, not assume grant-free "
        "scale. Confirmed buyers, government pathway, partner commitments, and pricing evidence: " + CONFIRM,
    ),
    (
        "8. Project team [1,500]",
        "FairBanks contributes local primary-care delivery, outreach, CHW/VHT links, and community "
        "follow-up. The project requires named leadership across clinical care, climate/epidemiology, "
        "product engineering, data protection, MEL, community engagement, finance, safeguarding, and "
        "partnerships. Attach roles and time commitments and close gaps through collaborators. Names, "
        "qualifications, employment links, CV evidence, and collaborator commitments: " + CONFIRM,
    ),
    (
        "9. Prior partnerships [1,500]",
        "Describe only verified work with community groups, health facilities, government, academic, "
        "private-sector, and development partners. State each partner, dates, role, activity, and "
        "result. Repository material supports a community-health ecosystem but does not verify every "
        "formal partnership. Evidence and approved wording: " + CONFIRM,
    ),
    (
        "10. Budget",
        "Complete Fluxx categories: remuneration; subcontractor fees; travel; goods and supplies; "
        "equipment; project administration; sub-grants; and indirect costs capped at 10% of direct "
        "costs. Certify that the budget is complete and that most activities and expenses occur in "
        "the eligible implementation country. Detailed amount, currency assumptions, quotations, "
        "indirect-cost basis, and majority-spend confirmation: " + CONFIRM,
    ),
]

TECH = [
    ("Climate inputs", "Rainfall, temperature, humidity, seasonality, and approved thresholds"),
    ("Health inputs", "CHW/VHT reports, outreach screening, facility trends, and medicine use"),
    ("Responsible analytics", "Explainable rules first; validation, monitoring, and clinician review"),
    ("Action layer", "GIS risk board, named response owner, protocol, referral, and follow-up"),
    ("Learning loop", "Response outcome, data quality, equity, safety, cost, and threshold refinement"),
]

MEL = [
    ("Data readiness", "Completeness, timeliness, missingness, geographic coverage"),
    ("Alert performance", "Sensitivity, specificity, false alerts, missed events, lead time"),
    ("Intermediary response", "Local actor capacity to interpret signals; surge/outreach adaptation; service continuity"),
    ("Priority outcomes", "Timely prevention, testing, care access, continuity; malaria and heat-sensitive care"),
    ("Climate adaptiveness", "Performance compared across rainfall, heat, and service-pressure periods"),
    ("Equity and safety", "Reach by priority group, complaints, adverse events, privacy incidents"),
    ("Feasibility", "Adoption, acceptability, uptime, cost per person reached, staff burden"),
]

CASCADE = [
    ("1. Community members", "Needs, participation, and household climate-health signals"),
    ("2. CHWs / VHTs", "Bridge: outreach, education, referral, offline data capture"),
    ("3. Community Reach programmes", "Screening, MCH, NCD, school/community education"),
    ("4. FairBanks Medical Centre", "Clinical review, diagnostics, pharmacy, referral QA"),
    ("5. Research · partnerships · skills", "Evidence, climate expertise, training, ethics"),
    ("6. Economic empowerment / CHIS", "Affordable access and resilient households where data exists"),
]

RISKS = [
    ("Overclaiming climate fit", "Make climate data and hazard-responsive action central, not decorative."),
    ("Weak evidence", "Set baseline, comparison logic, sample, tools, and analysis before launch."),
    ("Unsafe automation", "No diagnosis; clinician oversight, conservative thresholds, and pause rules."),
    ("Privacy and ethics", "Minimise data, obtain consent/approval, control access, and log activity."),
    ("Exclusion", "Offline CHW-mediated access; disability, gender, language, and poverty checks."),
    ("No capacity after alert", "Pre-agree service protocols, owners, supplies, referral routes, and escalation."),
    ("Financial misuse", "Segregated duties, procurement controls, declarations, audit trail, and reporting."),
]

SOURCES = [
    ("Official Nexa call page", OFFICIAL_PAGE),
    ("English Funding Opportunity", OFFICIAL_RFP),
    ("Proof-of-Concept questions", OFFICIAL_POC),
    ("Transition-to-Scale questions", OFFICIAL_TTS),
    ("Fluxx application portal", FLUXX),
    ("Nexa initiative site", "https://www.nexaclimate.org"),
]


def _shade(cell, color: str) -> None:
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))


def build_docx() -> None:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = sec.bottom_margin = Inches(0.65)
    sec.left_margin = sec.right_margin = Inches(0.72)
    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].font.color.rgb = RGBColor.from_string(SLATE)
    for name, size, color in (("Title", 29, NAVY), ("Heading 1", 20, NAVY), ("Heading 2", 14, TEAL)):
        style = styles[name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)

    def run(r, size=10, bold=False, color=SLATE, italic=False):
        r.font.name = "Aptos"
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = RGBColor.from_string(color)

    def para(value="", size=10, bold=False, color=SLATE, italic=False, align=None, after=5):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        if align is not None:
            p.alignment = align
        run(p.add_run(value), size, bold, color, italic)
        return p

    def heading(value, level=1):
        p = doc.add_heading(value, level)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(5)

    def bullets(items: Iterable[str]):
        for item in items:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            run(p.add_run(item))

    def table(headers, rows, widths=None, compact=False):
        t = doc.add_table(rows=1, cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for i, value in enumerate(headers):
            cell = t.rows[0].cells[i]
            _shade(cell, NAVY)
            cell.text = ""
            run(cell.paragraphs[0].add_run(str(value)), 8.5, True, WHITE)
        for idx, values in enumerate(rows):
            cells = t.add_row().cells
            for i, value in enumerate(values):
                cell = cells[i]
                cell.text = ""
                if idx % 2 == 0:
                    _shade(cell, PALE_TEAL)
                if "CONFIRM BEFORE SUBMISSION" in str(value):
                    _shade(cell, PALE_ORANGE)
                run(
                    cell.paragraphs[0].add_run(str(value)),
                    8 if compact else 9,
                    i == 0,
                    RED if "CONFIRM BEFORE SUBMISSION" in str(value) else SLATE,
                )
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        if widths:
            for row in t.rows:
                for i, width in enumerate(widths):
                    row.cells[i].width = Inches(width)
        doc.add_paragraph().paragraph_format.space_after = Pt(1)

    def add_photo(key, width=6.6, caption=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(photo(key)), width=Inches(width))
        if caption:
            para(caption, 8, color=MUTED, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(hp.add_run("FairBanks | Nexa Climate and Health 2026"), 8, color=MUTED)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(fp.add_run(f"{SLOGAN}  |  Proof-of-Concept working pack  |  Page "), 8, color=MUTED)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    fp._p.append(field)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(photo("logo")), width=Inches(1.8))
    para(PROGRAMME, 13, True, TEAL, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(TITLE)
    para(SUBTITLE, 14, True, ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
    add_photo("cover", 6.7)
    para(SLOGAN, 12, True, TEAL, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(
        "Proof-of-Concept application workbook | Uganda | Working draft",
        9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    doc.add_page_break()

    heading("Submission readiness gate")
    para(
        "The strongest fit is Proof of Concept. Do not apply as Transition to Scale without "
        "credible, quantitative climate-health results. Do not submit until every orange field "
        "has been verified by the applicant.",
        10, True, RED,
    )
    table(["Critical item", "Current position"], READINESS, [2.8, 3.8])

    heading("1. Official call brief")
    table(["Item", "Verified information"], CALL_FACTS, [1.8, 4.8])
    heading("1.1 Why FCHIP can fit", 2)
    bullets([
        "Direct fit with climate-informed early warning and monitoring systems.",
        "Climate-integrated design: climate shapes design, function, and deployment.",
        "A clear path from climate and health signals to named service-delivery actions.",
        "Existing medical centre, Community Reach cascade, and CHW/VHT links support field testing.",
        "Uganda is eligible and the target groups match Nexa's climate-vulnerable priorities.",
        "The proposal must prove climate integration, not simply add weather data to a health app.",
    ])
    heading("1.2 Relevant hazards and outcomes", 2)
    table(
        ["Nexa scope", "Proposed focus"],
        [
            ("Changing mosquito ecology", "Malaria risk, prevention, testing, referral, and supply readiness"),
            ("Extreme heat", "Pregnancy hypertension/heat stress, cardiovascular disease, and diabetes care"),
            ("Poor air quality", "Out of initial scope unless reliable local data and respiratory workflow are confirmed"),
            ("Priority groups", "Pregnant women, children, older people, underserved urban communities, climate-sensitive NCDs"),
        ],
        [2.0, 4.6],
    )

    heading("2. Proposed investment details")
    table(
        ["Portal field", "Draft / action"],
        [
            ("Applicant organisation", CONFIRM),
            ("Project Lead", CONFIRM + " - formally affiliated, key team member, one application only"),
            ("Project title (100 characters)", TITLE),
            ("Amount requested", CONFIRM + " - maximum USD 200,000"),
            ("Duration", CONFIRM + " - select 18, 21, or 24 months"),
            ("Collaborators", CONFIRM + " - list role and legal name"),
            ("Country of incorporation", "Uganda - confirm against incorporation documents"),
            ("Legal entity type", CONFIRM),
            ("Country of implementation", "Uganda"),
            ("Community connection", CONFIRM + " - likely community partnered/linked; choose from evidence"),
            ("Area of focus", "Climate-informed early warning and monitoring systems"),
            ("Language", "English"),
        ],
        [2.2, 4.4],
        compact=True,
    )

    heading("3. Innovation Overview - exact character-limited answer bank")
    para(
        "Only paste verified text. Character counts below include spaces and exclude the question label. "
        "Keep a small buffer because Fluxx may count formatting differently.",
        9, italic=True, color=MUTED,
    )
    for question, answer, limit in OVERVIEW_QUESTIONS:
        heading(question, 2)
        para(answer)
        count = len(answer)
        para(
            f"Draft count: {count:,} / {limit:,} characters | Buffer: {limit - count:,}",
            8, True, GREEN if count <= limit else RED,
        )

    heading("4. Full Proof-of-Concept application")
    para(
        "These fields are reviewed after the Innovation Screen. They are application-ready structures, "
        "but facts and numbers marked for confirmation must be supplied before submission.",
        9, italic=True,
    )
    for question, answer in EXECUTION:
        heading(question, 2)
        para(answer)

    heading("4.1 Budget workbook", 2)
    table(
        ["Fluxx category", "Amount (USD)", "Description / basis"],
        [
            ("Remuneration", CONFIRM, CONFIRM),
            ("Subcontractor fees", CONFIRM, CONFIRM),
            ("Travel costs", CONFIRM, CONFIRM),
            ("Goods and supplies", CONFIRM, CONFIRM),
            ("Equipment costs", CONFIRM, CONFIRM),
            ("Project administration costs", CONFIRM, CONFIRM),
            ("Sub-grants", CONFIRM, CONFIRM),
            ("Indirect costs", CONFIRM, "Maximum 10% of direct costs - verify calculation"),
            ("Total", CONFIRM, "Must not exceed USD 200,000"),
        ],
        [2.0, 1.25, 3.35],
        compact=True,
    )

    heading("5. Product and action pathway")
    add_photo("architecture", 6.4, "Illustrative FCHIP architecture; not evidence of a deployed climate-health product.")
    table(["Layer", "Proof-of-Concept role"], TECH, [1.8, 4.8])
    heading("5.1 Responsible AI and clinical safety", 2)
    bullets([
        "Start with explainable thresholds and rules; add machine learning only after sufficient data and validation.",
        "Do not diagnose, prescribe, or automatically deny or delay care.",
        "Show uncertainty, data quality, and the reason for each alert.",
        "Require clinical review for care decisions and define escalation and pause criteria.",
        "Monitor bias, false alerts, missed events, workload, privacy, and unintended harm.",
    ])

    heading("6. Monitoring, evaluation, and learning")
    add_photo("dashboard", 6.4, "Illustrative dashboard concept; measures and thresholds remain to be validated.")
    table(["Domain", "Measures"], MEL, [1.8, 4.8])
    para(
        "Nexa requires meaningful health access or health outcome evidence. Forecast accuracy alone is not enough.",
        10, True, ORANGE,
    )

    heading("7. Community Reach cascade and delivery")
    add_photo_path = concept("cascade")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(add_photo_path), width=Inches(6.4))
    para(
        "Canonical FairBanks Community Reach model. FCHIP sits on the Data and Feedback loop.",
        8, color=MUTED, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    table(["Cascade layer", "Role in this Proof of Concept"], CASCADE, [2.4, 4.2], compact=True)
    bullets([
        "Community members identify needs and test usability, language, and fairness of workflows.",
        "CHWs and VHTs bridge households to prevention, screening, referral, and follow-up.",
        "Community Reach programmes coordinate outreach and health education.",
        "FairBanks Medical Centre provides clinical review, diagnostics, pharmacy, referral, and quality assurance.",
        "Research and partners support evidence, climate expertise, training, and ethical evaluation.",
        "Economic empowerment and CHIS support affordable access where enrolment data exists.",
        "Learning returns to communities and strengthens resilience, not only project reporting.",
    ])

    heading("8. Risks and safeguards")
    table(["Risk", "Mitigation"], RISKS, [1.9, 4.7], compact=True)

    heading("9. Additional portal information")
    table(
        ["Field", "Action"],
        [
            ("Previous GCC application", CONFIRM + " - Yes/No and details"),
            ("Previous GCC award", CONFIRM + " - Yes/No, grant IDs, and difference from prior work"),
            ("Citations", "Add recent, relevant climate-malaria, heat-health, CHW, and early-action evidence"),
            ("Third-party IP", CONFIRM + " - include software, models, maps, climate and health datasets"),
            ("Licences/access", CONFIRM + " - verify rights allow implementation and future scale"),
            ("Certification", "Authorised signatory must accept call terms and privacy notice"),
            ("Confidentiality", "Do not submit trade secrets or sensitive data that are unnecessary for review"),
            ("AI disclosure", "Applicant remains responsible for original, true, accurate, complete, and cited content"),
        ],
        [2.0, 4.6],
        compact=True,
    )

    heading("10. Final submission checklist")
    bullets([
        "Recheck the live call, deadline, and any revised documents.",
        "Register in Fluxx early and confirm the legal organisation record.",
        "Confirm Proof-of-Concept stage, amount, duration, and Uganda implementation.",
        "Replace every CONFIRM BEFORE SUBMISSION marker.",
        "Keep all seven Innovation Overview answers within character limits.",
        "Verify that climate factors shape design, function, and deployment.",
        "Link every risk signal to a named health service action and outcome.",
        "Complete the 24-month plan, MEL design, objectives, safeguards, and budget.",
        "Verify Project Lead, team, collaborators, community role, approvals, IP, and prior GCC history.",
        "Review for confidential information, plagiarism, unsupported claims, and citation quality.",
        "Submit in Fluxx before the deadline; save the confirmation.",
    ])

    heading("11. Official sources")
    table(["Source", "URL"], SOURCES, [2.2, 4.4], compact=True)
    para(
        f"Source check date: {date.today().isoformat()}. Official Nexa documents and Fluxx always win.",
        8, italic=True, color=MUTED,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX)
    print(f"DOCX: {DOCX}")


def convert_pdf() -> None:
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        document.ExportAsFixedFormat(str(PDF.resolve()), 17)
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()
    print(f"PDF:  {PDF}")


def _transition(slide) -> None:
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    transition = etree.Element(f"{{{ns}}}transition", spd="slow", advClick="1")
    etree.SubElement(transition, f"{{{ns}}}fade")
    slide._element.insert(2, transition)


def build_pptx() -> None:
    from PIL import Image as PILImage
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def rgb(value):
        return RGBColor.from_string(value)

    def rect(slide, x, y, w, h, fill, line=None, rounded=False):
        kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
        shape.line.color.rgb = rgb(line or fill)
        return shape

    def text(slide, value, x, y, w, h, size=18, color=SLATE, bold=False,
             align=PP_ALIGN.LEFT, font="Aptos", valign=MSO_ANCHOR.TOP):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.04)
        tf.vertical_anchor = valign
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = value
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        return box

    def bullets(slide, items, x, y, w, h, size=16, color=SLATE):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + item
            p.space_after = Pt(8)
            p.font.name = "Aptos"
            p.font.size = Pt(size)
            p.font.color.rgb = rgb(color)
        return box

    def crop(slide, path, x, y, w, h):
        with PILImage.open(path) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
        frame_ratio, image_ratio = w / h, iw / ih
        if image_ratio > frame_ratio:
            amount = (1 - frame_ratio / image_ratio) / 2
            pic.crop_left = pic.crop_right = amount
        else:
            amount = (1 - image_ratio / frame_ratio) / 2
            pic.crop_top = pic.crop_bottom = amount
        return pic

    def slide():
        s = prs.slides.add_slide(blank)
        rect(s, 0, 0, 13.333, 7.5, CREAM)
        _transition(s)
        return s

    def band(s, kicker, title_value, subtitle=""):
        rect(s, 0, 0, 13.333, 0.12, TEAL)
        text(s, kicker.upper(), 0.55, 0.35, 5.2, 0.3, 10, ORANGE, True)
        text(s, title_value, 0.55, 0.72, 12.1, 0.65, 27, NAVY, True, font="Aptos Display")
        if subtitle:
            text(s, subtitle, 0.58, 1.42, 11.8, 0.38, 12, MUTED)

    def footer(s, number):
        text(s, "FairBanks | FCHIP Climate Health", 0.55, 7.12, 4.2, 0.2, 8, MUTED)
        text(s, f"{number:02}", 12.2, 7.1, 0.5, 0.2, 8, MUTED, align=PP_ALIGN.RIGHT)

    s = slide()
    crop(s, photo("cover"), 0, 0, 13.333, 7.5)
    rect(s, 0, 0, 7.7, 7.5, NAVY)
    text(s, PROGRAMME.upper(), 0.7, 0.62, 6.1, 0.4, 11, GOLD, True)
    text(s, TITLE, 0.7, 1.2, 6.15, 1.65, 31, WHITE, True, font="Aptos Display")
    text(s, SUBTITLE, 0.72, 3.25, 5.9, 1.1, 18, WHITE, True)
    text(s, SLOGAN, 0.72, 5.9, 4.0, 0.35, 13, GOLD, True)
    text(s, "Proof of Concept | Uganda | July 2026", 0.72, 6.38, 4.5, 0.28, 10, WHITE)

    s = slide()
    band(s, "Climate-health gap", "Local actors see the risk too late", "Climate and health signals remain separate from service action.")
    crop(s, photo("outreach"), 0.55, 1.95, 5.3, 4.85)
    bullets(s, [
        "Rainfall and heat are changing when and where health risks rise.",
        "CHW, outreach, facility, medicine, and weather signals stay fragmented.",
        "Pregnant women, children, older people, and people with NCDs face delayed care.",
        "Forecasts rarely name the local action, owner, referral, or follow-up.",
    ], 6.25, 2.1, 6.25, 3.8, 17)
    rect(s, 6.35, 6.05, 5.75, 0.55, PALE_ORANGE, ORANGE, True)
    text(s, "The missing link is climate signal -> timely health service action.", 6.6, 6.18, 5.3, 0.22, 12, ORANGE, True)
    footer(s, 2)

    s = slide()
    band(s, "The innovation", "FCHIP Climate Health Early Action", "An end-to-end climate-informed workflow built for low-connectivity care.")
    crop(s, photo("dashboard"), 6.55, 1.9, 6.15, 4.9)
    for i, (a, b) in enumerate([
        ("Sense", "Climate + community + facility signals"),
        ("Interpret", "Explainable malaria and heat-risk triggers"),
        ("Act", "Outreach, testing, referral, follow-up, supplies"),
        ("Learn", "Outcome, safety, equity, cost, and refinement"),
    ]):
        y = 2.0 + i * 1.08
        rect(s, 0.65, y, 5.35, 0.82, WHITE, LINE, True)
        rect(s, 0.65, y, 0.12, 0.82, TEAL)
        text(s, a, 0.95, y + 0.12, 1.2, 0.25, 14, NAVY, True)
        text(s, b, 2.15, y + 0.12, 3.45, 0.45, 11, MUTED)
    footer(s, 3)

    s = slide()
    band(s, "Climate integration", "Climate shapes design, function, and deployment", "Not a health app with weather added later.")
    crop(s, photo("gis"), 0.55, 1.9, 5.4, 4.9)
    bullets(s, [
        "Design: rainfall, temperature, humidity, seasonality, and local thresholds.",
        "Function: time-and-place malaria and heat-sensitive risk rules.",
        "Deployment: risk-based outreach, follow-up, staffing, and supply preparation.",
        "Performance: tested across rainfall, heat, connectivity, and service-pressure periods.",
    ], 6.3, 2.0, 6.1, 4.2, 17)
    footer(s, 4)

    s = slide()
    band(s, "Priority outcomes", "Alerts matter only when care improves", "Technical accuracy is necessary, but it is not the final outcome.")
    cards = [
        ("Malaria", "Earlier prevention, testing, referral, and supply readiness"),
        ("Pregnancy", "Heat stress, hypertension, and gestational diabetes follow-up"),
        ("NCD care", "Continuity for cardiovascular disease and diabetes during heat"),
        ("System response", "Faster action, completed referrals, and resilient services"),
    ]
    for i, (a, b) in enumerate(cards):
        col, row = i % 2, i // 2
        x, y = 0.7 + col * 6.2, 2.0 + row * 2.15
        rect(s, x, y, 5.75, 1.65, WHITE, LINE, True)
        rect(s, x, y, 5.75, 0.15, GREEN)
        text(s, a, x + 0.3, y + 0.38, 1.65, 0.35, 17, NAVY, True)
        text(s, b, x + 2.0, y + 0.35, 3.35, 0.85, 13, MUTED)
    footer(s, 5)

    s = slide()
    band(s, "People", "Designed with climate-vulnerable communities", "Access does not depend on owning a smartphone.")
    crop(s, photo("maternal"), 0.55, 1.9, 4.0, 4.85)
    crop(s, photo("elderly"), 4.75, 1.9, 3.2, 4.85)
    bullets(s, [
        "Pregnant women and children",
        "Older people",
        "People with cardiovascular disease or diabetes",
        "Underserved urban households",
        "CHWs and VHTs as frontline co-designers",
    ], 8.3, 2.05, 4.2, 4.2, 15)
    footer(s, 6)

    s = slide()
    band(s, "Proof of Concept", "What we must establish", "Feasible, accepted, safe, useful, and affordable enough to continue.")
    for i, item in enumerate([
        "Reliable core climate and health data",
        "Validated malaria and heat-risk triggers",
        "Faster documented service action",
        "Better referral and follow-up completion",
        "Equitable reach without serious harm",
        "A costed continuation and scale decision",
    ]):
        col, row = i % 3, i // 3
        x, y = 0.65 + col * 4.2, 2.0 + row * 2.0
        rect(s, x, y, 3.75, 1.5, WHITE, LINE, True)
        text(s, str(i + 1), x + 0.22, y + 0.28, 0.5, 0.45, 20, ORANGE, True)
        text(s, item, x + 0.85, y + 0.28, 2.55, 0.85, 13, NAVY, True)
    footer(s, 7)

    s = slide()
    band(s, "Execution", "A staged 24-month learning plan", "Final duration, sites, sample, owners, and targets require confirmation.")
    stages = [
        ("1-3", "Govern", "Approvals, users, baseline, climate data, protocols"),
        ("4-7", "Build", "Offline capture, pipeline, rules, GIS, audit trail"),
        ("8-10", "Test", "Training, usability, data quality, safety checks"),
        ("11-18", "Pilot", "Malaria and heat workflows across risk periods"),
        ("19-21", "Learn", "Outcomes, cost, equity, threshold refinement"),
        ("22-24", "Decide", "Evidence package and scale/no-scale decision"),
    ]
    for i, (m, a, b) in enumerate(stages):
        col, row = i % 3, i // 3
        x, y = 0.65 + col * 4.2, 1.95 + row * 2.1
        rect(s, x, y, 3.75, 1.62, WHITE, LINE, True)
        rect(s, x, y, 3.75, 0.44, TEAL)
        text(s, f"MONTHS {m}", x + 0.12, y + 0.11, 1.15, 0.18, 9, WHITE, True)
        text(s, a, x + 0.25, y + 0.7, 1.0, 0.3, 15, NAVY, True)
        text(s, b, x + 1.3, y + 0.66, 2.1, 0.65, 11, MUTED)
    footer(s, 8)

    s = slide()
    band(s, "MEL", "Measure the whole path to health action", "Mixed methods, climate adaptiveness, and no unsupported impact claims.")
    crop(s, photo("mobile"), 0.55, 1.9, 4.1, 4.9)
    for i, (a, b) in enumerate(MEL):
        y = 1.85 + i * 0.7
        text(s, a, 5.0, y, 1.85, 0.22, 11, TEAL, True)
        text(s, b, 6.9, y, 5.5, 0.4, 10, SLATE)
        rect(s, 5.0, y + 0.48, 7.0, 0.015, LINE)
    footer(s, 9)

    s = slide()
    band(s, "How FairBanks works", "Community Reach cascade + FCHIP intelligence", "FCHIP is the Data and Feedback layer — not a clinic-only app.")
    crop(s, concept("cascade"), 0.45, 1.85, 7.55, 5.0)
    bullets(s, [
        "Communities -> CHWs/VHTs -> programmes",
        "Medical centre anchors clinical action",
        "Research, skills, and partners strengthen evidence",
        "CHIS and livelihoods support access",
        "Climate + GIS alerts close the loop",
    ], 8.25, 2.05, 4.5, 4.4, 14)
    footer(s, 10)

    s = slide()
    band(s, "Readiness", "Critical facts before Fluxx submission", "Orange items cannot be guessed or inferred.")
    rect(s, 0.7, 1.95, 11.9, 4.85, PALE_ORANGE, ORANGE, True)
    text(s, "CONFIRM BEFORE SUBMISSION", 1.05, 2.25, 4.0, 0.35, 16, RED, True)
    bullets(s, [
        "Legal applicant, good standing, Project Lead, team, and collaborators",
        "Amount, duration, sites, sample, baseline, numeric targets, and full budget",
        "Climate-data source, licence, resolution, history, and thresholds",
        "Community category, approvals, safeguarding, ethics, and data protection",
        "Partner commitments, prior GCC history, third-party IP, and citations",
        "Every Innovation Overview answer within its exact character limit",
    ], 1.05, 2.9, 10.9, 3.4, 16)
    footer(s, 11)

    s = slide()
    crop(s, photo("cover"), 0, 0, 13.333, 7.5)
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    text(s, "THE PROOF-OF-CONCEPT ASK", 0.75, 0.75, 4.2, 0.35, 12, GOLD, True)
    text(s, "Help local health actors act before climate-driven risk becomes harm.", 0.75, 1.35, 8.6, 1.45, 30, WHITE, True, font="Aptos Display")
    bullets(s, [
        "Up to USD 200,000",
        "18, 21, or 24 months",
        "Climate-informed early warning linked to timely care",
    ], 0.78, 3.35, 6.3, 2.0, 18, WHITE)
    text(s, SLOGAN, 0.78, 6.35, 3.6, 0.35, 13, GOLD, True)
    text(s, FLUXX, 8.2, 6.4, 4.3, 0.25, 9, WHITE, align=PP_ALIGN.RIGHT)

    OUT.mkdir(parents=True, exist_ok=True)
    prs.save(PPTX)
    print(f"PPTX: {PPTX}")


def validate() -> None:
    from zipfile import BadZipFile, ZipFile
    import fitz
    from docx import Document
    from pptx import Presentation

    for path in (DOCX, PDF, PPTX):
        if not path.exists() or path.stat().st_size < 20_000:
            raise RuntimeError(f"Missing or unexpectedly small output: {path}")
    for path in (DOCX, PPTX):
        try:
            with ZipFile(path) as zf:
                bad = zf.testzip()
                if bad:
                    raise RuntimeError(f"Corrupt archive member: {bad}")
        except BadZipFile as exc:
            raise RuntimeError(f"Corrupt Office file: {path}") from exc

    doc = Document(DOCX)
    content = "\n".join(
        [p.text for p in doc.paragraphs]
        + [c.text for t in doc.tables for r in t.rows for c in r.cells]
    )
    for phrase in ("Submission readiness gate", "Innovation Overview", "CONFIRM BEFORE SUBMISSION", OFFICIAL_POC):
        if phrase not in content:
            raise RuntimeError(f"DOCX validation failed: {phrase}")
    for question, answer, limit in OVERVIEW_QUESTIONS:
        if len(answer) > limit:
            raise RuntimeError(f"Character limit exceeded: {question}")

    pdf = fitz.open(PDF)
    if pdf.page_count < 10:
        raise RuntimeError(f"PDF has too few pages: {pdf.page_count}")
    for page in pdf:
        blocks = page.get_text("blocks")
        for block in blocks:
            x0, y0, x1, y1, text_value = block[:5]
            if x0 < -2 or y0 < -2 or x1 > page.rect.width + 2 or y1 > page.rect.height + 2:
                raise RuntimeError(f"PDF text outside page bounds: {text_value[:60]}")
    deck = Presentation(PPTX)
    if len(deck.slides) != 12:
        raise RuntimeError(f"Expected 12 slides, found {len(deck.slides)}")
    print(f"Validated: {pdf.page_count} PDF pages | {len(deck.slides)} PPT slides")
    pdf.close()


def main() -> None:
    print(f"Building {PROGRAMME} application pack")
    print(f"Source check date: {date.today().isoformat()}")
    OUT.mkdir(parents=True, exist_ok=True)
    build_docx()
    convert_pdf()
    build_pptx()
    validate()
    print("Application pack complete.")


if __name__ == "__main__":
    main()
