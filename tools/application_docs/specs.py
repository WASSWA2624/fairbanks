"""Opportunity application content specs — one dict per tracker row (except awief)."""

from __future__ import annotations

COMMON_TRACTION = [
    "Functioning FairBanks Medical Centre with clinic, pharmacy, and diagnostics",
    "Active Community Reach outreach in Bukoto, Kyebando, Kisaasi, Kamwokya, Kikaaya",
    "Trusted CHW and VHT relationships for household visits and referrals",
    "Live programmes: maternal & child health, Gericare, chronic-disease screening",
    "Digital health records and outreach data ready to feed FCHIP",
    "Research and partnership mindset for ethical, evidence-based scale",
]

COMMON_DEEP_TECH = [
    ["Artificial Intelligence", "Outbreak early warning; maternal and NCD risk scoring"],
    ["Machine Learning", "Learn from community patterns, seasons, and outreach outcomes"],
    ["GIS Mapping", "Disease distribution, hotspots, and resource gaps"],
    ["Mobile Data Collection", "Offline-capable CHW/VHT capture at household level"],
    ["Cloud Computing", "Secure sync across facilities and partners"],
    ["Analytics Dashboards", "Real-time views for clinics, districts, and NGOs"],
]

COMMON_RISKS = [
    ["Data quality in the field", "Structured forms, validation, CHW training and supervision"],
    ["Low digital literacy", "Simple offline UX, local-language support, supervised rollout"],
    ["Privacy and consent", "Ethical governance aligned with Uganda data protection rules"],
    ["Adoption by partners", "Co-design with CHWs, clinics, and district users from day one"],
]


def _base(**kwargs):
    d = {
        "deep_tech": COMMON_DEEP_TECH,
        "traction": COMMON_TRACTION,
        "risks": COMMON_RISKS,
        "doc_label": "Win-Win Application Pack",
        "problem_bullets": [
            "Care often starts only after people fall sick",
            "Community health data sits in paper registers and silos",
            "High-risk pregnancies and NCDs are found too late",
            "Medicine stock-outs hit hard during disease surges",
            "Health shocks push poor families deeper into hardship",
        ],
        "solution_bullets": [
            "CHWs and VHTs capture structured data at household level",
            "FCHIP applies AI, ML, and GIS to predict risk early",
            "Alerts reach clinics, districts, and partners in time to act",
            "Built on a live medical centre and community programmes",
        ],
    }
    d.update(kwargs)
    return d


AUC = _base(
    slug="auc",
    url="https://opportunitiesforyouth.org/2026/07/16/auc-venture-lab-women-innovation-fellowship-empowering-women-entrepreneurs-in-egypt-and-beyond/",
    meta={
        "programme": "AUC Venture Lab Women Innovation Fellowship 2026",
        "title": "FairBanks Community Health Intelligence Platform (FCHIP)",
        "subtitle": "Women-led HealthTech ready for leadership, networks, and investment readiness",
        "applicant": "FairBanks / FCHIP — woman-led community health venture, Uganda",
        "deadline": "25 July 2026",
        "geography": "Uganda base · Fellowship in Cairo (Egypt–Denmark ecosystem)",
        "ask": "Select our founder/team for the Women Innovation Fellowship cohort (1 Sep–31 Oct 2026).",
    },
    exec_summary=[
        "FairBanks is building FCHIP — a community health intelligence platform that helps African communities move from reactive treatment to proactive prevention.",
        "A woman-led HealthTech venture rooted in a live medical centre and CHW outreach needs exactly what this fellowship offers: leadership training, mentorship, ecosystem access, and cross-regional exposure.",
        "Selecting FairBanks gives AUC Venture Lab a grounded African case study in deep-tech HealthTech — and gives our team the networks to grow responsibly across East Africa.",
    ],
    win_for_us=[
        "Leadership and investment-readiness coaching for a woman founder",
        "Access to Egypt–Denmark innovation networks and mentors",
        "Visibility that opens follow-on partners and capital conversations",
        "Peer learning with other women innovators across sectors",
    ],
    win_for_them=[
        "A real HealthTech / AI venture already serving communities in Uganda",
        "Storytelling power: deep tech rooted in African primary care, not imported soft services",
        "Contribution to gender-inclusive innovation across Africa and beyond Egypt",
        "Alumni potential that can mentor future women founders in HealthTech",
    ],
    problem=[
        "Women founders in African HealthTech face a double gap: building hard technology while also accessing leadership networks that are still uneven.",
        "Without structured mentorship and regional exposure, promising community-health platforms stay local and underfunded — even when the need is urgent.",
    ],
    solution=[
        "FCHIP connects CHWs, clinics, and partners through mobile capture, AI prediction, and GIS dashboards.",
        "The fellowship will sharpen how we pitch, partner, and scale this women-led deep-tech story without losing our community-first roots.",
    ],
    use_cases=[
        ("Maternal health", "Home-visit signals flag high-risk pregnancies before complications."),
        ("Disease surveillance", "Fever clusters trigger early malaria / outbreak response."),
        ("NCD hotspots", "Community BP and glucose data guide targeted screening."),
    ],
    fit_rows=[
        ["Women entrepreneurs & innovators", "Woman-led HealthTech venture with a clear founder growth need"],
        ["HealthTech / AI / sustainability", "FCHIP is AI + community data for preventive primary care"],
        ["Leadership & ecosystem access", "We seek coaching, mentors, and Egypt–Denmark exposure"],
        ["Beyond Egypt eligibility", "Ugandan women innovators are eligible; we bring East Africa context"],
    ],
    plan_bullets=[
        "Attend the full Cairo fellowship (1 Sep–31 Oct 2026) with focus on leadership and fundraising readiness",
        "Translate FairBanks Community Reach cascade into investor-ready narrative and metrics",
        "Map partnership opportunities with mentors and peers for CHW tools and district pilots",
        "Return with a 90-day execution plan for FCHIP MVP milestones in Kampala catchments",
    ],
    plan_slide_title="Fellowship journey and return plan",
    closing=[
        "AUC Venture Lab invests in women who will shape Africa's innovation future. FairBanks offers a live HealthTech platform and a community to prove impact.",
        "We ask to join the 2026 Women Innovation Fellowship so we can grow as leaders while advancing health for all.",
    ],
)

FEMINIST_AI = _base(
    slug="feminist-ai",
    url="https://opportunitiesforyouth.org/2026/07/16/ai-for-gender-equality-un-women-ai-school-opens-for-changemakers/",
    meta={
        "programme": "Young Feminist AI School 2026 (UN Women)",
        "title": "Gender-Responsive AI for Community Health",
        "subtitle": "Practical AI skills that put women, girls, and mothers at the centre of FCHIP",
        "applicant": "FairBanks team member(s) aged 18–30 — Uganda",
        "deadline": "28 July 2026",
        "geography": "Fully online · Uganda participants eligible worldwide",
        "ask": "Admit FairBanks youth / women team members to the 12-week AI School cohort.",
    },
    exec_summary=[
        "UN Women's Young Feminist AI School builds practical, no-code AI skills for advocacy, policy, and civic tech — with most places reserved for young women and girls.",
        "FairBanks needs exactly this capacity: gender-responsive AI applied to maternal health, adolescent health, and community data that too often ignores women's realities.",
        "Selecting our team strengthens FCHIP's responsible AI practice and gives UN Women a live health use case from Uganda.",
    ],
    win_for_us=[
        "Hands-on AI skills for young women on the FairBanks team",
        "Gender lens for FCHIP features (maternal risk, adolescent SRH, consent)",
        "A real project deliverable that can plug into community outreach",
        "Global peer network on feminist and responsible AI",
    ],
    win_for_them=[
        "Participants who will apply skills immediately in community health",
        "A Uganda case study linking AI for gender equality to primary care",
        "Amplification through FairBanks Community Reach programmes",
        "Contribution to the 70%+ young women participation goal with serious applicants",
    ],
    problem=[
        "AI tools in health can deepen inequality if they ignore how women and girls experience care, consent, and unpaid caregiving.",
        "Young changemakers need practical skills — not theory alone — to build gender-responsive systems.",
    ],
    solution=[
        "We will use the School's project track to design a gender-responsive AI workflow for FairBanks: safer maternal risk flags, clearer consent language, and community feedback loops.",
        "Skills return to CHW training and FCHIP product decisions, not a one-off certificate.",
    ],
    use_cases=[
        ("Maternal alerts", "Risk scores explained in plain language for CHWs and mothers."),
        ("Adolescent health", "Privacy-aware pathways for youth SRH education and referral."),
        ("Data dignity", "Consent and feedback designed with women users, not about them."),
    ],
    fit_rows=[
        ["Ages 18–30 worldwide", "Ugandan team members in range; women prioritised"],
        ["No-code practical AI", "We will ship a usable project for community health"],
        ["Advocacy / civic / responsible AI", "Health equity and gender-responsive product design"],
        ["~4 hours weekly", "Team can commit alongside FairBanks outreach duties"],
    ],
    plan_bullets=[
        "Complete all 12 weeks (Aug–Oct 2026) with full attendance",
        "Build a gender-responsive AI project tied to maternal / adolescent use cases",
        "Present findings to FairBanks clinical and CHW leads for adoption",
        "Share learnings back with UN Women alumni channels where useful",
    ],
    plan_slide_title="12-week learning and project plan",
    closing=[
        "This School turns AI skill into justice. FairBanks will turn those skills into safer care for women and girls in our communities.",
        "We respectfully request places for our eligible young women team members.",
    ],
)

GIRLCODE = _base(
    slug="girlcode",
    url="https://opportunitiesforyouth.org/2026/07/15/building-africas-next-generation-of-ai-powered-female-innovators-apply-for-the-2026-girlcode-hackathon-across-six-african-countries/",
    meta={
        "programme": "GirlCode Hackathon 2026 — Kampala",
        "title": "FCHIP Hack Challenge: AI for Community Health",
        "subtitle": "Women builders prototype tools that help CHWs act before crises",
        "applicant": "FairBanks-linked women innovators / team — Kampala, Uganda",
        "deadline": "Register before Kampala event 5–6 September 2026",
        "geography": "Kampala (UICT Nakawa) · Uganda chapter",
        "ask": "Register / host FairBanks women builders at the Kampala GirlCode Hackathon.",
    },
    exec_summary=[
        "GirlCode is growing Africa's next generation of AI-powered women innovators — and Kampala is one of the host cities.",
        "FairBanks brings a real problem set: offline CHW capture, maternal risk flags, and hotspot maps that communities need now.",
        "Together we create a win-win: women builders get a meaningful HealthTech challenge; FairBanks meets talent and prototypes that can feed FCHIP.",
    ],
    win_for_us=[
        "Visibility among Uganda's women-in-tech community",
        "Prototype ideas for FCHIP mobile and alert features",
        "Recruitment pipeline for women technologists",
        "Local proof that deep tech can be built with and for communities",
    ],
    win_for_them=[
        "A concrete HealthTech / AI challenge grounded in real CHW work",
        "Mentorship access from a live medical centre and outreach team",
        "Strong Kampala story aligned with GirlCode's six-country mission",
        "Path from hackathon demo to community impact, not shelfware",
    ],
    problem=[
        "Hackathons often invent problems. Community health already has them: paper registers, late referrals, and invisible hotspots.",
        "Women builders deserve challenges that matter — and partners who can take good ideas into the field.",
    ],
    solution=[
        "We propose an FCHIP challenge track: build lightweight AI-assisted tools for CHWs — symptom triage helpers, risk dashboards, or local-language education bots — designed for low connectivity.",
        "FairBanks mentors provide clinical and community context during the Kampala weekend.",
    ],
    use_cases=[
        ("CHW co-pilot", "Simple decision support for home visits."),
        ("Hotspot sketch", "Map fever or missed ANC clusters from sample data."),
        ("Mother messages", "Local-language tips linked to referral pathways."),
    ],
    fit_rows=[
        ["Women innovators in AI", "Team and challenge designed for women builders"],
        ["Kampala event 5–6 Sep 2026", "Local presence at UICT Nakawa"],
        ["Africa-wide GirlCode mission", "Uganda chapter contributes a HealthTech flagship challenge"],
        ["Talent & visibility", "FairBanks offers mentorship and follow-up pathways"],
    ],
    plan_bullets=[
        "Register Kampala team before the event weekend",
        "Publish a one-page FCHIP challenge brief for participants",
        "Provide on-site / virtual mentorship from FairBanks clinical and product leads",
        "Shortlist prototypes for post-hack follow-up and possible internship",
    ],
    plan_slide_title="Kampala hackathon plan",
    closing=[
        "GirlCode builds women who build Africa's future. FairBanks offers a health problem worth solving.",
        "We ask to participate fully in the Kampala 2026 hackathon with an FCHIP challenge and mentoring support.",
    ],
)

AUDA_SRH = _base(
    slug="auda-srh",
    url="https://opportunitiesforyouth.org/2026/06/13/african-union-youth-srh-ambassadors-initiative-2026-apply-to-become-an-auda-nepad-youth-srh-ambassador/",
    meta={
        "programme": "AUDA-NEPAD Youth SRH Ambassadors Initiative 2026",
        "title": "Youth SRH Ambassadors for Community Health Intelligence",
        "subtitle": "Advocacy and mobilisation powered by FairBanks community reach",
        "applicant": "Eligible FairBanks youth (18–35) — Uganda · voluntary ambassador role",
        "deadline": "17 July 2026 (closing)",
        "geography": "Uganda · African Union youth SRH network",
        "ask": "Appoint FairBanks youth as AUDA-NEPAD Youth SRH Ambassadors.",
    },
    exec_summary=[
        "AUDA-NEPAD is recruiting young Africans as voluntary Sexual and Reproductive Health ambassadors for advocacy, awareness, and community mobilisation.",
        "FairBanks already runs community outreach that touches adolescent and maternal health — so our youth can amplify AU messages with real community channels.",
        "This is a win-win: ambassadors gain recognition and training; AUDA-NEPAD gains Uganda-based advocates with digital presence and field access.",
    ],
    win_for_us=[
        "Continental recognition and mentorship for FairBanks youth",
        "Training that strengthens our adolescent and maternal health messaging",
        "Network access across African Union youth health spaces",
        "Sponsored participation opportunities in selected events",
    ],
    win_for_them=[
        "Ambassadors embedded in a live community health organisation",
        "Channels to CHWs, schools, and outreach events in Kampala communities",
        "Digital storytelling grounded in real SRH service pathways",
        "Diverse regional representation from East Africa / Uganda",
    ],
    problem=[
        "SRH information often fails to reach young people where they live — or arrives without a trusted referral path to care.",
        "Ambassadors need more than slogans: they need community platforms and health partners.",
    ],
    solution=[
        "FairBanks youth ambassadors will pair AUDA-NEPAD advocacy with Community Reach dialogues, school health touchpoints, and clear referral to clinical care.",
        "Where appropriate, FCHIP data (with privacy safeguards) can show where adolescent health gaps are rising — so advocacy targets the right places.",
    ],
    use_cases=[
        ("Community dialogues", "Youth-led SRH sessions linked to FairBanks outreach."),
        ("Digital advocacy", "Responsible content that drives people to care, not panic."),
        ("Referral bridge", "Ambassadors connect peers to CHWs and clinic services."),
    ],
    fit_rows=[
        ["Ages 18–35, AU citizens", "Ugandan youth applicants in range"],
        ["SRH commitment", "Active in maternal / adolescent community health work"],
        ["Digital presence", "Team can maintain responsible advocacy channels"],
        ["Voluntary service", "Clear understanding: recognition and training, not a salary"],
    ],
    plan_bullets=[
        "Complete ambassador onboarding and training requirements",
        "Run quarterly SRH awareness moments inside FairBanks outreach calendars",
        "Report activity and learning back to AUDA-NEPAD as required",
        "Protect privacy and avoid sharing identifiable patient data in advocacy",
    ],
    plan_slide_title="Ambassador activity plan",
    closing=[
        "AUDA-NEPAD needs youth who can move communities. FairBanks youth already work inside those communities.",
        "We ask for appointment as Youth SRH Ambassadors so continental goals meet local action.",
    ],
)

GADFLY = _base(
    slug="gadfly",
    url="https://opportunitiesforyouth.org/2026/07/16/the-gadfly-project-custom-web-application-grant-2026-apply-for-in-kind-software-development-support-worth-up-to-us100000/",
    meta={
        "programme": "The Gadfly Project Custom Web Application Grant 2026",
        "title": "Build FCHIP Community Capture & Partner Dashboards",
        "subtitle": "In-kind software development that expands community health impact",
        "applicant": "FairBanks Medical Centre / FairBanks Community Reach (social enterprise, Uganda)",
        "deadline": "24 July 2026 (next cycle 23 October 2026)",
        "geography": "Uganda communities · web/cloud tools usable worldwide",
        "ask": "Award in-kind engineering (valued toward US$5,000–100,000) to build FCHIP web and data tools.",
    },
    exec_summary=[
        "Gadfly funds impact through professional software development — not cash. That matches FairBanks: we have users, field workflows, and a clear product need.",
        "We request in-kind build support for offline-friendly community capture, secure sync, and partner dashboards that turn CHW data into action.",
        "Gadfly's engineers get a high-user, high-mission build; FairBanks communities get tools that prevent illness and hardship.",
    ],
    win_for_us=[
        "Professional engineering capacity we cannot yet fund at full strength",
        "Faster delivery of CHW capture and partner dashboard MVP",
        "Cleaner architecture for later AI and GIS modules",
        "Reduced risk of brittle custom code built under field pressure",
    ],
    win_for_them=[
        "Measurable users: CHWs, clinicians, and outreach teams already active",
        "Clear community impact story — prevention, referrals, stock awareness",
        "A social enterprise that will actually deploy and maintain the product",
        "Portfolio proof that in-kind software can change primary care access",
    ],
    problem=[
        "Our outreach still depends too much on paper and fragmented sheets. That slows referrals and hides emerging hotspots.",
        "We do not need another brochure website — we need durable web/app infrastructure for community health workflows.",
    ],
    solution=[
        "Scope for Gadfly: (1) CHW/VHT web/mobile-friendly capture with offline tolerance, (2) secure cloud sync, (3) role-based dashboards for clinic and programme managers, (4) export hooks for later AI models.",
        "FairBanks provides clinical workflows, user testing with CHWs, and hosting/ops ownership after delivery.",
    ],
    use_cases=[
        ("Household visit forms", "Structured maternal, child, and NCD screening capture."),
        ("Referral tracker", "See who was referred and who completed care."),
        ("Programme dashboard", "Managers see coverage gaps by community."),
    ],
    fit_rows=[
        ["Mission-driven org", "Community health social enterprise serving underserved areas"],
        ["Custom web / databases", "Capture + sync + dashboards are core Gadfly territory"],
        ["Users served over complexity", "Priority is CHW usability and community impact"],
        ["Worldwide eligibility", "Ugandan organisation qualifies"],
    ],
    plan_rows=[
        ["Discovery (weeks 1–3)", "Workflow mapping with CHWs and clinicians; prioritise MVP screens"],
        ["Build (weeks 4–14)", "Capture, sync, auth, dashboards; iterative field tests"],
        ["Handover (weeks 15–18)", "Docs, training, hosting handover, maintenance plan"],
    ],
    budget_rows=[
        ["In-kind engineering", "Gadfly team time for MVP build", "In-kind"],
        ["FairBanks counterpart", "Clinical/product time, devices, hosting, training", "In-kind + ops"],
        ["Total cash ask", "None — this is an in-kind grant", "US$0"],
    ],
    plan_slide_title="Discovery → build → handover",
    closing=[
        "Gadfly multiplies mission through code. FairBanks will put that code in the hands of CHWs who already walk our communities.",
        "We ask for selection in the current review cycle so communities benefit sooner.",
    ],
)

DOS_UGANDA = _base(
    slug="dos-uganda",
    url="https://opportunitiesforyouth.org/2026/06/15/u-s-department-of-state-launches-up-to-60-million-funding-opportunity-to-strengthen-ugandas-health-system-through-the-health-foreign-assistance-mou-implementation-plan/",
    meta={
        "programme": "U.S. Dept of State — Uganda Health System MOU (DFOP0017890)",
        "title": "Scaling Community Health Intelligence for Uganda",
        "subtitle": "Statement of Interest: digital health, community systems, and accountable care",
        "applicant": "FairBanks Medical Centre / FairBanks Community Reach — Uganda",
        "deadline": "31 July 2026 (Statement of Interest, 5:00 PM EDT)",
        "geography": "Uganda (national digital health & community health alignment)",
        "ask": "Invite FairBanks to submit a full application under Objective 4 (digital health) and related community tracks.",
    },
    exec_summary=[
        "The U.S.–Uganda Health MOU opens up to $60M across ~15 awards to strengthen Uganda's health system. Objective 4 funds scaling digital health systems including EMRs, DHIS2, eCHIS, iHRIS, and the national data warehouse.",
        "FCHIP is designed as a community intelligence layer that can complement national systems — feeding cleaner CHW data, early warnings, and facility decision support.",
        "A Statement of Interest from FairBanks offers the U.S. and Uganda a private-sector community partner with live operations — not a slide-deck-only vendor.",
    ],
    win_for_us=[
        "Pathway to multi-year partnership funding for digital community health",
        "Alignment with national systems (eCHIS / DHIS2 / data warehouse)",
        "Capacity to scale beyond Kampala catchments",
        "Stronger governance, M&E, and compliance muscle",
    ],
    win_for_them=[
        "Ugandan implementer with clinic + CHW footprint already in motion",
        "Contribution to Objective 4 digital health scaling goals",
        "Community-based complementary data that improves facility and district action",
        "Potential links to community financing, faith/community facilities, and supply accountability tracks",
    ],
    problem=[
        "Uganda's digital health stack is advancing, but community-generated signals still arrive late or incompletely — limiting outbreak response, maternal risk management, and medicine planning.",
        "Scaling digital tools without trusted community partners leaves last-mile gaps.",
    ],
    solution=[
        "FairBanks proposes FCHIP as a community intelligence complement: mobile CHW capture interoperable with national reporting where permitted, predictive alerts, GIS views, and facility dashboards.",
        "We start from live FairBanks operations, then partner outward with districts and public programmes under MOU priorities.",
    ],
    use_cases=[
        ("eCHIS-aligned capture", "Structured community forms that reduce double entry."),
        ("Surveillance support", "Fever and maternal risk signals for earlier action."),
        ("Supply awareness", "Demand signals that reduce avoidable stock-outs."),
    ],
    fit_rows=[
        ["Objective 4 — digital health", "EMR/community data, analytics, interoperability intent"],
        ["Ugandan organisations", "Local medical centre and community programmes"],
        ["Two-step process", "This pack supports a strong Statement of Interest"],
        ["Five-year systems strengthening", "Roadmap from pilot catchment to district partnership"],
    ],
    plan_rows=[
        ["SOI (by 31 Jul 2026)", "Submit concise Statement of Interest aligned to Objective 4"],
        ["Full proposal (if invited)", "Technical approach, consortium, budget, M&E, safeguarding"],
        ["Year 1–2", "Pilot interoperability + predictive modules in FairBanks catchment"],
        ["Year 3–5", "District partnerships and national reporting alignment"],
    ],
    budget_rows=[
        ["Digital platform & interoperability", "Capture, APIs, analytics, GIS", "TBD in full proposal"],
        ["CHW capacity & supervision", "Training, devices, quality assurance", "TBD"],
        ["M&E and learning", "Baselines, dashboards, reporting", "TBD"],
        ["Total", "Within award band if selected among ~15 awards", "Up to programme ceilings"],
    ],
    plan_slide_title="SOI → full proposal → scale",
    closing=[
        "The MOU succeeds when digital health reaches communities, not only servers. FairBanks is ready to help close that gap.",
        "We request consideration of our Statement of Interest for full application invitation.",
    ],
)

WHS = _base(
    slug="whs",
    url="https://opportunitiesforyouth.org/2026/06/26/apply-now-world-health-summit-2025-stipend-program-for-global-health-changemakers/",
    meta={
        "programme": "World Health Summit 2026 Youth Group Opportunity (Berlin)",
        "title": "FairBanks Youth Delegation to WHS 2026",
        "subtitle": "Complimentary group tickets for youth-led community health voices",
        "applicant": "FairBanks youth-led cohort (nominees aged 28 or under) — Uganda",
        "deadline": "31 July 2026",
        "geography": "Berlin summit · Uganda organisation · self-funded travel",
        "ask": "Award complimentary group tickets (3–15) for FairBanks youth delegates.",
    },
    exec_summary=[
        "The World Health Summit offers complimentary group tickets for youth-led organisations working in health and social impact.",
        "FairBanks youth can bring an authentic African community-health intelligence story to Berlin — and return with partnerships that strengthen FCHIP.",
        "WHS gains underrepresented regional voices; FairBanks gains global learning without ticket barriers (travel remains self-funded).",
    ],
    win_for_us=[
        "Access to three days of global health agenda-setting in Berlin",
        "Networking with ministries, NGOs, and innovators",
        "Visibility for FCHIP and FairBanks Community Reach",
        "Learning that feeds Uganda programme design",
    ],
    win_for_them=[
        "Youth delegates from an underrepresented region with real field work",
        "Concrete community health intelligence narrative for summit sessions",
        "Organisation that will share learnings back home, not only attend",
        "Diverse group size (3–15) ready to engage seriously",
    ],
    problem=[
        "Global health forums can miss the voices of community-level African implementers — especially youth who already do the work.",
        "Ticket costs block participation even when travel can be self-organised.",
    ],
    solution=[
        "A FairBanks youth group will attend WHS 2026, document key sessions, and run a post-summit learning clinic for CHWs and clinic staff in Kampala.",
        "We will focus on digital health, community systems, and partnership pathways relevant to FCHIP.",
    ],
    use_cases=[
        ("Summit learning", "Capture transferable practices for CHW digital tools."),
        ("Partnership mapping", "Identify research and NGO collaborators."),
        ("Homeward briefing", "Share insights with FairBanks and community partners."),
    ],
    fit_rows=[
        ["Youth-led health / impact org", "FairBanks youth cohort within community health ecosystem"],
        ["Nominees aged 28 or under", "Delegation will meet age rules"],
        ["Group tickets 3–15", "We request a mid-size learning group"],
        ["Self-funded travel", "Understood and planned separately from ticket request"],
    ],
    plan_bullets=[
        "Select diverse youth nominees (clinical, community, digital) under 28",
        "Prepare a one-page FairBanks brief for networking",
        "Attend all three summit days with a session tracking plan",
        "Host a post-summit community briefing within 30 days of return",
    ],
    plan_slide_title="Delegation and learning plan",
    closing=[
        "WHS is stronger when community practitioners are in the room. FairBanks youth are ready to listen, contribute, and bring lessons home.",
        "We request complimentary group tickets for our eligible delegation.",
    ],
)

OYW = _base(
    slug="oyw",
    url="https://opportunitiesforyouth.org/2026/07/15/ai-for-good-scholarship-2025-empowering-young-innovators-to-address-global-challenges/",
    meta={
        "programme": "Brandtech Group AI for Good Scholarship 2026 (One Young World)",
        "title": "AI for Community Health — Cape Town Summit",
        "subtitle": "Fully funded seat for a young African leader using AI for health impact",
        "applicant": "Eligible FairBanks young leader (18–30), Ugandan national & resident",
        "deadline": "2 August 2026",
        "geography": "One Young World Summit 2026 — Cape Town",
        "ask": "Award one of five AI for Good scholarships to a FairBanks young leader.",
    },
    exec_summary=[
        "Brandtech's AI for Good scholarship sends five young African leaders using AI for social impact to One Young World 2026 in Cape Town. Healthcare is a named priority.",
        "A FairBanks nominee can show responsible AI for community health — CHW data, maternal risk, and outbreak early warning — rooted in Uganda.",
        "The scholarship amplifies an African HealthTech practitioner; FairBanks returns with global alliances for FCHIP.",
    ],
    win_for_us=[
        "Fully funded summit access for a young AI-for-health leader",
        "Global network across One Young World delegates",
        "Platform to share FCHIP's responsible AI story",
        "Inspiration and partnerships for product and policy pathways",
    ],
    win_for_them=[
        "A scholar with proven community health AI intent, not abstract interest",
        "Healthcare priority area represented with African field credibility",
        "Nationality and residency eligibility met (Uganda)",
        "Alumni who will keep building after the summit lights fade",
    ],
    problem=[
        "AI for good needs builders who already serve communities — not only pitch decks. Too few young African health practitioners get summit access.",
    ],
    solution=[
        "Our nominee will present FCHIP's approach: community-generated data + predictive analytics for prevention, with ethics and gender awareness at the centre.",
        "Post-summit, learnings feed FairBanks product and partnership plans.",
    ],
    use_cases=[
        ("Responsible AI", "Explainable alerts for CHWs and clinicians."),
        ("Health for all", "Prioritise underserved communities in model design."),
        ("Youth leadership", "Carry summit commitments into Uganda programmes."),
    ],
    fit_rows=[
        ["Ages 18–30", "Nominee within range"],
        ["African nationality & residency", "Ugandan national living in Uganda"],
        ["AI for social / health impact", "FCHIP predictive community health"],
        ["Five scholarships", "Strong healthcare priority alignment"],
    ],
    plan_bullets=[
        "Submit nominee profile with AI-for-health evidence and references",
        "Prepare summit talking points on FCHIP and FairBanks cascade",
        "Engage Brandtech / OYW sessions on responsible AI",
        "Publish a public learning note for FairBanks community partners",
    ],
    plan_slide_title="Scholarship journey",
    closing=[
        "AI for Good is measured by lives improved. FairBanks is already walking that path in Uganda.",
        "We ask for a Brandtech AI for Good scholarship for our eligible young leader.",
    ],
)

AFRICA_CDC = _base(
    slug="africa-cdc",
    url="https://opportunitiesforyouth.org/2026/07/16/africa-cdc-african-epidemic-services-aes-epidemiology-fellowship-2026-2028-fully-funded-apply-by-26-august-2026/",
    meta={
        "programme": "Africa CDC African Epidemic Services (AES) Epidemiology Fellowship 2026–2028",
        "title": "Epidemiology & Informatics Talent for FCHIP",
        "subtitle": "Fully funded fellowship strengthening outbreak intelligence capacity",
        "applicant": "Eligible FairBanks public-health staff (AU citizen, under 35, employed)",
        "deadline": "26 August 2026",
        "geography": "Addis Ababa training + field placement · Uganda employer",
        "ask": "Select a FairBanks public-health professional for the AES Epidemiology (or Informatics) track.",
    },
    exec_summary=[
        "Africa CDC's AES fellowship builds outbreak investigation, surveillance, and data-analysis skills over two years — including a Public Health Informatics pathway.",
        "FairBanks needs this depth to make FCHIP's surveillance claims scientifically sound and operationally useful.",
        "Africa CDC gains a fellow employed in a live community health setting who will apply skills immediately.",
    ],
    win_for_us=[
        "World-class epidemiology / informatics training for staff",
        "Stronger outbreak and surveillance modules inside FCHIP",
        "Stipend, travel, insurance, hardware, and software covered",
        "Continental network across African Epidemic Services",
    ],
    win_for_them=[
        "Fellow already employed in public-facing community health work",
        "Field placement potential linked to real CHW and facility data flows",
        "Contribution to Africa's epidemic readiness from Uganda",
        "Informatics track alignment with digital community health tools",
    ],
    problem=[
        "Predictive community platforms fail without trained people who understand surveillance science, not only software.",
        "Africa needs more young epidemiologists and informaticians grounded in primary care realities.",
    ],
    solution=[
        "A FairBanks fellow will complete AES training and bring methods home: case investigation discipline, surveillance analysis, and informatics design for FCHIP alerts.",
        "Employer support ensures learning returns to community programmes and partner districts.",
    ],
    use_cases=[
        ("Outbreak workflows", "Standardise community signal → investigation → response."),
        ("Informatics", "Improve data quality pipelines for CHW capture."),
        ("Training cascade", "Teach CHW supervisors better data use."),
    ],
    fit_rows=[
        ["AU citizen under 35", "Ugandan staff candidate in range"],
        ["Employed in public health", "FairBanks clinical / community health employment"],
        ["Degree + 3 years' experience", "Candidate will meet stated requirements"],
        ["Epidemiology / Informatics tracks", "Direct fit to FCHIP surveillance goals"],
    ],
    plan_rows=[
        ["Months 1–3", "Core training in Addis Ababa"],
        ["Months 4–24", "Field placement with FairBanks / partner public health sites"],
        ["Throughout", "Apply methods to FCHIP surveillance and quality improvement"],
    ],
    plan_slide_title="Two-year fellowship pathway",
    closing=[
        "Epidemics are fought with skilled people. FairBanks will back our fellow completely so Africa CDC's investment returns to communities.",
        "We ask for selection into the 2026–2028 AES cohort.",
    ],
)

IFAD = _base(
    slug="ifad",
    url="https://opportunitiesforyouth.org/2026/07/14/ifad-announces-us1-2-million-grant-call-for-proposals-to-strengthen-rural-development-policies-worldwide/",
    meta={
        "programme": "IFAD Grant — Rural Sector Performance Assessment (up to US$1.2M)",
        "title": "Health–Livelihood Intelligence for Rural Policy",
        "subtitle": "Linking community health data to rural development performance",
        "applicant": "FairBanks (with research / policy partners as required) — Uganda",
        "deadline": "4 September 2026 (12:00 CET); optional EOI 31 July 2026",
        "geography": "Uganda rural / peri-urban catchments · policy learning global",
        "ask": "Consider FairBanks (and partners) for EOI and full proposal on health–livelihood policy intelligence.",
    },
    exec_summary=[
        "IFAD seeks proposals up to US$1.2M to leverage the Rural Sector Performance Assessment for better policy and investment — open to non-profits, private firms, universities, and think tanks.",
        "FairBanks' cascade already links health to livelihoods and resilient families. FCHIP can contribute community health and vulnerability signals that make rural performance assessment more human-centred.",
        "This pack frames a careful, partnership-ready fit — to be confirmed against IFAD's detailed thematic guidance before full proposal effort.",
    ],
    win_for_us=[
        "Resources to connect health intelligence with rural livelihood policy",
        "Partnerships with universities / think tanks on rigorous assessment",
        "Evidence that prevention protects income and food security",
        "Optional EOI path to test fit early (31 July 2026)",
    ],
    win_for_them=[
        "Ground-level health and vulnerability data from community programmes",
        "Private-sector / social enterprise implementer with field access",
        "SDG-aligned story: health shocks as rural poverty drivers",
        "Counterpart contribution potential via FairBanks operations (20–25%)",
    ],
    problem=[
        "Rural policy often tracks production and infrastructure while under-weighting health shocks that wipe out household gains.",
        "Without community health intelligence, rural performance assessments miss a core driver of poverty.",
    ],
    solution=[
        "Propose a workstream where FCHIP community indicators (illness clusters, maternal risk, missed care) inform rural performance dashboards and policy briefs, in partnership with research institutes.",
        "Counterpart effort includes FairBanks field operations, CHW networks, and data stewardship.",
    ],
    use_cases=[
        ("Health-shock markers", "Link disease surges to livelihood stress signals."),
        ("Policy briefs", "Translate community data into investment priorities."),
        ("Partner research", "Co-design methods with academic / think-tank partners."),
    ],
    fit_rows=[
        ["Rural Sector Performance Assessment", "Add health–livelihood intelligence layer"],
        ["Eligible entity types", "Social enterprise + research partners"],
        ["Counterpart 20–25%", "In-kind operations and staff time"],
        ["Verify thematic match", "EOI first; full proposal only if guidance confirms fit"],
    ],
    plan_rows=[
        ["By 31 Jul 2026", "Optional non-binding Expression of Interest"],
        ["By 4 Sep 2026", "Full proposal if thematic fit confirmed"],
        ["Year 1–2", "Pilot indicators in FairBanks / partner rural catchments"],
        ["Learning", "Policy notes for IFAD and national stakeholders"],
    ],
    budget_rows=[
        ["Research & assessment design", "Methods, partners, ethics", "Share of grant"],
        ["Community data & FCHIP modules", "Indicators, dashboards, quality", "Share of grant"],
        ["FairBanks counterpart", "Field ops, staff, outreach", "20–25% in-kind"],
        ["Ceiling", "Per call", "Up to US$1.2M"],
    ],
    plan_slide_title="EOI → proposal → policy learning",
    closing=[
        "Rural development succeeds when families stay healthy enough to work and thrive. FairBanks can help IFAD see that link clearly.",
        "We will submit an EOI and proceed to full proposal only where official guidance confirms thematic fit.",
    ],
)

GOVTECH = _base(
    slug="govtech",
    url="https://opportunitiesforyouth.org/2026/07/10/world-bank-govtech-innovation-challenge-2026-global-call-for-proposals-to-build-digital-solutions-for-governments/",
    meta={
        "programme": "World Bank GovTech Innovation Challenge 2026",
        "title": "District Health Intelligence for Government Partners",
        "subtitle": "Watching for health / Africa challenges — ready proof-of-concept partner",
        "applicant": "FairBanks / FCHIP — technology company & health operator, Uganda",
        "deadline": "Varies by challenge — monitor specific challenge pages",
        "geography": "Global challenge platform · Uganda / Africa health use cases",
        "ask": "Shortlist FairBanks when health or Africa digital-government challenges open; engage on relevant PoCs.",
    },
    exec_summary=[
        "The World Bank GovTech Innovation Challenge connects technology companies with government digitalisation needs — offering capacity building, a Switzerland bootcamp, and proof-of-concept work with government partners.",
        "Current open challenges may focus elsewhere (e.g. public-finance audit), but FairBanks maintains a ready response pack for health and Africa-aligned calls.",
        "When a matching challenge appears, we can move fast with a district health intelligence PoC concept rooted in live operations.",
    ],
    win_for_us=[
        "Access to government PoC pathways and Trust Valley capacity building",
        "Credibility from World Bank / SECO ecosystem visibility",
        "Learning that hardens FCHIP for public-sector deployment",
        "Potential Switzerland bootcamp for selected innovators",
    ],
    win_for_them=[
        "A vendor that already runs clinic and community operations — not only software",
        "Health intelligence use cases governments actually need (surveillance, CHW data, dashboards)",
        "African implementer ready for public-sector co-design",
        "Reusable PoC patterns for district health offices",
    ],
    problem=[
        "Governments struggle to turn scattered community health data into timely decisions. Vendors without field roots deliver brittle pilots.",
    ],
    solution=[
        "FCHIP offers a GovTech-ready pattern: secure community capture, analytics, GIS, and role-based dashboards co-designed with district users.",
        "We monitor challenge pages and submit only when the use case matches — protecting evaluator time and our own.",
    ],
    use_cases=[
        ("District dashboard", "Population risk and coverage views for managers."),
        ("CHW data quality", "Validation that improves government reporting."),
        ("Early warning", "Configurable alerts aligned to local protocols."),
    ],
    fit_rows=[
        ["Tech companies worldwide", "Ugandan HealthTech operator eligible"],
        ["Government digitalisation", "District health intelligence PoC ready"],
        ["Capacity building + bootcamp", "Team prepared to attend if selected"],
        ["Challenge-specific deadlines", "Active monitoring; no blind submissions"],
    ],
    plan_bullets=[
        "Maintain this application pack as a ready response baseline",
        "Monitor World Bank GovTech challenge pages weekly",
        "Submit tailored proposals only for health / Africa-aligned challenges",
        "If selected: co-design PoC with government partner and measure decision-cycle gains",
    ],
    plan_slide_title="Monitor → match → PoC",
    closing=[
        "GovTech works when government problems meet operators who already serve people. FairBanks is ready when the right challenge opens.",
        "We ask to be considered for matching health and Africa digital-government challenges as they appear.",
    ],
)

# AWIEF already has a full custom pack under applications/awief/
SPECS = {
    "auc": AUC,
    "feminist-ai": FEMINIST_AI,
    "girlcode": GIRLCODE,
    "auda-srh": AUDA_SRH,
    "gadfly": GADFLY,
    "dos-uganda": DOS_UGANDA,
    "whs": WHS,
    "oyw": OYW,
    "africa-cdc": AFRICA_CDC,
    "ifad": IFAD,
    "govtech": GOVTECH,
}
