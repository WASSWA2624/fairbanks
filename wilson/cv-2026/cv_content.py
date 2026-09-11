# -*- coding: utf-8 -*-
"""
Single source of truth for the Wasswa Wilson general-purpose CV.

Both the DOCX and the PDF renderers read from this module, so the two
documents can never drift apart. Edit here, then re-run build_wilson_cv.py.

House rule: no em dashes anywhere in this file. Use commas, colons,
parentheses, full stops or the middot separator instead.
"""

# --------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------

NAME = "WASSWA WILSON"

TAGLINE = "Biomedical Engineer  ·  Full-Stack Software Developer  ·  Health Technology Programme Lead"

CONTACT_LINES = [
    "Kampala, Uganda   ·   wasswawilson0001@gmail.com   ·   +256 783 230 321",
    "English (fluent, professional working proficiency)   ·   Luganda (native)",
]

# --------------------------------------------------------------------------
# Profile
# --------------------------------------------------------------------------

PROFILE = (
    "I am a biomedical engineer and full-stack software developer based in Kampala, with over "
    "nine years of work across hospital technology, health information systems and software "
    "engineering. I lead development of the FairBanks Community Health Intelligence Platform "
    "(FCHIP), where I handle the data capture tools, the database, the cloud data flows and the "
    "management dashboards. Alongside it I am building a cross-sector application that collects "
    "field data and generates finished reports in real time, for accounting, research, biology, "
    "agriculture, social media and other record-heavy work. Before FairBanks I managed the Gould "
    "Family Foundation biomedical programme across East and Central Africa, and ran the biomedical "
    "function at International Hospital Kampala through the COVID-19 response and through "
    "successful COHSASA accreditation. I write Java, Python, JavaScript and TypeScript, work "
    "across React, Next.js, React Native, Node.js and Express, and design AI agents and "
    "automations that are running in production today."
)

# --------------------------------------------------------------------------
# Career highlights
# --------------------------------------------------------------------------

HIGHLIGHTS = [
    "Lead developer of FCHIP, the FairBanks Community Health Intelligence Platform: community "
    "screening, referral and follow-up data capture, database, cloud sync and reporting "
    "dashboards, delivered across web and mobile.",

    "Building a cross-sector data collection and reporting application that writes the report as "
    "the user works, instead of leaving it to be compiled afterwards. Aimed at accounting, "
    "research, biology, agriculture, social media and similar fields.",

    "Developed the oxygen consumption formulas used to calculate patient demand during the "
    "COVID-19 response at International Hospital Kampala, and in support of Ministry of Health "
    "oxygen planning. The same method fed into capacity planning for later facility projects.",

    "Designed, built and deployed an AI reception agent for a law firm in Texas, USA. It handles "
    "live inbound client calls, books and reschedules appointments, and runs client follow-up. In "
    "active production use.",

    "Delivered hospital technology projects in five countries: Uganda, the Democratic Republic of "
    "Congo, Kenya, Tanzania and Somaliland. This included laboratory, theatre and radiology "
    "installations at the hospital serving the Kibali Gold Mine, DR Congo.",

    "Prepared and maintained the biomedical compliance evidence that carried International "
    "Hospital Kampala through successful COHSASA accreditation, and cut equipment downtime with a "
    "documented preventive maintenance programme.",

    "Upgraded partner health facilities across East and Central Africa for the Gould Family "
    "Foundation, including the neonatal intensive care unit at Mama Lucy Kibaki Hospital, Nairobi.",
]

# --------------------------------------------------------------------------
# Technical skills
# --------------------------------------------------------------------------

SKILLS = [
    ("Programming languages",
     "Java  ·  Python  ·  JavaScript  ·  TypeScript  ·  C  ·  C++  ·  C#  ·  PHP  ·  SQL  ·  "
     "MATLAB  ·  Arduino C  ·  HTML5  ·  CSS3"),

    ("Web and mobile",
     "React  ·  Next.js (App Router)  ·  React Native  ·  Expo  ·  Tailwind CSS  ·  "
     "styled-components  ·  Bootstrap  ·  JavaFX / FXML  ·  Jest  ·  Playwright"),

    ("Backend and APIs",
     "Node.js  ·  Express  ·  Laravel  ·  CodeIgniter  ·  Spring / Spring Boot  ·  Convex  ·  "
     "Prisma  ·  REST APIs  ·  JWT  ·  role-based access control"),

    ("Databases and data",
     "MySQL  ·  relational modelling and query optimisation  ·  SQL analytics  ·  "
     "data cleaning and reporting  ·  MATLAB modelling  ·  Excel analysis and dashboards"),

    ("Cloud and DevOps",
     "AWS  ·  Microsoft Azure  ·  Vercel  ·  EAS  ·  Bunny.net CDN  ·  Git and GitHub  ·  "
     "CI/CD pipelines  ·  release and environment management"),

    ("AI, agents and automation",
     "AI agent design, deployment and monitoring  ·  workflow and process automation  ·  "
     "third-party systems integration  ·  automated auditing and verification tooling  ·  "
     "Cursor  ·  Claude  ·  OpenAI GPT  ·  Codex"),

    ("Operating systems",
     "Windows  ·  Linux  ·  Android.  Administration, deployment and troubleshooting on all three."),

    ("Engineering and scientific",
     "Embedded systems design  ·  biosignal processing  ·  medical imaging  ·  "
     "biomedical modelling and simulation  ·  bioinformatics and functional genomics  ·  "
     "Proteus  ·  Eagle CAD  ·  Solid Edge"),

    ("Documents and productivity",
     "Microsoft Word  ·  Excel  ·  PowerPoint  ·  technical writing  ·  proposal and grant "
     "development  ·  document editing and design  ·  report and dashboard production"),
]

# --------------------------------------------------------------------------
# Professional experience
# --------------------------------------------------------------------------

EXPERIENCE = [
    {
        "role": "Lead Software Developer, FCHIP and Health Information Systems",
        "org": "FairBanks Medical Centre Limited, Kampala, Uganda",
        "dates": "2026 to present",
        "bullets": [
            "Lead the design and build of the FairBanks Community Health Intelligence Platform "
            "(FCHIP): data capture form design, database structure, cloud data flows, reporting "
            "dashboards, and delivery across both web and mobile.",

            "Building a cross-sector data collection and reporting application that turns work "
            "already being done into finished reports in real time, for accounting, research and "
            "laboratory science, biology, agriculture, social media and similar record-heavy "
            "fields.",

            "Architect and maintain the hospital and clinic software used daily by clinical and "
            "administrative staff, covering patient records, scheduling, referral tracking and "
            "management reporting.",

            "Design and deploy AI agents and automation pipelines that remove manual handling from "
            "reception, appointment management, client follow-up and routine reporting.",

            "Own technical documentation, data-protection practice, access control and release "
            "management across the platform.",

            "Write and edit institutional proposals, grant applications, partnership documents and "
            "investment materials. Chaired the FairBanks meeting with the Uganda Health Partners "
            "Cooperative on cooperative health insurance in July 2026.",
        ],
    },
    {
        "role": "Consultant Biomedical Engineer, Independent Practice",
        "org": "Contract assignments: Uganda, DR Congo, Kenya, Tanzania, Somaliland",
        "dates": "Mar 2025 to present",
        "bullets": [
            "Install, commission, calibrate and maintain hospital medical equipment under contract "
            "across five countries, working independently on site and to a fixed handover standard.",

            "Completed laboratory, theatre and radiology installations at several hospitals in the "
            "Democratic Republic of Congo, including the facility serving the Kibali Gold Mine.",

            "Train clinical users at handover and produce a written service report and maintenance "
            "record for every assignment, so the client keeps a verifiable equipment history.",

            "Handle preventive and corrective maintenance, on-site fault diagnosis and "
            "specification advice for procurement.",
        ],
    },
    {
        "role": "Biomedical Programs Manager",
        "org": "Gould Family Foundation (GFF), East and Central Africa",
        "dates": "Aug 2024 to Feb 2025",
        "bullets": [
            "Managed the regional biomedical team, covering programme planning, implementation and "
            "reporting across partner health facilities in multiple countries.",

            "Oversaw procurement, installation and commissioning of medical technologies, with "
            "programme documentation kept in line with healthcare technology standards.",

            "Developed equipment lifecycle-management and master-planning strategies that were "
            "applied across partner facilities.",

            "Upgraded several facilities, including the neonatal intensive care unit at Mama Lucy "
            "Kibaki Hospital, Nairobi, where oxygen and equipment demand modelling informed the "
            "capacity plan.",

            "Led biomedical and clinical user training programmes for engineers, technicians and "
            "clinical staff.",
        ],
    },
    {
        "role": "Biomedical Manager",
        "org": "International Hospital Kampala (IHK)",
        "dates": "Jan 2020 to Jan 2024",
        "bullets": [
            "Managed the biomedical engineering function of a major private hospital for four "
            "years, including team supervision, workplans and reporting to executive management.",

            "Led oxygen therapy management through the COVID-19 response and supported Ministry of "
            "Health oxygen planning. Developed the formulas used to calculate patient oxygen "
            "consumption, which gave managers a dependable basis for sizing plant capacity, "
            "cylinder stock and flow requirements rather than estimating them.",

            "Led installation and commissioning of CT and X-ray imaging, oxygen plant, laboratory, "
            "theatre and ICU systems, through to handover and trained clinical users.",

            "Built the preventive maintenance schedules and service-record systems that reduced "
            "equipment downtime and kept devices safe for clinical use.",

            "Coordinated procurement, technical specification, commissioning documentation and "
            "lifecycle management of hospital medical equipment.",

            "Prepared and maintained the compliance evidence files reviewed for COHSASA "
            "accreditation, supporting the hospital through successful accreditation.",
        ],
    },
    {
        "role": "Biomedical Engineer",
        "org": "Norvik Hospital Ltd, Kampala",
        "dates": "Apr 2019 to Jan 2020",
        "bullets": [
            "Maintained and serviced hospital diagnostic and patient-monitoring equipment.",
            "Supported installation and calibration of medical imaging and ICU equipment.",
            "Helped establish preventive maintenance schedules and service follow-up records.",
        ],
    },
    {
        "role": "Research Intern",
        "org": "Uganda Virus Research Institute (UVRI)",
        "dates": "Dec 2018 to Apr 2019",
        "bullets": [
            "Contributed to biomedical research projects and laboratory experimentation.",
            "Carried out laboratory data analysis, mathematical modelling and research pipeline "
            "design in Python.",
        ],
    },
    {
        "role": "Teaching Assistant",
        "org": "Makerere University, College of Health Sciences",
        "dates": "Mar 2016 to Aug 2017",
        "bullets": [
            "Assisted in teaching biomedical engineering course units at undergraduate level.",
            "Supervised and supported laboratory practical sessions for engineering students.",
        ],
    },
]

# --------------------------------------------------------------------------
# Selected projects  (project, role, context, period, detail)
# --------------------------------------------------------------------------

PROJECTS = [
    ("FairBanks Community Health Intelligence Platform (FCHIP)",
     "Lead Software Developer",
     "FairBanks Medical Centre Limited",
     "2026 to present",
     "Community, screening, referral and follow-up data capture; database structure; cloud data "
     "flows; management dashboards; web and mobile delivery. Built with Next.js, TypeScript, "
     "React Native and cloud data services."),

    ("Cross-sector data collection and reporting application",
     "Founder and Lead Developer",
     "In development",
     "2026 to present",
     "An application that simplifies data collection and reporting for professions that live on "
     "records: accounting, research, biology and laboratory science, agriculture and social media "
     "among them. Rather than collecting data now and writing the report later, the user works "
     "inside the app and the information and reports are generated in real time as they go."),

    ("COVID-19 oxygen therapy management and demand modelling",
     "Biomedical Manager",
     "International Hospital Kampala, with Ministry of Health planning support",
     "2020 to 2022",
     "Ran oxygen therapy management for the hospital through the pandemic and developed the "
     "formulas for calculating patient oxygen consumption. The figures were used to size plant "
     "capacity, cylinder stock and flow requirements, and to support planning decisions at "
     "managerial and national level."),

    ("AI reception agent",
     "Designer and Developer",
     "Law firm, Texas, USA",
     "2025 to present",
     "Built and deployed an AI agent that answers live inbound client calls, books and reschedules "
     "appointments, and runs post-contact client follow-up. In active production use."),

    ("Shulekeeper school information system",
     "Assistant Software Developer",
     "Shulekeeper",
     "2021 to 2024",
     "Multi-user records and reporting platform built with React and Next.js, TypeScript, Laravel, "
     "MySQL and AWS. Contributed front-end and back-end work, third-party integrations, testing, "
     "user training and documentation."),

    ("Endoscopy image capture and reporting software",
     "Lead Software Developer",
     "St. Catherine Hospital and St. Francis Hospital",
     "2017 to 2019",
     "Clinical image capture and procedure reporting system in JavaFX and MySQL. Led algorithm "
     "design, relational modelling, legacy data migration, on-site installation, user training and "
     "the user manual."),

    ("Wekebere foetal heart-rate monitoring system",
     "Electronics Developer and Programmer",
     "Wekebere",
     "2017 to 2024",
     "Android and Arduino monitoring system for third-trimester foetal heart-rate measurement, "
     "with Java serial communication between device and phone and Microsoft Azure cloud storage."),
]

# --------------------------------------------------------------------------
# Professional competencies
# --------------------------------------------------------------------------

COMPETENCIES = [
    ("Programme and project management",
     "Managed a regional biomedical team across East and Central Africa and a hospital biomedical "
     "function for four years: workplans, schedules, procurement coordination and management "
     "reporting."),

    ("Monitoring, evaluation and data",
     "Leads the FCHIP indicator, data-quality, dashboard and reporting build. Built equipment "
     "registers, maintenance schedules and routine management reports. Analysis in Python, SQL, "
     "Java, JavaScript and MATLAB."),

    ("Proposal, grant and technical writing",
     "Writes and edits institutional proposals, grant applications, investment materials, "
     "partnership documents, technical documentation and user manuals."),

    ("Data analysis and automated auditing",
     "Cleans, models and reports on operational and clinical data. Builds automated auditing and "
     "verification tooling to check records in bulk rather than by hand."),

    ("Procurement, specification and compliance",
     "Coordinated specification, procurement, commissioning and lifecycle documentation, and kept "
     "the compliance evidence reviewed for COHSASA accreditation."),

    ("Stakeholder and partnership management",
     "Chairs partner meetings and writes them up. Works with hospital departments, suppliers, "
     "trainees, accreditation assessors, donors and prospective partners."),

    ("Training and capacity building",
     "Trained biomedical engineers, technicians and clinical users across multiple facilities, and "
     "taught undergraduate engineering laboratory practicals at Makerere University."),

    ("Communication and adaptability",
     "Clear written and spoken English, comfortable presenting to clinical, executive and donor "
     "audiences. Settles quickly into new countries, teams, codebases and equipment brands."),
]

# --------------------------------------------------------------------------
# Medical equipment experience
# --------------------------------------------------------------------------

EQUIPMENT = [
    ("Radiology and imaging",
     "CT scanners  ·  X-ray machines  ·  ultrasound systems  ·  C-arms"),
    ("Theatre and intensive care",
     "Theatre systems  ·  ventilators  ·  patient monitors  ·  infusion pumps  ·  "
     "critical-care devices"),
    ("Oxygen and medical gas",
     "Oxygen plants  ·  concentrators  ·  cylinder and manifold systems  ·  piped medical gas  ·  "
     "oxygen demand and capacity calculation"),
    ("Laboratory",
     "Haematology analysers  ·  clinical chemistry analysers  ·  blood bank and transfusion "
     "equipment"),
    ("Maternity and neonatal",
     "Neonatal incubators  ·  phototherapy units  ·  foetal Dopplers  ·  CTG machines"),
    ("Support plant",
     "RO water treatment systems  ·  sterilisation and autoclave equipment"),
]

# --------------------------------------------------------------------------
# Education & professional development
# --------------------------------------------------------------------------

EDUCATION = [
    ("Certificate in Leadership and Management in Health",
     "University of Washington",
     "2022",
     "Grade A+. Planning, leadership and management of health programmes."),

    ("BSc Biomedical Engineering, Second Upper Honours",
     "Makerere University, Kampala",
     "2012 to 2017",
     "Key course units: Information and Communication Technology, Computer OOP Principles, "
     "Software Engineering, Database Systems, Computer-Based Medical Systems, Embedded Systems, "
     "Biosignal Processing and Analysis, Medical Imaging, Biomedical Modelling and Simulation, "
     "Bioinformatics and Functional Genomics, Clinical Engineering, Research Methods, "
     "Principles of Management."),

    ("Advanced Java Programming, Level 1",
     "Green Bridge School of Open Technologies",
     "2015 to 2016",
     "Grade A+."),

    ("Uganda Advanced Certificate of Education (UACE)",
     "Mengo Secondary School",
     "2011",
     "Principal subjects: Mathematics, Physics, Biology, Chemistry."),

    ("Uganda Certificate of Education (UCE)",
     "St. John's Wakiso Secondary School",
     "2009",
     ""),
]

TRAINING = [
    ("COHSASA accreditation standards: in-service preparation and compliance",
     "International Hospital Kampala", "2020 to 2024"),
    ("Fire Safety, Prevention, Firefighting and Emergency Scene Management",
     "Fire Technologies Limited (IHK)", "2021"),
]

# --------------------------------------------------------------------------
# Interests
# --------------------------------------------------------------------------

INTERESTS = (
    "Music, and playing the saxophone in particular.  ·  Film.  ·  Physical fitness and regular "
    "exercise.  ·  Travel and working in unfamiliar places, which the regional field work has "
    "made second nature."
)

# --------------------------------------------------------------------------
# Referees
# --------------------------------------------------------------------------

REFEREES = [
    ("Eng. Richard Ssejongo",
     "Biomedical Engineer, St. Francis Hospital Nsambya",
     "+256 753 818 754  ·  +256 777 132 489"),

    ("Dr. Annet Khingi",
     "Senior Radiologist and Administrator, Mengo Hospital",
     "+256 772 592 771  ·  +256 701 592 771"),

    ("Racheal Nabukeera",
     "Director and Founder, FairBanks Medical Centre",
     "+256 772 849 258  ·  nracheal017@gmail.com"),
]
