# -*- coding: utf-8 -*-
"""
Single source of truth for the Wasswa Wilson general-purpose CV.

Both the DOCX and the PDF renderers read from this module, so the two
documents can never drift apart. Edit here, then re-run build_wilson_cv.py.

House rules
-----------
* No em dashes. Use commas, colons, parentheses, full stops or the middot.
* First person throughout. "I led", "I built", "I contributed", not
  "Led" or "Responsible for". The reader should never be in doubt about
  which parts of this were mine.
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
    "Family Foundation biomedical programme across East and Central Africa, and I ran the "
    "biomedical function at International Hospital Kampala through the COVID-19 response and "
    "through successful COHSASA accreditation. I write Java, Python, JavaScript and TypeScript, "
    "I work across React, Next.js, React Native, Node.js and Express, and I design AI agents and "
    "automations that are running in production today."
)

# --------------------------------------------------------------------------
# Career highlights
# --------------------------------------------------------------------------

HIGHLIGHTS = [
    "I lead development of FCHIP, the FairBanks Community Health Intelligence Platform. I built "
    "the community screening, referral and follow-up data capture, the database, the cloud sync "
    "and the reporting dashboards, and I deliver the platform across web and mobile.",

    "I am building a cross-sector data collection and reporting application that writes the "
    "report as the user works, instead of leaving it to be compiled afterwards. I designed it for "
    "accounting, research, biology, agriculture, social media and similar fields.",

    "I developed the oxygen consumption formulas used to calculate patient demand during the "
    "COVID-19 response at International Hospital Kampala, and I supported Ministry of Health "
    "oxygen planning. My method went on to inform capacity planning for later facility projects.",

    "I designed, built and deployed an AI reception agent for a law firm in Texas, USA. It "
    "handles live inbound client calls, books and reschedules appointments, and runs client "
    "follow-up. It is in active production use.",

    "I have delivered hospital technology projects in five countries: Uganda, the Democratic "
    "Republic of Congo, Kenya, Tanzania and Somaliland. I completed the laboratory, theatre and "
    "radiology installations at the hospital serving the Kibali Gold Mine, DR Congo.",

    "I prepared and maintained the biomedical compliance evidence that carried International "
    "Hospital Kampala through successful COHSASA accreditation, and I cut equipment downtime with "
    "a preventive maintenance programme that I designed, documented and ran.",

    "I upgraded partner health facilities across East and Central Africa for the Gould Family "
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
     "Windows  ·  Linux  ·  Android.  I administer, deploy and troubleshoot on all three."),

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
            "I lead the design and build of the FairBanks Community Health Intelligence Platform "
            "(FCHIP). I do the data capture form design, the database structure, the cloud data "
            "flows, the reporting dashboards and the delivery across both web and mobile.",

            "I am building a cross-sector data collection and reporting application that turns "
            "work already being done into finished reports in real time, for accounting, research "
            "and laboratory science, biology, agriculture, social media and similar record-heavy "
            "fields.",

            "I architect and maintain the hospital and clinic software that clinical and "
            "administrative staff use daily, covering patient records, scheduling, referral "
            "tracking and management reporting.",

            "I design and deploy the AI agents and automation pipelines that remove manual "
            "handling from reception, appointment management, client follow-up and routine "
            "reporting.",

            "I own the technical documentation, the data-protection practice, the access control "
            "and the release management across the platform.",

            "I write and edit the institutional proposals, grant applications, partnership "
            "documents and investment materials, and I chaired the FairBanks meeting with the "
            "Uganda Health Partners Cooperative on cooperative health insurance in July 2026.",
        ],
    },
    {
        "role": "Consultant Biomedical Engineer, Independent Practice",
        "org": "Contract assignments: Uganda, DR Congo, Kenya, Tanzania, Somaliland",
        "dates": "Mar 2025 to present",
        "bullets": [
            "I install, commission, calibrate and maintain hospital medical equipment under "
            "contract across five countries, working independently on site and to a fixed "
            "handover standard.",

            "I completed the laboratory, theatre and radiology installations at several hospitals "
            "in the Democratic Republic of Congo, including the facility serving the Kibali Gold "
            "Mine.",

            "I train the clinical users at handover, and I produce a written service report and "
            "maintenance record for every assignment, so the client keeps a verifiable equipment "
            "history.",

            "I handle the preventive and corrective maintenance, the on-site fault diagnosis and "
            "the specification advice for procurement.",
        ],
    },
    {
        "role": "Biomedical Programs Manager",
        "org": "Gould Family Foundation (GFF), East and Central Africa",
        "dates": "Aug 2024 to Feb 2025",
        "bullets": [
            "I managed the regional biomedical team, covering programme planning, implementation "
            "and reporting across partner health facilities in multiple countries.",

            "I oversaw the procurement, installation and commissioning of medical technologies, "
            "and I kept the programme documentation in line with healthcare technology standards.",

            "I developed the equipment lifecycle-management and master-planning strategies that "
            "were then applied across partner facilities.",

            "I upgraded several facilities, including the neonatal intensive care unit at Mama "
            "Lucy Kibaki Hospital, Nairobi, where my oxygen and equipment demand modelling "
            "informed the capacity plan.",

            "I led the biomedical and clinical user training programmes for engineers, "
            "technicians and clinical staff.",
        ],
    },
    {
        "role": "Biomedical Manager",
        "org": "International Hospital Kampala (IHK)",
        "dates": "Jan 2020 to Jan 2024",
        "bullets": [
            "I managed the biomedical engineering function of a major private hospital for four "
            "years, including team supervision, workplans and reporting to executive management.",

            "I led oxygen therapy management through the COVID-19 response and I supported "
            "Ministry of Health oxygen planning. I developed the formulas used to calculate "
            "patient oxygen consumption, which gave managers a dependable basis for sizing plant "
            "capacity, cylinder stock and flow requirements rather than estimating them.",

            "I led the installation and commissioning of CT and X-ray imaging, oxygen plant, "
            "laboratory, theatre and ICU systems through to handover, and I trained the clinical "
            "users.",

            "I built the preventive maintenance schedules and service-record systems that reduced "
            "equipment downtime and kept devices safe for clinical use.",

            "I coordinated the procurement, technical specification, commissioning documentation "
            "and lifecycle management of hospital medical equipment.",

            "I prepared and maintained the compliance evidence files reviewed for COHSASA "
            "accreditation, and I supported the hospital through successful accreditation.",
        ],
    },
    {
        "role": "Biomedical Engineer",
        "org": "Norvik Hospital Ltd, Kampala",
        "dates": "Apr 2019 to Jan 2020",
        "bullets": [
            "I maintained and serviced the hospital diagnostic and patient-monitoring equipment.",
            "I supported the installation and calibration of medical imaging and ICU equipment.",
            "I helped establish the preventive maintenance schedules and service follow-up "
            "records.",
        ],
    },
    {
        "role": "Research Intern",
        "org": "Uganda Virus Research Institute (UVRI)",
        "dates": "Dec 2018 to Apr 2019",
        "bullets": [
            "I contributed to biomedical research projects and laboratory experimentation.",
            "I carried out laboratory data analysis, mathematical modelling and research pipeline "
            "design in Python.",
        ],
    },
    {
        "role": "Teaching Assistant",
        "org": "Makerere University, College of Health Sciences",
        "dates": "Mar 2016 to Aug 2017",
        "bullets": [
            "I assisted in teaching biomedical engineering course units at undergraduate level.",
            "I supervised and supported the laboratory practical sessions for engineering "
            "students.",
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
     "I build the community, screening, referral and follow-up data capture, the database "
     "structure, the cloud data flows, the management dashboards and the web and mobile "
     "delivery. I work in Next.js, TypeScript, React Native and cloud data services."),

    ("Cross-sector data collection and reporting application",
     "Founder and Lead Developer",
     "In development",
     "2026 to present",
     "I am building an application that simplifies data collection and reporting for professions "
     "that live on records: accounting, research, biology and laboratory science, agriculture and "
     "social media among them. Rather than collecting data now and writing the report later, the "
     "user works inside the app and it generates the information and reports in real time as "
     "they go."),

    ("COVID-19 oxygen therapy management and demand modelling",
     "Biomedical Manager",
     "International Hospital Kampala, with Ministry of Health planning support",
     "2020 to 2022",
     "I ran oxygen therapy management for the hospital through the pandemic and I developed the "
     "formulas for calculating patient oxygen consumption. My figures were used to size plant "
     "capacity, cylinder stock and flow requirements, and to support planning decisions at "
     "managerial and national level."),

    ("AI reception agent",
     "Designer and Developer",
     "Law firm, Texas, USA",
     "2025 to present",
     "I built and deployed an AI agent that answers live inbound client calls, books and "
     "reschedules appointments, and runs post-contact client follow-up. It is in active "
     "production use."),

    ("Shulekeeper school information system",
     "Assistant Software Developer",
     "Shulekeeper",
     "2021 to 2024",
     "I contributed front-end and back-end development, third-party integrations, testing, user "
     "training and documentation to a multi-user records and reporting platform built with React "
     "and Next.js, TypeScript, Laravel, MySQL and AWS."),

    ("Endoscopy image capture and reporting software",
     "Lead Software Developer",
     "St. Catherine Hospital and St. Francis Hospital",
     "2017 to 2019",
     "I led the build of a clinical image capture and procedure reporting system in JavaFX and "
     "MySQL. I did the algorithm design, the relational modelling, the legacy data migration, the "
     "on-site installation, the user training and the user manual."),

    ("Wekebere foetal heart-rate monitoring system",
     "Electronics Developer and Programmer",
     "Wekebere",
     "2017 to 2024",
     "I built an Android and Arduino monitoring system for third-trimester foetal heart-rate "
     "measurement, and I implemented the Java serial communication between the device and the "
     "phone, with Microsoft Azure cloud storage."),
]

# --------------------------------------------------------------------------
# Professional competencies
# --------------------------------------------------------------------------

COMPETENCIES = [
    ("Programme and project management",
     "I managed a regional biomedical team across East and Central Africa, and a hospital "
     "biomedical function for four years: workplans, schedules, procurement coordination and "
     "management reporting."),

    ("Monitoring, evaluation and data",
     "I lead the FCHIP indicator, data-quality, dashboard and reporting build. I built the "
     "equipment registers, maintenance schedules and routine management reports at IHK and GFF. "
     "I analyse in Python, SQL, Java, JavaScript and MATLAB."),

    ("Proposal, grant and technical writing",
     "I write and edit institutional proposals, grant applications, investment materials, "
     "partnership documents, technical documentation and user manuals."),

    ("Data analysis and automated auditing",
     "I clean, model and report on operational and clinical data, and I build automated auditing "
     "and verification tooling that checks records in bulk rather than by hand."),

    ("Procurement, specification and compliance",
     "I coordinated specification, procurement, commissioning and lifecycle documentation, and I "
     "kept the compliance evidence reviewed for COHSASA accreditation."),

    ("Stakeholder and partnership management",
     "I chair partner meetings and write them up, and I work with hospital departments, "
     "suppliers, trainees, accreditation assessors, donors and prospective partners."),

    ("Training and capacity building",
     "I trained biomedical engineers, technicians and clinical users across multiple facilities, "
     "and I taught undergraduate engineering laboratory practicals at Makerere University."),

    ("Communication and adaptability",
     "I write and speak clear English and I present comfortably to clinical, executive and donor "
     "audiences. I settle quickly into new countries, teams, codebases and equipment brands."),
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
     "Division 1."),
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
    "I play music, the saxophone in particular.  ·  I watch film.  ·  I train and exercise "
    "regularly.  ·  I enjoy travelling and working in unfamiliar places, which the regional field "
    "work has made second nature."
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
