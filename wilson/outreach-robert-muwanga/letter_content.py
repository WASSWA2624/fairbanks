# -*- coding: utf-8 -*-
"""
Content for the open introduction to Mr. Robert Muwanga, referred by Juliet.

House rule, same as the CV: no em dashes anywhere.
"""

DATE = "11 September 2026"

RECIPIENT = ["Mr. Robert Muwanga", "Kampala, Uganda"]

SENDER = {
    "name": "Wasswa Wilson",
    "title": "Biomedical Engineer, Researcher and Full-Stack Software Developer",
    "email": "wasswawilson0001@gmail.com",
    "phone": "+256 783 230 321",
    "city": "Kampala, Uganda",
}

SUBJECT_LINE = "An introduction, on the recommendation of Juliet"

SALUTATION = "Dear Mr. Muwanga,"

# --------------------------------------------------------------------------
# Cover letter body. Tuples of (heading or None, paragraph text).
# --------------------------------------------------------------------------

LETTER_BODY = [
    (None,
     "My colleague Juliet suggested that I write to you and share my CV, and I am glad to do so."),

    (None,
     "My name is Wasswa Wilson. I am a biomedical engineer, researcher and full-stack software "
     "developer based in Kampala, with over nine years of work across health research, data "
     "systems and software engineering. I am writing to introduce myself rather than to apply for "
     "anything in particular. If something in what follows is useful to you, or to someone in "
     "your network, I would welcome the conversation."),

    ("Research and analysis",
     "I began at the Uganda Virus Research Institute doing laboratory data analysis, mathematical "
     "modelling and research pipeline design in Python, and I have kept that method since. During "
     "the COVID-19 response at International Hospital Kampala I developed the formulas for "
     "calculating patient oxygen consumption: I defined the model, tested it against real "
     "hospital demand and handed it to management as a planning tool used to size plant capacity "
     "and supply. I analyse in Python, SQL, MATLAB and Excel."),

    ("Data management and reporting",
     "I design and run the data layer of the FairBanks Community Health Intelligence Platform: "
     "capture schemas, validation rules, cloud data flows and the dashboards management reads. "
     "Earlier I built and maintained equipment registers and service histories across "
     "multi-facility programmes, accurate enough to pass COHSASA accreditation review, which is "
     "data quality assurance against a published standard. I also build automated auditing and "
     "verification tooling that checks records in bulk rather than by hand, and I work with AI "
     "agents and agent-assisted development daily."),

    ("Proposal and grant writing",
     "I write and edit what FairBanks submits: the WA Foundation Category 2 full project proposal "
     "for FairBanks Community Health Reach, accelerator and fellowship applications, investment "
     "proposals for the maternity and diagnostics projects, partnership documents and pitch "
     "decks, handling the narrative, the annexes, the financial presentation and the evidence "
     "files."),

    (None,
     "Alongside this I am building a cross-sector application that simplifies data collection and "
     "reporting for professions that live on records, accounting and audit among them, so the "
     "report is produced in real time as the work is done. Before moving fully into software I "
     "managed the biomedical programme for the Gould Family Foundation across East and Central "
     "Africa and ran the biomedical function at International Hospital Kampala for four years."),

    (None,
     "I have attached my CV, together with a single file containing this letter, the CV and my "
     "academic and professional credentials. I would value a short conversation at your "
     "convenience, whether about work within your own firm or simply to be pointed towards others "
     "who may find this useful."),

    (None,
     "Thank you for your time, and please pass on my thanks to Juliet."),
]

CLOSING = "Yours sincerely,"

# --------------------------------------------------------------------------
# Email. Plain text, for pasting straight into a mail client.
# --------------------------------------------------------------------------

EMAIL_SUBJECT = ("Introduction from Juliet: Wasswa Wilson, biomedical engineer "
                 "and software developer")

EMAIL_BODY = """Dear Mr. Muwanga,

My colleague Juliet suggested I write to you and share my CV, so please allow me to introduce myself.

I am Wasswa Wilson, a biomedical engineer, researcher and full-stack software developer based in Kampala. I began at the Uganda Virus Research Institute doing laboratory data analysis, mathematical modelling and research pipeline design in Python, and research and data have stayed at the centre of my work since.

Three things I do that may be relevant to you:

  Research and analysis. During the COVID-19 response at International Hospital Kampala I developed the formulas for calculating patient oxygen consumption, tested them against real hospital demand, and handed them to management as a planning tool used to size plant capacity and supply. I analyse in Python, SQL, MATLAB and Excel.

  Data management and reporting. I design and run the data layer of the FairBanks Community Health Intelligence Platform: capture schemas, validation rules, cloud data flows and the dashboards management reads. I also build automated auditing and verification tooling that checks records in bulk rather than by hand, and I work with AI agents and agent-assisted development daily.

  Proposal and grant writing. I write what FairBanks submits, including a full project proposal to the WA Foundation, accelerator applications and investment proposals, with the annexes and evidence files that go with them.

Alongside this I am building an application that simplifies data collection and reporting for record-heavy professions, accounting and audit among them, producing the report in real time as the work is done.

I am not writing about a specific vacancy. This is an open introduction, and if anything here is useful to you or to someone in your network, I would be glad of a short conversation.

I have attached:

  1. My CV.
  2. A single file containing my cover letter, CV and supporting documents, that is my academic and professional credentials.

Thank you for your time, and my thanks to Juliet for the introduction.

Kind regards,

Wasswa Wilson
Biomedical Engineer, Researcher and Full-Stack Software Developer
wasswawilson0001@gmail.com
+256 783 230 321
"""

# --------------------------------------------------------------------------
# Dossier assembly
# --------------------------------------------------------------------------

DOSSIER_TITLE = "Professional Dossier"

CONTENTS = [
    ("1", "Cover letter", "Introduction to Mr. Robert Muwanga"),
    ("2", "Curriculum vitae", "Four pages"),
    ("3", "Academic and professional credentials", "Eight documents"),
    ("4", "Identification documents", "National identification card and driving licence"),
]

CONTENTS_NO_ID = CONTENTS[:3]

# Credentials, in the order they appear. (image name, caption)
CREDENTIALS = [
    ("image1.png", "Makerere University: Academic Transcript, BSc Biomedical Engineering, "
                   "Second Class Honours Upper Division"),
    ("image2.jpg", "Makerere University: Transcript key to grades and classification of awards"),
    ("image7.jpg", "University of Washington: Leadership and Management in Health, December 2022"),
    ("image6.jpg", "Greenbridge School of Open Technologies: Certificate in Java Programming, "
                   "Level 1, Grade A+, 2016"),
    ("image4.jpg", "Uganda National Examinations Board: Uganda Advanced Certificate of "
                   "Education, Mengo Secondary School, 2011"),
    ("image3.png", "Uganda National Examinations Board: Uganda Certificate of Education, "
                   "Division 1, St. John's Wakiso Secondary School, 2009"),
    ("image8.jpg", "Fire Technologies Limited: Fire safety, prevention, firefighting and "
                   "emergency scene management, International Hospital Kampala, April 2021"),
    ("image9.png", "COHSASA: Recognition of Achievement of Accreditation Status, Maintenance "
                   "Service, International Hospital Kampala, February 2022"),
]

# Images whose content sits rotated inside a portrait frame.
ROTATE_CCW = {"image6.jpg", "image7.jpg"}

# Identification pages, taken from Identification documents.pdf
IDENTIFICATION = [
    (0, "Republic of Uganda: National Identification Card, front"),
    (1, "Republic of Uganda: National Identification Card, reverse"),
    (2, "Republic of Uganda: Driving Licence, front"),
    (3, "Republic of Uganda: Driving Licence, reverse"),
]
