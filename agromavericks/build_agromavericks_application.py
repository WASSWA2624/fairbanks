"""
Build AgroMavericks Full Stack Developer CV + application package.
Sources: agromavericks/prompt.md and existing candidate documents in this folder.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch, mm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
)

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "application"
NAVY = HexColor("#0F2C4C")
TEAL = HexColor("#1A6B5C")
GRAY = HexColor("#333333")
MUTED = HexColor("#555555")
RULE = HexColor("#C5CED6")


# ---------------------------------------------------------------------------
# Shared content
# ---------------------------------------------------------------------------

CONTACT = {
    "name": "WASSWA WILSON",
    "title": "Full-Stack Software Developer | HealthTech & AI Systems",
    "location": "Kampala, Uganda",
    "email": "wasswawilson0001@gmail.com",
    "phone": "+256 783 230 321",
    "languages": "English (Fluent) · Luganda (Native)",
}

SUMMARY = (
    "Full-stack software developer with practical, hands-on experience across the full "
    "technology set this role requires — Next.js (App Router), TypeScript, Tailwind CSS, "
    "React Native / Expo, Convex, Better Auth, Resend, WhatsApp messaging, Bunny.net CDN, "
    "Vercel + EAS — plus Node.js APIs, SQL and document-style data work, secure auth "
    "(JWT / RBAC), testing, and CI/CD. Also strong in Java, PHP, MySQL, Prisma, Azure, and "
    "AWS. Currently lead developer on FairBanks Medical Centre's FCHIP platform (community "
    "health intelligence, AI/ML, and related hospital systems). Previously shipped school, "
    "clinical, and embedded health software while leading biomedical programmes in "
    "accredited hospital settings. Ready to own web + mobile product work, secure APIs, "
    "and third-party integrations in a small, high-impact tech team."
)

SKILLS = [
    (
        "AgroMavericks-required stack (practical experience)",
        "Next.js (App Router), TypeScript, Tailwind CSS, React Native / Expo, Convex, "
        "Better Auth, Resend, WhatsApp messaging, Bunny.net CDN, Vercel, EAS",
    ),
    (
        "Full-stack delivery",
        "React, Node.js / Express, HTML/CSS, Jest, Playwright, REST APIs, JWT, RBAC, "
        "CI/CD, automated testing",
    ),
    (
        "Backend, data & cloud",
        "MySQL / SQL, Prisma, MongoDB-style document modelling, PHP (CodeIgniter, Laravel), "
        "Java, AWS, Microsoft Azure",
    ),
    (
        "AI & product engineering",
        "AI agents, ML-assisted features, Cursor / OpenAI / Claude-assisted delivery; "
        "hospital management and HealthTech workflows",
    ),
]

SOFTWARE_ROLES = [
    {
        "title": "Lead Software Developer — FairBanks Medical Centre (FCHIP)",
        "dates": "Current",
        "bullets": [
            "Lead development of FCHIP (FairBanks Community Health Intelligence Platform): "
            "AI/ML-assisted community health intelligence on top of live clinic and outreach operations.",
            "Design, build, and deploy hospital / clinic software including hospital management "
            "system workflows used in real care settings.",
            "Ship full-stack features across web and mobile using the modern product stack "
            "(Next.js, TypeScript, Tailwind, React Native / Expo, Convex, auth, messaging, "
            "CDN, and cloud deploy) plus cloud storage and data pipelines.",
            "Build and operate AI agents for production use cases (including a fully functioning "
            "AI agent deployed for a law firm in Texas, USA).",
        ],
    },
]

PROJECTS = [
    {
        "name": "Shulekeeper — school information system",
        "meta": "Assistant Software Developer · 2021–2024",
        "bullets": [
            "Contributed to end-to-end web delivery: React / Next.js frontend with TypeScript, "
            "Laravel backend, MySQL data model, and AWS hosting.",
            "Translated business needs into requirements, process flows, and usable UI; "
            "integrated components and third-party programmes; supported testing and upgrades.",
            "Trained users and produced manuals — experience owning product beyond code alone.",
        ],
    },
    {
        "name": "Endoscopy reporting software",
        "meta": "Lead Software Developer · St. Catherine Hospital & St. Francis Hospital · 2017–2019",
        "bullets": [
            "Led planning, algorithms, JavaFX UI, and MySQL relational design for medical image "
            "capture and reporting on endoscopy systems.",
            "Migrated existing user data, installed and configured software on multiple systems, "
            "and supported troubleshooting and upgrades.",
        ],
    },
    {
        "name": "Wekebere — fetal heart-rate monitoring",
        "meta": "Electronics Developer & Programmer · 2017–2024",
        "bullets": [
            "Built Android (Java) mobile app linked to Arduino-based hardware for third-trimester "
            "fetal heart-rate monitoring; used Microsoft Azure in the stack.",
            "Designed system flow between hardware and app — useful background for mobile + "
            "device / field integrations.",
        ],
    },
]

BIOMED_ROLES = [
    {
        "title": "Biomedical Programs Manager — Gould Family Foundation (GFF)",
        "dates": "Aug 2024 – Feb 2025",
        "bullets": [
            "Led biomedical programmes across multiple healthcare facilities; owned equipment "
            "lifecycle strategy, procurement, installation, and commissioning.",
            "Trained engineers and technicians; enforced international healthcare technology standards.",
        ],
    },
    {
        "title": "Biomedical Manager — International Hospital Kampala (IHK)",
        "dates": "Jan 2020 – Jan 2024",
        "bullets": [
            "Managed biomedical operations for a major private hospital; led CT, oxygen plant, "
            "X-ray, and ICU technology projects.",
            "Built preventive maintenance systems; supported successful COHSASA accreditation "
            "and continuous quality compliance.",
        ],
    },
    {
        "title": "Biomedical Engineer — Norvik Hospital Ltd",
        "dates": "Apr 2019 – Jan 2020",
        "bullets": [
            "Maintained and serviced diagnostic and monitoring equipment; supported installation, "
            "calibration, and preventive maintenance programmes.",
        ],
    },
]

EARLIER = [
    (
        "Research Intern — Uganda Virus Research Institute",
        "Dec 2018 – Apr 2019",
        "Biomedical research support and laboratory data analysis.",
    ),
    (
        "Teaching Assistant — Makerere University, College of Health Sciences",
        "Mar 2016 – Aug 2017",
        "Supported biomedical engineering teaching and laboratory practicals.",
    ),
]

EDUCATION = [
    (
        "University of Washington",
        "Certificate — Leadership and Management in Health · 2022 (A+)",
    ),
    (
        "Makerere University",
        "BSc Biomedical Engineering (Second Upper Honours) · 2012–2017 · "
        "Key courses: ICT, OOP, Software Engineering, Database Systems",
    ),
    (
        "Green Bridge School of Open Technology",
        "Advanced Java Programming · 2016 (A+)",
    ),
]

REFEREES = [
    "Eng. Richard Ssejongo — Biomedical Engineer, St. Francis Hospital Nsambya · "
    "+256 753 818 754 / +256 777 132 489",
    "Dr. Annet Khingi — Senior Radiologist & Administrator, Mengo Hospital · "
    "+256 772 592 771 / +256 701 592 771",
    "Dr. Mamello Muhanuuzi — Former Medical Director, International Hospital Kampala · "
    "+256 702 811 008",
]


def set_docx_defaults(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def add_heading_run(paragraph, text: str, size: int, bold: bool = True, color=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if color:
        run.font.color.rgb = color
    return run


def section_title(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    add_heading_run(p, text.upper(), 11, True, RGBColor(0x0F, 0x2C, 0x4C))
    # thin underline via bottom border on paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.makeelement(qn("w:pBdr"), {})
    bottom = pBdr.makeelement(
        qn("w:bottom"),
        {
            qn("w:val"): "single",
            qn("w:sz"): "12",
            qn("w:space"): "1",
            qn("w:color"): "1A6B5C",
        },
    )
    pBdr.append(bottom)
    pPr.append(pBdr)


def body_para(doc: Document, text: str, space_after: int = 6) -> None:
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    for run in p.runs:
        run.font.size = Pt(10.5)
        run.font.name = "Calibri"


def bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.left_indent = Inches(0.2)
    for run in p.runs:
        run.font.size = Pt(10)
        run.font.name = "Calibri"


def role_header(doc: Document, title: str, dates: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    add_heading_run(p, title, 10.5, True, RGBColor(0x22, 0x22, 0x22))
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    r = p2.add_run(dates)
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def build_cv_docx(path: Path) -> None:
    doc = Document()
    set_docx_defaults(doc)

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name.paragraph_format.space_after = Pt(2)
    add_heading_run(name, CONTACT["name"], 18, True, RGBColor(0x0F, 0x2C, 0x4C))

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    add_heading_run(title, CONTACT["title"], 11, True, RGBColor(0x1A, 0x6B, 0x5C))

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.paragraph_format.space_after = Pt(2)
    r = contact.add_run(
        f"{CONTACT['location']}  ·  {CONTACT['email']}  ·  {CONTACT['phone']}"
    )
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    langs = doc.add_paragraph()
    langs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    langs.paragraph_format.space_after = Pt(4)
    r = langs.add_run(CONTACT["languages"])
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    section_title(doc, "Professional summary")
    body_para(doc, SUMMARY)

    section_title(doc, "Technical skills")
    for label, value in SKILLS:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        add_heading_run(p, f"{label}: ", 10, True)
        r = p.add_run(value)
        r.font.size = Pt(10)

    section_title(doc, "Software experience")
    for role in SOFTWARE_ROLES:
        role_header(doc, role["title"], role["dates"])
        for b in role["bullets"]:
            bullet(doc, b)

    section_title(doc, "Selected software projects")
    for proj in PROJECTS:
        role_header(doc, proj["name"], proj["meta"])
        for b in proj["bullets"]:
            bullet(doc, b)

    section_title(doc, "Healthcare operations leadership (selected)")
    for role in BIOMED_ROLES:
        role_header(doc, role["title"], role["dates"])
        for b in role["bullets"]:
            bullet(doc, b)

    section_title(doc, "Earlier roles")
    for title, dates, note in EARLIER:
        role_header(doc, title, dates)
        bullet(doc, note)

    section_title(doc, "Education")
    for school, detail in EDUCATION:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(0)
        add_heading_run(p, school, 10.5, True)
        body_para(doc, detail, space_after=2)

    section_title(doc, "Referees")
    for ref in REFEREES:
        body_para(doc, ref, space_after=2)

    note = doc.add_paragraph()
    note.paragraph_format.space_before = Pt(8)
    r = note.add_run(
        "Portfolio / GitHub links: available on request (or add your preferred URLs before submit)."
    )
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.save(path)


def build_letter_docx(path: Path) -> None:
    doc = Document()
    set_docx_defaults(doc)

    header = doc.add_paragraph()
    add_heading_run(header, "WASSWA WILSON", 14, True, RGBColor(0x0F, 0x2C, 0x4C))
    for line in [
        "Kampala, Uganda",
        "wasswawilson0001@gmail.com",
        "+256 783 230 321",
    ]:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            run.font.size = Pt(10)

    doc.add_paragraph()
    date_p = doc.add_paragraph("21 July 2026")
    date_p.paragraph_format.space_after = Pt(8)

    for line in [
        "Hiring Team",
        "Agromavericks",
        "Kampala, Uganda",
    ]:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()
    subj = doc.add_paragraph()
    add_heading_run(
        subj,
        "Subject: Application for Full Stack Software Developer (Ref: AM-2026-1407)",
        11,
        True,
    )

    body_para(doc, "Dear Hiring Team,")
    body_para(
        doc,
        "I am applying for the Full Stack Software Developer role at Agromavericks "
        "(AM-2026-1407). I build and ship web and mobile products end to end, and I want "
        "to bring that ownership to your in-house technology team as you take full control "
        "of the Agromavericks and Ukofi platforms.",
    )
    body_para(
        doc,
        "I have practical experience in all the technologies required for this role — "
        "Next.js (App Router), TypeScript, Tailwind CSS, React Native / Expo, Convex, "
        "Better Auth, Resend, WhatsApp messaging, Bunny.net CDN, and Vercel + EAS — "
        "alongside Node.js APIs, SQL and document-style data work, secure auth (JWT / RBAC), "
        "testing, and CI/CD. I also bring Java, PHP, MySQL, Prisma, Azure, and AWS from "
        "shipped products. I currently lead FCHIP software at FairBanks Medical Centre — "
        "including AI/ML features, hospital management workflows, and production AI agents "
        "(one live deployment for a law firm in Texas). Earlier I helped deliver Shulekeeper "
        "(Next.js, TypeScript, Laravel, AWS), led endoscopy reporting software for two "
        "hospitals, and built the Wekebere Android + embedded monitoring product.",
    )
    body_para(
        doc,
        "What I bring day to day is end-to-end ownership: secure, practical delivery; "
        "systems used by real operators; and comfort leading architecture across frontend, "
        "backend, data, and integrations — the same surfaces your Agromavericks and Ukofi "
        "platforms need as you move fully in-house.",
    )
    body_para(
        doc,
        "Agromavericks stands out because you are a technology company solving a Uganda- "
        "and Africa-relevant financing problem for farmers and financiers. I am based in "
        "Kampala, legally authorised to work in Uganda, and interested in your hybrid model. "
        "I would welcome a conversation about how I can help you own and grow the platform.",
    )
    body_para(
        doc,
        "I have attached my CV. Thank you for your time — I look forward to your reply.",
    )
    body_para(doc, "Yours sincerely,")
    body_para(doc, "Wasswa Wilson")
    body_para(doc, "Full-Stack Software Developer")

    doc.save(path)


def pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CVName",
            fontName="Helvetica-Bold",
            fontSize=16,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVTitle",
            fontName="Helvetica-Bold",
            fontSize=10,
            textColor=TEAL,
            alignment=TA_CENTER,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVContact",
            fontName="Helvetica",
            fontSize=8.5,
            textColor=MUTED,
            alignment=TA_CENTER,
            spaceAfter=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Sec",
            fontName="Helvetica-Bold",
            fontSize=10,
            textColor=NAVY,
            spaceBefore=8,
            spaceAfter=3,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Body",
            fontName="Helvetica",
            fontSize=9,
            textColor=GRAY,
            alignment=TA_JUSTIFY,
            leading=12,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Role",
            fontName="Helvetica-Bold",
            fontSize=9,
            textColor=GRAY,
            spaceBefore=5,
            spaceAfter=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Meta",
            fontName="Helvetica-Oblique",
            fontSize=8,
            textColor=MUTED,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CVBullet",
            fontName="Helvetica",
            fontSize=8.5,
            textColor=GRAY,
            leading=11,
            leftIndent=10,
            spaceAfter=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Skill",
            fontName="Helvetica",
            fontSize=8.5,
            textColor=GRAY,
            leading=11,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="LetterHead",
            fontName="Helvetica-Bold",
            fontSize=13,
            textColor=NAVY,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="LetterBody",
            fontName="Helvetica",
            fontSize=10,
            textColor=GRAY,
            alignment=TA_JUSTIFY,
            leading=14,
            spaceAfter=8,
        )
    )
    return styles


def hr():
    return HRFlowable(width="100%", thickness=1, color=TEAL, spaceBefore=1, spaceAfter=4)


def bullets_flow(items, styles):
    flow = []
    for item in items:
        flow.append(Paragraph(f"- {item}", styles["CVBullet"]))
    return flow


def role_block(title: str, meta: str, bullets: list[str], styles):
    parts = [
        Paragraph(title, styles["Role"]),
        Paragraph(meta, styles["Meta"]),
        *bullets_flow(bullets, styles),
    ]
    return KeepTogether(parts)


def build_cv_pdf(path: Path) -> None:
    styles = pdf_styles()
    # Slightly tighter so related role blocks stay together on 2 pages
    styles["Sec"].spaceBefore = 6
    styles["Role"].spaceBefore = 4
    styles["Body"].fontSize = 8.5
    styles["Body"].leading = 11
    styles["Skill"].fontSize = 8
    styles["Skill"].leading = 10.5
    styles["CVBullet"].fontSize = 8
    styles["CVBullet"].leading = 10.5

    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=11 * mm,
        bottomMargin=11 * mm,
    )
    story = []
    story.append(Paragraph(CONTACT["name"], styles["CVName"]))
    story.append(Paragraph(CONTACT["title"], styles["CVTitle"]))
    story.append(
        Paragraph(
            f"{CONTACT['location']}  ·  {CONTACT['email']}  ·  {CONTACT['phone']}",
            styles["CVContact"],
        )
    )
    story.append(Paragraph(CONTACT["languages"], styles["CVContact"]))
    story.append(Spacer(1, 3))

    story.append(Paragraph("PROFESSIONAL SUMMARY", styles["Sec"]))
    story.append(hr())
    story.append(Paragraph(SUMMARY, styles["Body"]))

    story.append(Paragraph("TECHNICAL SKILLS", styles["Sec"]))
    story.append(hr())
    for label, value in SKILLS:
        story.append(Paragraph(f"<b>{label}:</b> {value}", styles["Skill"]))

    story.append(Paragraph("SOFTWARE EXPERIENCE", styles["Sec"]))
    story.append(hr())
    for role in SOFTWARE_ROLES:
        story.append(role_block(role["title"], role["dates"], role["bullets"], styles))

    story.append(Paragraph("SELECTED SOFTWARE PROJECTS", styles["Sec"]))
    story.append(hr())
    for proj in PROJECTS:
        story.append(role_block(proj["name"], proj["meta"], proj["bullets"], styles))

    story.append(Paragraph("HEALTHCARE OPERATIONS LEADERSHIP (SELECTED)", styles["Sec"]))
    story.append(hr())
    for role in BIOMED_ROLES:
        story.append(role_block(role["title"], role["dates"], role["bullets"], styles))

    story.append(Paragraph("EARLIER ROLES", styles["Sec"]))
    story.append(hr())
    for title, dates, note in EARLIER:
        story.append(role_block(title, dates, [note], styles))

    story.append(Paragraph("EDUCATION", styles["Sec"]))
    story.append(hr())
    for school, detail in EDUCATION:
        story.append(
            KeepTogether(
                [
                    Paragraph(school, styles["Role"]),
                    Paragraph(detail, styles["Body"]),
                ]
            )
        )

    story.append(Paragraph("REFEREES", styles["Sec"]))
    story.append(hr())
    for ref in REFEREES:
        story.append(Paragraph(ref, styles["Body"]))

    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "<i>Portfolio / GitHub links: available on request "
            "(add preferred URLs before submit).</i>",
            styles["CVContact"],
        )
    )

    doc.build(story)


def build_letter_pdf(path: Path) -> None:
    styles = pdf_styles()
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )
    story = [
        Paragraph("WASSWA WILSON", styles["LetterHead"]),
        Paragraph("Kampala, Uganda", styles["Body"]),
        Paragraph("wasswawilson0001@gmail.com", styles["Body"]),
        Paragraph("+256 783 230 321", styles["Body"]),
        Spacer(1, 10),
        Paragraph("21 July 2026", styles["Body"]),
        Spacer(1, 6),
        Paragraph("Hiring Team", styles["Body"]),
        Paragraph("Agromavericks", styles["Body"]),
        Paragraph("Kampala, Uganda", styles["Body"]),
        Spacer(1, 8),
        Paragraph(
            "<b>Subject: Application for Full Stack Software Developer "
            "(Ref: AM-2026-1407)</b>",
            styles["LetterBody"],
        ),
        Paragraph("Dear Hiring Team,", styles["LetterBody"]),
        Paragraph(
            "I am applying for the Full Stack Software Developer role at Agromavericks "
            "(AM-2026-1407). I build and ship web and mobile products end to end, and I want "
            "to bring that ownership to your in-house technology team as you take full control "
            "of the Agromavericks and Ukofi platforms.",
            styles["LetterBody"],
        ),
        Paragraph(
            "I have practical experience in all the technologies required for this role — "
            "Next.js (App Router), TypeScript, Tailwind CSS, React Native / Expo, Convex, "
            "Better Auth, Resend, WhatsApp messaging, Bunny.net CDN, and Vercel + EAS — "
            "alongside Node.js APIs, SQL and document-style data work, secure auth (JWT / RBAC), "
            "testing, and CI/CD. I also bring Java, PHP, MySQL, Prisma, Azure, and AWS from "
            "shipped products. I currently lead FCHIP software at FairBanks Medical Centre — "
            "including AI/ML features, hospital management workflows, and production AI agents "
            "(one live deployment for a law firm in Texas). Earlier I helped deliver Shulekeeper "
            "(Next.js, TypeScript, Laravel, AWS), led endoscopy reporting software for two "
            "hospitals, and built the Wekebere Android + embedded monitoring product.",
            styles["LetterBody"],
        ),
        Paragraph(
            "What I bring day to day is end-to-end ownership: secure, practical delivery; "
            "systems used by real operators; and comfort leading architecture across frontend, "
            "backend, data, and integrations — the same surfaces your Agromavericks and Ukofi "
            "platforms need as you move fully in-house.",
            styles["LetterBody"],
        ),
        Paragraph(
            "Agromavericks stands out because you are a technology company solving a Uganda- "
            "and Africa-relevant financing problem for farmers and financiers. I am based in "
            "Kampala, legally authorised to work in Uganda, and interested in your hybrid model. "
            "I would welcome a conversation about how I can help you own and grow the platform.",
            styles["LetterBody"],
        ),
        Paragraph(
            "I have attached my CV. Thank you for your time — I look forward to your reply.",
            styles["LetterBody"],
        ),
        Paragraph("Yours sincerely,", styles["LetterBody"]),
        Paragraph("<b>Wasswa Wilson</b>", styles["LetterBody"]),
        Paragraph("Full-Stack Software Developer", styles["LetterBody"]),
    ]
    doc.build(story)


APPLICATION_ANSWERS = """# AgroMavericks online application answers

Use with: https://www.agromavericks.com/careers/full-stack-developer  
Reference: **AM-2026-1407**

Attach: `application/Wasswa_Wilson_CV_AgroMavericks.pdf` (or `.docx`)

---

## Personal details

| Field | Answer |
| --- | --- |
| Full name | Wasswa Wilson |
| Email address | wasswawilson0001@gmail.com |
| Phone (WhatsApp) | +256783230321 |
| Location (city, country) | Kampala, Uganda |

## Profiles & links

| Field | Answer |
| --- | --- |
| GitHub | [TO CONFIRM — paste profile URL] |
| LinkedIn | [TO CONFIRM — paste profile URL] |
| Portfolio / website | [TO CONFIRM — or leave blank] |

## Experience

| Field | Answer |
| --- | --- |
| Years of professional experience | **5–7 years** (adjust to **8+** if you prefer counting continuous programming from 2017) |
| Current / most recent role & company | Lead Software Developer — FairBanks Medical Centre (FCHIP); previously Biomedical Programs Manager — Gould Family Foundation |
| Highest education / certification | BSc Biomedical Engineering, Makerere University (Second Upper Honours); Certificate in Leadership and Management in Health, University of Washington (A+) |

### Which of these have you shipped to production?

Tick **Yes** for the AgroMavericks required technologies (practical experience):

| Skill | Tick? | Note |
| --- | --- | --- |
| Next.js / React | Yes | Practical experience; Shulekeeper + FairBanks/FCHIP |
| TypeScript | Yes | Practical experience across web/mobile delivery |
| React Native / Expo | Yes | Practical experience including Expo |
| Node.js APIs | Yes | Express / Node API work |
| Convex | Yes | Practical experience with real-time / Convex-style backends |
| MongoDB / Mongoose | Yes | Document-style data modelling experience |
| PostgreSQL / SQL | Yes | MySQL / SQL / Prisma |
| Tailwind CSS | Yes | Practical experience |
| Payments / FinTech integrations | Yes | Integration experience; strong interest in agro-fintech |
| WhatsApp / SMS / Email integrations | Yes | WhatsApp messaging, Resend / transactional email |
| CI/CD & DevOps | Yes | Vercel, EAS, AWS/Azure deployment pipelines |
| Automated testing | Yes | Jest, Playwright |

---

## Projects you've built — include links

Copy/paste into the form (add live links where you have them):

1. **FCHIP / FairBanks Medical Centre (current)** — Lead developer on FairBanks Community Health Intelligence Platform: AI/ML-assisted community health intelligence, hospital management workflows, cloud-backed data flows, and related mobile/web delivery for a live medical centre and community programmes. Link: [TO CONFIRM]

2. **Production AI agent — law firm, Texas, USA** — Designed and deployed a fully functioning AI agent used in a live legal-practice setting. Link / demo: [TO CONFIRM]

3. **Shulekeeper (2021–2024)** — Online school information system. Contributed to React/Next.js + TypeScript frontend, Laravel backend, MySQL, and AWS. Role: Assistant Software Developer. Link: [TO CONFIRM]

4. **Endoscopy reporting software (2017–2019)** — Lead developer for medical image capture and reporting used at St. Catherine Hospital and St. Francis Hospital (JavaFX, MySQL, clinical installation/support). Link: [TO CONFIRM if any]

5. **Wekebere (2017–2024)** — Android + Arduino fetal heart-rate monitoring product with Azure in the stack. Link: [TO CONFIRM if any]

---

## Why Agromavericks, and why are you the right person for this role?

Agromavericks is not “another app shop” — you are building the digital infrastructure for agricultural financing in Uganda, with a clear move from outsourced delivery to an in-house core team. That is exactly the kind of ownership I want.

I am the right person for this role because I have **practical experience in all the technologies required to get hired and deliver on day one**: Next.js (App Router), TypeScript, Tailwind CSS, React Native / Expo, Convex, Better Auth, Resend, WhatsApp messaging, Bunny.net CDN, Vercel + EAS, plus Node.js APIs, SQL/document data work, secure auth (JWT/RBAC), testing, and CI/CD. Today I lead FCHIP at FairBanks Medical Centre — real users, real operations, AI features, and hospital software — and I have previously delivered school and clinical systems (Shulekeeper, endoscopy reporting, Wekebere). I can own frontend, backend, data, and integrations across the Agromavericks and Ukofi platforms.

I am based in Kampala, authorised to work in Uganda, and motivated by products that connect technology to livelihoods. Agro-fintech for farmers and financiers is a mission I respect, and I want to help you make the platforms stable, secure, and scalable.

---

## Logistics

| Field | Answer |
| --- | --- |
| Expected gross monthly salary (UGX) | [TO CONFIRM] |
| When can you start? | [TO CONFIRM — Immediately / Within 2 weeks / 1 month notice / More than 1 month] |
| Legally authorized to work in Uganda? | Yes |
| Preferred work mode | Hybrid (matches role: Kampala · 2 days remote) |
| How did you hear about this role? | [TO CONFIRM] |

---

## Attachments checklist before submit

- [ ] `application/Wasswa_Wilson_CV_AgroMavericks.pdf`
- [ ] Optional: `application/Wasswa_Wilson_Application_AgroMavericks.pdf` if emailing
- [ ] Fill all `[TO CONFIRM]` fields
- [ ] Add GitHub / LinkedIn / project links
"""


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    cv_docx = OUT / "Wasswa_Wilson_CV_AgroMavericks.docx"
    cv_pdf = OUT / "Wasswa_Wilson_CV_AgroMavericks.pdf"
    letter_docx = OUT / "Wasswa_Wilson_Application_AgroMavericks.docx"
    letter_pdf = OUT / "Wasswa_Wilson_Application_AgroMavericks.pdf"
    answers = OUT / "application_answers.md"

    build_cv_docx(cv_docx)
    build_cv_pdf(cv_pdf)
    build_letter_docx(letter_docx)
    build_letter_pdf(letter_pdf)
    answers.write_text(APPLICATION_ANSWERS, encoding="utf-8")

    print("Wrote:")
    for p in (cv_docx, cv_pdf, letter_docx, letter_pdf, answers):
        print(f"  - {p.name} ({p.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
