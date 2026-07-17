#!/usr/bin/env python3
"""Scaffold opportunity application folders and generate document packs.

Creates applications/{slug}/ with:
  - build_{slug}_docs.py
  - rules/source_of_truth.mdc
  - documents/{slug}_{pdf,ppt,word}.*  (via generator)

Does not touch applications/awief (already complete) or applications/fid.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

from application_docs.specs import SPECS  # noqa: E402
from application_docs.generator import build_all  # noqa: E402

APPS = ROOT / "applications"


def sot_text(spec: dict) -> str:
    slug = spec["slug"]
    prog = spec["meta"]["programme"]
    url = spec["url"]
    return f"""---
description: {prog} compliance and deliverable sync
globs: applications/{slug}/**
alwaysApply: false
---

# {slug.upper()} — source of truth

All work for this pack must align with [{prog}]({url}). That page is the source of truth for eligibility, theme, format, and deadlines. If internal docs conflict with the call, fix the docs to match the call.

Also follow `.cursor/rules/documents.mdc` and `.cursor/rules/folder_structure.mdc`.

## Mandatory checks

1. Keep claims consistent with the live call (eligibility, geography, benefits, deadlines).
2. Frame a clear **win-win**: value for FairBanks / FCHIP and value for the programme.
3. Stay rooted in the FairBanks community health cascade and FCHIP vision.
4. Use simple, human language. Do not invent slogans beyond **Your health, our mission.**
5. Before finalising, re-check the official programme page.

## FairBanks deliverable set

| Role | Path |
|---|---|
| Pitch deck | `applications/{slug}/documents/{slug}_ppt.pptx` |
| PDF | `applications/{slug}/documents/{slug}_pdf.pdf` |
| Word | `applications/{slug}/documents/{slug}_word.docx` |

1. Treat these files as one synchronized content set.
2. Prefer the most recently edited format; if unclear, follow this SoT and the live call.
3. Preserve established document layout and branding.
4. Regenerate with `applications/{slug}/build_{slug}_docs.py`.
5. Keep only the three official deliverables in `documents/`.

## Branding

- Show **Your health, our mission.** on every cover/title page or slide.
- Do not invent slogans or alter core claims without user direction.
"""


def build_script_text(slug: str) -> str:
    const = slug.upper().replace("-", "_")
    # Map slug to SPECS key constant names in specs.py
    mapping = {
        "auc": "AUC",
        "feminist-ai": "FEMINIST_AI",
        "girlcode": "GIRLCODE",
        "auda-srh": "AUDA_SRH",
        "gadfly": "GADFLY",
        "dos-uganda": "DOS_UGANDA",
        "whs": "WHS",
        "oyw": "OYW",
        "africa-cdc": "AFRICA_CDC",
        "ifad": "IFAD",
        "govtech": "GOVTECH",
    }
    name = mapping[slug]
    return f'''#!/usr/bin/env python3
"""Build synced Word, PDF, and PowerPoint for {slug}."""

from pathlib import Path
import sys

PROJECT = Path(__file__).resolve().parent
ROOT = PROJECT.parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from application_docs.generator import build_all
from application_docs.specs import {name}

if __name__ == "__main__":
    build_all({name}, PROJECT)
'''


def scaffold_one(slug: str, spec: dict, generate: bool = True) -> None:
    proj = APPS / slug
    (proj / "rules").mkdir(parents=True, exist_ok=True)
    (proj / "documents").mkdir(parents=True, exist_ok=True)

    sot = proj / "rules" / "source_of_truth.mdc"
    sot.write_text(sot_text(spec), encoding="utf-8")

    script = proj / f"build_{slug}_docs.py"
    script.write_text(build_script_text(slug), encoding="utf-8")

    print(f"Scaffolded {proj}")
    if generate:
        build_all(spec, proj)


def main():
    generate = "--scaffold-only" not in sys.argv
    for slug, spec in SPECS.items():
        assert spec["slug"] == slug
        scaffold_one(slug, spec, generate=generate)
    print(f"\nCompleted {len(SPECS)} opportunity application packs.")
    print("Note: applications/awief already exists with a custom pack — left unchanged.")


if __name__ == "__main__":
    main()
