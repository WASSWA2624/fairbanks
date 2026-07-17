#!/usr/bin/env python3
"""
AWIEF Pitch n Grow 2026 — regenerate the full synced deliverable set.

Runs:
  build_awief_deck.py  → documents/awief_ppt.pptx
  build_photo_docs.py  → documents/awief_word.docx + documents/awief_pdf.pdf

Use: python applications/awief/build_awief_docs.py
"""

from pathlib import Path
import runpy
import sys

HERE = Path(__file__).resolve().parent


def main() -> int:
    scripts = [
        HERE / "build_awief_deck.py",
        HERE / "build_photo_docs.py",
    ]
    for script in scripts:
        print(f"=== {script.name} ===")
        runpy.run_path(str(script), run_name="__main__")
    print("AWIEF sync complete: awief_ppt.pptx, awief_word.docx, awief_pdf.pdf")
    return 0


if __name__ == "__main__":
    sys.exit(main())
