#!/usr/bin/env python3
"""Build synced Word, PDF, and PowerPoint for govtech."""

from pathlib import Path
import sys

PROJECT = Path(__file__).resolve().parent
ROOT = PROJECT.parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from application_docs.generator import build_all
from application_docs.specs import GOVTECH

if __name__ == "__main__":
    build_all(GOVTECH, PROJECT)
