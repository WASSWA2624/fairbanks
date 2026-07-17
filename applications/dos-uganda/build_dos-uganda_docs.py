#!/usr/bin/env python3
"""Build synced Word, PDF, and PowerPoint for dos-uganda."""

from pathlib import Path
import sys

PROJECT = Path(__file__).resolve().parent
ROOT = PROJECT.parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from application_docs.generator import build_all
from application_docs.specs import DOS_UGANDA

if __name__ == "__main__":
    build_all(DOS_UGANDA, PROJECT)
