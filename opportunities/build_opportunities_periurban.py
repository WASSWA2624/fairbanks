#!/usr/bin/env python3
"""Deprecated — peri-urban scan was merged into opportunities.xlsx (Jul 2026).

Use build_opportunities.py instead. This stub remains only so old docs or
habits do not recreate opportunities_periurban_fchip_2026-07.xlsx.
"""

from pathlib import Path
import runpy
import sys

print(
    "NOTE: opportunities_periurban_fchip_2026-07.xlsx was merged into "
    "opportunities/opportunities.xlsx and deleted.\n"
    "Running the canonical builder: build_opportunities.py"
)
runpy.run_path(str(Path(__file__).resolve().parent / "build_opportunities.py"), run_name="__main__")
sys.exit(0)
