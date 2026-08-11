#!/usr/bin/env python3
"""Build a simple one-sheet opportunities/activity_tracker.xlsx."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "activity_tracker.xlsx"
TRACKER = ROOT / "tracker.xlsx"
NEW_GRANTS = ROOT / "new_grants.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="0B3D2E")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
THIN = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
# Priority colours
P1 = PatternFill("solid", fgColor="FCE4D6")  # orange - do first
P2 = PatternFill("solid", fgColor="FFF2CC")  # yellow - this week
P3 = PatternFill("solid", fgColor="DDEBF7")  # blue - later
DONE = PatternFill("solid", fgColor="E2EFDA")  # green - done / applied


def load_tracker() -> list[dict]:
    wb = load_workbook(TRACKER, data_only=True)
    ws = wb["Opportunities"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        rows.append(
            {
                "title": str(r[0]).strip(),
                "url": (r[1] or "").strip(),
                "deadline": (r[3] or "").strip(),
                "status": (r[5] or "").strip(),
                "submission": (r[6] or "").strip(),
                "folder": (r[8] or "").strip(),
            }
        )
    return rows


def load_grants() -> list[dict]:
    wb = load_workbook(NEW_GRANTS, data_only=True)
    ws = wb["Opportunities"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        rows.append(
            {
                "title": str(r[0]).strip(),
                "url": (r[1] or "").strip(),
                "deadline": (r[3] or "").strip(),
                "status": (r[5] or "Not started").strip(),
                "number": (r[9] or "").strip() if len(r) > 9 else "",
            }
        )
    return rows


def short_title(text: str, n: int = 70) -> str:
    text = text.replace("\u2014", "-").replace("\u2013", "-")
    return text if len(text) <= n else text[: n - 1] + "..."


def is_applied(status: str) -> bool:
    s = status.lower()
    return any(
        x in s
        for x in ("complete", "submitted", "awarded", "submitted poc")
    )


def build() -> None:
    tracker = load_tracker()
    grants = load_grants()

    # One sheet only
    wb = Workbook()
    ws = wb.active
    ws.title = "Activity tracker"

    headers = [
        "Priority",
        "When",
        "Owner",
        "What to do",
        "Grant / item",
        "Application deadline",
        "Status",
        "Link or folder",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    rows: list[tuple[list, PatternFill]] = []

    # --- Priority 1: today / scheduled ---
    rows.append(
        (
            [
                "1 - Do first",
                "Today 11 Aug 2026, 08:00",
                "Racheal + team",
                "Meeting: prepare and apply for this CDC grant",
                "CDC-RFA-JG-26-0056 (detect / notify / respond globally)",
                "14 August 2026, 11:59 p.m. ET",
                "Meeting scheduled",
                "https://simpler.grants.gov/opportunity/8454e463-cd43-4d0d-97a2-8a4310e0ce6b",
            ],
            P1,
        )
    )
    rows.append(
        (
            [
                "1 - Do first",
                "Today 11 Aug 2026",
                "Racheal",
                "Contact someone at UVRI about possible collaboration",
                "UVRI partnership (helps CDC 0056 and Community Reach)",
                "14 August 2026 (CDC 0056)",
                "To do",
                "applications/cdc-ghs-global",
            ],
            P1,
        )
    )
    rows.append(
        (
            [
                "1 - Do first",
                "Today 11 Aug 2026",
                "Racheal",
                "Contact someone at IDI about possible collaboration",
                "IDI partnership (helps CDC 0056 and Community Reach)",
                "14 August 2026 (CDC 0056)",
                "To do",
                "applications/cdc-ghs-global",
            ],
            P1,
        )
    )
    rows.append(
        (
            [
                "1 - Do first",
                "Today 11 Aug 2026",
                "Team (together)",
                "Go through the FairBanks Blueprint together",
                "FairBanks Blueprint v1.0.1",
                "",
                "To do",
                ".cursor/source-of-truth/ (Blueprint v1.0.1 PDF)",
            ],
            P1,
        )
    )
    rows.append(
        (
            [
                "1 - Do first",
                "Today 11 Aug 2026",
                "Dianna",
                "Refuel Cursor subscription so FCHIP MVP work can continue",
                "FCHIP MVP tools",
                "",
                "To do",
                "Cursor subscription / billing",
            ],
            P1,
        )
    )

    # --- Priority 2: this week meetings / open drafts ---
    rows.append(
        (
            [
                "2 - This week",
                "Fri 14 Aug 2026, 11:30 p.m. Nairobi",
                "Team (invite from Emily Babirye)",
                "Join QuAM Plus training for CAPAID Partners (Zoom)",
                "QuAM Plus / CAPAID training",
                "",
                "Meeting scheduled",
                "Zoom: https://careorg.zoom.us/j/93605570080?pwd=8XhAa01MZazbvCTzf1kBccBBykTOVu.1 | ID 936 0557 0080 | Passcode 341875",
            ],
            P2,
        )
    )

    for t in tracker:
        if is_applied(t["status"]):
            continue
        st = t["status"].lower()
        if "draft" in st or "mvp" in st:
            title = short_title(t["title"])
            what = "Finish and submit application pack" if "draft" in st else "Finish MVP / pitch follow-up"
            if "0056" in t["title"]:
                # Already covered as today apply meeting
                continue
            rows.append(
                (
                    [
                        "2 - This week",
                        "Before deadline",
                        "Team",
                        what,
                        title,
                        t["deadline"],
                        t["status"],
                        t["folder"] or t["url"],
                    ],
                    P2,
                )
            )

    # Open grants from new_grants not already applied on tracker
    tracker_blob = " ".join(t["title"] for t in tracker)
    for g in grants:
        num = g.get("number") or ""
        # Skip if already complete on tracker
        matched_complete = False
        matched_draft = False
        for t in tracker:
            if num and num in t["title"]:
                if is_applied(t["status"]):
                    matched_complete = True
                else:
                    matched_draft = True
                break
        if matched_complete:
            continue
        if matched_draft and "0056" in num:
            continue  # today meeting
        if matched_draft:
            continue  # already listed as drafting above
        # Not on tracker / not started
        priority = "2 - This week"
        fill = P2
        if "February 2027" in g["deadline"] or "September 2027" in g["deadline"]:
            priority = "3 - Later"
            fill = P3
        elif "31 August" in g["deadline"]:
            priority = "2 - This week"
            fill = P2
        rows.append(
            (
                [
                    priority,
                    "Before deadline",
                    "Team",
                    "Decide if we should apply (read the page, then yes or no)",
                    short_title(f"{num} - {g['title']}" if num else g["title"]),
                    g["deadline"],
                    "Not started",
                    g["url"],
                ],
                fill,
            )
        )

    # --- Done / already applied (from tracker) ---
    for t in tracker:
        if not is_applied(t["status"]):
            continue
        rows.append(
            (
                [
                    "Done",
                    "Already done",
                    "Team",
                    "Already applied / completed - keep only if funder replies",
                    short_title(t["title"]),
                    t["deadline"],
                    t["status"],
                    t["folder"] or t["url"],
                ],
                DONE,
            )
        )

    # Write rows
    for i, (values, fill) in enumerate(rows, 2):
        for col, val in enumerate(values, 1):
            cell = ws.cell(i, col, val)
            cell.fill = fill
            cell.border = THIN
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[i].height = 48

    widths = {
        "A": 14,
        "B": 28,
        "C": 18,
        "D": 48,
        "E": 42,
        "F": 28,
        "G": 16,
        "H": 55,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{len(rows) + 1}"

    wb.save(OUT)
    print(f"Wrote {OUT} ({len(rows)} rows, 1 sheet)")


if __name__ == "__main__":
    build()
