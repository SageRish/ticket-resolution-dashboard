"""
Build the dashboard data payload from the raw Help Desk export.

Usage:  python scripts/build_data.py [path/to/ticket_data.xlsx]

Reads the raw xlsx, strips non-informative rows/columns, normalises dates and
turnaround times, and writes data/tickets.js (a plain JS assignment so the
dashboard works from file:// as well as from a web host).
"""
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SRC = Path.home() / "Downloads" / "ticket_data.xlsx"

# Header lives on row 3; rows 1-2 are the report title banner.
HEADER_ROW = 3

# Columns carrying exactly one value across every row -> no information.
DEAD_COLUMNS = {
    "Sub Category",              # always "None"
    "Assignee",                  # always "None"
    "On Hold Time",              # always 0
    "Assignment Tat",            # always 0
    "Expected Turnaround Time(Ett)",   # always "NA"
    "Resolution Tat Ratio (Rtat/Ett)", # always "NA"
}

# Statuses that mean the ticket was actually finished. Anything else is still
# open: its Resolution Tat column is stale, so we age it against a fixed cutoff.
RESOLVED_STATUSES = {"Closed", "Resolved"}

# Open tickets are measured as (CUTOFF - Created Date). Fixed, not "now", so the
# dashboard and the Excel formulas always agree on the same number.
CUTOFF = datetime(2026, 8, 20, 9, 30)

ESCALATION_LABELS = {"Level-1": "Urgent", "NA": "Usual"}

TAT_RE = re.compile(r"^(\d+) day\(s\): (\d+) hour\(s\): (\d+) minute\(s\)$")


def parse_tat(value):
    """'2 day(s): 3 hour(s): 5 minute(s)' -> minutes as int."""
    m = TAT_RE.match(str(value).strip())
    if not m:
        return None
    d, h, mi = (int(g) for g in m.groups())
    return d * 1440 + h * 60 + mi


def parse_dt(value):
    """'20-08-26, 08:56 AM' (DD-MM-YY) -> ISO 8601 string, or None."""
    s = str(value).strip()
    if not s or s in {"-", "None", "NA"}:
        return None
    try:
        return datetime.strptime(s, "%d-%m-%y, %I:%M %p").isoformat(timespec="minutes")
    except ValueError:
        return None


def raw_category(value):
    """Category exactly as stored, so the published Excel criteria still match."""
    if value is None:
        return "Other"
    s = str(value)
    return "Other" if s.strip() in {"", "-", "None", "NA"} else s


def clean(value):
    """Normalise the export's several spellings of 'empty'."""
    if value is None:
        return None
    s = str(value).strip()
    return None if s in {"", "-", "None", "NA"} else s


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SRC
    if not src.exists():
        sys.exit(f"Source workbook not found: {src}")

    ws = openpyxl.load_workbook(src, data_only=True)["Sheet1"]
    rows = list(ws.iter_rows(values_only=True))
    header = [h for h in rows[HEADER_ROW - 1] if h is not None]

    raw = [dict(zip(header, r)) for r in rows[HEADER_ROW:] if r[0] not in (None, "")]

    dropped_dead = sorted(c for c in DEAD_COLUMNS if c in header)
    tickets, skipped = [], 0

    for r in raw:
        flat = clean(r.get("Flat"))
        created = parse_dt(r.get("Created Date"))
        if not flat or not created:
            skipped += 1
            continue

        status = clean(r.get("Status")) or "Unknown"
        closed = parse_dt(r.get("Closed Time"))
        resolved = status in RESOLVED_STATUSES and closed is not None
        rating = clean(r.get("Rating"))

        if resolved:
            # Finished ticket: trust the exported Resolution Tat.
            minutes = parse_tat(r.get("Resolution Tat (Rtat)"))
        else:
            # Still open: age it from creation to the fixed cutoff.
            delta = CUTOFF - datetime.fromisoformat(created)
            minutes = round(delta.total_seconds() / 60)

        tickets.append({
            "id": int(r["Id"]),
            "flat": flat,
            "created": created,
            "closed": closed,
            # Deliberately NOT trimmed: the sheet stores "Gardening " with a
            # trailing space, and an Excel criterion has to match the raw value
            # or AVERAGEIFS finds none of those rows. The UI trims for display.
            "category": raw_category(r.get("Category")),
            "type": clean(r.get("Type")) or "personal",
            "subject": clean(r.get("Subject")) or "(no subject)",
            "by": clean(r.get("Created By")) or "Unknown",
            "status": status,
            # Raw escalation value kept as-is; the UI maps it to Urgent/Usual.
            "esc": clean(r.get("Escalation Level")) or "NA",
            # Elapsed minutes: resolution time if closed, else age at CUTOFF.
            "mins": minutes,
            "open": not resolved,
            "reopened": int(r.get("Reopen Count") or 0),
            "rating": int(rating) if rating and rating.isdigit() else None,
        })

    tickets.sort(key=lambda t: t["created"], reverse=True)

    dates = [t["created"] for t in tickets]
    meta = {
        "society": str(rows[0][0]).strip(),
        "generated": datetime.now().isoformat(timespec="minutes"),
        "source": src.name,
        "rangeStart": min(dates),
        "rangeEnd": max(dates),
        "total": len(tickets),
        "resolved": sum(1 for t in tickets if not t["open"]),
        "openCount": sum(1 for t in tickets if t["open"]),
        "cutoff": CUTOFF.isoformat(timespec="minutes"),
        "flats": len({t["flat"] for t in tickets}),
        "escalationLabels": ESCALATION_LABELS,
        "droppedColumns": dropped_dead,
        "droppedRows": skipped,
    }

    out = ROOT / "data" / "tickets.js"
    out.parent.mkdir(exist_ok=True)
    payload = json.dumps({"meta": meta, "tickets": tickets}, separators=(",", ":"))
    out.write_text(f"window.TICKET_DATA = {payload};\n", encoding="utf-8")

    (ROOT / "data" / "tickets.json").write_text(
        json.dumps({"meta": meta, "tickets": tickets}, indent=1), encoding="utf-8")

    print(f"wrote {out.relative_to(ROOT)}  ({out.stat().st_size/1024:.0f} KB)")
    print(f"  tickets kept : {len(tickets)}   skipped rows: {skipped}")
    print(f"  resolved     : {meta['resolved']}  still open: {meta['openCount']} (aged to {CUTOFF:%d %b %Y %H:%M})")
    print(f"  flats        : {meta['flats']}")
    print(f"  range        : {meta['rangeStart'][:10]} -> {meta['rangeEnd'][:10]}")
    print(f"  dropped cols : {', '.join(dropped_dead)}")
    print("  escalation   :", dict(Counter(t["esc"] for t in tickets)))


if __name__ == "__main__":
    main()
