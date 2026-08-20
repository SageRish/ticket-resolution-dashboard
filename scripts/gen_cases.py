"""
Build the audit case list: every formula shape the dashboard ships, paired with
the value the dashboard displays (taken from its own data payload).

audit_formulas.py then executes each formula against the raw workbook, so a
mismatch means a resident would not be able to reproduce what they were shown.

Usage:  python scripts/gen_cases.py > cases.json
"""
import json
import statistics
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
T = json.loads((ROOT / "data" / "tickets.json").read_text(encoding="utf-8"))["tickets"]

LAST = 3 + len(T)
R = lambda c: f"${c}$4:${c}${LAST}"
MR = R("W")
flats = sorted({t["flat"] for t in T}, key=lambda f: int(f.split("-")[1]))
AVGR = f"$AB$4:$AB${3 + len(flats)}"

mean = lambda a: statistics.fmean(a) if a else None
med = lambda a: statistics.median(a) if a else None
mins = lambda **kw: [t["mins"] for t in T if all(t[k] == v for k, v in kw.items())]

villa_avg = {f: mean(mins(flat=f)) for f in flats}
avgs = list(villa_avg.values())
by_avg = sorted(flats, key=lambda f: villa_avg[f])
counts = {f: len(mins(flat=f)) for f in flats}

sample = list(dict.fromkeys([
    by_avg[0], by_avg[-1], by_avg[len(by_avg) // 2],
    next(f for f in flats if counts[f] == 1),
    next(f for f in flats if any(t["flat"] == f and t["esc"] == "Level-1" for t in T)),
    "Villa-103", "Villa-151", "Villa-75", "Villa-138",
]))

cases = []
add = lambda label, code, expect: cases.append({"label": label, "code": code, "expect": expect})

for f in sample:
    for k in ("Level-1", "NA"):
        v = mins(flat=f, esc=k)
        if not v:
            continue
        add(f"{f} {k} avg", f'=AVERAGEIFS({MR},{R("G")},"{f}",{R("N")},"{k}")', mean(v))
        add(f"{f} {k} count", f'=COUNTIFS({R("G")},"{f}",{R("N")},"{k}")', len(v))
        add(f"{f} {k} median", f'=MEDIAN(IF(({R("G")}="{f}")*({R("N")}="{k}"),{MR}))', med(v))
    add(f"{f} overall avg", f'=AVERAGEIFS({MR},{R("G")},"{f}")', villa_avg[f])
    add(f"{f} ticket count", f'=COUNTIF({R("G")},"{f}")', counts[f])
    add(f"{f} open count", f'=COUNTIFS({R("G")},"{f}",{R("L")},"-")',
        len([t for t in T if t["flat"] == f and t["open"]]))
    me = villa_avg[f]
    add(f"{f} villas LONGER", f'=COUNTIF({AVGR},">"&AVERAGEIFS({MR},{R("G")},"{f}"))',
        len([a for a in avgs if a > me]))
    add(f"{f} villas SHORTER", f'=COUNTIF({AVGR},"<"&AVERAGEIFS({MR},{R("G")},"{f}"))',
        len([a for a in avgs if a < me]))
    add(f"{f} rank", f'=RANK(AVERAGEIFS({MR},{R("G")},"{f}"),{AVGR},1)',
        len([a for a in avgs if a < me]) + 1)

for k in ("Level-1", "NA"):
    add(f"society {k} avg", f'=AVERAGEIFS({MR},{R("N")},"{k}")', mean(mins(esc=k)))
add("society overall avg", f"=AVERAGE({MR})", mean([t["mins"] for t in T]))

for c in sorted({t["category"] for t in T}):
    add(f"society cat {c!r}", f'=AVERAGEIFS({MR},{R("D")},"{c}")', mean(mins(category=c)))
    v = mins(flat="Villa-151", category=c)
    if v:
        add(f"Villa-151 cat {c!r}", f'=AVERAGEIFS({MR},{R("G")},"Villa-151",{R("D")},"{c}")', mean(v))

json.dump([c for c in cases if c["expect"] is not None], sys.stdout, indent=1)
