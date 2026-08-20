# Breathe Villa Resolution Report

A web dashboard for the Breathe Villa Association help desk. A resident looks up
their villa and sees how long their tickets take to resolve, split by escalation
level, and how that compares with the society and with every other villa.

The point of the thing is **transparency**: every number on the page ships with
the Excel formula that reproduces it against the original export. Nothing is
"trust me" — open the spreadsheet, paste the formula, get the same number.

The headline measure is the **median**, not the mean. Society-wide the mean is
3.6x the median and 78% of tickets beat it, so it describes almost no real
ticket; ranking by mean also moves 99 of the 175 villas by more than twenty
places versus ranking by median. The mean and the slowest ticket stay visible
next to every median so the tail is never hidden. The dashboard argues this
case itself in its "Why the median leads" panel.

## Hosting

`index.html` is a static page with no build step at view time. Serve the repo
root from any static host:

- **GitHub Pages** — Settings → Pages → deploy from `main`, folder `/ (root)`.
- **Anything else** — copy `index.html` and the `data/` folder.

`dist/artifact.html` is the same dashboard as a single self-contained file
(data inlined), for emailing or dropping somewhere that can't serve a folder.

Locally:

```bash
python -m http.server 8765
```

## Rebuilding from a new export

```bash
python scripts/build_data.py path/to/ticket_data.xlsx && python scripts/build_site.py
```

`build_data.py` reads the raw workbook and writes `data/tickets.js` (loaded by
the page) plus `data/tickets.json` (readable, for the audit). `build_site.py`
assembles `src/dashboard.html` + `src/dashboard.js` into both outputs, so the
hosted page and the single-file version can never drift apart.

## Verifying the formulas

`audit_formulas.py` parses the Excel formulas the dashboard publishes, executes
them against the raw workbook, and compares each result to the value the
dashboard displays:

```bash
python scripts/gen_cases.py > cases.json && python scripts/audit_formulas.py path/to/ticket_data.xlsx cases.json
```

Current status: **236 / 236 formulas reproduce the displayed value exactly.**

## How the data is treated

The export is not clean, and the decisions below are all surfaced in the
dashboard's own "How every number on this page is built" panel.

| Thing | Decision |
|---|---|
| Header position | Row 1 is the report title, row 2 blank, headers on **row 3**, data in **rows 4:1180** |
| `Resolution Tat` | Stored as text (`2 day(s): 3 hour(s): 5 minute(s)`); converted to minutes by helper column `W` |
| Open tickets | 30 tickets never closed (`Closed Time` = `-`). Their exported TAT is stale, so they are aged from creation to a **fixed cutoff of 2026-08-20 09:30** and included in every average |
| Escalation levels | `Level-1` → **Urgent**, `NA` → **Usual**. No other value occurs |
| `Category` | **Not trimmed.** The sheet stores `"Gardening "` with a trailing space; an Excel criterion must match it exactly or `AVERAGEIFS` finds nothing. The UI trims for display only |
| Dropped columns | `Sub Category`, `Assignee`, `On Hold Time`, `Assignment Tat`, `Expected Turnaround Time(Ett)`, `Resolution Tat Ratio (Rtat/Ett)` — each holds one identical value on all 1177 rows |
| Headline measure | **Median**, computed with `MEDIAN(IF(...))` since Excel has no `MEDIANIFS` — an array formula (Ctrl+Shift+Enter before Excel 365). Rankings use the median too |
| Rank ties | The displayed rank mirrors Excel's `RANK`, which gives tied values the same rank, so the badge always agrees with the published formula |
| Averages | Plain arithmetic mean (`AVERAGEIFS`), no weighting, no outlier removal, shown beside every median |
| Trend chart | Tickets grouped by the month or week they were **raised**. Buckets where 10%+ of tickets are still open are drawn dashed and shaded: an open ticket can only be measured to the cutoff, so those points are a floor, not a result |

## Layout

```
index.html              generated — the hosted page
dist/artifact.html      generated — single self-contained file
data/tickets.js         generated — data payload the page loads
data/tickets.json       generated — same data, readable
src/dashboard.html      source: styles + markup
src/dashboard.js        source: all computation, rendering and formula building
scripts/build_data.py   xlsx  -> data/
scripts/build_site.py   src/  -> index.html + dist/
scripts/gen_cases.py    builds the audit case list
scripts/audit_formulas.py  executes published formulas against the raw workbook
```

Requires `openpyxl` for the build scripts. The page itself has no dependencies.
