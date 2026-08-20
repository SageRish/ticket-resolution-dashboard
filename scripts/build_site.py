"""
Assemble the dashboard from src/ into two shippable outputs.

Usage:  python scripts/build_site.py

  index.html          full HTML document, loads data/tickets.js  -> GitHub Pages
  dist/artifact.html  single self-contained file, data inlined   -> Artifact / email

Both are generated from the same src/dashboard.html + src/dashboard.js, so the
two can never drift apart.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

FONTS = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
    '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?'
    "family=IBM+Plex+Mono:wght@400;500;600&"
    "family=IBM+Plex+Sans:wght@400;500;600&"
    "family=IBM+Plex+Serif:ital,wght@0,600;1,600&display=swap\">"
)

TITLE = "Breathe Villa Resolution Report"
DESC = ("Look up any villa and see how long its help desk tickets take to resolve, "
        "with the Excel formula behind every number.")


def main():
    markup = (SRC / "dashboard.html").read_text(encoding="utf-8")
    script = (SRC / "dashboard.js").read_text(encoding="utf-8")
    data_js = (ROOT / "data" / "tickets.js").read_text(encoding="utf-8")

    # ---- index.html : a real document, data loaded as a sibling file ----
    doc = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{TITLE}</title>
<meta name="description" content="{DESC}">
{FONTS}
</head>
<body>
{markup}
<script src="data/tickets.js"></script>
<script>
{script}
</script>
</body>
</html>
"""
    (ROOT / "index.html").write_text(doc, encoding="utf-8")

    # ---- dist/artifact.html : no wrapper tags, everything inlined ----
    art = f"""<title>{TITLE}</title>
{FONTS}
{markup}
<script>
{data_js}
</script>
<script>
{script}
</script>
"""
    (ROOT / "dist").mkdir(exist_ok=True)
    (ROOT / "dist" / "artifact.html").write_text(art, encoding="utf-8")

    for p in (ROOT / "index.html", ROOT / "dist" / "artifact.html"):
        print(f"  {p.relative_to(ROOT).as_posix():24} {p.stat().st_size/1024:7.0f} KB")


if __name__ == "__main__":
    main()
