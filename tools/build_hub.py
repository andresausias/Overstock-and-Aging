"""
build_hub.py
------------
Generates the hub portal and assembles a self-contained output/ folder.

Steps:
  1. Copy all HTML reports from .tmp/reports/ → output/reports/
  2. Copy all PPTX slides from .tmp/slides/   → output/slides/
  3. Render hub portal → output/hub/index.html
     (iframe uses relative path ../reports/<file> — works offline/shareable)

Usage:
    python tools/build_hub.py

Output:
    output/
      reports/   ← HTML report copies
      slides/    ← PPTX slide copies
      hub/
        index.html      Portal (open this)
        hub_data.json   Data layer for charts
"""

import json
import re
import shutil
import webbrowser
from pathlib import Path
from datetime import datetime
from typing import Optional, List

import pandas as pd
from jinja2 import Environment, FileSystemLoader

ROOT = Path(__file__).parent.parent
REPORTS_DIR  = ROOT / ".tmp" / "reports"
SLIDES_DIR   = ROOT / ".tmp" / "slides"
FILTERED_DIR = ROOT / ".tmp" / "filtered"
TEMPLATES_DIR = Path(__file__).parent / "templates"

# ── Output folder (self-contained, shareable) ─────────────────────────────
OUTPUT_DIR         = ROOT / "output"
OUT_REPORTS_DIR    = OUTPUT_DIR / "reports"
OUT_SLIDES_DIR     = OUTPUT_DIR / "slides"
OUT_HUB_DIR        = OUTPUT_DIR / "hub"
for d in (OUT_REPORTS_DIR, OUT_SLIDES_DIR, OUT_HUB_DIR):
    d.mkdir(parents=True, exist_ok=True)

MONTH_NAMES = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December",
}
MONTH_ABBR = {v[:3].lower(): k for k, v in MONTH_NAMES.items()}

TOTAL_INVENTORY_DENOMINATOR = 50_935_732  # Update in run_params.json if changed


def extract_kpis_from_filtered(path: Path) -> dict:
    """Extract aged KPIs (Over 365) from a filtered CSV."""
    try:
        df = pd.read_csv(path)
        df["Total Amount $"] = pd.to_numeric(df["Total Amount $"], errors="coerce").fillna(0)
        df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
        aged = df[df["Range TOTAL"].astype(str).str.strip() == "Over 365"]
        val = aged["Total Amount $"].sum()
        units = int(aged["Qty"].sum())
        pct = val / TOTAL_INVENTORY_DENOMINATOR * 100 if TOTAL_INVENTORY_DENOMINATOR else 0
        return {"aged_value": round(val, 2), "aged_units": units, "aged_pct": round(pct, 2)}
    except Exception as e:
        return {"aged_value": 0, "aged_units": 0, "aged_pct": 0, "error": str(e)}


def parse_report_meta(filename: str) -> Optional[dict]:
    """Parse metadata from report filename."""
    # aging_jan_feb_2026.html
    m = re.match(r"aging_(\w+)_(\w+)_(\d{4})\.html", filename, re.IGNORECASE)
    if m:
        prev_abbr, curr_abbr, year = m.group(1), m.group(2), m.group(3)
        curr_month_num = MONTH_ABBR.get(curr_abbr.lower(), "01")
        return {
            "type": "mom",
            "filename": filename,
            "label": f"{prev_abbr.upper()} vs {curr_abbr.upper()} {year}",
            "date_key": f"{year}{curr_month_num}",
            "year": year,
            "month": curr_abbr.upper(),
        }
    # aging_evacuation_analysis.html
    if filename == "aging_evacuation_analysis.html":
        return {
            "type": "evacuation",
            "filename": filename,
            "label": "SKU Evacuation Analysis",
            "date_key": datetime.now().strftime("%Y%m"),
        }
    return None


def build_timeline(filtered_files: List[Path]) -> List[dict]:
    """Build KPI timeline from all filtered CSVs."""
    timeline = []
    for path in sorted(filtered_files):
        m = re.match(r"(\d{4})(\d{2})_Aging", path.name, re.IGNORECASE)
        if not m:
            continue
        year, month = m.group(1), m.group(2)
        kpis = extract_kpis_from_filtered(path)
        timeline.append({
            "date_key": f"{year}{month}",
            "label": f"{MONTH_NAMES.get(month, month)[:3]} {year}",
            **kpis,
        })
    return timeline


def main():
    today_prefix = datetime.now().strftime("%Y%m%d")

    # ── Copy HTML reports into output/reports/ with date prefix ──────────────
    reports = []
    if REPORTS_DIR.exists():
        for html_file in sorted(REPORTS_DIR.glob("*.html")):
            meta = parse_report_meta(html_file.name)
            if not meta:
                continue
            out_name = f"{today_prefix}_{html_file.name}"
            shutil.copy2(html_file, OUT_REPORTS_DIR / out_name)
            meta["filename"] = out_name      # hub uses this for iframe src
            meta["original_filename"] = html_file.name
            reports.append(meta)
            print(f"  Copied: output/reports/{out_name}")

    # Sort chronologically oldest → newest (loop.last in template = most recent)
    reports.sort(key=lambda r: r.get("date_key", ""))

    # ── Copy slides (PPTX + companion HTML) into output/slides/ ──────────────
    slides = []          # [{pptx, html, label}] for hub template
    slides_copied = 0
    if SLIDES_DIR.exists():
        for pptx in sorted(SLIDES_DIR.glob("*.pptx")):
            out_pptx = f"{today_prefix}_{pptx.name}"
            shutil.copy2(pptx, OUT_SLIDES_DIR / out_pptx)
            print(f"  Copied: output/slides/{out_pptx}")
            slides_copied += 1

            # Look for companion HTML (same stem, .html extension)
            html_src = pptx.with_suffix(".html")
            out_html = None
            if html_src.exists():
                out_html = f"{today_prefix}_{html_src.name}"
                shutil.copy2(html_src, OUT_SLIDES_DIR / out_html)
                print(f"  Copied: output/slides/{out_html}")

            slides.append({
                "label": pptx.stem.replace("_", " ").replace("weekly performance ", "").title(),
                "pptx":  out_pptx,
                "html":  out_html,
            })

    # ── Build KPI timeline from filtered CSVs ────────────────────────────────
    filtered_files = list(FILTERED_DIR.glob("*_filtered.csv")) if FILTERED_DIR.exists() else []
    timeline = build_timeline(filtered_files)
    latest = timeline[-1] if timeline else {}

    # ── Write hub_data.json ───────────────────────────────────────────────────
    hub_data = {
        "reports": reports,
        "slides":  slides,
        "timeline": timeline,
        "latest": latest,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    (OUT_HUB_DIR / "hub_data.json").write_text(json.dumps(hub_data, indent=2))

    # ── Render hub → output/hub/index.html ───────────────────────────────────
    env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))
    template = env.get_template("hub.html")
    html = template.render(
        reports=reports,
        slides=slides,
        timeline=timeline,
        latest=latest,
        reports_dir="../reports",
        slides_dir="../slides",
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )

    out_path = OUT_HUB_DIR / "index.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"\nHub saved:        output/hub/index.html")
    print(f"Reports indexed:  {len(reports)}")
    print(f"Slides collected: {slides_copied}")
    print(f"Timeline points:  {len(timeline)}")
    print("Opening hub in browser...")
    webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
