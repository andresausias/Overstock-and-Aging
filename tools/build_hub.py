"""
build_hub.py
------------
Generates the hub portal and assembles a self-contained output/ folder.

Steps:
  1. Copy aging HTML reports   → output/reports/aging/
  2. Copy overstock HTML reports → output/reports/overstock/
  3. Copy PPTX slides          → output/slides/
  4. Render hub portal         → output/hub/index.html

Usage:
    python tools/build_hub.py

Output:
    output/
      reports/
        aging/       ← MoM + evacuation HTML reports
        overstock/   ← Rolling forecast overstock HTML reports
      slides/        ← PPTX + companion HTML
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
REPORTS_DIR   = ROOT / ".tmp" / "reports"
SLIDES_DIR    = ROOT / ".tmp" / "slides"
FILTERED_DIR  = ROOT / ".tmp" / "filtered"
TEMPLATES_DIR = Path(__file__).parent / "templates"
CONFIG_DIR    = Path(__file__).parent / "config"
AGING_INPUT_DIR = ROOT / "input" / "aging"

# ── Output folder (self-contained, shareable) ─────────────────────────────
OUTPUT_DIR              = ROOT / "output"
OUT_REPORTS_AGING_DIR   = OUTPUT_DIR / "reports" / "aging"
OUT_REPORTS_OVERSTOCK_DIR = OUTPUT_DIR / "reports" / "overstock"
OUT_SLIDES_DIR          = OUTPUT_DIR / "slides"
OUT_HUB_DIR             = OUTPUT_DIR / "hub"
for d in (OUT_REPORTS_AGING_DIR, OUT_REPORTS_OVERSTOCK_DIR, OUT_SLIDES_DIR, OUT_HUB_DIR):
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


def parse_aging_report_meta(filename: str) -> Optional[dict]:
    """Parse metadata from aging report filename."""
    # aging_projection_*.html  — check BEFORE MoM pattern to avoid false match
    m = re.match(r"aging_projection_(.+)\.html", filename, re.IGNORECASE)
    if m:
        slug = m.group(1)
        label = slug.replace("_", " ").upper()
        return {
            "type": "projection",
            "section": "aging",
            "filename": filename,
            "label": label,
            "date_key": datetime.now().strftime("%Y%m"),
            "subfolder": "aging",
        }
    # aging_jan_feb_2026.html
    m = re.match(r"aging_(\w+)_(\w+)_(\d{4})\.html", filename, re.IGNORECASE)
    if m:
        prev_abbr, curr_abbr, year = m.group(1), m.group(2), m.group(3)
        curr_month_num = MONTH_ABBR.get(curr_abbr.lower(), "01")
        return {
            "type": "mom",
            "section": "aging",
            "filename": filename,
            "label": f"{prev_abbr.upper()} vs {curr_abbr.upper()} {year}",
            "date_key": f"{year}{curr_month_num}",
            "year": year,
            "month": curr_abbr.upper(),
            "subfolder": "aging",
        }
    # aging_evacuation_analysis.html
    if filename == "aging_evacuation_analysis.html":
        return {
            "type": "evacuation",
            "section": "aging",
            "filename": filename,
            "label": "SKU Evacuation Analysis",
            "date_key": datetime.now().strftime("%Y%m"),
            "subfolder": "aging",
        }
    return None


def _fmt_period_slug(slug: str) -> str:
    """Convert 'mar26' → 'Mar 26', 'mar27' → 'Mar 27'."""
    m = re.match(r"([a-z]+)(\d{2})$", slug.lower())
    if m:
        return f"{m.group(1).capitalize()} {m.group(2)}"
    return slug.upper()


def parse_overstock_report_meta(filename: str) -> Optional[dict]:
    """Parse metadata from overstock report filename.
    Expected pattern: overstock_<start>_<end>_report.html or YYYYMMDD_overstock_*.html
    """
    # With date prefix: 20260411_overstock_mar26_mar27_report.html
    m = re.match(r"(\d{8})_overstock_(\w+)_(\w+)_report\.html", filename, re.IGNORECASE)
    if m:
        date_prefix, start, end = m.group(1), m.group(2), m.group(3)
        label = f"{_fmt_period_slug(start)} → {_fmt_period_slug(end)}"
        return {
            "type": "overstock_forecast",
            "section": "overstock",
            "filename": filename,
            "label": label,
            "date_key": date_prefix,
            "subfolder": "overstock",
        }
    # Without date prefix: overstock_mar26_mar27_report.html
    m = re.match(r"overstock_(\w+)_(\w+)_report\.html", filename, re.IGNORECASE)
    if m:
        start, end = m.group(1), m.group(2)
        label = f"{_fmt_period_slug(start)} → {_fmt_period_slug(end)}"
        return {
            "type": "overstock_forecast",
            "section": "overstock",
            "filename": filename,
            "label": label,
            "date_key": datetime.now().strftime("%Y%m"),
            "subfolder": "overstock",
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

    # ── Copy aging HTML reports → output/reports/aging/ ──────────────────────
    aging_reports = []
    overstock_reports = []

    if REPORTS_DIR.exists():
        for html_file in sorted(REPORTS_DIR.glob("*.html")):
            name = html_file.name

            # Try aging first
            meta = parse_aging_report_meta(name)
            if meta:
                out_name = f"{today_prefix}_{name}"
                shutil.copy2(html_file, OUT_REPORTS_AGING_DIR / out_name)
                meta["filename"] = out_name
                meta["original_filename"] = name
                aging_reports.append(meta)
                print(f"  Copied: output/reports/aging/{out_name}")
                continue

            # Try overstock
            meta = parse_overstock_report_meta(name)
            if meta:
                out_name = f"{today_prefix}_{name}"
                shutil.copy2(html_file, OUT_REPORTS_OVERSTOCK_DIR / out_name)
                meta["filename"] = out_name
                meta["original_filename"] = name
                overstock_reports.append(meta)
                print(f"  Copied: output/reports/overstock/{out_name}")

    # Also check .tmp/overstock/ if it exists
    overstock_tmp = ROOT / ".tmp" / "overstock"
    if overstock_tmp.exists():
        for html_file in sorted(overstock_tmp.glob("*.html")):
            meta = parse_overstock_report_meta(html_file.name)
            if not meta:
                # Accept any HTML in overstock tmp dir
                meta = {
                    "type": "overstock_forecast",
                    "section": "overstock",
                    "filename": html_file.name,
                    "label": html_file.stem.replace("_", " ").title(),
                    "date_key": datetime.now().strftime("%Y%m"),
                    "subfolder": "overstock",
                }
            out_name = f"{today_prefix}_{html_file.name}"
            shutil.copy2(html_file, OUT_REPORTS_OVERSTOCK_DIR / out_name)
            meta["filename"] = out_name
            meta["original_filename"] = html_file.name
            overstock_reports.append(meta)
            print(f"  Copied: output/reports/overstock/{out_name}")

    # Sort chronologically
    aging_reports.sort(key=lambda r: r.get("date_key", ""))
    overstock_reports.sort(key=lambda r: r.get("date_key", ""))

    # Combined list for backward compat
    all_reports = aging_reports + overstock_reports

    # ── Copy slides (PPTX + companion HTML) into output/slides/ ──────────────
    # Only the most recently modified PPTX is kept — no accumulation.
    slides = []
    slides_copied = 0
    if SLIDES_DIR.exists():
        all_pptx = sorted(SLIDES_DIR.glob("*.pptx"), key=lambda f: f.stat().st_mtime, reverse=True)
        if all_pptx:
            pptx = all_pptx[0]  # newest only
            out_pptx = f"{today_prefix}_{pptx.name}"
            shutil.copy2(pptx, OUT_SLIDES_DIR / out_pptx)
            print(f"  Copied: output/slides/{out_pptx}")
            slides_copied += 1

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

    # ── Load overstock KPIs ───────────────────────────────────────────────────
    # Priority: .tmp/overstock/overstock_kpis.json (written by generate_overstock_report.py)
    # Fallback: run_params.json (manual entry)
    overstock_kpi = None
    kpi_json_path = ROOT / ".tmp" / "overstock" / "overstock_kpis.json"
    run_params_path = CONFIG_DIR / "run_params.json"

    overstock_trajectory = None
    if kpi_json_path.exists():
        try:
            ok = json.loads(kpi_json_path.read_text())
            overstock_kpi = {
                "value": float(ok.get("value", 0)),
                "units": int(ok.get("units", 0)),
                "pct":   float(ok.get("pct_eom", 0)),
                "start_label": ok.get("start_label", ""),
                "end_label":   ok.get("end_label", ""),
            }
            if ok.get("traj_units") and ok.get("traj_value"):
                overstock_trajectory = {
                    "month_labels": ok.get("month_labels", []),
                    "traj_units":   ok.get("traj_units", []),
                    "traj_value":   ok.get("traj_value", []),
                }
            print(f"  Overstock KPIs from: {kpi_json_path.name}")
        except Exception as e:
            print(f"  (overstock_kpis.json skipped: {e})")

    if overstock_kpi is None and run_params_path.exists():
        try:
            rp = json.loads(run_params_path.read_text())
            ov_val   = float(rp.get("overstock_valuation", 0) or 0)
            ov_units = int(rp.get("overstock_units", 0) or 0)
            total_inv = float(rp.get("total_inventory_denominator", 0) or 0)
            ov_pct   = round(ov_val / total_inv * 100, 2) if total_inv else 0.0
            if ov_val > 0:
                overstock_kpi = {"value": ov_val, "units": ov_units, "pct": ov_pct}
                print("  Overstock KPIs from: run_params.json (fallback)")
        except Exception as e:
            print(f"  (overstock KPI skipped: {e})")

    # ── Find latest aging Excel in input/aging/ ───────────────────────────────
    latest_aging_excel = None
    if AGING_INPUT_DIR.exists():
        candidates = sorted(AGING_INPUT_DIR.glob("*.xlsx"), key=lambda f: f.name, reverse=True)
        if candidates:
            latest_aging_excel = candidates[0].name
            print(f"  Latest aging Excel: {latest_aging_excel}")

    # ── Write hub_data.json ───────────────────────────────────────────────────
    hub_data = {
        "aging_reports": aging_reports,
        "overstock_reports": overstock_reports,
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
        reports=all_reports,           # kept for backward compat in template
        aging_reports=aging_reports,
        overstock_reports=overstock_reports,
        slides=slides,
        timeline=timeline,
        latest=latest,
        overstock_kpi=overstock_kpi,
        overstock_trajectory=overstock_trajectory,
        latest_aging_excel=latest_aging_excel,
        aging_input_dir="../../input/aging",
        aging_reports_dir="../reports/aging",
        overstock_reports_dir="../reports/overstock",
        reports_dir="../reports/aging",  # backward compat
        slides_dir="../slides",
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )

    out_path = OUT_HUB_DIR / "index.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"\nHub saved:              output/hub/index.html")
    print(f"Aging reports indexed:  {len(aging_reports)}")
    print(f"Overstock reports:      {len(overstock_reports)}")
    print(f"Slides collected:       {slides_copied}")
    print(f"Timeline points:        {len(timeline)}")
    print("Opening hub in browser...")
    webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
