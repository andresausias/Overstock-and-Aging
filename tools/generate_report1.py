"""
generate_report1.py
-------------------
Builds the 34-section Aging Inventory Analysis (MoM Comparison) HTML report.
Requires two filtered monthly CSV files (current + prior month) in .tmp/filtered/.

Output: .tmp/reports/aging_[prev]_[curr]_[year].html
        (self-contained, Chart.js v4.4.0 via CDN)

Usage:
    python tools/generate_report1.py

The script auto-detects the two most recent filtered CSVs by filename date prefix.
"""

import sys
import json
import re
import webbrowser
from pathlib import Path
from datetime import datetime
from typing import Optional, List

import pandas as pd
from jinja2 import Environment, FileSystemLoader

ROOT = Path(__file__).parent.parent
FILTERED_DIR = ROOT / ".tmp" / "filtered"
REPORTS_DIR = ROOT / ".tmp" / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR = Path(__file__).parent / "templates"
CONFIG_DIR = Path(__file__).parent / "config"

SILICONE_STYLES = set(
    json.loads((CONFIG_DIR / "silicone_styles.json").read_text())["styles"]
)

KNOWN_CATEGORIES = ["D2C", "WHOLESALE", "EMP", "BAD LOT"]
LOCATION_MARKET = {
    "JD NJ": "US", "JD ATL": "US", "JD CA": "US", "Lateral TJ": "US",
    "JD Canada": "Int", "JD UK": "Int", "JD AU": "Int",
    "JD SA": "Int", "CHE CN": "Int",
}
LOCATIONS = list(LOCATION_MARKET.keys())

MONTH_ABBR = {
    "01": "jan", "02": "feb", "03": "mar", "04": "apr",
    "05": "may", "06": "jun", "07": "jul", "08": "aug",
    "09": "sep", "10": "oct", "11": "nov", "12": "dec",
}

AGING_BUCKETS = ["0-90", "90-180", "180-270", "270-365", "Over 365"]


# ── Helpers ────────────────────────────────────────────────────────────────────

def fmt_m(v): return f"${v/1e6:.2f}M"
def fmt_k(v): return f"${v/1e3:.1f}K"
def fmt_val(v): return fmt_m(v) if abs(v) >= 1e6 else fmt_k(v)
def fmt_units(v): return f"{v/1e3:.1f}K" if abs(v) >= 1000 else str(int(v))
def is_silicone(style): return str(style).strip() in SILICONE_STYLES


def parse_date_from_filename(name: str) -> tuple[str, str, str]:
    """Extract YYYY, MM from YYYYMM_Aging*.csv filename."""
    m = re.match(r"(\d{4})(\d{2})_Aging", name, re.IGNORECASE)
    if not m:
        raise ValueError(f"Cannot parse date from filename: {name}")
    return m.group(1), m.group(2)


def load_filtered(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["Total Amount $"] = pd.to_numeric(df["Total Amount $"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Range TOTAL"] = df["Range TOTAL"].astype(str).str.strip()
    df["Category"] = df["Category"].astype(str).str.strip()
    df["Location"] = df["Location"].astype(str).str.strip()
    df["Style"] = df["Style"].astype(str).str.strip()
    df["Seller Product SKU"] = df["Seller Product SKU"].astype(str).str.strip()
    return df


def aged(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Range TOTAL"] == "Over 365"].copy()


def bar_color_absolute(style: str, value: float, is_wholesale: bool = False) -> str:
    """Return Chart.js rgba color for absolute value bar charts (S6, S7, S26-34)."""
    if is_silicone(style):
        return "rgba(156,39,176,0.6)"
    if value > 80000:
        return "rgba(244,67,54,0.6)"
    if value > 40000:
        return "rgba(255,152,0,0.6)"
    return "rgba(255,152,0,0.6)" if is_wholesale else "rgba(33,150,243,0.6)"


def compute_style_values(df: pd.DataFrame, categories: Optional[List[str]] = None) -> pd.DataFrame:
    """Aggregate aged value by style (optionally filtered by category)."""
    a = aged(df)
    if categories:
        a = a[a["Category"].isin(categories)]
    return (
        a.groupby("Style")["Total Amount $"]
        .sum()
        .reset_index()
        .rename(columns={"Total Amount $": "value"})
        .sort_values("value", ascending=False)
    )


def compute_movement(curr_styles: pd.DataFrame, prev_styles: pd.DataFrame) -> pd.DataFrame:
    """Compute MoM style movement (current - prior)."""
    merged = curr_styles.merge(prev_styles, on="Style", how="outer", suffixes=("_curr", "_prev"))
    merged = merged.fillna(0)
    merged["change"] = merged["value_curr"] - merged["value_prev"]
    merged["change_pct"] = merged.apply(
        lambda r: (r["change"] / r["value_prev"] * 100) if r["value_prev"] != 0 else 0, axis=1
    )
    return merged.sort_values("change", ascending=False)


def location_order_by_deterioration(curr_df: pd.DataFrame, prev_df: pd.DataFrame) -> list[str]:
    """Order locations by largest net deterioration (curr - prev aged value)."""
    curr_loc = aged(curr_df).groupby("Location")["Total Amount $"].sum()
    prev_loc = aged(prev_df).groupby("Location")["Total Amount $"].sum()
    diff = (curr_loc.reindex(LOCATIONS, fill_value=0) - prev_loc.reindex(LOCATIONS, fill_value=0))
    return diff.sort_values(ascending=False).index.tolist()


# ── Data computation ───────────────────────────────────────────────────────────

def compute_all(curr: pd.DataFrame, prev: pd.DataFrame) -> dict:
    """Compute all data needed for the 34-section report."""
    d = {}

    # Grand totals
    d["curr_aged_value"] = aged(curr)["Total Amount $"].sum()
    d["prev_aged_value"] = aged(prev)["Total Amount $"].sum()
    d["curr_aged_units"] = int(aged(curr)["Qty"].sum())
    d["prev_aged_units"] = int(aged(prev)["Qty"].sum())
    d["mom_value_change"] = d["curr_aged_value"] - d["prev_aged_value"]
    d["mom_pct_change"] = (
        d["mom_value_change"] / d["prev_aged_value"] * 100
        if d["prev_aged_value"] != 0 else 0
    )

    # US vs International aged values
    curr_us = aged(curr)[aged(curr)["Location"].map(LOCATION_MARKET).eq("US")]["Total Amount $"].sum()
    prev_us = aged(prev)[aged(prev)["Location"].map(LOCATION_MARKET).eq("US")]["Total Amount $"].sum()
    d["curr_us"] = curr_us
    d["prev_us"] = prev_us
    d["curr_int"] = d["curr_aged_value"] - curr_us
    d["prev_int"] = d["prev_aged_value"] - prev_us

    # Category breakdown (aged)
    d["curr_by_cat"] = {
        cat: aged(curr)[aged(curr)["Category"] == cat]["Total Amount $"].sum()
        for cat in KNOWN_CATEGORIES
    }
    d["prev_by_cat"] = {
        cat: aged(prev)[aged(prev)["Category"] == cat]["Total Amount $"].sum()
        for cat in KNOWN_CATEGORIES
    }

    # Aging bucket distribution (all buckets, all categories)
    d["curr_buckets"] = {
        b: curr[curr["Range TOTAL"] == b]["Total Amount $"].sum()
        for b in AGING_BUCKETS
    }
    d["prev_buckets"] = {
        b: prev[prev["Range TOTAL"] == b]["Total Amount $"].sum()
        for b in AGING_BUCKETS
    }

    # Top 15 styles — All categories (aged)
    d["top15_all"] = compute_style_values(curr).head(15)
    d["top15_wholesale"] = compute_style_values(curr, ["WHOLESALE"]).head(15)
    d["top15_d2c_group"] = compute_style_values(curr, ["D2C", "EMP", "BAD LOT"]).head(15)

    # MoM style movement
    curr_all_styles = compute_style_values(curr)
    prev_all_styles = compute_style_values(prev)
    movement = compute_movement(curr_all_styles, prev_all_styles)
    d["movement_all"] = movement
    d["top10_increases"] = movement.nlargest(10, "change")
    d["top10_decreases"] = movement.nsmallest(10, "change")

    # D2C group movement
    curr_d2c = compute_style_values(curr, ["D2C", "EMP", "BAD LOT"])
    prev_d2c = compute_style_values(prev, ["D2C", "EMP", "BAD LOT"])
    d["movement_d2c"] = compute_movement(curr_d2c, prev_d2c)

    # Wholesale movement
    curr_ws = compute_style_values(curr, ["WHOLESALE"])
    prev_ws = compute_style_values(prev, ["WHOLESALE"])
    d["movement_wholesale"] = compute_movement(curr_ws, prev_ws)

    # Top 25 increases / decreases (all styles)
    d["top25_increases"] = movement.nlargest(25, "change")
    d["top25_decreases"] = movement.nsmallest(25, "change")

    # Location heatmap
    d["location_heatmap"] = {}
    for loc in LOCATIONS:
        row = {}
        for bucket in AGING_BUCKETS:
            row[bucket] = curr[
                (curr["Location"] == loc) & (curr["Range TOTAL"] == bucket)
            ]["Total Amount $"].sum()
        row["aged_total"] = row["Over 365"]
        d["location_heatmap"][loc] = row

    # Location aged values (curr + prev) for S16
    d["curr_by_loc"] = {
        loc: aged(curr)[aged(curr)["Location"] == loc]["Total Amount $"].sum()
        for loc in LOCATIONS
    }
    d["prev_by_loc"] = {
        loc: aged(prev)[aged(prev)["Location"] == loc]["Total Amount $"].sum()
        for loc in LOCATIONS
    }
    d["loc_change"] = {
        loc: d["curr_by_loc"][loc] - d["prev_by_loc"][loc]
        for loc in LOCATIONS
    }
    d["loc_change_pct"] = {
        loc: (d["loc_change"][loc] / d["prev_by_loc"][loc] * 100
              if d["prev_by_loc"][loc] != 0 else 0)
        for loc in LOCATIONS
    }

    # Location order by deterioration
    d["loc_order"] = location_order_by_deterioration(curr, prev)

    # Per-location top styles (S17-34)
    d["per_loc_increases"] = {}
    d["per_loc_absolute"] = {}
    for loc in LOCATIONS:
        curr_loc_df = curr[curr["Location"] == loc]
        prev_loc_df = prev[prev["Location"] == loc]
        curr_loc_styles = compute_style_values(curr_loc_df)
        prev_loc_styles = compute_style_values(prev_loc_df)
        mv = compute_movement(curr_loc_styles, prev_loc_styles)
        d["per_loc_increases"][loc] = mv.nlargest(15, "change")
        d["per_loc_absolute"][loc] = curr_loc_styles.head(15)

    return d


# ── HTML rendering ─────────────────────────────────────────────────────────────

def render_report(d: dict, prev_label: str, curr_label: str, year: str) -> str:
    """Render the full 34-section HTML report from computed data."""
    env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))
    env.globals.update(
        fmt_m=fmt_m, fmt_k=fmt_k, fmt_val=fmt_val,
        fmt_units=fmt_units, is_silicone=is_silicone,
        bar_color_absolute=bar_color_absolute,
        zip=zip, enumerate=enumerate, abs=abs,
    )
    template = env.get_template("report1.html")
    cat_changes = [d["curr_by_cat"][c] - d["prev_by_cat"][c] for c in KNOWN_CATEGORIES]
    return template.render(
        d=d,
        prev_label=prev_label.upper(),
        curr_label=curr_label.upper(),
        year=year,
        locations=LOCATIONS,
        location_market=LOCATION_MARKET,
        categories=KNOWN_CATEGORIES,
        cat_changes=cat_changes,
        buckets=AGING_BUCKETS,
        SILICONE_STYLES=SILICONE_STYLES,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )


def main():
    # Find all filtered aging files sorted by date prefix
    filtered_files = sorted(FILTERED_DIR.glob("*_filtered.csv"))
    aging_files = sorted(
        [f for f in filtered_files if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)],
        key=lambda f: f.name[:6]
    )

    if len(aging_files) < 2:
        print(f"ERROR: Need at least 2 filtered aging files, found {len(aging_files)}.")
        print("Run filter_inventory.py first.")
        sys.exit(1)

    # Generate one MoM report for every consecutive pair
    pairs = list(zip(aging_files[:-1], aging_files[1:]))
    print(f"Found {len(aging_files)} filtered files → generating {len(pairs)} MoM report(s)\n")

    generated = []
    for prev_file, curr_file in pairs:
        prev_year, prev_month = parse_date_from_filename(prev_file.name)
        curr_year, curr_month = parse_date_from_filename(curr_file.name)
        prev_label = MONTH_ABBR[prev_month]
        curr_label = MONTH_ABBR[curr_month]

        out_name = f"aging_{prev_label}_{curr_label}_{curr_year}.html"
        out_path = REPORTS_DIR / out_name

        print(f"  {prev_label.upper()} → {curr_label.upper()} {curr_year}")
        prev_df = load_filtered(prev_file)
        curr_df  = load_filtered(curr_file)
        data = compute_all(curr_df, prev_df)
        html = render_report(data, prev_label, curr_label, curr_year)
        out_path.write_text(html, encoding="utf-8")
        generated.append(out_path)
        print(f"    Saved: {out_path.relative_to(ROOT)}")

    print(f"\n{len(generated)} report(s) saved.")
    print("Opening most recent in browser...")
    webbrowser.open(generated[-1].as_uri())


if __name__ == "__main__":
    main()
