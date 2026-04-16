"""
generate_report2.py
-------------------
Builds the Aged SKU Evacuation Analysis HTML report (10 sections).
Requires one filtered monthly CSV (current) + weekly shipment data.

Output: .tmp/reports/aging_evacuation_analysis.html
        (self-contained, Chart.js v4.4.0 via CDN)

Usage:
    python tools/generate_report2.py

The weekly shipment file must be in .tmp/raw/ and have columns:
    Seller Product SKU (or SKU), W01, W02, ... WNn
"""

import sys
import json
import re
import webbrowser
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd
from jinja2 import Environment, FileSystemLoader

ROOT = Path(__file__).parent.parent
FILTERED_DIR   = ROOT / ".tmp" / "filtered"
PROCESSED_DIR  = ROOT / ".tmp" / "processed"
REPORTS_DIR = ROOT / ".tmp" / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR = Path(__file__).parent / "templates"
CONFIG_DIR = Path(__file__).parent / "config"

SILICONE_STYLES = set(
    json.loads((CONFIG_DIR / "silicone_styles.json").read_text())["styles"]
)

# Segment definitions
BAD_LOT_STYLES = {"91401", "51422"}
EMP_SKU_PREFIX = "EMP-"
STYLE_10210_EXTENDED_SIZES = {"M+", "L+", "XL+", "2XL+"}
STYLE_10210_REGULAR_SIZES = {"S", "M", "L", "XL", "2XL", "3XL", "4XL"}

RISK_LABELS = {
    "on_track": ("On Track", "#D1FAE5", "#065F46"),
    "moderate": ("Moderate", "#FEF3C7", "#92400E"),
    "slow":     ("Slow",     "#FEE2E2", "#991B1B"),
    "stuck":    ("Stuck",    "#FEE2E2", "#991B1B"),
    "no_sales": ("No Sales", "#FEE2E2", "#991B1B"),
}


def weeks_to_risk(weeks: float) -> str:
    if weeks == 9999:
        return "no_sales"
    if weeks <= 8:
        return "on_track"
    if weeks <= 20:
        return "moderate"
    if weeks <= 52:
        return "slow"
    return "stuck"


def fmt_val(v: float) -> str:
    return f"${v/1e6:.2f}M" if abs(v) >= 1e6 else f"${v/1e3:.1f}K"


def fmt_units(v) -> str:
    return f"{v/1e3:.1f}K" if abs(v) >= 1000 else str(int(v))


def load_weekly_shipments(processed_dir: Path) -> Optional[pd.DataFrame]:
    """Load per-SKU weekly shipments from .tmp/processed/weekly_by_sku.csv."""
    path = processed_dir / "weekly_by_sku.csv"
    if not path.exists():
        print("  WARNING: .tmp/processed/weekly_by_sku.csv not found.")
        print("           Run: python tools/process_shipments.py first.")
        print("           Evacuation weeks will be N/A.")
        return None

    print(f"  Using: {path.name}")
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    # Normalize SKU column
    sku_col = next(
        (c for c in df.columns if "seller" in c.lower() or "sku" in c.lower()),
        df.columns[0]
    )
    df = df.rename(columns={sku_col: "SKU"})

    # Parse week columns (WK1, WK2, ... or W01, W02, ...)
    wk_cols = [c for c in df.columns if re.match(r"^WK?\d+$", c, re.IGNORECASE)]
    for c in wk_cols:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    df["shipped_ytd"] = df[wk_cols].sum(axis=1)
    df["num_weeks"] = len(wk_cols)
    df["weekly_avg"] = df["shipped_ytd"] / df["num_weeks"]
    df["wk_cols"] = [wk_cols] * len(df)
    df["weekly_detail"] = df[wk_cols].values.tolist()
    return df


def build_sku_records(aged: pd.DataFrame, shipments: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Merge aged inventory with shipment data to compute evacuation metrics."""
    records = aged.copy()
    records["SKU"] = records["Seller Product SKU"]

    if shipments is not None:
        records = records.merge(
            shipments[["SKU", "shipped_ytd", "weekly_avg", "num_weeks", "weekly_detail"]],
            on="SKU", how="left"
        )
    else:
        records["shipped_ytd"] = 0
        records["weekly_avg"] = 0
        records["num_weeks"] = 0
        records["weekly_detail"] = [[] for _ in range(len(records))]

    records = records.fillna({"shipped_ytd": 0, "weekly_avg": 0})

    records["weeks_to_evac"] = records.apply(
        lambda r: round(r["Total Aged"] / r["weekly_avg"], 1)
        if r.get("weekly_avg", 0) > 0 else 9999,
        axis=1
    )
    records["risk"] = records["weeks_to_evac"].apply(weeks_to_risk)
    return records


def segment_skus(df: pd.DataFrame) -> dict:
    """Split SKUs into the 6 report segments."""
    aged_df = df[df["Range TOTAL"].isin(["Over 365", "270-365"])].copy()
    aged_df["Total Aged"] = aged_df.groupby("Seller Product SKU")["Qty"].transform("sum")
    aged_df["Over 365 Qty"] = aged_df[aged_df["Range TOTAL"] == "Over 365"]["Qty"]
    aged_df["270-365 Qty"] = aged_df[aged_df["Range TOTAL"] == "270-365"]["Qty"]

    # Deduplicate to SKU level
    agg = (
        aged_df.groupby("Seller Product SKU")
        .agg({
            "Style": "first", "Color": "first", "Size": "first",
            "Total Amount $": "sum", "Qty": "sum",
            "Category": "first", "MKT Strategy": "first",
        })
        .reset_index()
        .rename(columns={"Qty": "Total Aged", "Total Amount $": "Valuation"})
    )

    style = agg["Style"].astype(str).str.strip()
    sku = agg["Seller Product SKU"].astype(str)

    segments = {
        "10210_extended": agg[style.eq("10210") & agg["Size"].isin(STYLE_10210_EXTENDED_SIZES)],
        "10210_regular":  agg[style.eq("10210") & ~agg["Size"].isin(STYLE_10210_EXTENDED_SIZES)],
        "10400":          agg[style.eq("10400")],
        "bad_lot":        agg[style.isin(BAD_LOT_STYLES) | (agg["Category"] == "BAD LOT")],
        "emp_brand":      agg[style.eq("62001") & sku.str.upper().str.startswith("EMP-")],
        "discontinued":   agg[
            agg["MKT Strategy"].astype(str).str.lower().str.contains("discontinued")
            & ~style.isin({"10210", "10400", "62001"} | BAD_LOT_STYLES)
        ],
    }
    return segments


def summarize_segment(df: pd.DataFrame) -> dict:
    return {
        "total_aged_units": int(df["Total Aged"].sum()),
        "total_valuation": df["Valuation"].sum(),
        "total_shipped": int(df.get("shipped_ytd", pd.Series([0] * len(df))).sum()),
        "stuck_count": int((df.get("weeks_to_evac", pd.Series([0] * len(df))) > 52).sum()),
    }


def main():
    # Load most recent filtered aging file
    filtered_files = sorted([
        f for f in FILTERED_DIR.glob("*_filtered.csv")
        if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)
    ])
    if not filtered_files:
        print("ERROR: No filtered aging files found. Run filter_inventory.py first.")
        sys.exit(1)

    curr_file = filtered_files[-1]
    print(f"Using: {curr_file.name}")

    df = pd.read_csv(curr_file)
    df["Total Amount $"] = pd.to_numeric(df["Total Amount $"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    for col in ["Seller Product SKU", "Style", "Category", "Range TOTAL", "MKT Strategy", "Size", "Color"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    print("Loading weekly shipments...")
    shipments = load_weekly_shipments(PROCESSED_DIR)

    print("Segmenting SKUs...")
    segments = segment_skus(df)

    # Merge shipment data into each segment
    segments = {k: build_sku_records(v, shipments) for k, v in segments.items()}

    # Compute summaries
    summaries = {k: summarize_segment(v) for k, v in segments.items()}

    # Load template and render
    env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))
    env.globals.update(
        fmt_val=fmt_val, fmt_units=fmt_units, is_silicone=lambda s: str(s) in SILICONE_STYLES,
        weeks_to_risk=weeks_to_risk, RISK_LABELS=RISK_LABELS,
        enumerate=enumerate, zip=zip, abs=abs,
    )
    template = env.get_template("report2.html")
    import re as _re
    wk_cols = [c for c in (shipments.columns.tolist() if shipments is not None else [])
               if _re.match(r"^WK?\d+$", c, _re.IGNORECASE)]
    html = template.render(
        segments=segments,
        summaries=summaries,
        shipments=shipments,
        wk_cols=wk_cols,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        silicone_styles=SILICONE_STYLES,
    )

    out_path = REPORTS_DIR / "aging_evacuation_analysis.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"\nReport saved: {out_path.relative_to(ROOT)}")
    print("Opening in browser...")
    webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
