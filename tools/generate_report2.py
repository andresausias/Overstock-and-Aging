"""
generate_report2.py
-------------------
Builds the Aged SKU Evacuation Analysis HTML report (7 sections).

Segmentation uses SKU_Segment.xlsx as the authoritative source — same mapping
as the shipments pipeline. Aged inventory scope: Over 365 only, D2C + EMP + BAD LOT.

Output: .tmp/reports/aging_evacuation_analysis.html
        (self-contained, Chart.js v4.4.0 via CDN)

Usage:
    python tools/generate_report2.py

Requires:
    .tmp/filtered/*_filtered.csv   — latest filtered aging CSV
    .tmp/processed/weekly_by_sku.csv — per-SKU weekly shipments
    input/reference/SKU_Segment.xlsx — SKU→segment mapping
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

ROOT          = Path(__file__).parent.parent
FILTERED_DIR  = ROOT / ".tmp" / "filtered"
PROCESSED_DIR = ROOT / ".tmp" / "processed"
REPORTS_DIR   = ROOT / ".tmp" / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR = Path(__file__).parent / "templates"
CONFIG_DIR    = Path(__file__).parent / "config"
REFERENCE_DIR = ROOT / "input" / "reference"

SILICONE_STYLES = set(
    json.loads((CONFIG_DIR / "silicone_styles.json").read_text())["styles"]
)

# SKU_Segment.xlsx Segment values → internal segment keys used in this report
SEGMENT_MAP = {
    "10210 EXTENDED (+ SIZES)": "ext",
    "10210 REGULAR NEW LABS":   "new_labs",
    "10210 REGULAR OLD LABS":   "old_labs",
    "10400 SILICONE BAND":      "sil",
    "BAD LOT":                  "bad_lot",
    "EMP BRAND":                "emp_brand",
}

SEGMENT_ORDER = ["ext", "new_labs", "old_labs", "sil", "bad_lot", "emp_brand"]

RISK_LABELS = {
    "on_track": ("On Track", "#D1FAE5", "#065F46"),
    "moderate": ("Moderate", "#FEF3C7", "#92400E"),
    "slow":     ("Slow",     "#FEE2E2", "#991B1B"),
    "stuck":    ("Stuck",    "#FEE2E2", "#991B1B"),
    "no_sales": ("No Sales", "#FEE2E2", "#991B1B"),
}

D2C_CATEGORIES = {"D2C", "EMP", "BAD LOT"}


def weeks_to_risk(weeks: float) -> str:
    if weeks == 9999:  return "no_sales"
    if weeks <= 8:     return "on_track"
    if weeks <= 20:    return "moderate"
    if weeks <= 52:    return "slow"
    return "stuck"


def fmt_val(v: float) -> str:
    return f"${v/1e6:.2f}M" if abs(v) >= 1e6 else f"${v/1e3:.1f}K"


def fmt_units(v) -> str:
    return f"{v/1e3:.1f}K" if abs(v) >= 1000 else str(int(v))


def load_sku_segment(reference_dir: Path) -> dict:
    """Load SKU → internal segment key mapping from SKU_Segment.xlsx."""
    candidates = list(reference_dir.glob("SKU_Segment*"))
    if not candidates:
        print("  WARNING: SKU_Segment.xlsx not found — segmentation will be empty.")
        return {}

    df = pd.read_excel(candidates[0])
    # Strip column names (handles 'SKU ' with trailing space)
    df.columns = df.columns.str.strip()

    sku_col = next((c for c in df.columns if c.strip().upper() == "SKU"), df.columns[0])
    seg_col = next((c for c in df.columns if "segment" in c.lower()), df.columns[1])

    df[sku_col] = df[sku_col].astype(str).str.strip()
    df[seg_col] = df[seg_col].astype(str).str.strip().str.upper()

    # Deduplicate — each SKU maps to exactly one segment
    df = df[[sku_col, seg_col]].drop_duplicates(subset=[sku_col])

    mapping = {}
    for _, row in df.iterrows():
        seg_key = SEGMENT_MAP.get(row[seg_col])
        if seg_key:
            mapping[row[sku_col]] = seg_key

    print(f"  Loaded {len(mapping):,} SKU→segment mappings from {candidates[0].name}")
    return mapping


def load_weekly_shipments(processed_dir: Path) -> Optional[pd.DataFrame]:
    """Load per-SKU weekly shipments from .tmp/processed/weekly_by_sku.csv."""
    path = processed_dir / "weekly_by_sku.csv"
    if not path.exists():
        print("  WARNING: .tmp/processed/weekly_by_sku.csv not found.")
        print("           Run: python tools/process_shipments.py first.")
        return None

    print(f"  Using: {path.name}")
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    sku_col = next(
        (c for c in df.columns if "seller" in c.lower() or "sku" in c.lower()),
        df.columns[0]
    )
    df = df.rename(columns={sku_col: "SKU"})

    wk_cols = [c for c in df.columns if re.match(r"^WK?\d+$", c, re.IGNORECASE)]
    for c in wk_cols:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    df["shipped_ytd"]   = df[wk_cols].sum(axis=1)
    df["num_weeks"]     = len(wk_cols)
    df["weekly_avg"]    = df["shipped_ytd"] / df["num_weeks"]
    df["wk_cols"]       = [wk_cols] * len(df)
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
        records["shipped_ytd"]   = 0
        records["weekly_avg"]    = 0
        records["num_weeks"]     = 0
        records["weekly_detail"] = [[] for _ in range(len(records))]

    records = records.fillna({"shipped_ytd": 0, "weekly_avg": 0})

    records["weeks_to_evac"] = records.apply(
        lambda r: round(r["Total Aged"] / r["weekly_avg"], 1)
        if r.get("weekly_avg", 0) > 0 else 9999,
        axis=1
    )
    records["risk"] = records["weeks_to_evac"].apply(weeks_to_risk)
    return records


def segment_skus(df: pd.DataFrame, sku_map: dict) -> dict:
    """
    Split aged SKUs into 6 segments using SKU_Segment.xlsx mapping.

    Filters:
      - Range TOTAL == 'Over 365'
      - Category in {D2C, EMP, BAD LOT}
      - SKU present in sku_map (i.e. one of the 6 tracked segments)
    """
    # Apply scope filters
    mask = (
        (df["Range TOTAL"] == "Over 365")
        & df["Category"].isin(D2C_CATEGORIES)
    )
    aged = df[mask].copy()

    # Map SKU → segment key
    aged["_seg"] = aged["Seller Product SKU"].map(sku_map)
    aged = aged[aged["_seg"].notna()]

    # Aggregate per SKU
    agg = (
        aged.groupby("Seller Product SKU")
        .agg({
            "_seg":          "first",
            "Style":         "first",
            "Color":         "first",
            "Size":          "first",
            "Total Amount $": "sum",
            "Qty":           "sum",
        })
        .reset_index()
        .rename(columns={"Qty": "Total Aged", "Total Amount $": "Valuation"})
    )

    # Split into per-segment DataFrames, sorted by Total Aged descending
    segments = {}
    for key in SEGMENT_ORDER:
        seg_df = agg[agg["_seg"] == key].drop(columns=["_seg"])
        segments[key] = seg_df.sort_values("Total Aged", ascending=False).reset_index(drop=True)

    return segments


def summarize_segment(df: pd.DataFrame) -> dict:
    shipped = df["shipped_ytd"].sum() if "shipped_ytd" in df.columns else 0
    stuck   = int((df["weeks_to_evac"] > 52).sum()) if "weeks_to_evac" in df.columns else 0
    return {
        "total_aged_units": int(df["Total Aged"].sum()),
        "total_valuation":  float(df["Valuation"].sum()),
        "total_shipped":    int(shipped),
        "stuck_count":      stuck,
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
    for col in ["Seller Product SKU", "Style", "Category", "Range TOTAL", "Size", "Color"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    print("\nLoading SKU→segment mapping...")
    sku_map = load_sku_segment(REFERENCE_DIR)

    print("\nLoading weekly shipments...")
    shipments = load_weekly_shipments(PROCESSED_DIR)

    print("\nSegmenting SKUs...")
    segments = segment_skus(df, sku_map)

    for key, seg_df in segments.items():
        print(f"  {key:<12}: {len(seg_df):>4} SKUs")

    # Merge shipment data
    segments = {k: build_sku_records(v, shipments) for k, v in segments.items()}

    # Summaries
    summaries = {k: summarize_segment(v) for k, v in segments.items()}

    # Week columns for chart
    wk_cols = [c for c in (shipments.columns.tolist() if shipments is not None else [])
               if re.match(r"^WK?\d+$", c, re.IGNORECASE)]

    # Render template
    env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))
    env.globals.update(
        fmt_val=fmt_val, fmt_units=fmt_units,
        weeks_to_risk=weeks_to_risk, RISK_LABELS=RISK_LABELS,
        enumerate=enumerate, zip=zip, abs=abs,
    )
    template = env.get_template("report2.html")
    html = template.render(
        segments=segments,
        summaries=summaries,
        shipments=shipments,
        wk_cols=wk_cols,
        segment_order=SEGMENT_ORDER,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )

    out_path = REPORTS_DIR / "aging_evacuation_analysis.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"\nReport saved: {out_path.relative_to(ROOT)}")

    # ── Append Target Calibration section ────────────────────────────────────
    try:
        import sys as _sys
        _sys.path.insert(0, str(Path(__file__).parent))
        from generate_target_calibration import build_calibration_html
        print("\nChecking for target calibration data...")
        calib_html = build_calibration_html(section_only=True)
        if calib_html:
            evac_html = out_path.read_text(encoding="utf-8")
            evac_html = evac_html.replace("</body>", f"{calib_html}\n</body>")
            out_path.write_text(evac_html, encoding="utf-8")
            print("  Target Calibration section appended.")
    except Exception as _e:
        print(f"  (Calibration skipped: {_e})")

    print("Opening in browser...")
    webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
