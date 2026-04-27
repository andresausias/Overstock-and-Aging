"""
process_shipments.py
--------------------
Transforms the raw fulfillment CSV (one row per shipment) into a weekly
aggregated summary by segment — the format expected by generate_report2.py
and generate_slides.js.

Steps:
  1. Load raw shipment CSV from input/shipments/
  2. Load SKU → Segment mapping from input/reference/SKU_Segment.xlsx
  3. Load week date ranges from tools/config/run_params.json
  4. Assign each shipment row to a week number based on package_created_at
  5. Aggregate shipped_qty by segment and week
  6. Write .tmp/processed/weekly_by_segment.csv

Segment name mapping (SKU_Segment.xlsx → config names):
  10022               → 10022 Bra
  10024               → 10024 Straps Bra
  10035               → 10035 Sweetheart Bra
  42075               → 42075 HW Legging
  62001               → 62001 Scoop Neck Cami
  10400 SILICONE BAND → 10400 Silicone Band
  10210 REGULAR OLD LABS  → 10210 Old Labs
  10210 REGULAR NEW LABS  → 10210 New Labs
  10210 EXTENDED (+ SIZES) → 10210 Extended
  BAD LOT             → Bad Lots
  EMP BRAND           → EMP Brand

Usage:
    python tools/process_shipments.py
"""

import re
import sys
import json
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd

ROOT          = Path(__file__).parent.parent
SHIPMENTS_DIR = ROOT / "input" / "shipments"
REFERENCE_DIR = ROOT / "input" / "reference"
CONFIG_DIR    = Path(__file__).parent / "config"
PROCESSED_DIR = ROOT / ".tmp" / "processed"
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

# Map from SKU_Segment.xlsx Segment values → config product names
SEGMENT_MAP = {
    "10022":                    "10022 Bra",
    "10024":                    "10024 Straps Bra",
    "10035":                    "10035 Sweetheart Bra",
    "42075":                    "42075 HW Legging",
    "62001":                    "62001 Scoop Neck Cami",
    "10400 SILICONE BAND":      "10400 Silicone Band",
    "10210 REGULAR OLD LABS":   "10210 Old Labs",
    "10210 REGULAR NEW LABS":   "10210 New Labs",
    "10210 EXTENDED (+ SIZES)": "10210 Extended",
    "BAD LOT":                  "Bad Lots",
    "EMP BRAND":                "EMP Brand",
}


def parse_week_ranges(params: dict) -> list:
    """
    Parse week_date_ranges from run_params.json into (start, end, label) tuples.
    Handles formats like "Jan 4-10", "Jan 25-31", "Jan 28-Feb 3".
    """
    year = int(params.get("report_date", "2026")[-4:])
    week_labels = params["week_labels"]
    week_date_ranges = params["week_date_ranges"]

    MONTHS = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
        "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
    }

    parsed = []
    for label, date_range in zip(week_labels, week_date_ranges):
        # Handle "Jan 4-10" or "Jan 28-Feb 3"
        parts = date_range.split("-")
        start_str = parts[0].strip()   # e.g. "Jan 4" or "Jan 28"
        end_str   = parts[1].strip()   # e.g. "10" or "Feb 3"

        start_parts = start_str.split()
        start_month = MONTHS[start_parts[0]]
        start_day   = int(start_parts[1])
        start_date  = datetime(year, start_month, start_day)

        # End: may be just a day number or "MonthName Day"
        if end_str[0].isalpha():
            end_parts  = end_str.split()
            end_month  = MONTHS[end_parts[0]]
            end_day    = int(end_parts[1])
            # Handle year rollover
            end_year   = year + 1 if end_month < start_month else year
            end_date   = datetime(end_year, end_month, end_day)
        else:
            end_date = datetime(year, start_month, int(end_str))

        parsed.append((start_date, end_date, label))

    return parsed


def assign_week(date: datetime, week_ranges: list) -> Optional[str]:
    """Return the week label for a given date, or None if outside all ranges."""
    for start, end, label in week_ranges:
        if start <= date <= end:
            return label
    return None


def load_sku_segment(reference_dir: Path) -> dict:
    """Load SKU → config segment name mapping."""
    candidates = list(reference_dir.glob("SKU_Segment*"))
    if not candidates:
        print("ERROR: SKU_Segment.xlsx not found in input/reference/")
        sys.exit(1)

    df = pd.read_excel(candidates[0])
    df.columns = df.columns.str.strip()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    df["Segment"] = df["Segment"].astype(str).str.strip().str.upper()

    mapping = {}
    for _, row in df.iterrows():
        raw_seg = row["Segment"]
        config_name = SEGMENT_MAP.get(raw_seg)
        if config_name:
            mapping[row["SKU"]] = config_name
        else:
            # Try numeric segment (stored as float/int in Excel)
            try:
                numeric = str(int(float(raw_seg)))
                config_name = SEGMENT_MAP.get(numeric)
                if config_name:
                    mapping[row["SKU"]] = config_name
            except (ValueError, OverflowError):
                pass

    print(f"  Loaded {len(mapping):,} SKU→segment mappings")
    return mapping


def main():
    # Load config
    params = json.loads((CONFIG_DIR / "run_params.json").read_text())
    week_ranges = parse_week_ranges(params)
    week_labels = [w[2] for w in week_ranges]
    print(f"Week ranges: {week_labels[0]} → {week_labels[-1]} ({len(week_labels)} weeks)")

    # Load SKU→segment mapping
    print("\nLoading SKU→segment mapping...")
    sku_map = load_sku_segment(REFERENCE_DIR)

    # Load raw shipment file
    shipment_files = [
        f for f in SHIPMENTS_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in (".csv", ".xlsx", ".xls")
        and f.name != ".gitkeep"
    ]
    if not shipment_files:
        print("ERROR: No shipment file found in input/shipments/")
        sys.exit(1)

    # Use the most recently modified file
    shipment_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    ship_file = shipment_files[0]
    print(f"\nLoading shipments: {ship_file.name} ({len(shipment_files)} file(s) found, using newest)")
    if ship_file.suffix.lower() == ".csv":
        df = pd.read_csv(ship_file)
    else:
        df = pd.read_excel(ship_file)

    df.columns = df.columns.str.strip()
    print(f"  {len(df):,} rows loaded")

    # Parse dates
    df["date"] = pd.to_datetime(df["package_created_at"], errors="coerce")
    df = df.dropna(subset=["date"])

    # Assign week labels
    df["week"] = df["date"].apply(lambda d: assign_week(d, week_ranges))
    outside = df["week"].isna().sum()
    if outside:
        print(f"  {outside:,} rows outside defined week ranges — excluded")
    df = df[df["week"].notna()]

    # Use shipped_qty (prefer over ordered_qty)
    qty_col = "shipped_qty" if "shipped_qty" in df.columns else "ordered_qty"
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)

    # Clean SKU column — keep full dataset for per-SKU output
    df["sku_clean"] = df["variant_sku"].astype(str).str.strip()
    df_all_skus = df.copy()   # all SKUs within week ranges (for evacuation analysis)

    # Map SKU → segment (filter to known segments for segment aggregation)
    df["segment"] = df["sku_clean"].map(sku_map)
    unmapped = df["segment"].isna().sum()
    print(f"  {unmapped:,} rows with unmapped SKUs — excluded from segment totals")
    df = df[df["segment"].notna()]

    # Aggregate by segment × week
    agg = (
        df.groupby(["segment", "week"])[qty_col]
        .sum()
        .reset_index()
        .rename(columns={qty_col: "units"})
    )

    # Pivot to wide format: segment | WK1 | WK2 | ... | WKn
    pivot = agg.pivot(index="segment", columns="week", values="units").fillna(0)

    # Ensure all week columns present in order
    for wk in week_labels:
        if wk not in pivot.columns:
            pivot[wk] = 0
    pivot = pivot[week_labels]
    pivot = pivot.reset_index().rename(columns={"segment": "Product/Segment"})

    out_path = PROCESSED_DIR / "weekly_by_segment.csv"
    pivot.to_csv(out_path, index=False)

    print(f"\nSegments found:")
    for seg in pivot["Product/Segment"].tolist():
        row = pivot[pivot["Product/Segment"] == seg].iloc[0]
        total = int(row[week_labels].sum())
        print(f"  {seg:<30} {total:>8,} units YTD")

    print(f"\nOutput saved: .tmp/processed/weekly_by_segment.csv")

    # Per-SKU aggregation for evacuation analysis — uses ALL SKUs (not segment-filtered)
    agg_sku = (
        df_all_skus.groupby(["sku_clean", "week"])[qty_col]
        .sum()
        .reset_index()
        .rename(columns={qty_col: "units"})
    )
    pivot_sku = agg_sku.pivot(index="sku_clean", columns="week", values="units").fillna(0)
    for wk in week_labels:
        if wk not in pivot_sku.columns:
            pivot_sku[wk] = 0
    pivot_sku = pivot_sku[week_labels]
    pivot_sku = pivot_sku.reset_index().rename(columns={"sku_clean": "Seller Product SKU"})

    out_sku_path = PROCESSED_DIR / "weekly_by_sku.csv"
    pivot_sku.to_csv(out_sku_path, index=False)
    print(f"Output saved: .tmp/processed/weekly_by_sku.csv ({len(pivot_sku):,} SKUs)")

    # ── Per-segment aged inventory from latest filtered aging file ────────────
    filtered_files = sorted([
        f for f in (ROOT / ".tmp" / "filtered").glob("*_filtered.csv")
        if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)
    ])
    if filtered_files:
        aging_df = pd.read_csv(filtered_files[-1], low_memory=False)
        aging_df.columns = aging_df.columns.str.strip()
        aging_df["Qty"] = pd.to_numeric(aging_df["Qty"], errors="coerce").fillna(0)
        aging_df["sku_clean"] = aging_df["Seller Product SKU"].astype(str).str.strip()
        aging_df["segment"] = aging_df["sku_clean"].map(sku_map)

        seg_aging = {}
        for seg_name in SEGMENT_MAP.values():
            seg_rows = aging_df[aging_df["segment"] == seg_name]
            aged_rows = seg_rows[seg_rows["Range TOTAL"].astype(str).str.strip() == "Over 365"]
            total_qty = int(seg_rows["Qty"].sum())
            aged_qty  = int(aged_rows["Qty"].sum())
            seg_aging[seg_name] = {
                "aged_units":  aged_qty,
                "total_units": total_qty,
                "aged_pct":    round(aged_qty / total_qty * 100, 1) if total_qty > 0 else 0.0,
            }

        out_aging = PROCESSED_DIR / "segment_aging.json"
        out_aging.write_text(json.dumps(seg_aging, indent=2))
        print(f"Output saved: .tmp/processed/segment_aging.json")


if __name__ == "__main__":
    main()
