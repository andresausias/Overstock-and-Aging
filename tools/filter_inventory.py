"""
filter_inventory.py
-------------------
Applies all mandatory business filters to raw aging Excel files in .tmp/raw/.
Outputs one clean CSV per aging file to .tmp/filtered/.

Mandatory filters (non-negotiable — failure to apply corrupts all outputs):
  1. Exclude packaging SKUs   (matched against Packaging_SKUs_Items_1.xlsx)
  2. Exclude Amazon category  (Category == "Amazon")
  3. Exclude SPADR- SKUs      (Seller Product SKU starts with "SPADR-")
  4. Exclude Eurofina owner   (Owner == "Eurofina")

Note: "Aged inventory" (Range TOTAL == "Over 365") is NOT applied here —
it is applied per-report as needed. This file outputs the full cleaned dataset.

Usage:
    python tools/filter_inventory.py

Output:
    .tmp/filtered/<filename>_filtered.csv   for each aging file
    .tmp/filtered/filter_summary.json       row counts before/after per file
"""

import sys
import json
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
AGING_DIR      = ROOT / "input" / "aging"
EXCLUSIONS_DIR = ROOT / "input" / "reference"
FILTERED_DIR   = ROOT / ".tmp" / "filtered"
FILTERED_DIR.mkdir(parents=True, exist_ok=True)

PACKAGING_FILE = "Packaging_SKUs_Items_1.xlsx"

# Regex to detect monthly aging files: YYYYMM_Aging*.xlsx / *.csv
AGING_FILE_PATTERN = re.compile(r"^\d{6}_Aging", re.IGNORECASE)


def load_packaging_skus(exclusions_dir: Path) -> set:
    """Load the packaging SKU exclusion list from the exclusions folder.
    Accepts any Excel/CSV file present — no strict filename matching."""
    if not exclusions_dir.exists():
        print(f"  WARNING: input/exclusions/ not found — packaging filter skipped.")
        return set()

    candidates = [
        f for f in exclusions_dir.iterdir()
        if f.is_file() and f.suffix.lower() in (".xlsx", ".xls", ".csv")
        and f.name != ".gitkeep"
    ]
    if not candidates:
        print(f"  WARNING: No exclusion file found in input/exclusions/ — packaging filter skipped.")
        return set()

    pkg_path = candidates[0]
    print(f"  Using exclusion file: {pkg_path.name}")

    df = pd.read_excel(pkg_path, sheet_name=0)
    # Accept any column named SKU, Seller Product SKU, or first column
    sku_col = next(
        (c for c in df.columns if "sku" in c.lower() or "seller" in c.lower()),
        df.columns[0]
    )
    skus = set(df[sku_col].dropna().astype(str).str.strip())
    print(f"  Loaded {len(skus):,} packaging SKUs from {PACKAGING_FILE}")
    return skus


def clean_numeric(series: pd.Series) -> pd.Series:
    """Strip $ and , from a string series and convert to numeric."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce"
    ).fillna(0)


def load_aging_file(path: Path) -> pd.DataFrame:
    """Load an aging Excel or CSV file with standard column cleaning.
    Handles files where real column names are in row 1 (blank/unnamed header row 0)."""
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=0)
    else:
        df = pd.read_csv(path, low_memory=False)

    # Detect shifted headers: real column names are in row 0 when the Excel header row is blank.
    # Heuristic: if "Seller Product SKU" doesn't appear in actual columns but DOES appear in
    # the first data row, promote that row to be the header.
    SENTINEL = "Seller Product SKU"
    has_sentinel_in_cols = any(SENTINEL.lower() in str(c).lower() for c in df.columns)
    has_sentinel_in_row0 = len(df) > 0 and any(
        SENTINEL.lower() in str(v).lower() for v in df.iloc[0].values
    )
    if not has_sentinel_in_cols and has_sentinel_in_row0:
        df.columns = [str(v).strip() for v in df.iloc[0].values]
        df = df.iloc[1:].reset_index(drop=True)

    # Strip whitespace from column names
    df.columns = [c.strip() for c in df.columns]

    # Clean key numeric columns
    if "Total Amount $" in df.columns:
        df["Total Amount $"] = clean_numeric(df["Total Amount $"])
    if "Qty" in df.columns:
        df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)

    # Strip string columns
    for col in ["Seller Product SKU", "Category", "Owner", "Range TOTAL", "Location"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df


def apply_filters(df: pd.DataFrame, packaging_skus: set[str], filename: str) -> tuple[pd.DataFrame, dict]:
    """Apply all 4 mandatory filters. Returns filtered df and a summary dict."""
    summary = {"file": filename, "rows_before": len(df), "filters": {}}

    # 1. Exclude packaging SKUs
    if packaging_skus and "Seller Product SKU" in df.columns:
        mask = ~df["Seller Product SKU"].isin(packaging_skus)
        removed = (~mask).sum()
        df = df[mask].copy()
        summary["filters"]["packaging_skus_removed"] = int(removed)
        print(f"    Filter 1 — Packaging SKUs:   removed {removed:,} rows")
    else:
        summary["filters"]["packaging_skus_removed"] = 0

    # 2. Exclude Amazon category
    if "Category" in df.columns:
        mask = df["Category"].str.upper() != "AMAZON"
        removed = (~mask).sum()
        df = df[mask].copy()
        summary["filters"]["amazon_rows_removed"] = int(removed)
        print(f"    Filter 2 — Amazon category:  removed {removed:,} rows")
    else:
        summary["filters"]["amazon_rows_removed"] = 0

    # 3. Exclude SPADR- prefixed SKUs
    if "Seller Product SKU" in df.columns:
        mask = ~df["Seller Product SKU"].str.upper().str.startswith("SPADR-")
        removed = (~mask).sum()
        df = df[mask].copy()
        summary["filters"]["spadr_rows_removed"] = int(removed)
        print(f"    Filter 3 — SPADR- SKUs:      removed {removed:,} rows")
    else:
        summary["filters"]["spadr_rows_removed"] = 0

    # 4. Exclude Eurofina owner
    if "Owner" in df.columns:
        mask = df["Owner"].str.upper() != "EUROFINA"
        removed = (~mask).sum()
        df = df[mask].copy()
        summary["filters"]["eurofina_rows_removed"] = int(removed)
        print(f"    Filter 4 — Eurofina owner:   removed {removed:,} rows")
    else:
        summary["filters"]["eurofina_rows_removed"] = 0

    summary["rows_after"] = len(df)
    total_removed = summary["rows_before"] - summary["rows_after"]
    print(f"    Total: {summary['rows_before']:,} → {summary['rows_after']:,} rows ({total_removed:,} removed)")
    return df, summary


def main():
    if not AGING_DIR.exists() or not any(f for f in AGING_DIR.iterdir() if f.name != ".gitkeep"):
        print("ERROR: input/aging/ is empty. Drop your aging Excel files there first.")
        sys.exit(1)

    # Find aging files
    aging_files = [
        f for f in AGING_DIR.iterdir()
        if f.is_file() and AGING_FILE_PATTERN.match(f.name)
        and f.suffix.lower() in (".xlsx", ".xls", ".csv")
    ]

    if not aging_files:
        print("No aging files found in input/aging/. Expected names like: 202602_AgingFebruary.xlsx")
        sys.exit(1)

    print(f"Found {len(aging_files)} aging file(s) to filter.\n")

    packaging_skus = load_packaging_skus(EXCLUSIONS_DIR)
    all_summaries = []

    for path in sorted(aging_files):
        print(f"\nProcessing: {path.name}")
        df = load_aging_file(path)
        print(f"  Loaded {len(df):,} rows")

        df_filtered, summary = apply_filters(df, packaging_skus, path.name)
        all_summaries.append(summary)

        out_path = FILTERED_DIR / f"{path.stem}_filtered.csv"
        df_filtered.to_csv(out_path, index=False)
        print(f"  Saved: {out_path.relative_to(ROOT)}")

    # Write summary JSON
    summary_path = FILTERED_DIR / "filter_summary.json"
    summary_path.write_text(json.dumps(all_summaries, indent=2))
    print(f"\nFilter summary saved to: {summary_path.relative_to(ROOT)}")
    print(f"\nDone. {len(aging_files)} file(s) filtered.")


if __name__ == "__main__":
    main()
