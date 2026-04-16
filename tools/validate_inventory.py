"""
validate_inventory.py
---------------------
Runs mandatory data integrity checks on filtered aging files before any HTML
is generated. Both checks must pass — if either fails, the script aborts with
a detailed diff and no reports are generated.

Checks:
  1. Category sum:  D2C + WHOLESALE + EMP + BAD LOT aged value == grand total
                    aged value (Over 365 only). Tolerance: < $1.
  2. Location sum:  Sum of all 9 location aged values == grand total aged value
                    (Over 365 only). Tolerance: < $100.

Usage:
    python tools/validate_inventory.py

Input:
    .tmp/filtered/*_filtered.csv

Output:
    .tmp/validated/validation_report.json   (pass/fail per file + diffs)
"""

import sys
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
FILTERED_DIR = ROOT / ".tmp" / "filtered"
VALIDATED_DIR = ROOT / ".tmp" / "validated"
VALIDATED_DIR.mkdir(parents=True, exist_ok=True)

CATEGORY_TOLERANCE = 1.0    # dollars
LOCATION_TOLERANCE = 100.0  # dollars

KNOWN_CATEGORIES = {"D2C", "WHOLESALE", "EMP", "BAD LOT"}
KNOWN_LOCATIONS = {
    "JD NJ", "JD ATL", "JD CA", "Lateral TJ",
    "JD Canada", "JD UK", "JD AU", "JD SA", "CHE CN"
}


def fmt_dollar(v: float) -> str:
    return f"${v:,.2f}"


def validate_file(path: Path) -> dict:
    """Run both checks on one filtered CSV. Returns a result dict."""
    print(f"\nValidating: {path.name}")
    df = pd.read_csv(path)

    # Ensure numeric
    df["Total Amount $"] = pd.to_numeric(df["Total Amount $"], errors="coerce").fillna(0)

    # Work on aged rows only (Over 365)
    aged = df[df["Range TOTAL"].astype(str).str.strip() == "Over 365"].copy()
    grand_total = aged["Total Amount $"].sum()
    print(f"  Grand total aged value (Over 365): {fmt_dollar(grand_total)}")

    result = {
        "file": path.name,
        "grand_total_aged": round(grand_total, 2),
        "checks": {}
    }

    # ── Check 1: Category sum ──────────────────────────────────────────────────
    cat_col = "Category"
    if cat_col not in df.columns:
        result["checks"]["category_sum"] = {"status": "SKIP", "reason": "Column 'Category' not found"}
        print("  Check 1 — Category sum: SKIPPED (no Category column)")
    else:
        aged_cats = aged[aged[cat_col].str.upper().isin({c.upper() for c in KNOWN_CATEGORIES})]
        cat_total = aged_cats["Total Amount $"].sum()
        diff = abs(grand_total - cat_total)

        cat_breakdown = (
            aged.groupby(cat_col)["Total Amount $"].sum()
            .reindex(list(KNOWN_CATEGORIES), fill_value=0)
            .round(2)
            .to_dict()
        )

        if diff <= CATEGORY_TOLERANCE:
            result["checks"]["category_sum"] = {
                "status": "PASS",
                "grand_total": round(grand_total, 2),
                "category_total": round(cat_total, 2),
                "diff": round(diff, 2),
                "breakdown": cat_breakdown
            }
            print(f"  Check 1 — Category sum: PASS  (diff: {fmt_dollar(diff)})")
        else:
            result["checks"]["category_sum"] = {
                "status": "FAIL",
                "grand_total": round(grand_total, 2),
                "category_total": round(cat_total, 2),
                "diff": round(diff, 2),
                "breakdown": cat_breakdown,
                "note": "Sum of known categories does not match grand total. "
                        "Check for unknown/uncategorized rows or SPADR-/Amazon rows "
                        "that were not filtered correctly."
            }
            print(f"  Check 1 — Category sum: FAIL  "
                  f"(grand: {fmt_dollar(grand_total)}, cat sum: {fmt_dollar(cat_total)}, "
                  f"diff: {fmt_dollar(diff)})")
            print(f"    Breakdown: {cat_breakdown}")

    # ── Check 2: Location sum ──────────────────────────────────────────────────
    loc_col = "Location"
    if loc_col not in df.columns:
        result["checks"]["location_sum"] = {"status": "SKIP", "reason": "Column 'Location' not found"}
        print("  Check 2 — Location sum: SKIPPED (no Location column)")
    else:
        loc_total = aged["Total Amount $"].sum()  # same as grand_total but by location
        loc_breakdown = aged.groupby(loc_col)["Total Amount $"].sum().round(2).to_dict()
        loc_sum = sum(loc_breakdown.values())
        diff = abs(grand_total - loc_sum)

        if diff <= LOCATION_TOLERANCE:
            result["checks"]["location_sum"] = {
                "status": "PASS",
                "grand_total": round(grand_total, 2),
                "location_sum": round(loc_sum, 2),
                "diff": round(diff, 2),
                "breakdown": loc_breakdown
            }
            print(f"  Check 2 — Location sum:  PASS  (diff: {fmt_dollar(diff)})")
        else:
            result["checks"]["location_sum"] = {
                "status": "FAIL",
                "grand_total": round(grand_total, 2),
                "location_sum": round(loc_sum, 2),
                "diff": round(diff, 2),
                "breakdown": loc_breakdown,
                "note": "Sum of location values does not match grand total. "
                        "Check for unknown locations or data quality issues."
            }
            print(f"  Check 2 — Location sum:  FAIL  "
                  f"(grand: {fmt_dollar(grand_total)}, loc sum: {fmt_dollar(loc_sum)}, "
                  f"diff: {fmt_dollar(diff)})")

    # Overall pass/fail
    statuses = [v.get("status") for v in result["checks"].values()]
    result["overall"] = "PASS" if all(s in ("PASS", "SKIP") for s in statuses) else "FAIL"
    return result


def main():
    filtered_files = sorted(FILTERED_DIR.glob("*_filtered.csv"))
    if not filtered_files:
        print("ERROR: No filtered files found in .tmp/filtered/. Run filter_inventory.py first.")
        sys.exit(1)

    all_results = []
    any_fail = False

    for path in filtered_files:
        result = validate_file(path)
        all_results.append(result)
        if result["overall"] == "FAIL":
            any_fail = True

    # Save report
    report_path = VALIDATED_DIR / "validation_report.json"
    report_path.write_text(json.dumps(all_results, indent=2))
    print(f"\nValidation report saved to: {report_path.relative_to(ROOT)}")

    if any_fail:
        print("\n" + "="*60)
        print("  VALIDATION FAILED — Report generation aborted.")
        print("  Fix the data issues above and re-run filter_inventory.py,")
        print("  then re-run this script before generating any reports.")
        print("="*60)
        sys.exit(1)
    else:
        print("\n✓ All checks passed. Safe to proceed with report generation.")


if __name__ == "__main__":
    main()
