# Workflow: Validate Inventory Data

## Objective
Ensure data integrity before generating any HTML reports. Two checks must pass.

## Run
```bash
python tools/validate_inventory.py
```

## What It Checks

### Check 1: Category Sum (tolerance < $1)
**Formula:** D2C + WHOLESALE + EMP + BAD LOT aged value = Grand Total aged value (Over 365)

**Why it matters:** If this fails, it means there are rows in the data that are aged but don't belong to a known category — usually Amazon rows that were not filtered, SPADR- SKUs still present, or Eurofina rows slipping through.

**If it fails:**
1. Check `filter_summary.json` in `.tmp/validated/` — were any rows removed for Amazon/SPADR-/Eurofina?
2. Open `.tmp/filtered/<file>_filtered.csv` and look for rows where `Category` is not D2C, WHOLESALE, EMP, or BAD LOT
3. Re-run `filter_inventory.py` and retry

### Check 2: Location Sum (tolerance < $100)
**Formula:** Sum of all 9 location aged values = Grand Total aged value (Over 365)

**Why it matters:** Discrepancy usually means rows with unrecognized location codes, or rows without a location value.

**If it fails:**
1. Look at the diff in the console output — which locations don't add up?
2. Check for new warehouse codes not in the standard 9 (`JD NJ`, `JD ATL`, `JD LA`, `Lateral TJ`, `JD Canada`, `JD UK`, `JD AU`, `JD SA`, `CHE CN`)
3. If a new location is legitimate, add it to `KNOWN_LOCATIONS` in `validate_inventory.py` and update the workflow

## Output
`.tmp/validated/validation_report.json` — per-file pass/fail with breakdown.

## Abort Behavior
If either check fails, `sys.exit(1)` is called. The report generators (`generate_report1.py`, `generate_report2.py`) should only be run after a clean validation.
