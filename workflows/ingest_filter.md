# Workflow: Ingest and Filter Inventory Files

## Objective
Download all inventory files from three separate Google Drive input folders and apply mandatory business filters.

## Step 1: Configure Drive Folders
In `.env`:
```
GOOGLE_DRIVE_INPUT_AGING_ID=        # Aging Excel files
GOOGLE_DRIVE_INPUT_EXCLUSIONS_ID=   # Packaging/exclusion lists
GOOGLE_DRIVE_INPUT_SHIPMENTS_ID=    # Weekly shipment data
```
Get each folder ID from its URL: `drive.google.com/drive/folders/<FOLDER_ID>`

## Step 2: Upload Files to the Correct Drive Folders

| Drive Folder | What to Upload |
|---|---|
| **Aging Files** (`AGING_ID`) | `YYYYMM_AgingMonth.xlsx` ×2 (current + prior month) |
| **Exclusions** (`EXCLUSIONS_ID`) | `Packaging_SKUs_Items_1.xlsx` + any future exclusion lists |
| **Shipments** (`SHIPMENTS_ID`) | Weekly units shipped file (any name, any `.xlsx`/`.csv`) |

File naming for aging files must follow: `202602_AgingFebruary.xlsx` (YYYYMM prefix required).

## Step 3: Download
```bash
python tools/drive_download.py
```

Files land in:
- `.tmp/raw/aging/`      ← aging Excel files
- `.tmp/raw/exclusions/` ← packaging exclusion lists
- `.tmp/raw/shipments/`  ← weekly shipment data

## Step 4: Filter
```bash
python tools/filter_inventory.py
```

Applies all 4 mandatory filters to aging files using exclusions from `.tmp/raw/exclusions/`:
1. Packaging SKUs (matched against `Packaging_SKUs_Items_1.xlsx`)
2. Amazon category (`Category == "Amazon"`)
3. SPADR- prefixed SKUs (shadow/virtual records — always remove)
4. Eurofina owner (`Owner == "Eurofina"`)

Output: `.tmp/filtered/<filename>_filtered.csv` + `filter_summary.json`

## Step 5: Verify
```bash
cat .tmp/filtered/filter_summary.json
```
Each file should show non-zero counts for at least the SPADR- and Amazon filters.

## Notes
- Adding a new exclusion type: place a new Excel/CSV in the Drive Exclusions folder and update the filter logic in `tools/filter_inventory.py` to load and apply it.
- Re-running is safe — all `.tmp/` subdirectories are overwritten on each run.
- The "Aged inventory" filter (`Range TOTAL == "Over 365"`) is applied per-report, not here.
