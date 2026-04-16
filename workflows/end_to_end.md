# Workflow: End-to-End Automation (Master SOP)

## Objective
Run the full automation pipeline from Drive ingestion to Slack delivery.

## Pre-Run Checklist
- [ ] Two monthly aging Excel files uploaded to Drive **Aging Files** folder (`AGING_ID`)
- [ ] `Packaging_SKUs_Items_1.xlsx` uploaded to Drive **Exclusions** folder (`EXCLUSIONS_ID`)
- [ ] Weekly shipment file uploaded to Drive **Shipments** folder (`SHIPMENTS_ID`)
- [ ] `tools/config/run_params.json` updated with current week/date/overstock totals
- [ ] Google OAuth done (`credentials.json` + `token.json` present)
- [ ] Node.js installed (`node --version`)
- [ ] Python deps installed (`pip install -r requirements.txt`)
- [ ] Node deps installed (`cd tools && npm install`)
- [ ] `.env` populated with all required keys

## Full Pipeline

```bash
# Step 1: Download from Drive
python tools/drive_download.py

# Step 2: Apply mandatory business filters
python tools/filter_inventory.py

# Step 3: Validate data integrity (MUST PASS before continuing)
python tools/validate_inventory.py

# Step 4: Generate 34-section MoM HTML report
python tools/generate_report1.py

# Step 5: Generate evacuation analysis HTML report
python tools/generate_report2.py

# Step 6: Build hub portal (opens in browser)
python tools/build_hub.py

# Step 7: Generate PPTX slides
cd tools && node generate_slides.js && cd ..

# Step 8: QA slides (optional but recommended)
soffice --headless --convert-to pdf .tmp/slides/weekly_performance_*.pptx
pdftoppm -jpeg -r 200 .tmp/slides/weekly_performance_*.pdf .tmp/slides/preview

# Step 9: Upload all outputs to Drive
python tools/drive_upload.py

# Step 10: Send Slack notification
python tools/send_slack.py
```

## Edge Cases

### Validation fails (Step 3)
- Read the diff printed to console
- Check if SPADR- or Eurofina rows were not filtered (re-run Step 2)
- Check if there are uncategorized rows (inspect `.tmp/filtered/*.csv`)
- Fix the source data if needed and restart from Step 1

### Missing aging files (Step 1)
- Verify `GOOGLE_DRIVE_INPUT_FOLDER_ID` in `.env` points to the correct folder
- Confirm both monthly Excel files are uploaded to that folder
- Check file naming: must match `YYYYMM_Aging*.xlsx`

### Slides show no shipment data (Step 7)
- Check that a shipment file exists in `.tmp/raw/` with "ship", "weekly", "wk", or "week" in the name
- Verify product name column matches names in `tools/config/targets.json`

### Re-run behavior
- All tools are idempotent — safe to re-run any step
- `.tmp/` directories are cleared and rebuilt each run
- Drive uploads replace existing files with the same name (no duplicates)

## Output Locations
| Output | Local | Drive |
|--------|-------|-------|
| HTML reports | `.tmp/reports/` | `FINALIZED_ID` folder |
| Raw files | `.tmp/raw/` | `RAW_ID` folder |
| PPTX slides | `.tmp/slides/` | `SLIDES_ID` folder |
| Hub portal | `.tmp/hub/index.html` | (local only, open in browser) |
