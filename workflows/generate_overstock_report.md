# Workflow: Generate Overstock Rolling Forecast Report

## Objective
Generate the 15-section Overstock Inventory Analysis HTML report from a Rolling Forecast Excel file.

## When to Run
- A new Rolling Forecast Excel file is available (typically monthly)
- After running, rebuild the hub: `python tools/build_hub.py`

## Required Input

Place the Rolling Forecast Excel file in:
```
input/overstock/Rolling_forecast_<Month>_PP_<Year>_-_All_channels.xlsx
```

Example: `input/overstock/Rolling_forecast_March_PP_2026_-_All_channels.xlsx`

**Two tabs must exist in the file:**
1. `Rolling forecast {Month}PP by SKU` — all active channels
2. `Disco&NB {Month}PP by SKU` — discontinued + national brands

**Key data requirements:**
- Header row at row 9 (0-indexed row 8)
- First 103 columns used (sheet duplicates after ~col 103)
- Required columns: Country, Channel, Brand, Category, Style, SKU, `EOM [MMM-YY]` (×13), `OVSTK [MMM-YY]` (×13), `LC - DDP (IS-OTB-Forecast)`

## Steps

1. **Place input file** in `input/overstock/`
2. **Run the generator:**
   ```bash
   python tools/generate_overstock_report.py
   ```
   Or pass the path explicitly:
   ```bash
   python tools/generate_overstock_report.py path/to/file.xlsx
   ```
3. **Output:** `.tmp/overstock/overstock_<start>_<end>_report.html`
4. **Rebuild hub:**
   ```bash
   python tools/build_hub.py
   ```
   The report will appear under `output/reports/overstock/` and show in the 🏷️ Overstock section of the hub sidebar.

## Output Structure
- 15 sections: Key Metrics, Trajectory, % of Inventory, By Country, By Channel, By Category, Top 15 All/D2C/Amazon, Evacuation Charts, Top 25 Value Table, Slowest Evac Table, Country Heatmap, Country Trajectory, By Brand
- Self-contained HTML (Chart.js v4.4.0 CDN)
- All data hardcoded into JavaScript — no dynamic fetching

## Known Data Pitfalls
1. **Duplicate columns**: Sheet duplicates all columns after ~col 103. Script uses `.iloc[:, :103]`.
2. **Header row offset**: Actual data header at row 9 (0-indexed row 8). Script uses `header=8`.
3. **Style .0 suffix**: Numeric styles load as floats. Stripped with `.str.replace('.0', '', regex=False)`.
4. **Negative overstock**: Some OVSTK cells may be negative. Clipped with `.clip(lower=0)` before summing.
5. **Missing LC-DDP**: Rows without landed cost contribute $0 to value totals — expected for some Disco&NB rows.

## Updating for New Months
When a new Rolling Forecast is available:
1. Update or replace the file in `input/overstock/`
2. Verify tab names match (month reference in the tab name may change)
3. Run `python tools/generate_overstock_report.py` — full rebuild from new data
4. Run `python tools/build_hub.py`
5. Never patch individual values — always do a full rebuild

## Silicone Band Styles (shown in purple in charts)
```
10400, 18437, 42001, 42004, 41005, 42024, 42066, 42075,
51001, 51009, 52002, 52005, 54008, 55021, 98099, 77001, 77002
```
Update the `SILICONE_STYLES` set in `tools/generate_overstock_report.py` if new styles are added.
