# Workflow: Generate Weekly Performance Slides

## Objective
Generate the 2-slide PPTX (PptxGenJS) for the weekly stakeholder review.

## Required Inputs
| File | Update Frequency |
|------|-----------------|
| `tools/config/run_params.json` | Every week |
| `tools/config/targets.json` | When targets change |
| `tools/config/levers.json` | When lever strategy changes |
| `tools/config/products.json` | When subtitle/action text changes |
| `.tmp/raw/<shipment file>` | Every week (downloaded by drive_download.py) |
| `.tmp/filtered/<latest>_filtered.csv` | Auto (from filter_inventory.py) |

## Weekly Update Steps

### 1. Update run_params.json
```json
{
  "report_date": "Mar 26, 2026",          ← update
  "week_current": 12,                      ← update
  "week_labels": ["WK1","WK2",...,"WK12"],
  "week_date_ranges": ["Jan 4-10",...],
  "num_weeks": 12,
  "overstock_valuation": 21700000,         ← update (manual, from Overstock tracker)
  "overstock_units": 4602646,              ← update (manual)
  "total_inventory_denominator": 50935732,
  "exec_summary_bullets": [               ← update
    "Overstock: ...",
    "Aging: ..."
  ]
}
```

### 2. Run the generator
```bash
cd tools && npm install  # first time only
node tools/generate_slides.js
```

### 3. QA the output
```bash
# Convert to PDF
soffice --headless --convert-to pdf .tmp/slides/weekly_performance_WK12_2026.pptx

# Convert to JPEG for visual inspection
pdftoppm -jpeg -r 200 .tmp/slides/weekly_performance_WK12_2026.pdf .tmp/slides/preview

# Open PDFs
open .tmp/slides/weekly_performance_WK12_2026.pdf
```

### QA Checklist
- [ ] Column sums match Grand Total row
- [ ] No text overflow or clipping
- [ ] Color coding is correct (green ≥100%, amber 70-99%, red <70%)
- [ ] All 11 products present across both slides
- [ ] Section header pills show correct valuation/units
- [ ] Exec summary bullets are updated

## Slide Structure Reference
- **Slide 1:** Per-product rows (5 Overstock + 6 Aging) + exec summary
- **Slide 2:** 5 lever groups (ACQ+CLR, ACQ+CLR+BOOST, CLR ONLY, BOOST+CLR, ACQ) + grand total

## Notes
- Aging valuation/units are auto-extracted from the latest filtered CSV (no manual entry needed)
- Overstock valuation/units must be entered manually in `run_params.json` each week
- `total_inventory_denominator` ($50,935,732) changes infrequently — update when told to
