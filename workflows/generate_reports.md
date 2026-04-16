# Workflow: Generate HTML Reports

## Objective
Generate both self-contained HTML reports from filtered aging data.

## Prerequisites
- `python tools/validate_inventory.py` passed (no FAIL status)
- `.tmp/filtered/` contains at least 2 filtered aging CSVs (for Report 1)

---

## Report 1: Aging MoM Analysis (34 sections)

```bash
python tools/generate_report1.py
```

**Auto-detection:** Picks the two most recent filtered CSVs by filename date prefix (`YYYYMM`).

**Output:** `.tmp/reports/aging_<prev>_<curr>_<year>.html` — opens in browser automatically.

### 34 Sections Include:
- S1: Key metrics summary table
- S2: Aging distribution (all buckets, grouped bar)
- S3: US vs International
- S4-5: Category doughnuts
- S6-7, S12: Top 15 style horizontal bars (all / wholesale / D2C group)
- S8-11: Movement variance charts
- S13-14: Top 25 increase/decrease tables
- S15: Location heatmap
- S16: Location comparison (3 charts)
- S17-25: Per-location style increases (ordered by worst deterioration first)
- S26-34: Per-location absolute values (same order)

---

## Report 2: SKU Evacuation Analysis (10 sections)

```bash
python tools/generate_report2.py
```

**Auto-detection:** Picks the most recent filtered aging CSV. Looks for a shipment file in `.tmp/raw/`.

**Output:** `.tmp/reports/aging_evacuation_analysis.html` — opens in browser automatically.

### 10 Sections Include:
- Summary cards for 6 segments
- S1-6: SKU tables per segment (10210 Extended, 10210 Regular, 10400, Bad Lot, EMP, Discontinued)
- S7: Critical slow-evacuation SKUs by valuation
- S8: Weekly trend charts (tabbed by segment)
- S9: Evacuation bubble chart (velocity vs aged units)
- S10: Size curve check

---

## Customization

### Adding a new segment to Report 2
Edit `tools/generate_report2.py` in the `segment_skus()` function. Add a new key to the returned dict and add the corresponding summary card and table in `tools/templates/report2.html`.

### Modifying silicone styles list
Edit `tools/config/silicone_styles.json`. Both reports read from this file.

### Adjusting insight box text
Insight boxes are in the Jinja2 templates: `tools/templates/report1.html` and `report2.html`. Edit the text inside `<div class="insight">` blocks directly.
