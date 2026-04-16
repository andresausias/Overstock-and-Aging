# Workflow: Build Hub Portal

## Objective
Generate the local intelligence hub — a static HTML portal that indexes all reports, shows KPI trends over time, and lets you open/embed any report.

## Run

```bash
python tools/build_hub.py
```

Opens `.tmp/hub/index.html` in your browser automatically.

---

## What the Hub Contains

| Feature | Description |
|---------|-------------|
| Sidebar | Links to all generated MoM and evacuation reports |
| KPI Timeline | Aged value, aged units, and % of total across all months |
| Overview tab | Charts + report index with Open buttons |
| Viewer tab | Embedded iframe — click any report to view it inline |
| Open in New Tab | Button to open any report in a full browser tab |

---

## How Reports Accumulate

Every time you run `generate_report1.py` or `generate_report2.py`, the output lands in `.tmp/reports/`. Running `build_hub.py` afterwards re-scans this directory and adds the new reports to the hub index.

**Historical data:** Each filtered aging CSV in `.tmp/filtered/` contributes one point to the KPI timeline. The more months you process, the richer the trend chart.

---

## Customization

### Changing the hub color scheme
Edit `tools/templates/hub.html`. The dark theme uses `#0f172a` (page), `#1e293b` (sidebar/cards). Adjust CSS variables at the top of `<style>`.

### Adding new sections to the hub
The hub template uses Jinja2. Add new tab buttons and panels in `tools/templates/hub.html`. Feed new data from `build_hub.py` by extending the `hub_data` dict and `template.render()` call.

### Updating total inventory denominator
The KPI % calculations use `TOTAL_INVENTORY_DENOMINATOR` in `build_hub.py`. Update this value when the denominator changes (or move it to `run_params.json`).

---

## Notes
- The hub is local-only — it is not uploaded to Drive (it links to local `.tmp/reports/` paths)
- If you move or rename the `.tmp/reports/` directory, update the `reports_dir` variable in `build_hub.py`
- Re-running is safe and always reflects the current state of `.tmp/`
