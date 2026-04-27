"""
generate_target_calibration.py
------------------------------
Generates the Target Calibration Analysis HTML section.

Reads (priority order):
  1. input/targets/YYYYMM_Targets*.xlsx  — targets Excel file (Q2 AGING FINAL ÷ 3 = monthly target)
     Accomplishment is derived automatically from Feb→Mar aging movement (max(0, prior_qty - curr_qty)).
  2. input/targets/targets_<month>.json  — manually-entered targets & accomplishments (fallback)

  .tmp/filtered/*_filtered.csv — two most recent filtered aging CSVs (current + prior)

The script can be used two ways:
  1. Imported as a module: call build_calibration_html() → returns HTML string to embed
  2. Run directly: generates standalone HTML + appends to aging_evacuation_analysis.html

Usage:
    python tools/generate_target_calibration.py

Output:
    .tmp/reports/target_calibration_<MonYYYY>.html   (standalone page)
    Appended section in .tmp/reports/aging_evacuation_analysis.html (if it exists)
"""

import json
import re
import sys
import webbrowser
from pathlib import Path
from typing import Optional

import pandas as pd

ROOT           = Path(__file__).parent.parent
FILTERED_DIR   = ROOT / ".tmp" / "filtered"
REPORTS_DIR    = ROOT / ".tmp" / "reports"
TARGETS_DIR    = ROOT / "input" / "targets"
CONFIG_DIR     = Path(__file__).parent / "config"
REFERENCE_DIR  = ROOT / "input" / "reference"

SILICONE_STYLES = set(
    json.loads((CONFIG_DIR / "silicone_styles.json").read_text())["styles"]
)

D2C_CATEGORIES = {"D2C", "EMP", "BAD LOT"}


# ── Data loading ──────────────────────────────────────────────────────────────

def load_packaging_skus() -> set:
    candidates = [
        f for f in REFERENCE_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in (".xlsx", ".xls", ".csv")
        and f.name != ".gitkeep"
    ] if REFERENCE_DIR.exists() else []
    if not candidates:
        return set()
    df = pd.read_excel(candidates[0], sheet_name=0)
    sku_col = next(
        (c for c in df.columns if "sku" in c.lower() or "seller" in c.lower()),
        df.columns[0]
    )
    return set(df[sku_col].dropna().astype(str).str.strip())


def load_d2c_aged(csv_path: Path, packaging_skus: set) -> pd.DataFrame:
    """Load a filtered CSV and return D2C Over 365, grouped by Style."""
    df = pd.read_csv(csv_path, low_memory=False)
    df.columns = df.columns.str.strip()

    # Numeric
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Total Amount $"] = pd.to_numeric(
        df["Total Amount $"].astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    # Strip strings
    for col in ["Category", "Range TOTAL", "Seller Product SKU", "Style"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Filter: D2C group + Over 365 + no packaging + no SPADR-
    mask = (
        df["Category"].isin(D2C_CATEGORIES)
        & (df["Range TOTAL"] == "Over 365")
        & ~df["Seller Product SKU"].isin(packaging_skus)
        & ~df["Seller Product SKU"].str.upper().str.startswith("SPADR-")
    )
    df = df[mask].copy()

    agg = (
        df.groupby("Style")
        .agg(aged_val=("Total Amount $", "sum"), aged_qty=("Qty", "sum"))
        .reset_index()
    )
    agg["aged_k"] = (agg["aged_val"] / 1000).round().astype(int)
    agg["aged_qty"] = agg["aged_qty"].round().astype(int)
    return agg


STRATEGY_DEFAULTS = {
    "10210": "ACQ", "10400": "ACQ", "42075": "ACQ", "62001": "ACQ", "10024": "ACQ",
    "10022": "ACQ", "81068": "ACQ", "91404": "ACQ", "51422": "ACQ", "55021": "ACQ", "10035": "ACQ",
    "95001": "Non ACQ", "52002": "Non ACQ", "91401": "Non ACQ", "95006": "Non ACQ", "95005": "Non ACQ",
    "51401": "Non ACQ", "62010": "Non ACQ", "13402": "Non ACQ", "51001": "Non ACQ", "81007": "Non ACQ",
    "95002": "Non ACQ", "48001": "Non ACQ", "56002": "Non ACQ", "42024": "Non ACQ", "97097": "Non ACQ",
    "81005": "Non ACQ", "91400": "Non ACQ", "31048": "Non ACQ", "62008": "Non ACQ", "41005": "Non ACQ",
    "41402": "Non ACQ", "73005": "Non ACQ", "51427": "Non ACQ", "73007": "Non ACQ", "81004": "Non ACQ",
    "52005": "Non ACQ", "51009": "Non ACQ", "98099": "Non ACQ", "42066": "Non ACQ",
    "18437": "Disc", "73003": "Disc",
    "77002": "New Launch",
}


def find_latest_targets_file() -> Optional[Path]:
    """Return newest file in input/targets/ — Excel preferred, JSON fallback."""
    if not TARGETS_DIR.exists():
        return None
    # Excel files take priority (YYYYMM prefix)
    xlsx_candidates = sorted(
        [f for f in TARGETS_DIR.glob("*.xlsx") if f.name != ".gitkeep"],
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    if xlsx_candidates:
        return xlsx_candidates[0]
    # Fallback: JSON
    json_candidates = sorted(
        [f for f in TARGETS_DIR.glob("targets_*.json") if f.name != ".gitkeep"],
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    return json_candidates[0] if json_candidates else None


def load_targets_from_excel(xlsx_path: Path, curr_aged: pd.DataFrame, prior_aged: pd.DataFrame) -> Optional[dict]:
    """
    Read a targets Excel (YYYYMM_Targets*.xlsx) and return a targets_data dict
    compatible with the JSON format, computing accomplishment from aging movement.

    Target  = Q2 AGING FINAL total (all countries) ÷ 3 (monthly)
    Accomp  = max(0, prior_qty − current_qty) per style (units that left Over 365)
    """
    df = pd.read_excel(xlsx_path, sheet_name=0)
    df.columns = df.columns.str.strip()
    df["Style"] = df["Style"].astype(str).str.strip()

    # Find target column (contains "TARGET" case-insensitive)
    tgt_col = next(
        (c for c in df.columns if "target" in c.lower() and "weekly" not in c.lower()),
        None
    )
    if tgt_col is None:
        print(f"  (Calibration) Could not find target column in {xlsx_path.name}")
        return None

    by_style = df.groupby("Style")[tgt_col].sum().reset_index()
    by_style["monthly_target"] = (by_style[tgt_col] / 3).round().astype(int)
    by_style = by_style[by_style["monthly_target"] > 0].copy()

    curr_map  = curr_aged.set_index("Style").to_dict("index") if not curr_aged.empty else {}
    prior_map = prior_aged.set_index("Style").to_dict("index") if not prior_aged.empty else {}

    # Parse month/year from filename prefix (YYYYMM_...)
    m = re.match(r"(\d{4})(\d{2})_", xlsx_path.name)
    curr_yyyymm  = m.group(1) + m.group(2) if m else ""
    prior_yyyymm = ""

    # Derive month labels from filtered CSVs
    all_filtered = sorted([
        f for f in FILTERED_DIR.glob("*_filtered.csv")
        if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)
    ])
    month_label  = "March 2026"
    prior_label  = "February 2026"
    if all_filtered:
        def _label_from(fname: str) -> str:
            import calendar
            mm = re.match(r"(\d{4})(\d{2})_Aging(\w+)?", fname, re.IGNORECASE)
            if mm:
                y, mo = mm.group(1), mm.group(2)
                return f"{calendar.month_name[int(mo)]} {y}"
            return fname
        curr_csv = next((f for f in all_filtered if f.name.startswith(curr_yyyymm)), all_filtered[-1])
        prior_csv_candidates = [f for f in all_filtered if f != curr_csv]
        month_label = _label_from(curr_csv.name)
        if prior_csv_candidates:
            prior_label = _label_from(prior_csv_candidates[-1].name)
            prior_yyyymm = prior_csv_candidates[-1].name[:6]

    styles = []
    for _, row in by_style.iterrows():
        style  = str(row["Style"])
        tgt    = int(row["monthly_target"])
        c_info = curr_map.get(style, {})
        p_info = prior_map.get(style, {})
        curr_qty = int(c_info.get("aged_qty", 0))
        prior_qty = int(p_info.get("aged_qty", 0))
        accomp = max(0, prior_qty - curr_qty)
        strategy = STRATEGY_DEFAULTS.get(style, "Non ACQ")
        styles.append({"style": style, "strategy": strategy, "target": tgt, "accomp": accomp})

    return {
        "month":          month_label,
        "prior_month":    prior_label,
        "current_yyyymm": curr_yyyymm,
        "prior_yyyymm":   prior_yyyymm,
        "styles":         styles,
        "_source":        "excel",
    }


# ── Quadrant & assessment logic ───────────────────────────────────────────────

def get_quadrant(mv: int, acc_pct: int) -> str:
    if mv > 0 and acc_pct < 100:  return "Q1"
    if mv > 0 and acc_pct >= 100: return "Q2"
    if mv <= 0 and acc_pct >= 100: return "Q3"
    return "Q4"


def get_assessment(q: str, acc_pct: int) -> str:
    if q == "Q1":
        if acc_pct < 20:   return "Target not met. Investigate root cause."
        if acc_pct >= 80:  return "Target too low. Nearly met but didn't offset inflow."
        return "Target too low and under-delivered."
    if q == "Q2":
        if acc_pct > 200:  return "Target far too low. Raise 3-5x."
        return "Target too low. Met but didn't offset inflow."
    if q == "Q3":
        if acc_pct > 200:  return "Target too low. Raise to match capacity."
        if acc_pct >= 80:  return "✅ Target well set."
        return "Target slightly high for actual capacity."
    # Q4
    if acc_pct < 50:       return "Target too high for actual capacity."
    return "Target slightly high. Aging improved naturally."


# ── HTML generation ───────────────────────────────────────────────────────────

def _tag(cls: str, text: str) -> str:
    return f'<span class="calib-tag calib-tag-{cls}">{text}</span>'


def _acc_dot(acc_pct: int) -> str:
    if acc_pct >= 100:   color = "#66bb6a"
    elif acc_pct >= 70:  color = "#ffa726"
    elif acc_pct >= 50:  color = "#ef5350"
    else:                color = "#ef5350"
    return f'<span class="calib-acc-dot" style="background:{color}"></span>'


def _acc_style(acc_pct: int) -> str:
    if acc_pct >= 100:   return "color:#2e7d32;font-weight:700"
    elif acc_pct < 50:   return "color:#c62828;font-weight:700"
    else:                return "color:#e65100;font-weight:700"


def _mv_style(mv: int) -> str:
    if mv > 0:    return "color:#c62828;font-weight:700"
    elif mv < 0:  return "color:#2e7d32;font-weight:700"
    return ""


def _tgt_pct_style(pct: float) -> str:
    if pct < 3:    return "color:#c62828;font-weight:700"
    elif pct < 6:  return "color:#e65100"
    return "color:#333"


def build_calibration_html(
    targets_file: Optional[Path] = None,
    section_only: bool = False
) -> Optional[str]:
    """
    Build the calibration HTML.

    Args:
        targets_file: path to targets JSON (auto-detected if None)
        section_only: if True, return only the inner div (no <html>/<head> wrapper)

    Returns:
        HTML string, or None if no targets file is found / targets are all zero.
    """
    if targets_file is None:
        targets_file = find_latest_targets_file()
    if targets_file is None:
        print("  (Calibration) No targets file found in input/targets/ — skipping.")
        return None

    # Load filtered CSVs (needed for both Excel and JSON paths)
    all_filtered = sorted([
        f for f in FILTERED_DIR.glob("*_filtered.csv")
        if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)
    ])
    if not all_filtered:
        print("  (Calibration) No filtered aging CSVs found — skipping.")
        return None

    packaging_skus = load_packaging_skus()

    # ── Excel path: derive everything from file ──────────────────────────────
    if targets_file.suffix.lower() in (".xlsx", ".xls"):
        curr_csv  = all_filtered[-1]
        prior_csv = all_filtered[-2] if len(all_filtered) >= 2 else None
        print(f"  (Calibration) Reading targets from: {targets_file.name}")
        print(f"  (Calibration) Current: {curr_csv.name}")
        if prior_csv:
            print(f"  (Calibration) Prior:   {prior_csv.name}")

        curr_aged  = load_d2c_aged(curr_csv, packaging_skus)
        prior_aged = load_d2c_aged(prior_csv, packaging_skus) if prior_csv else pd.DataFrame(columns=["Style", "aged_k", "aged_qty"])
        targets_data = load_targets_from_excel(targets_file, curr_aged, prior_aged)
        if targets_data is None:
            return None
        styles_input = targets_data["styles"]
    else:
        # ── JSON path: use manually-entered data ─────────────────────────────
        targets_data = json.loads(targets_file.read_text())
        styles_input = targets_data.get("styles", [])
        if all(s.get("target", 0) == 0 for s in styles_input):
            print(f"  (Calibration) {targets_file.name} has all-zero targets — skipping.")
            return None

        def find_csv(yyyymm: str) -> Optional[Path]:
            return next((f for f in all_filtered if f.name.startswith(yyyymm)), None)

        curr_yyyymm  = targets_data.get("current_yyyymm", "")
        prior_yyyymm = targets_data.get("prior_yyyymm", "")
        curr_csv  = find_csv(curr_yyyymm) or all_filtered[-1]
        prior_csv = find_csv(prior_yyyymm) or (all_filtered[-2] if len(all_filtered) >= 2 else None)
        print(f"  (Calibration) Current: {curr_csv.name}")
        if prior_csv:
            print(f"  (Calibration) Prior:   {prior_csv.name}")

        curr_aged  = load_d2c_aged(curr_csv, packaging_skus)
        prior_aged = load_d2c_aged(prior_csv, packaging_skus) if prior_csv else pd.DataFrame(columns=["Style", "aged_k", "aged_qty"])

    curr_map  = curr_aged.set_index("Style").to_dict("index")
    prior_map = prior_aged.set_index("Style").to_dict("index") if not prior_aged.empty else {}

    month_label  = targets_data.get("month", "")
    prior_label  = targets_data.get("prior_month", "")

    # Build rows
    rows = []
    for s in styles_input:
        style    = str(s["style"])
        strategy = s.get("strategy", "")
        target   = int(s.get("target", 0))
        accomp   = int(s.get("accomp", 0))

        if target == 0:
            continue

        curr_info  = curr_map.get(style, {"aged_k": 0, "aged_qty": 0})
        prior_info = prior_map.get(style, {"aged_k": 0, "aged_qty": 0})

        curr_k   = int(curr_info.get("aged_k", 0))
        curr_qty = int(curr_info.get("aged_qty", 0))
        prior_k  = int(prior_info.get("aged_k", 0))

        mv       = curr_k - prior_k
        acc_pct  = round(accomp / target * 100) if target > 0 else 0
        tgt_pct  = round(target / curr_qty * 100, 1) if curr_qty > 0 else 0.0
        q        = get_quadrant(mv, acc_pct)
        assess   = get_assessment(q, acc_pct)
        is_sil   = style in SILICONE_STYLES

        rows.append({
            "style": style, "strategy": strategy, "target": target, "accomp": accomp,
            "curr_k": curr_k, "curr_qty": curr_qty, "mv": mv,
            "acc_pct": acc_pct, "tgt_pct": tgt_pct, "q": q, "assess": assess,
            "is_sil": is_sil,
        })

    if not rows:
        print("  (Calibration) No rows with non-zero targets — skipping.")
        return None

    # Sort by curr_k descending
    rows.sort(key=lambda r: r["curr_k"], reverse=True)

    # KPI summary
    net_mv         = sum(r["mv"] for r in rows)
    tot_target     = sum(r["target"] for r in rows)
    tot_accomp     = sum(r["accomp"] for r in rows)
    overall_acc    = round(tot_accomp / tot_target * 100) if tot_target else 0
    grew_count     = sum(1 for r in rows if r["mv"] > 0)
    improved_count = sum(1 for r in rows if r["mv"] <= 0)
    n              = len(rows)
    mitigated_count = sum(1 for r in rows if r["mv"] > 0 and r["acc_pct"] >= 100)

    # Table rows HTML
    tbody_html = ""
    tot_curr_k = 0
    tot_curr_qty = 0
    for r in rows:
        tot_curr_k   += r["curr_k"]
        tot_curr_qty += r["curr_qty"]

        q_tag = (
            _tag("critical",  "🔴 Critical")  if r["q"] == "Q1" else
            _tag("mitigated", "🟡 Mitigated") if r["q"] == "Q2" else
            _tag("strong",    "🟢 Strong")    if r["q"] == "Q3" else
            _tag("watch",     "🔵 Watch")
        )
        sil_html = _tag("sil", "SIL") if r["is_sil"] else ""
        mv_sign  = "+" if r["mv"] > 0 else ""
        mv_html  = f'<td style="text-align:right;{_mv_style(r["mv"])}">{mv_sign}${r["mv"]}K</td>'
        acc_html = (
            f'<td class="calib-acc-cell">{_acc_dot(r["acc_pct"])} '
            f'<span style="{_acc_style(r["acc_pct"])}">{r["acc_pct"]}%</span></td>'
        )
        tgt_pct_html = f'<td style="text-align:right;{_tgt_pct_style(r["tgt_pct"])}">{r["tgt_pct"]}%</td>'

        tbody_html += f"""<tr>
            <td><strong>{r['style']}</strong></td>
            <td>{sil_html}</td>
            <td>{r['strategy']}</td>
            <td style="text-align:right">${r['curr_k']}K</td>
            {mv_html}
            <td style="text-align:right">{r['curr_qty']:,}</td>
            <td style="text-align:right">{r['target']/1000:.1f}K</td>
            <td style="text-align:right">{r['accomp']/1000:.1f}K</td>
            {acc_html}
            {tgt_pct_html}
            <td>{q_tag}</td>
            <td style="font-size:11px">{r['assess']}</td>
        </tr>"""

    # Footer
    net_sign = "+" if net_mv > 0 else ""
    ov_acc_style = "color:#e65100;font-weight:700" if overall_acc < 100 else "color:#2e7d32;font-weight:700"
    tfoot_summary = f"{net_sign}${net_mv}K D2C net aging despite {overall_acc}% accomplishment"
    tfoot_html = f"""<tr style="background:#f5f5f5;font-weight:bold;border-top:2px solid #37474f">
        <td colspan="3">TOTAL ({n} styles)</td>
        <td style="text-align:right">${tot_curr_k:,}K</td>
        <td style="text-align:right;{('color:#c62828;font-weight:700' if net_mv > 0 else 'color:#2e7d32;font-weight:700')}">{net_sign}${net_mv}K</td>
        <td style="text-align:right">{tot_curr_qty:,}</td>
        <td style="text-align:right">{tot_target/1000:.1f}K</td>
        <td style="text-align:right">{tot_accomp/1000:.1f}K</td>
        <td class="calib-acc-cell"><span style="{ov_acc_style}">{overall_acc}%</span></td>
        <td></td><td></td>
        <td style="font-size:11px">{tfoot_summary}</td>
    </tr>"""

    # KPI card HTML
    net_mv_sign = "+" if net_mv > 0 else ""
    kpi_html = f"""<div class="calib-mr">
        <div class="calib-mc calib-mc-dark">
            <div class="calib-v">{net_mv_sign}${net_mv}K</div>
            <div class="calib-l">D2C Net Aging Increase<br>{prior_label[:3]} → {month_label[:3]}</div>
        </div>
        <div class="calib-mc calib-mc-amber">
            <div class="calib-v">{overall_acc}%</div>
            <div class="calib-l">Overall Accomplishment<br>{tot_accomp/1000:.1f}K / {tot_target/1000:.1f}K units</div>
        </div>
        <div class="calib-mc calib-mc-red">
            <div class="calib-v">{grew_count} / {n}</div>
            <div class="calib-l">Styles Where Aging<br>Still Grew</div>
        </div>
        <div class="calib-mc calib-mc-green">
            <div class="calib-v">{improved_count} / {n}</div>
            <div class="calib-l">Styles Where Aging<br>Improved</div>
        </div>
    </div>"""

    # Insight box
    insight_html = f"""<div class="calib-insight calib-warn">
        <strong>⚠️ Core Problem:</strong> D2C aging (D2C + EMP + BAD LOT) {"grew" if net_mv > 0 else "changed"} {net_mv_sign}${abs(net_mv)}K despite {overall_acc}% target accomplishment.
        {grew_count} out of {n} tracked styles saw aging increase — including {mitigated_count} that exceeded their target.
        Targets are structurally insufficient to offset the monthly inflow from the 270-365 bucket into Over 365.
    </div>"""

    prior_abbr = prior_label[:3]
    curr_abbr  = month_label[:3]

    # Inner section content (shared by both standalone and embedded)
    section_content = f"""
<div class="calib-wrap">
<style>
  .calib-wrap {{font-family:Arial,sans-serif;color:#1a1a2e;margin-top:32px}}
  .calib-wrap h1{{color:#1e3a5f;margin-bottom:2px;font-size:24px;margin-top:0}}
  .calib-wrap .calib-sub{{color:#666;font-size:13px;margin-bottom:18px}}
  .calib-section{{background:white;padding:24px;margin-bottom:24px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.07)}}
  .calib-section h2{{margin:0 0 4px 0;color:#1e3a5f;font-size:17px}}
  .calib-section .calib-desc{{color:#777;font-size:12px;margin-bottom:14px}}
  .calib-insight{{background:#f0f7ff;border-left:4px solid #2979ff;padding:14px 16px;margin-bottom:20px;border-radius:4px;font-size:13px;line-height:1.55}}
  .calib-insight strong{{color:#1565c0}}
  .calib-warn{{background:#fff3e0;border-left-color:#ff6f00}}
  .calib-warn strong{{color:#e65100}}
  .calib-mr{{display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap}}
  .calib-mc{{flex:1;min-width:130px;padding:14px;border-radius:8px;text-align:center}}
  .calib-v{{font-size:22px;font-weight:800}}
  .calib-l{{font-size:11px;color:#555;margin-top:3px}}
  .calib-mc-red{{background:#ffebee;border-bottom:3px solid #e53935}}.calib-mc-red .calib-v{{color:#c62828}}
  .calib-mc-amber{{background:#fff8e1;border-bottom:3px solid #ffa000}}.calib-mc-amber .calib-v{{color:#e65100}}
  .calib-mc-green{{background:#e8f5e9;border-bottom:3px solid #43a047}}.calib-mc-green .calib-v{{color:#2e7d32}}
  .calib-mc-dark{{background:#eceff1;border-bottom:3px solid #546e7a}}.calib-mc-dark .calib-v{{color:#37474f}}
  .calib-wrap table{{width:100%;border-collapse:collapse;font-size:12px}}
  .calib-wrap th{{background:linear-gradient(135deg,#37474f,#263238);color:white;padding:10px 8px;text-align:left;font-size:11px;text-transform:uppercase;white-space:nowrap;position:sticky;top:0;z-index:1}}
  .calib-wrap td{{padding:8px;border-bottom:1px solid #eee}}
  .calib-wrap tr:hover td{{background:#f5f7fa}}
  .calib-tag{{padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap}}
  .calib-tag-critical{{background:#ffcdd2;color:#c62828}}
  .calib-tag-watch{{background:#fff9c4;color:#f57f17}}
  .calib-tag-strong{{background:#c8e6c9;color:#2e7d32}}
  .calib-tag-mitigated{{background:#ffe0b2;color:#e65100}}
  .calib-tag-sil{{background:#e1bee7;color:#6a1b9a}}
  .calib-acc-dot{{width:10px;height:10px;border-radius:50%;display:inline-block;vertical-align:middle;margin-right:4px}}
  .calib-acc-cell{{white-space:nowrap}}
</style>

<h1>Target Calibration Analysis — {month_label}</h1>
<p class="calib-sub">Are the Aged Monthly Targets correctly set? Accomplishment vs actual D2C aging movement
({prior_abbr}→{curr_abbr}, Over 365 days). D2C scope includes D2C, EMP, and BAD LOT categories.</p>

{kpi_html}

{insight_html}

<div class="calib-section">
  <h2>Style Detail: Accomplishment vs. Movement</h2>
  <p class="calib-desc">Sorted by {curr_abbr} D2C aged value (largest exposure first). Movement = actual change in D2C Over 365
  value from {prior_label} to {month_label}. Quadrant: 🔴 aging grew + under target | 🟡 aging grew + target met |
  🟢 aging improved + target met | 🔵 aging improved + under target.</p>
  <div style="overflow-x:auto">
    <table>
      <thead><tr>
        <th>Style</th>
        <th>Sil.</th>
        <th>Strategy</th>
        <th style="text-align:right">{curr_abbr} D2C<br>Aged $K</th>
        <th style="text-align:right">D2C<br>Movement $K</th>
        <th style="text-align:right">{curr_abbr} D2C<br>Aged Qty</th>
        <th style="text-align:right">Target</th>
        <th style="text-align:right">Accomp</th>
        <th>Acc %</th>
        <th style="text-align:right">Tgt as %<br>of Aged Qty</th>
        <th>Quadrant</th>
        <th>Target Assessment</th>
      </tr></thead>
      <tbody>{tbody_html}</tbody>
      <tfoot>{tfoot_html}</tfoot>
    </table>
  </div>
  <p style="font-size:11px;color:#777;margin-top:14px;line-height:1.6">
    <strong>Target (monthly):</strong> Q2 AGING FINAL ÷ 3, summed across all countries from the targets Excel file.
    &nbsp;|&nbsp;
    <strong>Target as % of Aged Qty:</strong> Monthly target ÷ current month D2C Over-365 Qty × 100.
    &nbsp;|&nbsp;
    <strong>Accomplishment:</strong> Max(0, prior month Over-365 qty − current month Over-365 qty) per style — net units that left the aged bucket.
  </p>
</div>
</div>
"""

    if section_only:
        return section_content

    # Full standalone page
    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Target Calibration Analysis — {month_label}</title>
<style>body{{padding:20px;background:#f4f5f7}}</style>
</head>
<body>
{section_content}
</body>
</html>"""
    return full_html


# ── Main entry point ──────────────────────────────────────────────────────────

def main():
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    targets_file = find_latest_targets_file()
    if targets_file is None:
        print("ERROR: No targets_*.json file found in input/targets/")
        print("       Create and fill in input/targets/targets_<month>.json first.")
        sys.exit(1)

    print(f"Using targets file: {targets_file.name}")

    html = build_calibration_html(targets_file=targets_file, section_only=False)
    if html is None:
        sys.exit(0)

    # Determine month label for filename
    if targets_file.suffix.lower() in (".xlsx", ".xls"):
        m = re.match(r"(\d{4})(\d{2})_", targets_file.name)
        if m:
            import calendar
            month_label = f"{calendar.month_name[int(m.group(2))]}{m.group(1)}"
        else:
            month_label = targets_file.stem
    else:
        targets_data = json.loads(targets_file.read_text())
        month_label  = targets_data.get("month", targets_file.stem).replace(" ", "").replace(",", "")

    # ── Save standalone HTML ──────────────────────────────────────────────────
    mon_slug = month_label.replace(" ", "").replace(",", "")
    out_path = REPORTS_DIR / f"target_calibration_{mon_slug}.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"\nStandalone report saved: {out_path.relative_to(ROOT)}")

    # ── Append section to aging_evacuation_analysis.html ─────────────────────
    evac_path = REPORTS_DIR / "aging_evacuation_analysis.html"
    if evac_path.exists():
        section = build_calibration_html(targets_file=targets_file, section_only=True)
        evac_html = evac_path.read_text(encoding="utf-8")
        if "</body>" in evac_html:
            evac_html = evac_html.replace("</body>", f"{section}\n</body>")
        else:
            evac_html += section
        evac_path.write_text(evac_html, encoding="utf-8")
        print(f"Appended calibration section to: {evac_path.relative_to(ROOT)}")
        webbrowser.open(evac_path.as_uri())
    else:
        print(f"Note: {evac_path.name} not found — run generate_report2.py first.")
        webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
