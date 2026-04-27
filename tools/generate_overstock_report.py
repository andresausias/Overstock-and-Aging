"""
generate_overstock_report.py
-----------------------------
Generates the 15-section Overstock Inventory Analysis HTML report from a
Rolling Forecast Excel file.

Two tabs are combined:
  - "Rolling forecast {Month}PP by SKU"   (all active channels)
  - "Disco&NB {Month}PP by SKU"           (discontinued + national brands)

The header row is at row 9 (0-indexed row 8). The sheet duplicates columns
after ~column 103 — only the first 103 columns are used.

Usage:
    python tools/generate_overstock_report.py
    python tools/generate_overstock_report.py path/to/Rolling_forecast_April_PP_2026.xlsx

Output:
    .tmp/overstock/overstock_<start>_<end>_report.html
    (self-contained, Chart.js v4.4.0 via CDN)
"""

import sys
import re
import json
import webbrowser
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd

ROOT = Path(__file__).parent.parent
INPUT_DIR  = ROOT / "input" / "rolling_forecast"
OUTPUT_DIR = ROOT / ".tmp" / "overstock"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── Config ─────────────────────────────────────────────────────────────────────

SILICONE_STYLES = {
    '10400', '18437', '42001', '42004', '41005', '42024', '42066',
    '42075', '51001', '51009', '52002', '52005', '54008', '55021',
    '98099', '77001', '77002',
}

MONTH_COLS = [
    'MAR-26', 'APR-26', 'MAY-26', 'JUN-26', 'JUL-26', 'AUG-26',
    'SEP-26', 'OCT-26', 'NOV-26', 'DEC-26', 'JAN-27', 'FEB-27', 'MAR-27',
]
MONTH_LABELS = [
    'Mar 26', 'Apr 26', 'May 26', 'Jun 26', 'Jul 26', 'Aug 26',
    'Sep 26', 'Oct 26', 'Nov 26', 'Dec 26', 'Jan 27', 'Feb 27', 'Mar 27',
]

CHANNEL_COLORS = ['#2196F3', '#FF9800', '#4CAF50', '#9C27B0', '#E91E63',
                  '#00BCD4', '#795548', '#607D8B', '#CDDC39']
CATEGORY_COLORS = ['#2196F3', '#FF9800', '#4CAF50', '#9C27B0', '#E91E63',
                   '#00BCD4', '#795548', '#FF5722', '#3F51B5', '#009688',
                   '#CDDC39', '#607D8B']
COUNTRY_COLORS = {
    'US': '#2196F3', 'CANADA': '#FF9800', 'UK': '#4CAF50',
    'AU': '#9C27B0', 'CHENKUN': '#E53935', 'SAUDI ARABIA': '#00BCD4',
}
COUNTRY_ORDER = ['US', 'CANADA', 'UK', 'AU', 'CHENKUN', 'SAUDI ARABIA']


# ── Helpers ────────────────────────────────────────────────────────────────────

def fmt_val(v):
    v = float(v)
    if abs(v) >= 1e6:
        return f'${v/1e6:.2f}M'
    return f'${v/1e3:.1f}K'

def fmt_units(v):
    v = float(v)
    if abs(v) >= 1e6:
        return f'{v/1e6:.2f}M'
    if abs(v) >= 1000:
        return f'{v/1e3:.1f}K'
    return str(int(v))

def fmt_int(v):
    return f'{int(v):,}'

def simplify_channel(ch):
    ch = str(ch)
    if 'Amazon' in ch:            return 'Amazon'
    if 'Wholesale' in ch:         return 'Wholesale'
    if ch == 'TSD':               return 'TSD'
    if ch == 'Revel':             return 'Revel'
    if 'Kenz' in ch or ('NB' in ch and 'D2C' in ch): return 'Kenz/NB'
    if 'Kenz NB' == ch:          return 'Kenz/NB'
    if 'Disco' in ch or 'Discontinued' in ch: return 'Discontinued'
    if 'Accesories' in ch or 'Accessories' in ch: return 'Accessories'
    if 'TV' in ch:                return 'TV Shows'
    return 'D2C'

def heatmap_color(value, global_max):
    if global_max == 0:
        return 'rgba(33,150,243,0.15)'
    intensity = min(value / global_max, 1.0)
    r = int(33  + (244 - 33)  * intensity)
    g = int(150 + (67  - 150) * intensity)
    b = int(243 + (54  - 243) * intensity)
    alpha = max(0.15, intensity * 0.85)
    return f'rgba({r},{g},{b},{alpha:.2f})'


# ── Load & process ─────────────────────────────────────────────────────────────

def find_input_file() -> Path:
    """Find the Rolling Forecast Excel in input/overstock/ or accept CLI arg."""
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if not p.exists():
            raise FileNotFoundError(f"File not found: {p}")
        return p

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    candidates = (list(INPUT_DIR.glob("Rolling_forecast*.xlsx")) +
                  list(INPUT_DIR.glob("rolling_forecast*.xlsx")) +
                  list(INPUT_DIR.glob("Rolling forecast*.xlsx")))
    if not candidates:
        raise FileNotFoundError(
            f"No Rolling Forecast file found in {INPUT_DIR}.\n"
            "Place the file there or pass the path as a CLI argument."
        )
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def detect_tab_names(xl: pd.ExcelFile) -> tuple[str, str]:
    """Auto-detect the two required tab names."""
    sheets = xl.sheet_names
    forecast_tab = next((s for s in sheets if 'Rolling forecast' in s or 'rolling forecast' in s.lower()), None)
    disco_tab    = next((s for s in sheets if 'Disco' in s or 'disco' in s.lower()), None)
    if not forecast_tab:
        raise ValueError(f"Cannot find 'Rolling forecast' tab. Available: {sheets}")
    if not disco_tab:
        raise ValueError(f"Cannot find 'Disco&NB' tab. Available: {sheets}")
    return forecast_tab, disco_tab


def detect_month_cols(df: pd.DataFrame, prefix: str) -> list[str]:
    """Find all columns starting with prefix (e.g. 'OVSTK ' or 'EOM '),
    excluding duplicates (no '.1' suffix)."""
    return [c for c in df.columns if str(c).startswith(prefix) and '.1' not in str(c)]


def load_data(filepath: Path) -> tuple[pd.DataFrame, list[str], list[str], str, str]:
    """Load and combine both tabs. Returns (df, ovstk_cols, eom_cols, start_label, end_label)."""
    xl = pd.ExcelFile(filepath)
    forecast_tab, disco_tab = detect_tab_names(xl)

    df1 = xl.parse(forecast_tab, header=8).iloc[:, :103].copy()
    df2 = xl.parse(disco_tab,    header=8).iloc[:, :103].copy()
    df1['Source'] = 'Rolling Forecast'
    df2['Source'] = 'Disco&NB'
    df = pd.concat([df1, df2], ignore_index=True)

    # Style cleanup
    df['Style'] = df['Style'].astype(str).str.replace('.0', '', regex=False).str.strip()

    # Numeric cols
    ovstk_cols = detect_month_cols(df, 'OVSTK ')
    eom_cols   = detect_month_cols(df, 'EOM ')
    for c in ovstk_cols + eom_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # Landed cost
    lc_col = 'LC - DDP (IS-OTB-Forecast)'
    if lc_col not in df.columns:
        lc_col = next((c for c in df.columns if 'LC' in str(c) and 'DDP' in str(c)), None)
    df['LC_DDP'] = pd.to_numeric(df[lc_col], errors='coerce').fillna(0) if lc_col else 0.0

    # Channel group
    df['Channel_Group'] = df['Channel'].apply(simplify_channel)

    # Silicone flag
    df['is_silicone'] = df['Style'].isin(SILICONE_STYLES)

    # Determine start/end month labels from column names
    # ovstk_cols e.g. ['OVSTK MAR-26', ..., 'OVSTK MAR-27']
    def col_label(c):
        raw = c.split(' ', 1)[1] if ' ' in c else c  # e.g. 'MAR-26'
        parts = raw.split('-')
        if len(parts) == 2:
            mon, yr = parts
            return f"{mon.capitalize()} {yr}"
        return raw
    start_label = col_label(ovstk_cols[0])  if ovstk_cols else 'Mar 26'
    end_label   = col_label(ovstk_cols[-1]) if ovstk_cols else 'Mar 27'

    return df, ovstk_cols, eom_cols, start_label, end_label


# ── Aggregations ───────────────────────────────────────────────────────────────

def agg_trajectory(df, ovstk_cols, eom_cols):
    traj_units = [df[c].clip(lower=0).sum() for c in ovstk_cols]
    traj_value = [(df[c].clip(lower=0) * df['LC_DDP']).sum() for c in ovstk_cols]
    eom_totals = [df[c].clip(lower=0).sum() for c in eom_cols]
    traj_pct   = [
        (u / e * 100) if e > 0 else 0
        for u, e in zip(traj_units, eom_totals)
    ]
    return traj_units, traj_value, traj_pct, eom_totals


def agg_by_country(df, ovstk_cols, eom_cols):
    col0, col_last = ovstk_cols[0], ovstk_cols[-1]
    results = {}
    for country in COUNTRY_ORDER:
        sub = df[df['Country'].astype(str).str.upper().str.strip() == country]
        traj_val  = [(sub[c].clip(lower=0) * sub['LC_DDP']).sum() for c in ovstk_cols]
        traj_units = [sub[c].clip(lower=0).sum() for c in ovstk_cols]
        results[country] = {'value': traj_val, 'units': traj_units}
    return results


def agg_by_channel(df, col0):
    grp = df.groupby('Channel_Group').apply(
        lambda g: pd.Series({
            'units': g[col0].clip(lower=0).sum(),
            'value': (g[col0].clip(lower=0) * g['LC_DDP']).sum(),
        })
    ).reset_index()
    grp = grp.sort_values('value', ascending=False).reset_index(drop=True)
    return grp


def agg_by_category(df, col0):
    grp = df.groupby('Category').apply(
        lambda g: pd.Series({
            'value': (g[col0].clip(lower=0) * g['LC_DDP']).sum(),
        })
    ).reset_index()
    grp = grp.sort_values('value', ascending=False).head(12).reset_index(drop=True)
    return grp


def agg_by_style(df, col0, col_last):
    grp = df.groupby('Style').apply(
        lambda g: pd.Series({
            'units_mar':   g[col0].clip(lower=0).sum(),
            'value_mar':   (g[col0].clip(lower=0) * g['LC_DDP']).sum(),
            'units_mar27': g[col_last].clip(lower=0).sum(),
            'value_mar27': (g[col_last].clip(lower=0) * g['LC_DDP']).sum(),
            'eom_mar':     g[col0].clip(lower=0).sum(),
            'is_silicone': g['is_silicone'].any(),
        })
    ).reset_index()
    grp['evac_pct'] = np.where(
        grp['units_mar'] > 0,
        (grp['units_mar'] - grp['units_mar27']) / grp['units_mar'] * 100,
        0
    )
    return grp


def agg_by_brand(df, col0):
    grp = df.groupby('Brand').apply(
        lambda g: pd.Series({
            'value': (g[col0].clip(lower=0) * g['LC_DDP']).sum(),
        })
    ).reset_index()
    grp = grp.sort_values('value', ascending=False).head(12).reset_index(drop=True)
    return grp


def top15_styles(style_df, channel_group=None, df_full=None, col0=None, col_last=None):
    """Top 15 styles by value for a channel group (or all channels)."""
    if channel_group and df_full is not None:
        sub = df_full[df_full['Channel_Group'] == channel_group]
        grp = agg_by_style(sub, col0, col_last)
    else:
        grp = style_df
    top = grp.sort_values('value_mar', ascending=False).head(15)
    return [
        {
            'style':      r['Style'],
            'units':      int(r['units_mar']),
            'value':      round(r['value_mar'], 0),
            'units_mar27': int(r['units_mar27']),
            'value_mar27': round(r['value_mar27'], 0),
            'is_silicone': bool(r['is_silicone']),
            'eom':         int(r['eom_mar']),
            'evac_pct':    round(r['evac_pct'], 1),
        }
        for _, r in top.iterrows()
    ]


def evac_data(style_df):
    """From top 25 by value, split into best/worst evacuators."""
    top25 = style_df.sort_values('value_mar', ascending=False).head(25)
    best10 = top25.sort_values('evac_pct', ascending=False).head(10)
    worst10 = top25.sort_values('evac_pct', ascending=True).head(10).sort_values('evac_pct', ascending=False)

    def to_list(df_):
        return [
            {
                'style':      r['Style'],
                'pct':        round(r['evac_pct'], 1),
                'value_mar':  round(r['value_mar'], 0),
                'units_mar':  int(r['units_mar']),
                'units_mar27': int(r['units_mar27']),
                'is_silicone': bool(r['is_silicone']),
            }
            for _, r in df_.iterrows()
        ]

    slow_list = to_list(worst10)
    # For slow chart, show 100 - evac_pct (% remaining)
    for item in slow_list:
        item['pct'] = round(100 - item['pct'], 1)

    return to_list(best10), slow_list


def slowest_table(style_df, n=25):
    """Top 25 slowest evacuators with >500 units."""
    filtered = style_df[style_df['units_mar'] > 500].sort_values('evac_pct', ascending=True).head(n)
    return [
        {
            'style':     r['Style'],
            'units_mar': int(r['units_mar']),
            'value_mar': round(r['value_mar'], 0),
            'evac_pct':  round(r['evac_pct'], 1),
            'value_chg': round(r['value_mar27'] - r['value_mar'], 0),
            'is_silicone': bool(r['is_silicone']),
        }
        for _, r in filtered.iterrows()
    ]


# ── HTML generation ────────────────────────────────────────────────────────────

def j(val):
    """JSON-safe repr for embedding in JS."""
    return json.dumps(val)


def build_html(filepath: Path) -> str:
    df, ovstk_cols, eom_cols, start_label, end_label = load_data(filepath)

    col0     = ovstk_cols[0]
    col_last = ovstk_cols[-1]

    # ── Trajectory ────────────────────────────────────────────────────────────
    traj_units, traj_value, traj_pct, eom_totals = agg_trajectory(df, ovstk_cols, eom_cols)

    total_ovstk_value = traj_value[0]
    total_ovstk_units = traj_units[0]
    total_eom_units   = eom_totals[0]
    ovstk_pct_inv     = traj_pct[0]
    proj_value        = traj_value[-1]
    proj_units        = traj_units[-1]
    proj_pct_eom      = traj_pct[-1]
    val_reduction_pct = (total_ovstk_value - proj_value) / total_ovstk_value * 100 if total_ovstk_value else 0
    unit_reduction_pct= (total_ovstk_units - proj_units) / total_ovstk_units * 100 if total_ovstk_units else 0

    # ── Styles with overstock ─────────────────────────────────────────────────
    n_styles_ovstk = int((df.groupby('Style')[col0].sum() > 0).sum())

    # ── Country ───────────────────────────────────────────────────────────────
    country_data = agg_by_country(df, ovstk_cols, eom_cols)

    # ── Channel ───────────────────────────────────────────────────────────────
    channel_df = agg_by_channel(df, col0)
    ch_labels  = channel_df['Channel_Group'].tolist()
    ch_values  = [round(v, 0) for v in channel_df['value'].tolist()]
    ch_units   = [int(u) for u in channel_df['units'].tolist()]

    # ── Category ─────────────────────────────────────────────────────────────
    cat_df     = agg_by_category(df, col0)
    cat_labels = cat_df['Category'].tolist()
    cat_values = [round(v, 0) for v in cat_df['value'].tolist()]

    # ── Style aggregations ────────────────────────────────────────────────────
    style_df = agg_by_style(df, col0, col_last)

    top15_all    = top15_styles(style_df)
    top15_d2c    = top15_styles(None, 'D2C',    df, col0, col_last)
    top15_amazon = top15_styles(None, 'Amazon', df, col0, col_last)

    fast_evac, slow_evac = evac_data(style_df)
    slow_tbl = slowest_table(style_df)

    # ── Top 25 value table ────────────────────────────────────────────────────
    top25_tbl = style_df.sort_values('value_mar', ascending=False).head(25)

    # ── Brand ─────────────────────────────────────────────────────────────────
    brand_df     = agg_by_brand(df, col0)
    brand_labels = brand_df['Brand'].tolist()
    brand_values = [round(v, 0) for v in brand_df['value'].tolist()]

    # ── Heatmap ───────────────────────────────────────────────────────────────
    hm_global_max = max(
        country_data[c]['units'][i]
        for c in COUNTRY_ORDER
        for i in range(len(ovstk_cols))
    )

    generated_date = datetime.now().strftime("%b %d, %Y")

    # ── Build month labels from actual column names ───────────────────────────
    def col_to_label(c):
        raw = c.split(' ', 1)[1] if ' ' in c else c
        parts = raw.split('-')
        if len(parts) == 2:
            mon, yr = parts
            return f"{mon.capitalize()} {yr[-2:]}"
        return raw
    month_labels_js = [col_to_label(c) for c in ovstk_cols]

    # slugify for filename
    def slug(s): return re.sub(r'[^a-z0-9]+', '', s.lower().replace(' ', ''))
    start_slug = slug(start_label)
    end_slug   = slug(end_label)

    # ── Heatmap HTML ──────────────────────────────────────────────────────────
    hm_rows_html = ""
    for country in COUNTRY_ORDER:
        cdata = country_data[country]
        units_start = cdata['units'][0]
        units_end   = cdata['units'][-1]
        reduction_pct = (units_start - units_end) / units_start * 100 if units_start else 0
        red_cls   = 'style="color:#2e7d32;font-weight:700"' if reduction_pct > 0 else 'style="color:#c62828;font-weight:700"'
        red_sign  = '+' if reduction_pct < 0 else ''  # negative pct = more stuck (bad)
        hm_rows_html += f'<tr><td style="background:#4a5568;color:#fff;font-weight:600">{country}</td>\n'
        for i, u in enumerate(cdata['units']):
            bg = heatmap_color(u, hm_global_max)
            hm_rows_html += f'<td style="background:{bg};color:#333">{int(u):,}</td>\n'
        hm_rows_html += f'<td {red_cls}>{red_sign}{abs(reduction_pct):.1f}%</td></tr>\n'

    # ── Heatmap header ────────────────────────────────────────────────────────
    hm_header_html = '<th>Country</th>\n'
    for lbl in month_labels_js:
        hm_header_html += f'<th>{lbl}</th>\n'
    hm_header_html += '<th>Reduction</th>\n'

    # ── Top 25 value table HTML ───────────────────────────────────────────────
    top25_rows_html = ""
    for idx, (_, r) in enumerate(top25_tbl.iterrows(), 1):
        v_chg = r['value_mar27'] - r['value_mar']
        chg_cls = 'change-neg' if v_chg < 0 else 'change-pos'
        chg_sign = '' if v_chg >= 0 else ''
        si_badge = '<span class="silicone-badge">S</span>' if r['is_silicone'] else '<span class="no-badge">NO</span>'
        top25_rows_html += (
            f'<tr><td>{idx}</td>'
            f'<td><strong>{r["Style"]}</strong></td>'
            f'<td>{fmt_int(r["units_mar"])}</td>'
            f'<td>{fmt_val(r["value_mar"])}</td>'
            f'<td>{fmt_int(r["units_mar27"])}</td>'
            f'<td>{fmt_val(r["value_mar27"])}</td>'
            f'<td class="{chg_cls}">{fmt_val(v_chg)}</td>'
            f'<td>{r["evac_pct"]:.1f}%</td>'
            f'<td>{si_badge}</td></tr>\n'
        )

    # ── Slowest evac table HTML ───────────────────────────────────────────────
    slow_rows_html = ""
    for idx, item in enumerate(slow_tbl, 1):
        ep = item['evac_pct']
        row_bg = 'background:#fff0f0' if ep < 10 else ('background:#fff8e1' if ep < 30 else '')
        row_style = f' style="{row_bg}"' if row_bg else ''
        chg_cls = 'change-neg' if item['value_chg'] < 0 else 'change-pos'
        si_badge = '<span class="silicone-badge">S</span>' if item['is_silicone'] else '<span class="no-badge">NO</span>'
        chg_fmt = fmt_val(item['value_chg']) if abs(item['value_chg']) > 1 else '$0'
        slow_rows_html += (
            f'<tr{row_style}>'
            f'<td>{idx}</td>'
            f'<td><strong>{item["style"]}</strong></td>'
            f'<td>{fmt_int(item["units_mar"])}</td>'
            f'<td>{fmt_val(item["value_mar"])}</td>'
            f'<td><strong>{ep:.1f}%</strong></td>'
            f'<td class="{chg_cls}">{chg_fmt}</td>'
            f'<td>{si_badge}</td></tr>\n'
        )

    # ── Country trajectory datasets ───────────────────────────────────────────
    country_datasets = []
    for country in COUNTRY_ORDER:
        cdata = country_data[country]
        country_datasets.append({
            "label": country,
            "data":  [round(v, 0) for v in cdata['value']],
            "borderColor": COUNTRY_COLORS.get(country, '#999'),
            "tension": 0.3,
            "pointRadius": 3,
            "fill": False,
        })

    # ── Country bar chart (Mar 26 vs Mar 27) ─────────────────────────────────
    country_mar26 = [round(country_data[c]['value'][0], 0) for c in COUNTRY_ORDER]
    country_mar27 = [round(country_data[c]['value'][-1], 0) for c in COUNTRY_ORDER]

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Overstock Inventory Analysis – {start_label} to {end_label}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f5f5f5;color:#333;padding:20px}}
h1{{color:#2c5aa0;text-align:center;margin-bottom:8px;font-size:28px}}
.subtitle{{text-align:center;color:#666;margin-bottom:30px;font-size:14px}}
.section{{background:#fff;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.1);padding:20px;margin-bottom:24px}}
.section h2{{color:#2c5aa0;font-size:18px;margin-bottom:14px;border-bottom:2px solid #e0e0e0;padding-bottom:8px}}
.section h3{{color:#444;font-size:15px;margin-bottom:10px}}
.chart-wrapper{{position:relative;height:400px}}
.chart-wrapper-lg{{position:relative;height:480px}}
.chart-wrapper-sm{{position:relative;height:300px}}
.insight-box{{background:#f0f7ff;border-left:4px solid #2196F3;padding:12px 16px;margin-top:16px;border-radius:0 4px 4px 0;font-size:13px;line-height:1.6}}
.insight-box strong{{color:#1565C0}}
.metrics-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:20px}}
.metric-card{{background:linear-gradient(135deg,#f8f9fa,#e9ecef);border-radius:8px;padding:16px;text-align:center;border:1px solid #dee2e6}}
.metric-card .value{{font-size:28px;font-weight:700;color:#2c5aa0}}
.metric-card .label{{font-size:12px;color:#666;margin-top:4px}}
.metric-card.green .value{{color:#2e7d32}}
.metric-card.red .value{{color:#c62828}}
.metric-card.amber .value{{color:#e65100}}
.styled-table{{width:100%;border-collapse:collapse;font-size:13px}}
.styled-table thead{{background:linear-gradient(135deg,#4a5568,#2d3748);color:#fff}}
.styled-table th{{padding:10px 12px;text-align:left;font-weight:600}}
.styled-table td{{padding:8px 12px;border-bottom:1px solid #e2e8f0}}
.styled-table tbody tr:nth-child(even){{background:#f7fafc}}
.styled-table tbody tr:hover{{background:#edf2f7}}
.silicone-badge{{background:#9c27b0;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}}
.no-badge{{background:#9e9e9e;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px}}
.change-pos{{color:#c62828;font-weight:600}}
.change-neg{{color:#2e7d32;font-weight:600}}
.section-banner{{background:linear-gradient(135deg,#1a237e,#283593);color:#fff;padding:16px 20px;border-radius:8px;margin-bottom:16px}}
.section-banner h2{{color:#fff;border:none;padding:0;margin:0;font-size:20px}}
.section-banner p{{color:rgba(255,255,255,0.8);font-size:13px;margin-top:4px}}
.heatmap-table{{width:100%;border-collapse:collapse;font-size:12px;text-align:center}}
.heatmap-table th{{padding:8px 6px;background:#4a5568;color:#fff;font-weight:600}}
.heatmap-table td{{padding:8px 6px;border:1px solid #e2e8f0;font-weight:600}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
@media(max-width:900px){{.two-col{{grid-template-columns:1fr}}}}
.summary-table{{width:100%;border-collapse:collapse;font-size:13px}}
.summary-table th{{background:#f1f5f9;padding:10px 12px;text-align:left;font-weight:600;color:#475569;border-bottom:2px solid #cbd5e1}}
.summary-table td{{padding:9px 12px;border-bottom:1px solid #e2e8f0}}
.summary-table tr:hover{{background:#f8fafc}}
</style>
</head>
<body>

<h1>Overstock Inventory Analysis</h1>
<p class="subtitle">Rolling Forecast: {start_label} to {end_label} | Combined: All Channels + Disco &amp; NB | Generated {generated_date}</p>

<div class="section">
<h2>1. Key Metrics Summary</h2>
<div class="metrics-grid">
<div class="metric-card"><div class="value">{fmt_val(total_ovstk_value)}</div><div class="label">Total Overstock Value ({start_label})</div></div>
<div class="metric-card"><div class="value">{fmt_units(total_ovstk_units)}</div><div class="label">Total Overstock Units ({start_label})</div></div>
<div class="metric-card amber"><div class="value">{ovstk_pct_inv:.1f}%</div><div class="label">% of Total Inventory</div></div>
<div class="metric-card"><div class="value">{n_styles_ovstk:,}</div><div class="label">Styles with Overstock</div></div>
<div class="metric-card green"><div class="value">-{val_reduction_pct:.1f}%</div><div class="label">Projected Value Reduction by {end_label}</div></div>
<div class="metric-card green"><div class="value">-{unit_reduction_pct:.1f}%</div><div class="label">Projected Unit Reduction by {end_label}</div></div>
<div class="metric-card"><div class="value">{fmt_val(proj_value)}</div><div class="label">Projected Value ({end_label})</div></div>
<div class="metric-card"><div class="value">{fmt_units(proj_units)}</div><div class="label">Projected Units ({end_label})</div></div>
</div>
<table class="summary-table">
<thead><tr><th>Metric</th><th>{start_label}</th><th>{end_label} (Projected)</th><th>Change</th></tr></thead>
<tbody>
<tr><td>Overstock Value</td><td><strong>{fmt_val(total_ovstk_value)}</strong></td><td>{fmt_val(proj_value)}</td>
<td class="change-neg">&darr; {fmt_val(total_ovstk_value - proj_value)} (-{val_reduction_pct:.1f}%)</td></tr>
<tr><td>Overstock Units</td><td><strong>{fmt_units(total_ovstk_units)}</strong></td><td>{fmt_units(proj_units)}</td>
<td class="change-neg">&darr; {fmt_units(total_ovstk_units - proj_units)} (-{unit_reduction_pct:.1f}%)</td></tr>
<tr><td>Total EOM Inventory</td><td><strong>{fmt_units(eom_totals[0])}</strong></td><td>{fmt_units(eom_totals[-1])}</td>
<td class="change-neg">&darr; {fmt_units(eom_totals[0] - eom_totals[-1])}</td></tr>
<tr><td>Overstock % of Inventory</td><td><strong>{traj_pct[0]:.1f}%</strong></td><td>{traj_pct[-1]:.1f}%</td>
<td class="change-pos">&uarr; {traj_pct[-1] - traj_pct[0]:.1f} pp</td></tr>
</tbody></table>
<div class="insight-box"><strong>Key Insight:</strong> Total overstock stands at {fmt_val(total_ovstk_value)} ({fmt_units(total_ovstk_units)} units), representing {ovstk_pct_inv:.1f}% of total EOM inventory. The rolling forecast projects a -{val_reduction_pct:.1f}% value reduction over the forecast period. Because total inventory declines faster than overstock, the <em>share</em> of overstock rises from {traj_pct[0]:.1f}% to {traj_pct[-1]:.1f}% — a structural concentration risk.</div>
</div>

<div class="section">
<h2>2. Overstock Forecast Trajectory – Units &amp; Value</h2>
<div class="chart-wrapper"><canvas id="chart_trajectory"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> Overstock units decline from {fmt_units(total_ovstk_units)} to {fmt_units(proj_units)} (-{unit_reduction_pct:.1f}%), while value drops from {fmt_val(total_ovstk_value)} to {fmt_val(proj_value)} (-{val_reduction_pct:.1f}%). Post-October, the curve flattens — residual overstock becomes increasingly sticky.</div>
</div>

<div class="section">
<h2>3. Overstock as % of Total Inventory</h2>
<div class="chart-wrapper-sm"><canvas id="chart_pct"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> While absolute overstock declines, overstock as a share of total inventory rises from {traj_pct[0]:.1f}% to {traj_pct[-1]:.1f}% by {end_label}. Inline inventory exits faster than overstock, leaving an increasingly concentrated overstock position.</div>
</div>

<div class="section">
<h2>4. Overstock by Country – Value ($)</h2>
<div class="chart-wrapper"><canvas id="chart_country"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> US dominates with {fmt_val(country_data['US']['value'][0])} ({country_data['US']['value'][0]/total_ovstk_value*100:.1f}% of total). CHENKUN is the only country where overstock value grows over the forecast period.</div>
</div>

<div class="section">
<h2>5. Overstock by Channel – {start_label}</h2>
<div class="chart-wrapper"><canvas id="chart_channel_val"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> {ch_labels[0] if ch_labels else 'D2C'} holds the largest overstock position at {fmt_val(ch_values[0]) if ch_values else '$0'}. The top 3 channels represent the majority of total overstock value.</div>
</div>

<div class="section">
<h2>6. Overstock by Product Category – {start_label}</h2>
<div class="chart-wrapper"><canvas id="chart_category"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> {cat_labels[0] if cat_labels else 'Bra'} accounts for the largest overstock position at {fmt_val(cat_values[0]) if cat_values else '$0'}. The top 3 categories represent over 60% of total overstock value.</div>
</div>

<div class="section-banner"><h2>Style-Level Analysis</h2><p>Deep dive into overstock by individual style – all channels combined</p></div>

<div class="section">
<h2>7. Top 15 Overstock Styles – All Channels (by Value)</h2>
<div class="chart-wrapper-lg"><canvas id="chart_top15_all"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> The top 15 styles by overstock value drive the majority of the overstock position. Silicone band styles (purple) appear across multiple ranking positions.</div>
</div>

<div class="section">
<h2>8. Top 15 Overstock Styles – D2C Channel</h2>
<div class="chart-wrapper-lg"><canvas id="chart_top15_d2c"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> D2C overstock is concentrated in a small number of high-volume styles. Focus clearance efforts here for maximum unit impact.</div>
</div>

<div class="section">
<h2>9. Top 15 Overstock Styles – Amazon Channel</h2>
<div class="chart-wrapper-lg"><canvas id="chart_top15_amazon"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> Amazon overstock includes styles with very low projected evacuation rates, indicating limited organic sell-through at current pricing or velocity.</div>
</div>

<div class="section-banner"><h2>Evacuation &amp; Movement Analysis</h2><p>Which styles are projected to reduce overstock fastest — and which are stuck</p></div>

<div class="section">
<h2>10. Projected Overstock Evacuation – Top 25 Styles by Value ({start_label} → {end_label})</h2>
<div class="two-col">
<div><h3>Best Evacuators – % Cleared by {end_label}</h3><div class="chart-wrapper"><canvas id="chart_fast_evac"></canvas></div></div>
<div><h3>Most Stuck – % of Overstock Remaining by {end_label}</h3><div class="chart-wrapper"><canvas id="chart_slow_evac"></canvas></div></div>
</div>
<div class="insight-box"><strong>Key Insight:</strong> Among the top 25 highest-value overstock styles, evacuation rates vary widely. Styles with &gt;100% remaining have growing overstock positions — they require direct intervention.</div>
</div>

<div class="section"><h2>11. Top 25 Styles by Overstock Value – {start_label}</h2>
<table class="styled-table"><thead><tr><th>#</th><th>Style</th><th>Units ({start_label})</th><th>Value ({start_label})</th><th>Units ({end_label})</th><th>Value ({end_label})</th><th>Value Change</th><th>Evac %</th><th>Silicone</th></tr></thead><tbody>
{top25_rows_html}
</tbody></table>
<div class="insight-box"><strong>Key Insight:</strong> The top 25 styles by value represent the core of the overstock challenge. Focus evacuation efforts on the highest-value, lowest-evacuation-rate combinations for maximum impact.</div></div>

<div class="section"><h2>12. Top 25 Styles with Slowest Projected Evacuation</h2>
<table class="styled-table"><thead><tr><th>#</th><th>Style</th><th>Units ({start_label})</th><th>Value ({start_label})</th><th>Evac %</th><th>Value Change</th><th>Silicone</th></tr></thead><tbody>
{slow_rows_html}
</tbody></table>
<div class="insight-box"><strong>Key Insight:</strong> Styles with &lt;10% projected evacuation (highlighted in red) may require aggressive intervention — deeper discounts, new channel placement, or liquidation consideration.</div></div>

<div class="section"><h2>13. Country × Month Overstock Heatmap (Units)</h2>
<table class="heatmap-table"><thead><tr>{hm_header_html}</tr></thead><tbody>
{hm_rows_html}
</tbody></table>
<div class="insight-box"><strong>Key Insight:</strong> Darker cells = higher overstock. CHENKUN stands out as the only country showing overstock growth. US has the largest absolute position but projects the strongest absolute reduction.</div></div>

<div class="section">
<h2>14. Country Overstock Value Trajectory</h2>
<div class="chart-wrapper"><canvas id="chart_country_traj"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> The US drives the bulk of overstock reduction in absolute terms. CHENKUN's growing trajectory requires monitoring — it is the only country where overstock value increases over the forecast period.</div>
</div>

<div class="section">
<h2>15. Overstock by Brand – {start_label}</h2>
<div class="chart-wrapper"><canvas id="chart_brand"></canvas></div>
<div class="insight-box"><strong>Key Insight:</strong> Overstock concentration in 2–3 brands suggests the issue is a product portfolio problem rather than a broad inventory management failure.</div>
</div>

<script>
const months = {j(month_labels_js)};
const trajUnits = {j([round(v, 0) for v in traj_units])};
const trajValue = {j([round(v, 0) for v in traj_value])};
const trajPct   = {j([round(v, 2) for v in traj_pct])};

Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
Chart.defaults.font.size = 12;

function fmtVal(v)   {{ return v >= 1e6 ? '$'+(v/1e6).toFixed(2)+'M' : '$'+(v/1e3).toFixed(1)+'K'; }}
function fmtUnits(v) {{ return v >= 1e6 ? (v/1e6).toFixed(2)+'M' : (v/1e3).toFixed(1)+'K'; }}

const barLabelPlugin = {{
  id: 'barLabels',
  afterDatasetsDraw(chart) {{
    if (chart.config.type !== 'bar') return;
    const ctx = chart.ctx;
    const isHorizontal = chart.options.indexAxis === 'y';
    ctx.save();
    ctx.font = '9px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillStyle = '#333';
    chart.data.datasets.forEach((dataset, di) => {{
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      meta.data.forEach((bar, i) => {{
        const val = dataset.data[i];
        if (val === 0 || val === null || val === undefined) return;
        let label = '';
        const chartId = chart.canvas.id || '';
        if (chartId.includes('evac'))      label = val.toFixed(1) + '%';
        else if (chartId.includes('unit')) label = fmtUnits(val);
        else                               label = fmtVal(val);
        if (isHorizontal) {{
          const x = bar.x + 6;
          const y = bar.y;
          ctx.textAlign = 'left';
          ctx.textBaseline = 'middle';
          if (x + ctx.measureText(label).width > chart.chartArea.right + 60) {{
            ctx.fillStyle = '#fff';
            ctx.fillText(label, bar.x - ctx.measureText(label).width - 6, y);
            ctx.fillStyle = '#333';
          }} else {{
            ctx.fillText(label, x, y);
          }}
        }} else {{
          const x = bar.x;
          const y = bar.y - 6;
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          ctx.fillText(label, x, y);
        }}
      }});
    }});
    ctx.restore();
  }}
}};
Chart.register(barLabelPlugin);
Chart.defaults.layout = {{ padding: {{ top: 20, right: 70 }} }};

new Chart(document.getElementById('chart_trajectory'), {{
  type:'line', data:{{ labels:months, datasets:[
    {{ label:'Overstock Value ($)', data:trajValue, borderColor:'#2196F3', backgroundColor:'rgba(33,150,243,0.1)', fill:true, yAxisID:'y', tension:0.3, pointRadius:4 }},
    {{ label:'Overstock Units', data:trajUnits, borderColor:'#FF9800', backgroundColor:'rgba(255,152,0,0.1)', fill:true, yAxisID:'y1', tension:0.3, pointRadius:4 }}
  ] }},
  options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ y:{{ position:'left', title:{{display:true,text:'Value ($)'}}, ticks:{{callback:v=>fmtVal(v)}} }},
      y1:{{ position:'right', title:{{display:true,text:'Units'}}, ticks:{{callback:v=>fmtUnits(v)}}, grid:{{drawOnChartArea:false}} }} }} }}
}});

new Chart(document.getElementById('chart_pct'), {{
  type:'line', data:{{ labels:months, datasets:[
    {{ label:'Overstock % of Total Inventory', data:trajPct, borderColor:'#E53935', backgroundColor:'rgba(229,57,53,0.1)', fill:true, tension:0.3, pointRadius:5 }}
  ] }},
  options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{display:false}} }},
    scales:{{ y:{{ title:{{display:true,text:'% of Inventory'}}, ticks:{{callback:v=>v+'%'}} }} }} }}
}});

new Chart(document.getElementById('chart_country'), {{
  type:'bar', data:{{ labels:{j(COUNTRY_ORDER)}, datasets:[
    {{ label:'{start_label}', data:{j(country_mar26)}, backgroundColor:'rgba(76,175,80,0.6)', borderColor:'rgba(76,175,80,1)', borderWidth:1 }},
    {{ label:'{end_label} (Projected)', data:{j(country_mar27)}, backgroundColor:'rgba(244,67,54,0.6)', borderColor:'rgba(244,67,54,1)', borderWidth:1 }}
  ] }},
  options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ y:{{ title:{{display:true,text:'Value ($)'}}, ticks:{{callback:v=>fmtVal(v)}} }} }} }}
}});

const chColors = {j(CHANNEL_COLORS)};
const chLabels = {j(ch_labels)};
const chValData = {j(ch_values)};
new Chart(document.getElementById('chart_channel_val'), {{
  type:'bar', data:{{ labels:chLabels, datasets:[{{ label:'Overstock Value', data:chValData, backgroundColor:chColors.map(c=>c+'99'), borderColor:chColors, borderWidth:1 }}] }},
  options:{{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
    plugins:{{ title:{{display:true,text:'By Value ($)'}}, legend:{{display:false}},
      tooltip:{{callbacks:{{label:ctx=>' '+ctx.label+': '+fmtVal(ctx.raw)+' ('+((ctx.raw/chValData.reduce((a,b)=>a+b,0))*100).toFixed(1)+'%)'}}}} }},
    scales:{{ x:{{ ticks:{{callback:v=>fmtVal(v)}} }} }} }}
}});

const catColors = {j(CATEGORY_COLORS)};
new Chart(document.getElementById('chart_category'), {{
  type:'bar', data:{{ labels:{j(cat_labels)}, datasets:[{{ data:{j(cat_values)}, backgroundColor:catColors.map(c=>c+'99'), borderColor:catColors, borderWidth:1 }}] }},
  options:{{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
    plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:ctx=>' '+fmtVal(ctx.raw)}}}} }},
    scales:{{ x:{{ ticks:{{callback:v=>fmtVal(v)}} }} }} }}
}});

function makeTop15Chart(canvasId, data) {{
  const labels = data.map(s => s.is_silicone ? s.style+' (S)' : s.style);
  const values = data.map(s => s.value);
  const colors = data.map(s => s.is_silicone ? 'rgba(156,39,176,0.6)' : 'rgba(33,150,243,0.6)');
  const borders = data.map(s => s.is_silicone ? 'rgba(156,39,176,1)' : 'rgba(33,150,243,1)');
  new Chart(document.getElementById(canvasId), {{
    type:'bar', data:{{ labels, datasets:[{{ data:values, backgroundColor:colors, borderColor:borders, borderWidth:1 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:ctx=>' '+fmtVal(ctx.raw)+' | '+fmtUnits(data[ctx.dataIndex].units)+' units | Evac: '+data[ctx.dataIndex].evac_pct+'%'}}}} }},
      scales:{{ x:{{ ticks:{{callback:v=>fmtVal(v)}} }} }} }}
  }});
}}
makeTop15Chart('chart_top15_all',    {j(top15_all)});
makeTop15Chart('chart_top15_d2c',    {j(top15_d2c)});
makeTop15Chart('chart_top15_amazon', {j(top15_amazon)});

function makeEvacChartV2(canvasId, data, axisLabel, colorGood) {{
  const labels = data.map(s => s.style + (s.is_silicone ? ' (S)' : ''));
  const values = data.map(s => s.pct);
  const bg = data.map(s => s.is_silicone ? 'rgba(156,39,176,0.6)' : (colorGood ? 'rgba(76,175,80,0.6)' : 'rgba(244,67,54,0.6)'));
  const bd = data.map(s => s.is_silicone ? 'rgba(156,39,176,1)'   : (colorGood ? 'rgba(76,175,80,1)'   : 'rgba(244,67,54,1)'));
  new Chart(document.getElementById(canvasId), {{
    type:'bar', data:{{ labels, datasets:[{{ data:values, backgroundColor:bg, borderColor:bd, borderWidth:1 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:ctx=>{{
        const d = data[ctx.dataIndex];
        return ' '+ctx.raw+'% | Value: '+fmtVal(d.value_mar)+' | '+fmtUnits(d.units_mar)+' → '+fmtUnits(d.units_mar27)+' units';
      }}}}}} }},
      scales:{{ x:{{ title:{{display:true,text:axisLabel}}, ticks:{{callback:v=>v+'%'}}, min:0 }} }} }}
  }});
}}
makeEvacChartV2('chart_fast_evac', {j(fast_evac)}, '% of overstock cleared by {end_label}', true);
makeEvacChartV2('chart_slow_evac', {j(slow_evac)}, '% of overstock remaining by {end_label} (>100% = grew)', false);

new Chart(document.getElementById('chart_country_traj'), {{
  type:'line', data:{{ labels:months, datasets:{j(country_datasets)} }},
  options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ y:{{ title:{{display:true,text:'Value ($)'}}, ticks:{{callback:v=>fmtVal(v)}} }} }} }}
}});

new Chart(document.getElementById('chart_brand'), {{
  type:'bar', data:{{ labels:{j(brand_labels)}, datasets:[{{ data:{j(brand_values)}, backgroundColor:'rgba(33,150,243,0.6)', borderColor:'rgba(33,150,243,1)', borderWidth:1 }}] }},
  options:{{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
    plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:ctx=>' '+fmtVal(ctx.raw)}}}} }},
    scales:{{ x:{{ ticks:{{callback:v=>fmtVal(v)}} }} }} }}
}});
</script>
</body></html>"""

    kpis = {
        "value":   round(total_ovstk_value, 2),
        "units":   int(total_ovstk_units),
        "pct_eom": round(ovstk_pct_inv, 2),
        "start_label": start_label,
        "end_label":   end_label,
        "month_labels": MONTH_LABELS,
        "traj_units": [int(round(v, 0)) for v in traj_units],
        "traj_value": [round(v, 0) for v in traj_value],
    }
    return html, f"overstock_{start_slug}_{end_slug}_report.html", kpis


def main():
    filepath = find_input_file()
    print(f"Processing: {filepath.name}")
    html, out_filename, kpis = build_html(filepath)
    out_path = OUTPUT_DIR / out_filename
    out_path.write_text(html, encoding="utf-8")
    print(f"Saved: {out_path}")

    # Write KPI snapshot so the hub header stays in sync with this report
    kpi_path = OUTPUT_DIR / "overstock_kpis.json"
    kpi_path.write_text(json.dumps(kpis, indent=2))
    print(f"KPIs written: {kpi_path}")

    webbrowser.open(out_path.as_uri())


if __name__ == "__main__":
    main()
