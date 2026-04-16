/**
 * generate_slides.js
 * ------------------
 * Generates the 2-slide Weekly Performance Tracking PPTX using PptxGenJS.
 *
 * Slide 1: Per-product performance (Overstock + Aging sections)
 * Slide 2: Per-lever grouped view (5 lever groups + grand total)
 *
 * Usage:
 *   node tools/generate_slides.js
 *
 * Required files:
 *   tools/config/run_params.json     — report date, week labels, overstock totals, exec bullets
 *   tools/config/targets.json        — weekly unit targets per product
 *   tools/config/levers.json         — lever assignment per product
 *   tools/config/products.json       — subtitle, action, section, inventory_units per product
 *   .tmp/raw/<shipment file>         — actual weekly units shipped (WK1...WKn)
 *   .tmp/filtered/<latest>_filtered.csv — for aging auto-extraction
 */

'use strict';

const path = require('path');
const fs   = require('fs');
const XLSX = require('xlsx');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

let pptxgen;
try { pptxgen = require('pptxgenjs'); } catch(e) {
  console.error('pptxgenjs not found. Run: cd tools && npm install');
  process.exit(1);
}

const ROOT              = path.join(__dirname, '..');
const CONFIG_DIR        = path.join(__dirname, 'config');
const PROCESSED_FILE    = path.join(ROOT, '.tmp', 'processed', 'weekly_by_segment.csv');
const SEGMENT_AGING_FILE = path.join(ROOT, '.tmp', 'processed', 'segment_aging.json');
const FILTERED_DIR      = path.join(ROOT, '.tmp', 'filtered');
const SLIDES_DIR        = path.join(ROOT, '.tmp', 'slides');
fs.mkdirSync(SLIDES_DIR, { recursive: true });

// ── Load config files ────────────────────────────────────────────────────────
const runParams = JSON.parse(fs.readFileSync(path.join(CONFIG_DIR, 'run_params.json'), 'utf8'));
const targets   = JSON.parse(fs.readFileSync(path.join(CONFIG_DIR, 'targets.json'), 'utf8'));
const levers    = JSON.parse(fs.readFileSync(path.join(CONFIG_DIR, 'levers.json'), 'utf8'));
const products  = JSON.parse(fs.readFileSync(path.join(CONFIG_DIR, 'products.json'), 'utf8'));

const NUM_WEEKS = runParams.num_weeks;
const WEEK_LABELS = runParams.week_labels.slice(0, NUM_WEEKS);
const WEEK_DATES  = runParams.week_date_ranges.slice(0, NUM_WEEKS);
const REPORT_DATE = runParams.report_date;
const EXEC_BULLETS = runParams.exec_summary_bullets;

// ── Layout constants ─────────────────────────────────────────────────────────
const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const TITLE_H = 0.72;
const LEGEND_H = 0.22;
const HDR_H   = 0.40;
const ROW_H   = 0.34;
const SEC_H   = 0.30;
const TOT_H   = 0.30;
const EXEC_H  = 0.70;

// Column widths (Slide 1)
const COL = {
  LEVER:      0.44,
  PRODUCT:    1.10,
  TARGET:     0.44,
  WEEK:       0.56,
  ACCOMP:     0.40,
  YTD_TOT:    0.56,
  YTD_PCT:    0.40,
  AGED_UNITS: 0.50,
  AGED_PCT:   0.40,
  EVAC:       0.48,
};
const ACTIONS_W = SLIDE_W - COL.LEVER - COL.PRODUCT - COL.TARGET
  - (COL.WEEK * NUM_WEEKS) - COL.ACCOMP - COL.YTD_TOT - COL.YTD_PCT - COL.EVAC;

// Colors
const C = {
  TITLE_BG:   '1A1A2E',
  HDR_BG:     '16213E',
  YTD_HDR_BG: '0F172A',
  WHITE:      'FFFFFF',
  GRAY_ROW:   'F9FAFB',
  TOTAL_BG:   'E5E7EB',
  GRID:       'E0E0E0',
  MUTED:      '6B7280',
  OV_ACCENT:  '3B82F6',
  AG_ACCENT:  'F59E0B',
  GREEN_BG:   'D1FAE5', GREEN_TXT:  '065F46',
  AMBER_BG:   'FEF3C7', AMBER_TXT:  '92400E',
  RED_BG:     'FEE2E2', RED_TXT:    '991B1B',
  OV_TGT_BG:  'DBEAFE', OV_TGT_TXT: '2563EB',
  AG_TGT_BG:  'E0E7FF', AG_TGT_TXT: '4338CA',
};

const LEVER_COLORS = {
  'ACQ+CLR':       '10B981',
  'ACQ+CLR+BOOST': '3B82F6',
  'CLR ONLY':      'EF4444',
  'BOOST+CLR':     '8B5CF6',
  'BOOSTERS':      'F97316',
  'ACQ':           '06B6D4',
};

// ── Load weekly shipments ────────────────────────────────────────────────────
function loadShipments() {
  if (!fs.existsSync(PROCESSED_FILE)) {
    console.warn('weekly_by_segment.csv not found — run: python tools/process_shipments.py first.');
    console.warn('All weekly values will be 0.');
    return {};
  }
  console.log(`Loading shipments: weekly_by_segment.csv`);
  const wb = XLSX.readFile(PROCESSED_FILE);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: 0 });

  const map = {};
  rows.forEach(row => {
    const skuKey = Object.keys(row).find(k => /product|segment/i.test(k)) || Object.keys(row)[0];
    const name = String(row[skuKey]).trim();
    const weekData = WEEK_LABELS.map(wk => {
      const val = row[wk] || row[wk.toLowerCase()] || 0;
      return typeof val === 'string' ? parseInt(val.replace(/,/g, ''), 10) || 0 : val;
    });
    map[name] = weekData;
  });
  return map;
}

// ── Extract aging KPIs from filtered CSV ─────────────────────────────────────
function extractAgingKPIs() {
  const files = fs.existsSync(FILTERED_DIR)
    ? fs.readdirSync(FILTERED_DIR).filter(f => /^\d{6}_Aging.*_filtered\.csv$/i.test(f)).sort()
    : [];
  if (!files.length) return { valuation: 0, units: 0, pct: 0 };

  const fpath = path.join(FILTERED_DIR, files[files.length - 1]);
  const wb = XLSX.readFile(fpath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

  let valuation = 0, units = 0;
  const DENOM = runParams.total_inventory_denominator || 50935732;

  rows.forEach(row => {
    const range = String(row['Range TOTAL'] || '').trim();
    if (range !== 'Over 365') return;
    const amt = parseFloat(String(row['Total Amount $'] || '0').replace(/[$,]/g, '')) || 0;
    const qty = parseFloat(row['Qty']) || 0;
    valuation += amt;
    units += qty;
  });

  return { valuation, units: Math.round(units), pct: valuation / DENOM * 100 };
}

// ── Helpers ──────────────────────────────────────────────────────────────────
const n = v => Math.round(v).toLocaleString('en-US');
const pct = v => Math.round(v) + '%';
const achieveColor = (ratio) => {
  if (ratio >= 1)   return { bg: C.GREEN_BG, txt: C.GREEN_TXT };
  if (ratio >= 0.7) return { bg: C.AMBER_BG, txt: C.AMBER_TXT };
  return              { bg: C.RED_BG,   txt: C.RED_TXT };
};

function rect(slide, x, y, w, h, fillHex, opts = {}) {
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: fillHex },
    line: opts.line || { color: opts.border || fillHex, width: opts.lineW || 0.3 },
    ...opts,
  });
}

function badge(slide, text, x, y, w, h, bgHex, txtHex, fontSize = 5) {
  rect(slide, x, y, w, h, bgHex, { rectRadius: 0.03 });
  slide.addText(text, {
    x, y, w, h,
    fontSize, bold: true, color: txtHex, align: 'center', valign: 'middle',
    fontFace: 'Calibri',
  });
}

function txt(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: 'Calibri',
    fontSize: opts.fontSize || 6,
    color: opts.color || '1F2937',
    bold: opts.bold || false,
    align: opts.align || 'left',
    valign: opts.valign || 'middle',
    wrap: opts.wrap !== false,
    ...opts,
  });
}

// ── Build product rows ───────────────────────────────────────────────────────
function buildRows(shipments, totalAgedUnitsAll) {
  const overstockOrder = ['10022 Bra','10024 Straps Bra','10035 Sweetheart Bra','42075 HW Legging','62001 Scoop Neck Cami'];
  const agingOrder     = ['10400 Silicone Band','10210 Old Labs','10210 New Labs','10210 Extended','Bad Lots','EMP Brand'];

  // Load per-segment aged inventory from latest aging file (produced by process_shipments.py)
  const segAging = fs.existsSync(SEGMENT_AGING_FILE)
    ? JSON.parse(fs.readFileSync(SEGMENT_AGING_FILE, 'utf8'))
    : {};

  const buildRow = (name) => {
    const wkData = shipments[name] || Array(NUM_WEEKS).fill(0);
    const tgt    = targets[name] || null;
    const ytdTotal = wkData.reduce((a, b) => a + b, 0);
    const lastWk   = wkData[wkData.length - 1] || 0;
    const accomp   = tgt ? lastWk / tgt : null;
    const ytdPct   = tgt ? ytdTotal / (tgt * NUM_WEEKS) : null;
    const aging    = segAging[name] || { aged_units: null, total_units: null };
    const section  = (products[name] || {}).section || 'overstock';
    return {
      name,
      lever: levers[name] || '—',
      target: tgt,
      section,
      subtitle: (products[name] || {}).subtitle || '',
      action: (products[name] || {}).action || '',
      aged_units:       aging.aged_units,
      aged_total_units: aging.total_units,
      aged_pct:         null,   // computed below after totals are known
      weeksToEvac:      aging.total_units && ytdTotal > 0
                          ? aging.total_units / (ytdTotal / NUM_WEEKS)
                          : null,
      weekData: wkData,
      ytdTotal, lastWk, accomp, ytdPct,
    };
  };

  const overstockRows = overstockOrder.map(buildRow);
  const agingRows     = agingOrder.map(buildRow);

  // Aging rows only: aged_pct = segment aged units / total aged inventory (global denominator)
  agingRows.forEach(r => {
    if (r.aged_units != null && totalAgedUnitsAll > 0) {
      r.aged_pct = r.aged_units / totalAgedUnitsAll * 100;
      r.subtitle = `${r.aged_units.toLocaleString('en-US')} units | ${r.aged_pct.toFixed(1)}% of aging`;
    }
  });
  // Overstock rows: no aged data shown
  overstockRows.forEach(r => { r.aged_units = null; r.aged_pct = null; });

  return {
    overstock: overstockRows,
    aging:     agingRows,
  };
}

// ── Draw title bar ───────────────────────────────────────────────────────────
function drawTitleBar(slide, titleText) {
  rect(slide, 0, 0, SLIDE_W, TITLE_H, C.TITLE_BG);
  txt(slide, 'Weekly Performance Review', 0.16, 0.04, 8, 0.18, { fontSize: 7.5, color: '9CA3AF' });
  txt(slide, titleText, 0.16, 0.22, 9, 0.38, { fontSize: 22, bold: true, color: C.WHITE });
  txt(slide, REPORT_DATE, SLIDE_W - 1.6, 0.28, 1.44, 0.22, { fontSize: 12, color: '9CA3AF', align: 'right' });
}

// ── Draw legend row ──────────────────────────────────────────────────────────
function drawLegend(slide, y, grandTotal) {
  rect(slide, 0, y, SLIDE_W, LEGEND_H, 'F8FAFC', { border: C.GRID });
  txt(slide, '📈 12-week performance (units shipped)', 0.12, y, 3.5, LEGEND_H, { fontSize: 7, color: C.MUTED });
  // Color key
  const swatchY = y + 0.04;
  const swH = 0.10;
  [['D1FAE5','≥100%'], ['FEF3C7','70-99%'], ['FEE2E2','<70%']].forEach(([c, lbl], i) => {
    const sx = 3.8 + i * 0.9;
    rect(slide, sx, swatchY, 0.14, swH, c, { lineW: 0.5, border: 'CBD5E1' });
    txt(slide, lbl, sx + 0.17, swatchY, 0.65, swH, { fontSize: 7, color: C.MUTED });
  });
  txt(slide, `Grand Total: ${n(grandTotal)} units`, SLIDE_W - 2.2, y, 2.0, LEGEND_H, { fontSize: 7.5, color: '374151', align: 'right', bold: true });
}

// ── Draw column headers ──────────────────────────────────────────────────────
function drawHeaders(slide, y) {
  rect(slide, 0, y, SLIDE_W, HDR_H, C.HDR_BG);
  let x = 0;
  const hTxt = (t, w, opts = {}) => {
    txt(slide, t, x, y, w, HDR_H, { fontSize: 6, bold: true, color: C.WHITE, align: 'center', ...opts });
    x += w;
  };
  hTxt('LEVER', COL.LEVER);
  hTxt('PRODUCT/SEGMENT', COL.PRODUCT, { align: 'left' });
  hTxt('TARGET', COL.TARGET);
  WEEK_LABELS.forEach((wk, i) => {
    txt(slide, wk, x, y, COL.WEEK, HDR_H * 0.55,
      { fontSize: 6, bold: true, color: C.WHITE, align: 'center' });
    txt(slide, WEEK_DATES[i] || '', x, y + HDR_H * 0.55, COL.WEEK, HDR_H * 0.4,
      { fontSize: 4.5, color: '9CA3AF', align: 'center' });
    x += COL.WEEK;
  });
  const ytdBg = C.YTD_HDR_BG;
  rect(slide, x, y, COL.ACCOMP + COL.YTD_TOT + COL.YTD_PCT, HDR_H, ytdBg);
  hTxt('WK\nACCOMP', COL.ACCOMP, { fontSize: 5 });
  hTxt('YTD\nTOTAL', COL.YTD_TOT, { fontSize: 5 });
  hTxt('YTD %', COL.YTD_PCT, { fontSize: 5 });
  hTxt('WKS TO\nEVAC', COL.EVAC, { fontSize: 5 });
  hTxt('ACTIONS', ACTIONS_W, { align: 'left', fontSize: 5.5 });
}

// ── Draw section header ──────────────────────────────────────────────────────
function drawSectionHeader(slide, y, label, emoji, pillText, accentHex) {
  rect(slide, 0, y, SLIDE_W, SEC_H, 'F1F5F9');
  // Accent left bar
  rect(slide, 0, y, 0.08, SEC_H, accentHex, { lineW: 0 });
  txt(slide, `${emoji}  ${label}`, 0.14, y, 2.0, SEC_H, { fontSize: 8, bold: true, color: accentHex });
  // Pill badge
  rect(slide, 2.2, y + 0.03, 4.6, SEC_H - 0.06, 'E5E7EB', { rectRadius: 0.03 });
  txt(slide, pillText, 2.2, y + 0.03, 4.6, SEC_H - 0.06,
    { fontSize: 6.5, bold: true, color: '1F2937', align: 'center' });
}

// ── Draw data row ────────────────────────────────────────────────────────────
function drawDataRow(slide, row, y, isEven, section) {
  const bg = isEven ? C.GRAY_ROW : C.WHITE;
  rect(slide, 0, y, SLIDE_W, ROW_H, bg, { lineW: 0.2, border: C.GRID });

  let x = 0;
  // Lever badge
  const lc = LEVER_COLORS[row.lever] || '6B7280';
  badge(slide, row.lever, x + 0.03, y + 0.05, COL.LEVER - 0.06, ROW_H - 0.10, lc, C.WHITE, 5);
  x += COL.LEVER;

  // Product name + subtitle
  txt(slide, row.name, x + 0.04, y + 0.02, COL.PRODUCT - 0.06, ROW_H * 0.55,
    { fontSize: 6.5, bold: true, color: '111827' });
  txt(slide, row.subtitle, x + 0.04, y + ROW_H * 0.55, COL.PRODUCT - 0.06, ROW_H * 0.40,
    { fontSize: 5, color: C.MUTED });
  x += COL.PRODUCT;

  // Target badge
  const tBg  = section === 'overstock' ? C.OV_TGT_BG : C.AG_TGT_BG;
  const tTxt = section === 'overstock' ? C.OV_TGT_TXT : C.AG_TGT_TXT;
  const tLabel = row.target ? n(row.target) : 'N/A';
  badge(slide, tLabel, x + 0.02, y + 0.07, COL.TARGET - 0.04, ROW_H - 0.14, tBg, tTxt, 6.5);
  x += COL.TARGET;

  // Week cells
  row.weekData.forEach(v => {
    if (row.target) {
      const ac = achieveColor(v / row.target);
      rect(slide, x, y, COL.WEEK, ROW_H, ac.bg, { border: C.GRID, lineW: 0.2 });
      txt(slide, n(v), x, y, COL.WEEK, ROW_H, { fontSize: 6, bold: true, color: ac.txt, align: 'center' });
    } else {
      txt(slide, n(v), x, y, COL.WEEK, ROW_H, { fontSize: 6, color: '374151', align: 'center' });
    }
    x += COL.WEEK;
  });

  // Week accomplishment
  if (row.accomp !== null) {
    const ac = achieveColor(row.accomp);
    badge(slide, pct(row.accomp * 100), x + 0.02, y + 0.07, COL.ACCOMP - 0.04, ROW_H - 0.14, ac.bg, ac.txt, 6.5);
  } else {
    txt(slide, '—', x, y, COL.ACCOMP, ROW_H, { fontSize: 6, color: C.MUTED, align: 'center' });
  }
  x += COL.ACCOMP;

  // YTD total
  txt(slide, n(row.ytdTotal), x, y, COL.YTD_TOT, ROW_H, { fontSize: 6, bold: true, color: '111827', align: 'center' });
  x += COL.YTD_TOT;

  // YTD %
  if (row.ytdPct !== null) {
    const ac = achieveColor(row.ytdPct);
    badge(slide, pct(row.ytdPct * 100), x + 0.02, y + 0.07, COL.YTD_PCT - 0.04, ROW_H - 0.14, ac.bg, ac.txt, 6.5);
  } else {
    txt(slide, '—', x, y, COL.YTD_PCT, ROW_H, { fontSize: 6, color: C.MUTED, align: 'center' });
  }
  x += COL.YTD_PCT;

  // Weeks to evacuate
  if (row.weeksToEvac != null) {
    const wc = row.weeksToEvac <= 26 ? { bg: C.GREEN_BG, txt: C.GREEN_TXT }
             : row.weeksToEvac <= 52 ? { bg: C.AMBER_BG, txt: C.AMBER_TXT }
             : { bg: C.RED_BG, txt: C.RED_TXT };
    badge(slide, row.weeksToEvac.toFixed(1), x + 0.02, y + 0.07, COL.EVAC - 0.04, ROW_H - 0.14, wc.bg, wc.txt, 6);
  } else {
    txt(slide, '—', x, y, COL.EVAC, ROW_H, { fontSize: 6, color: C.MUTED, align: 'center' });
  }
  x += COL.EVAC;

  // Actions
  txt(slide, row.action, x + 0.04, y, ACTIONS_W - 0.04, ROW_H, { fontSize: 5, color: C.MUTED, wrap: true });
}

// ── Draw total row ────────────────────────────────────────────────────────────
function drawTotalRow(slide, rows, y, accentHex, subtitle) {
  rect(slide, 0, y, SLIDE_W, TOT_H, C.TOTAL_BG, { lineW: 0 });
  slide.addShape('line', { x: 0, y, w: SLIDE_W, h: 0, line: { color: accentHex, width: 2 } });

  const combined = rows.reduce((a, r) => a + (r.target || 0), 0);
  const weekTotals = WEEK_LABELS.map((_, i) => rows.reduce((a, r) => a + r.weekData[i], 0));
  const lastWkTotal = weekTotals[weekTotals.length - 1];
  const ytdTotal = weekTotals.reduce((a, b) => a + b, 0);
  const accomp = combined ? lastWkTotal / combined : null;
  const ytdPct = combined ? ytdTotal / (combined * NUM_WEEKS) : null;

  const totalAged    = rows.reduce((a, r) => a + (r.aged_units || 0), 0);
  const totalAgedPct = rows.reduce((a, r) => a + (r.aged_pct   || 0), 0);
  // Sum of row aged_pct = section's share of total aging (each row already uses global denominator)
  const agedLine = totalAged > 0
    ? `${n(totalAged)} aged units | ${totalAgedPct.toFixed(1)}% of aging`
    : '';

  let x = 0;
  txt(slide, '', x, y, COL.LEVER, TOT_H); x += COL.LEVER;
  txt(slide, subtitle || 'TOTAL', x + 0.04, y, COL.PRODUCT - 0.06, TOT_H * 0.55,
    { fontSize: 6, bold: true, color: '374151' });
  txt(slide, agedLine, x + 0.04, y + TOT_H * 0.52, COL.PRODUCT - 0.06, TOT_H * 0.44,
    { fontSize: 5, color: '1E3A8A' });
  x += COL.PRODUCT;
  x += COL.TARGET;

  weekTotals.forEach(v => {
    txt(slide, n(v), x, y, COL.WEEK, TOT_H, { fontSize: 6, bold: true, color: '111827', align: 'center' });
    x += COL.WEEK;
  });

  if (accomp !== null) {
    const ac = achieveColor(accomp);
    badge(slide, pct(accomp * 100), x + 0.02, y + 0.05, COL.ACCOMP - 0.04, TOT_H - 0.10, ac.bg, ac.txt, 6.5);
  }
  x += COL.ACCOMP;
  txt(slide, n(ytdTotal), x, y, COL.YTD_TOT, TOT_H, { fontSize: 6, bold: true, align: 'center' });
  x += COL.YTD_TOT;
  if (ytdPct !== null) {
    const ac = achieveColor(ytdPct);
    badge(slide, pct(ytdPct * 100), x + 0.02, y + 0.05, COL.YTD_PCT - 0.04, TOT_H - 0.10, ac.bg, ac.txt, 6.5);
  }
  x += COL.YTD_PCT;

  // Section total weeks to evacuate
  const totalInv = rows.reduce((a, r) => a + (r.aged_total_units || 0), 0);
  const totalYtd = rows.reduce((a, r) => a + r.ytdTotal, 0);
  const sectEvac = totalInv > 0 && totalYtd > 0 ? totalInv / (totalYtd / NUM_WEEKS) : null;
  if (sectEvac != null) {
    const wc = sectEvac <= 26 ? { bg: C.GREEN_BG, txt: C.GREEN_TXT }
             : sectEvac <= 52 ? { bg: C.AMBER_BG, txt: C.AMBER_TXT }
             : { bg: C.RED_BG, txt: C.RED_TXT };
    badge(slide, sectEvac.toFixed(1), x + 0.02, y + 0.07, COL.EVAC - 0.04, TOT_H - 0.14, wc.bg, wc.txt, 6);
  }
}

// ── Draw executive summary ───────────────────────────────────────────────────
function drawExecSummary(slide, y) {
  rect(slide, 0.16, y, SLIDE_W - 0.32, EXEC_H, 'F8FAFC', {
    border: 'CBD5E1', lineW: 0.8, rectRadius: 0.04
  });
  txt(slide, 'Executive Summary (Board Level)', 0.28, y + 0.04, 4, 0.18,
    { fontSize: 8, bold: true, color: '1E293B' });
  EXEC_BULLETS.forEach((b, i) => {
    txt(slide, `• ${b}`, 0.28, y + 0.22 + i * 0.18, SLIDE_W - 0.56, 0.18,
      { fontSize: 7, color: '374151', wrap: true });
  });
}

// ── SLIDE 1 ──────────────────────────────────────────────────────────────────
function buildSlide1(pres, rows, agingKPIs) {
  const slide = pres.addSlide();

  const ovTotal = rows.overstock.reduce((a, r) => a + r.ytdTotal, 0);
  const agTotal = rows.aging.reduce((a, r) => a + r.ytdTotal, 0);
  const grandTotal = ovTotal + agTotal;

  let y = 0;
  drawTitleBar(slide, `Performance Tracking — WK1 to WK${NUM_WEEKS} ${new Date().getFullYear()}`);
  y += TITLE_H;
  drawLegend(slide, y, grandTotal);
  y += LEGEND_H;
  drawHeaders(slide, y);
  y += HDR_H;

  // OVERSTOCK
  const ovVal = runParams.overstock_valuation || 0;
  const ovUnits = runParams.overstock_units || 0;
  const ovPct = ovVal / (runParams.total_inventory_denominator || 50935732) * 100;
  const ovPill = `$${(ovVal/1e6).toFixed(1)}M USD | ${ovPct.toFixed(1)}% of Total Inventory | ${n(ovUnits)} units`;
  drawSectionHeader(slide, y, 'OVERSTOCK', '🏷️', ovPill, C.OV_ACCENT);
  y += SEC_H;
  rows.overstock.forEach((row, i) => {
    drawDataRow(slide, row, y, i % 2 === 1, 'overstock');
    y += ROW_H;
  });
  drawTotalRow(slide, rows.overstock, y, C.OV_ACCENT, runParams.overstock_total_subtitle || '5 styles');
  y += TOT_H;

  // AGING
  const agVal = agingKPIs.valuation;
  const agUnits = agingKPIs.units;
  const agPct = agingKPIs.pct;
  const agPill = `$${(agVal/1e6).toFixed(2)}M USD | ${agPct.toFixed(1)}% of Total Inventory | ${n(agUnits)} units`;
  drawSectionHeader(slide, y, 'AGING', '⏳', agPill, C.AG_ACCENT);
  y += SEC_H;
  rows.aging.forEach((row, i) => {
    drawDataRow(slide, row, y, i % 2 === 1, 'aging');
    y += ROW_H;
  });
  drawTotalRow(slide, rows.aging, y, C.AG_ACCENT, runParams.aging_total_subtitle || '6 segments');
  y += TOT_H;

  drawExecSummary(slide, y);
}

// ── SLIDE 2 ──────────────────────────────────────────────────────────────────
function buildSlide2(pres, rows) {
  const slide = pres.addSlide();
  const LEVER_ORDER = ['ACQ+CLR', 'ACQ+CLR+BOOST', 'CLR ONLY', 'BOOST+CLR', 'ACQ'];

  // Col widths for Slide 2
  const COL2 = {
    LEVER: 0.48, STYLES: 1.60, INV: 0.65, TARGET: 0.48,
    WEEK: 0.56, ACCOMP: 0.40, YTD_TOT: 0.56, YTD_PCT: 0.40, EVAC: 0.50,
  };

  const allRows = [...rows.overstock, ...rows.aging];
  const grandTotal = allRows.reduce((a, r) => a + r.ytdTotal, 0);

  let y = 0;
  drawTitleBar(slide, `Performance by Lever — WK1 to WK${NUM_WEEKS} ${new Date().getFullYear()}`);
  y += TITLE_H;
  drawLegend(slide, y, grandTotal);
  y += LEGEND_H;

  // Headers for Slide 2
  rect(slide, 0, y, SLIDE_W, HDR_H, C.HDR_BG);
  let hx = 0;
  const hT2 = (t, w) => {
    txt(slide, t, hx, y, w, HDR_H, { fontSize: 6, bold: true, color: C.WHITE, align: 'center' });
    hx += w;
  };
  hT2('LEVER', COL2.LEVER);
  hT2('STYLES', COL2.STYLES);
  hT2('INVENTORY\n(units)', COL2.INV);
  hT2('TARGET /wk', COL2.TARGET);
  WEEK_LABELS.forEach(() => { hT2('', COL2.WEEK); hx -= COL2.WEEK; }); // rewrite with week labels
  WEEK_LABELS.forEach((wk, i) => {
    txt(slide, wk, hx, y, COL2.WEEK, HDR_H * 0.55, { fontSize: 6, bold: true, color: C.WHITE, align: 'center' });
    txt(slide, WEEK_DATES[i] || '', hx, y + HDR_H * 0.55, COL2.WEEK, HDR_H * 0.4, { fontSize: 4.5, color: '9CA3AF', align: 'center' });
    hx += COL2.WEEK;
  });
  rect(slide, hx, y, COL2.ACCOMP + COL2.YTD_TOT + COL2.YTD_PCT, HDR_H, C.YTD_HDR_BG);
  hT2('WK\nACCOMP', COL2.ACCOMP);
  hT2('YTD\nTOTAL', COL2.YTD_TOT);
  hT2('YTD %', COL2.YTD_PCT);
  rect(slide, hx, y, COL2.EVAC, HDR_H, '1E3A5F');
  hT2('WKS TO\nEVAC', COL2.EVAC);
  y += HDR_H;

  // Lever groups
  const ROW2_H = 0.48;
  LEVER_ORDER.forEach((lv, li) => {
    const group = allRows.filter(r => r.lever === lv);
    if (!group.length) return;

    const bg = li % 2 === 0 ? C.WHITE : C.GRAY_ROW;
    rect(slide, 0, y, SLIDE_W, ROW2_H, bg, { border: C.GRID, lineW: 0.2 });

    let rx = 0;
    const lc = LEVER_COLORS[lv] || '6B7280';
    badge(slide, lv, rx + 0.03, y + 0.08, COL2.LEVER - 0.06, ROW2_H - 0.16, lc, C.WHITE, 5.5);
    rx += COL2.LEVER;

    // Styles column
    txt(slide, `${group.length} styles`, rx + 0.04, y + 0.04, COL2.STYLES - 0.08, 0.18, { fontSize: 7, bold: true, color: '111827' });
    txt(slide, group.map(r => r.name).join(', '), rx + 0.04, y + 0.22, COL2.STYLES - 0.08, ROW2_H - 0.24,
      { fontSize: 5, color: C.MUTED, wrap: true });
    rx += COL2.STYLES;

    // Inventory
    const invTotal = group.reduce((a, r) => a + (r.inventory || 0), 0);
    txt(slide, invTotal ? n(invTotal) : '—', rx, y, COL2.INV, ROW2_H, { fontSize: 6, color: '374151', align: 'center' });
    rx += COL2.INV;

    // Combined target
    const combinedTgt = group.reduce((a, r) => a + (r.target || 0), 0);
    badge(slide, n(combinedTgt), rx + 0.02, y + 0.10, COL2.TARGET - 0.04, ROW2_H - 0.20, C.OV_TGT_BG, C.OV_TGT_TXT, 6.5);
    rx += COL2.TARGET;

    // Week sums
    const weekSums = WEEK_LABELS.map((_, wi) => group.reduce((a, r) => a + r.weekData[wi], 0));
    const lastWk = weekSums[weekSums.length - 1];
    const ytdSum = weekSums.reduce((a, b) => a + b, 0);
    const accomp = combinedTgt ? lastWk / combinedTgt : null;
    const ytdPct  = combinedTgt ? ytdSum / (combinedTgt * NUM_WEEKS) : null;

    weekSums.forEach(v => {
      if (combinedTgt) {
        const ac = achieveColor(v / combinedTgt);
        rect(slide, rx, y, COL2.WEEK, ROW2_H, ac.bg, { border: C.GRID, lineW: 0.2 });
        txt(slide, n(v), rx, y, COL2.WEEK, ROW2_H, { fontSize: 6, bold: true, color: ac.txt, align: 'center' });
      } else {
        txt(slide, n(v), rx, y, COL2.WEEK, ROW2_H, { fontSize: 6, color: '374151', align: 'center' });
      }
      rx += COL2.WEEK;
    });

    if (accomp !== null) {
      const ac = achieveColor(accomp);
      badge(slide, pct(accomp * 100), rx + 0.02, y + 0.10, COL2.ACCOMP - 0.04, ROW2_H - 0.20, ac.bg, ac.txt, 6.5);
    }
    rx += COL2.ACCOMP;
    txt(slide, n(ytdSum), rx, y, COL2.YTD_TOT, ROW2_H, { fontSize: 6, bold: true, align: 'center' });
    rx += COL2.YTD_TOT;
    if (ytdPct !== null) {
      const ac = achieveColor(ytdPct);
      badge(slide, pct(ytdPct * 100), rx + 0.02, y + 0.10, COL2.YTD_PCT - 0.04, ROW2_H - 0.20, ac.bg, ac.txt, 6.5);
    }
    rx += COL2.YTD_PCT;

    // Weeks to evacuate for this lever group
    const leverInv = group.reduce((a, r) => a + (r.aged_total_units || 0), 0);
    const leverEvac = leverInv > 0 && ytdSum > 0 ? leverInv / (ytdSum / NUM_WEEKS) : null;
    if (leverEvac != null) {
      const wc = leverEvac <= 26 ? { bg: C.GREEN_BG, txt: C.GREEN_TXT }
               : leverEvac <= 52 ? { bg: C.AMBER_BG, txt: C.AMBER_TXT }
               : { bg: C.RED_BG, txt: C.RED_TXT };
      badge(slide, leverEvac.toFixed(1), rx + 0.02, y + 0.10, COL2.EVAC - 0.04, ROW2_H - 0.20, wc.bg, wc.txt, 6.5);
    } else {
      txt(slide, '—', rx, y, COL2.EVAC, ROW2_H, { fontSize: 6, color: C.MUTED, align: 'center' });
    }

    y += ROW2_H;
  });

  // Grand total row
  const gtBg = C.TOTAL_BG;
  rect(slide, 0, y, SLIDE_W, TOT_H, gtBg, { lineW: 0 });
  slide.addShape('line', { x: 0, y, w: SLIDE_W, h: 0, line: { color: '374151', width: 2 } });

  const grandCombinedTgt = allRows.reduce((a, r) => a + (r.target || 0), 0);
  const grandWeekSums = WEEK_LABELS.map((_, wi) => allRows.reduce((a, r) => a + r.weekData[wi], 0));
  const grandLastWk = grandWeekSums[grandWeekSums.length - 1];
  const grandYtd = grandWeekSums.reduce((a, b) => a + b, 0);
  const grandAccomp = grandCombinedTgt ? grandLastWk / grandCombinedTgt : null;
  const grandYtdPct = grandCombinedTgt ? grandYtd / (grandCombinedTgt * NUM_WEEKS) : null;

  let gx = 0;
  txt(slide, '', gx, y, COL2.LEVER, TOT_H); gx += COL2.LEVER;
  txt(slide, 'GRAND TOTAL', gx + 0.04, y, COL2.STYLES + COL2.INV + COL2.TARGET, TOT_H,
    { fontSize: 7.5, bold: true, color: '1F2937' });
  gx += COL2.STYLES + COL2.INV + COL2.TARGET;
  grandWeekSums.forEach(v => {
    txt(slide, n(v), gx, y, COL2.WEEK, TOT_H, { fontSize: 6, bold: true, align: 'center' });
    gx += COL2.WEEK;
  });
  if (grandAccomp !== null) {
    const ac = achieveColor(grandAccomp);
    badge(slide, pct(grandAccomp * 100), gx + 0.02, y + 0.05, COL2.ACCOMP - 0.04, TOT_H - 0.10, ac.bg, ac.txt, 6.5);
  }
  gx += COL2.ACCOMP;
  txt(slide, n(grandYtd), gx, y, COL2.YTD_TOT, TOT_H, { fontSize: 6, bold: true, align: 'center' });
  gx += COL2.YTD_TOT;
  if (grandYtdPct !== null) {
    const ac = achieveColor(grandYtdPct);
    badge(slide, pct(grandYtdPct * 100), gx + 0.02, y + 0.05, COL2.YTD_PCT - 0.04, TOT_H - 0.10, ac.bg, ac.txt, 6.5);
  }
  gx += COL2.YTD_PCT;
  const grandInv = allRows.reduce((a, r) => a + (r.aged_total_units || 0), 0);
  const grandEvac = grandInv > 0 && grandYtd > 0 ? grandInv / (grandYtd / NUM_WEEKS) : null;
  if (grandEvac != null) {
    const wc = grandEvac <= 26 ? { bg: C.GREEN_BG, txt: C.GREEN_TXT }
             : grandEvac <= 52 ? { bg: C.AMBER_BG, txt: C.AMBER_TXT }
             : { bg: C.RED_BG, txt: C.RED_TXT };
    badge(slide, grandEvac.toFixed(1), gx + 0.02, y + 0.05, COL2.EVAC - 0.04, TOT_H - 0.10, wc.bg, wc.txt, 6.5);
  }
}

// ── HTML export ──────────────────────────────────────────────────────────────
function buildHTMLReport(rows, agingKPIs) {
  const fmtN  = v => v == null ? '—' : Math.round(v).toLocaleString('en-US');
  const fmtPct = v => v == null ? '—' : Math.round(v * 100) + '%';
  const fmtM  = v => '$' + (v / 1e6).toFixed(2) + 'M';
  const cellColor = ratio => {
    if (ratio == null) return '#f8fafc';
    if (ratio >= 1)    return '#d1fae5';
    if (ratio >= 0.7)  return '#fef3c7';
    return '#fee2e2';
  };
  const txtColor = ratio => {
    if (ratio == null) return '#64748b';
    if (ratio >= 1)    return '#065f46';
    if (ratio >= 0.7)  return '#92400e';
    return '#991b1b';
  };

  const agePctColor = p => {
    if (p == null) return '#f8fafc';
    if (p >= 50)   return '#fee2e2';
    if (p >= 25)   return '#fef3c7';
    return '#d1fae5';
  };
  const agePctTxt = p => {
    if (p == null) return '#64748b';
    if (p >= 50)   return '#991b1b';
    if (p >= 25)   return '#92400e';
    return '#065f46';
  };

  const buildTable = (sectionRows, title, accent) => {
    const wkHeaders = WEEK_LABELS.map((wk, i) =>
      `<th>${wk}<br><span style="font-size:9px;font-weight:400;color:#94a3b8;">${WEEK_DATES[i]||''}</span></th>`
    ).join('');

    const dataRows = sectionRows.map(r => {
      const wkCells = r.weekData.map((v, i) => {
        const isLast = i === r.weekData.length - 1;
        const bg  = isLast ? cellColor(r.accomp) : '#fff';
        const clr = isLast ? txtColor(r.accomp)  : '#1e293b';
        return `<td style="background:${bg};color:${clr};font-weight:${isLast?'700':'400'};">${fmtN(v)}</td>`;
      }).join('');
      const evacColor = r.weeksToEvac == null ? '#f8fafc'
                      : r.weeksToEvac <= 26    ? '#d1fae5'
                      : r.weeksToEvac <= 52    ? '#fef3c7'
                      : '#fee2e2';
      const evacTxt   = r.weeksToEvac == null ? '#64748b'
                      : r.weeksToEvac <= 26    ? '#065f46'
                      : r.weeksToEvac <= 52    ? '#92400e'
                      : '#991b1b';
      const evacCell  = `<td style="background:${evacColor};color:${evacTxt};font-weight:700;text-align:center;">${r.weeksToEvac != null ? r.weeksToEvac.toFixed(1) : '—'}</td>`;
      return `
        <tr>
          <td style="font-size:10px;color:#64748b;">${r.lever}</td>
          <td><strong>${r.name}</strong>${r.subtitle ? `<br><span style="font-size:10px;color:#64748b;">${r.subtitle}</span>` : ''}</td>
          <td style="text-align:center;">${fmtN(r.target)}</td>
          ${wkCells}
          <td style="background:${cellColor(r.accomp)};color:${txtColor(r.accomp)};font-weight:700;text-align:center;">${fmtPct(r.accomp)}</td>
          <td style="text-align:center;font-weight:600;">${fmtN(r.ytdTotal)}</td>
          <td style="background:${cellColor(r.ytdPct)};color:${txtColor(r.ytdPct)};font-weight:700;text-align:center;">${fmtPct(r.ytdPct)}</td>
          ${evacCell}
          <td style="font-size:11px;color:#475569;">${r.action||'—'}</td>
        </tr>`;
    }).join('');

    // Subtotal row
    const totalTarget  = sectionRows.reduce((a, r) => a + (r.target || 0), 0);
    const weekTotals   = WEEK_LABELS.map((_, i) => sectionRows.reduce((a, r) => a + r.weekData[i], 0));
    const lastWkTotal  = weekTotals[weekTotals.length - 1];
    const ytdSectionTotal = weekTotals.reduce((a, b) => a + b, 0);
    const sectionAccomp  = totalTarget ? lastWkTotal / totalTarget : null;
    const sectionYtdPct  = totalTarget ? ytdSectionTotal / (totalTarget * NUM_WEEKS) : null;
    const totalAgedUnits = sectionRows.reduce((a, r) => a + (r.aged_units || 0), 0);
    const totalAgedPct   = sectionRows.reduce((a, r) => a + (r.aged_pct || 0), 0);
    const totalAgedPctOfAll = totalAgedUnits > 0 && agingKPIs.units > 0
      ? (totalAgedUnits / agingKPIs.units * 100).toFixed(1) + '%'
      : null;
    const agedSummary = totalAgedPctOfAll
      ? `<span style="font-size:10px;font-weight:400;color:#1e3a8a;margin-left:10px;">${fmtN(totalAgedUnits)} aged units | ${totalAgedPctOfAll} of aging</span>`
      : '';
    const sectInv  = sectionRows.reduce((a, r) => a + (r.aged_total_units || 0), 0);
    const sectYtd  = sectionRows.reduce((a, r) => a + r.ytdTotal, 0);
    const sectEvac = sectInv > 0 && sectYtd > 0 ? sectInv / (sectYtd / NUM_WEEKS) : null;
    const sectEvacColor = sectEvac == null ? '#f8fafc' : sectEvac <= 26 ? '#d1fae5' : sectEvac <= 52 ? '#fef3c7' : '#fee2e2';
    const sectEvacTxt   = sectEvac == null ? '#64748b' : sectEvac <= 26 ? '#065f46' : sectEvac <= 52 ? '#92400e' : '#991b1b';
    const wkTotalCells = weekTotals.map((v, i) => {
      const isLast = i === weekTotals.length - 1;
      const bg = isLast ? cellColor(sectionAccomp) : '#e2e8f0';
      const cl = isLast ? txtColor(sectionAccomp)  : '#0f172a';
      return `<td style="background:${bg};color:${cl};font-weight:700;text-align:center;">${fmtN(v)}</td>`;
    }).join('');
    const subtotalRow = `
      <tr style="background:#f1f5f9;border-top:2px solid ${accent};">
        <td colspan="2" style="padding:7px 10px;font-weight:700;color:#0f172a;">SECTION TOTAL ${agedSummary}</td>
        <td style="text-align:center;font-weight:700;">${fmtN(totalTarget)}/wk</td>
        ${wkTotalCells}
        <td style="background:${cellColor(sectionAccomp)};color:${txtColor(sectionAccomp)};font-weight:700;text-align:center;">${fmtPct(sectionAccomp)}</td>
        <td style="text-align:center;font-weight:700;">${fmtN(ytdSectionTotal)}</td>
        <td style="background:${cellColor(sectionYtdPct)};color:${txtColor(sectionYtdPct)};font-weight:700;text-align:center;">${fmtPct(sectionYtdPct)}</td>
        <td style="background:${sectEvacColor};color:${sectEvacTxt};font-weight:700;text-align:center;">${sectEvac != null ? sectEvac.toFixed(1) : '—'}</td>
        <td></td>
      </tr>`;

    return `
      <div style="margin-bottom:28px;">
        <div style="background:${accent};color:#fff;padding:8px 16px;border-radius:6px 6px 0 0;font-weight:700;font-size:13px;">${title}</div>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead style="background:#1e293b;color:#fff;">
              <tr>
                <th style="padding:8px 10px;text-align:left;">Lever</th>
                <th style="padding:8px 10px;text-align:left;">Product / Segment</th>
                <th style="padding:8px 10px;">Target</th>
                ${wkHeaders}
                <th style="padding:8px 10px;">WK Accomp</th>
                <th style="padding:8px 10px;">YTD Total</th>
                <th style="padding:8px 10px;">YTD %</th>
                <th style="padding:8px 10px;background:#1e3a5f;">Wks to Evac</th>
                <th style="padding:8px 10px;text-align:left;">Actions</th>
              </tr>
            </thead>
            <tbody>${dataRows}${subtotalRow}</tbody>
          </table>
        </div>
      </div>`;
  };

  const allRows = [...rows.overstock, ...rows.aging];
  const grandTotal = allRows.reduce((s, r) => s + r.ytdTotal, 0);
  const bullets = (runParams.exec_summary_bullets || [])
    .map(b => `<li style="margin-bottom:6px;">${b}</li>`).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Weekly Performance — ${REPORT_DATE}</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; background:#f8fafc; color:#1e293b; padding:32px; }
  .page { max-width:1400px; margin:0 auto; }
  .header { background:linear-gradient(135deg,#1a237e,#283593); color:#fff; border-radius:10px; padding:24px 28px; margin-bottom:24px; display:flex; justify-content:space-between; align-items:flex-start; }
  .header h1 { font-size:22px; font-weight:700; }
  .header p  { font-size:12px; color:rgba(255,255,255,0.7); margin-top:4px; }
  .kpi-row { display:flex; gap:12px; margin-bottom:24px; flex-wrap:wrap; }
  .kpi-card { flex:1; min-width:160px; background:#fff; border-radius:8px; padding:14px 18px; border-left:3px solid #1a237e; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
  .kpi-card .v { font-size:22px; font-weight:700; color:#1a237e; }
  .kpi-card .l { font-size:11px; color:#64748b; margin-top:2px; }
  .summary-card { background:#fff; border-radius:10px; padding:18px 20px; margin-bottom:24px; border:1px solid #e2e8f0; }
  .summary-card h3 { font-size:13px; font-weight:700; color:#1a237e; margin-bottom:10px; }
  .summary-card ul { padding-left:18px; color:#475569; font-size:13px; line-height:1.7; }
  table th { padding:8px 10px; white-space:nowrap; }
  table td { padding:7px 10px; border-top:1px solid #f1f5f9; }
  .legend { display:flex; gap:16px; font-size:11px; color:#64748b; margin-bottom:16px; align-items:center; }
  .swatch { width:12px; height:12px; border-radius:2px; display:inline-block; margin-right:4px; vertical-align:middle; }
  .footer { margin-top:24px; font-size:11px; color:#94a3b8; text-align:right; }
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <div>
      <h1>Weekly Performance Review</h1>
      <p>${REPORT_DATE} &nbsp;·&nbsp; Week ${runParams.week_current || NUM_WEEKS} of ${NUM_WEEKS}</p>
    </div>
    <div style="text-align:right;">
      <div style="font-size:11px;color:rgba(255,255,255,0.6);">Grand Total YTD</div>
      <div style="font-size:24px;font-weight:700;">${fmtN(grandTotal)} units</div>
    </div>
  </div>

  <div class="kpi-row">
    <div class="kpi-card"><div class="v">${fmtM(agingKPIs.valuation)}</div><div class="l">Aged Value (Over 365)</div></div>
    <div class="kpi-card"><div class="v">${fmtN(agingKPIs.units)}</div><div class="l">Aged Units</div></div>
    <div class="kpi-card"><div class="v">${agingKPIs.pct.toFixed(1)}%</div><div class="l">% of Total Inventory</div></div>
    <div class="kpi-card"><div class="v">${fmtN(grandTotal)}</div><div class="l">Total Units Shipped YTD</div></div>
  </div>

  ${bullets ? `<div class="summary-card"><h3>Executive Summary</h3><ul>${bullets}</ul></div>` : ''}

  <div class="legend">
    <span><span class="swatch" style="background:#d1fae5;border:1px solid #6ee7b7;"></span>≥ 100% (On Target)</span>
    <span><span class="swatch" style="background:#fef3c7;border:1px solid #fcd34d;"></span>70–99% (Near Target)</span>
    <span><span class="swatch" style="background:#fee2e2;border:1px solid #fca5a5;"></span>&lt; 70% (Below Target)</span>
  </div>

  ${buildTable(rows.overstock, '📦 Overstock', '#1e40af')}
  ${buildTable(rows.aging,     '📉 Aging',     '#6d28d9')}

  <div class="footer">Generated by Overstock &amp; Aging Automation · ${new Date().toLocaleString()}</div>
</div>
</body>
</html>`;
}

// ── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  console.log('Loading shipment data...');
  const shipments = loadShipments();

  console.log('Extracting aging KPIs...');
  const agingKPIs = extractAgingKPIs();
  console.log(`  Aging: $${(agingKPIs.valuation/1e6).toFixed(2)}M | ${n(agingKPIs.units)} units | ${agingKPIs.pct.toFixed(1)}%`);

  console.log('Building product rows...');
  const rows = buildRows(shipments, agingKPIs.units);

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.author  = 'Overstock & Aging Automation';

  console.log('Building Slide 1 (per-product)...');
  buildSlide1(pres, rows, agingKPIs);

  console.log('Building Slide 2 (per-lever)...');
  buildSlide2(pres, rows);

  const wk = runParams.week_current || NUM_WEEKS;
  const yr = new Date().getFullYear();
  const baseName = `weekly_performance_WK${wk}_${yr}`;
  const outPath  = path.join(SLIDES_DIR, baseName + '.pptx');

  await pres.writeFile({ fileName: outPath });
  console.log(`\nPPTX saved: .tmp/slides/${baseName}.pptx`);

  console.log('Building HTML export...');
  const htmlContent = buildHTMLReport(rows, agingKPIs);
  const htmlPath = path.join(SLIDES_DIR, baseName + '.html');
  fs.writeFileSync(htmlPath, htmlContent, 'utf8');
  console.log(`HTML saved: .tmp/slides/${baseName}.html`);
}

main().catch(e => { console.error(e); process.exit(1); });
