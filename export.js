"use strict";

/*
 * Builds the Excel workbook – one sheet per exported report page – in the
 * "Boardroom Brief" design: each sheet reads like a page of a consulting board pack.
 *
 *  1. Masthead – a thin rule in the customer accent, logo, "COMPANY · REPORT" eyebrow,
 *     the page title, a meta line (export time, filters, page n of m) closed by an ink
 *     rule, and a pale filter bar of "field  value" chips.
 *  2. Visuals, laid out either
 *     - "report": like the Power BI page. Visuals whose vertical extents overlap form a
 *       "band" and start on the same row; visuals sharing horizontal space in a band
 *       form a "stack". Stacks go left to right one gutter column apart, and stacks
 *       that start at the same x on the page snap to the same Excel column. Or
 *     - "stacked": one visual under another in reading order.
 *     Cards are scorecard tiles (a band of only cards packs into one key-figure strip).
 *     Every other visual is an "exhibit": a title, a caption row (legend of the chart's
 *     colours, or "Table · 24 rows"), and a real Excel Table with filter / sort
 *     buttons, styled only with ink rules and hairlines (or a built-in table style).
 *  3. Chart tables show the chart's exact colours: a colour key in the caption, the
 *     series colour on each series header, value cells filled with the colour of their
 *     bar / line / slice, and a colour tab in the gutter left of each row for charts
 *     that colour rows by category (kept right after a sort by conditional formatting).
 *     Colours come from the report file when it can be read (colors.js: exact theme and
 *     colours set by hand), else from the theme the embed API reports. The accent
 *     appears only in the top rule, KPI tile caps and the sheet tab.
 *  4. Long tables collapse rows past `visibleRows` behind [+] / [–]. Excel hides whole
 *     rows, so only rows holding nothing but those extra rows are collapsed, and
 *     side-by-side tables are aligned so none hides another's content.
 *  5. Notes, footer, print setup (landscape, fit to width, header / footer).
 * The file is white-label: no product name is written anywhere.
 * Uses globals from app.js: embeddedReport, currentReport, currentReportName, models, log,
 * errorText, filterParts, NON_DATA_TYPES; and from colors.js: loadReportColors,
 * reportColorsMessage, normalizeField, normalizeValue.
 */

const KPI_TYPES = ["card", "multiRowCard", "kpi", "cardVisual", "gauge"];
const PX_TOLERANCE = 10;       // overlap smaller than this does not count as overlapping
const PAGE_RENDER_TIMEOUT_MS = 20000;

/* ── Design tokens ────────────────────────────────────────────────────── */
const C = {
  INK: "FF0E1A2B",       // titles, KPI values, header text, rules
  TEXT: "FF1E2733",      // table body, filter values
  SLATE: "FF5B6B7F",     // labels, meta, captions, notes
  MUTED: "FF6A7787",     // secondary text (white background only)
  SURFACE: "FFF3F5F8",   // the only fills: table header, KPI tiles, filter bar
  HAIR: "FFE3E8EE",      // row separators
  RULE: "FFC9D1DC",      // dividers, notes / footer rules
  WHITE: "FFFFFFFF",
};
const DEFAULT_ACCENT = "FF0B7F56";
const ROW_H = 18;          // the row grid, in points
const HEAD_ROWS = 3;       // exhibit title, caption, table header
const CONTENT_COL = 2;     // everything starts at column B; column A is the margin
const MARGIN_W = 3;
const GUTTER_W = 3;
const MIN_PAGE_W = 100;    // minimum content width (Excel width units)
const SNAP_PX = 40;        // stacks starting within this many page pixels share a column
const BAND_GAP = 2;        // blank rows between bands
const MAX_CF_RULES = 24;   // above this many row colours, dots stay static (not sort-safe)
const WIDTH_STEPS = [10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 28, 32, 36, 40, 45, 50, 60];
const EXECUTIVE = "Executive";
const BUILT_IN_STYLES = ["TableStyleMedium2", "TableStyleMedium4", "TableStyleMedium7", "TableStyleMedium1", "TableStyleLight9"];

const SURFACE_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: C.SURFACE } };
const HEADER_BORDER = { top: { style: "medium", color: { argb: C.INK } } };
const BODY_BORDER = { top: { style: "thin", color: { argb: C.HAIR } } };
const INK_RULE = { style: "thin", color: { argb: C.INK } };
const LIGHT_RULE = { style: "thin", color: { argb: C.RULE } };
// Coloured cells are parted by white hairlines, so filled rows read like separate bars
const GAP_BORDER = { top: { style: "thin", color: { argb: C.WHITE } }, bottom: { style: "thin", color: { argb: C.WHITE } } };
const REPORT_COLORS_TIMEOUT_MS = 45000;   // longest wait for the report file after the pages are read
const KEY_GLYPH = "■";                     // colour key in captions: a swatch, like the filled cells

// Power BI's default theme data colours, used when the report's own theme can't be read
const DEFAULT_THEME_COLORS = ["#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2", "#D9B300", "#D64550"];
// Table-like visuals don't draw with the theme's data colours
const TABLE_VISUAL_TYPES = ["tableEx", "pivotTable", "table", "matrix"];
const HEX_COLOR = /^#[0-9A-F]{6}$/i;    // the leading # keeps codes like 202406 from counting as colours
// Number columns that are helpers, not series the chart draws (sort keys, ranks, IDs)
const HELPER_COLUMN = /\b(sort|rank|order|index|id|key)\b/i;
// Visuals that give every value of their first field its own colour (each slice / tile)
const COLOR_BY_CATEGORY_TYPES = ["treemap", "pieChart", "donutChart"];

/* "#1a73e8" → "FF1A73E8" */
function toArgb(hex, fallback = DEFAULT_ACCENT) {
  const h = String(hex || "").replace("#", "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(h) ? "FF" + h : fallback;
}

function solidFill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

/* WCAG relative luminance and contrast ratio of two ARGB colours */
function luminance(argb) {
  const n = parseInt(argb.slice(2), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
}

function contrast(a, b) {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

/* Where the customer accent goes. A very light accent (white, pale yellow) would vanish
   on white, so the rule and KPI caps fall back to ink; the sheet tab keeps the accent. */
function accentPlan(argb) {
  const tooLight = contrast(argb, C.WHITE) < 1.5;
  return { rule: tooLight ? C.INK : argb, cap: tooLight ? C.INK : argb, tab: argb };
}

/* Every text role's font. Calibri / Calibri Light ship with every Office since 2013,
   Mac Office and Excel for the web; Georgia is the optional editorial title font. */
function makeFonts(titleFont) {
  const georgia = titleFont === "Georgia";
  const f = (size, color, extra = {}) => ({ name: "Calibri", family: 2, size, color: { argb: color }, ...extra });
  return {
    eyebrow: f(9, C.SLATE, { bold: true }),
    title: georgia ? { name: "Georgia", family: 1, size: 22, color: { argb: C.INK } }
      : { name: "Calibri Light", family: 2, size: 24, color: { argb: C.INK } },
    meta: f(10, C.SLATE),
    filterLabel: f(9, C.SLATE, { bold: true }),
    chipField: f(10, C.SLATE),
    chipValue: f(10, C.TEXT, { bold: true }),
    chipSep: f(10, C.RULE),
    exNum: f(10, C.MUTED, { bold: true }),
    exTitle: georgia ? { name: "Georgia", family: 1, size: 13, color: { argb: C.INK } } : f(12, C.INK, { bold: true }),
    caption: f(9, C.SLATE),
    captionMore: f(9, C.MUTED),
    dot: (argb) => f(10, argb),
    header: f(10, C.INK, { bold: true }),
    headerOnStyle: f(10, C.WHITE, { bold: true }),
    body: f(10, C.TEXT),
    kpiLabel: f(9, C.SLATE, { bold: true }),
    kpiValue: { name: "Calibri Light", family: 2, size: 22, color: { argb: C.INK } },
    kpiSubLabel: f(9, C.SLATE),
    kpiSubValue: f(10, C.TEXT, { bold: true }),
    kpiExact: f(9, C.SLATE),
    noteCount: f(9, C.TEXT, { bold: true }),
    noteText: f(9, C.MUTED),
    notesHead: f(9, C.SLATE, { bold: true }),
    noteNum: f(9, C.MUTED),
    noteBody: f(9, C.SLATE),
    footerLeft: f(9, C.SLATE),
    footerRight: f(9, C.MUTED),
    empty: f(10, C.SLATE),
  };
}

/* Rich text from [text, font] pairs; empty runs are dropped */
function rich(...runs) {
  return { richText: runs.filter((r) => r && r[0]).map(([text, font]) => ({ text, font })) };
}

/* Sets fill and / or border on a run of cells, keeping borders already there */
function paintRange(ws, row, c1, c2, { fill, border } = {}) {
  for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(row, c);
    if (fill) cell.fill = fill;
    if (border) cell.border = { ...(cell.border || {}), ...border };
  }
}

/* ── Visual names ─────────────────────────────────────────────────────── */
const FRIENDLY_TYPES = {
  tableEx: "Table", pivotTable: "Matrix", card: "Card", cardVisual: "Card", multiRowCard: "Multi-row card",
  kpi: "KPI", gauge: "Gauge", lineChart: "Line chart", areaChart: "Area chart", stackedAreaChart: "Area chart",
  barChart: "Bar chart", clusteredBarChart: "Bar chart", hundredPercentStackedBarChart: "Bar chart",
  columnChart: "Column chart", clusteredColumnChart: "Column chart", hundredPercentStackedColumnChart: "Column chart",
  lineClusteredColumnComboChart: "Combo chart", lineStackedColumnComboChart: "Combo chart",
  ribbonChart: "Ribbon chart", waterfallChart: "Waterfall chart", funnel: "Funnel", scatterChart: "Scatter chart",
  pieChart: "Pie chart", donutChart: "Donut chart", treemap: "Treemap", map: "Map", filledMap: "Map",
  shapeMap: "Map", azureMap: "Map", decompositionTreeVisual: "Decomposition tree",
};

function friendlyType(type) {
  const t = String(type || "");
  if (FRIENDLY_TYPES[t]) return FRIENDLY_TYPES[t];
  if (/^deneb/i.test(t)) return "Deneb visual";
  if (/[0-9A-F]{32}$/i.test(t)) return "Custom visual";
  const words = t.replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
  return words ? words[0].toUpperCase() + words.slice(1) : "Visual";
}

const GENERIC_TITLES = new Set(["card", "multi-row card", "multirow card", "kpi", "gauge", "table", "matrix",
  "deneb", "visual", "chart", "untitled", "custom visual"]);

/* A title Power BI or this app filled in that says nothing about the data */
function isGenericTitle(title, type) {
  const t = String(title || "").trim();
  if (!t) return true;
  if (type && t.startsWith(`${type} (`) && t.endsWith(")")) return true;   // collectPage's "type (name)" fallback
  const low = t.toLowerCase();
  if (GENERIC_TITLES.has(low) || low === friendlyType(type).toLowerCase()) return true;
  const ty = String(type || "");
  return /[0-9A-F]{32}$/i.test(ty) && ty.toLowerCase().startsWith(low.replace(/\s+/g, ""));
}

/* Per-column facts about a visual's exported rows */
function classifyColumns(rows) {
  const [header, ...data] = rows;
  const filled = (ci) => data.map((r) => (r[ci] ?? "").trim()).filter(Boolean);
  const numeric = header.map((_, ci) => { const v = filled(ci); return v.length > 0 && v.every((x) => typeof toCellValue(x).value === "number"); });
  const hex = header.map((_, ci) => { const v = filled(ci); return v.length > 0 && v.every((x) => HEX_COLOR.test(x)); });
  const helper = header.map((h) => HELPER_COLUMN.test(h || ""));
  return { header, data, filled, numeric, hex, helper };
}

/* A title built from the columns, for visuals whose own title is generic */
function describeTitle(item) {
  const { header, numeric, helper, hex } = classifyColumns(item.rows);
  if (TABLE_VISUAL_TYPES.includes(item.type)) {
    const shown = header.slice(0, 3).join(", ");
    return header.length > 3 ? `${shown} +${header.length - 3}` : shown;
  }
  const series = header.filter((_, ci) => ci > 0 && numeric[ci] && !helper[ci] && !hex[ci]);
  const texts = header.filter((_, ci) => !numeric[ci] && !hex[ci]);
  if (!series.length) return `${friendlyType(item.type)} data`;
  const s = series.length <= 2 ? series.join(" and ") : `${series.slice(0, 2).join(", ")} and ${series.length - 2} more`;
  return texts.length ? `${s} by ${texts.slice(0, 2).join(" and ")}` : s;
}

/* ── Colours ──────────────────────────────────────────────────────────── */

/* The report theme's data colours (series 1, 2, 3 …), or Power BI's defaults */
async function readThemeColors() {
  try {
    const theme = await embeddedReport.getTheme();
    const json = theme && (theme.themeJson || theme);
    const colors = ((json && json.dataColors) || []).filter((c) => HEX_COLOR.test(c));
    if (colors.length) return { colors, fromReport: true };
  } catch (e) {
    log(`Report theme could not be read: ${errorText(e)}`);
  }
  return { colors: DEFAULT_THEME_COLORS, fromReport: false };
}

/*
 * Legend values in Power BI's series order. Rows come grouped by the axis, and within a
 * group in series order. A series missing from the first groups still lands in its place:
 * each group's sequence is merged in, a new value going after the previous value of its
 * group already listed (or before the next one).
 */
function seriesOrder(data, keyCol, groupCols) {
  const order = [];
  let group = null, seq = [];
  const flush = () => {
    seq.forEach((v, i) => {
      if (order.includes(v)) return;
      const prev = seq.slice(0, i).reverse().find((x) => order.includes(x));
      const next = seq.slice(i + 1).find((x) => order.includes(x));
      const at = prev !== undefined ? order.indexOf(prev) + 1 : next !== undefined ? order.indexOf(next) : order.length;
      order.splice(at, 0, v);
    });
    seq = [];
  };
  for (const row of data) {
    const g = groupCols.map((c) => row[c]).join("\u0001");
    if (g !== group) { flush(); group = g; }
    const v = (row[keyCol] ?? "").trim();
    if (v && !seq.includes(v)) seq.push(v);
  }
  flush();
  return order;
}

/*
 * Evaluates an fx colour rule (colors.js parseColorRule) for every row of a chart table.
 * The rule's input is a column of the table; else a column whose name holds all its words
 * ("Profit %" → "Profit Margin %"); else, for a rule per legend value, the legend order
 * (Power BI's scale over e.g. "Latest year" follows the years). Blank inputs count as zero,
 * Power BI's default. The last two are approximations and add a note.
 * Returns colours for columnColors: perRow (one colour per row), the filled columns, the
 * tab key, and a caption legend (the scale, or one entry per rule).
 */
function colorRuleRows({ scope, rule }, t) {
  const { header, data, cols, norm, lower, numeric, hex, seriesCols, findCol, info, seriesColor } = t;
  const targets = scope.startsWith("m:") ? seriesCols.filter((ci) => "m:" + norm[ci] === scope) : seriesCols;
  if (!targets.length || !data.length) return null;
  let approx = false;
  const resolve = (input) => {
    const ci = findCol({ ref: input.toLowerCase() });
    if (ci >= 0 && numeric[ci]) return ci;
    const words = input.toLowerCase().split(/\s+/).filter(Boolean);
    const close = cols.filter((c) => numeric[c] && words.every((w) => lower[c].split(/\s+/).includes(w)));
    if (close.length === 1) { approx = true; return close[0]; }
    return -1;
  };
  const num = (ri, ci) => { const v = toCellValue(data[ri][ci]).value; return typeof v === "number" ? v : 0; };
  const legendCol = info.series ? findCol(info.series) : -1;
  const categoryCol = findCol(info.category);
  const keyCol = scope === "series" && legendCol >= 0 ? legendCol
    : categoryCol >= 0 ? categoryCol : cols.find((ci) => !numeric[ci] && !hex[ci]) ?? -1;
  // A category hierarchy (Group, then Region) exports every level, but the chart shows its
  // first level: one bar per Group, coloured by that Group's total. Rows are coloured by
  // their group's total, so the table matches the bars (not for percentages, which don't add up).
  const levels = (info.categories || []).map((f) => findCol(f)).filter((ci) => ci >= 0);
  const grouped = scope !== "series" && levels.length > 1 && levels[0] === categoryCol;

  let perRow, legend, inputName = rule.input;
  if (rule.kind === "gradient") {
    const inputCol = resolve(rule.input);
    let values;
    if (inputCol >= 0) {
      values = data.map((_, ri) => num(ri, inputCol));
      inputName = header[inputCol];
      const percent = data.some((r) => String(r[inputCol] ?? "").trim().endsWith("%"));
      if (grouped && !percent) {
        const totals = {};
        const groupOf = (r) => (r[categoryCol] ?? "").trim();
        data.forEach((r, ri) => { totals[groupOf(r)] = (totals[groupOf(r)] || 0) + values[ri]; });
        values = data.map((r) => totals[groupOf(r)]);
      }
    } else if (scope === "series" && legendCol >= 0) {
      const order = seriesOrder(data, legendCol, cols.filter((ci) => ci !== legendCol && !seriesCols.includes(ci) && !hex[ci]));
      values = data.map((r) => order.indexOf((r[legendCol] ?? "").trim()));
      approx = true;
    } else {
      return null;
    }
    const lo = rule.min.value ?? Math.min(...values), hi = rule.max.value ?? Math.max(...values);
    const hasMid = rule.mid && rule.mid.color;
    const mid = hasMid ? (rule.mid.value ?? (lo + hi) / 2) : null;
    const part = (v, a, b) => (b === a ? 1 : Math.max(0, Math.min(1, (v - a) / (b - a))));
    const at = (v) => (hasMid
      ? (v <= mid ? mixHex(rule.min.color, rule.mid.color, part(v, lo, mid)) : mixHex(rule.mid.color, rule.max.color, part(v, mid, hi)))
      : mixHex(rule.min.color, rule.max.color, part(v, lo, hi)));
    perRow = values.map((v) => toArgb(at(v)));
    legend = { scale: [rule.min.color, hasMid ? rule.mid.color : null, rule.max.color].filter(Boolean).map((c) => toArgb(c)), label: `Colour scale on ${inputName}` };
  } else {
    const test = (c, ri) => {
      if (c.and) return c.and.every((x) => test(x, ri));
      const ci = resolve(c.input);
      if (ci < 0) return false;
      const v = num(ri, ci);
      return [v === c.value, v > c.value, v >= c.value, v <= c.value, v < c.value][c.op] || false;
    };
    const base = seriesColor(targets[0]);
    perRow = data.map((_, ri) => { const hit = rule.cases.find((c) => test(c.test, ri)); return hit ? toArgb(hit.color) : base; });
    const ops = ["=", ">", "≥", "≤", "<"];
    const describe = (c) => (c.and ? c.and.map(describe).join(" and ") : `${ops[c.op] || "?"} ${c.value}`);
    const firstCol = resolve(rule.input);
    inputName = firstCol >= 0 ? header[firstCol] : rule.input;
    legend = { entries: rule.cases.map((c) => ({ label: `${inputName} ${describe(c.test)}`, argb: toArgb(c.color) })) };
  }

  // Tabs follow a key column; they stay right after a sort only if each key has one colour
  let dot = null;
  if (keyCol >= 0) {
    const byKey = {};
    let sortSafe = true;
    data.forEach((r, ri) => {
      const k = (r[keyCol] ?? "").trim();
      if (byKey[k] && byKey[k] !== perRow[ri]) sortSafe = false;
      byKey[k] = perRow[ri];
    });
    dot = { keyCol, colorOf: (v) => byKey[v] || null, sortSafe };
  }
  const note = approx
    ? `Colours of "${t.title || "a chart"}" follow a rule on ${rule.input}, which isn't in the exported data, so they are approximated.`
    : null;
  return { series: {}, rowCols: targets, rowKey: null, dot, perRow, legend, note };
}

/*
 * Works out the colours of a chart's table, to match the chart. Returns null for cards,
 * table / matrix visuals, or when there is no palette. Sources, best first:
 *  1. A column of hex colours (e.g. a Deneb "Bar Color" field) is what the visual draws
 *     with: each row takes its own colour.
 *  2. The report file (`info`, from colors.js, when it could be read): which columns are
 *     drawn as series (tooltips aren't), which field is the legend or the category, and
 *     colours set by hand (per measure, per item, or for every data point).
 *  3. Otherwise inferred: the n-th number column after the axis is the n-th series
 *     (helper columns such as sort keys, ranks and IDs are skipped); treemap / pie /
 *     donut colour each value of their first field; one series beside one text column
 *     is split by a legend.
 * Series take the theme's data colours in field order, legend values in series order
 * (seriesOrder) and pie / treemap categories in the order they appear, as Power BI does.
 * Returns { series, rowCols, hexCols, rowKey, dot, exact }:
 *  - series: { col: argb } columns drawn in one colour (header band, filled cells)
 *  - rowCols: columns whose cells take their row's colour (rows coloured by key or hex)
 *  - rowKey: { col, colorOf: { value: argb } } legend / category colours, for the caption
 *  - dot: { keyCol, colorOf(value) } the column that decides each row's colour
 *  - exact: true when the report file described this visual
 */
function columnColors(item, palette, info) {
  if (!palette || !palette.length || item.isKpi || TABLE_VISUAL_TYPES.includes(item.type)) return null;
  const { header, data, filled, numeric, hex, helper } = classifyColumns(item.rows);
  const cols = header.map((_, ci) => ci);
  const hexCols = cols.filter((ci) => hex[ci]);
  const color = (i) => toArgb(palette[i % palette.length]);
  const norm = header.map((h) => normalizeField(h));
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  // A field's column: its renamed label first, then its name ("Sum of Sales" and "Average of
  // Sales" both read "sales", so a column already taken is skipped)
  const findCol = (field, taken = {}) => {
    if (!field) return -1;
    const free = cols.filter((ci) => taken[ci] === undefined);
    const byLabel = field.label ? free.find((ci) => lower[ci] === field.label) : undefined;
    return byLabel !== undefined ? byLabel : free.find((ci) => norm[ci] === field.ref) ?? -1;
  };

  // Series columns: the report file's value fields when they match, else every number column
  let seriesCols = [];
  const fieldOf = {};   // series column → its field in the report file
  const indexOf = {};   // series column → its position among the visual's series
  if (info && info.measures.length) {
    info.measures.forEach((m, i) => {
      const ci = findCol(m, indexOf);
      if (ci >= 0 && numeric[ci] && !hex[ci]) { seriesCols.push(ci); fieldOf[ci] = m; indexOf[ci] = i; }
    });
  }
  // Whole years (2024) are a category such as a legend or axis, never a series
  const yearLike = cols.map((ci) => numeric[ci] && filled(ci).every((x) => /^(19|20)\d\d$/.test(x)));
  if (!seriesCols.length) {
    const tooltips = new Set((info && info.tooltips) || []);
    seriesCols = cols.filter((ci) => ci > 0 && !hex[ci] && !helper[ci] && numeric[ci] && !yearLike[ci] && !tooltips.has(norm[ci]));
    seriesCols.forEach((ci, i) => { indexOf[ci] = i; });
  }
  if (!seriesCols.length && !hexCols.length) return null;
  const exact = !!info;

  // 1. A hex colour column: every row in its own colour (one band colour when they all agree)
  if (hexCols.length) {
    const colors = new Set(filled(hexCols[0]).map((x) => x.toUpperCase()));
    const dot = { keyCol: hexCols[0], colorOf: (v) => (HEX_COLOR.test(v) ? toArgb(v) : null) };
    if (colors.size === 1) {   // one colour for every row: a series colour, no row tabs needed
      const one = toArgb([...colors][0]);
      return { series: Object.fromEntries(seriesCols.map((ci) => [ci, one])), rowCols: [], hexCols, rowKey: null, dot: null, exact };
    }
    const counts = {};
    filled(hexCols[0]).forEach((x) => { counts[x] = (counts[x] || 0) + 1; });
    const common = toArgb(Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0]);
    return { series: {}, legendColor: common, rowCols: seriesCols, hexCols, rowKey: null, dot, exact };
  }

  const overrides = (info && info.overrides) || {};
  const itemOverride = (v) => overrides["i:" + normalizeValue(v)];
  const seriesColor = (ci) => {
    const f = fieldOf[ci];
    const set = overrides["m:" + norm[ci]] || (f && (overrides["m:" + f.ref] || (f.label && overrides["m:" + f.label])));
    if (set) return toArgb(set);
    if (overrides.all) return toArgb(overrides.all);
    if (info && info.rule === "single") return color(0);
    return color(indexOf[ci]);
  };
  const keyedRows = (keyCol, colorOfValue, bySeries) => {
    const values = bySeries
      ? seriesOrder(data, keyCol, cols.filter((ci) => ci !== keyCol && !seriesCols.includes(ci) && !hex[ci]))
      : [...new Set(filled(keyCol))];   // order of first appearance
    return { col: keyCol, colorOf: Object.fromEntries(values.map((v, i) => [v, colorOfValue(v, i)])) };
  };

  // 2. Conditional formatting (fx) in the report file: colour scales and rules, row by row
  const fx = info && (info.rules || []).find((r) => r.scope === "all" || r.scope === "series" || seriesCols.some((ci) => "m:" + norm[ci] === r.scope));
  if (fx) {
    const result = colorRuleRows(fx, { header, data, cols, norm, lower, numeric, hex, seriesCols, findCol, info, seriesColor, title: item.displayTitle || item.title });
    if (result) return { ...result, hexCols, exact };
  }

  // 3. Rows coloured by a field: the legend (Series) or, for pie / donut / treemap, the category
  let rowKey = null;
  if (info && (info.rule === "legend" || info.rule === "item")) {
    let keyCol = findCol(info.rule === "legend" ? info.series : info.category);
    if (keyCol < 0 && info.rule === "item" && !numeric[0] && !hex[0]) keyCol = 0;
    if (keyCol >= 0) rowKey = keyedRows(keyCol, (v, i) => toArgb(itemOverride(v) || palette[i % palette.length]), info.rule === "legend");
  } else if (info && seriesCols.length === 1 && Object.keys(overrides).some((k) => k.startsWith("i:"))) {
    // One series with colours set per bar ("Show all"): set bars keep theirs, the rest the series colour
    const keyCol = findCol(info.category);
    const base = seriesColor(seriesCols[0]);
    if (keyCol >= 0) rowKey = keyedRows(keyCol, (v) => (itemOverride(v) ? toArgb(itemOverride(v)) : base));
  } else if (!info) {
    let keyCol = null, legend = false;
    if (COLOR_BY_CATEGORY_TYPES.includes(item.type) && !numeric[0]) {
      keyCol = 0;
    } else if (seriesCols.length === 1) {
      const textCols = cols.filter((ci) => ci > 0 && !hex[ci] && (!numeric[ci] || yearLike[ci]));
      if (textCols.length === 1) { keyCol = textCols[0]; legend = true; }
    }
    if (keyCol !== null) {
      const values = new Set(filled(keyCol));
      if (values.size > 1 && values.size <= palette.length) rowKey = keyedRows(keyCol, (v, i) => color(i), legend);
    }
  }

  if (rowKey) {
    const dot = { keyCol: rowKey.col, colorOf: (v) => rowKey.colorOf[v] || null };
    return { series: {}, rowCols: seriesCols, hexCols, rowKey, dot, exact };
  }
  const series = Object.fromEntries(seriesCols.map((ci) => [ci, seriesColor(ci)]));
  return { series, rowCols: [], hexCols, rowKey: null, dot: null, exact };
}

/* ── CSV → cells ──────────────────────────────────────────────────────── */
function parseCsv(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  text = text.replace(/^﻿/, "");
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") { row.push(field); field = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(field); rows.push(row); row = []; field = "";
    } else field += c;
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }
  return rows.filter((r) => r.some((v) => v !== ""));
}

/*
 * Removes columns that carry no information:
 *  - Power BI "Blank" helper columns (empty, or the same value in every row)
 *  - exact duplicate columns (same name and same values)
 */
function cleanRows(rows) {
  const [header, ...data] = rows;
  const keep = [];
  const seen = new Set();
  header.forEach((h, i) => {
    const name = (h ?? "").trim();
    const values = data.map((r) => (r[i] ?? "").trim());
    const isBlankHelper = (name === "" || name.toLowerCase() === "blank") && new Set(values).size <= 1;
    if (isBlankHelper) return;
    const signature = name + "\u0000" + values.join("\u0001");
    if (seen.has(signature)) return;
    seen.add(signature);
    keep.push(i);
  });
  return rows.map((r) => keep.map((i) => r[i] ?? ""));
}

/*
 * Power BI's CSV carries values as formatted text ("$1,234.50", "12.5%", "(300)").
 * Convert them to real Excel numbers with a matching number format, so they can be
 * filtered, sorted and summed. Anything that is not clearly a number stays text.
 */
function toCellValue(raw) {
  const s = (raw ?? "").trim();
  if (s === "") return { value: null };
  if (/^0\d+$/.test(s)) return { value: s };             // codes with leading zeros (ZIP, IDs)

  let t = s, negative = false;
  if (/^\(.*\)$/.test(t)) { negative = true; t = t.slice(1, -1); }
  const isPercent = t.endsWith("%");
  const currency = (t.match(/^-?\s*([$€£₹¥])/) || [])[1];
  const cleaned = t.replace(/[$€£₹¥%,\s]/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return { value: s };

  const n = Number(cleaned) * (negative ? -1 : 1);
  const decimals = Math.min((cleaned.split(".")[1] || "").length, 4);
  const dec = decimals ? "." + "0".repeat(decimals) : "";
  const thousands = s.includes(",") ? "#,##0" : "0";      // keep years like 2024 as 2024

  if (isPercent) return { value: n / 100, numFmt: `0${dec}%` };
  if (currency) return { value: n, numFmt: `"${currency}"${thousands}${dec}` };
  return { value: n, numFmt: `${thousands}${dec}` };
}

/*
 * Power BI exports measures that have no format string as bare numbers ("93687.3").
 * Such number columns get thousands separators, and two decimals when they have any.
 * Years, sort keys, ranks and IDs keep their plain digits.
 * Returns one format per column (null = leave as exported).
 */
const PLAIN_FORMAT = /^0(\.0+)?$/;
function readableFormats(header, data) {
  return header.map((name, ci) => {
    if (HELPER_COLUMN.test(name) || /\byear\b/i.test(name)) return null;
    const cells = data.map((r) => r[ci]).filter((c) => typeof c.value === "number");
    if (!cells.length || !cells.every((c) => PLAIN_FORMAT.test(c.numFmt || ""))) return null;
    const isYear = (n) => Number.isInteger(n) && n >= 1900 && n <= 2100;
    const isYearMonth = (n) => Number.isInteger(n) && n >= 190001 && n <= 210012;
    if (cells.every((c) => isYear(c.value)) || cells.every((c) => isYearMonth(c.value))) return null;
    return cells.some((c) => !Number.isInteger(c.value)) ? "#,##0.00" : "#,##0";
  });
}

/* Roughly how Excel shows a value with a number format (used to size columns) */
function displayText(value, fmt) {
  if (typeof value !== "number") return String(value ?? "");
  const f = fmt || "";
  const prefix = (f.match(/^"([^"]*)"/) || [])[1] || "";
  const dec = ((f.match(/\.(0+)/) || [])[1] || "").length;
  if (f.includes("%")) return (value * 100).toFixed(dec) + "%";
  const body = Math.abs(value).toLocaleString("en-US", {
    minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: f.includes("#,##0"),
  });
  return (value < 0 ? "-" : "") + prefix + body;
}

/* A compact KPI format (2.52M) for very long numbers; the exact value is shown beneath */
function compactFormat(value, fmt) {
  if (typeof value !== "number" || Math.abs(value) < 1e6 || displayText(value, fmt).length <= 14) return null;
  const prefix = (fmt && (fmt.match(/^"[^"]*"/) || [])[0]) || "";
  const a = Math.abs(value);
  const [scale, unit] = a >= 1e12 ? [",,,,", "T"] : a >= 1e9 ? [",,,", "B"] : [",,", "M"];
  return `${prefix}#,##0.00${scale}"${unit}"`;
}

/* ── Names ────────────────────────────────────────────────────────────── */

/* Excel Table names must be unique across the whole workbook */
function makeTableNamer() {
  const used = new Set();
  return (title) => {
    const base = "T_" + (title || "Visual").replace(/[^A-Za-z0-9_]/g, "_").slice(0, 40);
    let name = base, n = 2;
    while (used.has(name.toLowerCase())) name = `${base}_${n++}`;
    used.add(name.toLowerCase());
    return name;
  };
}

/* Sheet names: unique, at most 31 characters, none of \ / * ? : [ ] ' – and not
   "History", which Excel reserves */
function makeSheetNamer() {
  const used = new Set();
  return (title) => {
    let base = (title || "Page").replace(/[\\/*?:[\]']/g, "").trim().slice(0, 31) || "Page";
    if (base.toLowerCase() === "history") base = "History (page)";
    let name = base, n = 2;
    while (used.has(name.toLowerCase())) name = `${base.slice(0, 27)} (${n++})`;
    used.add(name.toLowerCase());
    return name;
  };
}

/* Excel Tables need unique, non-empty header names */
function uniqueHeaders(header) {
  const seen = {};
  return header.map((h, i) => {
    const base = (h || "").trim() || `Column ${i + 1}`;
    seen[base] = (seen[base] || 0) + 1;
    return seen[base] > 1 ? `${base} (${seen[base]})` : base;
  });
}

/* ── Column widths ────────────────────────────────────────────────────── */
function stepUp(units) {
  if (units > 60) return Math.ceil(units);
  return WIDTH_STEPS.find((s) => s >= units) ?? 60;
}

/*
 * Collects the width each column needs (in Excel width units), then sets them:
 * column A is the margin, measured columns round up to fixed steps, unused columns
 * become narrow gutters, KPI tiles widen their columns last, and a filler column keeps
 * the page at least `minPage` wide so the masthead rules span the title.
 */
function makeWidthTracker() {
  const need = {};
  const spans = [];
  return {
    want(col, units) { need[col] = Math.max(need[col] || 0, units); },
    span(c1, c2, units) { spans.push([c1, c2, units]); },
    apply(ws, lastCol, minPage) {
      const width = {};
      for (let c = 1; c <= lastCol; c++) {
        width[c] = c === 1 ? MARGIN_W : need[c] !== undefined ? stepUp(need[c]) : GUTTER_W;
      }
      for (const [c1, c2, units] of spans) {
        let have = 0;
        for (let c = c1; c <= c2; c++) have += width[c] || 0;
        if (have < units) width[c2] = (width[c2] || 0) + (units - have);
      }
      let total = 0;
      for (let c = 1; c <= lastCol; c++) total += width[c];
      let contentEnd = lastCol;
      if (total < minPage) {
        contentEnd = lastCol + 1;
        width[contentEnd] = minPage - total;
      }
      const bandEnd = contentEnd + 1;
      width[bandEnd] = MARGIN_W;
      for (const [c, w] of Object.entries(width)) ws.getColumn(Number(c)).width = Math.round(w * 100) / 100;
      return { contentEnd, bandEnd };
    },
  };
}

/* ── Block writers ────────────────────────────────────────────────────── */

/* KPI tiles show a title row only when the title adds something beyond the field name */
function kpiHasTitle(item) {
  return !!item.displayTitle && item.displayTitle.toLowerCase() !== String(item.rows[0][0] || "").toLowerCase();
}

/*
 * Size of a visual's block in rows / columns.
 *  - KPI tile: [title] + label + value (2 rows) + [exact value] + one row per extra field
 *  - exhibit: title + caption + header + data rows + closing row (the closing rule, which
 *    doubles as the "rows grouped" note when collapsing is on)
 */
function blockSize(item, options) {
  if (item.isKpi) {
    const [header, values = []] = item.rows;
    const hero = toCellValue(values[0]);
    const fmt = readableFormats([header[0]], [[hero]])[0] || hero.numFmt;
    const compact = !!compactFormat(hero.value, fmt);
    const height = (kpiHasTitle(item) ? 1 : 0) + 3 + (compact ? 1 : 0) + (header.length - 1);
    return { width: 2, height, note: false };
  }
  const n = item.rows.length - 1;
  const note = !!(options.collapse && n > options.visibleRows);
  return { width: item.rows[0].length, height: HEAD_ROWS + n + 1, note };
}

/*
 * KPI / card as a scorecard tile: uppercase label, large value, optional sub-fields.
 * strip: { titleRow, height } for tiles in a key-figure strip, so every tile reserves the
 * same title row and reaches the same bottom row, and the caps join into one ribbon.
 */
function writeKpiBlock(ws, r, col, item, widths, lay, divider, strip) {
  const fonts = lay.fonts;
  const [header, values = []] = item.rows;
  const cells = header.map((_, i) => toCellValue(values[i]));
  const formats = readableFormats(header, [cells]);
  const fmtOf = (i) => formats[i] || cells[i].numFmt || null;
  // The value colours set on the card in the report, when the report file was read
  const info = lay.reportColors && item.name ? lay.reportColors.visuals[item.name] : null;
  const valueColor = (i) => {
    const vc = (info && info.valueColors) || {};
    const hex = vc["m:" + normalizeField(header[i])] || vc.all;
    return hex ? toArgb(hex) : null;
  };
  const tinted = (font, i) => (valueColor(i) ? { ...font, color: { argb: valueColor(i) } } : font);

  let top = r;
  if (kpiHasTitle(item)) {
    const t = ws.getCell(r, col);
    t.value = item.displayTitle;
    t.font = fonts.exTitle;
    t.alignment = { vertical: "bottom" };
    widths.span(col, col + 1, Math.ceil(item.displayTitle.length * (fonts.exTitle.name === "Georgia" ? 1.2 : 1.1)) + 2);
    top = r + 1;
  } else if (strip && strip.titleRow) {
    top = r + 1;
  }

  // Merges keep each cell's own style, so the cap / rule borders survive on every cell
  const labelText = String(header[0] ?? "").toLocaleUpperCase();
  ws.mergeCellsWithoutStyle(top, col, top, col + 1);
  const label = ws.getCell(top, col);
  label.value = labelText;
  label.font = fonts.kpiLabel;
  label.alignment = { horizontal: "left", vertical: "bottom", indent: 1 };

  const hero = cells[0];
  const heroFmt = fmtOf(0);
  const compact = compactFormat(hero.value, heroFmt);
  ws.mergeCellsWithoutStyle(top + 1, col, top + 2, col + 1);
  const v = ws.getCell(top + 1, col);
  v.value = hero.value;
  if (compact || heroFmt) v.numFmt = compact || heroFmt;
  v.font = tinted(fonts.kpiValue, 0);
  v.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

  let next = top + 3;
  if (compact) {
    ws.mergeCellsWithoutStyle(next, col, next, col + 1);
    const e = ws.getCell(next, col);
    e.value = hero.value;
    e.numFmt = heroFmt || "#,##0.00";
    e.font = fonts.kpiExact;
    e.alignment = { horizontal: "left", vertical: "top", indent: 1 };
    next++;
  }
  for (let i = 1; i < header.length; i++, next++) {
    const l = ws.getCell(next, col);
    l.value = header[i];
    l.font = fonts.kpiSubLabel;
    l.alignment = { vertical: "middle", indent: 1 };
    const sv = ws.getCell(next, col + 1);
    sv.value = cells[i].value;
    if (fmtOf(i)) sv.numFmt = fmtOf(i);
    sv.font = tinted(fonts.kpiSubValue, i);
    sv.alignment = { horizontal: "right", vertical: "middle", indent: 1 };
    widths.want(col, Math.min(30, String(header[i]).length + 3));
    widths.want(col + 1, displayText(cells[i].value, fmtOf(i)).length + 3);
  }

  const last = Math.max(next - 1, strip ? r + strip.height - 1 : 0);
  for (let rr = top; rr <= last; rr++) {
    for (let cc = col; cc <= col + 1; cc++) {
      const cell = ws.getCell(rr, cc);
      cell.fill = SURFACE_FILL;
      const b = {};
      if (rr === top) b.top = { style: "medium", color: { argb: lay.accent.cap } };
      if (rr === last) b.bottom = LIGHT_RULE;
      if (divider && cc === col) b.left = LIGHT_RULE;
      if (b.top || b.bottom || b.left) cell.border = b;
    }
  }

  const heroLen = compact ? 9 : displayText(hero.value, heroFmt).length;
  widths.want(col, 12);
  widths.want(col + 1, 10);
  widths.span(col, col + 1, Math.max(24, labelText.length + 4, Math.ceil(heroLen * 2.1) + 4));
}

/* Caption runs: a legend of the chart's colours, or "Table · 24 rows" */
function captionRuns(item, colors, header, fonts, units) {
  const n = item.rows.length - 1;
  const plain = [[`${friendlyType(item.type)}  ·  ${n} row${n === 1 ? "" : "s"}`, fonts.caption]];
  if (!colors) return plain;

  const legend = (entries) => {
    const runs = [];
    let len = 0, shown = 0;
    for (const e of entries) {
      const add = e.label.length * 0.8 + 4;
      if (shown && len + add > units - 6) break;
      runs.push([KEY_GLYPH + "  ", fonts.dot(e.argb)], [e.label + "      ", fonts.caption]);
      len += add;
      shown++;
    }
    return { runs, shown };
  };

  if (colors.legend && colors.legend.scale) {
    return [...colors.legend.scale.map((argb) => [KEY_GLYPH, fonts.dot(argb)]), ["  " + colors.legend.label, fonts.caption]];
  }
  if (colors.legend && colors.legend.entries) {
    const { runs, shown } = legend(colors.legend.entries);
    if (shown < colors.legend.entries.length) runs.push([`+${colors.legend.entries.length - shown} more`, fonts.captionMore]);
    return runs;
  }

  if (colors.rowKey) {
    const entries = Object.entries(colors.rowKey.colorOf).map(([label, argb]) => ({ label, argb }));
    const { runs, shown } = legend(entries);
    if (shown === entries.length) return runs;
    const few = entries.slice(0, 4).map((e) => [KEY_GLYPH + " ", fonts.dot(e.argb)]);
    return [...few, [` Colour by ${header[colors.rowKey.col]}`, fonts.caption]];
  }

  const seriesEntries = Object.entries(colors.series).map(([ci, argb]) => ({ label: header[ci], argb }));
  if (colors.hexCols.length) {
    const first = seriesEntries[0] || (colors.legendColor && colors.rowCols.length
      ? { label: header[colors.rowCols[0]], argb: colors.legendColor } : null);
    return [
      ...(first ? [[KEY_GLYPH + "  ", fonts.dot(first.argb)], [first.label, fonts.caption]] : []),
      [`${first ? "  ·  " : ""}row colours from ${header[colors.hexCols[0]]}`, fonts.caption],
    ];
  }
  const { runs, shown } = legend(seriesEntries);
  if (shown < seriesEntries.length) runs.push([`+${seriesEntries.length - shown} more`, fonts.captionMore]);
  return runs.length ? runs : plain;
}

/*
 * Exhibit: title row, caption row, then a real Excel Table. "Executive" uses Excel's
 * "None" table style with ink rules and hairlines; a built-in style keeps its own look.
 * Header cells keep exactly the column name (anything else makes Excel repair the file).
 * Style objects are made once per column and shared by every cell.
 */
function writeDataTable(ws, r, col, item, tableName, widths, lay, number) {
  const fonts = lay.fonts;
  const cls = classifyColumns(item.rows);
  const header = uniqueHeaders(item.rows[0]);
  const data = cls.data.map((row) => header.map((_, i) => toCellValue(row[i])));
  const formats = readableFormats(header, data);
  const n = data.length;
  const executive = !BUILT_IN_STYLES.includes(lay.tableStyle);

  // Column widths: header, longest text, widest formatted number; numbers in one table
  // share a width so their columns read as a set
  const units = header.map((h, ci) => {
    let u = Math.min(45, Math.ceil(h.length * 1.05) + 5);
    if (cls.numeric[ci]) {
      let widest = null, negative = false;
      for (const row of data) {
        const c = row[ci];
        if (typeof c.value !== "number") continue;
        if (c.value < 0) negative = true;
        if (!widest || Math.abs(c.value) > Math.abs(widest.value)) widest = c;
      }
      if (widest) {
        const len = displayText(Math.abs(widest.value), formats[ci] || widest.numFmt).length + (negative ? 1 : 0);
        u = Math.max(u, Math.ceil(len * 0.95) + 4);
      }
    } else {
      let longest = 0;
      for (const row of data) if (row[ci].value != null) longest = Math.max(longest, String(row[ci].value).length);
      u = Math.max(u, Math.min(45, Math.ceil(longest * 0.95) + 3));
    }
    return Math.max(u, 10);
  });
  const numericUnits = units.filter((_, ci) => cls.numeric[ci]);
  if (numericUnits.length > 1) {
    const target = Math.min(Math.max(...numericUnits), 18);
    units.forEach((u, ci) => { if (cls.numeric[ci]) units[ci] = Math.max(u, target); });
  }
  units.forEach((u, ci) => widths.want(col + ci, u));
  const totalUnits = units.reduce((a, b) => a + b, 0);
  // The title overflows to the right, up to the next visual's title one gutter along:
  // widen the table's last column when the title needs more (capped, so it stays a table)
  const titleLen = (number ? 4 : 0) + String(item.displayTitle || "").length;
  const titleUnits = Math.ceil(titleLen * (fonts.exTitle.name === "Georgia" ? 1.2 : 1.1)) + 1 - GUTTER_W;
  widths.span(col, col + header.length - 1, Math.min(titleUnits, totalUnits + 40));

  // Title and caption
  const t = ws.getCell(r, col);
  t.value = rich(number ? [String(number).padStart(2, "0") + "  ", fonts.exNum] : null, [item.displayTitle, fonts.exTitle]);
  t.alignment = { vertical: "bottom" };
  const cap = ws.getCell(r + 1, col);
  cap.value = rich(...captionRuns(item, item.colors, header, fonts, totalUnits));
  cap.alignment = { vertical: "middle" };

  // The table
  const hr = r + 2;
  ws.addTable({
    name: tableName,
    ref: ws.getCell(hr, col).address,
    headerRow: true,
    style: executive
      ? { theme: null, showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false }
      : { theme: lay.tableStyle, showRowStripes: true },
    columns: header.map((name) => ({ name, filterButton: true })),
    rows: data.map((row) => row.map((c) => c.value)),
  });

  header.forEach((_, ci) => {
    const num = cls.numeric[ci];
    const align = { vertical: "middle", horizontal: num ? "right" : "left", indent: num ? 2 : 1 };
    const h = ws.getCell(hr, col + ci);
    h.font = executive ? fonts.header : fonts.headerOnStyle;
    h.alignment = align;
    if (executive) {
      h.fill = SURFACE_FILL;
      h.border = HEADER_BORDER;
    }
    for (let ri = 0; ri < n; ri++) {
      const cell = ws.getCell(hr + 1 + ri, col + ci);
      cell.font = fonts.body;
      cell.alignment = align;
      if (executive) cell.border = BODY_BORDER;
      const fmt = formats[ci] || data[ri][ci].numFmt;
      if (fmt) cell.numFmt = fmt;
    }
  });

  if (item.colors) applyVisualColors(ws, hr, col, item, item.colors, lay);
}

/* Text colour that reads best on a filled cell: white or ink, whichever contrasts more */
function textOn(argb) {
  return contrast(argb, C.WHITE) >= contrast(argb, C.INK) ? C.WHITE : C.INK;
}

/*
 * Paints a chart table in the chart's colours:
 *  - series drawn in one colour: the header cell becomes a band in that colour and, with
 *    `lay.colorCells`, every value cell below is filled with it
 *  - rows coloured by a field (legend, pie, treemap) or a hex column: each row's value
 *    cells are filled with the row's colour (Excel moves fills along when sorting), and
 *    the row gets a colour tab in the column left of the table, painted by paintRowTabs
 *    once column widths are known (lay.pendingTabs)
 * Text on a filled cell is white or ink, whichever reads better.
 */
function applyVisualColors(ws, hr, col, item, colors, lay) {
  const fonts = lay.fonts;
  const data = item.rows.slice(1);
  const first = hr + 1, last = hr + data.length;
  const styles = new Map();   // argb → shared fill / font objects (big tables stay fast)
  const styleOf = (argb) => {
    if (!styles.has(argb)) {
      const fill = solidFill(argb);
      styles.set(argb, { fill, font: { ...fonts.body, color: { argb: textOn(argb) } }, headFont: { ...fonts.header, color: { argb: textOn(argb) } } });
    }
    return styles.get(argb);
  };
  const fillCell = (cell, argb) => {
    const s = styleOf(argb);
    cell.fill = s.fill;
    cell.font = s.font;
    cell.border = GAP_BORDER;
  };

  for (const [ci, argb] of Object.entries(colors.series)) {
    const c = col + Number(ci);
    const h = ws.getCell(hr, c);
    h.fill = styleOf(argb).fill;
    h.font = styleOf(argb).headFont;
    if (lay.colorCells) for (let r = first; r <= last; r++) fillCell(ws.getCell(r, c), argb);
  }

  const dot = colors.dot;
  if (!dot && !colors.perRow) return;
  const tabs = dot ? { g: col - 1, keyCol: col + dot.keyCol, first, last, cells: [], seen: new Map(), noRules: dot.sortSafe === false } : null;
  data.forEach((row, ri) => {
    const key = dot ? (row[dot.keyCol] ?? "").trim() : "";
    const argb = colors.perRow ? colors.perRow[ri] : key ? dot.colorOf(key) : null;
    if (!argb) return;
    if (lay.colorCells) for (const ci of colors.rowCols) fillCell(ws.getCell(first + ri, col + ci), argb);
    if (!tabs) return;
    tabs.cells.push([first + ri, argb]);
    if (key && !tabs.seen.has(key)) tabs.seen.set(key, argb);
  });
  if (!tabs) return;
  if (tabs.g >= 1 && tabs.cells.length && lay.pendingTabs) lay.pendingTabs.push(tabs);
}

/*
 * Row colour tabs, painted once column widths are known. In a narrow gutter a tab is a
 * filled cell; where the column is wider (another band's data column) it is a swatch
 * glyph at the right edge, so it never becomes a wide coloured block. The tabs aren't part
 * of the table, so up to MAX_CF_RULES colours conditional formatting alone colours each
 * tab from its own row's key cell, and tabs stay right after a sort or filter (rows with
 * no colour stay blank). Above that the colours are static.
 */
function paintRowTabs(ws, pendingTabs, fonts) {
  for (const t of pendingTabs) {
    const narrow = (ws.getColumn(t.g).width || GUTTER_W) <= GUTTER_W + 0.01;
    const K = ws.getColumn(t.keyCol).letter;
    let rules = !t.noRules && t.seen.size <= MAX_CF_RULES ? [] : null;
    for (const [key, argb] of rules ? t.seen : []) {
      const typed = toCellValue(key).value;
      const literal = typeof typed === "number" ? String(typed) : `"${String(key).replace(/"/g, '""')}"`;
      if (literal.length > 250) { rules = null; break; }
      rules.push({
        type: "expression",
        formulae: [`$${K}${t.first}=${literal}`],
        style: narrow ? { fill: { type: "pattern", pattern: "solid", bgColor: { argb } } } : { font: { color: { argb } } },
      });
    }
    const align = { horizontal: "right", vertical: "middle" };
    for (const [r, argb] of t.cells) {
      const cell = ws.getCell(r, t.g);
      if (narrow) {
        if (!rules) cell.fill = solidFill(argb);
        cell.border = GAP_BORDER;
      } else {
        cell.value = KEY_GLYPH;
        cell.font = fonts.dot(rules ? C.WHITE : argb);   // with rules, only matching rows show a colour
        cell.alignment = align;
      }
    }
    if (rules && rules.length) {
      const G = ws.getColumn(t.g).letter;
      ws.addConditionalFormatting({ ref: `${G}${t.first}:${G}${t.last}`, rules });
    }
  }
}

/* ── Page layout ──────────────────────────────────────────────────────── */

/* Bands: visuals whose vertical extents overlap, i.e. one "row" of the page */
function groupIntoBands(items) {
  const sorted = [...items].sort((a, b) => (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x));
  const bands = [];
  for (const item of sorted) {
    const band = bands[bands.length - 1];
    const bottom = item.layout.y + item.layout.height;
    if (band && item.layout.y < band.bottom - PX_TOLERANCE) {
      band.items.push(item);
      band.bottom = Math.max(band.bottom, bottom);
    } else {
      bands.push({ items: [item], bottom });
    }
  }
  return bands;
}

/* Stacks: visuals in a band that share horizontal space, kept top to bottom */
function groupIntoStacks(items) {
  const sorted = [...items].sort((a, b) => (a.layout.x - b.layout.x) || (a.layout.y - b.layout.y));
  const stacks = [];
  for (const item of sorted) {
    const left = item.layout.x, right = item.layout.x + item.layout.width;
    const stack = stacks.find((s) => left < s.right - PX_TOLERANCE && right > s.left + PX_TOLERANCE);
    if (stack) {
      stack.items.push(item);
      stack.left = Math.min(stack.left, left);
      stack.right = Math.max(stack.right, right);
    } else {
      stacks.push({ items: [item], left, right });
    }
  }
  stacks.forEach((s) => s.items.sort((a, b) => a.layout.y - b.layout.y));
  return stacks.sort((a, b) => a.left - b.left);
}

/* Rows of a planned table that may collapse: data rows past its visible rows, plus padding before its note */
function collapsibleRange(p) {
  if (p.item.isKpi) return null;
  const n = p.item.rows.length - 1;
  if (n <= p.vis) return null;
  const end = p.noteRow ? p.noteRow - 1 : p.r + HEAD_ROWS + n - 1;
  return [p.r + HEAD_ROWS + p.vis, end];
}

/* Rows of a planned block that must always stay visible: a whole tile, or a table's
   head and first rows (and its closing row when nothing of it collapses) */
function topRange(p) {
  if (p.item.isKpi) return [p.r, p.r + p.size.height - 1];
  const n = p.item.rows.length - 1;
  if (n <= p.vis) return [p.r, p.r + HEAD_ROWS + n];
  return [p.r, p.r + HEAD_ROWS - 1 + p.vis];
}

/* Moves a block, and everything below it in the same stack, down by `delta` rows */
function shiftStack(plan, p, delta) {
  const from = p.r;
  for (const q of plan) {
    if (q.stackIndex === p.stackIndex && q.r >= from) {
      q.r += delta;
      if (q.noteRow) q.noteRow += delta;
    }
  }
}

/*
 * Excel hides whole rows, so side-by-side blocks must not get in each other's way:
 *  1. A table's collapsible rows must not sit next to a neighbour's visible rows.
 *     The table first shows a few more rows (up to `maxExtend`), so band tops stay
 *     level; if that isn't enough it moves down instead.
 *  2. A table's note row must not sit inside a neighbour's collapsible rows,
 *     otherwise it splits the neighbour. The note moves down below them, so both
 *     notes share one row and one [+] / [–] button.
 */
function alignNoteRows(plan, maxExtend) {
  for (let guard = 0; guard < 200; guard++) {
    let moved = false;
    for (const p of plan) {
      const myRange = collapsibleRange(p);
      for (const o of plan) {
        if (o === p || o.stackIndex === p.stackIndex) continue;

        const top = topRange(o);
        if (myRange && myRange[0] <= top[1] && top[0] <= myRange[1]) {
          const need = top[1] + 1 - myRange[0];
          if (p.extra + need <= maxExtend) {
            p.extra += need;
            p.vis += need;
          } else {
            shiftStack(plan, p, need);
          }
          moved = true;
          break;
        }

        const range = collapsibleRange(o);
        if (p.noteRow && range && p.noteRow >= range[0] && p.noteRow <= range[1]) {
          const delta = range[1] + 1 - p.noteRow;
          p.noteRow += delta;
          p.size.height += delta;
          for (const q of plan) {
            if (q.stackIndex === p.stackIndex && q.r > p.r) {
              q.r += delta;
              if (q.noteRow) q.noteRow += delta;
            }
          }
          moved = true;
          break;
        }
      }
      if (moved) break;
    }
    if (!moved) break;
  }
}

/*
 * Writes every visual starting at `startRow`. Returns the first free row below the
 * layout, the last column used, the number of tables, whether rows were collapsed and
 * the note rows to fill in once column widths are known.
 */
function writePageLayout(ws, items, startRow, widths, options, lay, nameTable) {
  const stacked = options.layout === "stacked";
  const threshold = options.collapse ? options.visibleRows : Infinity;
  const maxExtend = Math.max(12, options.visibleRows);
  const tables = [];
  const rowUse = [];  // per Excel row: does it hold collapsible table rows ("deep") / anything else ("other")?
  const mark = (r, kind) => { (rowUse[r] = rowUse[r] || { deep: false, other: false })[kind] = true; };

  // Titles (generic ones described from their columns, duplicates numbered) and colours
  const usedTitles = {};
  const colorNotes = [];
  for (const item of items) {
    let title = isGenericTitle(item.title, item.type) ? (item.isKpi ? null : describeTitle(item)) : item.title;
    if (title) {
      const key = title.toLowerCase();
      usedTitles[key] = (usedTitles[key] || 0) + 1;
      if (usedTitles[key] > 1) title = `${title} (${usedTitles[key]})`;
    }
    item.displayTitle = title;
    const info = lay.reportColors && item.name ? lay.reportColors.visuals[item.name] || null : null;
    item.colors = lay.palette && !item.isKpi ? columnColors(item, lay.palette, info) : null;
    if (item.colors && item.colors.note) colorNotes.push(item.colors.note);
  }
  const executive = !BUILT_IN_STYLES.includes(lay.tableStyle);
  const hasTab = (it) => !it.isKpi && !!(it.colors && it.colors.dot);

  const bands = stacked
    ? groupIntoBands(items).flatMap((b) => [...b.items].sort((a, c) => a.layout.x - c.layout.x).map((it) => ({ items: [it] })))
    : groupIntoBands(items);

  const anchors = [];   // { x, col }: page x positions already given a column
  let bandTop = startRow;
  let nextRow = startRow;
  let lastCol = CONTENT_COL;
  let exhibit = 0;

  bands.forEach((band, bandIndex) => {
    // 1. Plan where every visual of this band goes
    const plan = [];
    const allCards = !stacked && band.items.every((it) => it.isKpi);
    const stacks = stacked ? [{ items: band.items, left: 0 }] : groupIntoStacks(band.items);
    // Key-figure strip: the first tile of every stack shares one title row and one height
    let strip = null;
    if (allCards) {
      const heads = stacks.map((s) => s.items[0]);
      const titleRow = heads.some(kpiHasTitle);
      const height = Math.max(...heads.map((it) => blockSize(it, options).height + (titleRow && !kpiHasTitle(it) ? 1 : 0)));
      strip = { titleRow, height };
    }
    let cursorCol = CONTENT_COL;
    stacks.forEach((stack, stackIndex) => {
      // Row colour tabs sit in the column left of a table: keep a blank gutter before them,
      // so they don't touch the visual to the left
      // (lining up with an earlier visual at the same x wins; paintRowTabs copes with wide columns)
      let col = cursorCol + (stackIndex > 0 && stack.items.some(hasTab) ? 1 : 0);
      if (!stacked && !allCards) {
        const anchor = anchors.find((a) => Math.abs(a.x - stack.left) <= SNAP_PX && a.col >= cursorCol);
        if (anchor) col = anchor.col;
        else anchors.push({ x: stack.left, col });
      }
      let r = bandTop, stackWidth = 0;
      stack.items.forEach((item, i) => {
        const inStrip = strip && i === 0;
        const size = blockSize(item, options);
        if (inStrip) size.height = strip.height;
        const n = item.rows.length - 1;
        plan.push({
          item, col, r, size, stackIndex, vis: threshold, extra: 0,
          noteRow: size.note ? r + HEAD_ROWS + n : null,
          divider: allCards && stackIndex > 0,
          strip: inStrip ? strip : null,
        });
        stackWidth = Math.max(stackWidth, size.width);
        r += size.height + 1;
      });
      cursorCol = col + stackWidth + (allCards ? 0 : 1);
    });

    // 2. Keep side-by-side tables from splitting each other when rows collapse
    if (options.collapse) alignNoteRows(plan, maxExtend);

    // 3. Write the band
    let bandBottom = bandTop;
    for (const p of plan) {
      const { item, col, r, size } = p;
      lastCol = Math.max(lastCol, col + size.width - 1);
      if (item.isKpi) {
        writeKpiBlock(ws, r, col, item, widths, lay, p.divider, p.strip);
        for (let i = 0; i < size.height; i++) mark(r + i, "other");
      } else {
        const number = options.exhibitNumbers ? ++exhibit : null;
        writeDataTable(ws, r, col, item, nameTable(item.displayTitle), widths, lay, number);
        const n = item.rows.length - 1;
        for (let i = 0; i < HEAD_ROWS; i++) mark(r + i, "other");
        for (let i = 0; i < n; i++) mark(r + HEAD_ROWS + i, i >= p.vis ? "deep" : "other");
        const closeRow = p.noteRow || r + HEAD_ROWS + n;
        for (let x = r + HEAD_ROWS + n; x < closeRow; x++) mark(x, "deep");   // padding lined up with a neighbour
        mark(closeRow, "other");
        if (executive) paintRange(ws, closeRow, col, col + size.width - 1, { border: { top: INK_RULE } });
        tables.push({ top: r, col, w: size.width, n, vis: p.vis, noteRow: p.noteRow });
      }
      bandBottom = Math.max(bandBottom, r + size.height);
    }
    nextRow = bandBottom;
    if (bandIndex < bands.length - 1) bandTop = bandBottom + BAND_GAP;
  });

  // Collapse only rows that hold nothing but a long table's extra rows
  let grouped = false;
  if (options.collapse) {
    rowUse.forEach((use, r) => {
      if (use && use.deep && !use.other) {
        const excelRow = ws.getRow(r);
        excelRow.outlineLevel = 1;
        excelRow.hidden = true;
        grouped = true;
      }
    });
  }

  const pendingNotes = [];
  for (const t of tables) {
    if (!t.noteRow) continue;
    let hidden = 0;
    for (let i = t.vis; i < t.n; i++) if (ws.getRow(t.top + HEAD_ROWS + i).hidden) hidden++;
    if (hidden) pendingNotes.push({ row: t.noteRow, col: t.col, w: t.w, n: t.n, hidden });
  }

  return { nextRow, lastCol, tableCount: tables.length, grouped, pendingNotes, colorNotes };
}

/* ── Collecting a page ────────────────────────────────────────────────── */

/* A slicer with nothing selected has no filter to name its field, so read the field
   names from the header of its data (one row is enough). */
async function slicerFieldNames(visual, fallback) {
  try {
    const result = await visual.exportData(models.ExportDataType.Summarized, 1);
    const header = parseCsv(result.data)[0];
    if (header && header.length) return header.join(", ");
  } catch (e) {
    // fall through to the visual's title
  }
  return fallback;
}

function isHiddenVisual(v) {
  return !!(v.layout && v.layout.displayState && v.layout.displayState.mode === models.VisualContainerDisplayMode.Hidden);
}

/* Visuals worth offering in the options checklist (not slicers, not decorative) */
async function listDataVisuals(page) {
  const visuals = await page.getVisuals();
  return visuals
    .filter((v) => v.type !== "slicer" && !NON_DATA_TYPES.includes(v.type) && !isHiddenVisual(v))
    .sort((a, b) => (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x))
    .map((v) => ({ name: v.name, title: v.title || v.type, type: v.type }));
}

/* Makes `page` the active page and waits until Power BI has rendered it */
async function activatePage(page) {
  const active = await embeddedReport.getActivePage();
  if (active.name === page.name) return;
  await new Promise((resolve) => {
    let finished = false;
    const done = () => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      embeddedReport.off("rendered", done);
      resolve();
    };
    const timer = setTimeout(done, PAGE_RENDER_TIMEOUT_MS);
    embeddedReport.on("rendered", done);
    page.setActive().catch(done);
  });
}

/*
 * Reads one page: filter chips, exportable items and notes. Visuals are exported one
 * at a time (Power BI rejects concurrent exportData calls).
 * options.excludedVisuals: Set of visual names to leave out (current-page mode only).
 */
async function collectPage(page, options, onProgress, isCancelled) {
  const visuals = (await page.getVisuals()).sort((a, b) =>
    (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x));

  const filterChips = [];
  if (options.includeFilters) {
    filterParts(await embeddedReport.getFilters()).forEach((p) => filterChips.push({ all: false, ...p }));
    filterParts(await page.getFilters()).forEach((p) => filterChips.push({ all: false, ...p }));
  }

  const items = [];
  const notes = [];
  for (let i = 0; i < visuals.length; i++) {
    if (isCancelled()) throw { cancelled: true };
    const v = visuals[i];
    const title = v.title || `${v.type} (${v.name})`;
    onProgress(i + 1, visuals.length, title);

    if (isHiddenVisual(v)) { notes.push(`Skipped hidden visual: ${title}`); continue; }
    if (options.excludedVisuals && options.excludedVisuals.has(v.name)) continue;

    // Slicers become filter chips: "Field  selection", or "Field  All"
    if (v.type === "slicer") {
      if (!options.includeFilters) continue;
      try {
        const state = await v.getSlicerState();
        const parts = filterParts(state.filters);
        if (parts.length) {
          parts.forEach((p) => filterChips.push({ all: false, ...p }));
        } else {
          const target = (state.targets && state.targets[0]) || {};
          const field = target.column || target.hierarchyLevel || target.measure || target.hierarchy ||
            await slicerFieldNames(v, title);
          filterChips.push({ field, value: "All", all: true });
        }
      } catch (e) {
        notes.push(`Could not read the selection of "${title}": ${errorText(e)}`);
      }
      continue;
    }

    // Every other visual: try to export. Decorative ones (text box, shape, button…)
    // usually have no data; they are included only if Power BI returns some.
    try {
      const result = await v.exportData(models.ExportDataType.Summarized, 30000);
      const rows = cleanRows(parseCsv(result.data));
      if (rows.length < 2 || rows[0].length === 0) {
        if (!NON_DATA_TYPES.includes(v.type)) notes.push(`No data in: ${title}`);
        continue;
      }
      const isKpi = KPI_TYPES.includes(v.type) && rows.length === 2;
      if (!isKpi && rows[0].length === 1) {
        if (!NON_DATA_TYPES.includes(v.type)) {
          notes.push(`Skipped "${title}": a single list of values with no measures (for example a what-if parameter)`);
        }
        continue;
      }
      items.push({ title, name: v.name, type: v.type, layout: v.layout, rows, isKpi });
      log(`Exported "${title}" (${rows.length - 1} rows)`);
    } catch (e) {
      if (NON_DATA_TYPES.includes(v.type)) continue;
      notes.push(`Export failed for "${title}" (${v.type}): ${errorText(e)}`);
      log(`Export failed for "${title}": ${errorText(e)}`);
    }
  }
  return { filterChips, items, notes };
}

/* ── Writing a sheet ──────────────────────────────────────────────────── */

/* Filter chips split into lines of about `max` characters; a chip is never split */
function chipLines(chips, max) {
  const lines = [];
  let line = [], len = 0;
  for (const chip of chips) {
    const add = chip.field.length + 2 + Math.min(String(chip.value).length, 80) + 7;
    if (line.length && len + add > max) { lines.push(line); line = []; len = 0; }
    line.push(chip);
    len += add;
  }
  if (line.length) lines.push(line);
  return lines;
}

function chipRuns(line, fonts) {
  const runs = [];
  line.forEach((chip, i) => {
    if (i) runs.push(["   │   ", fonts.chipSep]);
    const value = String(chip.value);
    runs.push([chip.field + "  ", fonts.chipField], [value.length > 80 ? value.slice(0, 79) + "…" : value, chip.all ? fonts.chipField : fonts.chipValue]);
  });
  return runs;
}

/* Header / footer text: & doubled, each part kept short (Excel caps a section at 255) */
function hfText(s) {
  return String(s || "").slice(0, 60).replace(/&/g, "&&");
}

/*
 * ctx: { pageIndex, pageCount, dateText }. Order: masthead text → layout → column
 * widths → rules and fills (need the final width) → logo → collapse notes (wording
 * fits the table width) → notes and footer → row heights, outline, print setup.
 */
function writePageSheet(wb, sheetName, page, data, options, branding, logo, nameTable, ctx) {
  const accent = accentPlan(toArgb(branding.accent));
  const fonts = makeFonts(branding.titleFont);
  const lay = {
    fonts, accent,
    tableStyle: branding.tableStyle || EXECUTIVE,
    palette: options.matchColors ? options.palette : null,
    reportColors: options.matchColors ? options.reportColors || null : null,
    colorCells: options.colorCells !== false,
    pendingTabs: [],   // row colour tabs, painted once column widths are known
  };
  const ws = wb.addWorksheet(sheetName, {
    views: [{ showGridLines: false, zoomScale: 100 }],
    properties: { tabColor: { argb: accent.tab }, defaultRowHeight: ROW_H },
  });
  const widths = makeWidthTracker();
  const B = CONTENT_COL;
  const heights = {};
  let row = 1;

  // 1. Masthead text (fills and rules are painted once the page width is known)
  heights[row++] = 4.5;                          // accent rule
  heights[row++] = logo ? 12 : 15;               // spacer
  let logoRow = null;
  if (logo) { logoRow = row; heights[row++] = 36; }

  const eyebrowText = [branding.companyName, currentReportName].filter(Boolean).join("  ·  ").toLocaleUpperCase();
  const eyebrow = ws.getCell(row, B);
  eyebrow.value = eyebrowText;
  eyebrow.font = fonts.eyebrow;
  eyebrow.alignment = { vertical: "bottom" };
  heights[row++] = 15;

  const title = ws.getCell(row, B);
  title.value = page.displayName;
  title.font = fonts.title;
  title.alignment = { vertical: "bottom" };
  heights[row++] = 34;

  const chips = options.includeFilters ? data.filterChips : [];
  const active = chips.filter((c) => !c.all).length;
  const metaText = [
    ctx.dateText,
    options.includeFilters && active ? `${active} filter${active === 1 ? "" : "s"} applied` : null,
    ctx.pageCount > 1 ? `Page ${ctx.pageIndex} of ${ctx.pageCount}` : null,
  ].filter(Boolean).join("  ·  ");
  const metaRow = row;
  const meta = ws.getCell(row, B);
  meta.value = metaText;
  meta.font = fonts.meta;
  meta.alignment = { vertical: "middle" };
  heights[row++] = 20;
  heights[row++] = 10;                           // spacer

  const filterRows = [];
  let longestFilterLine = 0;
  if (options.includeFilters) {
    const lines = chipLines(chips, 100);
    lines.forEach((line) => {
      longestFilterLine = Math.max(longestFilterLine, chipRuns(line, fonts).reduce((a, [t]) => a + t.length, 0) + 11);
    });
    if (lines.length <= 1) {
      const c = ws.getCell(row, B);
      c.value = rich(["FILTERS    ", fonts.filterLabel], ...(lines.length ? chipRuns(lines[0], fonts) : [["No filters applied", fonts.chipField]]));
      c.alignment = { vertical: "middle", indent: 1 };
      filterRows.push(row);
      heights[row++] = 21;
    } else {
      const label = ws.getCell(row, B);
      label.value = "FILTERS";
      label.font = fonts.filterLabel;
      label.alignment = { vertical: "bottom", indent: 1 };
      filterRows.push(row);
      heights[row++] = 16;
      for (const line of lines) {
        const c = ws.getCell(row, B);
        c.value = rich(...chipRuns(line, fonts));
        c.alignment = { vertical: "middle", indent: 1 };
        filterRows.push(row);
        heights[row++] = 20;
      }
    }
  }
  heights[row++] = 18;                           // spacer
  const layoutStart = row;

  // 2. Visuals
  let layout;
  let emptyRow = null;
  if (data.items.length) {
    layout = writePageLayout(ws, data.items, layoutStart, widths, options, lay, nameTable);
  } else {
    emptyRow = layoutStart;
    const c = ws.getCell(emptyRow, B);
    c.value = "No visual on this page could be exported.";
    c.font = fonts.empty;
    c.alignment = { vertical: "middle", indent: 1 };
    layout = { nextRow: layoutStart + 1, lastCol: B + 7, tableCount: 0, grouped: false, pendingNotes: [] };
  }

  // 3. Column widths; the page is at least as wide as its masthead text
  const titleUnits = Math.ceil(String(page.displayName).length * (branding.titleFont === "Georgia" ? 2.3 : 2.1)) + 6;
  const minPage = Math.max(MIN_PAGE_W, titleUnits, Math.ceil(eyebrowText.length * 0.95) + 6,
    Math.ceil(metaText.length * 0.9) + 6, Math.ceil(longestFilterLine * 0.9) + 6);
  const { contentEnd, bandEnd } = widths.apply(ws, Math.max(layout.lastCol, B), minPage);

  // 4. Rules and fills, row colour tabs
  paintRange(ws, 1, 1, bandEnd, { fill: solidFill(accent.rule) });
  paintRowTabs(ws, lay.pendingTabs, fonts);
  paintRange(ws, metaRow, B, contentEnd, { border: { bottom: INK_RULE } });
  for (const fr of filterRows) paintRange(ws, fr, B, contentEnd, { fill: SURFACE_FILL });
  if (emptyRow) paintRange(ws, emptyRow, B, contentEnd, { fill: SURFACE_FILL });

  // 5. Logo: 32px high, at most 200px wide, anchored by native offsets at B
  if (logo) {
    const aspect = logo.width && logo.height ? logo.width / logo.height : 1;
    let height = 32, width = Math.round(32 * aspect);
    if (width > 200) { width = 200; height = Math.round(200 / aspect); }
    ws.addImage(logo.id, {
      tl: { nativeCol: B - 1, nativeColOff: 0, nativeRow: logoRow - 1, nativeRowOff: 6 * 9525 },
      ext: { width, height },
      editAs: "oneCell",
    });
  }

  // 6. Notes under collapsed tables, worded to fit the table's width
  for (const note of layout.pendingNotes) {
    let units = 0;
    for (let c = note.col; c < note.col + note.w; c++) units += ws.getColumn(c).width || GUTTER_W;
    const rest = units >= 62 ? `  ·  ${note.hidden} grouped under [+] / [–] in the left margin  ·  expand before sorting or filtering`
      : units >= 44 ? `  ·  ${note.hidden} grouped under [+] / [–]  ·  expand before sorting`
        : `  ·  ${note.hidden} under [+]`;
    const c = ws.getCell(note.row, note.col);
    c.value = rich([`${note.n} rows`, fonts.noteCount], [rest, fonts.noteText]);
    c.alignment = { vertical: "middle", indent: 1 };
  }

  // 7. Notes block and footer
  if (layout.colorNotes) data.notes.push(...layout.colorNotes);
  row = layout.nextRow;
  if (options.includeNotes && data.notes.length) {
    row += 2;
    const head = ws.getCell(row, B);
    head.value = "NOTES";
    head.font = fonts.notesHead;
    head.alignment = { vertical: "bottom" };
    paintRange(ws, row, B, contentEnd, { border: { top: LIGHT_RULE } });
    row++;
    data.notes.forEach((text, i) => {
      const c = ws.getCell(row++, B);
      c.value = rich([String(i + 1), fonts.noteNum], ["   " + text, fonts.noteBody]);
      c.alignment = { vertical: "middle" };
    });
  }
  const footerRow = row + 1;
  paintRange(ws, footerRow, B, contentEnd, { border: { top: LIGHT_RULE } });
  if (branding.footer) {
    const left = ws.getCell(footerRow, B);
    left.value = branding.footer;
    left.font = fonts.footerLeft;
    left.alignment = { vertical: "middle" };
  }
  if (contentEnd > B) {
    const right = ws.getCell(footerRow, contentEnd);
    right.value = `Source: Power BI  ·  ${currentReportName}`;
    right.font = fonts.footerRight;
    right.alignment = { horizontal: "right", vertical: "middle" };
  }

  // 8. Row heights, outline, print setup
  for (let r = 1; r < layoutStart; r++) ws.getRow(r).height = heights[r] ?? ROW_H;
  for (let r = layoutStart; r <= footerRow; r++) ws.getRow(r).height = ROW_H;
  // outlineProperties is never set: Excel's default puts [+] on the note row, and
  // setting it together with fitToPage makes ExcelJS write invalid sheet XML
  if (layout.grouped) ws.properties.outlineLevelRow = 1;

  const letterPaper = /^(en-US|en-CA|fr-CA|es-MX|es-US|en-PH)/i.test((typeof navigator !== "undefined" && navigator.language) || "");
  Object.assign(ws.pageSetup, {
    paperSize: letterPaper ? 1 : 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.55, bottom: 0.55, header: 0.25, footer: 0.25 },
    printArea: `A1:${ws.getColumn(bandEnd).letter}${footerRow}`,
  });
  const header = `&L&"Calibri,Bold"&9&K0E1A2B${hfText(currentReportName)} – ${hfText(page.displayName)}` +
    `&R&"Calibri,Regular"&9&K5B6B7F${hfText(branding.companyName)}`;
  const footer = `&L&"Calibri,Regular"&9&K5B6B7F${hfText(branding.footer || branding.companyName)}` +
    `&C&"Calibri,Regular"&9&K5B6B7FPage &P of &N&R&"Calibri,Regular"&9&K5B6B7FSource: Power BI`;
  ws.headerFooter = { differentFirst: true, oddHeader: header, oddFooter: footer, firstFooter: footer };

  return { tables: layout.tableCount, kpis: data.items.length - layout.tableCount };
}

/* ── Orchestrator ─────────────────────────────────────────────────────── */
/*
 * pages: Power BI Page objects to export, in order.
 * onProgress({ pageIndex, pageCount, pageName, visualIndex, visualCount, visualTitle })
 * Returns { fileName, sheets, tables, kpis, notes } or throws { cancelled: true }.
 */
async function exportToExcel(pages, options, branding, onProgress, isCancelled) {
  const original = await embeddedReport.getActivePage();
  const collected = [];

  // Chart colours: the theme from the embed API now, and the report file (exact theme and
  // colours set by hand) downloaded while the pages are read
  let theme = null, reportColorsJob = null;
  if (options.matchColors) {
    theme = await readThemeColors();
    options = { ...options, palette: theme.colors };
    reportColorsJob = typeof loadReportColors === "function"
      ? loadReportColors(currentReport, theme.fromReport ? { dataColors: theme.colors } : null)
      : Promise.reject(new Error("NOT_LOADED"));
    reportColorsJob.catch(() => {});   // handled below, after the pages
  }
  try {
    for (let p = 0; p < pages.length; p++) {
      const page = pages[p];
      onProgress({ pageIndex: p + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: 0, visualCount: 0, visualTitle: "Opening page…" });
      await activatePage(page);
      const data = await collectPage(page, options,
        (i, n, title) => onProgress({ pageIndex: p + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: i, visualCount: n, visualTitle: title }),
        isCancelled);
      collected.push({ page, data });
    }
  } finally {
    // Put the user back on the page they were looking at
    if (pages.length > 1 || pages[0].name !== original.name) {
      await activatePage(original).catch(() => {});
    }
  }

  if (reportColorsJob) {
    onProgress({ pageIndex: pages.length, pageCount: pages.length, pageName: pages[pages.length - 1].displayName, visualIndex: 0, visualCount: 0, visualTitle: "Reading the report's colours…" });
    let colorNote = null, timer = null, poll = null;
    try {
      const stop = new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error("TIMEOUT")), REPORT_COLORS_TIMEOUT_MS);
        poll = setInterval(() => { if (isCancelled()) reject(new Error("CANCELLED")); }, 250);
      });
      const rc = await Promise.race([reportColorsJob, stop]);
      options = { ...options, palette: rc.palette, reportColors: rc };
    } catch (e) {
      const code = String((e && e.message) || "");
      if ((code === "TIMEOUT" || code === "CANCELLED") && reportColorsJob.abort) reportColorsJob.abort();
      if (code === "CANCELLED") throw { cancelled: true };
      log(`Exact report colours not used: ${errorText(e)}`);
      colorNote = theme.fromReport ? reportColorsMessage(e)
        : "The report's colour theme could not be read, so chart colours use Power BI's default colours, which may differ from this report.";
    } finally {
      clearTimeout(timer);
      clearInterval(poll);
    }
    if (isCancelled()) throw { cancelled: true };
    if (colorNote && collected.length) collected[0].data.notes.push(colorNote);
  }

  const exportTime = new Date();
  const dateText = exportTime.toLocaleString(undefined, { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });

  // White-label workbook properties: the customer's name only, never the product's
  const wb = new ExcelJS.Workbook();
  if (branding.companyName) {
    wb.creator = branding.companyName;
    wb.lastModifiedBy = branding.companyName;
    wb.company = branding.companyName;
  }
  wb.title = currentReportName;
  wb.subject = "Power BI report export";
  wb.keywords = "Power BI; export";
  wb.category = "Report";
  wb.description = `Exported from Power BI on ${dateText}`;
  wb.created = exportTime;
  wb.modified = exportTime;

  let logo = null;
  if (branding.logo) {
    const extension = /^data:image\/png/i.test(branding.logo) ? "png" : "jpeg";
    logo = { id: wb.addImage({ base64: branding.logo, extension }), width: branding.logoWidth, height: branding.logoHeight };
  }

  const nameSheet = makeSheetNamer();
  const nameTable = makeTableNamer();
  const totals = { tables: 0, kpis: 0, notes: [] };
  collected.forEach(({ page, data }, i) => {
    const ctx = { pageIndex: i + 1, pageCount: collected.length, dateText };
    const counts = writePageSheet(wb, nameSheet(page.displayName), page, data, options, branding, logo, nameTable, ctx);
    totals.tables += counts.tables;
    totals.kpis += counts.kpis;
    data.notes.forEach((n) => totals.notes.push(collected.length > 1 ? `${page.displayName}: ${n}` : n));
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const baseName = collected.length === 1 ? `${currentReportName} - ${collected[0].page.displayName}` : currentReportName;
  const fileName = (baseName.replace(/[\\/:*?"<>|]/g, "").trim() || "Power BI export") + ".xlsx";
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(link.href), 10000);

  log(`Saved "${fileName}" – ${collected.length} sheet(s), ${totals.tables} table(s), ${totals.kpis} KPI(s), ${totals.notes.length} note(s)`);
  return { fileName, sheets: collected.length, tables: totals.tables, kpis: totals.kpis, notes: totals.notes };
}
