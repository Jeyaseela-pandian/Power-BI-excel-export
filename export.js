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
 *  3. Chart tables show the chart's exact colours where the chart's legend would: on the
 *     cell naming the series (each row's legend value, slice or tile; the header of a
 *     measure drawn as a series), with a colour key in the caption. Colours that depend
 *     on the value (colour scales, rules, waterfall steps) stay on the value cells.
 *     Excel moves the fills with the rows when sorting.
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
const OVERLAP_SHARE = 0.25;    // …nor overlap under a quarter of the smaller visual (padding, not placement)
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
const NEUTRAL_ACCENT = "FF1F3A5F";   // deep navy: the "match the report" accent when no theme can be read
const MIN_RULE_CONTRAST = 2;         // accent rules and caps stay at least this visible on white
const ROW_H = 18;          // the row grid, in points
const HEAD_ROWS = 3;       // exhibit title, caption, table header
const TREND_ROWS = 4;      // a trend card: label, value (two rows) beside its sparkline, change
const CONTENT_COL = 2;     // everything starts at column B; column A is the margin
const MARGIN_W = 3;
const GUTTER_W = 3;
const MIN_PAGE_W = 100;    // minimum content width (Excel width units)
const SNAP_PX = 40;        // stacks starting within this many page pixels share a column
const BAND_GAP = 2;        // blank rows between bands
// Matrices: normal Excel tables by default; options.matrixLayout === "grid" gives Power BI's
// stepped layout (groups and subtotals, a column per column-field value)
const flatMatrix = (item) => item.type === "pivotTable" && !item.matrixGrid;
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
// ("Order Count" and "Order Qty" are amounts; "Sort Order" is a helper)
const HELPER_COLUMN = /\b(sort|rank|index|id|key)\b|\border\b(?!\s*(count|qty|quantity|total|value|amount|lines?|date|no\b|number))/i;
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

/* A colour darkened step by step until it shows clearly on white: pale yellow becomes a
   deep gold, not black, so the colour stays recognisable */
function readableOnWhite(argb, min = MIN_RULE_CONTRAST) {
  let n = parseInt(argb.slice(2), 16);
  for (let i = 0; i < 12 && contrast("FF" + n.toString(16).padStart(6, "0").toUpperCase(), C.WHITE) < min; i++) {
    n = (Math.round(((n >> 16) & 255) * 0.88) << 16) | (Math.round(((n >> 8) & 255) * 0.88) << 8) | Math.round((n & 255) * 0.88);
  }
  return "FF" + n.toString(16).padStart(6, "0").toUpperCase();
}

/* Where the customer accent goes: the top rule and KPI caps (a darker shade when the accent
   is too pale to see on white), and the sheet tab (the accent itself) */
function accentPlan(argb) {
  const strong = readableOnWhite(argb);
  return { rule: strong, cap: strong, tab: argb };
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
    kpiChange: f(10, C.INK, { bold: true }),
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
  keyDriversVisual: "Key influencers", aiNarratives: "Smart narrative", qnaVisual: "Q&A",
};

function friendlyType(type) {
  const t = String(type || "");
  if (FRIENDLY_TYPES[t]) return FRIENDLY_TYPES[t];
  if (customVisual(t)) return customVisual(t).name;
  if (/^deneb/i.test(t)) return "Deneb visual";
  if (isCustomType(t)) {
    // "SankeyDiagram1446463184954" → "Sankey diagram"; an id alone stays "Custom visual"
    const name = t.replace(/(_[0-9A-F]{8}_.*|[0-9A-F]{32}|\d{12,})$/i, "").replace(/_/g, " ").replace(/([a-z])([A-Z])/g, "$1 $2").trim().toLowerCase();
    return name.length > 2 ? name[0].toUpperCase() + name.slice(1) : "Custom visual";
  }
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
  const { header, filled, numeric, helper, hex } = classifyColumns(item.rows);
  if (TABLE_VISUAL_TYPES.includes(item.type)) {
    const shown = header.slice(0, 3).join(", ");
    return header.length > 3 ? `${shown} +${header.length - 3}` : shown;
  }
  // Measures can come first in the export (a card split by category, an area chart);
  // whole years (2024) are an axis, not a series
  const yearLike = (ci) => numeric[ci] && filled(ci).every((x) => /^(19|20)\d\d$/.test(x));
  const series = header.filter((_, ci) => numeric[ci] && !yearLike(ci) && !helper[ci] && !hex[ci]);
  const texts = header.filter((_, ci) => (!numeric[ci] || yearLike(ci)) && !hex[ci]);
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
 * Legend values in the order Power BI gives them colours (and stacks and lists them). The
 * export's own order (seriesOrder) is Power BI's series order, except when the visual sorts by
 * a measure: then each axis group's rows run largest first (or smallest first), which says
 * nothing about the series order, and Power BI's own order applies: the legend field's values
 * ascending (numbers by value, dates and month names by date, text alphabetically).
 */
function legendOrder(data, keyCol, groupCols, valueCol) {
  const order = seriesOrder(data, keyCol, groupCols);
  if (valueCol === undefined || valueCol < 0 || order.length < 2) return order;
  const num = (r) => { const v = toCellValue(r[valueCol]).value; return typeof v === "number" ? v : null; };
  const groups = new Map();
  for (const r of data) {
    if (!String(r[keyCol] ?? "").trim() || num(r) === null) continue;
    const g = groupCols.map((c) => r[c]).join("\u0001");
    if (!groups.has(g)) groups.set(g, []);
    groups.get(g).push(num(r));
  }
  const seqs = [...groups.values()].filter((s) => s.length > 1);
  const runs = (s, dir) => s.every((v, i) => i === 0 || (dir > 0 ? v >= s[i - 1] : v <= s[i - 1]));
  const strict = seqs.some((s) => s.some((v, i) => i > 0 && v !== s[i - 1]));
  const byValue = seqs.length > 0 && strict && (seqs.every((s) => runs(s, -1)) || seqs.every((s) => runs(s, 1)));
  if (!byValue) return order;
  // Ascending, as Power BI sorts a legend field
  const plain = (v) => v.replace(/,/g, "");
  const month = (v) => { const t = Date.parse(`${v} 1, 2000`); return /\d/.test(v) || isNaN(t) ? null : new Date(t).getMonth(); };
  const sortKey = order.every((v) => /^-?\d+(\.\d+)?$/.test(plain(v))) ? (v) => Number(plain(v))
    : order.every((v) => parseDate(v)) ? (v) => parseDate(v).date.getTime()
      : order.every((v) => month(v) !== null) ? month : null;
  return [...order].sort(sortKey ? (a, b) => sortKey(a) - sortKey(b)
    : (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
}

/* "≥ 0", "> 80% of range", "= Bikes": a rule case for caption legends */
function describeTest(c) {
  if (c.and) return c.and.map(describeTest).join(" and ");
  if (c.or) return c.or.map(describeTest).join(" or ");
  if (c.not) return "not " + describeTest(c.not);
  const ops = ["=", ">", "≥", "≤", "<"];
  const at = c.at.num !== undefined ? c.at.num.toLocaleString() : c.at.pct !== undefined ? `${Math.round(c.at.pct * 100)}% of range`
    : c.at.pctl !== undefined ? `${Math.round(c.at.pctl * 100)}th percentile` : c.at.str;
  return `${ops[c.op] || "?"} ${at}`;
}

// Measures whose group value isn't the sum of its rows
const NON_ADDITIVE = /%|\b(avg|average|mean|median|rate|ratio|margin|percent|pct|distinct|min|max|count of)\b/i;

/*
 * Applies the report's fx colour rules (colors.js) to a chart table, row by row.
 * A rule reads a field: a column of this table; else a column whose name holds all the
 * field's words ("Profit %" → "Profit Margin %"); else, for a scale per legend value, the
 * legend order (a scale over e.g. "Latest year" follows the years). Those two are
 * approximations and add a note; a rule whose field can't be found is left out.
 * A category hierarchy (Group, then Region) is drawn at its first level, so an additive
 * field is totalled per group first, as the bars show it.
 * Rules for one measure win over rules for every series; series without a rule keep
 * their series colour, and rows a rule leaves uncoloured keep it too.
 * Returns colours for columnColors: perCol (a colour per row for each ruled column), the
 * remaining series, the row colours by key column and a caption legend.
 */
function colorRuleRows(rules, t) {
  const { header, data, cols, norm, lower, numeric, hex, seriesCols, findCol, info, seriesColor } = t;
  const n = data.length;
  if (!n) return null;
  const legendCol = info.series ? findCol(info.series) : -1;
  const categoryCol = findCol(info.category);
  const levels = (info.categories || []).map((f) => findCol(f)).filter((ci) => ci >= 0);
  const grouped = levels.length > 1 && levels[0] === categoryCol;
  const approx = [];    // rule fields matched to a similar column, or approximated
  const skipped = [];   // rule fields not in the exported data at all
  // A rule's field: its column; else a column whose name holds all its words ("Profit %" →
  // "Profit Margin %"); else one with the same words apart from little ones ("Var Plan %" →
  // "Var to Plan %"). The last two are approximations and add a note.
  const LITTLE = new Set(["to", "of", "the", "a", "an", "by", "vs", "and", "in", "for"]);
  const wordKey = (s) => s.split(/\s+/).filter((w) => w && !LITTLE.has(w)).sort().join(" ");
  const columnFor = (name) => {
    const low = String(name).toLowerCase();
    let ci = findCol({ ref: normalizeField(low) });
    if (ci < 0) {
      const words = low.split(/\s+/).filter(Boolean);
      let close = cols.filter((c) => words.every((w) => lower[c].split(/\s+/).includes(w)));
      if (close.length !== 1) close = cols.filter((c) => wordKey(lower[c]) === wordKey(low));
      if (close.length === 1) { ci = close[0]; if (!approx.includes(name)) approx.push(name); }
    }
    return ci;
  };
  const column = (scope) => (name) => {
    const ci = columnFor(name);
    if (ci < 0) return null;
    const raw = data.map((r) => String(r[ci] ?? "").trim());
    let v = raw.map((x) => (x ? toCellValue(x).value : null));
    const additive = !NON_ADDITIVE.test(header[ci]) && !raw.some((x) => x.endsWith("%"));
    // A rule per legend value reads the field once per legend value (its total), so every
    // row of one legend value shares a colour; for a ratio the average stands in (a note)
    if (scope === "series" && legendCol >= 0 && numeric[ci]) {
      const sum = {}, count = {};
      const keyOf = (r) => String(r[legendCol] ?? "").trim();
      data.forEach((r, ri) => { const k = keyOf(r); sum[k] = (sum[k] || 0) + (typeof v[ri] === "number" ? v[ri] : 0); count[k] = (count[k] || 0) + 1; });
      if (!additive && !approx.includes(name)) approx.push(name);
      return data.map((r) => (additive ? sum[keyOf(r)] : sum[keyOf(r)] / count[keyOf(r)]));
    }
    if (grouped && scope !== "series" && numeric[ci] && additive) {
      const totals = {};
      const groupOf = (r) => String(r[categoryCol] ?? "").trim();
      data.forEach((r, ri) => { totals[groupOf(r)] = (totals[groupOf(r)] || 0) + (typeof v[ri] === "number" ? v[ri] : 0); });
      v = data.map((r) => totals[groupOf(r)]);
    }
    return v;
  };

  const perCol = {};
  let first = null;
  // A rule for one measure wins over a rule for every series; otherwise the first rule wins
  // (an Azure Map's bubble colour is listed before its leftover data point colour)
  const ordered = [...rules].sort((a, b) => Number(b.scope.startsWith("m:")) - Number(a.scope.startsWith("m:")));
  for (const r of ordered) {
    // A map whose measures are all tooltips colours by a rule on one of them: that column
    // shows the colours (Power BI paints the regions with it)
    const inputCol = !seriesCols.length && r.rule.input ? columnFor(r.rule.input) : -1;
    const targets = (r.scope.startsWith("m:") ? seriesCols.filter((ci) => "m:" + norm[ci] === r.scope)
      : seriesCols.length ? seriesCols : inputCol >= 0 && numeric[inputCol] ? [inputCol] : [])
      .filter((ci) => !perCol[ci]);
    if (!targets.length) {
      // nothing to colour because the rule's field isn't exported either: say so
      if (!seriesCols.length && inputCol < 0 && r.rule.input && !skipped.includes(r.rule.input)) skipped.push(r.rule.input);
      continue;
    }
    let { colors, missing } = evaluateColorRule(r.rule, column(r.scope), n);
    if (missing.length && r.rule.kind === "gradient" && r.scope === "series" && legendCol >= 0) {
      const order = legendOrder(data, legendCol, cols.filter((ci) => ci !== legendCol && !seriesCols.includes(ci) && !hex[ci]), seriesCols[0]);
      ({ colors } = evaluateColorRule(r.rule, () => data.map((row) => order.indexOf(String(row[legendCol] ?? "").trim())), n));
      approx.push(r.rule.input);
    } else if (missing.length) {
      missing.forEach((m) => { if (!skipped.includes(m)) skipped.push(m); });
      continue;
    }
    for (const ci of targets) perCol[ci] = colors.map((c) => (c ? toArgb(c) : null));
    if (!first) first = { rule: r.rule, scope: r.scope, col: targets[0] };
  }
  const skippedNote = skipped.length
    ? `"${t.title || "A chart"}" is coloured by ${skipped.join(", ")} in the report, which isn't in the exported data, so its default colours are shown.`
    : null;
  if (!first) return skippedNote ? { none: true, note: skippedNote } : null;
  for (const ci of Object.keys(perCol)) perCol[ci] = perCol[ci].map((c) => c || seriesColor(Number(ci)));
  const series = {};
  for (const ci of seriesCols) if (!perCol[ci]) series[ci] = seriesColor(ci);

  // Row colours by key column (caption and cell fills)
  const keyCol = first.scope === "series" && legendCol >= 0 ? legendCol
    : categoryCol >= 0 ? categoryCol : cols.find((ci) => !numeric[ci] && !hex[ci]) ?? -1;
  const rowColors = perCol[first.col];
  let dot = null;
  if (keyCol >= 0) {
    const byKey = {};
    data.forEach((r, ri) => { byKey[String(r[keyCol] ?? "").trim()] = rowColors[ri]; });
    dot = { keyCol, colorOf: (v) => byKey[v] || null };
  }

  const rule = first.rule;
  const ci0 = columnFor(rule.input);
  const inputName = ci0 >= 0 ? header[ci0] : rule.input;
  const legend = rule.kind === "gradient"
    ? { scale: [rule.min.color, rule.mid && rule.mid.color, rule.max.color].filter(Boolean).map((c) => toArgb(c)), label: `Colour scale on ${inputName}` }
    : rule.kind === "conditional"
      ? { entries: rule.cases.map((c) => ({ label: `${inputName} ${describeTest(c.test)}`, argb: toArgb(c.color) })) }
      : { text: `Colours from ${inputName}` };
  const note = [approx.length
    ? `Colours of "${t.title || "a chart"}" follow a rule on ${approx.join(", ")}, which isn't in the exported data as such, so they are approximated.`
    : null, skippedNote].filter(Boolean).join(" ") || null;
  return { series, rowCols: Object.keys(perCol).map(Number), perCol, rowColors, rowKey: null, dot, legend, note };
}

/* The Nested Bars visual's automatic subset colour, as its code works it out with d3: the
   hue turned half way round, lightness moved 0.2 towards the middle, saturation raised 20% */
function complementaryColor(hex) {
  const n = parseInt(String(hex).replace("#", ""), 16);
  if (!/^#?[0-9A-F]{6}$/i.test(String(hex))) return "#FD625E";
  const r = (n >> 16 & 255) / 255, g = (n >> 8 & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = NaN, s = max - min, l = (max + min) / 2;
  if (s) {
    h = r === max ? (g - b) / s + (g < b) * 6 : g === max ? (b - r) / s + 2 : (r - g) / s + 4;
    s /= l < 0.5 ? max + min : 2 - max - min;
    h *= 60;
  } else s = l > 0 && l < 1 ? 0 : h;
  h = (h + 180) % 360;
  l = l > 0.6 ? Math.max(0.3, l - 0.2) : Math.min(0.7, l + 0.2);
  s = Math.min(0.9, 1.2 * s);
  // back to RGB (d3.hsl → rgb)
  h = h % 360 + (h < 0) * 360;
  const sat = isNaN(h) || isNaN(s) ? 0 : s;
  const m2 = l + (l < 0.5 ? l : 1 - l) * sat, m1 = 2 * l - m2;
  const ch = (x) => (x < 60 ? m1 + (m2 - m1) * x / 60 : x < 180 ? m2 : x < 240 ? m1 + (m2 - m1) * (240 - x) / 60 : m1) * 255;
  const hx = (v) => Math.max(0, Math.min(255, Math.round(v) || 0)).toString(16).padStart(2, "0");
  return ("#" + hx(ch(h >= 240 ? h - 240 : h + 120)) + hx(ch(h)) + hx(ch(h < 120 ? h + 240 : h - 120))).toUpperCase();
}

/* d3's ordinal scale, as custom visuals use it: a value it doesn't know joins its list and
   takes the next colour round */
function ordinalScale(domain, range) {
  const index = new Map();
  for (const d of domain) if (!index.has(d)) index.set(d, index.size);
  return (d) => {
    if (!index.has(d)) index.set(d, index.size);
    return range[index.get(d) % range.length];
  };
}

/*
 * A double donut chart (NeenOpal's Double Donut Chart), coloured as its code does it: the
 * inner ring shows the categories, the outer ring their subcategories.
 *  - Each category takes the next theme colour in the order the data lists them, unless a
 *    colour is set for it.
 *  - The visual stores a category's colour under its value as text ("null" for a blank, a
 *    number or date unformatted) but looks it up by its label ("(Blank)", the formatted
 *    number or date). Where the two differ, the inner slice takes the colour of the 1st,
 *    2nd … category, and the outer ring asks the theme for a new colour.
 *  - An outer slice is its category's colour mixed with white: 30% for the first
 *    subcategory, evenly up to 80% for the last.
 * Without the report file, the category is the text column the rows are grouped by (the
 * one with the fewest values), the subcategory the other. Returns null without a category.
 */
function ringColors(t) {
  const { data, cols, numeric, hex, findCol, palette, info } = t;
  const text = (row, ci) => String(row[ci] ?? "").trim();
  let catCol = info && info.categories.length ? findCol(info.categories[0]) : -1;
  let subCol = catCol >= 0 && info.categories[1] ? findCol(info.categories[1], { [catCol]: true }) : -1;
  if (catCol < 0) {
    const yearLike = (ci) => data.every((row) => !text(row, ci) || /^(19|20)\d\d$/.test(text(row, ci)));
    const textCols = cols.filter((ci) => !hex[ci] && (!numeric[ci] || yearLike(ci)));
    const distinct = (ci) => new Set(data.map((row) => text(row, ci))).size;
    const runs = (ci) => data.filter((row, ri) => !ri || text(row, ci) !== text(data[ri - 1], ci)).length;
    const grouped = textCols.filter((ci) => runs(ci) === distinct(ci));
    catCol = [...(grouped.length ? grouped : textCols)].sort((a, b) => distinct(a) - distinct(b))[0] ?? -1;
    subCol = textCols.find((ci) => ci !== catCol) ?? -1;
  }
  if (catCol < 0) return null;

  // A value's label, and the text the visual stores its colour under
  const label = (v) => (v === "" ? "(Blank)" : v);
  const asTyped = (v) => {   // raw and formatted text agree: plain text, unformatted numbers
    if (v === "" || /^(true|false)$/i.test(v) || /^\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}/.test(v)) return false;
    const n = toCellValue(v).value;
    return typeof n !== "number" || String(n) === v;
  };
  const key = (v) => (v === "" ? "null" : asTyped(v) ? v : "\u0001" + v);
  // Theme colours, handed out in the order they are asked for (the visual's colour palette)
  const handed = new Map();
  const themeColor = (k) => {
    if (!handed.has(k)) handed.set(k, palette[handed.size % palette.length]);
    return handed.get(k);
  };
  const overrides = (info && info.overrides) || {};
  const cats = [...new Set(data.map((row) => text(row, catCol)))];
  const colorOf = new Map();
  for (const v of cats) {
    const theme = themeColor(key(v));   // asked for even when a colour is set
    colorOf.set(key(v), overrides["i:" + (v === "" ? "null" : normalizeValue(v))] || theme);
  }
  const inner = ordinalScale(cats.map(key), cats.map((v) => colorOf.get(key(v))));
  const innerColor = Object.fromEntries(cats.map((v) => [v, toArgb(inner(label(v)))]));

  let rings = null;
  if (subCol >= 0) {
    const labels = data.map((row) => `${label(text(row, catCol))} - ${label(text(row, subCol))}`);
    // Slices are grouped by their label's text before the first " - ", as the visual splits it
    const groups = new Map();
    for (const l of labels) {
      const [parent, child] = l.split(" - ");
      if (!groups.has(parent)) groups.set(parent, []);
      groups.get(parent).push(child);
    }
    const shade = new Map();
    for (const [parent, kids] of groups) {
      const base = colorOf.get(parent) || themeColor(parent);
      kids.forEach((child, k) => shade.set(`${parent} - ${child}`, mixHex(base, "#FFFFFF", 0.3 + k / Math.max(kids.length - 1, 1) * 0.5)));
    }
    const outer = ordinalScale(labels, labels.map((l) => shade.get(l)));
    rings = { col: subCol, colors: labels.map((l) => toArgb(outer(l) || "#000000")) };   // a slice it finds no colour for is black
  }
  return {
    series: {}, rowCols: [], rowKey: { col: catCol, colorOf: innerColor }, rings,
    legend: { entries: cats.map((v) => ({ label: label(v), argb: innerColor[v] })) },
    dot: { keyCol: catCol, colorOf: (v) => innerColor[v] || null },
  };
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
 * Supported AppSource visuals (colors.js CUSTOM_VISUALS) are coloured as their code does
 * it: a double donut by ringColors.
 * Series take the theme's data colours in field order, legend values in series order
 * (seriesOrder) and pie / treemap categories in the order they appear, as Power BI does.
 * Returns { series, rowCols, hexCols, rowKey, rings, dot, exact }:
 *  - series: { col: argb } columns drawn in one colour (header band, filled cells)
 *  - rowCols: columns whose cells take their row's colour (rows coloured by key or hex)
 *  - rowKey: { col, colorOf: { value: argb } } legend / category colours, for the caption
 *  - rings: { col, colors: [argb per row] } a double donut's outer ring (subcategory cells)
 *  - dot: { keyCol, colorOf(value) } the column that decides each row's colour
 *  - exact: true when the report file described this visual
 */
function columnColors(item, palette, info) {
  if (!palette || !palette.length || item.isKpi || TABLE_VISUAL_TYPES.includes(item.type)) return null;
  const own = customVisual(item.type);
  if (own && own.palette) palette = own.palette;   // a visual with its own fixed colours instead of the theme's
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

  // Supported AppSource visuals, coloured as their code does it
  const custom = customVisual(item.type);
  if (custom && custom.rule === "table") return null;   // its own formatting, as a table's (data bars)
  if (custom && custom.rule === "rings") {
    const rings = ringColors({ data, cols, numeric, hex, findCol, palette, info });
    if (rings) return { ...rings, hexCols, exact: !!info };
  }
  if (custom && custom.rule === "nodes") {
    // Sankey: every node (source or destination) its theme colour by name, in the order the
    // visual meets them: sources first, then destination-only nodes. Its code asks the palette
    // for an empty key right after the first node, so node k (k ≥ 1) takes slot k + 1
    const texts = cols.filter((ci) => !numeric[ci] && !hex[ci]);
    const src = info ? findCol(info.category) : texts[0] ?? -1;
    const dst = info ? findCol((info.roles.Destination || [])[0], { [src]: true }) : texts.find((ci) => ci !== src) ?? -1;
    if (src >= 0 && dst >= 0) {
      const text = (r, ci) => String(r[ci] ?? "").trim();
      const sources = [...new Set(data.map((r) => text(r, src)))];
      const nodes = [...sources, ...[...new Set(data.map((r) => text(r, dst)))].filter((v) => !sources.includes(v))];
      const colorOf = {};
      const set = (info && info.overrides) || {};   // node colours set in the report, or one for every node
      nodes.forEach((v, k) => { colorOf[v] = toArgb(set["i:" + normalizeValue(v)] || set.all || palette[(k === 0 ? 0 : k + 1) % palette.length]); });
      const rowKey = { col: src, colorOf: Object.fromEntries(sources.map((v) => [v, colorOf[v]])) };
      return { series: {}, rowCols: [], hexCols, rowKey, rings: { col: dst, colors: data.map((r) => colorOf[text(r, dst)]) },
        dot: { keyCol: src, colorOf: (v) => colorOf[v] || null }, exact: !!info };
    }
  }
  if (custom && custom.rule === "fixed" && info) {
    // Fixed colours per value field (a bullet chart's bands…): each field's header, from the
    // report's setting, else the visual's default; a setting may fall back to another's
    const series = {};
    for (const [role, keys] of Object.entries(custom.seriesColors || {})) {
      const ci = findCol((info.roles[role] || [])[0], series);
      const c = keys.map((k) => customSetting(item.type, info, k)).find(Boolean);
      if (ci >= 0 && c) series[ci] = toArgb(c);
    }
    return Object.keys(series).length ? { series, rowCols: [], hexCols, rowKey: null, dot: null, exact: true } : null;
  }
  if (custom && custom.rule === "nested") {
    // Nested bars: the category value in the outer colour, the subset value inside it in the
    // subset colour (by default the outer colour's complement); the visual's own colours,
    // not the theme's
    const valueCols = cols.filter((ci) => numeric[ci] && !hex[ci]);
    const placed = item.nestedCols;   // rows rebuilt by nestedBarsRows: category, total, chosen subset
    const outerCol = placed ? placed.outer : info ? findCol(info.measures.find((m) => m.role === "Y")) : valueCols[0] ?? -1;
    const innerCol = placed ? placed.inner : info ? findCol(info.measures.find((m) => m.role === "Y2"), { [outerCol]: true }) : valueCols[1] ?? -1;
    const outer = customSetting(item.type, info, "outer");
    const inner = customSetting(item.type, info, "autoInner") === false ? customSetting(item.type, info, "inner") : complementaryColor(outer);
    const series = {};
    if (outerCol >= 0) series[outerCol] = toArgb(outer);
    if (innerCol >= 0) series[innerCol] = toArgb(inner);
    if (Object.keys(series).length) return { series, rowCols: [], hexCols, rowKey: null, dot: null, exact: !!info };
  }

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
  const hasRules = !!(info && (info.rules || []).length);
  if (!seriesCols.length && !hexCols.length && !hasRules) return null;
  const exact = !!info;

  // 1. A hex colour column: every row in its own colour (one band colour when they all agree)
  if (hexCols.length) {
    const colors = new Set(filled(hexCols[0]).map((x) => x.toUpperCase()));
    const dot = { keyCol: hexCols[0], colorOf: (v) => (HEX_COLOR.test(v) ? toArgb(v) : null) };
    if (colors.size === 1) {   // one colour for every row: a series colour
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
  // A bar or column chart of one measure without a legend has one colour in Power BI's format
  // pane, its "Default colour" (set for no field); a colour an earlier version of the chart left
  // for the measure isn't shown
  const barSpec = chartSpecOf(item.type);
  const defaultOnly = barSpec && barSpec.kind === "bar" && info && !info.series && (info.measures || []).length === 1 && overrides.all;
  const seriesColor = (ci, slot = indexOf[ci] ?? 0) => {
    if (defaultOnly) return toArgb(overrides.all);
    const f = fieldOf[ci];
    const set = overrides["m:" + norm[ci]] || (f && (overrides["m:" + f.ref] || (f.label && overrides["m:" + f.label])));
    if (set) return toArgb(set);
    if (overrides.all) return toArgb(overrides.all);
    if (info && info.rule === "single") return color(0);
    return color(slot);
  };
  // A treemap draws (and colours) its groups largest first; the exported rows come in the
  // category's own order, so its groups are put in order of their totals before the colours
  const largestFirst = (keyCol, values) => {
    const ci = seriesCols[0];
    if (item.type !== "treemap" || ci === undefined) return values;
    const total = {};
    data.forEach((r) => {
      const k = (r[keyCol] ?? "").trim(), v = toCellValue(r[ci]).value;
      if (k && typeof v === "number") total[k] = (total[k] || 0) + v;
    });
    return values.map((v, i) => ({ v, i })).sort((a, b) => ((total[b.v] || 0) - (total[a.v] || 0)) || (a.i - b.i)).map((x) => x.v);
  };
  const keyedRows = (keyCol, colorOfValue, bySeries) => {
    const values = bySeries
      ? legendOrder(data, keyCol, cols.filter((ci) => ci !== keyCol && !seriesCols.includes(ci) && !hex[ci]), seriesCols[0])
      : largestFirst(keyCol, [...new Set(filled(keyCol))]);   // order of first appearance (a treemap's largest first)
    return { col: keyCol, colorOf: Object.fromEntries(values.map((v, i) => [v, colorOfValue(v, i)])) };
  };

  // 2. Waterfall: each step in its increase / decrease colour (a "Total" step in the total colour)
  if (info && info.waterfall && seriesCols.length) {
    const w = info.waterfall, ci = seriesCols[0];
    const keyCol = cols.find((c) => !numeric[c] && !hex[c]) ?? -1;
    const rowColors = data.map((r) => {
      const label = keyCol >= 0 ? String(r[keyCol] ?? "").trim() : "";
      const v = toCellValue(r[ci]).value;
      return toArgb(/^total$/i.test(label) ? w.total : typeof v === "number" && v < 0 ? w.decrease : w.increase);
    });
    const legend = { entries: [{ label: "Increase", argb: toArgb(w.increase) }, { label: "Decrease", argb: toArgb(w.decrease) }, { label: "Total", argb: toArgb(w.total) }] };
    return { series: {}, rowCols: [ci], perCol: { [ci]: rowColors }, rowColors, hexCols, rowKey: null, dot: null, legend, exact };
  }

  // An AppSource visual whose colouring isn't known: most give each value of their first
  // category its own theme colour (Power BI's palette, in data order), with colours set per
  // item kept. A best match, so a note says so.
  const unknownCustom = info ? info.rule === "generic" : isCustomType(item.type) && !customVisual(item.type) && !/^deneb/i.test(item.type);
  if (unknownCustom) {
    const keyCol = cols.find((ci) => !numeric[ci] && !hex[ci] && !helper[ci]);
    if (keyCol !== undefined && new Set(filled(keyCol)).size > 1) {
      const rowKey = keyedRows(keyCol, (v, i) => toArgb(itemOverride(v) || palette[i % palette.length]));
      return { series: {}, rowCols: [], hexCols, rowKey, dot: { keyCol, colorOf: (v) => rowKey.colorOf[v] || null }, exact: !!info,
        note: `"${item.displayTitle || item.title}" (${friendlyType(item.type)}) is an AppSource visual whose colouring isn't known in detail: its ${header[keyCol]} values take the theme colours in order, which may differ from the chart.` };
    }
  }

  // 3. Conditional formatting (fx) in the report file: colour scales, rules, colours from a field
  let ruleNote = null;
  // With a legend, Power BI applies colour rules per legend value only; a rule for every
  // data point is a leftover from before the legend was added, and isn't drawn
  const hasLegend = !!(info && info.series && findCol(info.series) >= 0);
  const fxRules = info ? (info.rules || []).filter((r) => (r.scope === "all" && !hasLegend) || r.scope === "series" || seriesCols.some((ci) => "m:" + norm[ci] === r.scope)) : [];
  if (fxRules.length) {
    const result = colorRuleRows(fxRules, { header, data, cols, norm, lower, numeric, hex, seriesCols, findCol, info, seriesColor, title: item.displayTitle || item.title });
    if (result && !result.none) return { ...result, hexCols, exact };
    if (result) ruleNote = result.note;   // the rule couldn't be applied: default colours, with a note
  }

  // An expanded treemap, pie or donut (a tile / slice per Region and State, "East New York"):
  // each in its own theme colour (a treemap's largest first, a pie's in data order), or the
  // colour set by hand for it, as Power BI draws it; in the table, on its last level's cell
  const levels = COLOR_BY_CATEGORY_TYPES.includes(item.type) ? expandedLevelCols(item, info, header) : null;
  if (levels && seriesCols.length) {
    const label = (r) => levelsLabel(r, levels);
    const total = {}, byLabel = {};
    data.forEach((r) => {
      const v = toCellValue(r[seriesCols[0]]).value;
      if (typeof v === "number") total[label(r)] = (total[label(r)] || 0) + v;
      if (!byLabel[label(r)]) byLabel[label(r)] = levels.map((ci) => normalizeValue(String(r[ci] ?? "").trim()));
    });
    const seen = [...new Set(data.map(label))];
    const order = item.type === "treemap" ? seen.sort((a, b) => (total[b] || 0) - (total[a] || 0)) : seen;
    const setFor = (v) => overrides["i:" + byLabel[v].join("|")] || overrides["i:" + [...byLabel[v]].reverse().join("|")];
    const colorOf = Object.fromEntries(order.map((v, i) => [v, toArgb(setFor(v) || overrides.all || palette[i % palette.length])]));
    const rowColors = data.map((r) => colorOf[label(r)]);
    const last = levels[levels.length - 1];
    return { series: {}, rowCols: [last], perCol: { [last]: rowColors }, rowColors, hexCols, rowKey: null, dot: null,
      legend: { entries: order.map((v) => ({ label: v, argb: colorOf[v] })) }, tiles: { levels, colorOf }, exact, note: ruleNote };
  }

  // 4. Rows coloured by a field: the legend (Series) or, for pie / donut / treemap, the category
  let rowKey = null;
  if (info && (info.rule === "legend" || info.rule === "item")) {
    let keyCol = findCol(info.rule === "legend" ? info.series : info.category);
    if (keyCol < 0 && info.rule === "item" && !numeric[0] && !hex[0]) keyCol = 0;
    // One colour for every item: a word cloud's, or a box and whisker chart's "one colour" (its own, else the first)
    const oneColor = (custom && custom.settings && custom.settings.all && overrides.all) ||
      (custom && custom.settings && custom.settings.oneColor && customSetting(item.type, info, "oneColor") === true
        ? customSetting(item.type, info, "oneFill") || palette[0] : null);
    if (keyCol >= 0) rowKey = keyedRows(keyCol, (v, i) => toArgb(itemOverride(v) || oneColor || palette[i % palette.length]), info.rule === "legend");
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
    // A combo chart's legend splits its column values only: line values (Y2) stay one series
    // each, in their own colour (set by hand, else the theme colours after the legend's)
    const lines = info && info.rule === "legend" ? seriesCols.filter((ci) => fieldOf[ci] && fieldOf[ci].role === "Y2") : [];
    const legendCount = Object.keys(rowKey.colorOf).length;
    const series = Object.fromEntries(lines.map((ci, k) => [ci, seriesColor(ci, legendCount + k)]));
    return { series, rowCols: seriesCols.filter((ci) => !lines.includes(ci)), hexCols, rowKey, dot, exact, note: ruleNote };
  }
  if (!seriesCols.length && !ruleNote) return null;
  const series = Object.fromEntries(seriesCols.map((ci) => [ci, seriesColor(ci)]));
  return { series, rowCols: [], hexCols, rowKey: null, dot: null, exact, note: ruleNote };
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
/*
 * A date as Power BI exports it, "2023-01-15 00:00:00" or "1/15/2023 12:00:00 AM" (day
 * first when the first number can't be a month), as a UTC Date (what ExcelJS writes as an
 * Excel date), or null. hasTime: the value carries a time of day other than midnight.
 */
const EXPORTED_ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?Z?$/;
const EXPORTED_SLASH_DATE = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?)?$/i;
function parseDate(text, dayFirst = false) {
  let y, mo, d, h = 0, mi = 0, sec = 0, m;
  if ((m = text.match(EXPORTED_ISO_DATE))) [, y, mo, d, h = 0, mi = 0, sec = 0] = m;
  else if ((m = text.match(EXPORTED_SLASH_DATE))) {
    let a = Number(m[1]), b = Number(m[2]);
    if ((dayFirst || a > 12) && b <= 12) [a, b] = [b, a];   // 15/1/2023 (or a day-first column): day first
    [mo, d, y] = [a, b, m[3]];
    h = Number(m[4] || 0); mi = m[5] || 0; sec = m[6] || 0;
    if (m[7]) h = (h % 12) + (/p/i.test(m[7]) ? 12 : 0);
  } else return null;
  [y, mo, d, h, mi, sec] = [y, mo, d, h, mi, sec].map(Number);
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || sec > 59) return null;
  const date = new Date(Date.UTC(y, mo - 1, d, h, mi, sec));
  if (date.getUTCDate() !== d) return null;   // 31 February
  return { date, hasTime: !!(h || mi || sec) };
}
const SHORT_DATE = "mm-dd-yy";   // Excel's built-in short date: shown in the reader's own date order

/* Columns whose dates are written day first (dd/mm/yyyy): some value's first number can't be a month */
function dayFirstColumns(header, data) {
  return header.map((_, ci) => data.some((r) => {
    const m = String(r[ci] ?? "").trim().match(EXPORTED_SLASH_DATE);
    return m && Number(m[1]) > 12;
  }));
}

function toCellValue(raw, dayFirst = false) {
  const s = (raw ?? "").trim();
  if (s === "") return { value: null };
  if (/^0\d+$/.test(s)) return { value: s };             // codes with leading zeros (ZIP, IDs)
  const date = parseDate(s, dayFirst);
  if (date) return { value: date.date, numFmt: date.hasTime ? "m/d/yyyy h:mm" : SHORT_DATE };

  let t = s, negative = false;
  if (/^\(.*\)$/.test(t)) { negative = true; t = t.slice(1, -1); }
  const isPercent = t.endsWith("%");
  const currency = (t.match(/^-?\s*([$€£₹¥])/) || [])[1];
  const cleaned = t.replace(/[$€£₹¥%,\s]/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return { value: s };

  const n = Number(cleaned) * (negative ? -1 : 1);
  const decimals = Math.min((cleaned.split(".")[1] || "").length, 4);
  const dec = decimals ? "." + "0".repeat(decimals) : "";
  // Years like 2024 stay 2024; amounts with a currency sign are always grouped, as Power BI
  // shows them (it exports "$2627668" for $2,627,668)
  const thousands = s.includes(",") || currency ? "#,##0" : "0";

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
    // Power BI exports a percentage measure as a fraction (0.44): a "%" field reads as one
    if (/%|\b(percent|pct)\b/i.test(name) && cells.every((c) => Math.abs(c.value) <= 10)) return "0.0%";
    return cells.some((c) => !Number.isInteger(c.value)) ? "#,##0.00" : "#,##0";
  });
}

/* A table's number columns in display units, as Power BI's "Auto" units show them: one unit
   per column from its largest value (thousands from 10,000), two decimals; percentages, years,
   codes and small numbers unchanged. Formats already set for a column are kept. */
/* "Numbers: As in the report": each column the report gives display units or decimals, in them
   (the cells keep their exact values); null when the report sets none */
function reportFormats(item, info) {
  const set = info && info.columnUnits;
  if (!set) return null;
  const own = item.cellFormats || [];
  const { header, data } = classifyColumns(item.rows);
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci])));
  const readable = readableFormats(header, cells);
  // a column renamed in the visual ("Qualified Pipeline") has its settings under its field
  // ("Opportunities.Revenue In Pipeline")
  const entries = info.roles ? fieldsByColumn(header, Object.values(info.roles).flat()) : header.map(() => null);
  const settingFor = (name, ci) => {
    const e = entries[ci];
    const keys = [name, e && e.queryRef, e && e.field && e.field.prop].filter(Boolean).map((k) => "m:" + normalizeField(k));
    return keys.map((k) => set[k]).find(Boolean) || null;
  };
  let changed = false;
  item.fixedUnits = header.map(() => false);
  const out = header.map((name, ci) => {
    const s = settingFor(name, ci);
    const nums = cells.map((r) => r[ci].value).filter((v) => typeof v === "number");
    if (!s || !nums.length) return own[ci] || null;
    const base = own[ci] || readable[ci] || (cells.map((r) => r[ci]).find((c) => c.numFmt) || {}).numFmt || "#,##0";
    if (/[dmyhs]/i.test(base.replace(/"[^"]*"/g, ""))) return own[ci] || null;   // dates
    changed = true;
    item.fixedUnits[ci] = s.units > 1;   // the column's unit set in the report: its total keeps it
    // a table's "Auto" units show the full number, as Power BI's tables do
    return labelNumFmt(base, nums, { units: s.units > 0 ? s.units : 1, precision: s.precision });
  });
  return changed ? out : null;
}

function shortFormats(item) {
  const own = item.cellFormats || [];
  const { header, data, numeric, helper } = classifyColumns(item.rows);
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci])));
  const readable = readableFormats(header, cells);
  const isCode = (n) => Number.isInteger(n) && n >= 190001 && n <= 210012;   // 202301: a year and month
  return header.map((name, ci) => {
    // A column already given a format keeps it, unless it is a plain number format
    if (own[ci] && !/^("[^"]*"|[$€£¥₹])?#,##0(\.0+)?$/.test(own[ci])) return own[ci];
    const nums = cells.map((r) => r[ci]).filter((c) => typeof c.value === "number");
    if (!numeric[ci] || helper[ci] || !nums.length || /\b(year|month|code|zip|postal|phone)\b/i.test(name) || nums.every((c) => isCode(c.value))) return own[ci] || null;
    const base = own[ci] || readable[ci] || (nums.find((c) => c.numFmt) || {}).numFmt || "#,##0";
    if (base.includes("%") || /[dmyhs]/i.test(base.replace(/"[^"]*"/g, ""))) return own[ci] || null;   // percentages, dates
    const top = maxAbsOf(nums.map((c) => c.value));
    return top >= 1e4 ? unitFormat(base, autoUnit(top), 2) : own[ci] || null;
  });
}

/* A total in short numbers: a larger unit when it outgrows its column's (1.16M under a column
   of thousands); other formats unchanged */
function unitForTotal(fmt, value) {
  const m = String(fmt || "").match(/^(.*?)(,+)"(K|M|bn|T)"$/);
  if (!m || typeof value !== "number") return fmt;
  const unit = autoUnit(Math.abs(value));
  return unit > Math.pow(1000, m[2].length) ? unitFormat(m[1], unit, fmtDecimals(m[1])) : fmt;
}

/* Roughly how Excel shows a value with a number format (used to size columns) */
function displayText(value, fmt) {
  if (value instanceof Date) return dateText(value, /h/.test(fmt || ""));
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

/*
 * Excel width units a text needs in its cell, measured in the cell's own font (a column
 * of w units is about w × 7 + 5 pixels at 100%; an indent level is about 9 pixels).
 * Where the browser doesn't have the font (Office fonts on a Mac), the fallback font is
 * wider, so the column comes out roomier, never too narrow. The 15% margin covers the
 * way Excel rounds text to whole pixels, more so when zoomed in or out.
 */
const TEXT_MARGIN = 1.15;
let measureContext = null;
function cellUnits(text, font, indent = 0) {
  const s = String(text ?? "");
  if (!s) return 0;
  const px = ((font && font.size) || 11) * 96 / 72;
  let width;
  try {
    if (!measureContext) measureContext = document.createElement("canvas").getContext("2d");
    measureContext.font = `${font && font.bold ? "bold " : ""}${px}px "${(font && font.name) || "Calibri"}", Calibri, Carlito, Arial, sans-serif`;
    width = measureContext.measureText(s).width;
  } catch (e) {
    width = s.length * px * 0.6 * (font && font.bold ? 1.06 : 1);   // no canvas: a generous average
  }
  return Math.max(0, (width * TEXT_MARGIN + 6 + indent * 9 - 5) / 7);
}

/* The text a compact KPI format (2.52M) shows */
/* "Numbers: As in the report": a card's value in the report's display units (Auto: as Power BI
   picks them) and decimals (Auto: none from ten of the unit, one below: 18K, $110M, 3.9M); null
   when it shows as exported */
function reportValueFormat(item, info, value, fmt) {
  if (typeof value !== "number") return null;
  const set = (info && info.valueUnits) || (VALUE_UNIT_SETTINGS[item.type] ? { units: 0, precision: null } : null);
  if (!set) return null;
  const f = labelNumFmt(fmt || "#,##0", [value], set, 0, autoUnit);   // cards shorten from 1,000, Auto decimals as labels
  return f && f !== (fmt || "#,##0") ? f : null;
}

function compactText(value, fmt) {
  const prefix = ((fmt || "").match(/^"([^"]*)"/) || [])[1] || "";
  const a = Math.abs(value);
  const [scale, unit] = a >= 1e12 ? [1e12, "T"] : a >= 1e9 ? [1e9, "B"] : [1e6, "M"];
  return `${value < 0 ? "-" : ""}${prefix}${(a / scale).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}${unit}`;
}
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

/*
 * A gauge or KPI shows one field as its main value and the others (target, minimum,
 * maximum) as context, but Power BI exports them in the order they were added. The main
 * value is moved first, so it becomes the tile's large value. The report file names it;
 * without the file, a gauge's title names it first ("This Year Sales and Last Year Sales").
 */
function mainValueFirst(item, info) {
  const header = item.rows[0].map((h) => String(h ?? "").trim());
  let k = -1;
  if (info && info.mainValue) {
    const f = info.mainValue;
    if (f.label) k = header.findIndex((h) => h.toLowerCase() === f.label);
    if (k < 0) k = header.findIndex((h) => normalizeField(h) === f.ref);
  } else if (!info && item.type === "gauge" && item.title) {
    const title = item.title.toLowerCase();
    let at = Infinity;
    header.forEach((h, i) => {
      const p = h ? title.indexOf(h.toLowerCase()) : -1;
      if (p >= 0 && (p < at || (p === at && h.length > header[k].length))) { at = p; k = i; }
    });
  }
  if (k > 0) item.rows = item.rows.map((row) => [row[k], ...row.filter((_, i) => i !== k)]);
}

/*
 * Nested Bars draws one bar per category: the outer bar is the category's total of its
 * category value, the inner bar the subset value of ONE subset (the one chosen in the
 * visual, else the first in the data). Power BI exports a row per category and subset, so
 * the rows are rebuilt to read like the chart: category, its total, the chosen subset's
 * value (headed "Value · Subset"). A total of a ratio can't be added up from the rows: it
 * is averaged, with a note.
 */
function nestedBarsRows(item, info, notes) {
  const { header, data, numeric, hex } = classifyColumns(item.rows);
  const cols = header.map((_, ci) => ci);
  const norm = header.map((h) => normalizeField(h)), lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const taken = new Set();
  const colOf = (f) => {
    if (!f) return -1;
    let ci = f.label ? cols.find((c) => !taken.has(c) && lower[c] === f.label) : undefined;
    if (ci === undefined) ci = cols.find((c) => !taken.has(c) && norm[c] === f.ref);
    if (ci !== undefined) taken.add(ci);
    return ci ?? -1;
  };
  const texts = cols.filter((ci) => !numeric[ci] && !hex[ci]), nums = cols.filter((ci) => numeric[ci] && !hex[ci]);
  const roles = (info && info.roles) || {};
  let cat = colOf((roles.Category || [])[0]), sub = colOf((roles.Subset || [])[0]);
  let outer = colOf((roles.Y || [])[0]), inner = colOf((roles.Y2 || [])[0]);
  if (cat < 0) cat = texts[0] ?? -1;
  if (sub < 0) sub = texts.find((ci) => ci !== cat) ?? -1;
  if (outer < 0) outer = nums[0] ?? -1;
  if (inner < 0) inner = nums.find((ci) => ci !== outer) ?? -1;
  if (cat < 0 || sub < 0 || outer < 0) return;   // one level only: the rows already read like the chart
  const key = (r, ci) => String(r[ci] ?? "").trim();
  const num = (r, ci) => { const v = toCellValue(r[ci]).value; return typeof v === "number" ? v : null; };
  const subsets = [...new Set(data.map((r) => key(r, sub)))];
  const chosen = customSetting(item.type, info, "subset");
  const shown = chosen && subsets.includes(chosen) ? chosen : subsets[0];
  const additive = !NON_ADDITIVE.test(header[outer]) && !data.some((r) => key(r, outer).endsWith("%"));
  const raw = (r, ci) => String(r[ci] ?? "").trim();
  const fmtLike = (ci, v) => {   // write the total the way the column's values are written ("$7052", "12.5%")
    const sample = data.map((r) => raw(r, ci)).find(Boolean) || "";
    const cur = (sample.match(/^-?\s*([$€£₹¥])/) || [])[1] || "";
    if (sample.endsWith("%")) return (v * 100).toFixed(2) + "%";
    return (v < 0 ? "-" : "") + cur + String(Math.abs(Math.round(v * 100) / 100));
  };
  const rows = [];
  for (const c of [...new Set(data.map((r) => key(r, cat)))]) {
    const mine = data.filter((r) => key(r, cat) === c);
    const vals = mine.map((r) => num(r, outer)).filter((v) => v !== null);
    const total = vals.length ? (additive ? vals.reduce((a, b) => a + b, 0) : vals.reduce((a, b) => a + b, 0) / vals.length) : null;
    const pick = mine.find((r) => key(r, sub) === shown);
    rows.push([mine[0][cat], total === null ? "" : fmtLike(outer, total), inner >= 0 && pick ? pick[inner] : ""]);
  }
  const head = [header[cat], header[outer], ...(inner >= 0 ? [`${header[inner]} · ${shown || "subset"}`] : [])];
  item.rows = [head, ...rows.map((r) => (inner >= 0 ? r : r.slice(0, 2)))];
  item.nestedCols = { outer: 1, inner: inner >= 0 ? 2 : -1 };
  if (!additive) notes.push(`"${item.title}": each bar's total is averaged from Power BI's rows (a ratio can't be added up), so it may differ slightly from the chart.`);
}

/*
 * The Comparison Table draws a grid: a row per metric, its benchmark (the first region's), then
 * a column per region holding that region's value, each cell with a bar: blue for the
 * benchmark, green for a value (red when negative). Power BI exports a row per metric and
 * region, so the rows are rebuilt into that grid and the bars become data bars on its
 * columns (header colours set in the report still apply). Needs the report file to tell the
 * benchmark from the values.
 */
function comparisonRows(item, info) {
  const roles = (info && info.roles) || {};
  const { header, data } = classifyColumns(item.rows);
  const norm = header.map((h) => normalizeField(h)), lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const taken = new Set();
  const colOf = (f) => {
    if (!f) return -1;
    let ci = f.label ? header.findIndex((_, c) => !taken.has(c) && lower[c] === f.label) : -1;
    if (ci < 0) ci = header.findIndex((_, c) => !taken.has(c) && norm[c] === f.ref);
    if (ci >= 0) taken.add(ci);
    return ci;
  };
  const metric = colOf((roles.Rows || [])[0]), region = colOf((roles.Columns || [])[0]);
  const bench = colOf((roles.Values || [])[0]), value = colOf((roles.Values || [])[1]);
  if (metric < 0 || region < 0 || (bench < 0 && value < 0)) return;
  const key = (r, ci) => String(r[ci] ?? "").trim();
  const metrics = [...new Set(data.map((r) => key(r, metric)))];
  const regions = [...new Set(data.map((r) => key(r, region)))];
  const cell = (m, g, ci) => { const r = data.find((x) => key(x, metric) === m && key(x, region) === g); return r && ci >= 0 ? r[ci] : ""; };
  const head = [header[metric], ...(bench >= 0 ? [header[bench]] : []), ...(value >= 0 ? regions.map((g) => g || "(Blank)") : [])];
  const rows = metrics.map((m) => [m, ...(bench >= 0 ? [cell(m, regions[0], bench)] : []), ...(value >= 0 ? regions.map((g) => cell(m, g, value)) : [])]);
  item.rows = [head, ...rows];
  const custom = customVisual(item.type);
  const bar = (name, b) => ({ target: "m:" + normalizeField(name), kind: "bar", bar: { color: b.color, negative: b.negative || null, axis: null, reverse: false, showValue: true } });
  item.gridFormat = { measures: [], tableFormats: [
    ...(bench >= 0 ? [bar(head[1], custom.bars.benchmark)] : []),
    ...(value >= 0 ? head.slice(bench >= 0 ? 2 : 1).map((g) => bar(g, custom.bars.values)) : []),
    ...((info && info.tableFormats) || []).filter((f) => f.target === "header"),
  ] };
}

/*
 * The Histogram draws its values in bins, as its code (d3 v5) makes them: the bin count is
 * the visual's setting, else d3's own (Sturges' rule on a nice step); the bins split the range
 * from the smallest to the largest value evenly, a value falls in (x0, x1] (the first bin also
 * takes its lower edge), and each value adds its frequency (the Frequency field, 1 without
 * one). Power BI exports a row per value, so the rows are rebuilt as one per bin.
 */
function histogramRows(item, info) {
  const { header, data } = classifyColumns(item.rows);
  const roles = (info && info.roles) || {};
  const find = (f) => (f ? header.findIndex((h) => normalizeField(h) === f.ref || String(h).trim().toLowerCase() === f.label) : -1);
  let vc = find((roles.Category || [])[0]), fc = find((roles.Frequency || [])[0]);
  if (vc < 0) vc = 0;
  if (fc < 0 && header.length > 1) fc = vc === 0 ? 1 : 0;
  const num = (x) => { const v = toCellValue(x).value; return typeof v === "number" ? v : null; };
  const points = data.map((r) => ({ v: num(r[vc]), f: fc >= 0 ? num(r[fc]) : null })).filter((p) => p.v !== null)
    .map((p) => ({ v: p.v, f: p.f !== null && p.f > 1 ? p.f : 1 }));
  if (points.length < 2) return;
  const values = points.map((p) => p.v);
  const min = Math.min(...values), max = Math.max(...values);
  if (min === max) return;
  // d3 v5: tickStep and range, as d3.histogram uses them
  const tickStep = (a, b, count) => {
    const step0 = Math.abs(b - a) / Math.max(0, count);
    let step1 = Math.pow(10, Math.floor(Math.log(step0) / Math.LN10));
    const error = step0 / step1;
    if (error >= Math.sqrt(50)) step1 *= 10; else if (error >= Math.sqrt(10)) step1 *= 5; else if (error >= Math.sqrt(2)) step1 *= 2;
    return b < a ? -step1 : step1;
  };
  const range = (a, b, step) => Array.from({ length: Math.max(0, Math.ceil((b - a) / step)) }, (_, i) => a + i * step);
  const binCount = (thresholds) => {
    const t = [...thresholds];
    while (t.length && t[0] <= min) t.shift();
    while (t.length && t[t.length - 1] > max) t.pop();
    return t;
  };
  const set = Number(String(customSetting(item.type, info, "bins") ?? "").replace(/[LD]$/, ""));
  let count = set > 0 ? Math.round(set) : 0;
  if (!count) {
    const sturges = Math.ceil(Math.log(values.length) / Math.LN2) + 1;
    const step = tickStep(min, max, sturges);
    count = binCount(range(Math.ceil(min / step) * step, max, step)).length + 1;
  }
  const edges = [min, ...binCount(range(min, max, (max - min) / count)), max];
  const fmt = (x) => Number(x.toFixed(2)).toLocaleString("en-US", { maximumFractionDigits: 2 });
  const rows = [];
  for (let i = 0; i < edges.length - 1; i++) {
    const [x0, x1] = [edges[i], edges[i + 1]];
    const freq = points.filter((p) => (i === 0 ? p.v >= x0 : p.v > x0) && p.v <= x1).reduce((a, p) => a + p.f, 0);
    rows.push([`${fmt(x0)} – ${fmt(x1)}`, String(freq)]);
  }
  item.rows = [[`${header[vc]} (bin)`, fc >= 0 ? header[fc] : "Frequency"], ...rows];
}

/* ── Trend cards ──────────────────────────────────────────────────────────
 * A KPI drawn with its trend becomes a card per value field above the visual's data table:
 * the latest value, its change, and an Excel sparkline of the values over time (added to the
 * finished file by addSparklinesToFile). item.trend:
 *   { dateCol, points, cards: [{ col, label, value, numFmt, text, change, since, color, min, max }],
 *     caption: { title, days, from, to, stale, staleText, warning } }
 * The rows are put in date order, rows without a date last, so each sparkline reads the
 * first `points` data rows.
 */

/* Dual KPI's percentages: one decimal, cut (not rounded), "+" when asked */
function dualKpiPercent(v, sign) {
  return `${sign && v >= 0 ? "+" : ""}${Math.floor(v * 10) / 10}%`;
}

/* d3's ".3s" (1.19k): three significant digits and an SI unit, as an Excel format for this
   value, and as text */
function siUnit(value) {
  const a = Math.abs(value);
  return a >= 1e12 ? [1e12, ",,,,", "T"] : a >= 1e9 ? [1e9, ",,,", "G"] : a >= 1e6 ? [1e6, ",,", "M"] : a >= 1e3 ? [1e3, ",", "k"] : [1, "", ""];
}
function siFormat(value) {
  const [div, commas, unit] = siUnit(value);
  const a = Math.abs(value) / div;
  const dec = 3 - (a >= 100 ? 3 : a >= 10 ? 2 : 1);
  return `0${dec ? "." + "0".repeat(dec) : ""}${commas}${unit ? `"${unit}"` : ""}`;
}
function siText(value) {
  const [div, , unit] = siUnit(value);
  return (value / div).toPrecision(3) + unit;
}

/* A number format with `places` decimals, shown in thousands / millions … (Power BI's display
   units: 1 none, 0 auto, else the unit) */
const DISPLAY_UNITS = [[1e12, ",,,,", "T"], [1e9, ",,,", "bn"], [1e6, ",,", "M"], [1e3, ",", "K"]];
function displayUnit(units, value) {
  if (units === 0) return DISPLAY_UNITS.find(([u]) => Math.abs(value) >= u) || null;
  return DISPLAY_UNITS.find(([u]) => u === units) || null;
}
function unitsFormat(fmt, places, unit) {
  const tail = (places > 0 ? "." + "0".repeat(places) : "") + (unit ? `${unit[1]}"${unit[2]}"` : "");
  return (fmt || "#,##0").split(";").map((section) => section.replace(/(#,##0|0)(\.0+)?/, `$1${tail}`)).join(";");
}

/*
 * Dual KPI draws two charts, one per value field (top and bottom), each with the field's
 * latest value and its change since the first date (or the "% change start date" set in the
 * visual or given as a field); "Last N days" spans the first to the last date. Its code skips
 * rows without a date and sorts the rest by date. Percent fields show the change in points.
 */
function dualKpiTrend(item, info, colorsOn) {
  const set = (key) => customSetting(item.type, info, key);
  const num = (x) => {
    const n = Number(String(x ?? "").replace(/[LDM]$/i, ""));
    return x === null || x === undefined || x === "" || !Number.isFinite(n) ? null : n;
  };
  const { header, data, numeric } = classifyColumns(item.rows);
  const dayFirst = dayFirstColumns(header, data);
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci], dayFirst[ci])));
  const formats = readableFormats(header, cells);
  const roles = (info && info.roles) || {};
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const taken = new Set();
  const colOf = (f) => {
    if (!f) return -1;
    let ci = f.label ? lower.findIndex((h, c) => !taken.has(c) && h === f.label) : -1;
    if (ci < 0) ci = header.findIndex((h, c) => !taken.has(c) && normalizeField(h) === f.ref);
    return ci;
  };
  const isDate = (ci) => cells.some((r) => r[ci].value instanceof Date) && cells.every((r) => r[ci].value === null || r[ci].value instanceof Date);

  // The axis: the visual's field, else the first column of dates
  let dateCol = colOf((roles.Category || [])[0]);
  if (dateCol < 0 || !isDate(dateCol)) dateCol = header.findIndex((_, ci) => isDate(ci));
  if (dateCol < 0) return;
  taken.add(dateCol);
  const fieldCol = (role) => { const ci = colOf((roles[role] || [])[0]); if (ci >= 0) taken.add(ci); return ci; };
  const startCols = { top: fieldCol("toppercentdate"), bottom: fieldCol("bottompercentdate") };
  const warnCol = fieldCol("warningstate");
  // The values: the visual's fields, else (without the report file) the first two number columns
  const valueCol = (role) => {
    if (info && !(roles[role] || []).length) return -1;
    let ci = colOf((roles[role] || [])[0]);
    if (ci < 0) ci = header.findIndex((_, c) => numeric[c] && !taken.has(c));
    if (ci >= 0) taken.add(ci);
    return ci;
  };
  const cols = { top: valueCol("topvalues"), bottom: valueCol("bottomvalues") };

  const points = cells.map((r, i) => ({ i, date: r[dateCol].value })).filter((p) => p.date instanceof Date)
    .sort((a, b) => a.date - b.date || a.i - b.i);
  if (!points.length) return;

  const DAY = 86400000;
  const daysBetween = (a, b) => Math.round(Math.abs(a - b) / DAY);
  const first = points[0].date, last = points[points.length - 1].date;
  const topColor = colorsOn ? set("top") || "#01B8AA" : null;
  const colors = { top: topColor, bottom: colorsOn ? (set("matchBottom") !== false ? topColor : set("bottom") || topColor) : null };

  const cards = [];
  for (const pos of ["top", "bottom"]) {
    const ci = cols[pos];
    if (ci < 0 || set(pos === "top" ? "showTop" : "showBottom") === false) continue;
    const fmt = formats[ci] || cells[points[points.length - 1].i][ci].numFmt || null;
    const asPercent = /%/.test(fmt || "");
    // The visual reads a blank as 0, and a percent field in percent points
    const values = points.map((p) => { const v = cells[p.i][ci].value; return (typeof v === "number" ? v : 0) * (asPercent ? 100 : 1); });
    const latest = values[values.length - 1];

    // The change's start: the last date on or before the start date (from the field, else the setting), else the first
    const given = startCols[pos] >= 0 && cells.length ? cells[0][startCols[pos]].value : null;
    const typed = String(set(pos === "top" ? "topStart" : "bottomStart") || "").trim();
    const calc = given instanceof Date ? given : typed ? (parseDate(typed) || {}).date || null : null;
    let k = 0;
    if (calc) {
      for (let i = 0; i < points.length; i++) {
        const t = points[i].date.getTime();
        if (t === calc.getTime()) { k = i; break; }
        if (t < calc.getTime()) k = i; else break;
      }
    }
    const start = values[k];
    let change;
    if (asPercent) change = dualKpiPercent(latest - start, true);
    else if (start === 0) change = "n/a";
    else change = dualKpiPercent((latest < start ? -1 : 1) * Math.abs((latest - start) / start) * 100, true);

    // The latest value as the visual formats it: its decimal places and display units, or abbreviated
    const raw = cells[points[points.length - 1].i][ci].value;
    let value = raw, numFmt = null, text;
    if (set(pos === "top" ? "showTopValue" : "showBottomValue") === false) {
      value = text = String(set(pos === "top" ? "topDefault" : "bottomDefault") ?? "N/A");
      change = "";
    } else if (typeof raw !== "number") {
      value = text = "(Blank)";
    } else if (asPercent) {
      numFmt = "0.0%";
      text = dualKpiPercent(latest, false);
    } else if (set("abbreviate") === true) {
      numFmt = siFormat(raw);
      text = siText(raw);
    } else {
      const places = Math.min(17, Math.max(0, num(set("precision")) ?? 0));
      const unit = displayUnit(num(set("units")) ?? 1, raw);
      numFmt = unitsFormat(fmt, places, unit);
      text = displayText(unit ? raw / unit[0] : raw, unitsFormat(fmt, places, null)) + (unit ? unit[2] : "");
    }
    const axis = (key) => { const v = num(set(pos + key)); return v === null ? null : asPercent ? v / 100 : v; };
    cards.push({ col: ci, label: header[ci], value, numFmt, text, change, since: change ? dateText(points[k].date) : "",
      color: colors[pos], min: axis("Min"), max: axis("Max") });
  }
  if (!cards.length) return;

  const title = String(set("titleText") || "").trim();
  const age = daysBetween(last.getTime(), Date.now());
  const warn = warnCol >= 0 && cells.length ? cells[cells.length - 1][warnCol].value : null;
  const dated = new Set(points.map((p) => p.i));
  item.rows = [header, ...points.map((p) => data[p.i]), ...data.filter((_, i) => !dated.has(i))];
  item.trend = {
    dateCol, points: points.length, cards,
    caption: {
      title: title && title !== "Title" ? title : null,   // "Title" is the visual's placeholder
      days: daysBetween(first.getTime(), last.getTime()), from: dateText(first), to: dateText(last),
      stale: set("staleWarning") !== false && age >= (num(set("staleDays")) ?? 2) ? age : null,
      staleText: String(set("staleText") || "").trim(),
      warning: typeof warn === "number" && warn < 0 ? String(set("warningText") || "").trim() || "Warning message" : null,
    },
  };
}

/*
 * Power BI's KPI visual with a trend axis, as its code draws it: the indicator's value on the
 * last point of the axis, "Goal: x" (two goals: "Goal: x, y") and the distance to a single
 * goal ("(+12.5%)", or the amount, or both), the trend as an area over evenly spaced points,
 * and a status colour from the theme: good, bad (below the goal when high is good, above it
 * when low is good), neutral (between two goals), grey without a goal. The value and the
 * trend take the status colour, so the card reads like the visual.
 */
const KPI_STATUS_TEXT = { good: "Target met", neutral: "Between targets", bad: "Target not met" };
function kpiTrend(item, info, colorsOn) {
  const { header, data, numeric } = classifyColumns(item.rows);
  if (data.length < 2) return;
  const dayFirst = dayFirstColumns(header, data);
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci], dayFirst[ci])));
  const formats = readableFormats(header, cells);
  const roles = (info && info.roles) || {};
  const b = (info && info.builtin) || {};
  const found = fieldsByColumn(header, [...(roles.TrendLine || []), ...(roles.Indicator || []), ...(roles.Goal || [])]);
  const roleCol = (role) => found.findIndex((e) => e && (roles[role] || []).includes(e));
  let trendCol = roleCol("TrendLine"), valueCol = roleCol("Indicator");
  let goalCols = (roles.Goal || []).map((e) => found.indexOf(e)).filter((c) => c >= 0).slice(0, 2);
  // Without the report file: the axis is the first column that isn't a number, then the
  // indicator and up to two goals
  if (trendCol < 0) trendCol = header.findIndex((_, ci) => !numeric[ci]);
  if (valueCol < 0) valueCol = header.findIndex((_, ci) => numeric[ci] && ci !== trendCol);
  if (!info) goalCols = header.map((_, ci) => ci).filter((ci) => numeric[ci] && ci !== trendCol && ci !== valueCol).slice(0, 2);
  if (trendCol < 0 || valueCol < 0) return;

  // Points in axis order (dates by date; anything else as Power BI sorted it)
  const order = cells.map((r, i) => i);
  const dates = cells.every((r) => r[trendCol].value instanceof Date);
  if (dates) order.sort((x, y) => cells[x][trendCol].value - cells[y][trendCol].value || x - y);
  item.rows = [header, ...order.map((i) => data[i])];
  const last = cells[order[order.length - 1]];
  const actual = last[valueCol].value;
  const goals = goalCols.map((c) => last[c].value).filter((v) => typeof v === "number");
  const fmtOf = (c) => formats[c] || (cells.find((r) => typeof r[c].value === "number") || [])[c]?.numFmt || null;
  const fmt = fmtOf(valueCol);

  let status = null;
  if (goals.length && typeof actual === "number") {
    const hi = Math.max(...goals), lo = Math.min(...goals);
    if (goals.length === 2 && actual >= lo && actual <= hi) status = "neutral";
    else if (b.direction === "Negative") status = actual > hi ? "bad" : "good";
    else status = actual < lo ? "bad" : "good";
  }
  const statusHex = status ? b[status] || KEYWORD_DEFAULTS[status] : "#8C8C8C";   // the card takes hex colours
  const statusColor = toArgb(statusHex);

  const note = [];
  if (goals.length && b.showGoal !== false) {
    note.push([`${b.goalText || "Goal"}: ${goals.map((g, i) => displayText(g, fmtOf(goalCols[i]) || fmt)).join(", ")}`, "plain"]);
  }
  if (goals.length === 1 && goals[0] !== 0 && typeof actual === "number" && b.showDistance !== false) {
    const highIsGood = (b.distanceDirection || (b.direction === "Negative" ? "Low is good" : "High is good")) !== "Low is good";
    const sign = (highIsGood ? actual < goals[0] : actual > goals[0]) ? "-" : "+";
    const diff = Math.abs(actual - goals[0]);
    const pct = `${sign}${(diff / Math.abs(goals[0]) * 100).toLocaleString("en-US", { maximumFractionDigits: 2 })}%`;
    const amount = `${sign}${displayText(diff, fmt)}`;
    const label = String(b.distanceLabel || "Percent");
    const text = /both/i.test(label) ? `(${amount} ${pct})` : /value/i.test(label) ? `(${amount})` : `(${pct})`;
    note.push([`${note.length ? "  " : ""}${text}`, "strong", colorsOn ? readableOnWhite(statusColor) : null]);
  }
  const label = (v) => (v instanceof Date ? dateText(v) : String(v ?? ""));
  const compact = compactFormat(actual, fmt);
  item.trend = {
    dateCol: trendCol, points: order.length, dateAxis: false,
    cards: [{
      col: valueCol, label: header[valueCol], value: actual ?? "(Blank)", numFmt: compact || fmt,
      text: typeof actual === "number" ? (compact ? compactText(actual, fmt) : displayText(actual, fmt)) : "(Blank)",
      note, color: colorsOn ? statusHex : null, valueColor: colorsOn && status ? statusHex : null,
      noSpark: b.showTrend === false, min: null, max: null,
      // for the KPI drawn as a chart (charts mode): its numbers as the visual shows them
      kpi: { actual, goals, goalText: b.showGoal !== false ? b.goalText || "Goal" : null, fmt, status,
        distance: note.length > (goals.length && b.showGoal !== false ? 1 : 0) ? note[note.length - 1][0].trim() : null },
    }],
    captionRuns: [
      [`${header[trendCol]}  ${label(cells[order[0]][trendCol].value)} – ${label(last[trendCol].value)}`, "plain"],
      b.showLastDate ? [`  ·  as of ${label(last[trendCol].value)}`, "plain"] : null,
      status ? [`  ·  ${KPI_STATUS_TEXT[status]}`, "strong", colorsOn ? readableOnWhite(statusColor) : null] : null,
    ],
  };
}

/* ── Totals and matrices ──────────────────────────────────────────────────
 * Power BI exports neither a table's totals nor a matrix's subtotals, and a matrix comes
 * flat (row fields, values, column fields last; a row per row and column value). Totals are
 * written as live formulas (SUBTOTAL 9: what the user filters out in Excel leaves the total,
 * rows folded under [+] stay in it), only where adding up the rows gives Power BI's own
 * total: sums and counts, minimum / maximum, and measures that aren't ratios, averages or
 * distinct counts by their name or format. Other columns' totals stay blank.
 */
const COUNT_LIKE = /\b(unique|distinct|dcount|customers?|clients?|users?|visitors?|members?|accounts?|headcount|active)\b/i;
const RATE_LIKE = /\b(price|rate|per|unit|avg|average)\b/i;
const MEASURE_RATE = /\b(price|per|avg|average)\b/i;   // a measure's rate words (rate itself: NON_ADDITIVE)
const CHANGE_LIKE = /\b(yoy|mom|qoq|wow|growth|change|var|variance|index|score|share)\b/i;
const SUBTOTAL_FN = { sum: 9, min: 5, max: 4 };

/* How a column totals: "sum", "min", "max", or null (left blank) */
function totalKind(field, name, fmt) {
  if (/%/.test(fmt || "")) return null;
  const names = [name, field && field.prop].filter(Boolean).join(" ");
  if (field) {
    if (field.kind === "agg") return field.agg === 0 || field.agg === 5 ? "sum" : field.agg === 3 ? "min" : field.agg === 4 ? "max" : null;
    // a rate or average ("Sales Per Sq Ft", "Avg Price") never adds up across rows ("Business
    // Unit" in a count's name doesn't make it one)
    if (field.kind === "measure") return NON_ADDITIVE.test(names) || COUNT_LIKE.test(names) || CHANGE_LIKE.test(names) || MEASURE_RATE.test(names) ? null : "sum";
    return null;   // a column shown as it is, a hierarchy level
  }
  return NON_ADDITIVE.test(names) || COUNT_LIKE.test(names) || CHANGE_LIKE.test(names) || RATE_LIKE.test(names) ? null : "sum";
}

/* Each column's field in the report: its renamed label first, then its name */
function fieldsByColumn(header, entries) {
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const norm = header.map((h) => normalizeField(h));
  const out = header.map(() => null), taken = new Set();
  // Two different fields shown under one name (Seps YoY and Actives YoY, both renamed "YoY"):
  // which column is which can't be told from the export, so those columns stay unmatched and
  // nothing (a total, units, a colour rule) is worked out for them on a guess
  const fieldsByLabel = {};
  for (const e of entries || []) {
    if (e && e.label) (fieldsByLabel[e.label] = fieldsByLabel[e.label] || new Set()).add(e.queryRef || e.ref || JSON.stringify(e.field));
  }
  const ambiguous = new Set(Object.keys(fieldsByLabel).filter((l) => fieldsByLabel[l].size > 1));
  const sameName = (h, l) => h === l || h.startsWith(l + " (") && /^\(\d+\)$/.test(h.slice(l.length + 1));
  lower.forEach((h, c) => { if ([...ambiguous].some((l) => sameName(h, l))) taken.add(c); });
  for (const e of entries || []) {
    if (e && e.label && ambiguous.has(e.label)) continue;
    let ci = e.label ? lower.findIndex((h, c) => !taken.has(c) && h === e.label) : -1;
    if (ci < 0) ci = norm.findIndex((h, c) => !taken.has(c) && h === e.ref);
    if (ci >= 0) { out[ci] = e; taken.add(ci); }
  }
  return out;
}

/* A table's (or flat matrix's) totals row: { kinds: per column "sum" | "min" | "max" | null },
   or null when the report turns totals off or nothing adds up */
function totalsPlan(item, info) {
  const b = (info && info.builtin) || {};
  if ((item.type === "tableEx" && b.totals === false) || (item.type === "pivotTable" && b.rowSubtotals === false)) return null;
  const { header, data, numeric, helper } = classifyColumns(item.rows);
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci])));
  const formats = readableFormats(header, cells);
  const entries = info ? fieldsByColumn(header, [...((info.roles || {}).Values || []), ...((info.roles || {}).Rows || [])]) : header.map(() => null);
  const kinds = header.map((h, ci) => {
    if (item.columnKinds) return item.columnKinds[ci] || null;
    if (!numeric[ci] || helper[ci] || /\byear\b/i.test(h)) return null;
    const fmt = formats[ci] || (cells.find((r) => typeof r[ci].value === "number") || [])[ci]?.numFmt;
    const kind = totalKind(info && entries[ci] ? entries[ci].field || null : null, h, fmt);
    // a rate, share or average: Power BI's own total, asked of it (loadTableTotals), else blank
    if (!kind && item.reportTotals && typeof item.reportTotals[normalizeField(h)] === "number") return "report";
    return kind;
  });
  return kinds.some(Boolean) ? { kinds, report: item.reportTotals || null } : null;
}

/* The total of numbers as Excel's SUBTOTAL would give it */
function subtotalOf(kind, values) {
  const nums = values.filter((v) => typeof v === "number");
  if (!nums.length) return 0;
  return kind === "min" ? Math.min(...nums) : kind === "max" ? Math.max(...nums) : nums.reduce((a, b) => a + b, 0);
}

/* A1 reference of a run of rows in one column */
function colRange(ws, c, r1, r2) {
  const L = ws.getColumn(c).letter;
  return r1 === r2 ? `${L}${r1}` : `${L}${r1}:${L}${r2}`;
}

/* The totals row: "Total", then a live SUBTOTAL per column that adds up (data rows hr+1 … hr+n) */
function writeTotalsRow(ws, row, col, item, hr, n, lay, table) {
  const fonts = lay.fonts;
  const kinds = item.totals.kinds;
  const font = { ...fonts.header };
  kinds.forEach((kind, ci) => {
    const cell = ws.getCell(row, col + ci);
    cell.font = font;
    if (!kind) return;
    const values = table.data.map((r) => r[ci].value);
    const fmt = table.formats[ci] || (table.data.find((r) => typeof r[ci].value === "number") || [])[ci]?.numFmt;
    if (kind === "report") {
      // Power BI's own total (10.8% for all surgeons together): a value, not a formula
      cell.value = item.totals.report[normalizeField(item.rows[0][ci])];
      if (fmt) cell.numFmt = fmt;
      cell.alignment = { vertical: "middle", horizontal: "right", indent: 2 };
      return;
    }
    cell.value = { formula: `SUBTOTAL(${SUBTOTAL_FN[kind]},${colRange(ws, col + ci, hr + 1, hr + n)})`, result: subtotalOf(kind, values) };
    // a unit the report sets for the column (thousands: $11,429.83K) stays on its total
    if (fmt) cell.numFmt = item.fixedUnits && item.fixedUnits[ci] ? fmt : unitForTotal(fmt, cell.value.result);
    cell.alignment = { vertical: "middle", horizontal: "right", indent: 2 };
  });
  if (!kinds[0]) {
    const label = ws.getCell(row, col);
    label.value = "Total";
    label.alignment = { vertical: "middle", indent: 1 };
  }
}

/*
 * A visual's field wells at the drill level the report shows while exporting. The report file
 * lists every level of a hierarchy (Year, Quarter, Month), but Power BI exports a visual as it
 * is drawn: drilled down to Quarter, only Quarter; expanded one level, Year and Quarter; a
 * matrix only the row and column levels it has expanded. So the hierarchy roles keep just the
 * levels in the export, and everything after (chart categories, colours, the matrix layout)
 * follows the report's current level. The embedding API can't report drill state, but the
 * export always reflects it. AppSource visuals keep their roles (theirs mean something else).
 */
const HIERARCHY_ROLES = ["Category", "Group", "Rows", "Columns"];
// Field wells that group the data (the rest hold values)
const KEY_ROLES = [...HIERARCHY_ROLES, "Series", "Legend", "Details", "Location", "Breakdown", "Axis"];

/* Whether a matrix's saved state opens any member below its top level shown: a level
   expanded by default, or a member opened by hand (Apparel expanded, the rest closed) */
function someExpanded(expansion, role) {
  const st = (expansion || []).find((e) => (e.roles || []).includes(role));
  if (!st) return false;
  if ((st.levels || []).some((l) => l && l.isCollapsed === false)) return true;
  const opened = (n) => !!n && (n.isToggled === true || (n.children || []).some(opened));
  return opened(st.root);
}

/* The levels of one hierarchy role the export keeps. A matrix exported as a flat table can't
   show some members open and others closed, so when the report opens any of them it keeps
   every level from the first one shown (Category and Segment for every category); a matrix
   drilled down past a level, or with every member closed, keeps just the levels shown */
function levelsShown(item, info, role) {
  const list = ((info && info.roles) || {})[role] || [];
  const shown = shownLevels(list);
  if (!flatMatrix(item) || (role !== "Rows" && role !== "Columns") || shown.length === list.length) return shown;
  if (!someExpanded(info.expansion, role)) return shown;
  const first = list.indexOf(shown[0]);
  return first < 0 ? list : list.slice(first);
}

/* The levels of a hierarchy the report shows: those its file marks active (the drill level;
   several when expanded, "Chain Name"), else every level (no drill state saved) */
function shownLevels(list) {
  if (!list || list.length < 2 || !list.some((f) => f.active === true)) return list || [];
  return list.filter((f) => f.active === true);
}

function atDrillLevel(item, info) {
  if (!info || !info.roles || !item.rows || !item.rows.length || customVisual(item.type)) return info;
  const header = item.rows[0].map((h) => String(h ?? "").trim());
  const lower = header.map((h) => h.toLowerCase());
  const norm = header.map((h) => normalizeField(h));
  const present = (f) => !!f && ((f.label && lower.includes(f.label)) || norm.includes(f.ref));
  const roles = { ...info.roles };
  let changed = false;
  for (const r of HIERARCHY_ROLES) {
    const list = roles[r];
    if (!list || list.length < 2) continue;
    // The levels shown that are in the export; none found means the export names them
    // differently, so the report's list stays
    let kept = (item.drillKept ? list : levelsShown(item, info, r)).filter(present);
    if (!kept.length) kept = list.filter(present);
    if (kept.length && kept.length < list.length) { roles[r] = kept; changed = true; }
  }
  if (!changed) return info;
  const cats = roles.Category || roles.Group || [];
  return { ...info, roles, category: cats[0] || null, categories: cats };
}

/*
 * Power BI exports every level of a visual's hierarchy (Region and VP), even when the report
 * shows only the top one (a bar per Region). The rows are rolled up to the levels shown, as
 * the visual draws them: sums, counts, minimums and maximums from the exported rows; values
 * that don't add up (ratios, percentages, averages) are queried at that level from the
 * semantic model (needs Build permission); if that fails, the visual keeps every level Power BI
 * exported (item.drillKept), with a note. Returns a note, or null.
 */
async function rollUpToDrillLevel(item, info, filters) {
  if (!info || !info.roles || customVisual(item.type) || !item.rows || item.rows.length < 2) return null;
  const roles = info.roles;
  const hiddenFields = HIERARCHY_ROLES.flatMap((r) => {
    const list = roles[r] || [], shown = levelsShown(item, info, r);
    return list.filter((f) => !shown.includes(f));
  });
  if (!hiddenFields.length) return null;
  const header = item.rows[0], data = item.rows.slice(1);
  const keyFields = new Set(KEY_ROLES.flatMap((r) => levelsShown(item, info, r)));
  const entries = fieldsByColumn(header, [...hiddenFields, ...keyFields, ...Object.values(roles).flat()]);
  const hidden = header.map((_, ci) => hiddenFields.includes(entries[ci]));
  if (!hidden.some(Boolean)) return null;   // already exported at the level shown
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci])));
  const isNum = header.map((_, ci) => cells.some((r) => typeof r[ci].value === "number") && cells.every((r) => r[ci].value === null || r[ci].value === "" || typeof r[ci].value === "number"));
  const keyCols = header.map((_, ci) => ci).filter((ci) => !hidden[ci] && (keyFields.has(entries[ci]) || (!entries[ci] && !isNum[ci])));
  const valueCols = header.map((_, ci) => ci).filter((ci) => !hidden[ci] && !keyCols.includes(ci));
  const fmtOf = (ci) => (cells.find((r) => typeof r[ci].value === "number" && r[ci].numFmt) || [])[ci]?.numFmt || null;
  const kinds = Object.fromEntries(valueCols.map((ci) => [ci, totalKind(entries[ci] && entries[ci].field, header[ci], fmtOf(ci))]));

  // Group the rows by the levels (and legend) shown, in the export's order
  const groups = new Map();
  data.forEach((row, ri) => {
    const k = keyCols.map((ci) => matchKey(row[ci])).join("\u0001");
    if (!groups.has(k)) groups.set(k, { row, rows: [] });
    groups.get(k).rows.push(ri);
  });
    // Power BI sorted the exported rows by one value (Revenue Won, descending) at the detail level,
  // so the groups first appear in the wrong order (the biggest single product's group first):
  // put the groups in that value's order again, by their totals
  const sortBy = valueCols.find((ci) => {
    if (!kinds[ci]) return false;
    const v = cells.map((r) => r[ci].value).filter((x) => typeof x === "number");
    if (v.length < 3 || new Set(v).size < 2) return false;
    return v.every((x, i) => !i || x <= v[i - 1]) || v.every((x, i) => !i || x >= v[i - 1]);
  });
  if (sortBy !== undefined) {
    const v = cells.map((r) => r[sortBy].value).filter((x) => typeof x === "number");
    const desc = v[0] > v[v.length - 1];
    const total = (g) => subtotalOf(kinds[sortBy], g.rows.map((ri) => cells[ri][sortBy].value));
    const sorted = [...groups.entries()].sort((a, b) => (desc ? total(b[1]) - total(a[1]) : total(a[1]) - total(b[1])));
    groups.clear();
    sorted.forEach(([k, g]) => groups.set(k, g));
  }
  const out = [...groups.values()].map((g) => header.map((_, ci) => {
    if (!valueCols.includes(ci)) return g.row[ci];
    const kind = kinds[ci];
    if (!kind) return "";
    const vals = g.rows.map((ri) => cells[ri][ci].value).filter((v) => typeof v === "number");
    return vals.length ? displayText(subtotalOf(kind, vals), fmtOf(ci)) : "";
  }));

  // Values that don't add up: queried at the level shown. If that can't be done, nothing is
  // rolled up (a value is never lost): the visual keeps every level Power BI exported
  const open = valueCols.filter((ci) => !kinds[ci]);
  if (open.length) {
    const levelNames = keyCols.filter((ci) => entries[ci] && HIERARCHY_ROLES.some((r) => (roles[r] || []).includes(entries[ci]))).map((ci) => header[ci]);
    let error = null;
    const keyOk = keyCols.every((ci) => entries[ci] && entries[ci].field && ["column", "level"].includes(entries[ci].field.kind) && entries[ci].field.entity);
    const measures = open.map((ci) => daxValue(entries[ci] && entries[ci].field));
    if (!keyOk || measures.some((m) => !m)) error = "its fields can't be queried";
    // A filter on a measure ("Actives is not blank") only drops rows, and the exported rows
    // already reflect it, so the query leaves it out
    let filterTables = [];
    if (!error) {
      try { filterTables = (filters || []).filter((f) => !(f && f.target && f.target.measure)).map(daxFilter).filter(Boolean); } catch (e) { error = `a ${e.message}`; }
    }
    if (!error) {
      const groupBy = keyCols.map((ci) => daxColumn(entries[ci].field.entity, entries[ci].field.prop));
      const query = `EVALUATE SUMMARIZECOLUMNS(${[...groupBy, ...filterTables, ...measures.map((m, i) => `"v${i}", ${m}`)].join(", ")})`;
      const res = await executeDaxQuery(query);
      let byKey = null, missingIsEmpty = false;
      if (res.error) {
        // No permission to query the model (a Viewer): Power BI's own query service, used only
        // when it reproduces the exported rows exactly
        const detailCols = header.map((_, ci) => ci).filter((ci) => hidden[ci] || keyCols.includes(ci));
        const site = await websiteLevelValues(header, entries, keyCols, detailCols, open, data, cells, filters);
        if (site.byKey) {
          byKey = site.byKey;
          missingIsEmpty = !!site.missingIsEmpty;
          log(`"${item.title}": values at the drill level from Power BI's query service (checked against the export)`);
        } else {
          error = res.error;
          item.siteQueryError = site.error;   // for the visual check
          log(`"${item.title}": Power BI's query service not used: ${site.error}`);
        }
      } else {
        byKey = new Map(res.rows.map((r) => [keyCols.map((ci) => matchKey(r[`${entries[ci].field.entity}[${entries[ci].field.prop}]`])).join("\u0001"), open.map((_, i) => r[`[v${i}]`])]));
      }
      if (byKey) {
        const keys = [...groups.keys()];
        if (!missingIsEmpty && keys.some((k) => !byKey.has(k))) error = "its answer didn't match the exported rows";
        else keys.forEach((k, gi) => {
          const r = byKey.get(k) || [];
          open.forEach((ci, i) => {
            const v = typeof r[i] === "number" ? r[i] : r[i] !== null && r[i] !== undefined && r[i] !== "" && Number.isFinite(Number(r[i])) ? Number(r[i]) : null;
            out[gi][ci] = v !== null ? displayText(v, fmtOf(ci)) : "";
          });
        });
      }
    }
    if (error) {
      item.drillKept = true;   // the layout uses every exported level too
      const per = levelNames.length ? `per ${levelNames.join(" / ")}` : "at its drill level";
      const all = header.filter((_, ci) => hidden[ci] || levelNames.includes(header[ci]));
      const names = open.map((ci) => header[ci]).join(", "), many = open.length > 1;
      return `"${item.title}": shown at every level Power BI exported (${all.join(" › ")}), not ${per} as in the report: ` +
        `${names} ${many ? "don't" : "doesn't"} add up from the exported rows, and querying ${many ? "them" : "it"} ${per} failed (${error}).`;
    }
  }
  item.rows = [header.filter((_, ci) => !hidden[ci]), ...out.map((row) => row.filter((_, ci) => !hidden[ci]))];
  log(`"${item.title}": rolled up to the drill level shown (${header.filter((_, ci) => keyCols.includes(ci)).join(", ")}), ${data.length} → ${out.length} rows`);
  return null;
}

/*
 * The drill state of the report as it is in the app right now: Power BI's embedding API can't
 * report a visual's drill level, but a bookmark captured from the session holds it (each
 * visual's active hierarchy levels, "activeProjections", and a matrix's expand / collapse
 * state, "expansionStates"). Returns { page: { visual: singleVisual state } }, or null.
 */
async function captureLiveDrill(report) {
  try {
    const bm = await report.bookmarksManager.capture({ allPages: true });
    const es = await decodeBookmarkState(bm && bm.state);
    const sections = es && (es.sections || (es.explorationState && es.explorationState.sections));
    if (!sections) { log("Live drill levels not read: the captured view had no pages"); return null; }
    const out = {};
    for (const [page, s] of Object.entries(sections)) {
      out[page] = {};
      for (const [name, vc] of Object.entries((s && s.visualContainers) || {})) {
        if (vc && vc.singleVisual) out[page][name] = vc.singleVisual;
      }
    }
    return out;
  } catch (e) {
    log(`Live drill levels not read (the saved levels are used): ${errorText(e)}`);
    return null;
  }
}

/*
 * Click selections in the report (a bar, a slice, a row clicked: Power BI cross-filters or
 * cross-highlights the other visuals), per page, from the embed's "dataSelected" event: the
 * visual clicked and the values picked, by field. Clicking empty space clears a page's.
 */
const liveSelections = {};
function recordSelection(detail) {
  const page = detail && detail.page && detail.page.name;
  if (!page) return;
  const byField = new Map();
  for (const point of detail.dataPoints || []) {
    for (const id of point.identity || []) {
      const t = id.target || {};
      const field = t.column || t.hierarchyLevel || t.hierarchy || t.measure;
      if (!field || id.equals === undefined || id.equals === null) continue;
      const value = id.equals instanceof Date ? id.equals.toLocaleDateString() : String(id.equals);
      if (!byField.has(field)) byField.set(field, []);
      if (!byField.get(field).includes(value)) byField.get(field).push(value);
    }
  }
  if (!byField.size) { delete liveSelections[page]; return; }
  liveSelections[page] = {
    visual: (detail.visual && (detail.visual.title || detail.visual.name)) || "",
    parts: [...byField].map(([field, values]) => ({ field, values: values.length > 6 ? [...values.slice(0, 6), `+${values.length - 6} more`] : values })),
  };
}
function clearSelections() {
  Object.keys(liveSelections).forEach((k) => delete liveSelections[k]);
}

/* A captured bookmark's state: JSON, or base64 of JSON or of gzipped JSON */
async function decodeBookmarkState(state) {
  if (!state) return null;
  const s = String(state).trim();
  if (s.startsWith("{")) return JSON.parse(s);
  const bytes = Uint8Array.from(atob(s.replace(/-/g, "+").replace(/_/g, "/")), (ch) => ch.charCodeAt(0));
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) {
    const text = await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
    return JSON.parse(text);
  }
  return JSON.parse(new TextDecoder().decode(bytes));
}

/* A visual's report-file info with its live drill state on top: the levels it shows now and
   a matrix's expand / collapse state (the file's saved state where the session has none) */
function withLiveDrill(item, info) {
  const sv = item.liveDrill;
  if (!info || !info.roles || !sv) return info;
  const key = (f) => (f ? `${String(f.entity || "").toLowerCase()}|${String(f.prop || "").toLowerCase()}` : "");
  const roles = { ...info.roles };
  for (const [role, exprs] of Object.entries(sv.activeProjections || {})) {
    const list = roles[role];
    if (!list || !list.length || !Array.isArray(exprs)) continue;
    const live = exprs.map((e) => fieldInfo(e)).filter(Boolean);
    // A level with no query field (some layouts, like the website's, leave the visual's query
    // out) matches by its name
    const isLive = (f) => (f.field
      ? live.some((l) => key(l) === key(f.field) || (!f.field.entity && String(l.prop).toLowerCase() === String(f.field.prop || "").toLowerCase()))
      : !!f.ref && live.some((l) => normalizeField(l.prop) === f.ref));
    if (!list.some(isLive)) continue;   // the fields didn't match: keep the saved levels
    roles[role] = list.map((f) => ({ ...f, active: isLive(f) }));
  }
  return { ...info, roles, expansion: sv.expansionStates || info.expansion };
}

/* For every visual: its rows at the drill level the report shows (needs the report file) */
async function loadDrillLevels(collected, options) {
  for (const { data } of collected) {
    for (const item of data.items) {
      item.matrixGrid = options.matrixLayout === "grid";
      const info = withLiveDrill(item, visualInfo(options.reportColors, item.page, item.name));
      const note = await rollUpToDrillLevel(item, info, [...(data.queryFilters || []), ...(item.visualFilters || [])]);
      if (note) data.notes.push(note);
    }
  }
}

/* A matrix group's saved expand / collapse state: collapsed when its level is collapsed by
   default and the group isn't toggled, or expanded by default and toggled. path: the
   group's labels from the top level */
function groupCollapsed(expansion, role, path) {
  const st = (expansion || []).find((e) => (e.roles || []).includes(role));
  if (!st || !st.levels) return false;
  const lit = (v) => {
    const x = v && v.Literal ? String(v.Literal.Value) : "";
    return matchKey(x.replace(/^datetime'(.*)'$/, "$1").replace(/^'(.*)'$/s, "$1").replace(/^(-?\d+(\.\d+)?)[LDM]$/, "$1").replace(/''/g, "'"));
  };
  let node = st.root || {};
  for (const label of path) {
    node = (node.children || []).find((c) => (c.identityValues || []).length && lit(c.identityValues[c.identityValues.length - 1]) === matchKey(label)) || {};
  }
  const level = st.levels[path.length - 1] || {};
  return !!level.isCollapsed !== !!node.isToggled;
}

/*
 * A matrix as Power BI draws it, from its flat export (needs the report file to know its
 * row, column and value fields):
 *  - column fields: a column per column value (per value and measure, "2024 · Sales" when
 *    there are several measures), then Total columns (live sums across the row) when the
 *    matrix shows column subtotals
 *  - one row field: a row per row value, an Excel Table with a totals row
 *  - several row fields: Power BI's stepped layout (item.stepped): each group's name, its
 *    rows indented, its total row (live SUBTOTALs), Excel outline levels to fold the groups
 *    (the grouping buttons beside the row numbers), and a grand total
 * item.valueCols: each value field's columns (for the report's formatting); item.columnKinds:
 * how each column totals.
 */
function matrixLayout(item, info) {
  const roles = (info && info.roles) || {};
  const rowF = roles.Rows || [], colF = roles.Columns || [], valF = roles.Values || [];
  if (!rowF.length || !valF.length) return;
  const { header, data } = classifyColumns(item.rows);
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const norm = header.map((h) => normalizeField(h));
  const taken = new Set();
  const colOf = (f) => {
    let ci = f.label ? lower.findIndex((h, c) => !taken.has(c) && h === f.label) : -1;
    if (ci < 0) ci = norm.findIndex((h, c) => !taken.has(c) && h === f.ref);
    if (ci >= 0) taken.add(ci);
    return ci;
  };
  const rc = rowF.map(colOf), cc = colF.map(colOf), vc = valF.map(colOf);
  if (rc.some((c) => c < 0) || cc.some((c) => c < 0) || vc.every((c) => c < 0)) return;
  if (!cc.length && rc.length < 2) return;   // a plain list: stays a table (totals row only)
  const b = info.builtin || {};
  const text = (row, c) => String(row[c] ?? "").trim();
  const cells = data.map((row) => header.map((_, ci) => toCellValue(row[ci])));
  const formats = readableFormats(header, cells);
  const measures = valF.map((f, k) => {
    const ci = vc[k];
    if (ci < 0) return null;
    const fmt = formats[ci] || (cells.find((r) => typeof r[ci].value === "number") || [])[ci]?.numFmt;
    return { k, ci, name: header[ci], fmt, kind: totalKind(f.field, header[ci], fmt) };
  }).filter(Boolean);

  // Column values in Power BI's order (a value missing from the first rows still lands in place)
  let keys = [null];
  const keyOf = (row) => cc.map((c) => text(row, c) || "(Blank)").join(" · ");
  if (cc.length) {
    const tagged = data.map((row) => [...row, keyOf(row)]);
    keys = seriesOrder(tagged, header.length, rc);
  }
  // Leaf rows: one per row value(s), in the export's order
  const leaves = new Map();
  for (const row of data) {
    const labels = rc.map((c) => text(row, c) || "(Blank)");
    const id = labels.join("\u0001");
    if (!leaves.has(id)) leaves.set(id, { labels, at: {} });
    leaves.get(id).at[cc.length ? keyOf(row) : "\u0000"] = row;
  }

  // Columns: values per column value and measure, then the Total columns
  const several = measures.length > 1;
  const valueCols = valF.map(() => []);
  const columns = [];   // { head, m, key } or { head, m, total: true }
  for (const key of keys) {
    for (const m of measures) {
      columns.push({ head: key === null ? m.name : several ? `${key} · ${m.name}` : key, m, key });
    }
  }
  if (cc.length && b.columnSubtotals !== false) {
    for (const m of measures) if (m.kind) columns.push({ head: several ? `Total · ${m.name}` : "Total", m, total: true });
  }
  const valueOf = (leaf, c) => {
    if (c.total) {
      const vals = columns.filter((x) => !x.total && x.m === c.m).map((x) => toCellValue(text(leaf.at[x.key] || [], x.m.ci)).value);
      return vals.some((v) => typeof v === "number") ? String(subtotalOf(c.m.kind, vals)) : "";
    }
    const row = leaf.at[c.key === null ? "\u0000" : c.key];
    return row ? row[c.m.ci] ?? "" : "";
  };

  const labelHead = rc.length > 1 ? rc.map((c) => header[c]).join(" › ") : header[rc[0]];
  const leafList = [...leaves.values()];
  columns.forEach((c, i) => valueCols[c.m.k].push(i + 1));
  item.valueCols = valueCols;
  item.columnKinds = [null, ...columns.map((c) => c.m.kind)];
  // Total columns sum across the row: { col index: [the measure's value columns] }
  item.rowTotals = {};
  columns.forEach((c, i) => {
    if (c.total) item.rowTotals[i + 1] = { kind: c.m.kind, of: columns.map((x, j) => (!x.total && x.m === c.m ? j + 1 : -1)).filter((j) => j > 0) };
  });
  // Each value column in its measure's format: a readable one, else the export's own ($, %)
  item.cellFormats = [null, ...columns.map((c) => c.m.fmt || null)];

  if (rc.length === 1) {
    item.rows = [[labelHead, ...columns.map((c) => c.head)], ...leafList.map((l) => [l.labels[0], ...columns.map((c) => valueOf(l, c))])];
    return;
  }

  // Stepped: group names, indented rows, group totals; levels for Excel's outline
  const levels = rc.length;
  const body = [];   // { type: "group" | "leaf" | "total", depth, label, values, from, to (body indices a total covers) }
  // A group the report shows collapsed is folded in Excel's outline (its total row stays in
  // view, as Power BI shows the collapsed group's row); the [+] beside it opens it
  const walk = (list, depth, path) => {
    const groups = new Map();
    for (const l of list) {
      const g = l.labels[depth];
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(l);
    }
    for (const [label, members] of groups) {
      if (depth === levels - 1) {
        for (const l of members) body.push({ type: "leaf", depth, label, values: columns.map((c) => valueOf(l, c)) });
        continue;
      }
      const collapsed = groupCollapsed(info.expansion, "Rows", [...path, label]);
      const head = body.length;
      body.push({ type: "group", depth, label, values: columns.map(() => ""), collapsed });
      const from = body.length;
      walk(members, depth + 1, [...path, label]);
      if (b.rowSubtotals !== false) {
        body.push({ type: "total", depth, label: collapsed ? label : `${label} total`, values: columns.map(() => ""), from, to: body.length - 1 });
      }
      if (collapsed) body[head].fold = { from: b.rowSubtotals !== false ? head : from, to: body.length - (b.rowSubtotals !== false ? 2 : 1) };
    }
  };
  walk(leafList, 0, []);
  item.stepped = { levels, body, subtotals: b.rowSubtotals !== false };
  item.rows = [[labelHead, ...columns.map((c) => c.head)], ...body.map((x) => [x.label, ...x.values])];
}

/*
 * The new card draws one card per measure in its Data field well, each with its own reference
 * labels beneath (fields set in the card's formatting, e.g. "vs last year +4.2%"). Power BI
 * exports them all in one row, so the row is split into one tile per measure, side by side in
 * the visual's place, each with its reference labels as sub-rows (the visual's title goes on
 * the first). Without the report file, every number column is a card.
 */
/*
 * A key influencers visual as Power BI shows it: what makes the analysed field (Status) take
 * its value (Won). Power BI exports only the rows behind it, so each influencer is worked out
 * from the rows counted per value (loadKeyInfluencers) the way Power BI states it: the share
 * with that value where the field has this value, over the share among all other rows
 * ("1.4x more likely").
 * Influencers whose difference could be chance (a two-proportion test) or that rest on few
 * rows are left out. Returns the insight table where the visual sits; or nothing (with a note)
 * when it can't be worked out.
 */
const KI_MIN_ROWS = 10;      // rows a value needs on each side of the comparison
const KI_Z = 1.96;           // 95% confidence
const KI_TOP = 10;           // influencers listed
function keyInfluencerItems(item, info, notes) {
  if (item.type !== "keyDriversVisual") return [item];
  const name = item.title && !isGenericTitle(item.title, item.type) ? item.title : friendlyType(item.type);
  const skip = (why) => { notes.push(`Not included: "${name}" – ${why || ON_SCREEN_ONLY.keyDriversVisual}`); return []; };
  const kd = info && info.keyDrivers;
  if (!kd || !kd.target || kd.targetValue === null || kd.analysis !== "keyInfluencers") return skip();
  // Rows counted per value by a query (loadKeyInfluencers): the exported rows are merged ones
  if (!item.kiCounts) return skip(item.kiError ? `${item.kiError}. View it in the report.` : null);
  const header = item.rows[0] || [];
  const nameOf = (f) => {
    const ci = header.findIndex((h) => (f.label && String(h).trim().toLowerCase() === f.label) || normalizeField(h) === f.ref);
    return ci >= 0 ? header[ci] : (f.field && f.field.prop) || f.ref;
  };
  const goal = String(kd.targetValue).trim().toLowerCase();
  const isGoal = (t) => String(t ?? "").trim().toLowerCase() === goal;
  const targetName = nameOf(kd.target);
  let targetValue = kd.targetValue;
  const found = [];
  for (const fc of item.kiCounts.fields) {
    const groups = new Map();
    let N = 0, T = 0;
    for (const [value, target, rows] of fc.counts) {
      const n = Number(rows) || 0;
      if (!n) continue;
      N += n;
      if (isGoal(target)) { T += n; targetValue = String(target).trim(); }
      const v = String(value ?? "").trim();
      if (!v) continue;   // blank values count in the totals, not as an influencer
      const g = groups.get(v) || { n: 0, t: 0 };
      g.n += n; if (isGoal(target)) g.t += n;
      groups.set(v, g);
    }
    if (!T || T === N) continue;
    const field = nameOf(fc);
    const p = T / N;
    for (const [value, g] of groups) {
      const nr = N - g.n, tr = T - g.t;
      if (g.n < KI_MIN_ROWS || nr < KI_MIN_ROWS || !tr) continue;
      const pv = g.t / g.n, pr = tr / nr;
      const se = Math.sqrt(p * (1 - p) * (1 / g.n + 1 / nr));
      if (pv <= pr || !se || (pv - pr) / se < KI_Z) continue;
      found.push({ field, value, pv, pr, ratio: pv / pr });
    }
  }
  const numberFields = kd.explainBy.filter((f) => !f.field || f.field.kind !== "column").map(nameOf);
  if (!found.length) return skip(`none of its influencers could be confirmed from the rows behind it. View it in the report.`);
  found.sort((a, b) => b.ratio - a.ratio);
  const top = found.slice(0, KI_TOP);
  const pct = (x) => `${(x * 100).toFixed(1)}%`;
  const shareHead = `${targetName} is ${targetValue}`;
  // The list alone, where the visual sits (its detail chart depends on the influencer picked on
  // screen). Only the shares, which are exact counts: Power BI's "x more likely" comes from its
  // own model of all the fields together, which no count reproduces
  const table = {
    ...item,
    name: `${item.name}~influencers`,
    title: `What influences ${targetName} to be ${targetValue}?`,
    type: "keyInfluencers",   // a plain table: no totals row (times more likely don't add up)
    captionText: `Exact shares from the report's data · strongest influence first`,
    rows: [["When", "Is", shareHead, `${shareHead} (others)`],
      ...top.map((f) => [f.field, f.value, pct(f.pv), pct(f.pr)])],
    gridFormat: { measures: [], tableFormats: [shareHead, `${shareHead} (others)`].map((h, i) => ({ target: "m:" + normalizeField(h), kind: "bar",
      bar: { color: i ? "#C8C6C4" : DEFAULT_THEME_COLORS[0], negative: null, axis: null, reverse: false, showValue: true, share: true } })) },
  };
  if (numberFields.length) notes.push(`"${name}": ${numberFields.join(", ")} (number fields) are explained in the report only.`);
  return [table];
}

/*
 * A decomposition tree as the report shows it: the total, then each level the report has open
 * (Product category, then Product, then Region), with only the branches opened in the report
 * expanded (Software, then Design app), largest first and up to its bars per level (the rest
 * added up in one row). Power BI exports every combination of the fields, so each value is
 * added up from those rows: only for a sum or a count, which add up exactly. Otherwise, or
 * without the report's layout, the visual stays the table it was.
 */
const TREE_INDENT = "  ";   // two em spaces per level
/* A decomposition tree's open levels, in order (the report's expansion state): { state, levelFields } */
function treeLevels(info) {
  const state = ((info.expansion || []).find((s) => (s.roles || []).includes("ExplainBy"))) || null;
  const explain = info.roles.ExplainBy || [];
  const byQuery = (q) => explain.find((f) => f.queryRef === q);
  const levelFields = state ? (state.levels || []).filter((l) => l.isPinned).map((l) => byQuery((l.queryRefs || [])[0])) : [];
  return { state, levelFields: levelFields.some((f) => !f) ? [null] : levelFields };
}
const joinNames = (n) => (n.length <= 1 ? n.join("") : `${n.slice(0, -1).join(", ")} and ${n[n.length - 1]}`);
function decompositionTreeItem(item, info) {
  if (item.type !== "decompositionTreeVisual" || !info || !info.roles) return item;
  const [header, ...data] = item.rows;
  const colOf = (f) => (f ? header.findIndex((h) => (f.label && String(h).trim().toLowerCase() === f.label) || normalizeField(h) === f.ref) : -1);
  const measure = (info.roles.Analyze || [])[0];
  const mc = colOf(measure);
  const additive = measure && measure.field && measure.field.kind === "agg" && [0, 5].includes(measure.field.agg);   // Sum, Count
  // A measure that doesn't add up (Var Plan): its values as Power BI worked them out
  // (loadTreeValues), else the visual stays the table it was
  const asked = !additive && item.treeValues ? item.treeValues : null;
  if (mc < 0 || (!additive && !asked)) return item;
  // The open levels, in order, and the opened branch at each (the report's expansion state)
  const { state, levelFields } = treeLevels(info);
  const levels = levelFields.map(colOf);
  if (!levels.length || levels.some((ci) => ci < 0)) return item;
  const literal = (v) => { const s = String(((v || {}).Literal || {}).Value ?? ""); return /^'.*'$/s.test(s) ? s.slice(1, -1).replace(/''/g, "'") : s; };
  const text = (r, ci) => String(r[ci] ?? "").trim();
  const num = (r) => { const v = toCellValue(text(r, mc)).value; return typeof v === "number" ? v : 0; };
  // The measure as exported ("$1,234.5"): sums keep its currency sign and decimals
  const sample = text(data.find((r) => text(r, mc)) || [], mc);
  const currency = (sample.match(/^-?\s*([$€£₹¥])/) || [])[1] || "";
  const decimals = ((sample.replace(/[^\d.]/g, "").split(".")[1]) || "").length;
  const show = (n) => `${n < 0 ? "-" : ""}${currency}${Math.abs(n).toFixed(Math.min(decimals, 4))}`;
  const pct = (x) => `${(x * 100).toFixed(1)}%`;
  const limit = Math.max(1, info.barsPerLevel || 10);
  const out = [];
  const total = asked ? asked.total ?? 0 : data.reduce((s, r) => s + num(r), 0);
  out.push(["Total", "", show(total), "100.0%"]);
  // One level under a branch: its values largest first, the opened one expanded beneath it
  const walk = (rows, depth, node, parentTotal, path = []) => {
    const ci = levels[depth];
    const groups = new Map();
    for (const r of rows) { const k = text(r, ci) || "(Blank)"; groups.set(k, (groups.get(k) || 0) + num(r)); }
    if (asked) for (const k of groups.keys()) groups.set(k, asked.byPath[[...path, k].map(matchKey).join("\u0001")] ?? 0);
    const sorted = [...groups].sort((a, b) => b[1] - a[1]);
    const open = new Set(((node && node.children) || []).filter((c) => c.isToggled).map((c) => literal((c.identityValues || [])[0])));
    // values asked of Power BI don't add up, so none are put together: every bar is listed
    const shown = asked ? sorted.length : limit;
    sorted.slice(0, shown).forEach(([name, value]) => {
      out.push([TREE_INDENT.repeat(depth) + header[ci], name, show(value), pct(parentTotal ? value / parentTotal : 0)]);
      if (open.has(name) && depth + 1 < levels.length) {
        const child = node.children.find((c) => c.isToggled && literal((c.identityValues || [])[0]) === name);
        walk(rows.filter((r) => (text(r, ci) || "(Blank)") === name), depth + 1, child, value, [...path, name]);
      }
    });
    const rest = sorted.slice(shown);
    if (rest.length) {
      const value = rest.reduce((s, [, v]) => s + v, 0);
      out.push([TREE_INDENT.repeat(depth) + header[ci], `${rest.length} more`, show(value), pct(parentTotal ? value / parentTotal : 0)]);
    }
  };
  walk(data, 0, state.root, total);
  const path = [];
  for (let node = state.root; node; ) {
    const next = (node.children || []).find((c) => c.isToggled);
    if (next) path.push(literal((next.identityValues || [])[0]));
    node = next;
  }
  const shareHead = "Share of parent";
  return {
    ...item,
    type: "decompositionTree",   // a plain table: its own Total row, no totals row added
    title: isGenericTitle(item.title, item.type) ? `${header[mc]} by ${joinNames(levels.map((ci) => header[ci]))}` : item.title,
    captionText: `As opened in the report${path.length ? `: ${path.join(" › ")}` : ""} · ${asked ? "values from Power BI" : "values added up from the report's data"}`,
    rows: [["Level", "Name", header[mc], shareHead], ...out],
    gridFormat: { measures: [], tableFormats: [{ target: "m:" + normalizeField(shareHead), kind: "bar",
      bar: { color: DEFAULT_THEME_COLORS[0], negative: null, axis: null, reverse: false, showValue: true, share: true } }] },
  };
}

/*
 * A chart with small multiples (a field in Power BI's "Small multiples" well, stored as Rows):
 * one small chart per value of that field, laid out in a grid under the visual's title as
 * Power BI lays them out, with the visual's look and one value axis shared by all, joined into
 * one Excel object (addChartsToFile). Charts mode only; the data table keeps every row. Without
 * the report's layout, with one value, or with more panels than fit, the chart stays as it was.
 */
const SMALL_MULTIPLES_MAX = 12;
function smallMultipleItems(item, info) {
  const spec = chartSpecOf(item.type);
  if (!spec || item.isKpi || !info || !info.roles || !(info.roles.Rows || []).length) return [item];
  if (!["bar", "line", "area", "combo"].includes(spec.kind)) return [item];
  const [header, ...data] = item.rows;
  const f = info.roles.Rows[0];
  const rc = header.findIndex((h) => (f.label && String(h).trim().toLowerCase() === f.label) || normalizeField(h) === f.ref);
  if (rc < 0) return [item];
  const key = (r) => String(r[rc] ?? "").trim() || "(Blank)";
  const exported = [...new Set(data.map(key))];
  if (exported.length < 2 || exported.length > SMALL_MULTIPLES_MAX) return [item];
  const drop = (r) => r.filter((_, ci) => ci !== rc);
  // In the report's order: the panels by their field, the bars in every panel in one order
  const { panelOrder, rowOrder } = smallMultipleOrder(item, info, header, data, rc);
  const values = panelOrder(exported);
  const panelRows = (v) => rowOrder(data.filter((r) => key(r) === v));
  // One value axis for all panels, as Power BI shares it (bars from zero, no negatives)
  const nums = data.flatMap((r) => drop(r).map((x) => toCellValue(String(x ?? "")).value)).filter((v) => typeof v === "number");
  const sharedMax = nums.length && nums.every((v) => v >= 0) ? Math.max(...nums) : null;
  const cols = Math.ceil(Math.sqrt(values.length));
  return [{
    ...item,
    panels: values.map((v) => ({ value: v, rows: [drop(header), ...panelRows(v).map(drop)] })),
    panelGrid: { cols, rows: Math.ceil(values.length / cols) },
    sharedMax,
  }];
}

/*
 * The order Power BI draws small multiples in. Panels: by their field, A to Z (numbers by
 * value, months and dates in calendar order) unless the visual is sorted by that field the other
 * way. Bars: one order shared by every panel (Power BI shares the axis), from the visual's sort
 * as shown in the session, else as saved: by the axis field's values, or by a measure's total
 * across all panels; with no sort, the rows keep the order Power BI exported them in.
 */
function smallMultipleOrder(item, info, header, data, rc) {
  const live = item.liveDrill && item.liveDrill.orderBy ? sortOf(item.liveDrill.orderBy) : null;
  const sort = live || info.sort || [];
  const colOf = (f) => header.findIndex((h, ci) => ci !== rc && ((f.label && String(h).trim().toLowerCase() === f.label) || normalizeField(h) === f.ref));
  const same = (a, b) => !!a && !!b && String(a.prop || "").toLowerCase() === String(b.prop || "").toLowerCase() &&
    (!a.entity || !b.entity || String(a.entity).toLowerCase() === String(b.entity).toLowerCase());
  const matches = (e, f) => same(e.field, f.field) || (!f.field && normalizeField(e.field.prop) === f.ref);
  const panelSort = sort.find((e) => matches(e, info.roles.Rows[0]));
  const panelOrder = (values) => naturalOrder(values, !!(panelSort && panelSort.desc));
  const keep = { panelOrder, rowOrder: (rows) => rows };
  // the axis: the first sort entry on one of its fields or on a measure shown
  const catFields = (info.roles.Category || info.roles.Axis || []).filter((f) => colOf(f) >= 0);
  const catCols = catFields.map(colOf);
  if (!catCols.length || !sort.length) return keep;
  const text = (r, c) => String(r[c] ?? "").trim();
  const catKey = (r) => catCols.map((c) => text(r, c)).join("\u0001");
  const measures = ["Y", "Values", "Value", "Y2"].flatMap((role) => info.roles[role] || []).filter((f) => colOf(f) >= 0);
  let rank = null;
  for (const e of sort) {
    const onCat = catFields.findIndex((f) => matches(e, f));
    if (onCat >= 0) {
      const c = catCols[onCat];
      const pos = new Map(naturalOrder([...new Set(data.map((r) => text(r, c)))], e.desc).map((v, i) => [v, i]));
      rank = (r) => pos.get(text(r, c)) ?? 1e9;
      break;
    }
    const m = measures.find((f) => matches(e, f));
    if (m) {
      const mc = colOf(m), totals = new Map();
      for (const r of data) {
        const v = toCellValue(text(r, mc)).value;
        totals.set(catKey(r), (totals.get(catKey(r)) || 0) + (typeof v === "number" ? v : 0));
      }
      const order = [...totals.keys()].map((k, i) => [k, i])
        .sort((a, b) => (e.desc ? totals.get(b[0]) - totals.get(a[0]) : totals.get(a[0]) - totals.get(b[0])) || a[1] - b[1]).map(([k]) => k);
      const pos = new Map(order.map((k, i) => [k, i]));
      rank = (r) => pos.get(catKey(r)) ?? 1e9;
      break;
    }
  }
  if (!rank) return keep;
  if (typeof log === "function") {
    log(`"${item.title}": small multiples sorted as ${live ? "shown in the report now" : "saved in the report"}: ` +
      sort.map((e) => `${e.field.prop} ${e.desc ? "descending" : "ascending"}`).join(", "));
  }
  return { panelOrder, rowOrder: (rows) => rows.map((r, i) => [r, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(([r]) => r) };
}

/* Values in Power BI's natural order: numbers by value, months and dates in calendar order, text A to Z */
function naturalOrder(values, desc = false) {
  const t = values.map((v) => String(v ?? "").trim());
  const num = (v) => Number(v.replace(/[,$%]/g, ""));
  const month = (v) => MONTHS.indexOf(v.slice(0, 3).toLowerCase());
  const cmp = t.every((v) => v !== "" && isFinite(num(v))) ? (a, b) => num(a) - num(b)
    : t.every((v) => /^[a-z]+$/i.test(v) && month(v) >= 0) ? (a, b) => month(a) - month(b)
      : t.every((v) => parseDate(v)) ? (a, b) => parseDate(a).date - parseDate(b).date
        : (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  const order = values.map((_, i) => i).sort((a, b) => cmp(t[a], t[b]) || a - b);
  if (desc) order.reverse();
  return order.map((i) => values[i]);
}

function splitCards(item, info) {
  // A multi-row card with one row shows each field side by side, its value over its name: a
  // tile per field, across the visual's place (a card with several rows stays a table)
  if (item.type === "multiRowCard" && item.isKpi && item.rows.length === 2 && item.rows[0].length >= 2) {
    const [header, values] = item.rows;
    const { x, y, width, height } = item.layout || { x: 0, y: 0, width: 0, height: 0 };
    return header.map((h, k) => ({
      ...item, title: k === 0 ? item.title : "",
      rows: [[h], [values[k]]],
      layout: { ...item.layout, x: x + (width / header.length) * k, y, width: width / header.length, height },
    }));
  }
  if (item.type !== "cardVisual" || item.rows.length !== 2 || item.rows[0].length < 2) return [item];
  const [header, values] = item.rows;
  const data = (info && info.roles && info.roles.Data) || [];
  const refLabels = (info && info.builtin && info.builtin.refLabels) || [];
  const owners = header.map(() => -1);   // per column: the card it belongs to (-1: a card of its own)
  let cards;
  if (data.length) {
    const found = fieldsByColumn(header, data);
    cards = found.map((e, ci) => (e ? ci : -1)).filter((ci) => ci >= 0);
    if (cards.length < 1) return [item];
    header.forEach((h, ci) => {
      if (cards.includes(ci)) return;
      const rl = refLabels.find((x) => x.ref === normalizeField(h));
      const ownerField = rl && rl.owner >= 0 ? data[rl.owner] : null;
      const at = ownerField ? found.indexOf(ownerField) : -1;
      owners[ci] = at >= 0 ? at : cards[0];
    });
  } else {
    cards = header.map((_, ci) => ci).filter((ci) => typeof toCellValue(values[ci]).value === "number");
    if (cards.length < 2) return [item];
  }
  if (cards.length === 1) {
    // one card: its measure is the large value, its reference labels beneath
    const order = [cards[0], ...header.map((_, ci) => ci).filter((ci) => ci !== cards[0])];
    return [{ ...item, rows: [order.map((ci) => header[ci]), order.map((ci) => values[ci])] }];
  }
  const { x, y, width, height } = item.layout || { x: 0, y: 0, width: 0, height: 0 };
  return cards.map((ci, k) => {
    const cols = [ci, ...header.map((_, c) => c).filter((c) => owners[c] === ci)];
    return {
      ...item, title: k === 0 ? item.title : "", isKpi: true,
      rows: [cols.map((c) => header[c]), cols.map((c) => values[c])],
      layout: { ...item.layout, x: x + (width / cards.length) * k, y, width: width / cards.length, height },
    };
  });
}

/* A table's sparklines (points read by querySparklines) as empty columns at their place in
   the table, where the Excel sparklines go (item.sparkCols) */
function addSparklineColumns(item) {
  const cols = [];
  for (const sp of [...item.sparklines].sort((a, b) => a.index - b.index)) {
    const at = Math.max(0, Math.min(sp.index, item.rows[0].length));
    item.rows = item.rows.map((row, ri) => [...row.slice(0, at), ri === 0 ? sp.name : "", ...row.slice(at)]);
    cols.forEach((c) => {
      if (c.ci >= at) c.ci++;
      c.keyCols = c.keyCols.map((k) => (k >= at ? k + 1 : k));
    });
    cols.push({ ci: at, axis: sp.axis, series: sp.series, keyCols: sp.keyCols.map((k) => (k >= at ? k + 1 : k)) });
  }
  item.sparkCols = cols;
}

/*
 * The page's hidden sheet of sparkline points. Excel sparklines read fixed ranges and stay
 * in their cells when a table is sorted, so each table row's line reads a "live" row that
 * looks its points up by the row's key cells (INDEX / MATCH): after a sort, every line
 * still shows its own row's trend.
 * add(axis, rows: [{ key, keyExpr, points }]) → each table row's range of live points.
 */
function makeTrendSheet(name) {
  const blocks = [];
  let next = 1;
  return {
    name, blocks,
    add(axis, rows) {
      const top = next, v0 = top + 1 + rows.length;
      blocks.push({ top, axis, rows });
      next = v0 + rows.length + 1;
      const last = columnLetter(1 + axis.length);
      return rows.map((_, i) => `'${name}'!B${v0 + i}:${last}${v0 + i}`);
    },
  };
}

/* A key cell's value as Excel's "&" turns it into text: numbers and dates (serials) plainly */
function keyText(v) {
  if (v instanceof Date) return String(Math.round((v.getTime() / 86400000 + 25569) * 1e6) / 1e6);
  if (v === null || v === undefined) return "";
  return String(v);
}
function columnLetter(n) {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
function writeTrendSheet(wb, sheet) {
  const ws = wb.addWorksheet(sheet.name, { views: [{ showGridLines: false }] });
  ws.state = "veryHidden";   // Excel's sparklines read it; users see only the page
  for (const { top, axis, rows } of sheet.blocks) {
    // The axis, then each row's key and points as exported, then the live rows
    ws.getCell(top, 1).value = "Key";
    axis.forEach((a, k) => { ws.getCell(top, 2 + k).value = a instanceof Date ? a : String(a ?? ""); });
    const s0 = top + 1, s1 = top + rows.length, v0 = s1 + 1;
    const last = columnLetter(1 + axis.length);
    rows.forEach((r, i) => {
      ws.getCell(s0 + i, 1).value = r.key;
      r.points.forEach((v, k) => { ws.getCell(s0 + i, 2 + k).value = typeof v === "number" ? v : null; });
    });
    rows.forEach((r, i) => {
      const vr = v0 + i;
      ws.getCell(vr, 1).value = { formula: r.keyExpr, result: r.key };
      axis.forEach((_, k) => {
        const v = r.points[k];
        ws.getCell(vr, 2 + k).value = {
          formula: `IFERROR(INDEX($B$${s0}:$${last}$${s1},MATCH($A$${vr},$A$${s0}:$A$${s1},0),${k + 1}),"")`,
          result: typeof v === "number" ? v : "",
        };
      });
    });
  }
}

/* KPI tiles show a title row only when the title adds something beyond the field name */
function kpiHasTitle(item) {
  if (item.type === "kpi") return false;   // the KPI visual's title is its tile's label (writeKpiBlock)
  return !!item.displayTitle && item.displayTitle.toLowerCase() !== String(item.rows[0][0] || "").toLowerCase();
}

/* Rows an exhibit takes above its data rows: title, caption, [trend cards and a blank row], header */
function headRows(item) {
  return HEAD_ROWS + (item.trend ? item.trend.cards.length * TREND_ROWS + 1 : 0);
}

/*
 * Size of a visual's block in rows / columns.
 *  - KPI tile: [title] + label + value (2 rows) + [exact value] + one row per extra field
 *  - exhibit: title + caption + [trend cards] + header + data rows + closing row (the closing
 *    rule, which doubles as the "rows grouped" note when collapsing is on)
 */
function blockSize(item, options) {
  if (item.chart) return chartBlockSize(item);   // charts mode: the chart's own size
  if (item.isKpi) {
    const [header, values = []] = item.rows;
    const hero = toCellValue(values[0]);
    const fmt = readableFormats([header[0]], [[hero]])[0] || hero.numFmt;
    const compact = options.tableNumbers !== "report" && !!compactFormat(hero.value, fmt);
    const height = (kpiHasTitle(item) ? 1 : 0) + 3 + (compact ? 1 : 0) + (header.length - 1);
    return { width: 2, height, note: false };
  }
  const n = item.rows.length - 1;
  // A stepped matrix folds by its own groups; a totals row puts the note one row lower
  const note = !item.stepped && !!(options.collapse && n > options.visibleRows);
  return { width: item.rows[0].length, height: headRows(item) + n + 1 + (item.totals && note ? 1 : 0), note };
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
  const info = visualInfo(lay.reportColors, item.page, item.name);
  const valueColor = (i) => {
    const vc = (info && info.valueColors) || {}, vr = (info && info.valueRules) || {};
    const key = "m:" + normalizeField(header[i]);
    let hex = vc[key] || vc.all;
    const rule = vr[key] || vr.all;
    if (!hex && rule) {
      // the rule reads this card's own fields
      const column = (name) => {
        const k = header.findIndex((h) => normalizeField(h) === normalizeField(name));
        return k >= 0 ? [cells[k].value ?? null] : null;
      };
      hex = evaluateColorRule(rule, column, 1).colors[0];
    }
    return hex ? toArgb(hex) : null;
  };
  // "Chart background: As in the report": the tile in the card's own background, its text
  // readable on it; otherwise the tile's usual light fill
  const frame = lay.reportBackground && info && info.look && info.look.frame;
  const ownBg = frame && frame.background && (frame.transparency || 0) < 100 ? toArgb(frame.background) : null;
  const tileFill = ownBg ? solidFill(ownBg) : SURFACE_FILL;
  const onTile = (font) => (ownBg && luminance(ownBg) < 0.5 ? { ...font, color: { argb: C.WHITE } } : font);
  const tinted = (font, i) => (valueColor(i) ? { ...onTile(font), color: { argb: valueColor(i) } } : onTile(font));
  // The cap takes the card's own colour from the report (its accent bar, else its value
  // colour), so the tile reads as one unit; otherwise the brand accent
  const fieldValues = (name) => {
    const k = header.findIndex((h) => normalizeField(h) === normalizeField(name));
    return k >= 0 ? [cells[k].value ?? null] : null;
  };
  const accentBar = info && info.accent
    ? (info.accent.color || evaluateColorRule(info.accent.rule, fieldValues, 1).colors[0]) : null;
  const ownAccent = customSetting(item.type, info, "accent");   // e.g. a progress icon's fill colour
  const capSource = accentBar ? toArgb(accentBar) : ownAccent ? toArgb(ownAccent) : valueColor(0);
  const cap = capSource ? readableOnWhite(capSource) : lay.accent.cap;

  let top = r;
  if (kpiHasTitle(item)) {
    const t = ws.getCell(r, col);
    t.value = item.displayTitle;
    t.font = fonts.exTitle;
    t.alignment = { vertical: "bottom" };
    widths.span(col, col + 1, cellUnits(item.displayTitle, fonts.exTitle));
    top = r + 1;
  } else if (strip && strip.titleRow) {
    top = r + 1;
  }

  // Merges keep each cell's own style, so the cap / rule borders survive on every cell. The KPI
  // visual shows its title over the value (its own, else the measure's name; none when hidden),
  // not the measure's name as a card does
  const kpiVisual = item.type === "kpi";
  const kpiTitle = info && info.titleHidden ? "" : String(item.displayTitle || header[0] || "");
  const labelText = kpiVisual ? kpiTitle : String(header[0] ?? "").toLocaleUpperCase();
  ws.mergeCellsWithoutStyle(top, col, top, col + 1);
  const label = ws.getCell(top, col);
  label.value = labelText;
  label.font = onTile(fonts.kpiLabel);
  // Shrunk to fit when Excel draws it wider than measured, so a long label is never cut off
  label.alignment = { horizontal: "left", vertical: "bottom", indent: 1, shrinkToFit: true };

  const hero = cells[0];
  const report = lay.reportNumbers && reportValueFormat(item, info, hero.value, fmtOf(0));
  const heroFmt = report || fmtOf(0);
  const compact = report ? null : compactFormat(hero.value, heroFmt);
  ws.mergeCellsWithoutStyle(top + 1, col, top + 2, col + 1);
  const v = ws.getCell(top + 1, col);
  const heroNone = hero.value === null || hero.value === undefined || hero.value === "";
  v.value = heroNone ? "(Blank)" : hero.value;
  if ((compact || heroFmt) && !heroNone) v.numFmt = compact || heroFmt;
  v.font = tinted(fonts.kpiValue, 0);
  v.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

  let next = top + 3;
  if (compact) {
    ws.mergeCellsWithoutStyle(next, col, next, col + 1);
    const e = ws.getCell(next, col);
    e.value = hero.value;
    e.numFmt = heroFmt || "#,##0.00";
    e.font = onTile(fonts.kpiExact);
    e.alignment = { horizontal: "left", vertical: "top", indent: 1 };
    next++;
  }
  for (let i = 1; i < header.length; i++, next++) {
    const l = ws.getCell(next, col);
    l.value = header[i];
    l.font = onTile(fonts.kpiSubLabel);
    l.alignment = { vertical: "middle", indent: 1 };
    const sv = ws.getCell(next, col + 1);
    // An empty value reads "(Blank)", as on the card in Power BI
    const none = cells[i].value === null || cells[i].value === undefined || cells[i].value === "";
    sv.value = none ? "(Blank)" : cells[i].value;
    if (fmtOf(i) && !none) sv.numFmt = fmtOf(i);
    sv.font = tinted(fonts.kpiSubValue, i);
    sv.alignment = { horizontal: "right", vertical: "middle", indent: 1 };
    widths.want(col, Math.min(60, cellUnits(header[i], fonts.kpiSubLabel, 1)));
    widths.want(col + 1, cellUnits(none ? "(Blank)" : displayText(cells[i].value, fmtOf(i)), fonts.kpiSubValue, 1));
  }

  const last = Math.max(next - 1, strip ? r + strip.height - 1 : 0);
  for (let rr = top; rr <= last; rr++) {
    for (let cc = col; cc <= col + 1; cc++) {
      const cell = ws.getCell(rr, cc);
      cell.fill = tileFill;
      const b = {};
      if (rr === top) b.top = { style: "medium", color: { argb: cap } };
      if (rr === last) b.bottom = LIGHT_RULE;
      if (divider && cc === col) b.left = LIGHT_RULE;
      if (b.top || b.bottom || b.left) cell.border = b;
    }
  }

  // The tile is as wide as its widest text: the label, the large value, the exact value
  const heroText = compact ? compactText(hero.value, heroFmt) : displayText(hero.value, heroFmt);
  widths.want(col, 12);
  widths.want(col + 1, 10);
  widths.span(col, col + 1, Math.max(24,
    cellUnits(labelText, fonts.kpiLabel, 1) * 1.1,
    cellUnits(heroText, v.font, 1),
    compact ? cellUnits(displayText(hero.value, heroFmt || "#,##0.00"), fonts.kpiExact, 1) : 0));
}

/* A date as short text, "1/15/2023" (with the time when it has one): sizes columns, labels keys */
function dateText(date, withTime) {
  const t = `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`;
  return withTime ? `${t} ${date.getUTCHours()}:${String(date.getUTCMinutes()).padStart(2, "0")}` : t;
}

/* A colour key's label: a date exported as "2023-01-15 00:00:00" reads "1/15/2023" */
function keyLabel(text) {
  const d = parseDate(String(text ?? "").trim());
  return d ? dateText(d.date, d.hasTime) : text;
}

/* A category on a chart axis: as keyLabel, and "(Blank)" for an empty one, as Power BI labels it
   (an empty label leaves its column or point unnamed on the axis) */
function axisLabel(text) {
  const t = keyLabel(text);
  return t === null || t === undefined || String(t).trim() === "" ? "(Blank)" : t;
}

/* Caption runs: a legend of the chart's colours, or "Table · 24 rows" */
function captionRuns(item, colors, header, fonts, units) {
  if (item.captionText) return [[item.captionText, fonts.caption]];   // an exhibit's own note (key influencers)
  // the visual's own subtitle first, as Power BI shows it under the title
  if (item.subtitle) return [[item.subtitle + "  ·  ", fonts.caption], ...captionRuns({ ...item, subtitle: null }, colors, header, fonts, units)];
  const n = item.rows.length - 1;
  const plain = [[`${friendlyType(item.type)}  ·  ${n} row${n === 1 ? "" : "s"}`, fonts.caption]];
  if (!colors) return plain;

  const legend = (entries) => {
    const runs = [];
    let len = 0, shown = 0;
    for (const e of entries) {
      const add = String(keyLabel(e.label)).length * 0.8 + 4;
      if (shown && len + add > units - 6) break;
      runs.push([KEY_GLYPH + "  ", fonts.dot(e.argb)], [keyLabel(e.label) + "      ", fonts.caption]);
      len += add;
      shown++;
    }
    return { runs, shown };
  };

  if (colors.legend && colors.legend.scale) {
    return [...colors.legend.scale.map((argb) => [KEY_GLYPH, fonts.dot(argb)]), ["  " + colors.legend.label, fonts.caption]];
  }
  if (colors.legend && colors.legend.text) return [[colors.legend.text, fonts.caption]];
  if (colors.legend && colors.legend.entries) {
    const { runs, shown } = legend(colors.legend.entries);
    if (shown < colors.legend.entries.length) runs.push([`+${colors.legend.entries.length - shown} more`, fonts.captionMore]);
    return runs;
  }

  if (colors.rowKey) {
    const entries = [
      ...Object.entries(colors.rowKey.colorOf).map(([label, argb]) => ({ label, argb })),
      ...Object.entries(colors.series).map(([ci, argb]) => ({ label: header[ci], argb })),   // combo line values
    ];
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

/* A trend visual's caption: its title, the days it spans and its dates, and its warnings */
function trendCaptionRuns(trend, fonts) {
  if (trend.captionRuns) return runsOf(trend.captionRuns, fonts);
  const c = trend.caption;
  return [
    c.title ? [c.title + "  ·  ", fonts.caption] : null,
    [`Last ${c.days} day${c.days === 1 ? "" : "s"}`, fonts.caption],
    [`  ·  ${c.from} – ${c.to}`, fonts.caption],
    c.stale !== null ? [`  ·  latest data ${c.stale} days old${c.staleText ? ". " + c.staleText : ""}`, fonts.captionMore] : null,
    c.warning ? [`  ·  ⚠ ${c.warning}`, fonts.kpiChange] : null,
  ];
}

/* Text runs as [text, "plain" | "muted" | "strong", argb?] → rich-text runs */
function runsOf(spec, fonts) {
  return spec.filter(Boolean).map(([text, style, argb]) => {
    const font = style === "strong" ? fonts.kpiChange : style === "muted" ? fonts.captionMore : fonts.caption;
    return [text, argb ? { ...font, color: { argb } } : font];
  });
}

/*
 * Trend cards, one under the other, as wide as the table beneath:
 *   TOTAL SALES
 *   $1,192        ╱╲_╱‾╲_   (an Excel sparkline of the values, over a date axis)
 *   +27.2%  since 1/15/2023
 * The cap takes the chart's colour (the brand accent without the report's colours); the
 * sparkline reads the table's first `trend.points` rows (hr: the table's header row).
 */
function writeTrendCards(ws, r, col, w, item, widths, lay, hr) {
  const fonts = lay.fonts, trend = item.trend;
  const range = (ci) => `'${lay.sheetName}'!${ws.getCell(hr + 1, col + ci).address}:${ws.getCell(hr + trend.points, col + ci).address}`;
  let needLeft = 12, needLabel = 0;
  trend.cards.forEach((card, k) => {
    const a = r + k * TREND_ROWS;
    const color = card.color ? readableOnWhite(toArgb(card.color)) : lay.accent.cap;

    const label = ws.getCell(a, col);
    label.value = String(card.label ?? "").toLocaleUpperCase();
    label.font = fonts.kpiLabel;
    label.alignment = { vertical: "bottom", indent: 1 };

    ws.mergeCellsWithoutStyle(a + 1, col, a + 2, col);
    const v = ws.getCell(a + 1, col);
    v.value = card.value;
    if (card.numFmt) v.numFmt = card.numFmt;
    v.font = card.valueColor ? { ...fonts.kpiValue, color: { argb: readableOnWhite(toArgb(card.valueColor)) } } : fonts.kpiValue;
    v.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

    // The line under the value: the change since a date, or a KPI's goal and distance
    const noteRuns = card.note ? runsOf(card.note, fonts) : [[card.change, fonts.kpiChange], [card.since ? `  since ${card.since}` : "", fonts.caption]];
    const change = ws.getCell(a + 3, col);
    change.value = rich(...noteRuns);
    change.alignment = { vertical: "middle", indent: 1 };

    if (w > 1) {
      ws.mergeCellsWithoutStyle(a + 1, col + 1, a + 3, col + w - 1);
      if (!card.noSpark) {
        (lay.sparklines = lay.sparklines || []).push({
          cell: ws.getCell(a + 1, col + 1).address, values: range(card.col), dates: trend.dateAxis === false ? null : range(trend.dateCol),
          color, min: card.min, max: card.max,
        });
      }
    }

    for (let rr = a; rr < a + TREND_ROWS; rr++) {
      for (let cc = col; cc < col + w; cc++) {
        const cell = ws.getCell(rr, cc);
        cell.fill = SURFACE_FILL;
        if (rr === a) cell.border = { top: { style: "medium", color: { argb: color } } };
        else if (rr === a + TREND_ROWS - 1 && k === trend.cards.length - 1) cell.border = { bottom: LIGHT_RULE };
      }
    }
    needLeft = Math.max(needLeft, cellUnits(card.text, fonts.kpiValue, 1),
      noteRuns.reduce((u, [text, font], i) => u + cellUnits(text, font, i ? 0 : 1), 0));
    needLabel = Math.max(needLabel, cellUnits(label.value, fonts.kpiLabel, 1));
  });
  // The value and change fit their column; the sparkline gets room; the label stays inside the card
  widths.want(col, Math.min(40, needLeft));
  if (w === 2) widths.want(col + 1, 20);
  widths.span(col, col + w - 1, needLabel);
}

/*
 * Exhibit: title row, caption row, then a real Excel Table. "Executive" uses Excel's
 * "None" table style with ink rules and hairlines; a built-in style keeps its own look.
 * Header cells keep exactly the column name (anything else makes Excel repair the file).
 * Style objects are made once per column and shared by every cell.
 */
/* Column widths: header, longest text, widest formatted number; numbers in one table share
   a width so their columns read as a set. indentOf(ri): a stepped row's label indent. */
function tableColumnUnits(header, data, numeric, formats, indentOf) {
  const units = header.map((h, ci) => {
    let u = Math.min(45, Math.ceil(h.length * 1.05) + 5);
    if (numeric[ci]) {
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
      data.forEach((row, ri) => {
        if (row[ci].value != null) longest = Math.max(longest, displayText(row[ci].value, row[ci].numFmt).length + (ci === 0 && indentOf ? indentOf(ri) * 1.3 : 0));
      });
      u = Math.max(u, Math.min(45, Math.ceil(longest * 0.95) + 3));
    }
    return Math.max(u, 10);
  });
  const numericUnits = units.filter((_, ci) => numeric[ci]);
  if (numericUnits.length > 1) {
    const target = Math.min(Math.max(...numericUnits), 18);
    units.forEach((u, ci) => { if (numeric[ci]) units[ci] = Math.max(u, target); });
  }
  return units;
}

/* An exhibit's title and caption rows, and its columns' widths */
function writeExhibitHead(ws, r, col, item, header, units, widths, fonts, number) {
  units.forEach((u, ci) => widths.want(col + ci, u));
  const totalUnits = units.reduce((a, b) => a + b, 0);
  // The title overflows to the right, up to the next visual's title one gutter along:
  // widen the table's last column when the title needs more (capped, so it stays a table)
  const titleUnits = cellUnits(item.displayTitle, fonts.exTitle) + (number ? cellUnits(String(number).padStart(2, "0") + "  ", fonts.exNum) : 0) - GUTTER_W;
  widths.span(col, col + header.length - 1, Math.min(titleUnits, totalUnits + 40));

  const t = ws.getCell(r, col);
  t.value = rich(number ? [String(number).padStart(2, "0") + "  ", fonts.exNum] : null, [item.displayTitle, fonts.exTitle]);
  t.alignment = { vertical: "bottom" };
  const cap = ws.getCell(r + 1, col);
  cap.value = rich(...(item.trend ? trendCaptionRuns(item.trend, fonts) : captionRuns(item, item.colors, header, fonts, totalUnits)));
  cap.alignment = { vertical: "middle" };
}

/* A formula over cells of one row: a range when they're side by side, else a list */
function rowFormula(ws, fn, row, cols) {
  const sorted = [...cols].sort((a, b) => a - b);
  const runs = [];
  for (const c of sorted) {
    const last = runs[runs.length - 1];
    if (last && c === last[1] + 1) last[1] = c; else runs.push([c, c]);
  }
  const refs = runs.map(([a, b]) => (a === b ? `${ws.getColumn(a).letter}${row}` : `${ws.getColumn(a).letter}${row}:${ws.getColumn(b).letter}${row}`));
  return `${fn}(${refs.join(",")})`;
}
const ROW_TOTAL_FN = { sum: "SUM", min: "MIN", max: "MAX" };

/* Returns what the totals row needs: the cells and number formats */
function writeDataTable(ws, r, col, item, tableName, widths, lay, number) {
  const fonts = lay.fonts;
  const cls = classifyColumns(item.rows);
  const header = uniqueHeaders(item.rows[0]);
  const dayFirst = dayFirstColumns(header, cls.data);   // 05/01/2023 is 5 January in a column that also has 15/01/2023
  const data = cls.data.map((row) => header.map((_, i) => toCellValue(row[i], dayFirst[i])));
  const formats = readableFormats(header, data).map((f, ci) => (item.cellFormats && item.cellFormats[ci]) || f);
  const n = data.length;
  const executive = !BUILT_IN_STYLES.includes(lay.tableStyle);

  const units = tableColumnUnits(header, data, cls.numeric, formats);
  writeExhibitHead(ws, r, col, item, header, units, widths, fonts, number);

  // Trend cards, then the table
  const hr = r + headRows(item) - 1;
  if (item.trend) writeTrendCards(ws, r + 2, col, header.length, item, widths, lay, hr);
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

  // Sparkline columns: each row's points on the page's hidden "Trends" sheet, its line in the cell
  for (const sc of item.sparkCols || []) {
    if (!lay.trendSheet) break;
    const color = lay.palette ? readableOnWhite(toArgb(lay.palette[0])) : lay.accent.cap;
    widths.want(col + sc.ci, 18);
    const page = `'${lay.sheetName}'!`;
    const rows = data.map((row, ri) => ({
      key: sc.keyCols.map((k) => keyText(row[k].value)).join("|"),
      keyExpr: sc.keyCols.map((k) => page + "$" + ws.getColumn(col + k).letter + "$" + (hr + 1 + ri)).join('&"|"&'),
      points: sc.series[ri] || [],
    }));
    const ranges = lay.trendSheet.add(sc.axis, rows);
    ranges.forEach((at, ri) => {
      (lay.sparklines = lay.sparklines || []).push({ cell: ws.getCell(hr + 1 + ri, col + sc.ci).address, values: at, dates: null, color, min: null, max: null });
    });
  }

  // A matrix's Total columns: live sums (minimum / maximum) across the row
  for (const [ci, t] of Object.entries(item.rowTotals || {})) {
    for (let ri = 0; ri < n; ri++) {
      const cell = ws.getCell(hr + 1 + ri, col + Number(ci));
      const result = data[ri][ci].value;
      cell.value = { formula: rowFormula(ws, ROW_TOTAL_FN[t.kind], hr + 1 + ri, t.of.map((c) => col + c)), result: typeof result === "number" ? result : 0 };
    }
  }

  if (item.colors) applyVisualColors(ws, hr, col, item, item.colors, lay);
  if (item.tableFormat) applyTableFormats(ws, hr, col, item, header, data, lay);
  return { data, formats };
}

/*
 * A matrix in Power BI's stepped layout (item.stepped, from matrixLayout): a header row, each
 * group's name (bold), its rows indented one step per level, the group's total row, and a
 * grand total on the closing row (`closeRow`). Totals are live SUBTOTALs over the group's rows
 * (SUBTOTAL leaves out the subtotals inside its range, so nothing counts twice). Rows get Excel
 * outline levels, folding each group into its total row with the [–] buttons beside the row
 * numbers. Not an Excel Table: a stepped report doesn't sort or filter as a list.
 * Returns the deepest outline level used.
 */
function writeSteppedMatrix(ws, r, col, item, widths, lay, number, closeRow) {
  const fonts = lay.fonts, st = item.stepped;
  const executive = !BUILT_IN_STYLES.includes(lay.tableStyle);
  const header = uniqueHeaders(item.rows[0]);
  const cls = classifyColumns(item.rows);
  const data = cls.data.map((row) => header.map((_, i) => toCellValue(row[i])));
  const formats = readableFormats(header, data).map((f, ci) => (item.cellFormats && item.cellFormats[ci]) || f);
  const numeric = header.map((_, ci) => ci > 0);
  const units = tableColumnUnits(header, data, numeric, formats, (ri) => st.body[ri].depth);
  // Totals can outgrow every row: size the value columns for the grand total too
  const kinds = item.columnKinds || [];
  units.forEach((u, ci) => {
    if (!kinds[ci]) return;
    const total = subtotalOf(kinds[ci], data.filter((_, ri) => st.body[ri].type === "leaf").map((row) => row[ci].value));
    units[ci] = Math.max(u, Math.ceil(displayText(Math.abs(total), formats[ci]).length * 0.95) + 5);
  });
  writeExhibitHead(ws, r, col, item, header, units, widths, fonts, number);

  const hr = r + 2;
  const bold = { ...fonts.body, bold: true, color: { argb: C.INK } };
  header.forEach((h, ci) => {
    const cell = ws.getCell(hr, col + ci);
    cell.value = h;
    cell.font = executive ? fonts.header : fonts.headerOnStyle;
    cell.fill = executive ? SURFACE_FILL : solidFill(C.INK);
    if (executive) cell.border = HEADER_BORDER;
    cell.alignment = { vertical: "middle", horizontal: ci ? "right" : "left", indent: ci ? 2 : 1 };
  });

  const fnOf = (ci) => SUBTOTAL_FN[kinds[ci]];
  const resultOf = (ci, from, to) => subtotalOf(kinds[ci], data.slice(from, to + 1).filter((_, i) => st.body[from + i].type === "leaf").map((row) => row[ci].value));
  let deepest = 0;
  st.body.forEach((x, ri) => {
    const sr = hr + 1 + ri;
    const level = x.type === "total" ? x.depth : x.depth + (x.type === "leaf" ? 0 : 1);
    const outline = Math.min(7, x.type === "leaf" ? st.levels - 1 : level);
    if (outline) ws.getRow(sr).outlineLevel = outline;
    deepest = Math.max(deepest, outline);
    const strong = x.type !== "leaf";
    header.forEach((_, ci) => {
      const cell = ws.getCell(sr, col + ci);
      cell.font = strong ? bold : fonts.body;
      if (executive) cell.border = x.type === "total" ? { top: { style: "thin", color: { argb: C.RULE } } } : BODY_BORDER;
      if (ci === 0) {
        cell.value = x.label;
        cell.alignment = { vertical: "middle", indent: 1 + x.depth };
        return;
      }
      cell.alignment = { vertical: "middle", horizontal: "right", indent: 2 };
      if (formats[ci]) cell.numFmt = formats[ci];
      if (x.type === "leaf") {
        const t = item.rowTotals && item.rowTotals[ci];
        const v = data[ri][ci].value;
        cell.value = t ? { formula: rowFormula(ws, ROW_TOTAL_FN[t.kind], sr, t.of.map((c) => col + c)), result: typeof v === "number" ? v : 0 } : v;
      } else if (x.type === "total" && fnOf(ci)) {
        cell.value = { formula: `SUBTOTAL(${fnOf(ci)},${colRange(ws, col + ci, hr + 1 + x.from, hr + 1 + x.to)})`, result: resultOf(ci, x.from, x.to) };
        if (formats[ci]) cell.numFmt = unitForTotal(formats[ci], cell.value.result);
      }
    });
  });

  // Groups the report shows collapsed: folded (their rows hidden under the outline's [+])
  st.body.forEach((x) => {
    if (x.fold) for (let i = x.fold.from; i <= x.fold.to; i++) ws.getRow(hr + 1 + i).hidden = true;
  });

  // Grand total, on the closing row
  const n = st.body.length;
  header.forEach((_, ci) => {
    const cell = ws.getCell(closeRow, col + ci);
    cell.font = bold;
    if (ci === 0) { cell.value = "Total"; cell.alignment = { vertical: "middle", indent: 1 }; return; }
    cell.alignment = { vertical: "middle", horizontal: "right", indent: 2 };
    if (formats[ci]) cell.numFmt = formats[ci];
    if (fnOf(ci) && n) {
      cell.value = { formula: `SUBTOTAL(${fnOf(ci)},${colRange(ws, col + ci, hr + 1, hr + n)})`, result: resultOf(ci, 0, n - 1) };
      if (formats[ci]) cell.numFmt = unitForTotal(formats[ci], cell.value.result);
    }
  });

  if (item.tableFormat) {
    const leafRows = [], leafData = [];
    st.body.forEach((x, ri) => { if (x.type === "leaf") { leafRows.push(hr + 1 + ri); leafData.push(data[ri]); } });
    applyTableFormats(ws, hr, col, item, header, leafData, lay, leafRows);
  }
  return deepest;
}

/* Excel icon sets for Power BI icon rules: the same number of icons, the closest family */
function excelIconSet(cases) {
  const n = cases.length;
  const names = cases.map((c) => c.icon).join(" ");
  const family = /signal|bar/i.test(names) ? "Rating" : /arrow/i.test(names) ? "Arrows" : /triangle/i.test(names) ? "Triangles"
    : /star/i.test(names) ? "Stars" : /flag/i.test(names) ? "Flags" : /quarter|pie/i.test(names) ? "Quarters" : "Traffic";
  const sets = {
    3: { Arrows: "3Arrows", Triangles: "3Triangles", Stars: "3Stars", Flags: "3Flags" },
    4: { Rating: "4Rating", Arrows: "4Arrows" },
    5: { Rating: "5Rating", Arrows: "5Arrows", Quarters: "5Quarters" },
  };
  if (n < 3 || n > 5) return null;
  return sets[n][family] || { 3: "3TrafficLights1", 4: "4TrafficLights", 5: "5Rating" }[n];
}

/* How "full" an icon looks, to tell whether the report's icons run high-to-low */
function iconRank(name) {
  const order = [/empty|none|red|down|low\b|bad/i, /low|one/i, /medium(?!2)|yellow|side|two/i, /medium2|three/i, /full|green|up|high|good/i];
  const i = order.findIndex((re) => re.test(name));
  return i < 0 ? 2 : i;
}

/*
 * Table / matrix visuals keep their Power BI formatting (from the report file):
 *  - background and font colours, fixed or from an fx rule (colour scale or rules), per cell,
 *    with readable text on filled cells; Excel moves them with the rows when sorting
 *  - data bars as Excel data bars, icon rules as the closest Excel icon set
 * A format names its column by field; matrices export rows then values in field order,
 * so a value field not found by name is found by its position among the values.
 */
function applyTableFormats(ws, hr, col, item, header, data, lay, rowsOf) {
  const info = item.tableFormat;
  // Sheet row of each data row (a stepped matrix's value rows sit between group rows)
  const rowAt = (ri) => (rowsOf ? rowsOf[ri] : hr + 1 + ri);
  const norm = header.map((h) => normalizeField(h));
  const lowerHeader = header.map((h) => String(h).trim().toLowerCase());
  const values = info.measures || [];
  // Each value field's column: its renamed label, else its name; matrices export the row
  // fields, the column field and the values in varying order, so a field matched by neither
  // (e.g. two fields both renamed "YoY") is placed by the offset the matched ones show
  const valueCol = values.map(() => -1);
  const taken = new Set();
  values.forEach((m, k) => {
    const ci = m.label ? lowerHeader.findIndex((h, c) => !taken.has(c) && h === m.label) : -1;
    if (ci >= 0) { valueCol[k] = ci; taken.add(ci); }
  });
  values.forEach((m, k) => {
    if (valueCol[k] >= 0) return;
    const ci = norm.findIndex((h, c) => !taken.has(c) && h === m.ref);
    if (ci >= 0) { valueCol[k] = ci; taken.add(ci); }
  });
  const offsets = {};
  valueCol.forEach((ci, k) => { if (ci >= 0) offsets[ci - k] = (offsets[ci - k] || 0) + 1; });
  const offset = Object.keys(offsets).sort((a, b) => offsets[b] - offsets[a])[0];
  if (offset !== undefined) {
    valueCol.forEach((ci, k) => {
      const c = k + Number(offset);
      if (ci < 0 && c >= 0 && c < header.length && !taken.has(c)) { valueCol[k] = c; taken.add(c); }
    });
  }
  // A field's columns: one, or a pivoted matrix's column per column value (item.valueCols)
  const columnsOf = (name) => {
    const n = normalizeField(String(name || "").toLowerCase());
    const k = values.findIndex((m) => m.ref === n);
    if (item.valueCols && k >= 0) return item.valueCols[k] || [];
    if (k >= 0 && valueCol[k] >= 0) return [valueCol[k]];
    const ci = norm.indexOf(n);
    return ci >= 0 ? [ci] : [];
  };
  if (!data.length) return;
  // Excel's range for some columns over the data rows: runs of rows, joined by spaces
  const runs = [];
  data.forEach((_, ri) => {
    const r = rowAt(ri), last = runs[runs.length - 1];
    if (last && r === last[1] + 1) last[1] = r; else runs.push([r, r]);
  });
  const refOf = (cols) => cols.flatMap((ci) => runs.map(([a, b]) => colRange(ws, col + ci, a, b))).join(" ");

  // One colour per row for a column: a fixed colour, or its fx rule (null = left as is). A
  // rule reads its fields in the same column value's columns as the cell it colours
  const rowColors = (f, ci) => {
    if (f.color) return data.map(() => toArgb(f.color));
    const at = item.valueCols ? item.valueCols.findIndex((cols) => cols.includes(ci)) : -1;
    const pos = at >= 0 ? item.valueCols[at].indexOf(ci) : -1;
    const column = (name) => {
      const cols = columnsOf(name);
      const c = pos >= 0 && cols.length > pos ? cols[pos] : cols[0];
      return c !== undefined ? data.map((r) => r[c].value ?? null) : null;
    };
    return evaluateColorRule(f.rule, column, data.length).colors.map((c) => (c ? toArgb(c) : null));
  };

  const fontCache = new Map();
  const fontIn = (argb, bold) => {
    const key = argb + (bold ? "b" : "");
    if (!fontCache.has(key)) fontCache.set(key, { ...lay.fonts.body, color: { argb } });
    return fontCache.get(key);
  };
  for (const f of info.tableFormats) {
    if (f.target === "header") {
      header.forEach((_, ci) => {
        const h = ws.getCell(hr, col + ci);
        if (f.kind === "back") { h.fill = solidFill(toArgb(f.color)); h.font = { ...h.font, color: { argb: textOn(toArgb(f.color)) } }; }
        else h.font = { ...h.font, color: { argb: toArgb(f.color) } };
      });
      continue;
    }
    const targets = f.target === "all"
      ? header.map((_, ci) => ci).filter((ci) => data.some((r) => typeof r[ci].value === "number"))
      : [...columnsOf(f.target.slice(2))];
    // Data bars and icons run over all of a field's columns as one range (one scale)
    if ((f.kind === "bar" || f.kind === "icon") && targets.length > 1 && item.valueCols) {
      targets.splice(1);
      targets.ref = refOf(columnsOf(f.target.slice(2)));
    }
    for (const ci of targets) {
      const c = col + ci;
      const ref = targets.ref || refOf([ci]);
      if (f.kind === "back" || f.kind === "font") {
        const colors = rowColors(f, ci);
        colors.forEach((argb, ri) => {
          if (!argb) return;
          const cell = ws.getCell(rowAt(ri), c);
          if (f.kind === "back") {
            cell.fill = solidFill(argb);
            cell.font = fontIn(textOn(argb));
            cell.border = GAP_BORDER;
          } else {
            cell.font = fontIn(argb);
          }
        });
      } else if (f.kind === "bar") {
        const bar = { type: "dataBar", gradient: false, border: false, showValue: f.bar.showValue, minLength: 0, maxLength: 100,
          // a share (0–100%) on one fixed scale, so columns compare; else the column's own range
          cfvo: f.bar.share ? [{ type: "num", value: 0 }, { type: "num", value: 1 }] : [{ type: "min" }, { type: "max" }], color: { argb: toArgb(f.bar.color) } };
        if (f.bar.negative) { bar.negativeBarColorSameAsPositive = false; bar.negativeFillColor = { argb: toArgb(f.bar.negative) }; }
        if (f.bar.axis) bar.axisColor = { argb: toArgb(f.bar.axis) };
        if (f.bar.reverse) bar.direction = "rightToLeft";
        ws.addConditionalFormatting({ ref, rules: [bar] });
      } else if (f.kind === "icon") {
        // Excel thresholds: % of range, a number or a percentile, like the report's
        const cfvo = (at) => (at.num !== undefined ? { type: "num", value: at.num }
          : at.pctl !== undefined ? { type: "percentile", value: Math.round(at.pctl * 100) }
            : { type: "percent", value: Math.round((at.pct || 0) * 100) });
        const sortKey = (at) => (at.num !== undefined ? at.num : (at.pct ?? at.pctl ?? 0));
        const cases = [...f.icons.cases].sort((a, b) => sortKey(a.from) - sortKey(b.from));
        const iconSet = excelIconSet(cases);
        if (!iconSet) continue;
        const reverse = iconRank(cases[cases.length - 1].icon) < iconRank(cases[0].icon);
        const steps = cases.map((x) => cfvo(x.from));
        if (steps[0].type === "num") steps[0] = { type: "min" };   // the lowest icon starts at the lowest value
        ws.addConditionalFormatting({
          ref,
          rules: [{ type: "iconSet", iconSet, reverse, showValue: true, cfvo: steps }],
        });
      }
    }
  }
}

/*
 * A chart's look on a white background (Export options → Chart background: White). The chart
 * area turns white; on a visual that was dark in the report, text, gridlines and analytics
 * lines too pale to read on white take Power BI's light-theme defaults. Data colours (bars,
 * lines, slices) and anything already readable are kept; a light visual only loses its tint.
 */
function lookOnWhite(look) {
  if (!look || !look.frame) return look;
  const bg = look.frame.background;
  const pale = (c) => !!c && contrast(toArgb(c, C.WHITE), C.WHITE) < 3;
  // A visual without a background shows the page behind it: made for a dark page when its
  // title or axis labels are pale (on a light page they would be unreadable in the report too)
  const dark = bg ? luminance(toArgb(bg, C.WHITE)) < 0.5 : (!look.title.background && pale(look.title.color)) || pale(look.cat.text.color);
  if (!bg && !dark) return look;   // no background on a light page: the sheet's white shows through, as before
  const out = structuredClone(look);
  if (bg) {
    out.frame.background = "#FFFFFF";
    out.frame.transparency = 0;
  }
  if (!dark) return out;   // a light visual: its colours read on white
  const hex = (argb) => "#" + argb.slice(2);
  const ink = hex(C.INK);
  const fix = (t, dflt) => { if (t && pale(t.color)) t.color = dflt; };
  if (!out.title.background) fix(out.title, ink);
  fix(out.subtitle, LABEL_GREY);
  fix(out.legend.text, LABEL_GREY);
  for (const a of [out.cat, out.val, out.sec]) {
    fix(a.text, LABEL_GREY);
    fix(a.titleFont, ink);
    if (a.grid) a.grid.color = GRID_GREY;   // a dark visual's gridlines are drawn for dark: Power BI's light grey
  }
  if (!out.labels.background) fix(out.labels.text, LABEL_GREY);
  fix(out.totals.text, LABEL_GREY);
  for (const l of out.refLines || []) {
    if (pale(l.color)) l.color = LABEL_GREY;
    fix(l.label, LABEL_GREY);
  }
  fix(out.trend, ink);
  return out;
}

/* Text colour that reads best on a filled cell: white or ink, whichever contrasts more */
function textOn(argb) {
  return contrast(argb, C.WHITE) >= contrast(argb, C.INK) ? C.WHITE : C.INK;
}

/*
 * Paints a chart table in the chart's colours, where the chart's legend puts them:
 *  - a colour that names a series goes on the cell that names it, so the numbers stay
 *    neutral: each row's legend value, slice or tile (category cell; a double donut's
 *    outer ring on the subcategory cell), or the header of a measure drawn as one series
 *    (also a combo chart's line values)
 *  - a colour that comes from the value (colour scales, rules, waterfall increase /
 *    decrease, colours from a field) stays on the value cells it describes
 * Text on a filled cell is white or ink, whichever reads better. Excel moves fills along
 * with their rows when sorting.
 */
function applyVisualColors(ws, hr, col, item, colors, lay) {
  const fonts = lay.fonts;
  const data = item.rows.slice(1);
  const first = hr + 1;
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

    // Series drawn in one colour (a measure such as Revenue Won): its value cells, like the bars
  // the chart draws. The header stays neutral; the caption's colour key names each colour
  for (const [ci, argb] of Object.entries(colors.series)) {
    data.forEach((row, ri) => { if (String(row[ci] ?? "").trim()) fillCell(ws.getCell(first + ri, col + Number(ci)), argb); });
  }

  // Colour that identifies a category (legend value, slice, tile, a bar set by hand): the
  // category cell in each row, like the chart's legend
  if (colors.rowKey) {
    const k = colors.rowKey.col;
    const rings = colors.rings;
    data.forEach((row, ri) => {
      const argb = colors.rowKey.colorOf[(row[k] ?? "").trim()];
      if (argb) fillCell(ws.getCell(first + ri, col + k), argb);
      if (rings && rings.colors[ri]) fillCell(ws.getCell(first + ri, col + rings.col), rings.colors[ri]);
    });
    return;
  }

  // Colour that comes from the value (rules, scales, waterfall steps, a colour field): the values
  const dot = colors.dot;
  if (!dot && !colors.perCol) return;
  data.forEach((row, ri) => {
    const key = dot ? (row[dot.keyCol] ?? "").trim() : "";
    const argb = key ? dot.colorOf(key) : null;
    for (const ci of colors.rowCols) {
      const cellColor = colors.perCol && colors.perCol[ci] ? colors.perCol[ci][ri] : argb;
      if (cellColor) fillCell(ws.getCell(first + ri, col + ci), cellColor);
    }
  });
}

/* ── Page layout ──────────────────────────────────────────────────────── */

/* Visuals overlap only by a real share of the smaller one: report authors leave visuals
   touching or overlapping by a few pixels of padding, which isn't a placement */
function overlapsEnough(overlap, sizeA, sizeB) {
  return overlap > Math.max(PX_TOLERANCE, Math.min(sizeA, sizeB) * OVERLAP_SHARE);
}

/* Bands: visuals whose vertical extents overlap, i.e. one "row" of the page */
function groupIntoBands(items) {
  const sorted = [...items].sort((a, b) => (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x));
  const bands = [];
  for (const item of sorted) {
    const band = bands[bands.length - 1];
    const bottom = item.layout.y + item.layout.height;
    if (band && overlapsEnough(band.bottom - item.layout.y, item.layout.height, band.bottom - band.top)) {
      band.items.push(item);
      band.bottom = Math.max(band.bottom, bottom);
    } else {
      bands.push({ items: [item], top: item.layout.y, bottom });
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
    const stack = stacks.find((s) => overlapsEnough(Math.min(right, s.right) - Math.max(left, s.left), item.layout.width, s.right - s.left));
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
  if (p.item.isKpi || p.item.chart) return null;
  const n = p.item.rows.length - 1, head = headRows(p.item);
  if (n <= p.vis) return null;
  const end = p.noteRow ? p.noteRow - 1 : p.r + head + n - 1;
  return [p.r + head + p.vis, end];
}

/* Rows of a planned block that must always stay visible: a whole tile, or a table's
   head and first rows (and its closing row when nothing of it collapses) */
function topRange(p) {
  if (p.item.isKpi || p.item.chart) return [p.r, p.r + p.size.height - 1];
  const n = p.item.rows.length - 1, head = headRows(p.item);
  if (n <= p.vis) return [p.r, p.r + head + n];
  return [p.r, p.r + head - 1 + p.vis];
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
 *     The table first shows a few more rows (up to `maxExtend`), so band tops stay level.
 *     If that isn't enough:
 *     - a neighbour that sits below a long table of its own column moves down instead,
 *       below this table's extra rows (it is already far from the band top, and the rows
 *       it skips are hidden, so nothing visibly moves);
 *     - otherwise this table moves down below the neighbour's visible rows, at most a
 *       few times per column: two columns of long tables could otherwise keep pushing
 *       each other down. When that limit is reached, the rows beside the neighbour simply
 *       stay visible.
 *  2. A table's note row must not sit inside a neighbour's collapsible rows, otherwise it
 *     splits the neighbour. The note moves down below them, so both notes share one row
 *     and one [+] / [–] button; what follows in its column moves only as far as it must.
 */
const MAX_STACK_SHIFTS = 3;
function alignNoteRows(plan, maxExtend) {
  const shifts = {};
  const firstInStack = (q) => !plan.some((x) => x !== q && x.stackIndex === q.stackIndex && x.r < q.r);
  // Moves the blocks of a column below `row`, keeping one blank row between blocks
  const pushBelow = (stackIndex, fromR, row) => {
    let next = row;
    for (const q of plan.filter((x) => x.stackIndex === stackIndex && x.r >= fromR).sort((a, b) => a.r - b.r)) {
      if (q.r < next) {
        const d = next - q.r;
        q.r += d;
        if (q.noteRow) q.noteRow += d;
      }
      next = q.r + q.size.height + 1;
    }
  };
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
          } else if (!firstInStack(o)) {
            pushBelow(o.stackIndex, o.r, myRange[1] + 1);
          } else if ((shifts[p.stackIndex] || 0) < MAX_STACK_SHIFTS) {
            shifts[p.stackIndex] = (shifts[p.stackIndex] || 0) + 1;
            shiftStack(plan, p, need);
          } else {
            continue;
          }
          moved = true;
          break;
        }

        const range = collapsibleRange(o);
        if (p.noteRow && range && p.noteRow >= range[0] && p.noteRow <= range[1]) {
          const delta = range[1] + 1 - p.noteRow;
          p.noteRow += delta;
          p.size.height += delta;
          pushBelow(p.stackIndex, p.r + 1, p.noteRow + 2);
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
  lay.reportNumbers = options.tableNumbers === "report";   // cards in the report's display units
  lay.reportBackground = options.chartBackground === "report";   // cards on their own background
  const threshold = options.collapse ? options.visibleRows : Infinity;
  const maxExtend = Math.max(12, options.visibleRows);
  const tables = [];
  let outline = 0;     // deepest Excel outline level (stepped matrices)
  const rowUse = [];  // per Excel row: does it hold collapsible table rows ("deep") / anything else ("other")?
  const mark = (r, kind) => { (rowUse[r] = rowUse[r] || { deep: false, other: false })[kind] = true; };

  // A KPI's main value first; titles (generic ones described from their columns, duplicates
  // numbered) and colours
  const usedTitles = {};
  const colorNotes = [];
  // A new card with several measures: one tile per measure
  items = items.flatMap((item) => splitCards(item, visualInfo(lay.reportColors, item.page, item.name)));
  items = items.flatMap((item) => keyInfluencerItems(item, visualInfo(lay.reportColors, item.page, item.name), colorNotes));
  items = items.map((item) => decompositionTreeItem(item, visualInfo(lay.reportColors, item.page, item.name)));
  if (options.visualsAs === "charts") items = items.flatMap((item) => smallMultipleItems(item, visualInfo(lay.reportColors, item.page, item.name)));
  for (const item of items) {
    item.matrixGrid = options.matrixLayout === "grid";
    const info = atDrillLevel(item, withLiveDrill(item, visualInfo(lay.reportColors, item.page, item.name)));
    if (item.isKpi) mainValueFirst(item, info);
    if ((customVisual(item.type) || {}).rule === "nested") nestedBarsRows(item, info, colorNotes);
    if ((customVisual(item.type) || {}).layout === "grid") comparisonRows(item, info);
    if ((customVisual(item.type) || {}).layout === "bins") histogramRows(item, info);
    if ((customVisual(item.type) || {}).layout === "trend") dualKpiTrend(item, info, !!lay.palette);
    if (item.type === "kpi" && !item.isKpi) kpiTrend(item, info, !!lay.palette);
    if (item.type === "pivotTable" && info && item.matrixGrid) matrixLayout(item, info);
    // Short numbers (2.47M): each number column in display units; the cells keep exact values
    if (options.tableNumbers === "short" && !item.isKpi && !item.trend) item.cellFormats = shortFormats(item);
    if (options.tableNumbers === "report" && !item.isKpi && !item.trend) item.cellFormats = reportFormats(item, info) || item.cellFormats;
    if (item.sparklines && item.sparklines.length && !item.stepped) addSparklineColumns(item);
    if (TABLE_VISUAL_TYPES.includes(item.type) && !item.stepped) item.totals = totalsPlan(item, info);
    // A title turned off in the report isn't shown either (it is often an internal name such
    // as "TYS Card"): a table is titled from its columns instead, a KPI tile by its label
    const hidden = !!(info && info.titleHidden);
    // A chart whose title the author left automatic (the embed API gives no title): the title
    // Power BI shows (as it reads on screen, else built as Power BI builds it)
    const automatic = !hidden && !item.isKpi && item.title === `${item.type} (${item.name})`;
    const live = item.liveSettings || {};
    const shownTitle = automatic ? (typeof live.titleText === "string" && live.titleText.trim()) || autoTitle(item, info) : null;
    let title = shownTitle || (hidden || isGenericTitle(item.title, item.type) ? (item.isKpi ? null : describeTitle(item)) : item.title);
    const chartTitle = title;   // a chart keeps its own title; "(2)" only tells the sheet's blocks apart
    if (title) {
      const key = title.toLowerCase();
      usedTitles[key] = (usedTitles[key] || 0) + 1;
      if (usedTitles[key] > 1) title = `${title} (${usedTitles[key]})`;
    }
    item.displayTitle = title;
    // a table's subtitle of its own ("by profitability"), shown under its title
    const ownSub = info && info.look && info.look.subtitle && info.look.subtitle.text;
    if (ownSub && !item.isKpi && !hidden) item.subtitle = ownSub;
    // A trend visual's colours are on its cards; its table stays plain
    // (a decomposition tree kept as its table: its bars' colour isn't a column's, so none)
    item.colors = lay.palette && !item.isKpi && !item.trend && !item.captionText && item.type !== "decompositionTreeVisual" ? columnColors(item, lay.palette, info) : null;
    item.tableFormat = item.gridFormat || (info && info.tableFormats && info.tableFormats.length ? info : null);
    item.colorInfo = info;
    item.themePalette = lay.palette && lay.palette.length ? lay.palette : null;   // colours for charts without their own settings
    // charts mode only; a trend visual (Dual KPI, the KPI visual) is drawn as its area charts
    // (trendChart), never as trend cards there; tables mode keeps the cards
    item.chart = options.visualsAs !== "charts" ? null : item.trend ? trendChart(item) : chartPlan(item);
    if (item.chart) {
      // The visual's format settings (Power BI's defaults without the report file)
      const reportLook = (info && info.look) || chartLook({ type: item.type }, null);
      let look = options.chartBackground === "report" ? reportLook : lookOnWhite(reportLook);   // white unless asked otherwise
      look = tornadoLook(look, item, info);
      // Small multiples share one value axis, rounded as the chart's own would be
      if (item.sharedMax !== null && item.sharedMax !== undefined && look.val && look.val.end == null) {
        const shared = valueScale([item.sharedMax], 200, 75, { zero: true });
        const end = shared && shared.step ? Number((Math.ceil(item.sharedMax / shared.step - 1e-9) * shared.step).toFixed(12)) : shared && shared.max;
        if (end > 0) look = { ...look, val: { ...look.val, start: look.val.start ?? 0, end } };
      }
      if (chartTitle) item.chart.title = chartTitle;
      // A subtitle shown without its own text is Power BI's automatic one: the automatic title
      // splits into "Orders" and "by Mth and Year"; with a title of the author's, no subtitle
      if (look.subtitle && look.subtitle.text == null) {
        const at = automatic && chartTitle ? chartTitle.indexOf(" by ") : -1;
        if (at > 0) {
          item.chart.title = chartTitle.slice(0, at);
          look = { ...look, subtitle: { ...look.subtitle, text: chartTitle.slice(at + 1) } };
        } else {
          look = { ...look, subtitle: null };
        }
      }
      Object.assign(item.chart, { look, sizePx: item.chart.sizePx || chartSizePx(item), showTitle: !(info && info.titleHidden) && look.title.show !== false });
      item.chart = withDateAxis(withNumberAxis(withNumberOrder(item.chart)));
      // A chart Excel's part can't be written for stays a table, so the rest of the file is still made
      const broken = chartBuildError(item.chart);
      if (broken) {
        log(`Chart for "${item.title}" not drawn, kept as a table: ${errorText(broken)}`);
        item.chart = null;
      }
      if (item.chart && item.panels) item.chart.panels = panelPlans(item, info, look, lay.palette);
      if (item.chart && item.chart.trendCards) item.chart.panels = trendPanels(item, look);
      if (item.chart && !item.chart.panels) item.chart = growForLegend(item.chart);
    }
    if (item.chart && item.chart.note) colorNotes.push(item.chart.note);
    // Two columns under one name in the report (two measures both renamed "YoY"): the second is
    // "YoY (2)" here, as an Excel table can't repeat a column name
    if (!item.chart && !item.isKpi && item.rows && item.rows[0]) {
      const names = item.rows[0].map((h) => String(h ?? "").trim()).filter(Boolean);
      const twice = [...new Set(names.filter((h, i) => names.findIndex((x) => x.toLowerCase() === h.toLowerCase()) !== i))];
      for (const name of twice) {
        colorNotes.push(`"${item.displayTitle || item.title}": two columns are named "${name}" in the report (two different fields); ` +
          `the second is "${name} (2)" here, as an Excel table can't repeat a column name. Their totals aren't worked out, as the export can't tell the two apart.`);
      }
    }
    // The anomaly band and expected values: Power BI's AI works them out on screen only
    if (info && info.anomalies) colorNotes.push(`Not included: the anomaly detection on "${item.displayTitle || item.title}" (its expected range and values) – Power BI's AI works it out on screen. The values themselves are exported.`);
    if (item.colors && item.colors.note) colorNotes.push(item.colors.note);
  }
  const executive = !BUILT_IN_STYLES.includes(lay.tableStyle);

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
      let col = cursorCol;
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
          item, col, r, size, stackIndex, vis: item.stepped ? Infinity : threshold, extra: 0,
          noteRow: size.note ? r + headRows(item) + n : null,
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
      if (item.chart) {
        // Charts mode: reserve the chart's place; the chart itself is added to the file at the end
        const { w, h } = item.chart.sizePx || chartSizePx(item);
        widths.span(col, col + size.width - 1, Math.ceil(w / UNIT_PX));
        (lay.charts = lay.charts || []).push({ plan: item.chart, col, row: r, wPx: w, hPx: h });
        for (let i = 0; i < size.height; i++) mark(r + i, "other");
      } else if (item.isKpi) {
        writeKpiBlock(ws, r, col, item, widths, lay, p.divider, p.strip);
        for (let i = 0; i < size.height; i++) mark(r + i, "other");
      } else {
        const number = options.exhibitNumbers ? ++exhibit : null;
        const n = item.rows.length - 1, head = headRows(item);
        const closeRow = p.noteRow || r + head + n;
        if (item.stepped) {
          outline = Math.max(outline, writeSteppedMatrix(ws, r, col, item, widths, lay, number, closeRow));
        } else {
          const tableName = nameTable(item.displayTitle);
          const table = writeDataTable(ws, r, col, item, tableName, widths, lay, number);
          if (lay.slicerFields.length) tableSlicers(item, tableName, r + head - 1, lay);
          // The totals row is the closing row; a "rows grouped" note goes on the row below
          if (item.totals) writeTotalsRow(ws, closeRow, col, item, r + head - 1, n, lay, table);
        }
        for (let i = 0; i < head; i++) mark(r + i, "other");
        for (let i = 0; i < n; i++) mark(r + head + i, i >= p.vis ? "deep" : "other");
        for (let x = r + head + n; x < closeRow; x++) mark(x, "deep");   // padding lined up with a neighbour
        mark(closeRow, "other");
        if (item.totals && p.noteRow) mark(closeRow + 1, "other");
        if (executive || item.totals || item.stepped) paintRange(ws, closeRow, col, col + size.width - 1, { border: { top: INK_RULE } });
        tables.push({ top: r, head, col, w: size.width, n, vis: p.vis, noteRow: p.noteRow ? p.noteRow + (item.totals ? 1 : 0) : null });
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
    for (let i = t.vis; i < t.n; i++) if (ws.getRow(t.top + t.head + i).hidden) hidden++;
    if (hidden) pendingNotes.push({ row: t.noteRow, col: t.col, w: t.w, n: t.n, hidden });
  }

  return { nextRow, lastCol, tableCount: tables.length, grouped, outline: Math.max(outline, grouped ? 1 : 0), pendingNotes, colorNotes };
}

/* ── Collecting a page ────────────────────────────────────────────────── */

/* A slicer with nothing selected has no filter to name its field, so read the field
   names from the header of its data (one row is enough). */
async function slicerFieldNames(visual, fallback) {
  try {
    const result = await visual.exportData(models.ExportDataType.Summarized, 1);
    const header = parseCsv(result.data)[0];
    // An AppSource filter visual may export its value measure too (an attribute slicer's
    // bars): its field is the first column
    if (header && header.length) return customVisual(visual.type) ? header[0] : header.join(", ");
  } catch (e) {
    // fall through to the visual's title
  }
  return fallback;
}

/* A table's Excel slicers: one per page slicer field among its columns, with its item count */
function tableSlicers(item, tableName, headerRow, lay) {
  const header = uniqueHeaders(item.rows[0]);
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  for (const field of lay.slicerFields) {
    const ci = lower.indexOf(String(field).trim().toLowerCase());
    if (ci < 0) continue;
    const count = new Set(item.rows.slice(1).map((r) => String(r[ci] ?? ""))).size;
    (lay.slicers = lay.slicers || []).push({ tableName, column: header[ci], caption: header[ci], title: item.displayTitle, row: headerRow, count });
  }
}

/* A slicer's fields, named as the visuals' exports name them (a hierarchy slicer: each level) */
async function slicerColumns(visual) {
  try {
    const result = await visual.exportData(models.ExportDataType.Summarized, 1);
    const header = parseCsv(result.data)[0] || [];
    return (customVisual(visual.type) ? header.slice(0, 1) : header).map((h) => String(h).trim()).filter(Boolean);
  } catch (e) {
    return [];
  }
}

function isHiddenVisual(v) {
  return !!(v.layout && v.layout.displayState && v.layout.displayState.mode === models.VisualContainerDisplayMode.Hidden);
}

/* Visuals worth offering in the options checklist (not slicers, not decorative) */
async function listDataVisuals(page) {
  const visuals = await page.getVisuals();
  return visuals
    .filter((v) => !isSlicerType(v.type) && !NON_DATA_TYPES.includes(v.type) && !isHiddenVisual(v))
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

/* A report's, page's or visual's filters, or none when Power BI won't say */
async function filtersOf(target) {
  try {
    return (target && typeof target.getFilters === "function" && (await target.getFilters())) || [];
  } catch (e) {
    return [];
  }
}

/*
 * Reads one page: filter chips, exportable items and notes. Visuals are exported one
 * at a time (Power BI rejects concurrent exportData calls).
 * options.excludedVisuals: Set of visual names to leave out (current-page mode only).
 */
async function collectPage(page, options, onProgress, isCancelled) {
  const visuals = (await page.getVisuals()).sort((a, b) =>
    (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x));
  if (typeof notePageVisuals === "function") notePageVisuals(page.name, visuals);   // finds visuals by place where names differ

  const filterChips = [];
  // The filters as models too: table sparklines are queried under them
  const queryFilters = [...await filtersOf(embeddedReport), ...await filtersOf(page)];
  if (options.includeFilters) {
    filterParts(queryFilters).forEach((p) => filterChips.push({ all: false, ...p }));
    // What the user clicked in the report (a bar, a slice): it cross-filters the page as exported
    if (options.selection) options.selection.parts.forEach((p) => filterChips.push({ all: false, field: `${p.field} (selected)`, value: p.values.join(", ") }));
  }

  const items = [];
  const notes = [];
  const slicerFields = [];   // "Excel slicers": the fields the page's slicers filter
  for (let i = 0; i < visuals.length; i++) {
    if (isCancelled()) throw { cancelled: true };
    const v = visuals[i];
    const title = v.title || `${v.type} (${v.name})`;
    onProgress(i + 1, visuals.length, title);

    // Slicers become filter chips: "Field  selection", or "Field  All". A hidden slicer still
    // filters the page (bookmarks often hide them), so it gets a chip too.
    if (isSlicerType(v.type)) {
      if (options.excelSlicers && !isHiddenVisual(v)) slicerFields.push(...await slicerColumns(v));
      if (!options.includeFilters) {
        try { queryFilters.push(...((await v.getSlicerState()).filters || [])); } catch (e) { /* not readable */ }
        continue;
      }
      try {
        const state = await v.getSlicerState();
        queryFilters.push(...(state.filters || []));
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
        // Newer slicer types may not report their selection: name the field, say so in the bar
        log(`Could not read the selection of "${title}": ${errorText(e)}`);
        filterChips.push({ field: await slicerFieldNames(v, title), value: "selection not readable", all: true });
      }
      continue;
    }

    // Hidden visuals aren't exported; only hidden data visuals are worth a note (decorative
    // ones are routinely hidden to drive bookmarks)
    if (isHiddenVisual(v)) {
      if (!NON_DATA_TYPES.includes(v.type)) notes.push(`Skipped hidden visual: ${title}`);
      continue;
    }
    if (options.excludedVisuals && options.excludedVisuals.has(v.name)) continue;
    if ((customVisual(v.type) || {}).rule === "none") continue;   // an AppSource visual with no data (a button)
    // Visuals whose content Power BI works out on screen: their export would only be raw rows, or nothing
    // (key influencers are exported: their rows become an insight table, see keyInfluencerItems)
    if (ON_SCREEN_ONLY[v.type] && v.type !== "keyDriversVisual") {
      notes.push(`Not included: "${v.title || friendlyType(v.type)}" – ${ON_SCREEN_ONLY[v.type]}`);
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
      const isKpi = (KPI_TYPES.includes(v.type) || (customVisual(v.type) || {}).rule === "kpi") && rows.length === 2;
      if (!isKpi && rows[0].length === 1) {
        if (!NON_DATA_TYPES.includes(v.type)) {
          notes.push(`Skipped "${title}": a single list of values with no measures (for example a what-if parameter)`);
        }
        continue;
      }
      // Visual-level filters: table sparklines and drill-level queries run under them too
      const visualFilters = await filtersOf(v);
      // A chart's settings as Power BI resolves them on screen (axis titles, title text): for every role
      const liveSettings = !isKpi && chartSpecOf(v.type) ? await liveVisualSettings(v) : null;
      items.push({ title, name: v.name, page: page.name, type: v.type, layout: v.layout, rows, isKpi, visualFilters, liveSettings });
      log(`Exported "${title}" (${rows.length - 1} rows)`);
    } catch (e) {
      if (NON_DATA_TYPES.includes(v.type)) continue;
      // Visuals Power BI never exports from: say so plainly rather than as a failure
      const blocked = NOT_EXPORTABLE[v.type];
      notes.push(blocked ? `Not included: "${title}" – Power BI doesn't export data from ${blocked}.`
        : `Export failed for "${title}" (${v.type}): ${errorText(e)}`);
      log(`Export failed for "${title}": ${errorText(e)}`);
    }
  }
  return { filterChips, items, notes, queryFilters, slicerFields: [...new Set(slicerFields)] };
}

/* Visuals whose data Power BI doesn't export, by type: what to call them in a note */
const ON_SCREEN_ONLY = {
  keyDriversVisual: "a key influencers visual is an AI analysis Power BI works out on screen; its export would only be the raw rows behind it. View it in the report.",
  aiNarratives: "a smart narrative's text is written by Power BI on screen and isn't exported. View it in the report.",
  qnaVisual: "a Q&A visual answers questions on screen and has no data to export.",
};
const NOT_EXPORTABLE = {
  keyDriversVisual: "key influencers visuals", qnaVisual: "Q&A visuals", scriptVisual: "R visuals", pythonVisual: "Python visuals",
  rdlVisual: "paginated report visuals", powerApps: "Power Apps visuals", PowerApps_PowerBI_Visual: "Power Apps visuals",
};

/*
 * A visual's settings as Power BI resolves them on screen (its own setting, else the report's
 * themes, which Viewers can't read), through the embed API's visual properties: whether each
 * axis shows its title, and the title's text. Each one undefined when the API doesn't say;
 * null when it says nothing at all.
 */
async function liveVisualSettings(v) {
  if (typeof v.getProperty !== "function") return null;
  const read = async (objectName, propertyName) => {
    try {
      const r = await v.getProperty({ objectName, propertyName });
      return r && r.value !== undefined && r.value !== null ? r.value : undefined;
    } catch (e) {
      return undefined;
    }
  };
  // (The API has no axis-title setting: those come from the themes, see baseThemeAxisStyles)
  const out = { titleText: await read("title", "titleText") };
  return Object.values(out).some((x) => x !== undefined) ? out : null;
}

/* An auto title's field name, as Power BI writes it: the name it was given in the visual,
   else an aggregated column's "Sum of Sales" / "Count of Product (Distinct)", else its name */
const AGG_NAMES = { 0: "Sum of", 1: "Average of", 2: "Count of", 3: "Min of", 4: "Max of", 5: "Count of", 6: "Median of", 7: "Standard deviation of", 8: "Variance of" };
function autoFieldName(f, header) {
  if (f.label) {
    const h = header.find((x) => String(x ?? "").trim().toLowerCase() === f.label);
    if (h) return String(h).trim();
  }
  const fi = f.field;
  if (!fi || !fi.prop) return null;
  if (fi.kind === "agg" && AGG_NAMES[fi.agg]) return `${AGG_NAMES[fi.agg]} ${fi.prop}${fi.agg === 2 ? " (Distinct)" : ""}`;
  return String(fi.prop);
}

/*
 * The title Power BI writes on a chart whose author left it automatic: its values, then "by"
 * and its axis and legend fields ("Sum of Sales and Profit Margin by Month", "Sum of Sales by
 * Country and Category"); a hierarchy by the level shown. null when the layout doesn't say.
 */
const AUTO_TITLE_VALUES = ["Y", "Y2", "Values", "ColumnY", "LineY"];
const AUTO_TITLE_GROUPS = ["Category", "Group", "Series", "Details", "Breakdown"];
function autoTitle(item, info) {
  if (!info || !info.roles || !EXCEL_CHARTS[item.type] || customVisual(item.type) || EXCEL_CHARTS[item.type].kind === "scatter") return null;
  const header = item.rows[0] || [];
  const names = (list) => list.map((f) => autoFieldName(f, header));
  const values = names(AUTO_TITLE_VALUES.flatMap((r) => info.roles[r] || []));
  const groups = names(AUTO_TITLE_GROUPS.flatMap((r) => {
    const list = info.roles[r] || [];
    const shown = list.filter((f) => f.active === true);
    return shown.length ? shown : list.slice(0, 1);
  }));
  if (!values.length || values.some((n) => !n) || groups.some((n) => !n)) return null;
  const join = (l) => (l.length <= 2 ? l.join(" and ") : `${l.slice(0, -1).join(", ")} and ${l[l.length - 1]}`);
  return groups.length ? `${join(values)} by ${join(groups)}` : join(values);
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
/*
 * The page's title text box, from its design: a short heading in a text box, one line in large
 * type (16pt or more), at least a fifth of the page wide, set above every chart and table on
 * the page ("Warehouse Fulfillment Dashboard"); the largest when there are several, with its
 * alignment on the page. null when the page has none.
 */
const PAGE_TITLE_MIN_PT = 14;
/* A text box that heads the page: one short line in large type (14pt or more), a fifth of the
   page wide or more, above every chart and table */
function isPageHeading(v, onPage, pageW) {
  const top = (x) => Number(x.position.y);
  const dataTops = onPage.filter((x) => x.visualType && !NON_DATA_TYPES.includes(x.visualType) && x.visualType !== "slicer").map(top).filter(isFinite);
  const firstData = dataTops.length ? Math.min(...dataTops) : Infinity;
  const W = pageW || Math.max(...onPage.map((x) => Number(x.position.x || 0) + Number(x.position.width || 0)).filter(isFinite), 1);
  // a heading, not a paragraph: one to three short lines ("Opportunity Analysis Report" over
  // "Opportunity detail"), none a sentence
  const lines = v.textBox ? String(v.textBox.text).split("\n") : [];
  return !!(v.visualType === "textbox" && v.textBox && lines.length >= 1 && lines.length <= 3 && v.textBox.size >= PAGE_TITLE_MIN_PT &&
    lines.every((l) => l.length <= 70 && !/[.!?]$/.test(l)) && Number(v.position.width || 0) >= W * 0.2 &&
    isFinite(top(v)) && isFinite(firstData) && top(v) + Number(v.position.height || 0) <= firstData + 1);
}
function pageTitleBox(reportColors, page) {
  const visuals = (reportColors && reportColors.visuals) || {};
  const prefix = String(page && page.name) + "\u0001";
  const onPage = Object.keys(visuals).filter((k) => k.startsWith(prefix)).map((k) => visuals[k]).filter((v) => v && v.position);
  const top = (v) => Number(v.position.y);
  const dataTops = onPage.filter((v) => v.visualType && !NON_DATA_TYPES.includes(v.visualType) && v.visualType !== "slicer").map(top).filter(isFinite);
  const firstData = dataTops.length ? Math.min(...dataTops) : Infinity;
  if (!isFinite(firstData)) return null;   // a page of text only: no heading over a chart or table
  if (!onPage.some((v) => isPageHeading(v, onPage))) return null;
  // a title spans a good part of the page (a small badge such as "Free" isn't one)
  const pageW = Math.max(...onPage.map((v) => Number(v.position.x || 0) + Number(v.position.width || 0)).filter(isFinite), 1);
  // a heading, not a sentence: short, and not ending in a full stop
  const titles = onPage.filter((v) => isPageHeading(v, onPage, pageW));
  if (!titles.length) return null;
  titles.sort((a, b) => b.textBox.size - a.textBox.size || top(a) - top(b));
  return asTitle(titles[0], pageW);
}
/* Every heading box over the page's visuals, top to bottom ("REGIONAL SALES", then "Sales
   Overview"), at most three */
function pageTitleBoxes(reportColors, page) {
  const first = pageTitleBox(reportColors, page);
  if (!first) return [];
  const visuals = (reportColors && reportColors.visuals) || {};
  const prefix = String(page && page.name) + "\u0001";
  const onPage = Object.keys(visuals).filter((k) => k.startsWith(prefix)).map((k) => visuals[k]).filter((v) => v && v.position);
  const pageW = Math.max(...onPage.map((v) => Number(v.position.x || 0) + Number(v.position.width || 0)).filter(isFinite), 1);
  const heads = onPage.filter((v) => isPageHeading(v, onPage, pageW));
  heads.sort((a, b) => Number(a.position.y) - Number(b.position.y) || Number(a.position.x) - Number(b.position.x));
  return heads.flatMap((v) => headingLines(asTitle(v, pageW))).slice(0, 4);
}
/* A heading box of several lines as one title per line, each in its own type */
function headingLines(box) {
  if (box.lines <= 1 || !box.runs) return [box];
  const lines = [[]];
  for (const r of box.runs) { if (r.text === "\n") lines.push([]); else lines[lines.length - 1].push(r); }
  return lines.filter((runs) => runs.some((r) => r.text.trim())).map((runs) => {
    const sizes = runs.map((r) => r.size ?? TEXT_BOX_DEFAULT_PT);
    const main = runs[sizes.indexOf(Math.max(...sizes))];
    return { ...box, text: runs.map((r) => r.text).join("").trim(), lines: 1, size: Math.max(...sizes), bold: main.bold, color: main.color, face: main.face, runs };
  });
}
function asTitle(t, pageW) {
  // left, centre or right, by where its middle sits across the page
  const mid = (Number(t.position.x || 0) + Number(t.position.width || 0) / 2) / pageW;
  return { ...t.textBox, align: mid < 0.34 ? "left" : mid > 0.66 ? "right" : "center" };
}
function pageTitleText(reportColors, page) {
  const box = pageTitleBox(reportColors, page);
  return box ? box.text : null;
}

/* The title text box across the page's width, placed as in the report, each part of its text in
   its own type ("CORPORATE SPEND SAMPLE PBIX" big and bold, "YTD IT Spend Trend Analysis" small);
   type over 26pt is scaled down, every part alike */
const PAGE_TITLE_MAX_PT = 26;
function writeTitleBox(ws, row, c1, c2, box) {
  if (c2 > c1) ws.mergeCellsWithoutStyle(row, c1, row, c2);
  const c = ws.getCell(row, c1);
  const scale = box.size > PAGE_TITLE_MAX_PT ? PAGE_TITLE_MAX_PT / box.size : 1;
  const fontOf = (r) => ({ name: r.face || box.face || "Segoe UI", size: Math.max(6, Math.round((r.size ?? TEXT_BOX_DEFAULT_PT) * scale)),
    // white or pale type (made for a dark banner the export doesn't draw) reads dark on the sheet
    bold: !!r.bold, italic: !!r.italic, color: { argb: ((c) => (luminance(c) > 0.6 ? toArgb("#252423") : readableOnWhite(c)))(toArgb(r.color || "#252423")) } });
  const runs = (box.runs || []).filter((r) => r.text !== "\n");
  if (runs.length > 1) {
    c.value = { richText: runs.map((r) => ({ text: r.text, font: fontOf(r) })) };
  } else {
    c.value = box.text;
    c.font = fontOf(runs[0] || box);
  }
  c.alignment = { vertical: "middle", horizontal: box.align, indent: box.align === "center" ? 0 : 1 };
}

function writePageSheet(wb, sheetName, page, data, options, branding, logo, nameTable, ctx) {
  const accent = accentPlan(toArgb(branding.accent));
  const fonts = makeFonts(branding.titleFont);
  const lay = {
    fonts, accent, sheetName,
    trendSheet: ctx.trendSheet ? makeTrendSheet(ctx.trendSheet) : null,
    tableStyle: branding.tableStyle || EXECUTIVE,
    palette: options.matchColors ? options.palette : null,
    reportColors: options.matchColors ? options.reportColors || null : null,
    slicerFields: options.excelSlicers ? data.slicerFields || [] : [],
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
  // The page's title text box ("Warehouse Fulfillment Dashboard"), above the visuals as in the report
  const titleBoxes = data.items.length ? pageTitleBoxes(lay.reportColors, page) : [];
  const titleBoxRows = titleBoxes.map((b) => { const at = row; heights[row++] = Math.round(Math.min(b.size, PAGE_TITLE_MAX_PT) * 1.6); return at; });
  if (titleBoxes.length) heights[row++] = 8;     // spacer
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

  titleBoxes.forEach((b, k) => writeTitleBox(ws, titleBoxRows[k], B, contentEnd, b));

  // 4. Rules and fills
  paintRange(ws, 1, 1, bandEnd, { fill: solidFill(accent.rule) });
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
  if (layout.outline) ws.properties.outlineLevelRow = layout.outline;

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

    return {
    tables: layout.tableCount,
    kpis: data.items.filter((it) => it.isKpi).length,
    charts: lay.charts || [],   // charts mode: charts to add to this sheet (Step 9)
    slicers: (lay.slicers || []).map((x) => ({ ...x, col: bandEnd + 2 })),   // Excel slicers, right of the page (addSlicersToFile)
    sparklines: lay.sparklines || [],   // trend cards' sparklines, added by addSparklinesToFile
    trendSheet: lay.trendSheet && lay.trendSheet.blocks.length ? lay.trendSheet : null,
  };
}

/* ── Charts mode ("Visuals as: Charts") ───────────────────────────────────
 * Chart visuals become real Excel charts, placed like the report. ExcelJS can't write
 * charts, so: the page sheet reserves each chart's place, the chart's numbers go on the
 * page's "Data" sheet, and addChartsToFile() adds the chart parts to the finished .xlsx
 * (with JSZip). Only used when options.visualsAs === "charts"; the data tables mode
 * never runs this code. Visuals without an Excel chart equivalent stay data tables.
 * With the report file, a chart draws the fields the visual draws (tooltip-only measures
 * are left out; a combo's line values are its Y2 fields), and scatter charts know their
 * X, Y and Size fields.
 */
const EXCEL_CHARTS = {
  clusteredColumnChart: { kind: "bar", dir: "col", grouping: "clustered" },
  columnChart: { kind: "bar", dir: "col", grouping: "stacked" },
  hundredPercentStackedColumnChart: { kind: "bar", dir: "col", grouping: "percentStacked" },
  clusteredBarChart: { kind: "bar", dir: "bar", grouping: "clustered" },
  barChart: { kind: "bar", dir: "bar", grouping: "stacked" },
  hundredPercentStackedBarChart: { kind: "bar", dir: "bar", grouping: "percentStacked" },
  lineChart: { kind: "line" },
  areaChart: { kind: "area", grouping: "standard" },
  stackedAreaChart: { kind: "area", grouping: "stacked" },
  hundredPercentStackedAreaChart: { kind: "area", grouping: "percentStacked" },
  lineClusteredColumnComboChart: { kind: "combo", dir: "col", grouping: "clustered" },
  lineStackedColumnComboChart: { kind: "combo", dir: "col", grouping: "stacked" },
  pieChart: { kind: "pie" },
  donutChart: { kind: "doughnut" },
  scatterChart: { kind: "scatter" },   // a bubble chart when the visual has a Size field
  waterfallChart: { kind: "waterfall" },
  funnel: { kind: "funnel" },
  treemap: { kind: "treemap" },   // an Excel treemap (Excel 2016 and later)
  RadarChart1446119667547: { kind: "radar" },   // the AppSource Radar Chart: Excel's own filled radar
  Sunburst1445472000808: { kind: "sunburst" },   // the AppSource Sunburst: an Excel sunburst (Excel 2016 and later)
  Histogram1445664487616: { kind: "bar", dir: "col", grouping: "clustered", gap: 0 },   // the AppSource Histogram: its own bins, as touching columns
  TornadoChart1452517688218: { kind: "bar", dir: "bar", grouping: "clustered", mirror: true },   // the AppSource Tornado: two sides back to back
};
// AppSource box & whisker visuals (several makers): Excel's own box & whisker chart (Excel 2016 and later)
const BOX_WHISKER_TYPE = /box.?(and.?)?whisker|boxplot/i;
function chartSpecOf(type) {
  return EXCEL_CHARTS[type] || (BOX_WHISKER_TYPE.test(String(type || "")) ? { kind: "boxWhisker" } : null);
}
const CHART_SCALE = 1;          // Power BI page pixels → chart pixels
const CHART_MIN_W = 360;        // smallest chart, in pixels
const CHART_MIN_H = 220;
const CHART_PX_TO_EMU = 9525;   // Excel drawing units per pixel
const ROW_PX = ROW_H * 4 / 3;   // one sheet row, in pixels
const UNIT_PX = 7;              // about one Excel column width unit, in pixels

/* The chart's size in pixels: the visual's size on the report page */
function chartSizePx(item) {
  const w = Math.max(CHART_MIN_W, Math.round(((item.layout && item.layout.width) || 400) * CHART_SCALE));
  const h = Math.max(CHART_MIN_H, Math.round(((item.layout && item.layout.height) || 260) * CHART_SCALE));
  return { w, h };
}

/* Rows and columns a chart covers on the page sheet */
function chartBlockSize(item) {
  const { w, h } = (item.chart && item.chart.sizePx) || chartSizePx(item);
  return { width: Math.max(2, Math.round(w / 90)), height: Math.ceil(h / ROW_PX) + 1, note: false };
}

/*
 * What a chart visual draws, from its exported rows: categories (the axis), one series
 * per measure (or per legend value), and colours when "Match the report's chart colours"
 * is on. Returns null when the visual can't be an Excel chart (it then stays a table).
 */
function chartPlan(item) {
  const spec = chartSpecOf(item.type);
  if (!spec || item.isKpi) return null;
  // A tornado is a bar chart of its two sides, the first drawn to the left
  if (spec.mirror) {
    const plan = chartPlan({ ...item, type: "clusteredBarChart" });
    return plan && mirrorPlan(plan, spec);
  }
  // Scatter, waterfall and funnel have their own builders; any problem keeps the visual as a table
  if (["scatter", "waterfall", "funnel", "treemap", "sunburst", "boxWhisker"].includes(spec.kind)) {
    try {
      return spec.kind === "scatter" ? scatterPlan(item) : spec.kind === "waterfall" ? waterfallPlan(item)
        : spec.kind === "funnel" ? funnelPlan(item) : spec.kind === "boxWhisker" ? boxPlan(item) : treemapPlan(item, spec);
    } catch (e) {
      log(`Chart for "${item.title}" not drawn, kept as a table: ${errorText(e)}`);
      return null;
    }
  }
  const { header, data, numeric, hex, helper, filled } = classifyColumns(item.rows);
  if (!data.length) return null;
  const cols = header.map((_, ci) => ci);
  const colors = item.colors || null;
  // A drawn measure with no value at all (a Target not set yet): Power BI still lists it in
  // the legend and keeps its room beside the others
  const empty = (ci) => filled(ci).length === 0;
  const info = item.colorInfo || null;
  const key = (r, ci) => String(r[ci] ?? "").trim();
  const num = (raw) => { const v = toCellValue(raw).value; return typeof v === "number" ? v : null; };
  const formatOf = (ci) => {
    const cells = data.map((r) => [toCellValue(r[ci])]);
    const readable = readableFormats([header[ci]], cells)[0];
    const exported = cells.map((c) => c[0]).find((c) => typeof c.value === "number" && c.numFmt);
    return readable || (exported && exported.numFmt) || "General";
  };
  // A report field's column: its renamed label first, then its name (each column used once)
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const norm = header.map((h) => normalizeField(h));
  const taken = new Set();
  const colOf = (f) => {
    if (!f) return -1;
    const free = cols.filter((c) => !taken.has(c));
    let ci = f.label ? free.find((c) => lower[c] === f.label) : undefined;
    if (ci === undefined) ci = free.find((c) => norm[c] === f.ref);
    if (ci === undefined) return -1;
    taken.add(ci);
    return ci;
  };

  const textCols = cols.filter((ci) => !numeric[ci] && !hex[ci]);
  let catCol = textCols.length ? textCols[0] : 0;
  let valueCols = cols.filter((ci) => ci !== catCol && numeric[ci] && !hex[ci] && !helper[ci]);
  // With the report file, the series are the fields the visual draws: tooltip-only measures
  // aren't drawn, and a combo's line values are its Y2 fields
  let lineCols = null;
  if (info) {
    const c = colOf(info.category);
    if (c >= 0) catCol = c; else taken.add(catCol);
    const drawn = (info.measures || []).map((m) => ({ ci: colOf(m), role: m.role }))
      .filter((m) => m.ci >= 0 && (numeric[m.ci] || empty(m.ci)) && !hex[m.ci]);
    // Only alongside a measure with values: a chart of nothing but empty measures stays as it was
    if (drawn.length && drawn.every((m) => empty(m.ci))) drawn.length = 0;
    if (drawn.length) {
      valueCols = drawn.map((m) => m.ci);
      lineCols = drawn.filter((m) => m.role === "Y2").map((m) => m.ci);
    } else {
      const tooltips = new Set(info.tooltips || []);
      valueCols = valueCols.filter((ci) => ci !== catCol && !tooltips.has(norm[ci]));
    }
  }
  if (!valueCols.length) return null;
  // A category hierarchy Power BI shows expanded (Chain, then Name: a bar per store) is
  // exported at its lowest level shown: every level in the export makes up the bar's label,
  // joined as Power BI joins them, so the rows aren't totalled back up to the top level
  const levelCols = info ? [catCol, ...(info.categories || []).slice(1).map((f) => colOf(f)).filter((c) => c >= 0)] : [catCol];
  const catOf = (r) => levelCols.map((c) => key(r, c)).filter(Boolean).join(" ");
  const categories = [...new Set(data.map((r) => catOf(r)))];
  const firstRowOf = (c) => data.findIndex((r) => catOf(r) === c);
  // One value per category; a category with several rows (a hierarchy) is totalled
  const valuesFor = (ci, rows) => {
    const additive = !NON_ADDITIVE.test(header[ci]) && !data.some((r) => key(r, ci).endsWith("%"));
    return categories.map((c) => {
      const vals = rows.filter((r) => catOf(r) === c).map((r) => num(r[ci])).filter((v) => v !== null);
      if (!vals.length) return null;
      return additive ? vals.reduce((a, b) => a + b, 0) : vals[0];
    });
  };
  const base = { spec, title: item.displayTitle || item.title, catName: levelCols.map((c) => header[c]).join(" / "), categories, pointColors: null,
    axes: (info && info.axes) || null, legendName: null };
  // A hierarchy shown with its labels not joined (Power BI's default): Excel's multi-level
  // axis labels, each higher level's value under its first category (Quarter, then Year below)
  const look = (info && info.look) || null;
  if (levelCols.length > 1 && !(look && look.cat.concatenate) && ["bar", "line", "area", "combo"].includes(spec.kind)) {
    const firstRows = categories.map((c) => data[firstRowOf(c)]);
    const label = (r, ci) => String(axisLabel(key(r, ci)));
    base.multiLevel = true;
    base.catTitle = header[levelCols[levelCols.length - 1]];   // Power BI titles the axis by its lowest level
    base.catText = firstRows.map((r) => label(r, levelCols[levelCols.length - 1]));
    base.catLevels = levelCols.slice(0, -1).map((ci, k) => {
      const values = firstRows.map((r) => label(r, ci));
      const outer = (i) => levelCols.slice(0, k + 1).map((c) => key(firstRows[i], c)).join("\u0001");
      return { name: header[ci], values, starts: values.map((_, i) => i === 0 || outer(i) !== outer(i - 1)) };
    });
  }

  // Pie / donut: one series, a colour per slice
  if (spec.kind === "pie" || spec.kind === "doughnut") {
    const vc = valueCols[0];
    // Levels shown together: each slice ("East New York") in its own colour (columnColors)
    const byKey = colors && colors.tiles ? colors.tiles.colorOf
      : colors && colors.rowKey && colors.rowKey.col === catCol ? colors.rowKey.colorOf : null;
    // No colour per slice from the legend: a colour rule's row colours (as a bar chart's), else
    // each slice's colour set by hand or the theme's in order, as Power BI draws it
    let pointColors = byKey ? categories.map((c) => byKey[c] || null) : null;
    if (!pointColors && colors && colors.perCol && colors.perCol[vc]) pointColors = categories.map((c) => colors.perCol[vc][firstRowOf(c)] || null);
    else if (!pointColors && colors && colors.tabColors) pointColors = categories.map((c) => colors.tabColors[firstRowOf(c)] || null);
    if (!pointColors) pointColors = themeColorsOf(item, categories);
    return {
      ...base,
      legendName: base.catName,
      series: [{ name: header[vc], values: valuesFor(vc, data), fmt: formatOf(vc), color: null }],
      pointColors,
    };
  }

  // A legend splits one measure into one series per legend value
  let legendCol = -1;
  if (info && info.rule === "legend" && info.series) {
    legendCol = header.findIndex((h) => normalizeField(h) === info.series.ref || String(h).trim().toLowerCase() === info.series.label);
  } else if (!info && textCols.length === 2 && valueCols.length === 1) {
    legendCol = textCols[1];
  }
  const liveLegend = legendCol < 0 && !info && textCols.length > 2 && valueCols.length === 1;
  if (liveLegend) legendCol = liveLegendCol(item, header, textCols);
  if (legendCol >= 0 && legendCol !== catCol) {
    // A combo's legend splits only its column values: its line values (Y2) stay one line
    // each, on a second axis. A line value repeats on every legend row of a category, so it
    // is read once per category, not added up.
    const lines = spec.kind === "combo" && lineCols ? valueCols.filter((ci) => lineCols.includes(ci)) : [];
    const vc = valueCols.find((ci) => !lines.includes(ci)) ?? valueCols[0];
    const byKey = colors && colors.rowKey && colors.rowKey.col === legendCol ? colors.rowKey.colorOf : null;
    const legends = byKey ? Object.keys(byKey) : legendOrder(data, legendCol, [catCol], vc);
    // A line value is either the same on every legend row of a category (read once), or one per
    // legend value (Seps SPLY per separation reason: Power BI's line is the month's, their total
    // when the measure adds up; otherwise read once, as before)
    const lineEntries = info && info.roles ? fieldsByColumn(header, info.roles.Y2 || []) : header.map(() => null);
    const perLegend = (ci) => categories.some((c) => new Set(data.filter((x) => catOf(x) === c).map((x) => num(x[ci])).filter((v) => v !== null)).size > 1);
    const once = (ci) => {
      const add = perLegend(ci) && totalKind((lineEntries[ci] && lineEntries[ci].field) || null, header[ci], formatOf(ci)) === "sum";
      return categories.map((c) => {
        const vals = data.filter((x) => catOf(x) === c).map((x) => num(x[ci])).filter((v) => v !== null);
        return !vals.length ? null : add ? vals.reduce((a, b) => a + b, 0) : vals[0];
      });
    };
    // A colour rule per legend value (fx on the series): each legend value's colour from the rule
    const byRule = !byKey && colors && colors.dot && colors.dot.keyCol === legendCol ? colors.dot.colorOf : null;
    // No legend colours worked out (colors): each value's colour set by hand, else the theme's in order
    const fallback = byKey || liveLegend ? null : themeColorsOf(item, legends);
    // A legend with no values at all (every legend cell blank): drawn as if it had no legend
    if (legends.length) return {
      ...base,
      valueName: header[vc],
      legendName: header[legendCol],
      legendSplit: lines.length > 0,   // a combo split by a legend: Power BI draws its line on the columns' scale
      series: [
        ...legends.map((name, i) => ({
          name,
          values: valuesFor(vc, data.filter((r) => key(r, legendCol) === name)),
          fmt: formatOf(vc),
          // A legend found without the layout: the theme's colours in order, as Power BI gives them
          color: (byKey && byKey[name]) || (byRule && byRule(name)) ||
            (liveLegend && !byKey && item.themePalette ? toArgb(item.themePalette[i % item.themePalette.length]) : null) ||
            (fallback && fallback[i]) || null,
        })),
        ...lines.map((ci) => ({ name: header[ci], values: once(ci), fmt: formatOf(ci), color: (colors && colors.series && colors.series[ci]) || null, line: true })),
      ],
    };
  }

  // One series per measure; a combo draws its line values (its last measure, without the
  // report file) as lines on a second axis
  const overrides = (info && info.overrides) || {};
  const palette = item.themePalette || [];
  const series = valueCols.map((ci, i) => ({
    name: header[ci],
    values: valuesFor(ci, data),
    fmt: formatOf(ci),
    // An empty measure has no colour worked out (it has no values): its own, else the theme's for its place
    color: (colors && colors.series && colors.series[ci]) ||
      (empty(ci) && (overrides["m:" + normalizeField(header[ci])] || palette.length) ? toArgb(overrides["m:" + normalizeField(header[ci])] || palette[i % palette.length]) : null),
    // A line chart's secondary y-axis values (Y2) go on the second axis too, as in Power BI
    line: spec.kind === "combo" ? (lineCols ? lineCols.includes(ci) : valueCols.length > 1 && i === valueCols.length - 1)
      : spec.kind === "line" && !!lineCols && lineCols.includes(ci) && lineCols.length < valueCols.length,
  }));
  // A combo showing only line values is a line chart
  if (series.every((s) => s.line)) {
    series.forEach((s) => { s.line = false; });
    base.spec = { kind: "line" };
  }
  // One measure coloured per bar (colours set by hand, or a colour rule)
  let pointColors = null;
  const withValues = valueCols.filter((ci) => !empty(ci));   // an empty measure draws nothing to colour
  if (withValues.length === 1 && colors) {
    const ci = withValues[0];
    if (colors.perCol && colors.perCol[ci]) pointColors = categories.map((c) => colors.perCol[ci][firstRowOf(c)] || null);
    else if (colors.rowKey && colors.rowKey.col === catCol) pointColors = categories.map((c) => colors.rowKey.colorOf[c] || null);
    else if (colors.tabColors) pointColors = categories.map((c) => colors.tabColors[firstRowOf(c)] || null);
  }
  return { ...base, series, pointColors };
}

/* A number format without a minus sign: a tornado's left side reads as positive */
function unsignedNumFmt(fmt) {
  const first = String(fmt || "General").split(";")[0];
  return `${first};${first}`;
}

/*
 * The Tornado visual as it draws itself (its own defaults, else the report's settings): the
 * values inside each bar, the names on the left, a legend without a title, and no value axis,
 * gridlines or axis titles (it has none). Other charts: their look unchanged.
 */
function tornadoLook(look, item, info) {
  if (!(chartSpecOf(item.type) || {}).mirror || !item.chart) return look;
  const set = (key) => customSetting(item.type, info, key);
  const axes = item.chart.axes || {};
  item.chart.axes = { ...axes, cat: { ...(axes.cat || {}), title: false, show: set("categories") !== false }, val: { ...(axes.val || {}), show: false, title: false } };
  return { ...look,
    labels: { ...look.labels, show: set("labels") !== false, position: "InsideCenter" },
    legend: { ...look.legend, title: set("legendTitle") === true } };
}

/* Tornado: its first two sides (legend values or measures), the first drawn to the left as
   negative bars, both labelled without a minus. One side only: a plain bar chart. */
function mirrorPlan(plan, spec) {
  const sides = plan.series.filter((s) => !s.line);
  if (sides.length < 2) return plan;
  const [left, right] = sides;
  return {
    ...plan, spec: { ...spec, hideValueAxis: true },   // the Tornado draws no value axis or gridlines
    series: [
      { ...left, values: left.values.map((v) => (typeof v === "number" ? -v : v)), fmt: unsignedNumFmt(left.fmt) },
      { ...right, fmt: unsignedNumFmt(right.fmt) },
    ],
  };
}

/* Box & whisker: every row's value with its category; Excel's chart works out each category's
   quartiles, whiskers and outliers itself (one box when there is no category) */
function boxPlan(item) {
  const t = chartColumns(item);
  const info = item.colorInfo || null;
  const roles = (info && info.roles) || {};
  const entries = info ? fieldsByColumn(t.header, Object.values(roles).flat()) : t.header.map(() => null);
  const inRole = (ci, ...names) => names.some((r) => (roles[r] || []).includes(entries[ci]));
  const valueCol = t.measureCols.find((ci) => inRole(ci, "Values", "Value", "Y", "Measure")) ?? t.measureCols[t.measureCols.length - 1];
  if (valueCol === undefined) return null;
  // The boxes' category (Jan Pieter Posthuma's visual calls it Groups); its Samples are the points
  // in each box, never a category. With the report file and no category field: one box.
  const grouped = t.textCols.find((ci) => inRole(ci, "Groups", "Category", "Axis", "Group"));
  const catCol = grouped ?? (info && info.roles ? -1 : t.textCols.find((ci) => !inRole(ci, "Samples")) ?? -1);
  const rows = t.data.filter((r) => t.num(r[valueCol]) !== null);
  if (rows.length < 2) return null;
  const fmt = t.formatOf(valueCol);
  // Each category box in its own colour (Excel colours a box & whisker series as one): a series
  // per category, named in the legend, side by side. One colour for every box: one series.
  const catOf = (r) => t.key(r, catCol) || "(Blank)";
  const byKey = item.colors && item.colors.rowKey && item.colors.rowKey.col === catCol ? item.colors.rowKey.colorOf : null;
  const groups = catCol >= 0 ? [...new Set(rows.map(catOf))] : [];
  const colorOf = (g, k) => (byKey && byKey[g]) || (item.colors ? null : toArgb(DEFAULT_THEME_COLORS[k % DEFAULT_THEME_COLORS.length]));
  const oneColor = customSetting(item.type, info, "oneColor") === true || new Set(groups.map((g, k) => colorOf(g, k))).size <= 1;
  if (groups.length > 1 && groups.length <= 12 && !oneColor) {
    const lists = groups.map((g) => rows.filter((r) => catOf(r) === g).map((r) => t.num(r[valueCol])));
    const n = Math.max(...lists.map((l) => l.length));
    return {
      spec: { kind: "boxWhisker" }, perGroup: true,
      title: item.displayTitle || item.title,
      catName: null,   // the legend names the boxes
      valueName: t.header[valueCol], legendName: t.header[catCol],
      axes: (info && info.axes) || null,
      categories: Array.from({ length: n }, () => ""),
      series: groups.map((g, k) => ({ name: g, values: Array.from({ length: n }, (_, i) => (i < lists[k].length ? lists[k][i] : null)), fmt, color: colorOf(g, k) })),
      pointColors: null,
    };
  }
  const color = (groups.length && colorOf(groups[0], 0)) || (item.colors && item.colors.series && item.colors.series[valueCol]) || null;
  return {
    spec: { kind: "boxWhisker" },
    title: item.displayTitle || item.title,
    catName: catCol >= 0 ? t.header[catCol] : null,   // one box: no category axis title
    valueName: t.header[valueCol],
    axes: (info && info.axes) || null,
    categories: rows.map((r) => (catCol >= 0 ? t.key(r, catCol) || "(Blank)" : t.header[valueCol])),
    series: [{ name: t.header[valueCol], values: rows.map((r) => t.num(r[valueCol])), fmt, color }],
    pointColors: null,
  };
}

/*
 * Dates on a line or area chart's axis are spaced by date, as Power BI's continuous axis
 * spaces them (its default for a date field): an Excel date axis with labels every few months
 * or years ("Jul 2023"), and the blank date left out as Power BI leaves it out.
 */
const DATE_STEPS = [[1, "days", 1], [7, "days", 7], [1, "months", 30.44], [2, "months", 60.9], [3, "months", 91.3], [6, "months", 182.6],
  [1, "years", 365.25], [2, "years", 730.5], [5, "years", 1826], [10, "years", 3652]];
function withDateAxis(plan) {
  // a combo's columns and line share the axis: dates spaced by date there too ("May 2021")
  if (!plan || !["line", "area", "combo"].includes(plan.spec.kind) || plan.dateAxis) return plan;
  if (plan.look && plan.look.cat.continuous === false) return plan;
  // A Year › Month (› Day) hierarchy on a continuous axis: Power BI draws it as dates
  // ("Jul 2020"); any other hierarchy keeps its levels of labels
  if (plan.multiLevel && !(plan = hierarchyDates(plan))) return arguments[0];
  const parsed = plan.categories.map((c) => { const t = String(c ?? "").trim(); return t ? parseDate(t) : null; });
  const blank = plan.categories.map((c) => !String(c ?? "").trim());
  if (parsed.filter(Boolean).length < 2 || parsed.some((p, i) => !p && !blank[i])) return plan;
  const pick = (list) => list.filter((_, i) => !blank[i]);
  const serials = pick(parsed).map((p) => Math.round((p.date.getTime() / 86400000 + 25569) * 1e6) / 1e6);
  const span = Math.max(...serials) - Math.min(...serials);
  const most = Math.max(2, ((plan.sizePx ? plan.sizePx.w : 400) - 70) / 85);
  // a year › month hierarchy steps by whole months at least, and its points read "Jul 2020"
  const steps = plan.monthsOnly ? DATE_STEPS.filter(([, t]) => t !== "days") : DATE_STEPS;
  const [unit, timeUnit, days] = steps.find(([, , d]) => span / d <= most) || steps[steps.length - 1];
  const fmt = timeUnit === "years" ? "yyyy" : timeUnit === "months" ? "mmm yyyy" : "d mmm";
  return {
    ...plan,
    categories: pick(plan.categories),
    series: plan.series.map((s) => ({ ...s, values: pick(s.values) })),
    pointColors: plan.pointColors ? pick(plan.pointColors) : null,
    dateAxis: withLastLabelRoom({ serials, unit, timeUnit, days, ...alignedDateTicks(serials, unit, timeUnit, fmt), ...(plan.monthsOnly ? { cellFmt: "mmm yyyy" } : {}) },
      fmt, plan.sizePx, plan.look && plan.look.cat && plan.look.cat.text),
  };
}

/*
 * Excel centres each date label under its date and cuts off one that runs past the chart's
 * right edge ("May 20…"). When the last label sits that close to the end, the axis runs on
 * just far enough for all of it to show (Power BI keeps that room beside its plot too).
 */
function withLastLabelRoom(axis, fmt, sizePx, text) {
  if (!["months", "years"].includes(axis.timeUnit) && axis.unit !== 7 && axis.unit !== 1) return axis;
  const first = Math.min(...axis.serials), last = Math.max(...axis.serials);
  const min = axis.min ?? first;
  const toDate = (v) => new Date(Math.round((v - 25569) * 86400000));
  const serial = (d) => Math.round(d.getTime() / 86400000 + 25569);
  // the labels' dates: from the axis's start, a step at a time, up to the last point
  let t = toDate(min), lastTick = null;
  for (let k = 0; k < 2000; k++) {
    const v = serial(t);
    if (v > last) break;
    lastTick = t;
    t = axis.timeUnit === "years" ? new Date(Date.UTC(t.getUTCFullYear() + axis.unit, t.getUTCMonth(), t.getUTCDate()))
      : axis.timeUnit === "months" ? new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + axis.unit, t.getUTCDate()))
        : new Date(t.getTime() + axis.unit * 86400000);
  }
  if (!lastTick || last <= min) return axis;
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const y = lastTick.getUTCFullYear(), mo = MON[lastTick.getUTCMonth()], dd = lastTick.getUTCDate();
  const label = /^yyyy$/.test(fmt) ? String(y) : /mmm yyyy/.test(fmt) ? `${mo} ${y}` : `${dd} ${mo}`;
  const half = textPx(label, text || {}) / 2 + 6;
  const plotPx = Math.max(120, ((sizePx && sizePx.w) || 400) - 90);
  // the end at which the last label's half fits: (end - tick) / (end - min) × plot = half
  const tick = serial(lastTick);
  const end = (tick - (half / plotPx) * min) / (1 - half / plotPx);
  if (!(end > last)) return axis;
  return { ...axis, min, max: Math.ceil(end) };
}

/*
 * Power BI puts a date axis's labels on round dates, not on the first point: every 2 months is
 * Jan, Mar, May, Jul, Sep, Nov (Jul 2020, Sep 2020 … for points from Jun 2020), every 3 months
 * Jan, Apr, Jul, Oct, every 6 Jan and Jul, every n years the years divisible by n, every week
 * Sundays. Excel labels from the axis's start, so the axis starts at the round date at or before
 * the first point, and a label there before the first point is left blank.
 */
function alignedDateTicks(serials, unit, timeUnit, fmt) {
  const first = Math.min(...serials), last = Math.max(...serials);
  const d = new Date(Math.round((Math.floor(first) - 25569) * 86400000));
  const y = d.getUTCFullYear(), m = d.getUTCMonth();
  let start;
  if (timeUnit === "years") start = Date.UTC(y - (((y % unit) + unit) % unit), 0, 1);
  else if (timeUnit === "months") start = Date.UTC(y, m - (m % unit), 1);
  else if (unit === 7) start = Date.UTC(y, m, d.getUTCDate() - d.getUTCDay());
  else return { fmt };
  const min = Math.round(start / 86400000 + 25569);
  if (!(min < first)) return { fmt };
  return { fmt: `[<${Math.floor(first)}]"";${fmt}`, min, max: last };
}

/*
 * A hierarchy whose levels are a year, a month (name or number) and optionally a quarter or a
 * day, as dates in date order: the plan with one date per point; null for any other hierarchy
 */
function hierarchyDates(plan) {
  const levels = [...(plan.catLevels || []).map((l) => l.values), plan.catText || []];
  const n = plan.categories.length;
  if (levels.length < 2 || levels.some((l) => l.length !== n)) return null;
  const kinds = levels.map((vals) => {
    const t = vals.map((v) => String(v ?? "").trim());
    if (t.every((v) => /^\d{4}$/.test(v) && v >= 1900 && v <= 2100)) return "year";
    if (t.every((v) => MONTHS.includes(v.slice(0, 3).toLowerCase()) && /^[a-z]+$/i.test(v))) return "month";
    if (t.every((v) => /^(q|qtr|quarter)\s*[1-4]$/i.test(v))) return "quarter";
    if (t.every((v) => /^\d{1,2}$/.test(v) && v >= 1 && v <= 31)) return "number";
    return null;
  });
  const year = kinds.indexOf("year");
  // a number level is the month right after the year (or quarter), else the day after the month
  let month = kinds.indexOf("month");
  const nums = kinds.map((k, i) => (k === "number" ? i : -1)).filter((i) => i >= 0);
  if (month < 0 && nums.length && levels[nums[0]].every((v) => Number(v) <= 12)) month = nums.shift();
  const day = nums.length === 1 && nums[0] > month ? nums[0] : -1;
  if (year < 0 || month < 0 || month < year || nums.length > (day >= 0 ? 1 : 0) || kinds.some((k) => k === null)) return null;
  const monthOf = (v) => (/^\d+$/.test(String(v).trim()) ? Number(v) : MONTHS.indexOf(String(v).trim().slice(0, 3).toLowerCase()) + 1);
  const pad = (x) => String(x).padStart(2, "0");
  const dates = plan.categories.map((_, i) => `${levels[year][i]}-${pad(monthOf(levels[month][i]))}-${pad(day >= 0 ? Number(levels[day][i]) : 1)}`);
  if (dates.some((d) => !parseDate(d)) || new Set(dates).size !== n) return null;
  const order = dates.map((_, i) => i).sort((a, b) => (dates[a] < dates[b] ? -1 : dates[a] > dates[b] ? 1 : 0));
  const pick = (list) => (Array.isArray(list) ? order.map((i) => list[i]) : list);
  return { ...plan, multiLevel: false, catLevels: null, catText: null, categories: pick(dates), pointColors: pick(plan.pointColors),
    series: plan.series.map((s) => ({ ...s, values: pick(s.values) })), monthsOnly: day < 0 };
}

/*
 * Numbers on a line or area chart's axis (Period 1…12) run in ascending order, as on Power BI's
 * continuous axis, whatever order the rows were exported in (sorted by the measure, say)
 */
function withNumberOrder(plan) {
  if (!plan || !["line", "area"].includes(plan.spec.kind) || plan.multiLevel) return plan;
  if (plan.look && plan.look.cat.continuous === false) return plan;
  const nums = plan.categories.map((c) => { const t = String(c ?? "").trim().replace(/,/g, ""); return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : null; });
  if (nums.length < 2 || nums.some((n) => n === null)) return plan;
  const order = nums.map((_, i) => i).sort((a, b) => nums[a] - nums[b]);
  if (order.every((i, k) => i === k)) return plan;
  const pick = (list) => (Array.isArray(list) ? order.map((i) => list[i]) : list);
  return { ...plan, categories: pick(plan.categories), catText: pick(plan.catText), pointColors: pick(plan.pointColors),
    series: plan.series.map((s) => ({ ...s, values: pick(s.values) })) };
}

/*
 * Charts mode: a Dual KPI as the visual draws it: its two fields (top, then bottom) as area
 * charts over the dates, stacked, each in its colour with its axis labelled at its lowest and
 * highest value (or the axis range set in the visual), the field's name, change and latest value
 * over its middle ("Total Sales (+27.2%)", "$1,192"), and "Last 670 days" under them. The data
 * is the trend cards' (dualKpiTrend), in date order.
 */
const DUAL_KPI_FOOT_PX = 18;
function trendChart(item) {
  const t = item.trend;
  // a card without its trend (the KPI visual's "Trend axis" off) stays a card
  if (!t || !t.cards.length || t.points < 2 || t.cards.some((c) => c.noSpark)) return null;
  const [header, ...rows] = item.rows;
  const pts = rows.slice(0, t.points);
  const series = t.cards.map((card) => ({ name: card.label, values: pts.map((r) => { const v = toCellValue(String(r[card.col] ?? "")).value; return typeof v === "number" ? v : null; }),
    fmt: (pts.map((r) => toCellValue(String(r[card.col] ?? ""))).find((c) => typeof c.value === "number") || {}).numFmt || "#,##0", color: card.color ? toArgb(card.color) : null }));
  return { spec: { kind: "area", grouping: "standard" }, title: item.displayTitle || item.title, catName: header[t.dateCol],
    categories: pts.map((r) => String(r[t.dateCol] ?? "")), series, pointColors: null, axes: null, trendCards: true };
}
/* The KPI visual's numbers: display units Auto, two decimals (114.65K, $1.20M) */
function kpiShort(v, fmt) {
  if (typeof v !== "number") return "(Blank)";
  const unit = autoUnit(Math.abs(v));
  if (unit <= 1) return displayText(v, fmt);
  const prefix = fmtPrefix(String(fmt || "")).replace(/^"|"$/g, "").replace(/^\\/, "");
  return `${v < 0 ? "-" : ""}${prefix}${(Math.abs(v) / unit).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}${PBI_UNITS[unit] ? PBI_UNITS[unit][0] : ""}`;
}
function trendPanels(item, look) {
  const base = item.chart, t = item.trend;
  const total = base.sizePx || chartSizePx(item);
  const head = base.showTitle && base.title ? PANEL_TITLE_PX : 0;
  const foot = DUAL_KPI_FOOT_PX;
  const n = t.cards.length;
  const w = total.w, h = Math.floor((total.h - head - foot) / n);
  const plainLook = { ...look, title: { ...look.title, show: false }, legend: { ...look.legend, show: false }, labels: { ...look.labels, show: false },
    cat: { ...look.cat, grid: null }, val: { ...look.val, grid: null }, frame: { ...look.frame, border: null, shadow: null }, refLines: [], trend: null, totals: { ...look.totals, show: false },
    areaOpacity: 0.3, lines: { ...look.lines, width: 0.75, style: "solid", show: true, markers: null, series: {} } };
  const valueFont = { ...LOOK_TEXT, size: 20, color: "#252423" };
  const labelFont = { ...LOOK_TEXT, size: 10, color: "#252423" };
  const plans = t.cards.map((card, k) => {
    const s = base.series[k];
    const vals = s.values.filter((v) => typeof v === "number");
    const lo = card.min ?? Math.min(...vals), hi = card.max ?? Math.max(...vals);
    // The KPI visual: no axis; its value in the status colour with the status mark, the goal under it
    const kpi = card.kpi;
    const lines = kpi
      ? [{ text: kpiShort(kpi.actual, kpi.fmt) + (kpi.status === "bad" ? " !" : kpi.status === "good" ? " \u2713" : ""), font: { ...valueFont, color: card.valueColor || valueFont.color } },
        ...(kpi.goalText && kpi.goals.length ? [{ text: `${kpi.goalText}: ${kpi.goals.map((g) => kpiShort(g, kpi.fmt)).join(", ")}${kpi.distance ? ` ${kpi.distance}` : ""}`, font: { ...labelFont, size: 9 } }] : [])]
      : [{ text: card.change ? `${card.label} (${card.change})` : card.label, font: labelFont }, { text: card.text, font: valueFont }];
    return { spec: base.spec, title: card.label, catName: base.catName, categories: base.categories, pointColors: null,
      series: [kpi && card.color ? { ...s, color: toArgb(card.color) } : s], look: plainLook, sizePx: { w, h }, showTitle: false,
      axes: { cat: { show: false, title: false }, val: { show: !kpi, title: false } },
      axisEnds: hi > lo ? { min: lo, max: hi } : null,
      overlay: { height: 44, lines } };
  });
  const c = t.caption || {};
  const footer = c.days !== undefined ? { text: `Last ${c.days} days`, align: "right", font: { ...LOOK_TEXT, size: 9, italic: true, color: "#605E5C" } } : null;
  return { plans, cols: 1, rows: n, w, h, head, foot: footer ? foot : 0, footer, title: head ? base.title : null, look };
}

/* Small multiples: each panel's chart (its own colours, the visual's look, centred value as its
   title, no frame) sized to its cell of the grid under the title; null when any can't be drawn */
const PANEL_TITLE_PX = 30;
function panelPlans(item, info, look, palette) {
  const { cols, rows } = item.panelGrid;
  const total = item.chart.sizePx || chartSizePx(item);
  const head = item.chart.showTitle && item.chart.title ? PANEL_TITLE_PX : 0;
  const w = Math.floor(total.w / cols), h = Math.floor((total.h - head) / rows);
  const panelLook = { ...look, title: { ...look.title, align: "center", show: true, background: null }, frame: { ...look.frame, border: null, shadow: null } };
  const plans = [];
  for (const pn of item.panels) {
    const part = { ...item, rows: pn.rows, panels: null };
    part.colors = palette ? columnColors(part, palette, info) : null;
    let plan = chartPlan(part);
    if (!plan) return null;
    // Power BI draws the panels' axes once, on the grid's outer edge: the bottom row, the left column
    const k = plans.length, count = item.panels.length;
    Object.assign(plan, { title: pn.value, look: panelLook, sizePx: { w, h }, showTitle: true,
      panelEdges: { bottom: k + cols >= count, left: k % cols === 0 }, panelTop: item.sharedMax ?? null });
    plan = withDateAxis(withNumberAxis(withNumberOrder(plan)));
    if (chartBuildError(plan)) return null;
    plans.push(plan);
  }
  return { plans, cols, rows, w, h, head, title: head ? item.chart.title : null, look };
}

/*
 * A chart whose legend can't show every entry in its size (a pie of 37 products) is made taller
 * by the rows its legend needs, so every entry shows and the chart itself keeps its room; at
 * most three times its height, beyond which the legend shows the first entries that fit.
 */
const LEGEND_GROWTH_MAX = 3;
function growForLegend(plan) {
  const ref = "'Data'!$A$1:$A$2";
  const need = (p) => {
    try {
      return chartXml({ ...p, refs: { catRef: ref, series: (p.series || []).map(() => ({ nameRef: "'Data'!$A$1", valRef: ref })), range: () => ref } }).legendNeed;
    } catch (e) {
      return null;
    }
  };
  const n = need(plan);
  if (!n || !n.cut || !plan.sizePx) return plan;
  const h = Math.ceil(plan.sizePx.h + n.extraRows * n.rowH + 8);
  if (h > plan.sizePx.h * LEGEND_GROWTH_MAX) return plan;
  return { ...plan, fullLegend: true, sizePx: { ...plan.sizePx, h } };
}

/* Builds a chart's part once with stand-in data references (the real ones come with the Data
   sheet): the error when it can't be built, else null */
function chartBuildError(plan) {
  const ref = "'Data'!$A$1:$A$2";
  const test = { ...plan, refs: { catRef: ref, series: (plan.series || []).map(() => ({ nameRef: "'Data'!$A$1", valRef: ref })), range: () => ref } };
  try {
    const kind = plan.spec.kind;
    if (kind === "treemap" || kind === "sunburst") treemapXml(test, 1);
    else if (kind === "boxWhisker") boxWhiskerXml(test);
    else chartXml(test);
    return null;
  } catch (e) {
    return e;
  }
}

/*
 * Whole numbers on a line or area chart's axis (RecoveryDay 1, 3, 7, 14, 30, 60) are spaced by
 * value, as on Power BI's continuous axis: Excel's date axis takes the numbers as its positions
 * and shows them as numbers, on round steps up to the highest (0, 20, 40, 60). Negative or
 * fractional numbers, or an axis the report sets to Categorical, keep the category axis.
 */
function withNumberAxis(plan) {
  if (!plan || !["line", "area"].includes(plan.spec.kind) || plan.multiLevel || plan.dateAxis) return plan;
  if (plan.look && plan.look.cat.continuous === false) return plan;
  const nums = plan.categories.map((c) => { const t = String(c ?? "").trim().replace(/,/g, ""); return /^\d+$/.test(t) ? Number(t) : null; });
  if (nums.length < 2 || nums.some((n) => n === null || n > 2958465)) return plan;
  const plotPx = Math.max(120, ((plan.sizePx && plan.sizePx.w) || 400) - 90);
  const scale = valueScale(nums, plotPx, 90, { zero: false, round: true });
  if (!scale || !(scale.step > 0)) return plan;
  // from the step at or below the lowest number (ticks then fall on round numbers) to the highest
  const lo = Math.max(0, Math.floor(Math.min(...nums) / scale.step) * scale.step);
  const hi = Math.max(...nums);
  const years = nums.every((n) => n >= 1900 && n <= 2100);
  // Power BI's axis starts at the lowest number (Period 1, labels 2, 4 … 12): Excel's starts on
  // the step below it, and the label there, before the first point, is left blank
  const shown = years ? "0" : "#,##0";
  const first = Math.min(...nums);
  return { ...plan, dateAxis: { serials: nums, unit: scale.step, timeUnit: "days", fmt: lo < first ? `[<${first}]"";${shown}` : shown, cellFmt: shown,
    min: lo, max: hi > lo ? hi : lo + scale.step, numbers: true } };
}

/* Shared helpers for the chart builders below */
/* Without the report's layout: a chart's legend when the export has more text columns than
   its axis (Region, VP, Ethnicity: VP a level drilled out of view). The live view names the
   axis field; Power BI exports the legend after the axis levels, so the last text column is
   the legend unless it is on the axis. -1 when the live view doesn't say. */
function liveLegendCol(item, header, textCols) {
  const axis = ((item.liveDrill && item.liveDrill.activeProjections) || {}).Category;
  if (!Array.isArray(axis) || !axis.length) return -1;
  const onAxis = new Set(axis.map((e) => fieldInfo(e)).filter(Boolean).map((f) => normalizeField(f.prop)));
  if (!onAxis.has(normalizeField(header[textCols[0]]))) return -1;   // the axis isn't where expected
  const last = textCols[textCols.length - 1];
  return onAxis.has(normalizeField(header[last])) ? -1 : last;
}

/* Colours for values (legend values, slices) the colour rules gave none: each one's colour set
   by hand in the report, else one for every value, else the theme's in order, as Power BI draws
   them. null without a theme (Excel's own colours, as before). */
function themeColorsOf(item, values) {
  const palette = item.themePalette;
  const overrides = (item.colorInfo && item.colorInfo.overrides) || {};
  if (!palette || !palette.length || !values.length) return null;
  return values.map((v, i) => toArgb(overrides["i:" + normalizeValue(v)] || overrides.all || palette[i % palette.length]));
}

function chartColumns(item) {
  const c = classifyColumns(item.rows);
  const cols = c.header.map((_, ci) => ci);
  const key = (r, ci) => String(r[ci] ?? "").trim();
  const num = (raw) => { const v = toCellValue(raw).value; return typeof v === "number" ? v : null; };
  const formatOf = (ci) => {
    const cells = c.data.map((r) => [toCellValue(r[ci])]);
    const readable = readableFormats([c.header[ci]], cells)[0];
    const exported = cells.map((x) => x[0]).find((x) => typeof x.value === "number" && x.numFmt);
    return readable || (exported && exported.numFmt) || "General";
  };
  return {
    ...c, cols, key, num, formatOf,
    textCols: cols.filter((ci) => !c.numeric[ci] && !c.hex[ci]),
    measureCols: cols.filter((ci) => c.numeric[ci] && !c.hex[ci] && !c.helper[ci]),
  };
}

/* One value per category, in the order they appear ("Total" rows are left out). levels: the
   category's levels shown together ("East New York"), else just catCol */
function chartSteps(t, catCol, valueCol, levels = null) {
  const additive = !NON_ADDITIVE.test(t.header[valueCol]);
  const steps = [];
  for (const r of t.data) {
    const label = levels ? levelsLabel(r, levels) : t.key(r, catCol), v = t.num(r[valueCol]);
    if (!label || v === null || /^total$/i.test(label)) continue;
    const s = steps.find((x) => x.label === label);
    if (!s) steps.push({ label, value: v });
    else if (additive) s.value += v;
  }
  return steps;
}

/* Scatter (or bubble, when there is a Size field): one Excel series per legend value */
function scatterPlan(item) {
  const t = chartColumns(item);
  const { header, data, textCols, key, num } = t;
  const info = item.colorInfo || null, colors = item.colors || null;
  const colOf = (f) => (f ? header.findIndex((h) => String(h).trim().toLowerCase() === f.label || normalizeField(h) === f.ref) : -1);
  let legendCol = info && info.rule === "legend" ? colOf(info.series) : -1;
  if (legendCol < 0 && !info && textCols.length >= 2) legendCol = textCols[1];
  let labelCol = info ? colOf(info.category) : -1;
  if (labelCol < 0 || labelCol === legendCol) labelCol = textCols.find((ci) => ci !== legendCol) ?? -1;
  const nums = t.measureCols.filter((ci) => ci !== legendCol && ci !== labelCol);
  if (nums.length < 2) return null;
  const [xc, yc, sc] = scatterFields(t, item, info, nums);
  if (xc === undefined || yc === undefined) return null;
  const valueCols = sc !== undefined ? [xc, yc, sc] : [xc, yc];
  const rows = data.filter((r) => num(r[xc]) !== null && num(r[yc]) !== null);
  if (!rows.length) return null;
  const groupOf = (r) => (legendCol >= 0 ? key(r, legendCol) : "");

  // One bubble per name (and legend value). Power BI can export finer rows than the chart
  // draws (a field in the tooltips, say): those rows are combined. Amounts are added up;
  // averages, rates and percentages are averaged, weighted by the bubble size when there is one.
  const fmts = valueCols.map((ci) => t.formatOf(ci));
  const averaged = valueCols.map((ci, k) => NON_ADDITIVE.test(header[ci]) || /\bper\b/i.test(header[ci]) || fmts[k].includes("%"));
  const weighted = sc !== undefined && !averaged[2];
  const combine = (list, k) => {
    const ci = valueCols[k];
    const vals = list.map((r) => [num(r[ci]), weighted ? num(r[sc]) || 0 : 1]).filter(([v]) => v !== null);
    if (!vals.length) return null;
    if (!averaged[k]) return vals.reduce((a, [v]) => a + v, 0);
    const w = vals.reduce((a, [, wt]) => a + wt, 0);
    return w ? vals.reduce((a, [v, wt]) => a + v * wt, 0) / w : vals.reduce((a, [v]) => a + v, 0) / vals.length;
  };
  let merged = false;
  let points;
  if (labelCol >= 0) {
    const byKey = new Map();
    for (const r of rows) {
      const k = groupOf(r) + "\u0001" + key(r, labelCol);
      if (!byKey.has(k)) byKey.set(k, []);
      byKey.get(k).push(r);
    }
    points = [...byKey.values()].map((list) => {
      if (list.length > 1) merged = true;
      return { group: groupOf(list[0]), label: key(list[0], labelCol), values: valueCols.map((_, k) => combine(list, k)) };
    });
  } else {
    points = rows.map((r, i) => ({ group: groupOf(r), label: String(i + 1), values: valueCols.map((ci) => num(r[ci])) }));
  }
  const names = [...new Set(points.map((p) => p.group))];
  points = names.flatMap((n) => points.filter((p) => p.group === n));

  const fallback = (colors && colors.series && Object.values(colors.series)[0]) || toArgb(DEFAULT_THEME_COLORS[0]);
  const colorOf = (n) => (colors && colors.rowKey && colors.rowKey.colorOf[n]) || fallback;
  // Without a legend, each bubble is its own Excel series, so its name shows beside it and in
  // Excel's hover tooltip, like Power BI (up to 60 bubbles, to stay readable)
  const perPoint = legendCol < 0 && labelCol >= 0 && points.length <= 60;
  let groups;
  if (perPoint) {
    groups = points.map((p, i) => ({ name: p.label, start: i, end: i, color: colorOf(p.group) }));
  } else {
    let at = 0;
    groups = names.map((n) => {
      const count = points.filter((p) => p.group === n).length;
      const g = { name: n || header[yc], start: at, end: at + count - 1, color: colorOf(n) };
      at += count;
      return g;
    });
  }
  const title = item.displayTitle || item.title;
  return {
    spec: { kind: sc !== undefined ? "bubble" : "scatter" },
    title,
    catName: labelCol >= 0 ? header[labelCol] : "Point",
    categories: points.map((p) => p.label),
    series: valueCols.map((ci, k) => ({ name: header[ci], values: points.map((p) => p.values[k]), fmt: fmts[k], color: null })),
    groups,
    perPoint,
    legendName: legendCol >= 0 ? header[legendCol] : null,
    sizePx: chartSizePx(item),   // the visual's size on the report page, for Power BI-like axis steps
    pointColors: null,
    note: merged ? `"${title}" shows one bubble per ${header[labelCol]}, but Power BI exported several rows for some of them. They were combined, so averages and percentages in that chart are approximate.` : null,
  };
}
/* Waterfall: stacked columns on an invisible base, in increase / decrease / total colours.
   Steps that cross zero are split into a part above and a part below zero. */
function waterfallPlan(item) {
  const t = chartColumns(item);
  const catCol = t.textCols.length ? t.textCols[0] : -1;
  const valueCol = t.measureCols.find((ci) => ci !== catCol);
  if (catCol < 0 || valueCol === undefined) return null;
  // Levels shown together: a step per Region and State ("East New York")
  const steps = chartSteps(t, catCol, valueCol, expandedLevelCols(item, item.colorInfo || null, t.header));
  if (!steps.length) return null;
  const w = (item.colorInfo && item.colorInfo.waterfall) || {};
  const up = toArgb(w.increase || "#1AAB40"), down = toArgb(w.decrease || "#D64554"), total = toArgb(w.total || (item.themePalette || [])[0] || "#118DFF");
  const fmt = t.formatOf(valueCol);
  const negFmt = `"-"${fmt}`;   // a decrease is stored as its size and shown with a minus: -25.36%
  const starts = [], ends = [];
  let running = 0;
  steps.forEach((s) => { starts.push(running); running += s.value; ends.push(running); });
  // The axis like Power BI's (0M to 2M in 1M steps for a 2.47M total), on to the next step
  // when the labels need the room
  const size = chartSizePx(item);
  const tidy = (x) => Number(x.toFixed(12));
  const lookLabels = item.colorInfo && item.colorInfo.look ? item.colorInfo.look.labels : null;
  const showLabels = !!(lookLabels && lookLabels.show === true);   // Power BI's waterfall has no labels unless the report turns them on
  const axis = valueScale([0, ...starts, ...ends], size.h - 100, 75, { zero: true, round: showLabels }) || { min: Math.min(0, ...ends), max: Math.max(0, ...ends), step: 0 };
  // Each label needs room outside its bar (above an increase, below a decrease). Where the
  // axis edge leaves too little, the label goes inside the end of the bar in white, as Power
  // BI does; a bar too short to hold it makes the axis one step longer instead. The total is
  // labelled inside its bar when it is tall enough, like Power BI.
  const points = [...steps.map((s, i) => ({ v: s.value, end: ends[i], total: false })), { v: running, end: running, total: true }];
  const inside = points.map(() => false);
  if (showLabels && axis.step) {
    const plotPx = Math.max(60, size.h - 100), LABEL_PX = 16;   // the plot's height; a label's height
    const px = (v) => (Math.abs(v) * plotPx) / (axis.max - axis.min);
    points.forEach((p, i) => {
      const tall = px(p.v) >= LABEL_PX + 4;
      if (p.total && tall) { inside[i] = true; return; }
      const room = p.v >= 0 ? axis.max - p.end : p.end - axis.min;
      if (px(room) >= LABEL_PX) return;
      if (tall) inside[i] = true;
      else if (p.v >= 0) axis.max = tidy(axis.max + axis.step);
      else axis.min = tidy(axis.min - axis.step);
    });
  }
  // Every bar is drawn as ONE column holding the step's real size, so hovering shows the true
  // value. The columns stand on an invisible base, measured from the bottom of the axis
  // (shift), which keeps every value above zero: a step that crosses zero is never split.
  const shift = -axis.min;
  const base = [], inc = [], dec = [], tot = [];
  steps.forEach((s, i) => {
    base.push(tidy(Math.min(starts[i], ends[i]) + shift));
    inc.push(s.value >= 0 ? tidy(s.value) : null);
    dec.push(s.value < 0 ? tidy(-s.value) : null);
    tot.push(null);
  });
  base.push(tidy(Math.min(0, running) + shift)); inc.push(null); dec.push(null); tot.push(tidy(Math.abs(running)));
  return {
    spec: { kind: "waterfall" },
    title: item.displayTitle || item.title,
    catName: t.header[catCol],
    valueName: t.header[valueCol],
    axes: (item.colorInfo && item.colorInfo.axes) || null,
    categories: [...steps.map((s) => s.label), "Total"],
    series: [
      { name: " ", values: base, fmt: ";;;", hidden: true, noLegend: true },   // ";;;" keeps the base out of tooltips
      { name: "Increase", values: inc, fmt, color: up },
      { name: "Decrease", values: dec, fmt: negFmt, color: down },
      { name: "Total", values: tot, fmt: running < 0 ? negFmt : fmt, color: total },
      { name: t.header[valueCol], values: [...ends, running], fmt, hidden: true, noLegend: true },   // carries the labels
      { name: t.header[valueCol], values: [...steps.map((s) => s.value), running], fmt, hidden: true, noLegend: true, dataOnly: true },   // Excel's own waterfall
    ],
    steps: [...steps.map((s) => s.value), running],
    axis, fmt,
    showLabels,
    labelInside: inside,   // per bar: label inside the bar's end (white) instead of beside it
    labelPrecision: lookLabels && lookLabels.precision !== null ? lookLabels.precision : undefined,
    pointColors: null,
  };
}

/* Funnel: centred bars (an invisible offset pushes each bar to the middle), largest on top */
function funnelPlan(item) {
  const t = chartColumns(item);
  const catCol = t.textCols.length ? t.textCols[0] : -1;
  const valueCol = t.measureCols.find((ci) => ci !== catCol);
  if (catCol < 0 && t.measureCols.length > 1) return measureFunnelPlan(item, t);
  if (catCol < 0 || valueCol === undefined) return null;
  // Levels shown together: a stage per Region and State ("East New York")
  const levels = expandedLevelCols(item, item.colorInfo || null, t.header);
  const stages = chartSteps(t, catCol, valueCol, levels);
  if (!stages.length || stages.some((s) => s.value < 0)) return null;
  const max = Math.max(...stages.map((s) => s.value));
  const colors = item.colors || null;
  const color = (colors && colors.series && colors.series[valueCol]) || null;
  const pointColors = colors && colors.rowKey && colors.rowKey.col === catCol
    ? stages.map((s) => colors.rowKey.colorOf[s.label] || color) : null;
  const fmt = t.formatOf(valueCol);
  return {
    spec: { kind: "bar", dir: "bar", grouping: "stacked", hideValueAxis: true },
    funnel: true,   // also written as Excel's own funnel chart (addChartsToFile), these bars its fallback
    title: item.displayTitle || item.title,
    catName: levels ? levels.map((ci) => t.header[ci]).join(", ") : t.header[catCol],
    categories: stages.map((s) => s.label),
    series: [
      { name: "Offset", values: stages.map((s) => (max - s.value) / 2), fmt, color: null, hidden: true, noLegend: true },
      { name: t.header[valueCol], values: stages.map((s) => s.value), fmt, color, labels: true },
    ],
    pointColors,
  };
}

/* A funnel of measures with no category (Sales, Cost, Profit): a stage per measure, in the
   visual's field order, each in its measure's colour. Its export is one row; anything else
   keeps the visual as a table. */
function measureFunnelPlan(item, t) {
  if (t.data.length !== 1) return null;
  const cols = t.measureCols;
  const values = cols.map((ci) => t.num(t.data[0][ci]));
  if (values.some((v) => v === null || v < 0)) return null;
  const max = Math.max(...values);
  const colors = item.colors || null;
  const palette = item.themePalette || [];
  const pointColors = cols.map((ci, i) => (colors && colors.series && colors.series[ci]) || (palette.length ? toArgb(palette[i % palette.length]) : null));
  const fmt = t.formatOf(cols[0]);
  return {
    spec: { kind: "bar", dir: "bar", grouping: "stacked", hideValueAxis: true },
    funnel: true,
    title: item.displayTitle || item.title,
    catName: "Measure",
    categories: cols.map((ci) => t.header[ci]),
    series: [
      { name: "Offset", values: values.map((v) => (max - v) / 2), fmt, color: null, hidden: true, noLegend: true },
      { name: "Value", values, fmt, color: null, labels: true },
    ],
    pointColors: pointColors.some(Boolean) ? pointColors : null,
  };
}

/* A visual's category levels shown together ("Expand all down one level": a tile, slice or
   stage per Region and State, labelled "East New York"): their columns in the hierarchy's
   order, when the report marks at least two of them shown and the export holds them; else
   null (with no drill state a visual shows its top level, as before). A treemap's levels are
   its Group (not with Details, which split its tiles instead); the others' their Category. */
function expandedLevelCols(item, info, header) {
  if (!info || !info.roles || customVisual(item.type)) return null;
  const role = item.type === "treemap" ? "Group" : "Category";
  if (role === "Group" && (info.roles.Details || []).length) return null;
  const shown = (info.roles[role] || []).filter((f) => f.active === true);
  if (shown.length < 2) return null;
  const entries = fieldsByColumn(header, shown);
  const cols = shown.map((f) => entries.indexOf(f)).filter((ci) => ci >= 0);
  return cols.length > 1 ? cols : null;
}

/* A row's label at its levels shown together, joined as Power BI joins them ("East New York") */
function levelsLabel(row, levels) {
  return levels.map((ci) => String(row[ci] ?? "").trim()).filter(Boolean).join(" ");
}

/*
 * A Power BI treemap as an Excel treemap: a tile per group (Group), split into its Details when
 * the visual has them, sized by the value, each group in its colour from the report (its
 * Details share it, as in Power BI), largest first. Tiles can't be negative: those rows are left
 * out of the chart (the table still lists them in Data tables mode).
 */
function treemapPlan(item, spec) {
  const t = chartColumns(item);
  const info = item.colorInfo || null;
  const roles = (info && info.roles) || {};
  const entries = info ? fieldsByColumn(t.header, Object.values(roles).flat()) : t.header.map(() => null);
  const inRole = (ci, ...names) => names.some((r) => (roles[r] || []).includes(entries[ci]));
  const groupCol = t.textCols.find((ci) => !info || inRole(ci, "Group", "Category")) ?? t.textCols[0];
  if (groupCol === undefined) return null;
  // Group levels shown together ("East New York"): each combination is one tile
  const levels = spec.kind === "treemap" ? expandedLevelCols(item, info, t.header) : null;
  const groupOf = (r) => (levels ? levelsLabel(r, levels) : t.key(r, groupCol));
  // A treemap's tiles are its Details; a sunburst's outer ring is its second node level
  const detailCol = levels ? undefined
    : t.textCols.find((ci) => ci !== groupCol && (!info || inRole(ci, "Details") || (spec.kind === "sunburst" && inRole(ci, "Category", "Group"))));
  const valueCol = t.measureCols.find((ci) => !info || inRole(ci, "Values", "Y", "Size")) ?? t.measureCols[0];
  if (valueCol === undefined) return null;
  const additive = !NON_ADDITIVE.test(t.header[valueCol]);
  // One tile per group (and detail): a hierarchy's rows added up, a ratio's first value kept
  const tiles = [];
  for (const r of t.data) {
    const group = groupOf(r), detail = detailCol !== undefined ? t.key(r, detailCol) : "";
    const v = t.num(r[valueCol]);
    if (v === null) continue;
    const tile = tiles.find((x) => x.group === group && x.detail === detail);
    if (!tile) tiles.push({ group, detail, value: v });
    else if (additive) tile.value += v;
  }
  const shown = tiles.filter((x) => x.value > 0);
  if (!shown.length) return null;
  // Largest group first, its largest tile first (Power BI's order)
  const totals = {};
  shown.forEach((x) => { totals[x.group] = (totals[x.group] || 0) + x.value; });
  // (a sunburst keeps the data's order, clockwise from the top, as the visual draws it)
  if (spec.kind !== "sunburst") shown.sort((a, b) => (totals[b.group] - totals[a.group]) || (a.group < b.group ? -1 : a.group > b.group ? 1 : 0) || (b.value - a.value));
  else { const first = [...new Set(shown.map((x) => x.group))]; shown.sort((a, b) => first.indexOf(a.group) - first.indexOf(b.group)); }
  const colors = item.colors || null;
  const byGroup = levels ? (colors && colors.tiles ? colors.tiles.colorOf : null)
    : colors && colors.rowKey && colors.rowKey.col === groupCol ? colors.rowKey.colorOf : null;
  const one = colors && colors.series && colors.series[valueCol] || null;
  const label = (v) => axisLabel(v);
  return {
    spec,
    title: item.displayTitle || item.title,
    catName: levels ? levels.map((ci) => t.header[ci]).join(", ") : t.header[detailCol ?? groupCol],
    catLevels: detailCol !== undefined ? [{ name: t.header[groupCol], values: shown.map((x) => label(x.group)) }] : [],
    categories: shown.map((x) => (detailCol !== undefined ? x.detail : x.group)),
    series: [{ name: t.header[valueCol], values: shown.map((x) => x.value), fmt: t.formatOf(valueCol), color: one }],
    // What's drawn: each group in its colour, with its tiles (without Details a group is one tile)
    groups: shown.reduce((out, x) => {
      const last = out[out.length - 1];
      const tile = { name: label(detailCol !== undefined ? x.detail : x.group), value: x.value };
      if (last && last.key === x.group && detailCol !== undefined) { last.value += x.value; last.tiles.push(tile); }
      else out.push({ key: x.group, name: label(x.group), value: x.value, color: (byGroup && byGroup[x.group]) || one || null, tiles: [tile] });
      return out;
    }, []),
    nested: detailCol !== undefined,
  };
}

const NO_BOX = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>";
const labelFlags = ({ val = false, cat = false, ser = false, pct = false } = {}) =>
  `<c:showLegendKey val="0"/><c:showVal val="${val ? 1 : 0}"/><c:showCatName val="${cat ? 1 : 0}"/><c:showSerName val="${ser ? 1 : 0}"/>` +
  `<c:showPercent val="${pct ? 1 : 0}"/><c:showBubbleSize val="0"/>`;

/* Value labels in the middle of each bar (used by the funnel) */
function dataLabelsXml(fmt, argb, t) {
  // Labels inside a bar: white or ink, whichever reads on the bar's colour (Power BI's white)
  const text = { ...LOOK_TEXT, ...(t || {}), ...(argb ? { color: textOn(argbOf(argb)) } : {}) };
  return `<c:dLbls><c:numFmt formatCode="${xmlEsc(fmt || "General")}" sourceLinked="0"/>${NO_BOX}${txPrA(text)}` +
    `<c:dLblPos val="ctr"/>${labelFlags({ val: true })}</c:dLbls>`;
}

/*
 * A series' data labels as the report sets them: position (only where Excel allows it for the
 * chart type), display units and decimals, font, background. Inside a bar they take white or
 * ink, whichever reads on the bar, unless the report sets their colour.
 */
const LABEL_POSITIONS = { insideend: "inEnd", outsideend: "outEnd", insidecenter: "ctr", insidebase: "inBase", above: "t", under: "b", below: "b", left: "l", right: "r", center: "ctr" };
function labelPosOf(kind, grouping, labels) {
  let pos = LABEL_POSITIONS[String(labels.position || "").toLowerCase()] || null;
  if (kind === "bar") {
    const allowed = grouping === "clustered" ? ["ctr", "inEnd", "inBase", "outEnd"] : ["ctr", "inEnd", "inBase"];
    if (!allowed.includes(pos)) pos = grouping === "clustered" ? "outEnd" : "ctr";
  } else if (kind === "line") {
    if (!["t", "b", "l", "r", "ctr"].includes(pos)) pos = "t";
  } else {
    pos = null;   // area: Excel places them
  }
  return pos;
}
/* hide: the points whose label isn't shown (a bar too small to hold its label inside) */
function seriesLabelsXml(kind, grouping, color, numFmt, labels, hide = []) {
  const pos = labelPosOf(kind, grouping, labels);
  const inside = kind === "bar" && pos !== "outEnd";
  const own = labels.text.color && hexOf(labels.text.color) !== hexOf(LOOK_TEXT.color);
  const text = inside && !own && color ? { ...labels.text, color: textOn(argbOf(color)) } : labels.text;
  const back = labels.background ? fillA(labels.background.color, 1 - labels.background.transparency / 100) : "<a:noFill/>";
  const off = hide.map((i) => `<c:dLbl><c:idx val="${i}"/><c:delete val="1"/></c:dLbl>`).join("");
  return `<c:dLbls>${off}<c:numFmt formatCode="${xmlEsc(numFmt || "General")}" sourceLinked="0"/><c:spPr>${back}<a:ln><a:noFill/></a:ln></c:spPr>${txPrA(text)}` +
    (pos ? `<c:dLblPos val="${pos}"/>` : "") + `${labelFlags({ val: true })}</c:dLbls>`;
}

/* Labels written as text, one per point (pie slices, totals, analytics lines): [{ idx, text, pos }] */
function textLabelsXml(points, t) {
  return `<c:dLbls>${points.map((p) => `<c:dLbl><c:idx val="${p.idx}"/>${richA(p.text, t)}${NO_BOX}` +
    (p.pos ? `<c:dLblPos val="${p.pos}"/>` : "") + `${labelFlags({ val: true })}</c:dLbl>`).join("")}` +
    `${NO_BOX}${labelFlags()}${points.some((p) => p.leader) ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
}

/* A linear trend line in the report's colour and dash */
function trendlineXml(trend) {
  return `<c:trendline><c:spPr>${lineA({ color: trend.color, width: 1.5, style: trend.style, alpha: 1 - trend.transparency / 100 })}</c:spPr>` +
    `<c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>`;
}

/* Scatter / bubble chart XML: each legend group reads its own block of rows on the Data sheet */
function scatterXml(plan) {
  const look = plan.look || chartLook({ type: "scatterChart" }, null);
  const bubble = plan.spec.kind === "bubble";
  const [xs, ys, ss] = plan.series;
  const range = plan.refs.range;
  // Each point's name beside it only when the report turns category labels on (it still shows on hover)
  const labels = plan.perPoint && look.categoryLabels ? nameLabelsXml(look.labels.text) : "";
  const trend = look.trend && !bubble && !plan.perPoint ? trendlineXml(look.trend) : "";
  const numRef = (i, g, s) => `<c:numRef><c:f>${xmlEsc(range(i, g.start, g.end))}</c:f>` +
    `${numCacheXml(s.values.slice(g.start, g.end + 1), s.fmt)}</c:numRef>`;
  const fill = (argb) => (argb ? `<a:solidFill><a:srgbClr val="${String(argb).slice(-6).toUpperCase()}"><a:alpha val="85000"/></a:srgbClr></a:solidFill>` : "");
  const sers = plan.groups.map((g, i) => {
    const head = `<c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xmlEsc(g.name)}</c:v></c:tx>`;
    const xv = `<c:xVal>${numRef(0, g, xs)}</c:xVal>`, yv = `<c:yVal>${numRef(1, g, ys)}</c:yVal>`;
    const sp = `<c:spPr>${fill(g.color)}<a:ln w="9525">${fillXml(C.WHITE)}</a:ln></c:spPr>`;
    if (bubble) {
      return `<c:ser>${head}${sp}<c:invertIfNegative val="0"/>${labels}${xv}${yv}<c:bubbleSize>${numRef(2, g, ss)}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
    }
    return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>` +
      `<c:marker><c:symbol val="circle"/><c:size val="8"/>${sp}</c:marker>${labels}${trend}${xv}${yv}<c:smooth val="0"/></c:ser>`;
  }).join("");
  // Axis range and steps like Power BI's automatic scatter axes: not from zero, a step about
  // every 200 px across and 75 px up, so the gridlines fall on the same values as in the
  // report; constant lines are kept in view
  const size = plan.sizePx || { w: 0, h: 0 };
  const constants = (k) => look.refLines.filter((l) => l.kind === "constant" && l.axis === k).map((l) => l.value);
  // Bubbles need room beyond the outermost centres, so their edges aren't cut off
  const roomy = (vals) => { const v = vals.filter((x) => typeof x === "number" && isFinite(x)); if (!bubble || v.length < 2) return v; const d = (Math.max(...v) - Math.min(...v)) * 0.08; return [...v, Math.min(...v) - d, Math.max(...v) + d]; };
  const xAxis = valueScale([...roomy(xs.values), ...constants("x")], size.w - 90, 200, { round: false, start: look.cat.start, end: look.cat.end });
  const yAxis = valueScale([...roomy(ys.values), ...constants("y")], size.h - 100, 75, { round: false, start: look.val.start, end: look.val.end });
  // Constant and analytics lines: two-point lines across the plot, on the same axes
  const lines = [];
  for (const l of look.refLines) {
    const vertical = l.axis === "x";
    const v = l.kind === "constant" ? l.value : analyticValue(l, vertical ? xs.values : ys.values);
    if (v === null || !xAxis || !yAxis) continue;
    const pts = vertical ? [[v, yAxis.min], [v, yAxis.max]] : [[xAxis.min, v], [xAxis.max, v]];
    lines.push(refLineSeriesXml(plan.groups.length + lines.length, l, pts, vertical ? xs.fmt : ys.fmt, vertical, v));
  }
  plan.hiddenLegend = lines.map((_, k) => plan.groups.length + k);
  const A1 = 50110, A2 = 50120;
  const chart = bubble
    ? `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="40"/><c:showNegBubbles val="0"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:bubbleChart>` +
      (lines.length ? `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${lines.join("")}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:scatterChart>` : "")
    : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${lines.join("")}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:scatterChart>`;
  // Axes along the bottom and left edges with their field names, as Power BI draws them
  const a = plan.axes || {};
  const on = (x, key) => !x || x[key] !== false;
  const xTitle = on(a.cat, "title") ? look.cat.titleText || xs.name : null, yTitle = on(a.val, "title") ? look.val.titleText || ys.name : null;
  return chart +
    valAxXml(A1, A2, "b", axisNumFmt(xs.fmt, xAxis, look.cat, xs.values), { crosses: "min", title: xTitle, axis: xAxis, hidden: !on(a.cat, "show"), look: { ...look.cat, grid: look.cat.grid || look.val.grid } }) +
    valAxXml(A2, A1, "l", axisNumFmt(ys.fmt, yAxis, look.val, ys.values), { crosses: "min", title: yTitle, axis: yAxis, hidden: !on(a.val, "show"), look: look.val });
}

/* An analytics line's value from the points it summarises: min, max, average, median, percentile */
function analyticValue(line, values) {
  const v = values.filter((x) => typeof x === "number" && isFinite(x)).sort((a, b) => a - b);
  if (!v.length) return null;
  switch (line.kind) {
    case "min": return v[0];
    case "max": return v[v.length - 1];
    case "average": return v.reduce((a, b) => a + b, 0) / v.length;
    case "median": return v.length % 2 ? v[(v.length - 1) / 2] : (v[v.length / 2 - 1] + v[v.length / 2]) / 2;
    case "percentile": {
      const p = Math.max(0, Math.min(100, line.percentile ?? 50)) / 100 * (v.length - 1);
      const lo = Math.floor(p);
      return v[lo] + (v[Math.min(v.length - 1, lo + 1)] - v[lo]) * (p - lo);
    }
    default: return null;
  }
}

/*
 * A line Power BI draws "stepped": each value flat across its own category (from the first
 * category's middle to the last's), a straight drop at each boundary to the next, as scatter
 * pieces across the plot (0 to 1 over the categories). Excel draws no more points of one such
 * series than the chart has categories, so each flat and each drop is a piece of its own; the
 * first carries the line's name for the legend, the rest are left out of it. Labels over each
 * category's middle. alloc(): a new series number, kept out of the legend.
 */
function steppedLineSeriesXml([s, i], n, plan, refs, ls, label, alloc) {
  const lit = (vals) => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, k) => `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numLit>`;
  const ref = refs.series[i];
  const sp = `<c:spPr>${lineA({ color: s.color, width: ls.width, style: ls.style })}</c:spPr><c:marker><c:symbol val="none"/></c:marker>`;
  let first = true;
  const piece = (pts, labelAt) => {
    const idx = first ? i : alloc();
    const name = first && ref ? `<c:tx><c:strRef><c:f>${xmlEsc(ref.nameRef)}</c:f>${strCacheXml([s.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${xmlEsc(s.name)}</c:v></c:tx>`;
    first = false;
    const hide = pts.map((_, k) => k).filter((k) => k !== labelAt);
    const labels = label && labelAt >= 0 ? seriesLabelsXml("line", "standard", s.color, label.fmt, label.labels, hide) : "";
    return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${name}${sp}${labels}` +
      `<c:xVal>${lit(pts.map((p) => p[0]))}</c:xVal><c:yVal>${lit(pts.map((p) => p[1]))}</c:yVal><c:smooth val="0"/></c:ser>`;
  };
  const shown = s.values.map((v, k) => [k, v]).filter(([, v]) => typeof v === "number");
  const out = [];
  shown.forEach(([k, v], j) => {
    const left = j === 0 ? null : [k / n, v], mid = [(k + 0.5) / n, v], right = j === shown.length - 1 ? null : [(k + 1) / n, v];
    const pts = [left, mid, right].filter(Boolean);
    out.push(piece(pts, pts.indexOf(mid)));
    // the drop at the boundary to the next value
    if (right) out.push(piece([[(k + 1) / n, v], [(k + 1) / n, shown[j + 1][1]]], -1));
  });
  return out.join("");
}

/* One constant or analytics line: a two-point scatter series ([x, y] pairs), with its label
   (the value, its name, or both) at the line's start, as Power BI places it */
function refLineSeriesXml(idx, line, pts, fmt, vertical, value) {
  const lit = (vals) => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numLit>`;
  let labels = "";
  if (line.label) {
    const unit = line.label.units > 0 ? line.label.units : 1;
    const valueText = displayValue(value, fmt, unit, line.label.precision);
    const name = line.label.name || (line.kind === "constant" ? "" : line.kind[0].toUpperCase() + line.kind.slice(1));
    const text = /^name$/i.test(line.label.text) ? name : /name/i.test(line.label.text) ? `${name} ${valueText}`.trim() : valueText;
    labels = textLabelsXml([{ idx: vertical ? 1 : 0, text, pos: vertical ? "r" : "t" }], { ...LOOK_TEXT, color: line.label.color });
  }
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${xmlEsc(line.label && line.label.name ? line.label.name : line.kind)}</c:v></c:tx>` +
    `<c:spPr>${lineA({ color: line.color, width: line.width, style: line.style, alpha: 1 - line.transparency / 100 })}</c:spPr>` +
    `<c:marker><c:symbol val="none"/></c:marker>${labels}<c:xVal>${lit(pts.map((p) => p[0]))}</c:xVal><c:yVal>${lit(pts.map((p) => p[1]))}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* Number format for axis labels, shortened the way Power BI does: "$4,000,000" → "$4.0M", "34.28%" → "34%" */
function axisFormat(fmt, values) {
  const f = String(fmt || "General");
  const nums = (values || []).filter((v) => typeof v === "number" && isFinite(v));
  const max = nums.length ? Math.max(...nums.map((v) => Math.abs(v))) : 0;
  if (f.includes("%")) return max >= 0.05 ? "0%" : "0.0%";
  const prefix = (f.match(/^"[^"]*"/) || [""])[0];
  if (max >= 1e8) return `${prefix}#,##0.0,,,"B"`;
  if (max >= 1e5) return `${prefix}#,##0.0,,"M"`;   // axes switch a step early, as Power BI's: $0.4M
  if (max >= 1e4) return `${prefix}#,##0.0,"K"`;
  return f;
}

/* An axis title in the report's font (Power BI's DIN, about 10pt, by default) */
const AXIS_TITLE_TEXT = { face: "Bahnschrift", size: 10, color: "#252423", bold: false, italic: false };
function axisTitleXml(text, font) {
  return `<c:title>${richA(text, font || AXIS_TITLE_TEXT)}<c:overlay val="0"/></c:title>`;
}

/* Bubble labels: each bubble's name beside it ("050-Shoes") */
function nameLabelsXml(t) {
  return `<c:dLbls>${NO_BOX}${txPrA({ ...LOOK_TEXT, ...(t || {}) })}<c:dLblPos val="r"/>${labelFlags({ ser: true })}</c:dLbls>`;
}

/*
 * Pie and donut labels as Power BI writes them, per its label style ("Data value, percent of
 * total" → "$6K (26.95%)", "Category, percent of total" → "Shoes 39.08%"), outside the slices
 * with leader lines. Written as text, so they read exactly as in the report.
 */
function pieLabelPoints(plan, look) {
  const s = plan.series[0];
  const values = s.values.map((v) => (typeof v === "number" && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const L = look.labels;
  const style = String(L.style || "Data value, percent of total").toLowerCase();
  const unit = L.units > 0 ? L.units : labelAutoUnit(maxAbsOf(values));
  const dec = L.precision !== null && L.precision !== undefined ? L.precision : null;
  const pctDec = L.percentPrecision !== null && L.percentPrecision !== undefined ? L.percentPrecision : 2;
  const pctText = (v) => (v / total * 100).toLocaleString("en-US", { minimumFractionDigits: L.percentPrecision !== null && L.percentPrecision !== undefined ? pctDec : 0, maximumFractionDigits: pctDec }) + "%";
  return values.map((v, idx) => {
    if (!v) return null;
    const cat = String(plan.catText ? plan.catText[idx] : axisLabel(plan.categories[idx]));
    const val = displayValue(v, s.fmt, unit, dec);
    const pct = pctText(v);
    const text = style === "category" ? cat : style === "data value" ? val : style === "percent of total" ? pct
      : style === "category, data value" ? `${cat} ${val}` : style === "category, percent of total" ? `${cat} ${pct}`
        : style === "all detail labels" ? `${cat} ${val} (${pct})` : `${val} (${pct})`;
    // A long label takes two lines, the share under the value, as Power BI fits them beside the slice
    return { idx, text: textPx(text, L.text) > 70 ? text.replace(/ (?=\()/, "\n") : text, pos: "outEnd", leader: true };
  }).filter(Boolean);
}

/*
 * A pie's labels that fit: each sits outside its slice's middle, on the slice's side (right or
 * left), at that height; one that would overlap a label already placed on its side is left out.
 * Largest slices are placed first, so the small ones give way, as in Power BI. Every label fits
 * when there's room, so pies with few slices are unchanged.
 */
function fitPieLabels(points, values, startAngle, radius, lineH) {
  if (points.length < 2) return points;
  const v = values.map((x) => (typeof x === "number" && x > 0 ? x : 0));
  const total = v.reduce((a, b) => a + b, 0);
  if (!total) return points;
  const mid = [];
  let at = startAngle;
  v.forEach((x, i) => { const sweep = (x / total) * 360; mid[i] = at + sweep / 2; at += sweep; });
  const placed = { left: [], right: [] };
  const keep = new Set();
  for (const p of [...points].sort((a, b) => v[b.idx] - v[a.idx])) {
    const rad = (mid[p.idx] * Math.PI) / 180;
    const side = Math.sin(rad) >= 0 ? "right" : "left";
    const y = -Math.cos(rad) * (radius + lineH);
    const h = String(p.text).split("\n").length * lineH;
    if (placed[side].some((q) => Math.abs(q.y - y) < (q.h + h) / 2 + 2)) continue;
    placed[side].push({ y, h });
    keep.add(p.idx);
  }
  return points.filter((p) => keep.has(p.idx));
}

/* Which columns are X, Y and Size: the report file's field wells when it could be read, else
   Power BI's automatic title, which lists them in order ("X, Y and Size by Details"), else column order */
function scatterFields(t, item, info, nums) {
  const pick = (role) => {
    const f = info && info.roles && (info.roles[role] || [])[0];
    if (!f) return -1;
    return nums.find((ci) => String(t.header[ci]).trim().toLowerCase() === f.label || normalizeField(t.header[ci]) === f.ref) ?? -1;
  };
  const x = pick("X"), y = pick("Y");
  if (x >= 0 && y >= 0 && x !== y) {
    const s = pick("Size");
    return [x, y, s >= 0 && s !== x && s !== y ? s : undefined];
  }
  const title = String(item.title || "").toLowerCase();
  const cut = title.lastIndexOf(" by ");
  const head = cut > 0 ? title.slice(0, cut) : title;
  const hits = nums.map((ci) => {
    const name = String(t.header[ci] || "").trim().toLowerCase();
    return { ci, name, at: name ? head.indexOf(name) : -1 };
  }).filter((h) => h.at >= 0);
  const found = hits.filter((h) => !hits.some((o) => o !== h && o.name.length > h.name.length && o.at <= h.at && h.at + h.name.length <= o.at + o.name.length))
    .sort((a, b) => a.at - b.at);
  if (found.length >= 2) return [found[0].ci, found[1].ci, found[2] ? found[2].ci : undefined];
  return [nums[0], nums[1], nums[2]];
}

/* Axis label format with as many decimals as the step needs: 10% steps → "-30%", 0.5 steps → "$12.5" */
function stepFormat(fmt, axis, values) {
  const f = String(fmt || "General");
  if (!axis) return axisFormat(fmt, values);
  const pct = f.includes("%");
  if (!pct && Math.max(Math.abs(axis.min), Math.abs(axis.max)) >= 1e4) return axisFormat(fmt, values);
  const s = pct ? axis.step * 100 : axis.step;
  const d = Math.max(0, Math.min(4, -Math.floor(Math.log10(s) + 1e-9)));
  const dec = d ? "." + "0".repeat(d) : "";
  if (pct) return `0${dec}%`;
  const prefix = (f.match(/^"[^"]*"/) || [""])[0];
  return `${prefix}#,##0${dec}`;
}

/* A waterfall step as text, the way Power BI labels it: "-10%", "32%" */
function waterfallLabel(v, fmt, precision) {
  const f = String(fmt || "");
  if (f.includes("%")) return (v * 100).toFixed(precision ?? 0) + "%";
  const prefix = (f.match(/^"([^"]*)"/) || [])[1] || "";
  // Without decimals set in the report, large steps read like Power BI's auto units (2.47M)
  const a = Math.abs(v);
  if (precision === undefined && a >= 1e4) {
    const [div, unit] = a >= 1e9 ? [1e9, "bn"] : a >= 1e6 ? [1e6, "M"] : [1e3, "K"];
    return (v < 0 ? "-" : "") + prefix + (a / div).toLocaleString("en-US", { maximumFractionDigits: 2 }) + unit;
  }
  const dec = precision ?? (((f.match(/\.(0+)/) || [])[1] || "").length);
  return (v < 0 ? "-" : "") + prefix + Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

/*
 * Waterfall chart: each step is ONE bar from its start to its end value (Excel "up/down bars"
 * between an invisible Start line and an invisible End line), so a step that crosses zero is
 * never split in two. Up bars take the increase colour, down bars the decrease colour; the
 * total is a normal column. Labels show each step's value, like Power BI.
 */
function waterfallXml(plan) {
  const { series, refs, categories, axis } = plan;
  const look = plan.look || chartLook({ type: "waterfallChart" }, null);
  const A1 = 50210, A2 = 50220, A3 = 50230, A4 = 50240;
  const last = categories.length - 1;
  const noBox = NO_BOX;
  const flags = (show) => labelFlags({ val: show });
  // Labels like Power BI's: in the label font beside a bar, white bold inside one
  const label = (i, pos, argb, bold) => `<c:dLbl><c:idx val="${i}"/>` +
    richA(waterfallLabel(plan.steps[i], plan.fmt, plan.labelPrecision ?? look.labels.precision ?? undefined), { ...look.labels.text, color: argb === C.WHITE ? C.WHITE : look.labels.text.color, bold: bold || look.labels.text.bold }) +
    `${noBox}<c:dLblPos val="${pos}"/>${flags(true)}</c:dLbl>`;
  const cat = `<c:cat><c:strRef><c:f>${xmlEsc(refs.catRef)}</c:f>${strCacheXml(categories.map(axisLabel))}</c:strRef></c:cat>`;
  const head = (i) => `<c:idx val="${i}"/><c:order val="${i}"/>` +
    `<c:tx><c:strRef><c:f>${xmlEsc(refs.series[i].nameRef)}</c:f>${strCacheXml([series[i].name])}</c:strRef></c:tx>`;
  const val = (i) => `<c:val><c:numRef><c:f>${xmlEsc(refs.series[i].valRef)}</c:f>${numCacheXml(series[i].values, series[i].fmt)}</c:numRef></c:val>`;
  const bars = [0, 1, 2, 3].map((i) => {
    const sp = series[i].hidden ? noBox : `<c:spPr>${fillXml(series[i].color)}<a:ln><a:noFill/></a:ln></c:spPr>`;
    return `<c:ser>${head(i)}${sp}<c:invertIfNegative val="0"/>${cat}${val(i)}</c:ser>`;
  }).join("");
  // Labels sit on an invisible line through each bar's end value, on the visible axis: beside
  // the bar (above an increase, below a decrease), or inside its end in white where the axis
  // edge leaves no room, so a label never runs into the axis labels
  const labels = plan.showLabels
    ? `<c:dLbls>${plan.steps.map((v, i) => {
        const inside = plan.labelInside ? plan.labelInside[i] : i === last;
        return label(i, (v >= 0) !== inside ? "t" : "b", inside ? C.WHITE : C.SLATE, inside);
      }).join("")}${noBox}${flags(false)}</c:dLbls>` : "";
  const line = `<c:ser>${head(4)}<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>${labels}${cat}${val(4)}<c:smooth val="0"/></c:ser>`;
  const titles = axisTitles(plan, [], []);
  // Columns first, so legend entries follow the series order (helper entries are removed by index).
  // The columns use a second axis without labels, running from 0 to the axis height (the shift);
  // the visible axis shows the real values, and both cover the same range.
  return `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>${bars}` +
      `<c:gapWidth val="60"/><c:overlap val="100"/><c:axId val="${A3}"/><c:axId val="${A4}"/></c:barChart>` +
    `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${line}<c:marker val="1"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:lineChart>` +
    catAxXml(A1, A2, "b", { title: titles.cat, hidden: !titles.catShow, count: categories.length, look: look.cat }) +
    valAxXml(A2, A1, "l", axisNumFmt(plan.fmt, axis, look.val), { axis, title: titles.val, hidden: !titles.valShow, look: look.val }) +
    catAxXml(A3, A4, "b", { hidden: true }) +
    valAxXml(A4, A3, "r", "General", { crosses: "max", grid: false, labels: false,
      axis: { min: 0, max: Number((axis.max - axis.min).toFixed(12)), step: axis.step } });
}

/* Writes every chart's numbers on the page's "Data" sheet and records the cell ranges */
function writeChartDataSheet(wb, name, charts) {
  const ws = wb.addWorksheet(name, { views: [{ showGridLines: false }] });
    // The chart numbers stay in the file (Excel charts need them) but the sheet is invisible:
  // "veryHidden" also keeps it out of Excel's Unhide list, so users see only the visual sheet
  ws.state = "veryHidden";
  const sheetRef = `'${name.replace(/'/g, "''")}'`;
  const abs = (r, c) => `$${ws.getColumn(c).letter}$${r}`;
  let row = 2, widest = 1;
  for (const p of charts.flatMap((ch) => (ch.plan.panels ? ch.plan.panels.plans : [ch.plan]))) {
    const t = ws.getCell(row, 2);
    t.value = p.title;
    t.font = { name: "Calibri", size: 12, bold: true, color: { argb: C.INK } };
    const hr = row + 1;
    // Higher category levels (a treemap's groups) come first: Excel reads the levels together
    const levels = p.catLevels || [], off = levels.length;
    [...levels.map((l) => l.name), p.catName || "Category"].forEach((name, k) => {
      const h = ws.getCell(hr, 2 + k);
      h.value = name;
      h.font = { name: "Calibri", size: 10, bold: true, color: { argb: C.INK } };
      h.fill = SURFACE_FILL;
    });
    p.series.forEach((s, i) => {
      const h = ws.getCell(hr, 3 + off + i);
      h.value = s.name;
      h.font = { name: "Calibri", size: 10, bold: true, color: { argb: C.INK } };
      h.fill = SURFACE_FILL;
      h.alignment = { horizontal: "right" };
    });
    p.categories.forEach((c, ri) => {
      // A multi-level axis groups by blank cells: a higher level's value only where its group starts
      levels.forEach((l, k) => { if (!p.multiLevel || l.starts[ri]) ws.getCell(hr + 1 + ri, 2 + k).value = l.values[ri]; });
      const cat = ws.getCell(hr + 1 + ri, 2 + off);
      if (p.dateAxis) {
        cat.value = p.dateAxis.serials[ri];
        cat.numFmt = p.dateAxis.cellFmt || "m/d/yyyy";
      } else {
        cat.value = p.catText ? p.catText[ri] : axisLabel(c);   // a date reads "1/15/2023" on the axis, an empty one "(Blank)"
      }
      p.series.forEach((s, i) => {
        const cell = ws.getCell(hr + 1 + ri, 3 + off + i);
        cell.value = s.values[ri];
        if (s.fmt && s.fmt !== "General") cell.numFmt = s.fmt;
      });
    });
    const first = hr + 1, last = hr + p.categories.length;
    p.refs = {
      catRef: `${sheetRef}!${abs(first, 2)}:${abs(last, 2 + off)}`,
      series: p.series.map((s, i) => ({
        nameRef: `${sheetRef}!${abs(hr, 3 + off + i)}`,
        valRef: `${sheetRef}!${abs(first, 3 + off + i)}:${abs(last, 3 + off + i)}`,
      })),
    };
    // Any block of rows in one Data column (scatter charts read one block per legend value)
    p.refs.range = (i, from, to) => `${sheetRef}!${abs(first + from, 3 + off + i)}:${abs(first + to, 3 + off + i)}`;
    widest = Math.max(widest, p.series.length + off);
    row = last + 3;
  }
  ws.getColumn(1).width = 3;
  ws.getColumn(2).width = 30;
  for (let c = 3; c < 3 + widest; c++) ws.getColumn(c).width = 18;
  return ws;
}

/* ── Chart XML (DrawingML) ── */
function xmlEsc(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function fillXml(argb) {
  return argb ? `<a:solidFill><a:srgbClr val="${String(argb).slice(-6).toUpperCase()}"/></a:solidFill>` : "";
}
function strCacheXml(values) {
  return `<c:strCache><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(v)}</c:v></c:pt>`).join("") + `</c:strCache>`;
}
function numCacheXml(values, fmt) {
  return `<c:numCache><c:formatCode>${xmlEsc(fmt || "General")}</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => (v === null || v === undefined ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("") + `</c:numCache>`;
}

/* ── The visual's look in chart XML: text, fills, lines (sizes in points, widths in pixels) ── */
const LOOK_TEXT = { face: "Segoe UI", size: 9, color: "#605E5C", bold: false, italic: false };
const r4 = (x) => Number(Math.max(0, Math.min(1, x)).toFixed(4));
function hexOf(c) {
  return String(c || "000000").replace(/^#/, "").slice(-6).toUpperCase();
}
const argbOf = (c) => "FF" + hexOf(c);
/* A fill, optionally see-through (alpha 0–1); no colour: no fill */
function fillA(c, alpha = 1) {
  if (!c) return "<a:noFill/>";
  return `<a:solidFill><a:srgbClr val="${hexOf(c)}">${alpha < 1 ? `<a:alpha val="${Math.round(Math.max(0, alpha) * 100000)}"/>` : ""}</a:srgbClr></a:solidFill>`;
}
const DASHES = { dashed: "dash", dotted: "sysDot", dash: "dash", dot: "sysDot" };
function lineA({ color = null, width = 1, style = "solid", alpha = 1 } = {}) {
  if (!color) return "<a:ln><a:noFill/></a:ln>";
  const dash = DASHES[style] ? `<a:prstDash val="${DASHES[style]}"/>` : "";
  return `<a:ln w="${Math.round(width * CHART_PX_TO_EMU)}" cap="${dash ? "flat" : "rnd"}">${fillA(color, alpha)}${dash}<a:round/></a:ln>`;
}
function runPrA(t, tag = "a:defRPr") {
  const x = { ...LOOK_TEXT, ...(t || {}) }, face = xmlEsc(x.face);
  return `<${tag}${tag === "a:rPr" ? ' lang="en-US"' : ""} sz="${Math.round(x.size * 100)}" b="${x.bold ? 1 : 0}" i="${x.italic ? 1 : 0}">` +
    `${fillA(x.color)}<a:latin typeface="${face}"/><a:cs typeface="${face}"/></${tag}>`;
}
function txPrA(t, rot, noWrap = false) {
  return `<c:txPr><a:bodyPr${rot !== undefined ? ` rot="${rot}" vert="horz"` : ""}${noWrap ? ' wrap="none"' : ""}/><a:lstStyle/><a:p><a:pPr>${runPrA(t)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}
function paraA(text, t, align) {
  return `<a:p><a:pPr${align ? ` algn="${align}"` : ""}>${runPrA(t)}</a:pPr><a:r>${runPrA(t, "a:rPr")}<a:t>${xmlEsc(text)}</a:t></a:r></a:p>`;
}
/* Rich text; "\n" starts a new line */
function richA(text, t) {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/>${String(text).split("\n").map((line) => paraA(line, t)).join("")}</c:rich></c:tx>`;
}
/* A position inside the chart, in fractions of its width and height */
function manualLayoutA(x, y, w, h, inner) {
  return `<c:layout><c:manualLayout>${inner ? '<c:layoutTarget val="inner"/>' : ""}<c:xMode val="edge"/><c:yMode val="edge"/>` +
    `<c:x val="${r4(x)}"/><c:y val="${r4(y)}"/>${w !== undefined ? `<c:w val="${r4(w)}"/><c:h val="${r4(h)}"/>` : ""}</c:manualLayout></c:layout>`;
}
/* About how wide a line of text is, in pixels (Bahnschrift, Power BI's DIN, runs narrower than Segoe UI) */
function textPx(text, t) {
  const x = { ...LOOK_TEXT, ...(t || {}) };
  const per = /bahnschrift|din/i.test(x.face) ? 0.48 : 0.45;
  return String(text ?? "").length * x.size * (4 / 3) * per * (x.bold ? 1.06 : 1);
}
const lineHeightPx = (t) => ({ ...LOOK_TEXT, ...(t || {}) }).size * (4 / 3) * 1.3;
/* Text broken into lines no wider than maxPx, word by word, as Power BI wraps titles */
function wrapLines(text, t, maxPx) {
  const lines = [];
  let line = "";
  for (const word of String(text ?? "").split(/\s+/).filter(Boolean)) {
    const next = line ? `${line} ${word}` : word;
    if (line && textPx(next, t) > maxPx) { lines.push(line); line = word; } else line = next;
  }
  if (line) lines.push(line);
  return lines.length ? lines : [""];
}

/* ── Number formats as Power BI shows them: display units (K, M, bn, T) and decimals ── */
const PBI_UNITS = { 1000: ["K", ","], 1000000: ["M", ",,"], 1000000000: ["bn", ",,,"], 1000000000000: ["T", ",,,,"] };
function fmtPrefix(f) {
  const m = String(f || "").match(/^(?:"[^"]*"|\\.|[$€£¥₹])/);
  return m ? m[0] : "";
}
function fmtDecimals(f) {
  const m = String(f || "").split(";")[0].match(/\.(0+)/);
  return m ? m[1].length : 0;
}
function unitFormat(f, unit, dec) {
  const s = String(f || "General");
  if (s.includes("%")) return `0${dec ? "." + "0".repeat(dec) : ""}%`;
  const d = dec ?? fmtDecimals(s);
  const body = `${fmtPrefix(s)}#,##0${d ? "." + "0".repeat(d) : ""}`;
  return PBI_UNITS[unit] ? `${body}${PBI_UNITS[unit][1]}"${PBI_UNITS[unit][0]}"` : body;
}
/* An axis's Auto unit as Power BI picks it: the next unit a step early, "0.1M, 0.2M … rather
   than 100K, 200K" (millions from 100,000, billions from 100 million, trillions from 100 billion) */
function axisAutoUnit(maxAbs) {
  return maxAbs >= 1e11 ? 1e12 : maxAbs >= 1e8 ? 1e9 : maxAbs >= 1e5 ? 1e6 : maxAbs >= 1e3 ? 1e3 : 1;
}
function autoUnit(maxAbs) {
  return maxAbs >= 1e12 ? 1e12 : maxAbs >= 1e9 ? 1e9 : maxAbs >= 1e6 ? 1e6 : maxAbs >= 1e3 ? 1e3 : 1;
}
/* Display units "Auto" for data labels: Power BI shows thousands in full below 10,000 (1,037 on a
   line reaching 2,128) and in K from there (6.8K on stacks reaching 13K) */
function labelAutoUnit(maxAbs) {
  return maxAbs >= 1e4 ? autoUnit(maxAbs) : 1;
}
function maxAbsOf(values) {
  let m = 0;
  for (const v of values) if (typeof v === "number" && isFinite(v)) m = Math.max(m, Math.abs(v));
  return m;
}
/* A data label's number format. Display units "Auto" pick one unit for the whole visual from
   its largest value; with them Power BI keeps the measure's decimals, at least minDecimals
   (bar, line and area labels show one: $7.1K, $0.4K) */
function labelNumFmt(fmt, values, cfg, minDecimals = 0, auto = labelAutoUnit) {
  const s = String(fmt || "General");
  const precision = cfg && cfg.precision !== null && cfg.precision !== undefined ? cfg.precision : null;
  if (s.includes("%")) return precision !== null ? unitFormat(s, 1, precision) : s;
  const unit = cfg && cfg.units > 0 ? cfg.units : auto(maxAbsOf(values));
  if (unit <= 1 && precision === null) return s;
  // A whole-number measure shortened (data labels): Power BI gives the whole chart one number of
  // decimals: one when any label is below ten of the unit ($15.2M, $9.0M, $0.5M on one chart;
  // 13.6K beside 4.0K), none when every label is ten or more (33K, 31K, 24K, 15K)
  if (unit > 1 && precision === null && minDecimals === 0 && fmtDecimals(s) === 0) {
    const shown = values.filter((v) => typeof v === "number" && v !== 0);
    return unitFormat(s, unit, shown.length && shown.every((v) => Math.abs(v) >= 9.95 * unit) ? 0 : 1);
  }
  return unitFormat(s, unit, precision !== null ? precision : unit > 1 ? Math.max(minDecimals, fmtDecimals(s)) : null);
}
/* A number as label text, for labels written as text (pie slices, analytics lines) */
function displayValue(v, fmt, unit = 1, dec = null) {
  const s = String(fmt || "General");
  const fixed = (x, d) => x.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
  if (s.includes("%")) return fixed(v * 100, dec ?? fmtDecimals(s)) + "%";
  const prefix = fmtPrefix(s).replace(/^"|"$/g, "").replace(/^\\/, "");
  const d = dec ?? fmtDecimals(s);   // Power BI keeps the measure's own decimals with auto units: 1.10M, $8K
  return `${v < 0 ? "-" : ""}${prefix}${fixed(Math.abs(v) / (unit || 1), d)}${PBI_UNITS[unit] ? PBI_UNITS[unit][0] : ""}`;
}

/* A tick step of 1, 2 or 5 × 10ⁿ giving about `count` steps across lo…hi (d3's rule, which Power BI's axes follow) */
function tickStep(lo, hi, count) {
  const raw = Math.abs(hi - lo) / Math.max(1, count);
  if (!raw || !isFinite(raw)) return 1;
  const power = Math.pow(10, Math.floor(Math.log10(raw)));
  const err = raw / power;
  return (err >= 7.0710678 ? 10 : err >= 3.1622777 ? 5 : err >= 1.4142136 ? 2 : 1) * power;
}

/*
 * A value axis like Power BI's automatic one: about one tick every gapPx pixels (Power BI
 * spaces them ~75 px up, ~200 px across), ending just past the data. round: run on to the
 * next tick when that adds no more than a quarter, then take the steps that range wants (0 to
 * $10K in $5K steps for an $8.4K top), as Power BI does to make room for labels over bars.
 * start / end: the report's own axis range.
 */
function valueScale(values, plotPx, gapPx, { zero = false, round = true, start = null, end = null } = {}) {
  const v = values.filter((x) => typeof x === "number" && isFinite(x));
  if (zero) v.push(0);
  if (!v.length && (start === null || end === null)) return null;
  const tidy = (x) => Number(x.toFixed(12));
  let lo = start !== null ? start : Math.min(...v), hi = end !== null ? end : Math.max(...v);
  if (lo === hi) { const d = Math.abs(lo) || 1; lo -= d / 2; hi += d / 2; if (zero && lo < 0 && hi > 0 && Math.min(...v) >= 0) lo = 0; }
  // Values all but equal (a margin of 43.934% everywhere): a little room around them
  else if (hi - lo < Math.abs(hi) * 1e-3 && start === null && end === null) { const d = Math.abs(hi) * 5e-3; lo -= d; hi += d; }
  const count = Math.max(1.5, Math.max(60, plotPx) / gapPx);
  let step = tickStep(lo, hi, count);
  const floorTo = (x, s) => tidy(Math.floor(x / s + 1e-9) * s);
  let min = start !== null ? start : zero && lo >= 0 ? 0 : floorTo(lo, step);
  let max;
  if (end !== null) max = end;
  else {
    const niceMax = tidy(Math.ceil(hi / step - 1e-9) * step);
    if (round && hi > 0 && (niceMax - hi) / hi <= 0.25) {
      step = tickStep(min, niceMax, count);
      max = tidy(Math.ceil(niceMax / step - 1e-9) * step);
      if (start === null && !(zero && lo >= 0)) min = floorTo(lo, step);
    } else {
      max = tidy(hi + (hi - min) * 0.03);
    }
  }
  if (max <= min) max = tidy(min + step);
  return { min, max, step };
}
/* A value axis's number format: display units from the step ($5K steps, $500 steps), as many decimals as the step needs */
function axisNumFmt(fmt, axis, cfg, values) {
  const s = String(fmt || "General");
  const precision = cfg && cfg.precision !== null && cfg.precision !== undefined ? cfg.precision : null;
  if (!axis) return axisFormat(fmt, values || []);
  if (s.includes("%")) return precision !== null ? unitFormat(s, 1, precision) : stepFormat(s, axis, values || []);
  const step = axis.step || Math.abs(axis.max - axis.min) || 1;
  // Power BI's auto units: from the axis's largest value once it reaches 10,000 ($20K, $0.4M);
  // below that, thousands only when the steps are whole thousands ($2K, but $1,500)
  const top = Math.max(Math.abs(axis.min), Math.abs(axis.max));
  const unit = cfg && cfg.units > 0 ? cfg.units : top >= 1e4 ? axisAutoUnit(top) : autoUnit(step);
  const dec = precision !== null ? precision : Math.max(0, Math.min(4, -Math.floor(Math.log10(step / unit) + 1e-9)));
  return unitFormat(s, unit, dec);
}

/* A series' categories: dates on a date axis, Excel's multi-level labels for a hierarchy, else text */
function categoryXml(plan) {
  const ref = xmlEsc(plan.refs.catRef);
  if (plan.dateAxis) return `<c:cat><c:numRef><c:f>${ref}</c:f>${numCacheXml(plan.dateAxis.serials, "m/d/yyyy")}</c:numRef></c:cat>`;
  if (plan.multiLevel) {
    const lvl = (vals, starts) => `<c:lvl>${vals.map((v, i) => (starts && !starts[i] ? "" : `<c:pt idx="${i}"><c:v>${xmlEsc(v)}</c:v></c:pt>`)).join("")}</c:lvl>`;
    return `<c:cat><c:multiLvlStrRef><c:f>${ref}</c:f><c:multiLvlStrCache><c:ptCount val="${plan.categories.length}"/>` +
      lvl(plan.catText) + [...plan.catLevels].reverse().map((l) => lvl(l.values, l.starts)).join("") + `</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
  }
  return `<c:cat><c:strRef><c:f>${ref}</c:f>${strCacheXml(plan.catText || plan.categories.map(axisLabel))}</c:strRef></c:cat>`;
}

const MARKER_SYMBOLS = { circle: "circle", square: "square", diamond: "diamond", triangle: "triangle", x: "x", plus: "plus", shortdash: "dash", longdash: "dash", star: "star" };
const DEFAULT_LINE_LOOK = { width: 3, style: "solid", markers: null, smooth: false };

/*
 * One series. s.literal: values written into the chart (overlay lines, total labels), not
 * read from the Data sheet. s.labelsXml / s.trendXml: its data labels and trend line;
 * s.lineLook: its line (width, dash, markers); s.alpha: a see-through fill (areas).
 */
function chartSeriesXml(kind, s, i, plan, ref) {
  const head = s.literal ? `<c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xmlEsc(s.name)}</c:v></c:tx>`
    : `<c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>${xmlEsc(ref.nameRef)}</c:f>${strCacheXml([s.name])}</c:strRef></c:tx>`;
  const cat = categoryXml(plan);
  const val = s.literal
    ? `<c:val><c:numLit><c:formatCode>${xmlEsc(s.fmt || "General")}</c:formatCode><c:ptCount val="${s.values.length}"/>` +
      s.values.map((v, k) => (v === null || v === undefined ? "" : `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`)).join("") + `</c:numLit></c:val>`
    : `<c:val><c:numRef><c:f>${xmlEsc(ref.valRef)}</c:f>${numCacheXml(s.values, s.fmt)}</c:numRef></c:val>`;
  // Hidden helper series (waterfall base, funnel offset) take no bar colours
  const points = s.hidden ? "" : (plan.pointColors || []).map((c, pi) => (c
    ? `<c:dPt><c:idx val="${pi}"/>${kind === "bar" ? '<c:invertIfNegative val="0"/>' : ""}<c:bubble3D val="0"/><c:spPr>${fillXml(c)}</c:spPr></c:dPt>`
    : "")).join("");
  const look = plan.look;
  const labels = s.labelsXml || (s.labels
    ? dataLabelsXml(labelNumFmt(s.fmt, s.values, look && look.labels), s.color || (plan.pointColors && plan.pointColors.find(Boolean)) || null, look && look.labels.text) : "");
  const trend = s.trendXml || "";
  if (kind === "line") {
    const ls = s.lineLook || (look && look.lines) || DEFAULT_LINE_LOOK;
    // the report's "Show line" off: the markers alone (when it has some)
    const lineOff = ls.show === false && ls.markers;
    const sp = s.hidden ? NO_BOX : lineOff ? `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>`
      : s.color ? `<c:spPr>${lineA({ color: s.color, width: ls.width, style: ls.style })}</c:spPr>` : "";
    const m = ls.markers && !s.hidden
      ? `<c:marker><c:symbol val="${MARKER_SYMBOLS[String(ls.markers.shape).toLowerCase()] || "circle"}"/><c:size val="${Math.max(2, Math.min(72, Math.round(ls.markers.size * 0.9)))}"/>` +
        `<c:spPr>${fillA(ls.markers.color || s.color)}${lineA({ color: ls.markers.color || s.color, width: 1 })}</c:spPr></c:marker>`
      : `<c:marker><c:symbol val="none"/></c:marker>`;
    return `<c:ser>${head}${sp}${m}${labels}${trend}${cat}${val}<c:smooth val="${ls.smooth && !s.hidden ? 1 : 0}"/></c:ser>`;
  }
  const sp = s.hidden ? NO_BOX
    : s.color ? `<c:spPr>${fillA(s.color, s.alpha ?? 1)}${s.alpha !== undefined ? "<a:ln><a:noFill/></a:ln>" : ""}</c:spPr>` : "";
  if (kind === "bar") return `<c:ser>${head}${sp}<c:invertIfNegative val="0"/>${points}${labels}${trend}${cat}${val}</c:ser>`;
  if (kind === "area") return `<c:ser>${head}${sp}${labels}${cat}${val}</c:ser>`;
  return `<c:ser>${head}${sp}${points}${labels}${cat}${val}</c:ser>`;   // pie
}

/* Gridlines in the report's colour, width and dash (Power BI's: dotted, light grey) */
function gridXml(grid) {
  return grid ? `<c:majorGridlines><c:spPr>${lineA({ color: grid.color, width: grid.width, style: grid.style })}</c:spPr></c:majorGridlines>` : "";
}

/* A category axis, without an axis line as in Power BI. Every label is drawn (Excel otherwise
   skips some when it thinks they are crowded, so labels came and went with the chart's size);
   up to 60 categories, beyond which Excel's own spacing reads better */
function catAxXml(id, crossId, pos, { reverse = false, hidden = false, title = null, count = 0, look = null, multi = false, noWrap = false } = {}) {
  const lk = look || {};
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${reverse ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${hidden && !title ? 1 : 0}"/><c:axPos val="${pos}"/>${hidden ? "" : gridXml(lk.grid)}` + (title ? axisTitleXml(title, lk.titleFont) : "") +
    `<c:numFmt formatCode="General" sourceLinked="1"/>` +
    `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="${hidden ? "none" : "low"}"/>` +
    `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPrA(lk.text, noWrap ? 0 : undefined, noWrap)}` +
    `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>` +
    (count && count <= 60 && !hidden && !multi ? `<c:tickLblSkip val="1"/>` : "") + `<c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/* A date axis: dates spaced by time, labelled every few months or years ("Jul 2023") */
function dateAxXml(id, crossId, { hidden = false, title = null, look = null, date }) {
  const lk = look || {};
  const range = date.min !== undefined && date.max !== undefined ? `<c:max val="${date.max}"/><c:min val="${date.min}"/>` : "";
  return `<c:dateAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${range}</c:scaling>` +
    `<c:delete val="${hidden && !title ? 1 : 0}"/><c:axPos val="b"/>${hidden ? "" : gridXml(lk.grid)}` + (title ? axisTitleXml(title, lk.titleFont) : "") +
    `<c:numFmt formatCode="${xmlEsc(date.fmt)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${hidden ? "none" : "low"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPrA(lk.text)}` +
    `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="0"/><c:lblOffset val="100"/><c:baseTimeUnit val="days"/>` +
    `<c:majorUnit val="${date.unit}"/><c:majorTimeUnit val="${date.timeUnit}"/></c:dateAx>`;
}

/* An axis the chart doesn't show, giving overlay lines their scale */
function hiddenValAxXml(id, crossId, pos, axis) {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${axis.max}"/><c:min val="${axis.min}"/></c:scaling>` +
    `<c:delete val="1"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:crossAx val="${crossId}"/><c:crosses val="min"/><c:crossBetween val="midCat"/></c:valAx>`;
}

/* A chart's axis titles as Power BI draws them by default: the category field, and the value
   fields ("Last Year Sales and This Year Sales"); a title or axis the report turns off isn't
   drawn. Returns { cat, val, sec, catShow, valShow, secShow } */
function axisTitles(plan, main, lines) {
    // A funnel has no axis titles in Power BI: only its stage names and values
  if (plan.funnel) return { catShow: true, valShow: true, secShow: true, cat: null, val: null, sec: null };
  const a = plan.axes || {};
  const lk = plan.look || {};
  const own = (k) => (lk[k] && lk[k].titleText ? { text: lk[k].titleText } : null);
  const on = (x, key) => !x || x[key] !== false;
  const names = (list) => {
    const n = [...new Set(list.map(([s]) => s.name))];
    return n.length <= 2 ? n.join(" and ") : `${n.slice(0, -1).join(", ")} and ${n[n.length - 1]}`;
  };
  return {
    catShow: on(a.cat, "show"), valShow: on(a.val, "show"), secShow: on(a.sec, "show"),
    // An axis turned off loses its labels, not its title (Power BI still draws the title)
        // The author's own title text ("Revenue") when the report sets one, else the field names
    cat: on(a.cat, "title") ? (own("cat") || a.cat || {}).text || plan.catTitle || plan.catName : null,
    val: on(a.val, "title") ? (own("val") || a.val || {}).text || plan.valueName || names(main) : null,
    sec: lines.length && on(a.sec, "title") ? (own("sec") || a.sec || {}).text || names(lines) : null,
  };
}

/*
 * A value axis in the report's look: labels' font, gridlines (none when the axis is hidden, as
 * in Power BI), log scale, reversed order, title font.
 */
function valAxXml(id, crossId, pos, fmt, { crosses = "autoZero", grid = true, hidden = false, title = null, axis = null, labels = true, zero = false, look = null } = {}) {
  const lk = look || {};
  const log = !!lk.log;
  // zero: bars and columns start at 0 as Power BI's do; Excel's own minimum can cut close values short
  const bounds = log ? (lk.end > 0 ? `<c:max val="${lk.end}"/>` : "") + (lk.start > 0 ? `<c:min val="${lk.start}"/>` : "")
    : axis ? `<c:max val="${axis.max}"/><c:min val="${axis.min}"/>` : zero ? `<c:min val="0"/>` : "";
  const lines = grid && !hidden ? (look ? gridXml(lk.grid) : `<c:majorGridlines><c:spPr><a:ln w="6350">${fillXml(C.HAIR)}</a:ln></c:spPr></c:majorGridlines>`) : "";
  return `<c:valAx><c:axId val="${id}"/><c:scaling>${log ? '<c:logBase val="10"/>' : ""}<c:orientation val="${lk.invert ? "maxMin" : "minMax"}"/>${bounds}</c:scaling>` +
    `<c:delete val="${hidden && !title ? 1 : 0}"/><c:axPos val="${pos}"/>${lines}` +
    (title ? axisTitleXml(title, lk.titleFont) : "") +
    `<c:numFmt formatCode="${xmlEsc(fmt || "General")}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="${labels && !hidden ? "nextTo" : "none"}"/>` +
    `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPrA(lk.text)}<c:crossAx val="${crossId}"/><c:crosses val="${crosses}"/><c:crossBetween val="between"/>` +
    (axis && !log ? `<c:majorUnit val="${axis.step}"/>` : "") + `</c:valAx>`;
}

/* Where a legend goes in Power BI → Excel's side, and whether it starts at the edge or is centred */
const LEGEND_SPOTS = { top: ["t", "start"], topcenter: ["t", "center"], topright: ["t", "end"], bottom: ["b", "start"], bottomcenter: ["b", "center"],
  bottomright: ["b", "end"], left: ["l", "start"], leftcenter: ["l", "center"], right: ["r", "start"], rightcenter: ["r", "center"] };
const DONUT_HOLE = 0.62;   // Power BI's donut hole, as a share of the donut's width, when the report leaves it
const USER_SHAPES_RID = "rId1";

/* Shapes drawn over a chart (a legend's title, a donut's hole), placed in fractions of the chart */
function shapeAnchorXml(x, y, w, h, inner) {
  return `<cdr:relSizeAnchor xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"><cdr:from><cdr:x>${r4(x)}</cdr:x><cdr:y>${r4(y)}</cdr:y></cdr:from>` +
    `<cdr:to><cdr:x>${r4(x + w)}</cdr:x><cdr:y>${r4(y + h)}</cdr:y></cdr:to>${inner}</cdr:relSizeAnchor>`;
}
function textShapeXml(id, text, t, x, y, w, h) {
  return shapeAnchorXml(x, y, w, h, `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="${id}" name="Legend title"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>` +
    `<cdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></cdr:spPr>` +
    `<cdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>` +
    `<a:p><a:r>${runPrA(t, "a:rPr")}<a:t>${xmlEsc(text)}</a:t></a:r></a:p></cdr:txBody></cdr:sp>`);
}
function ellipseShapeXml(id, color, x, y, w, h) {
  return shapeAnchorXml(x, y, w, h, `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="${id}" name="Donut hole"/><cdr:cNvSpPr/></cdr:nvSpPr>` +
    `<cdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${fillA(color)}<a:ln><a:noFill/></a:ln></cdr:spPr></cdr:sp>`);
}
function userShapesXml(shapes) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${shapes.join("")}</c:userShapes>`;
}

/* Per category: the tops of the positive and the bottoms of the negative stacks */
function stackExtents(list, count) {
  const out = [];
  for (let ci = 0; ci < count; ci++) {
    let up = 0, down = 0;
    for (const s of list) { const v = s.values[ci]; if (typeof v === "number") { if (v >= 0) up += v; else down += v; } }
    out.push(up, down);
  }
  return out;
}

/*
 * The chart part for a chart plan, in the visual's look (plan.look): title and subtitle where the
 * report puts them, legend (with its title) where the report puts it, fonts, gridlines, axis
 * ranges and number formats as Power BI draws them, data labels, total labels, constant and
 * analytics lines, trend lines, the visual's background, border and shadow.
 * Returns { xml, shapes }: shapes is the chart's drawing of shapes over it, or null.
 */
function chartXml(plan) {
  const look = plan.look || chartLook({ type: "" }, null);
  const { spec, series } = plan;
  const kind = spec.kind;
  const refs = plan.refs;
  const W = (plan.sizePx && plan.sizePx.w) || 400, H = (plan.sizePx && plan.sizePx.h) || 300;
  const scatter = kind === "scatter" || kind === "bubble";
  const round = kind === "pie" || kind === "doughnut";
  const shapes = [];
  const catLabel = (i) => String(plan.catText ? plan.catText[i] : axisLabel(plan.categories[i]));

  // Title (and subtitle) at the top, aligned as in the report
  const T = look.title;
  const showTitle = plan.showTitle !== false && !!plan.title;
  // Lines broken where Power BI breaks them (Excel would wrap a title at three quarters of the chart)
  const titleLines = showTitle ? wrapLines(plan.title, T, W - 16) : [];
  const subLines = showTitle && look.subtitle ? wrapLines(look.subtitle.text, look.subtitle, W - 16) : [];
  const titleH = showTitle ? titleLines.length * lineHeightPx(T) + subLines.length * lineHeightPx(look.subtitle) + 8 : 4;
  let title = `<c:autoTitleDeleted val="1"/>`;
  if (showTitle) {
    const algn = T.align === "center" ? "ctr" : T.align === "right" ? "r" : "l";
    const paras = titleLines.map((l) => paraA(l, T, algn)).join("") + subLines.map((l) => paraA(l, look.subtitle, algn)).join("");
    const widest = Math.min(W - 16, Math.max(...titleLines.map((l) => textPx(l, T)), ...subLines.map((l) => textPx(l, look.subtitle))));
    const layout = algn === "ctr" ? "" : manualLayoutA(algn === "l" ? 8 / W : 1 - (widest + 10) / W, 4 / H);
    title = `<c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz" wrap="none" anchor="ctr" anchorCtr="1"/><a:lstStyle/>${paras}</c:rich></c:tx>${layout}<c:overlay val="0"/>` +
      `${T.background ? `<c:spPr>${fillA(T.background)}</c:spPr>` : ""}</c:title><c:autoTitleDeleted val="0"/>`;
  }

  // Legend: one entry per visible series (a slice, a scatter group); shown as Power BI shows it
  const legendNames = scatter ? (plan.perPoint ? [] : plan.groups.map((g) => g.name))
    : round ? plan.categories.map((_, i) => catLabel(i)) : series.filter((s) => !s.noLegend).map((s) => s.name);
  const legendOn = look.legend.show !== false && (round ? legendNames.length > 0 : legendNames.length > 1);
  const LT = look.legend.text;
  const [side, align] = LEGEND_SPOTS[String(look.legend.position).toLowerCase()] || ["t", "start"];
  const legendTitle = legendOn && look.legend.title ? look.legend.titleText || plan.legendName || null : null;
  const legendTitleFont = { ...LT, bold: true };
  const lineH = lineHeightPx(LT) + 2;
  // Each entry's width as Excel draws it: its key (a line's key is wider than a box's), the
  // name and the gap to the next entry
  const legendLines = scatter || round ? [] : series.filter((s) => !s.noLegend).map((s) => kind === "line" || !!s.line);
  const entryW = legendNames.map((n, i) => textPx(n, LT) * 1.12 + (legendLines[i] ? 40 : 26));
  const titleW = legendTitle ? textPx(legendTitle, legendTitleFont) + 8 : 0;
  const across = side === "t" || side === "b";
  // Excel lays a top or bottom legend out in equal columns, each as wide as its widest entry
  const room = Math.max(40, W - 16 - titleW);
  const colW = Math.max(20, ...entryW);
  const perRow = Math.max(1, Math.min(legendNames.length, Math.floor(room / colW)));
  // A box a little too small makes Excel leave entries out, so it gets room to spare
  const rowH = lineH + 5;
  // More entries than fit (a slice per product): the first ones, as Power BI's legend shows its
  // first entries and scrolls to the rest, so the legend never squeezes or covers the chart.
  // A top or bottom legend takes at most 30% of the height, a side legend the height there is.
  // (a chart made taller for its legend, plan.fullLegend, shows every entry)
  const maxRows = plan.fullLegend ? Infinity : across ? Math.max(1, Math.floor((H - titleH) * 0.3 / rowH))
    : Math.max(1, Math.floor((H - titleH - 12 - (legendTitle ? lineH : 0)) / rowH));
  const legendIdx = scatter ? null : round ? legendNames.map((_, i) => i) : series.map((s, i) => (s.noLegend ? -1 : i)).filter((i) => i >= 0);
  const shownCount = legendIdx ? Math.min(legendNames.length, across ? perRow * maxRows : maxRows) : legendNames.length;
  const legendCut = legendIdx ? legendIdx.slice(shownCount) : [];
  const legendRows = Math.max(1, Math.ceil(shownCount / perRow));
  const legendW = across ? Math.min(room, perRow * colW + 12) : Math.min(W * 0.4, Math.max(titleW, ...entryW.slice(0, shownCount), 20) + 12);
  const legendH = across ? rowH * legendRows + 6 : rowH * shownCount + 6 + (legendTitle ? lineH : 0);
  let legendBox = null;
  if (legendOn) {
    let x, y;
    if (across) {
      x = align === "center" ? Math.max(8, (W - titleW - legendW) / 2) : align === "end" ? Math.max(8, W - titleW - legendW - 8) : 8;
      y = side === "t" ? titleH + 2 : H - legendH - 6;
      if (legendTitle) shapes.push(textShapeXml(shapes.length + 2, legendTitle, legendTitleFont, x / W, y / H, titleW / W, lineH / H));
      x += titleW;
    } else {
      x = side === "l" ? 6 : W - legendW - 8;
      y = align === "center" ? Math.max(titleH + 4, (H - legendH) / 2) : titleH + 6;
      if (legendTitle) { shapes.push(textShapeXml(shapes.length + 2, legendTitle, legendTitleFont, x / W, y / H, Math.max(titleW, 20) / W, lineH / H)); y += lineH; }
    }
    // The legend's whole box, not just its corner: given only a corner, Excel sizes the legend
    // itself and slides it left when its entries run wider, over the legend's title
    legendBox = { x, y, w: Math.min(legendW, W - x - 4), h: across ? legendH : legendH - (legendTitle ? lineH : 0) };
  }

  let body = "";
  let hidden = series.map((s, i) => (s.noLegend && !s.dataOnly ? i : -1)).filter((i) => i >= 0);
  let plotLayout = "<c:layout/>";
  if (scatter) {
    body = scatterXml(plan);
    hidden = plan.hiddenLegend || [];
  } else if (kind === "waterfall") {
    body = waterfallXml(plan);
  } else if (round) {
    // Pie and donut: slices from 12 o'clock, labels outside with leader lines. A donut is a pie
    // with its hole drawn over the centre, so its labels can sit outside as Power BI's do;
    // the pie is placed by hand in the room the title, legend and labels leave.
    let pts = look.labels.show === false ? [] : pieLabelPoints(plan, look);
    const startAngle = Math.round(((look.startAngle || 0) % 360 + 360) % 360);
    let x0 = 0, x1 = W, y0 = titleH, y1 = H;
    if (legendOn) {
      if (side === "l") x0 += legendW + 14; else if (side === "r") x1 -= legendW + 14;
      else if (side === "t") y0 += legendH + 4; else y1 -= legendH + 8;
    }
    const mx = pts.length ? Math.min(110, Math.max(...pts.flatMap((p) => p.text.split("\n").map((l) => textPx(l, look.labels.text)))) + 16) : 8;
    const my = pts.length ? lineHeightPx(look.labels.text) * Math.max(...pts.map((p) => p.text.split("\n").length)) + 6 : 8;
    const d = Math.max(40, Math.min(x1 - x0 - 2 * mx, y1 - y0 - 2 * my));
    const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
    plotLayout = manualLayoutA((cx - d / 2) / W, (cy - d / 2) / H, d / W, d / H, true);
    // Only the labels that fit beside their slice, as Power BI shows them (the rest on hover)
    pts = fitPieLabels(pts, series[0].values, startAngle, d / 2, lineHeightPx(look.labels.text));
    const ser = chartSeriesXml("pie", { ...series[0], labelsXml: pts.length ? textLabelsXml(pts, look.labels.text) : "" }, 0, plan, refs.series[0]);
    body = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="${startAngle}"/></c:pieChart>`;
    if (kind === "doughnut") {
      const ratio = look.donutHole !== null && look.donutHole !== undefined ? Math.min(0.95, Math.max(0.05, look.donutHole / 100 * 0.6 + 0.37)) : DONUT_HOLE;
      const hd = d * ratio;
      shapes.push(ellipseShapeXml(shapes.length + 2, look.frame.background || "#FFFFFF", (cx - hd / 2) / W, (cy - hd / 2) / H, hd / W, hd / H));
    }
  } else if (kind === "radar") {
    // A radar is a square web: placed by hand in the room the title, legend and the spoke
    // labels leave
    let x0 = 0, x1 = W, y0 = titleH, y1 = H;
    if (legendOn) {
      if (side === "l") x0 += legendW + 14; else if (side === "r") x1 -= legendW + 14;
      else if (side === "t") y0 += legendH + 4; else y1 -= legendH + 8;
    }
    const mx = Math.min(90, Math.max(20, ...plan.categories.map((_, i) => textPx(catLabel(i), look.cat.text))) + 12);
    const my = lineHeightPx(look.cat.text) + 10;
    const d = Math.max(40, Math.min(x1 - x0 - 2 * mx, y1 - y0 - 2 * my));
    const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
    plotLayout = manualLayoutA((cx - d / 2) / W, (cy - d / 2) / H, d / W, d / H, true);
    body = radarXml(plan, look, catLabel);
  } else {
    body = cartesianXml(plan, look, { titleH, legendOn, side, legendW, legendH, catLabel }, hidden);
    // Small multiples: every panel's plot the same size and place, room kept for the axes
    if (plan.panelEdges && !legendOn) plotLayout = panelPlotLayout(plan, look, titleH, catLabel, W, H);
  }

  const hiddenEntries = [...new Set([...hidden, ...legendCut])].sort((a, b) => a - b).map((i) => `<c:legendEntry><c:idx val="${i}"/><c:delete val="1"/></c:legendEntry>`).join("");
  const legend = legendOn
    ? `<c:legend><c:legendPos val="${side}"/>${hiddenEntries}${manualLayoutA(legendBox.x / W, legendBox.y / H, legendBox.w / W, legendBox.h / H)}<c:overlay val="0"/>${txPrA(LT)}</c:legend>` : "";
  const F = look.frame;
  const shadow = F.shadow
    ? `<a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="${hexOf(F.shadow.color)}"><a:alpha val="${Math.round((1 - F.shadow.transparency / 100) * 100000)}"/></a:srgbClr></a:outerShdw></a:effectLst>` : "";
  const frame = `<c:spPr>${F.background ? fillA(F.background, 1 - (F.transparency || 0) / 100) : "<a:noFill/>"}` +
    `${F.border ? lineA({ color: F.border.color, width: F.border.width }) : "<a:ln><a:noFill/></a:ln>"}${shadow}</c:spPr>`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:roundedCorners val="${F.border && F.border.radius > 0 ? 1 : 0}"/><c:chart>${title}<c:plotArea>${plotLayout}${body}${NO_BOX}</c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${frame}${txPrA(LT)}` +
    (shapes.length ? `<c:userShapes r:id="${USER_SHAPES_RID}"/>` : "") + `</c:chartSpace>`;
  // What the legend left out, and the rows it would need to show everything (legendRoom)
  const legendNeed = { cut: legendCut.length, rowH, extraRows: legendCut.length ? Math.ceil(legendNames.length / (across ? perRow : 1)) - (across ? legendRows : shownCount) : 0 };
  return { xml, shapes: shapes.length ? userShapesXml(shapes) : null, legendNeed };
}

/*
 * Radar (the AppSource Radar Chart): Excel's filled radar, a spoke per category and a
 * polygon per measure in its colour, see-through so overlapping polygons stay visible as in
 * the visual, with rings of the value axis behind them
 */
function radarXml(plan, look, catLabel) {
  const { series, refs } = plan;
  const A1 = 60010, A2 = 60020;
  const cats = plan.categories.map((_, i) => catLabel(i));
  const sers = series.map((s, i) => {
    const color = s.color || null;
    const sp = color ? `<c:spPr>${fillA(color, 0.45)}${lineA({ color, width: 2 })}</c:spPr>` : "";
    const ref = refs.series[i];
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>${xmlEsc(ref.nameRef)}</c:f>${strCacheXml([s.name])}</c:strRef></c:tx>${sp}` +
      `<c:cat><c:strRef><c:f>${xmlEsc(refs.catRef)}</c:f>${strCacheXml(cats)}</c:strRef></c:cat>` +
      `<c:val><c:numRef><c:f>${xmlEsc(ref.valRef)}</c:f>${numCacheXml(s.values, s.fmt)}</c:numRef></c:val></c:ser>`;
  }).join("");
  const values = series.flatMap((s) => s.values).filter((v) => typeof v === "number");
  const fmt = axisNumFmt((series[0] || {}).fmt, null, look.val, values);
  const web = lineA({ color: GRID_GREY, width: 1 });
  return `<c:radarChart><c:radarStyle val="filled"/><c:varyColors val="0"/>${sers}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:radarChart>` +
    `<c:catAx><c:axId val="${A1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
    `<c:majorGridlines><c:spPr>${web}</c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/>` +
    `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr>${web}</c:spPr>${txPrA(look.cat.text)}` +
    `<c:crossAx val="${A2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${A2}"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
    `<c:majorGridlines><c:spPr>${web}</c:spPr></c:majorGridlines><c:numFmt formatCode="${xmlEsc(fmt || "General")}" sourceLinked="0"/>` +
    // No ring values, as the visual shows none (they'd sit over the top spoke's polygons)
    `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPrA(look.val.text)}` +
    `<c:crossAx val="${A1}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

/*
 * Bar, column, line, area and combo charts: their chart groups and axes. hidden collects the
 * series kept out of the legend (overlay lines, total labels, analytics lines).
 */
/*
 * A small multiple's plot area, the same in every panel as Power BI draws them: room on the left
 * for the names (bars) or numbers (columns, lines) and below for the axis under it, kept in every
 * panel, so the bars line up across the grid and a name is never cut ("Weiler,...") or wrapped
 */
function panelPlotLayout(plan, look, titleH, catLabel, W, H) {
  const horizontal = plan.spec.dir === "bar";
  const nameW = Math.max(0, ...plan.categories.map((_, i) => textPx(String(catLabel(i)), look.cat.text)));
  const values = plan.series.flatMap((s) => s.values).filter((v) => typeof v === "number" && isFinite(v));
  const top = plan.panelTop ?? maxAbsOf(values);   // the axis all panels share
  const fmt = (plan.series[0] || {}).fmt;
  const numberW = textPx(displayValue(top, fmt, autoUnit(top), 1), horizontal ? look.val.text : look.val.text) + 10;
  // room for the names (and Excel's gap before the plot), or for the numbers
  const left = Math.min(W * 0.55, horizontal ? nameW + 18 : numberW);
  // the bottom axis's first and last numbers sit half outside the plot
  const half = horizontal ? Math.max(10, numberW / 2) : 10;
  const pw = Math.max(20, W - left - half);
  // under the plot: one line of labels, or slanted names when they don't fit across
  let below = lineHeightPx(horizontal ? look.val.text : look.cat.text) + 8;
  if (!horizontal && plan.categories.length * 4 + plan.categories.reduce((a, _, i) => a + textPx(String(catLabel(i)), look.cat.text), 0) > pw) {
    below = Math.min(H * 0.4, nameW * 0.75 + 10);
  }
  const y = titleH + 4, ph = Math.max(20, H - y - below);
  const x = plan.panelEdges.left ? left : Math.min(left, half);
  return manualLayoutA(x / W, y / H, pw / W, ph / H, true);
}

function cartesianXml(plan, look, box, hidden) {
  const { spec, series, refs } = plan;
  const kind = spec.kind;
  const W = (plan.sizePx && plan.sizePx.w) || 400, H = (plan.sizePx && plan.sizePx.h) || 300;
  const indexed = series.map((s, i) => [s, i]);
  const main = indexed.filter(([s]) => !s.line);
  const lineSeries = indexed.filter(([s]) => s.line);
  // A combo whose Y-axis the report turns off: Power BI draws no second axis either, and its line
  // on the columns' own scale (Plan's $13.9M just under Actual's $15.2M)
  // (also a combo whose columns a legend splits: Power BI draws no second axis for its line then)
  const shareScale = lineSeries.length > 0 && kind !== "line" && (!!(plan.axes && plan.axes.val && plan.axes.val.show === false) || !!plan.legendSplit);
  const lines = shareScale ? [] : lineSeries;
  const sameScaleLines = shareScale ? lineSeries : [];
  const horizontal = spec.dir === "bar";
  // One series "stacked" on nothing (Power BI's bar and column charts with one measure) is drawn
  // as plain bars: they look the same, and only then can Excel put labels past the bar's end, as
  // Power BI does
  const single = kind === "bar" && spec.grouping === "stacked" && !lines.length && main.filter(([s]) => !s.hidden).length === 1 && !(look.totals && look.totals.show);
  const grouping = single ? "clustered" : spec.grouping || (kind === "area" ? "standard" : "clustered");
  const pct = grouping === "percentStacked", stacked = grouping === "stacked";
  const isBar = kind === "bar" || kind === "combo";
  const shown = main.filter(([s]) => !s.hidden).map(([s]) => s);
  const n = plan.categories.length;
  const A1 = 50010, A2 = 50020, A3 = 50030, A4 = 50040, A7 = 50070, A8 = 50080;
  let nextIdx = series.length;
  const overlays = [];   // line series on the main axes: area tops, trend helper

  // The value axis like Power BI's: room for its labels, a tick every ~55 px up or ~100 px across
  const titles = axisTitles(plan, main, lines);
  const mainVals = (stacked ? stackExtents(main.map(([s]) => s), n) : main.flatMap(([s]) => s.values))
    .concat(sameScaleLines.flatMap(([s]) => s.values));
  // A tornado's axis runs as far left as right
  if (spec.mirror) { const m = maxAbsOf(mainVals); mainVals.push(m, -m); }
  const refUsable = (l) => l.kind !== "constant" || l.axis === (horizontal ? "x" : "y");
  const refConst = look.refLines.filter((l) => l.kind === "constant" && refUsable(l)).map((l) => l.value);
  const legendBand = box.legendOn ? (box.side === "t" || box.side === "b" ? box.legendH + 6 : 0) : 0;
  const legendSide = box.legendOn && (box.side === "l" || box.side === "r") ? box.legendW + 14 : 0;
  const catW = horizontal ? Math.min(W * 0.35, Math.max(30, ...plan.categories.map((_, i) => textPx(box.catLabel(i), look.cat.text)))) + 8 : 0;
  const plotH = H - box.titleH - legendBand - (titles.cat ? 22 : 0) - 28;
  const plotW = W - catW - legendSide - (titles.val ? 22 : 0) - 30;
  // Power BI runs a bar axis on to the next tick only to make room for labels outside the bars
  const labelsOn = look.labels.show === true;
  const roomForLabels = isBar && ((labelsOn && (grouping === "clustered" || /outside/i.test(look.labels.position))) || (stacked && look.totals.show));
  const axisPx = horizontal ? plotW : plotH, gapPx = horizontal ? 200 : 75;
  const axis = look.val.log ? null
    : pct ? valueScale([0, 1], axisPx, gapPx, { zero: true, round: true, start: 0, end: 1 })
      : valueScale([...mainVals, ...refConst], axisPx, gapPx, { zero: isBar, round: roomForLabels, start: look.val.start, end: look.val.end });
  if (spec.mirror && axis) { axis.max = Math.max(axis.max, -axis.min); axis.min = -axis.max; }
  if (plan.axisEnds && axis) Object.assign(axis, plan.axisEnds, { step: plan.axisEnds.max - plan.axisEnds.min || 1 });
  const fmt0 = (main[0] ? main[0][0] : series[0]).fmt;
  // Display units "Auto" for the labels on the main axis follow where the axis ends: the
  // gridline at or above the largest value (8,468 under a 10K gridline shows 8.5K; 2,128 under
  // 2,500 stays 2,128), taken before any room added for the labels themselves
  const gridTop = axis && !pct && axis.step > 0 ? Math.ceil(maxAbsOf(mainVals) / axis.step - 1e-9) * axis.step : 0;
  const labelAuto = (m) => labelAutoUnit(m > 0 ? Math.max(m, gridTop) : m);
  // Labels past the bar's end need room before the axis ends: as Power BI, the axis runs on a
  // step at a time until the largest label fits (not when the report sets the axis's end)
  const outside = roomForLabels && !pct && !spec.mirror && !/inside|center|base/i.test(String(look.labels.position || ""));
  if (outside && axis && axis.step > 0 && (look.val.end === null || look.val.end === undefined)) {
    const tops = stacked
      ? plan.categories.map((_, ci) => shown.reduce((a, s) => a + (typeof s.values[ci] === "number" ? s.values[ci] : 0), 0))
      : mainVals.filter((v) => typeof v === "number");
    const top = Math.max(0, ...tops);
    if (top > 0) {
      const L = stacked && look.totals.show ? look.totals : look.labels;
      const unit = L.units > 0 ? L.units : labelAuto(maxAbsOf(tops));
      // the widest label as it shows: shortened labels take a decimal unless every one is ten or more of the unit
      const dec = L.precision ?? (unit > 1 && fmtDecimals(String(fmt0 || "")) === 0 ? (tops.every((v) => !v || Math.abs(v) >= 9.95 * unit) ? 0 : 1) : null);
      const text = displayValue(top, fmt0, unit, dec);
      const labelPx = horizontal ? textPx(text, L.text) + 10 : lineHeightPx(L.text) + 6;
      const fits = () => ((top - axis.min) / (axis.max - axis.min)) * axisPx + labelPx <= axisPx;
      // first to the next whole step (the axis ends on a gridline), then a step more at a time
      if (!fits()) axis.max = Number((Math.ceil(axis.max / axis.step - 1e-9) * axis.step).toFixed(12));
      for (let i = 0; i < 10 && !fits(); i++) axis.max = Number((axis.max + axis.step).toFixed(12));
    }
    // A negative bar's label sits past its end too, towards the names: the axis runs on below
    // zero, a step at a time, until "-0.61M" fits clear of them (not when the report sets the start)
    const bottom = Math.min(0, ...tops);
    if (bottom < 0 && (look.val.start === null || look.val.start === undefined)) {
      const L = stacked && look.totals.show ? look.totals : look.labels;
      const unit = L.units > 0 ? L.units : labelAuto(maxAbsOf(tops));
      const dec = L.precision ?? (unit > 1 && fmtDecimals(String(fmt0 || "")) === 0 ? (tops.every((v) => !v || Math.abs(v) >= 9.95 * unit) ? 0 : 1) : null);
      const labelPx = horizontal ? textPx(displayValue(bottom, fmt0, unit, dec), L.text) + 10 : lineHeightPx(L.text) + 6;
      const fitsLow = () => ((bottom - axis.min) / (axis.max - axis.min)) * axisPx >= labelPx;
      if (!fitsLow()) axis.min = Number((Math.floor(axis.min / axis.step + 1e-9) * axis.step).toFixed(12));
      for (let i = 0; i < 10 && !fitsLow(); i++) axis.min = Number((axis.min - axis.step).toFixed(12));
    }
  }
  const valFmt = pct ? stepFormat("0%", axis, []) : spec.mirror ? unsignedNumFmt(axisNumFmt(fmt0, axis, look.val, mainVals)) : axisNumFmt(fmt0, axis, look.val, mainVals);
  // Labels take their units from the chart's size: a stacked chart's stack totals, else its values
  const labelVals = stacked
    ? plan.categories.map((_, ci) => shown.reduce((a, s) => a + (typeof s.values[ci] === "number" ? Math.abs(s.values[ci]) : 0), 0))
    : shown.flatMap((s) => s.values);
  const lineLookOf = (s) => look.lines.series[normalizeField(s.name)] || look.lines;
  const trendOn = !!look.trend && !pct && !stacked && (kind === "line" || (isBar && grouping === "clustered"));
  const combineTrend = trendOn && look.trend.combine && shown.length > 1;

  let groups = "";
  if (isBar) {
    // Bars fill the share of each category Power BI gives them (its inner padding: about a
    // quarter of the category is gap); clustered bars share that room
    const share = look.cat.padding === null || look.cat.padding === undefined ? 0.25 : Math.min(0.9, Math.max(0, look.cat.padding / 100 + 0.05));
    const per = grouping === "clustered" ? Math.max(1, shown.length) : 1;
    const gapWidth = spec.gap !== undefined ? spec.gap : Math.round(Math.min(500, share / (1 - share) * per * 100));
    // Labels inside the bars: Power BI leaves out the ones a bar is too small to hold
    const insideBars = labelsOn && axis && axis.max > axis.min && labelPosOf("bar", grouping, look.labels) !== "outEnd";
    const tooSmall = (s) => {
      if (!insideBars) return [];
      const unit = look.labels.units > 0 ? look.labels.units : labelAuto(maxAbsOf(labelVals));
      return s.values.map((v, i) => {
        if (typeof v !== "number") return -1;
        const barPx = (Math.abs(v) / (axis.max - axis.min)) * axisPx;
        const needPx = horizontal ? textPx(displayValue(v, s.fmt, unit, look.labels.precision ?? (unit > 1 ? 1 : null)), look.labels.text) + 6 : lineHeightPx(look.labels.text) + 4;
        return barPx < needPx ? i : -1;
      }).filter((i) => i >= 0);
    };
    let sers = main.map(([s, i]) => chartSeriesXml("bar", {
      ...s,
      labelsXml: labelsOn && !s.hidden ? seriesLabelsXml("bar", grouping, s.color || (plan.pointColors || []).find(Boolean),
        spec.mirror ? unsignedNumFmt(labelNumFmt(s.fmt, labelVals, look.labels, 0, labelAuto)) : labelNumFmt(s.fmt, labelVals, look.labels, 0, labelAuto), look.labels, tooSmall(s)) : "",
      trendXml: trendOn && !combineTrend && !s.hidden ? trendlineXml(look.trend) : "",
    }, i, plan, refs.series[i])).join("");
    // Total labels over each stack: an empty last segment labelled with the stack's total
    if (stacked && look.totals.show && !spec.hideValueAxis) {
      const totals = plan.categories.map((_, ci) => shown.reduce((a, s) => a + (typeof s.values[ci] === "number" ? s.values[ci] : 0), 0));
      const unit = look.totals.units > 0 ? look.totals.units : labelAuto(maxAbsOf(totals));
      const dec = look.totals.precision !== null && look.totals.precision !== undefined ? look.totals.precision : null;
      const pts = totals.map((v, idx) => ({ idx, text: displayValue(v, fmt0, unit, dec), pos: "inBase" }));
      const idx = nextIdx++;
      hidden.push(idx);
      sers += chartSeriesXml("bar", { name: "Total", values: totals.map(() => 0), fmt: "General", literal: true, hidden: true, labelsXml: textLabelsXml(pts, look.totals.text) }, idx, plan, null);
    }
    groups += `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}` +
      `<c:gapWidth val="${gapWidth}"/>${grouping !== "clustered" || spec.mirror ? '<c:overlap val="100"/>' : ""}` +
      `<c:axId val="${A1}"/><c:axId val="${A2}"/></c:barChart>`;
  } else if (kind === "area") {
    // Power BI's area: a see-through fill under a solid line along its top
    const tops = [];
    let running = plan.categories.map(() => 0);
    const totals = plan.categories.map((_, ci) => main.reduce((a, [s]) => a + Math.abs(typeof s.values[ci] === "number" ? s.values[ci] : 0), 0) || 1);
    for (const [s] of main) {
      running = running.map((r, ci) => r + (typeof s.values[ci] === "number" ? s.values[ci] : 0));
      tops.push({ s, values: running.map((r, ci) => (s.values[ci] === null || s.values[ci] === undefined ? null : grouping === "standard" ? s.values[ci] : pct ? r / totals[ci] : r)) });
    }
    groups += `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>` +
      main.map(([s, i]) => chartSeriesXml("area", { ...s, alpha: look.areaOpacity,
        labelsXml: labelsOn && grouping !== "standard" ? seriesLabelsXml("area", grouping, s.color, labelNumFmt(s.fmt, labelVals, look.labels, 0, labelAuto), look.labels) : "" }, i, plan, refs.series[i])).join("") +
      `<c:axId val="${A1}"/><c:axId val="${A2}"/></c:areaChart>`;
    for (const t of tops) {
      overlays.push({ ...t.s, name: t.s.name, values: t.values, literal: true, noLegend: true, lineLook: lineLookOf(t.s),
        labelsXml: labelsOn && grouping === "standard" ? seriesLabelsXml("line", grouping, t.s.color, labelNumFmt(t.s.fmt, labelVals, look.labels, 0, labelAuto), look.labels) : "" });
    }
    // Total labels over a stacked area (Power BI's "Total labels"): an unseen line along the
    // stacks' tops, labelled with each total above it
    if (stacked && look.totals.show) {
      const sums = plan.categories.map((_, ci) => main.reduce((a, [s]) => a + (typeof s.values[ci] === "number" ? s.values[ci] : 0), 0));
      const unit = look.totals.units > 0 ? look.totals.units : labelAuto(maxAbsOf(sums));
      const dec = look.totals.precision !== null && look.totals.precision !== undefined ? look.totals.precision : null;
      const pts = sums.map((v, idx) => ({ idx, text: displayValue(v, fmt0, unit, dec), pos: "t" }));
      overlays.push({ name: "Total", values: sums, fmt: fmt0, literal: true, hidden: true, noLegend: true, labelsXml: textLabelsXml(pts, look.totals.text) });
    }
  }
  if (combineTrend) {
    // One trend over the series together (Power BI's "combine series"): through their average
    const mean = plan.categories.map((_, ci) => { const v = shown.map((s) => s.values[ci]).filter((x) => typeof x === "number"); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; });
    overlays.push({ name: "Trend", values: mean, fmt: fmt0, literal: true, hidden: true, noLegend: true, trendXml: trendlineXml(look.trend) });
  }
  // A combo's line the report draws "stepped" (flat from one value to the next): drawn as steps
  // below, on the hidden axes the analytics lines use; any other line as before
  const sameVals = sameScaleLines.flatMap(([s]) => s.values);
  const steppedLine = ([s]) => !horizontal && !plan.dateAxis && !plan.multiLevel && n > 1 && !!lineLookOf(s).stepped;
  const sameScaleFlat = sameScaleLines.filter((e) => !steppedLine(e));
  const sameScaleSteps = sameScaleLines.filter(steppedLine);
  if (kind === "line") {
    groups += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
      main.map(([s, i]) => chartSeriesXml("line", {
        ...s, lineLook: lineLookOf(s),
        labelsXml: labelsOn ? seriesLabelsXml("line", "standard", s.color, labelNumFmt(s.fmt, labelVals, look.labels, 0, labelAuto), look.labels) : "",
        trendXml: trendOn && !combineTrend ? trendlineXml(look.trend) : "",
      }, i, plan, refs.series[i])).join("") +
      overlays.map((s) => { const idx = nextIdx++; hidden.push(idx); return chartSeriesXml("line", s, idx, plan, null); }).join("") +
      `<c:marker val="1"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:lineChart>`;
  } else if (overlays.length || sameScaleFlat.length) {
    groups += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
      sameScaleFlat.map(([s, i]) => chartSeriesXml("line", { ...s, lineLook: lineLookOf(s),
        labelsXml: labelsOn ? seriesLabelsXml("line", "standard", s.color, labelNumFmt(s.fmt, sameVals, look.labels, 0, labelAuto), look.labels) : "" }, i, plan, refs.series[i])).join("") +
      overlays.map((s) => { const idx = nextIdx++; hidden.push(idx); return chartSeriesXml("line", s, idx, plan, null); }).join("") +
      `<c:marker val="1"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:lineChart>`;
  }

  // Axes. A funnel hides its value axis and gridlines: the offsets would make the numbers misleading
  // A small multiple's panel shows an axis only on the grid's outer edge (labels off, gridlines kept)
  const edges = plan.panelEdges || null;
  const catOut = !edges || (horizontal ? edges.left : edges.bottom), valOut = !edges || (horizontal ? edges.bottom : edges.left);
  const catXml = plan.dateAxis
    ? dateAxXml(A1, A2, { hidden: !titles.catShow || !catOut, title: catOut ? titles.cat : null, look: look.cat, date: plan.dateAxis })
    : catAxXml(A1, A2, horizontal ? "l" : "b", { reverse: horizontal, title: catOut ? titles.cat : null, hidden: !titles.catShow || !catOut, count: n, look: look.cat, multi: plan.multiLevel,
        noWrap: !!edges && horizontal });   // a panel's names on one line, in the room its layout keeps
  let axes = catXml + valAxXml(A2, A1, horizontal ? "b" : "l", valFmt, { crosses: horizontal ? "max" : "autoZero", grid: !spec.hideValueAxis, labels: valOut,
    hidden: !!spec.hideValueAxis || !titles.valShow, title: spec.hideValueAxis || !valOut ? null : titles.val, axis,
    zero: isBar && main.every(([s]) => s.values.every((v) => v === null || v >= 0)), look: look.val });
  if (lines.length) {
    // A combo's line values on the second axis, scaled like Power BI's
    const lineVals = lines.flatMap(([s]) => s.values);
    const secAxis = look.sec.log ? null : valueScale(lineVals, plotH, 75, { round: false, start: look.sec.start, end: look.sec.end });
    groups += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
      lines.map(([s, i]) => chartSeriesXml("line", { ...s, lineLook: lineLookOf(s),
        labelsXml: labelsOn ? seriesLabelsXml("line", "standard", s.color, labelNumFmt(s.fmt, lineVals, look.labels), look.labels) : "" }, i, plan, refs.series[i])).join("") +
      `<c:marker val="1"/><c:axId val="${A3}"/><c:axId val="${A4}"/></c:lineChart>`;
    // the second axis's line follows the same dates as the columns under it
    axes += (plan.dateAxis ? dateAxXml(A3, A4, { hidden: true, date: plan.dateAxis }) : catAxXml(A3, A4, "b", { hidden: true })) +
      valAxXml(A4, A3, "r", axisNumFmt(lines[0][0].fmt, secAxis, look.sec, lineVals), { crosses: "max", grid: false, hidden: !titles.secShow, title: titles.sec, axis: secAxis, look: look.sec });
  }

  // Constant and analytics lines (average, median…): straight lines across the plot, on hidden
  // axes that share the value axis's range
  const refSers = [];
  if (axis && !spec.hideValueAxis) {
    for (const l of look.refLines) {
      if (!refUsable(l)) continue;
      const of = (m) => shown.filter((s) => !m || normalizeField(s.name) === m || normalizeField(plan.valueName || "") === m);
      const pool = of(l.measure).length ? of(l.measure) : shown;
      const v = l.kind === "constant" ? l.value : analyticValue(l, pool.flatMap((s) => s.values));
      if (v === null) continue;
      const idx = nextIdx++;
      hidden.push(idx);
      refSers.push(refLineSeriesXml(idx, l, horizontal ? [[v, 0], [v, 1]] : [[0, v], [1, v]], fmt0, horizontal, v));
    }
  }
  const legendless = () => { const idx = nextIdx++; hidden.push(idx); return idx; };
  if (axis) for (const e of sameScaleSteps) refSers.push(steppedLineSeriesXml(e, n, plan, refs, lineLookOf(e[0]),
    labelsOn ? { fmt: labelNumFmt(e[0].fmt, sameVals, look.labels, 0, labelAuto), labels: look.labels } : null, legendless));
  if (refSers.length) {
    const unitAxis = { min: 0, max: 1 };
    groups += `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${refSers.join("")}<c:axId val="${A7}"/><c:axId val="${A8}"/></c:scatterChart>`;
    axes += hiddenValAxXml(A7, A8, "b", horizontal ? axis : unitAxis) + hiddenValAxXml(A8, A7, "l", horizontal ? unitAxis : axis);
  }
  return groups + axes;
}

/* An Excel treemap or sunburst (chartEx): the levels and sizes from the Data sheet, each group in
   its colour, tiles labelled with their names and group names over each group, no legend */
function treemapXml(plan, n) {
  const s = plan.series[0], r = plan.refs;
  // Excel numbers a treemap's points in tree order: each group, then its tiles. A group's colour
  // goes on the group itself (its tiles take it); without Details each tile is a group of its own
  let idx = 0;
  const points = (plan.groups || []).map((g) => {
    const at = idx;
    idx += plan.nested ? 1 + g.tiles.length : 1;
    return g.color ? `<cx:dataPt idx="${at}"><cx:spPr>${fillXml(g.color)}</cx:spPr></cx:dataPt>` : "";
  }).join("");
  const look = plan.look || chartLook({ type: "treemap" }, null);
  const layout = plan.spec.kind === "sunburst" ? "sunburst" : "treemap";
  const T = look.title;
  const align = T.align === "center" ? "ctr" : T.align === "right" ? "max" : "min";
  const title = plan.showTitle === false ? "" : `<cx:title pos="t" align="${align}" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/>` +
    `${paraA(plan.title, T, align === "ctr" ? "ctr" : align === "max" ? "r" : "l")}</cx:rich></cx:tx></cx:title>`;
  const labels = `<cx:dataLabels pos="inEnd"><cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${runPrA({ ...look.labels.text, color: C.WHITE })}</a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>` +
    `<cx:visibility seriesName="0" categoryName="1" value="0"/></cx:dataLabels>`;
  const id = `{${String(n).padStart(8, "0")}-0000-4000-8000-00000000AE02}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${CHARTEX_NS}">` +
    `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${xmlEsc(r.catRef)}</cx:f></cx:strDim>` +
    `<cx:numDim type="size"><cx:f>${xmlEsc(r.series[0].valRef)}</cx:f></cx:numDim></cx:data></cx:chartData>` +
    `<cx:chart>${title}<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}" uniqueId="${id}">` +
    `<cx:tx><cx:txData><cx:f>${xmlEsc(r.series[0].nameRef)}</cx:f><cx:v>${xmlEsc(s.name)}</cx:v></cx:txData></cx:tx>${points}${labels}` +
    `<cx:dataId val="0"/>${layout === "treemap" ? `<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>` : ""}</cx:series></cx:plotAreaRegion></cx:plotArea>` +
    `</cx:chart>${chartExFrame(look, "")}</cx:chartSpace>`;   // no legend: the tiles name their groups, as in Power BI
}

/* A chartEx chart's frame (a treemap, a sunburst): Excel 2016 and later show the chart, older Excel a note */
function chartExAnchorXml(ch, rid, id) {
  const cx = Math.round(ch.wPx * CHART_PX_TO_EMU), cy = Math.round(ch.hPx * CHART_PX_TO_EMU);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"><xdr:graphicFrame macro="">` +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"${altAttr(ch.plan)}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CHARTEX_NS}">` +
    `<cx:chart xmlns:cx="${CHARTEX_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Chart ${id}"${altAttr(ch.plan)}/>` +
    `<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F4F6F5"/></a:solidFill></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>` +
    `<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>${xmlEsc(ch.plan.title)}: a ${chartKindName(ch.plan.spec)}, shown in Excel 2016 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>` +
    `</mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor>`;
}

/* Small multiples on the page sheet: one group (moved and sized as one) holding the visual's
   title and a chart per panel in its grid */
function panelGroupXml(ch, rids, id) {
  const E = CHART_PX_TO_EMU, g = ch.plan.panels;
  const W = Math.round(ch.wPx * E), H = Math.round(ch.hPx * E);
  const frame = (k) => {
    const x = (k % g.cols) * g.w * E, y = (g.head + Math.floor(k / g.cols) * g.h) * E;
    return `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id + 1 + k}" name="Chart ${id + 1 + k}"${altAttr(g.plans[k])}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(g.w * E)}" cy="${Math.round(g.h * E)}"/></xdr:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rids[k]}"/>` +
      `</a:graphicData></a:graphic></xdr:graphicFrame>`;
  };
  const T = g.look.title;
  const title = g.title
    ? `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Title ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${Math.round(g.head * E)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>` +
      `<xdr:txBody><a:bodyPr wrap="none" lIns="76200" tIns="0" rIns="76200" bIns="0" anchor="ctr"/><a:lstStyle/>${paraA(g.title, T, T.align === "center" ? "ctr" : T.align === "right" ? "r" : "l")}</xdr:txBody></xdr:sp>`
    : "";
  // Text over a panel's middle (a Dual KPI's label, change and value), and a line under the grid
  const textBox = (sid, x, y, w, h, paras, anchor = "ctr") =>
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${sid}" name="Text ${sid}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="${Math.round(x * E)}" y="${Math.round(y * E)}"/><a:ext cx="${Math.round(w * E)}" cy="${Math.round(h * E)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp>`;
  const overlays = g.plans.map((p, k) => {
    const o = p.overlay;
    if (!o) return "";
    const x = (k % g.cols) * g.w, y = g.head + Math.floor(k / g.cols) * g.h;
    return textBox(id + 200 + k, x + g.w * 0.2, y + g.h / 2 - o.height / 2, g.w * 0.6, o.height, o.lines.map((l) => paraA(l.text, l.font, "ctr")).join(""));
  }).join("");
  const foot = g.footer
    ? textBox(id + 300, 0, H / E - g.foot, W / E, g.foot, paraA(g.footer.text, g.footer.font, g.footer.align === "right" ? "r" : "l"))
    : "";
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${W}" cy="${H}"/><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${id + 100}" name="Small multiples ${id}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></xdr:grpSpPr>` +
    `${title}${g.plans.map((_, k) => frame(k)).join("")}${overlays}${foot}</xdr:grpSp><xdr:clientData/></xdr:oneCellAnchor>`;
}

/* A chart's frame on the page sheet: its top-left cell and its size */
function chartAnchorXml(ch, rid, id) {
  const cx = Math.round(ch.wPx * CHART_PX_TO_EMU), cy = Math.round(ch.hPx * CHART_PX_TO_EMU);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"${altAttr(ch.plan)}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
}

/* Excel's standard chart style and colour parts: a chartEx chart (a treemap) won't load without them */
const CHARTEX_STYLE_XML = `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="410"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:categoryAxis><cs:chartArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000"/></cs:chartArea><cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="lt1"/></cs:fontRef><cs:defRPr sz="900"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/><cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1"><a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout symbol="circle" size="5"/><cs:dataPointWireframe><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:dataTable><cs:downBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downBar><cs:dropLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorBar><cs:floor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:floor><cs:gridlineMajor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:legend><cs:plotArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea><cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1400"/></cs:title><cs:trendline><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:trendlineLabel><cs:upBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upBar><cs:valueAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:valueAxis><cs:wall><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:wall></cs:chartStyle>`;
// The same style with data labels in dark grey, as Excel styles a waterfall's (a treemap's are
// white, on its tiles); Excel draws chartEx labels in the style's colour
const CHARTEX_STYLE_DARK_LABELS = CHARTEX_STYLE_XML.replace(/(<cs:dataLabel>(?:(?!<\/cs:dataLabel>).)*?<cs:fontRef idx="minor">)<a:schemeClr val="lt1"\/>/,
  '$1<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>');
const CHARTEX_COLORS_XML = `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/><cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>`;

/* ── Adding the charts to the finished .xlsx ── */
/* ── Funnel charts: Excel's own funnel ───────────────────────────────────
 * Excel 2019 and Microsoft 365 have a real funnel chart ("chartEx", a newer chart format).
 * A funnel visual is written as one, with the centred-bar funnel as its fallback: Excel
 * versions without funnel charts, Excel on older Macs, LibreOffice and Google Sheets show the
 * bars instead, so every reader sees a funnel. NATIVE_FUNNEL = false keeps the bars only.
 */
const NATIVE_FUNNEL = true;
const CHARTEX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CHARTEX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const CHARTEX_TYPE = "application/vnd.ms-office.chartex+xml";

/* The funnel as Excel's own funnel chart: stages largest first, a colour per stage when the
   report sets them, values inside the bars */
function funnelChartExXml(plan) {
  const s = plan.series[1], ref = plan.refs.series[1];
  const cats = plan.categories.map(axisLabel);
  const pts = (list) => list.map((v, i) => (v === null || v === undefined || v === "" ? "" : `<cx:pt idx="${i}">${xmlEsc(v)}</cx:pt>`)).join("");
  const colors = plan.pointColors || [];
  const main = s.color || colors.find(Boolean) || toArgb(DEFAULT_THEME_COLORS[0]);
  const font = (argb, sz, b) => `<a:defRPr sz="${sz}" b="${b ? 1 : 0}">${fillXml(argb)}<a:latin typeface="Calibri"/></a:defRPr>`;
  const txPr = (argb, sz, b) => `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${font(argb, sz, b)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>`;
  const dataPts = colors.map((c, i) => (c ? `<cx:dataPt idx="${i}"><cx:spPr>${fillXml(c)}</cx:spPr></cx:dataPt>` : "")).join("");
  // The title as the report draws it, as on the other charts
  const look = plan.look || chartLook({ type: "funnel" }, null);
  const title = chartExTitleXml(plan, look);
  // The chart area in the visual's background (white without one); on a dark one the stage
  // names take the report's label colour, so they stay readable
  const bg = look.frame && look.frame.background ? toArgb(look.frame.background, C.WHITE) : C.WHITE;
  const catText = luminance(bg) < 0.5 ? toArgb(look.cat.text.color, C.SLATE) : C.SLATE;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${CHARTEX_NS}">` +
    `<cx:chartData><cx:data id="0">` +
    `<cx:strDim type="cat"><cx:f>${xmlEsc(plan.refs.catRef)}</cx:f><cx:lvl ptCount="${cats.length}">${pts(cats)}</cx:lvl></cx:strDim>` +
    `<cx:numDim type="val"><cx:f>${xmlEsc(ref.valRef)}</cx:f><cx:lvl ptCount="${s.values.length}" formatCode="${xmlEsc(s.fmt || "General")}">${pts(s.values)}</cx:lvl></cx:numDim>` +
    `</cx:data></cx:chartData>` +
    `<cx:chart>${title}` +
    `<cx:plotArea><cx:plotAreaRegion>` +
    `<cx:series layoutId="funnel" uniqueId="{6F7B2C1A-3E4D-4B5A-9C8D-1A2B3C4D5E6F}">` +
    `<cx:tx><cx:txData><cx:f>${xmlEsc(ref.nameRef)}</cx:f><cx:v>${xmlEsc(s.name)}</cx:v></cx:txData></cx:tx>` +
    `<cx:spPr>${fillXml(main)}</cx:spPr>${dataPts}` +
    `<cx:dataLabels pos="ctr"><cx:numFmt formatCode="${xmlEsc(labelNumFmt(s.fmt, s.values, look.labels) || "General")}" sourceLinked="0"/>${txPr(textOn(main), 900, true)}` +
    `<cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>` +
    `<cx:dataId val="0"/></cx:series>` +
    `</cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.06"/><cx:tickLabels/><cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr>${txPr(catText, 900, false)}</cx:axis></cx:plotArea>` +
    `</cx:chart>` +
    `<cx:spPr>${fillXml(bg)}<a:ln w="9525">${fillXml(C.HAIR)}</a:ln></cx:spPr>` +
    `</cx:chartSpace>`;
}

/* A chartEx chart's title, as the report draws it */
function chartExTitleXml(plan, look) {
  const T = look.title;
  const align = T.align === "center" ? "ctr" : T.align === "right" ? "max" : "min";
  return plan.showTitle === false || !plan.title ? "" : `<cx:title pos="t" align="${align}" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/>` +
    `${paraA(plan.title, T, align === "ctr" ? "ctr" : align === "max" ? "r" : "l")}</cx:rich></cx:tx></cx:title>`;
}
/* A chartEx text block in the report's font */
function chartExTxPr(t) {
  return `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${runPrA(t)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>`;
}
/* A chartEx legend where the report puts it (top left by default), below the title */
function chartExLegendXml(look) {
  const [side, align] = LEGEND_SPOTS[String(look.legend.position).toLowerCase()] || ["t", "start"];
  return `<cx:legend pos="${side}" align="${align === "center" ? "ctr" : align === "end" ? "max" : "min"}" overlay="0">${chartExTxPr(look.legend.text)}</cx:legend>`;
}
/* A chartEx axis title */
function chartExAxisTitle(text, t) {
  return text ? `<cx:title><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/>${paraA(text, t, "ctr")}</cx:rich></cx:tx><cx:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cx:spPr></cx:title>` : "";
}
/* A chartEx value axis like Power BI's: its range and steps, units, gridlines, labels */
function chartExValueAxis(id, axis, fmt, look, title, show) {
  const scale = axis ? `<cx:valScaling max="${axis.max}" min="${axis.min}"${axis.step ? ` majorUnit="${axis.step}"` : ""}/>` : "<cx:valScaling/>";
  return `<cx:axis id="${id}"${show ? "" : ' hidden="1"'}>${scale}${chartExAxisTitle(title, look.val.titleFont || look.val.text)}` +
    `<cx:majorGridlines><cx:spPr>${lineA({ color: GRID_GREY, width: 1 })}</cx:spPr></cx:majorGridlines><cx:tickLabels/>` +
    `<cx:numFmt formatCode="${xmlEsc(fmt)}" sourceLinked="0"/><cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr>${chartExTxPr(look.val.text)}</cx:axis>`;
}
const CHARTEX_SPACE = `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${CHARTEX_NS}">`;
const CHARTEX_FRAME = `<cx:spPr>${fillXml(C.WHITE)}<a:ln w="9525">${fillXml(C.HAIR)}</a:ln></cx:spPr>`;
/* A chartEx chart's frame in its visual's background ("Chart background: As in the report");
   white, as before, when the background is white or none ("" for a chart Excel frames itself) */
function chartExFrame(look, plain = CHARTEX_FRAME) {
  const bg = look && look.frame && look.frame.background ? toArgb(look.frame.background, C.WHITE) : C.WHITE;
  return bg === C.WHITE ? plain : `<cx:spPr>${fillXml(bg)}<a:ln w="9525">${fillXml(C.HAIR)}</a:ln></cx:spPr>`;
}
/* A chartEx part's colours: its own, in order (a waterfall's increase, decrease and total) */
function chartExColorsXml(colors) {
  return `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">` +
    colors.map((c) => `<a:srgbClr val="${hexOf(c)}"/>`).join("") + `<cs:variation/></cs:colorStyle>`;
}

/*
 * Excel's own waterfall (Excel 2016 and later): the signed steps, the last bar the total, each
 * bar in the report's increase, decrease or total colour, no connector lines (Power BI draws
 * none). Older Excel and other spreadsheet apps show the stacked-column waterfall instead.
 */
const NATIVE_WATERFALL = true;
function waterfallChartExXml(plan) {
  const s = plan.series[5], ref = plan.refs.series[5];
  const look = plan.look || chartLook({ type: "waterfallChart" }, null);
  const cats = plan.categories.map(axisLabel);
  const pts = (list) => list.map((v, i) => (v === null || v === undefined || v === "" ? "" : `<cx:pt idx="${i}">${xmlEsc(v)}</cx:pt>`)).join("");
  const last = s.values.length - 1;
  const colorOf = (i) => (i === last ? plan.series[3].color : s.values[i] < 0 ? plan.series[2].color : plan.series[1].color);
  const dataPts = s.values.map((_, i) => `<cx:dataPt idx="${i}"><cx:spPr>${fillXml(colorOf(i))}</cx:spPr></cx:dataPt>`).join("");
  // Labels as the stacked waterfall writes them: display units from 10,000 up (2.47M), else the
  // measure's own format; Excel places each outside its bar and sizes the axis to fit them
  const precision = plan.labelPrecision ?? look.labels.precision ?? null;
  const labelFmt = maxAbsOf(plan.steps) >= 1e4 || look.labels.units > 0 ? labelNumFmt(plan.fmt, plan.steps, look.labels, precision === null ? 2 : 0) : precision !== null ? unitFormat(plan.fmt, 1, precision) : plan.fmt;
  const labels = plan.showLabels
    ? `<cx:dataLabels><cx:numFmt formatCode="${xmlEsc(labelFmt || "General")}" sourceLinked="0"/>` +
      `${chartExTxPr(look.labels.text)}<cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>` : "";
  const titles = axisTitles(plan, [], []);
  const legend = look.legend.show === false ? "" : chartExLegendXml(look);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CHARTEX_SPACE}` +
    `<cx:chartData><cx:data id="0">` +
    `<cx:strDim type="cat"><cx:f>${xmlEsc(plan.refs.catRef)}</cx:f><cx:lvl ptCount="${cats.length}">${pts(cats)}</cx:lvl></cx:strDim>` +
    `<cx:numDim type="val"><cx:f>${xmlEsc(ref.valRef)}</cx:f><cx:lvl ptCount="${s.values.length}" formatCode="${xmlEsc(s.fmt || "General")}">${pts(s.values)}</cx:lvl></cx:numDim>` +
    `</cx:data></cx:chartData>` +
    `<cx:chart>${chartExTitleXml(plan, look)}<cx:plotArea><cx:plotAreaRegion>` +
    `<cx:series layoutId="waterfall" uniqueId="{2B5D8E41-7A3C-4F19-9E62-0C4B1D7A3E58}">` +
    `<cx:tx><cx:txData><cx:f>${xmlEsc(ref.nameRef)}</cx:f><cx:v>${xmlEsc(s.name)}</cx:v></cx:txData></cx:tx>${dataPts}${labels}` +
    `<cx:dataId val="0"/><cx:layoutPr><cx:visibility connectorLines="0"/><cx:subtotals><cx:idx val="${last}"/></cx:subtotals></cx:layoutPr></cx:series>` +
    `</cx:plotAreaRegion>` +
    `<cx:axis id="0"${titles.catShow ? "" : ' hidden="1"'}><cx:catScaling gapWidth="0.6"/>${chartExAxisTitle(titles.cat, look.cat.titleFont || look.cat.text)}<cx:tickLabels/>` +
    `<cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr>${chartExTxPr(look.cat.text)}</cx:axis>` +
    chartExValueAxis(1, plan.showLabels ? null : plan.axis, axisNumFmt(plan.fmt, plan.axis, look.val), look, titles.val, titles.valShow) +
    `</cx:plotArea>${legend}</cx:chart>${chartExFrame(look)}</cx:chartSpace>`;
}

/* Excel's own box & whisker chart: a box per category, with its mean marker and outliers */
function boxWhiskerXml(plan) {
  const r = plan.refs, series = plan.series;
  const look = plan.look || chartLook({ type: "boxWhisker" }, null);
  const pts = (list) => list.map((v, i) => (v === null || v === undefined || v === "" ? "" : `<cx:pt idx="${i}">${xmlEsc(v)}</cx:pt>`)).join("");
  const vals = series.flatMap((s) => s.values).filter((v) => typeof v === "number");
  const fmt = series[0].fmt || "General";
  const axis = valueScale(vals, ((plan.sizePx && plan.sizePx.h) || 300) - 100, 75, { round: false });
  const titles = axisTitles(plan, [[series[0]]], []);
  // A box per category: each category its own series (its colour, its legend entry), drawn side
  // by side with no category axis; else one series, its boxes along the category axis
  const cats = plan.perGroup ? null : plan.categories.map(axisLabel);
  const data = series.map((s, k) => `<cx:data id="${k}">` +
    (cats ? `<cx:strDim type="cat"><cx:f>${xmlEsc(r.catRef)}</cx:f><cx:lvl ptCount="${cats.length}">${pts(cats)}</cx:lvl></cx:strDim>` : "") +
    `<cx:numDim type="val"><cx:f>${xmlEsc(r.series[k].valRef)}</cx:f><cx:lvl ptCount="${s.values.length}" formatCode="${xmlEsc(s.fmt || "General")}">${pts(s.values)}</cx:lvl></cx:numDim></cx:data>`).join("");
  const boxes = series.map((s, k) => {
    const color = s.color || toArgb(DEFAULT_THEME_COLORS[k % DEFAULT_THEME_COLORS.length]);
    return `<cx:series layoutId="boxWhisker" uniqueId="{7C1E4A92-3B6D-4E85-A0F7-${String(k).padStart(12, "0")}}">` +
      `<cx:tx><cx:txData><cx:f>${xmlEsc(r.series[k].nameRef)}</cx:f><cx:v>${xmlEsc(s.name)}</cx:v></cx:txData></cx:tx>` +
      `<cx:spPr>${fillA(color, 0.85)}<a:ln w="9525">${fillXml(color)}</a:ln></cx:spPr>` +
      `<cx:dataId val="${k}"/><cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr></cx:series>`;
  }).join("");
  const legend = plan.perGroup && look.legend.show !== false ? chartExLegendXml(look) : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CHARTEX_SPACE}` +
    `<cx:chartData>${data}</cx:chartData>` +
    `<cx:chart>${chartExTitleXml(plan, look)}<cx:plotArea><cx:plotAreaRegion>${boxes}</cx:plotAreaRegion>` +
    `<cx:axis id="0"${titles.catShow && !plan.perGroup ? "" : ' hidden="1"'}><cx:catScaling gapWidth="${plan.perGroup ? 0.3 : 1}"/>` +
    `${plan.perGroup ? "" : chartExAxisTitle(titles.cat, look.cat.titleFont || look.cat.text)}<cx:tickLabels/>` +
    `<cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr>${chartExTxPr(look.cat.text)}</cx:axis>` +
    chartExValueAxis(1, axis, axisNumFmt(fmt, axis, look.val, vals), look, titles.val, titles.valShow) +
    `</cx:plotArea>${legend}</cx:chart>${chartExFrame(look)}</cx:chartSpace>`;
}

/* A chart's alt text, for screen readers and Excel's Alt Text pane: what it shows, and its
   largest value when it draws one measure */
const CHART_KIND_NAMES = { line: "line chart", area: "area chart", combo: "column and line chart", pie: "pie chart", doughnut: "donut chart",
  scatter: "scatter chart", bubble: "bubble chart", waterfall: "waterfall chart", funnel: "funnel chart", treemap: "treemap",
  sunburst: "sunburst chart", radar: "radar chart", boxWhisker: "box and whisker chart" };
function chartKindName(spec) {
  if (spec.mirror) return "tornado chart";
  if (spec.gap === 0) return "histogram";
  if (spec.kind === "bar") return spec.dir === "bar" ? "bar chart" : "column chart";
  return CHART_KIND_NAMES[spec.kind] || "chart";
}
function chartAltText(plan) {
  const kind = chartKindName(plan.spec);
  const shown = (plan.series || []).filter((s) => !s.noLegend && !s.hidden);
  const names = plan.spec.kind === "waterfall" ? [plan.valueName] : plan.legendName && plan.valueName ? [plan.valueName] : [...new Set(shown.map((s) => s.name))];
  const list = names.filter(Boolean).slice(0, 4);
  const what = list.length ? ` of ${list.length > 1 ? `${list.slice(0, -1).join(", ")} and ${list[list.length - 1]}` : list[0]}` : "";
  const byNames = [plan.catName, plan.legendName && plan.valueName ? plan.legendName : null].filter(Boolean);
  const by = byNames.length ? ` by ${byNames.join(" and ")}` : "";
  let text = `${plan.title ? plan.title + ": " : ""}${/^[aeiou]/i.test(kind) ? "an" : "a"} ${kind}${what}${by}.`;
  const one = shown.length === 1 && !["scatter", "boxWhisker", "waterfall"].includes(plan.spec.kind) ? shown[0] : null;
  if (one && plan.categories && plan.categories.length > 1) {
    let best = -1;
    one.values.forEach((v, i) => { if (typeof v === "number" && (best < 0 || v > one.values[best])) best = i; });
    if (best >= 0) {
      const unit = autoUnit(maxAbsOf(one.values));
      text += ` Highest: ${axisLabel(plan.categories[best])}, ${displayValue(one.values[best], one.fmt, unit, unit > 1 ? 1 : null)}.`;
    }
  }
  return text;
}
const altAttr = (plan) => ` descr="${xmlEsc(chartAltText(plan))}"`;

/* A chart frame that shows Excel's own chart (a funnel: cx2; a waterfall: cx1) where Excel has
   it, else the classic chart drawn the same way */
const CHARTEX_REQUIRES = { cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex", cx2: "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" };
function chartAltAnchorXml(ch, exRid, rid, id, requires = "cx2") {
  const cx = Math.round(ch.wPx * CHART_PX_TO_EMU), cy = Math.round(ch.hPx * CHART_PX_TO_EMU);
  const frame = (uri, inner) => `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"${altAttr(ch.plan)}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></xdr:graphicFrame>`;
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:${requires}="${CHARTEX_REQUIRES[requires]}" Requires="${requires}">` +
    frame(CHARTEX_NS, `<cx:chart xmlns:cx="${CHARTEX_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${exRid}"/>`) +
    `</mc:Choice><mc:Fallback>` +
    frame("http://schemas.openxmlformats.org/drawingml/2006/chart", `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>`) +
    `</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor>`;
}


const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

function addRel(rels, type, target) {
  const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = "rId" + ((ids.length ? Math.max(...ids) : 0) + 1);
  const rel = `<Relationship Id="${id}" Type="${/^https?:/.test(type) ? type : REL_TYPE + type}" Target="${target}"/>`;
  const xml = /<Relationships[^>]*\/>/.test(rels)
    ? rels.replace(/<Relationships([^>]*)\/>/, `<Relationships$1>${rel}</Relationships>`)
    : rels.replace("</Relationships>", rel + "</Relationships>");
  return { xml, id };
}

function relTarget(rels, id) {
  const tag = (rels.match(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*>`)) || [])[0] || "";
  return (tag.match(/\bTarget="([^"]+)"/) || [])[1] || null;
}

function unescapeXml(s) {
  return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

/* "../drawings/drawing1.xml" relative to a file in xl/<folder>/ → "xl/drawings/drawing1.xml" */
function xlPath(target) {
  if (target.startsWith("/")) return target.slice(1);
  return "xl/" + target.replace(/^\.\.\//, "").replace(/^\.\//, "");
}

function insertDrawing(sheetXml, rid) {
  if (!/<worksheet[^>]*xmlns:r=/.test(sheetXml)) {
    sheetXml = sheetXml.replace("<worksheet", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  }
  // Look only past the parts that come before <drawing>: data bars carry their own <extLst>
  // inside <conditionalFormatting>, and a chart placed there is ignored by Excel
  const from = Math.max(0, ...["</sheetData>", "</conditionalFormatting>", "</dataValidations>", "</hyperlinks>", "</ignoredErrors>"]
    .map((t) => sheetXml.lastIndexOf(t)));
  // <drawing> must come before these elements in a worksheet
  const after = ["<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst"];
  const found = after.map((t) => sheetXml.indexOf(t, from)).filter((i) => i >= 0);
  const at = found.length ? Math.min(...found) : sheetXml.lastIndexOf("</worksheet>");
  return sheetXml.slice(0, at) + `<drawing r:id="${rid}"/>` + sheetXml.slice(at);
}

/* Sheet name → its part in the file ("xl/worksheets/sheet3.xml") */
async function sheetFileMap(zip) {
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const workbookRels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const files = {};
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unescapeXml((m[0].match(/\bname="([^"]*)"/) || [])[1] || "");
    const rid = (m[0].match(/\br:id="([^"]+)"/) || [])[1];
    const target = rid && relTarget(workbookRels, rid);
    if (target) files[name] = xlPath(target);
  }
  return files;
}

/*
 * Excel slicers ("Include: Excel slicers"): each of the page's slicer fields becomes an Excel
 * slicer on every exported table with that column, to the right of the page (outside the print
 * area). A slicer filters one table; the totals rows use SUBTOTAL, so they follow it. The parts,
 * as Excel writes them: a slicer cache per slicer (workbook level, with a #N/A defined name),
 * the sheet's slicers part, and a frame in the sheet's drawing.
 * slicerSheets: [{ sheetName, slicers: [{ tableName, column, caption, title, row, col, count }] }]
 */
const SLICER_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SLICER_W_PX = 190, SLICER_ROW_PX = 26;
async function addSlicersToFile(buffer, slicerSheets) {
  const zip = await JSZip.loadAsync(buffer);
  const read = (p) => zip.file(p).async("string");
  let types = await read("[Content_Types].xml");
  const addType = (part, type) => {
    if (!types.includes(`PartName="/${part}"`)) types = types.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  let workbookXml = await read("xl/workbook.xml");
  let workbookRels = await read("xl/_rels/workbook.xml.rels");
  // The tables ExcelJS wrote: name → id and column names
  const tables = {};
  for (const path of Object.keys(zip.files).filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))) {
    const xml = await read(path);
    const tag = (xml.match(/<table\b[^>]*>/) || [""])[0];
    const name = unescapeXml((tag.match(/\bname="([^"]*)"/) || [])[1] || "");
    const id = (tag.match(/\bid="(\d+)"/) || [])[1];
    if (name && id) tables[name] = { id, cols: [...xml.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map((m) => unescapeXml(m[1])) };
  }
  const sheetFiles = await sheetFileMap(zip);
  let drawingNo = Object.keys(zip.files).filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).length;
  let cacheNo = 0, partNo = 0, frameId = 5000;
  const cacheNames = new Set([...workbookXml.matchAll(/<definedName\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1].toLowerCase()));
  const slicerNames = new Set();
  const unique = (base, used) => {
    let name = base, k = 1;
    while (used.has(name.toLowerCase())) name = `${base}${++k}`;
    used.add(name.toLowerCase());
    return name;
  };
  const cacheRids = [], names = [];
  for (const { sheetName, slicers } of slicerSheets) {
    const sheetPath = sheetFiles[sheetName];
    const list = slicers.filter((x) => tables[x.tableName] && tables[x.tableName].cols.includes(x.column));
    if (!sheetPath || !list.length) continue;
    const sheetRelsPath = sheetPath.replace(/([^/]+)$/, "_rels/$1.rels");
    let sheetXml = await read(sheetPath);
    let sheetRels = zip.file(sheetRelsPath) ? await read(sheetRelsPath) : EMPTY_RELS;
    // The sheet's drawing (the logo's, the charts'), or a new one
    let drawingPath, drawingXml;
    const existing = sheetXml.match(/<drawing r:id="([^"]+)"\s*\/>/);
    const existingTarget = existing && relTarget(sheetRels, existing[1]);
    if (existingTarget) {
      drawingPath = "xl/drawings/" + existingTarget.split("/").pop();
      drawingXml = await read(drawingPath);
    } else {
      drawingNo++;
      drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>`;
      const rel = addRel(sheetRels, "drawing", `../drawings/drawing${drawingNo}.xml`);
      sheetRels = rel.xml;
      sheetXml = insertDrawing(sheetXml, rel.id);
      addType(drawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml");
    }

    // Stacked down the right of the page, each level with its table where there is room
    let anchors = "", entries = "", nextRow = 0;
    for (const x of list) {
      const t = tables[x.tableName];
      cacheNo++;
      const cacheName = unique(`Slicer_${x.column.replace(/[^A-Za-z0-9_]/g, "_").replace(/^(\d)/, "_$1")}`, cacheNames);
      const slicerName = unique(x.column, slicerNames);
      const cachePath = `xl/slicerCaches/slicerCache${cacheNo}.xml`;
      zip.file(cachePath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<slicerCacheDefinition xmlns="${SLICER_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" ` +
        `xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${xmlEsc(cacheName)}" sourceName="${xmlEsc(x.column)}">` +
        `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
        `<x15:tableSlicerCache tableId="${t.id}" column="${t.cols.indexOf(x.column) + 1}"/></x:ext></extLst></slicerCacheDefinition>`);
      addType(cachePath, "application/vnd.ms-excel.slicerCache+xml");
      const wRel = addRel(workbookRels, "http://schemas.microsoft.com/office/2007/relationships/slicerCache", `slicerCaches/slicerCache${cacheNo}.xml`);
      workbookRels = wRel.xml;
      cacheRids.push(wRel.id);
      names.push(cacheName);
      // Captioned by its field (it sits beside its own table), in Excel's neutral grey style
      entries += `<slicer name="${xmlEsc(slicerName)}" cache="${xmlEsc(cacheName)}" caption="${xmlEsc(x.caption)}" style="SlicerStyleOther1" rowHeight="${SLICER_ROW_PX * CHART_PX_TO_EMU}"/>`;
      // Its frame: as tall as its items (up to 8; more scroll), beside its table
      const hPx = 34 + Math.min(8, Math.max(1, x.count)) * (SLICER_ROW_PX + 3) + 8;
      const row = Math.max(x.row - 1, nextRow);
      nextRow = row + Math.ceil(hPx / ROW_PX) + 1;
      frameId++;
      anchors += `<xdr:oneCellAnchor><xdr:from><xdr:col>${x.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:ext cx="${SLICER_W_PX * CHART_PX_TO_EMU}" cy="${Math.round(hPx * CHART_PX_TO_EMU)}"/>` +
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${frameId}" name="${xmlEsc(slicerName)}" descr="${xmlEsc(`Slicer for ${x.column} on ${x.title || x.tableName}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">` +
        `<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${xmlEsc(slicerName)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
        `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${frameId}" name="${xmlEsc(slicerName)}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
        `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SLICER_W_PX * CHART_PX_TO_EMU}" cy="${Math.round(hPx * CHART_PX_TO_EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:solidFill><a:srgbClr val="F4F6F5"/></a:solidFill></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>` +
        `<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>${xmlEsc(`Slicer for ${x.column}: shown in Excel 2013 or later.`)}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>` +
        `</mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor>`;
    }
    // The sheet's slicers part, listed in its extLst (always the worksheet's last element)
    partNo++;
    const partPath = `xl/slicers/slicer${partNo}.xml`;
    zip.file(partPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicers xmlns="${SLICER_NS}" ` +
      `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${entries}</slicers>`);
    addType(partPath, "application/vnd.ms-excel.slicer+xml");
    const sRel = addRel(sheetRels, "http://schemas.microsoft.com/office/2007/relationships/slicer", `../slicers/slicer${partNo}.xml`);
    sheetRels = sRel.xml;
    if (!/<worksheet[^>]*xmlns:r=/.test(sheetXml)) {
      sheetXml = sheetXml.replace("<worksheet", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }
    const ext = `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
      `<x14:slicerList xmlns:x14="${SLICER_NS}"><x14:slicer r:id="${sRel.id}"/></x14:slicerList></ext>`;
    sheetXml = /<\/extLst>\s*<\/worksheet>\s*$/.test(sheetXml)
      ? sheetXml.replace(/<\/extLst>\s*<\/worksheet>\s*$/, `${ext}</extLst></worksheet>`)
      : sheetXml.replace(/<\/worksheet>\s*$/, `<extLst>${ext}</extLst></worksheet>`);
    drawingXml = drawingXml.includes("</xdr:wsDr>")
      ? drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>")
      : drawingXml.replace(/<xdr:wsDr([^>]*)\/>/, `<xdr:wsDr$1>${anchors}</xdr:wsDr>`);
    zip.file(sheetPath, sheetXml);
    zip.file(sheetRelsPath, sheetRels);
    zip.file(drawingPath, drawingXml);
  }
  if (!cacheRids.length) return buffer;
  // The workbook: a #N/A name per slicer cache, and the caches listed in its extLst
  const defined = names.map((n) => `<definedName name="${xmlEsc(n)}">#N/A</definedName>`).join("");
  workbookXml = workbookXml.includes("</definedNames>")
    ? workbookXml.replace("</definedNames>", `${defined}</definedNames>`)
    : workbookXml.replace(/(<\/sheets>)/, `$1<definedNames>${defined}</definedNames>`);
  if (!/<workbook[^>]*xmlns:r=/.test(workbookXml)) {
    workbookXml = workbookXml.replace("<workbook", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  }
  const wbExt = `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
    `<x15:slicerCaches xmlns:x14="${SLICER_NS}">${cacheRids.map((id) => `<x14:slicerCache r:id="${id}"/>`).join("")}</x15:slicerCaches></ext>`;
  workbookXml = workbookXml.includes("</extLst>")
    ? workbookXml.replace(/<\/extLst>\s*<\/workbook>\s*$/, `${wbExt}</extLst></workbook>`)
    : workbookXml.replace(/<\/workbook>\s*$/, `<extLst>${wbExt}</extLst></workbook>`);
  zip.file("xl/workbook.xml", workbookXml);
  zip.file("xl/_rels/workbook.xml.rels", workbookRels);
  zip.file("[Content_Types].xml", types);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/* chartSheets: [{ sheetName, charts: [{ plan, col, row, wPx, hPx }] }]. Returns the new file. */
async function addChartsToFile(buffer, chartSheets) {
  const zip = await JSZip.loadAsync(buffer);
  const read = (p) => zip.file(p).async("string");
  let types = await read("[Content_Types].xml");
  const addType = (part, type) => {
    if (!types.includes(`PartName="/${part}"`)) types = types.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  let chartNo = Object.keys(zip.files).filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)).length;
  let exNo = Object.keys(zip.files).filter((n) => /^xl\/charts\/chartEx\d+\.xml$/.test(n)).length;
  let drawingNo = Object.keys(zip.files).filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).length;
  const sheetFiles = await sheetFileMap(zip);

  for (const { sheetName, charts } of chartSheets) {
    const sheetPath = sheetFiles[sheetName];
    if (!sheetPath || !charts.length) continue;
    const sheetRelsPath = sheetPath.replace(/([^/]+)$/, "_rels/$1.rels");
    let sheetXml = await read(sheetPath);
    let sheetRels = zip.file(sheetRelsPath) ? await read(sheetRelsPath) : EMPTY_RELS;

    // The sheet's drawing: the one ExcelJS made for the logo, or a new one
    let drawingPath, drawingXml;
    const existing = sheetXml.match(/<drawing r:id="([^"]+)"\s*\/>/);
    const existingTarget = existing && relTarget(sheetRels, existing[1]);
    if (existingTarget) {
      drawingPath = "xl/drawings/" + existingTarget.split("/").pop();
      drawingXml = await read(drawingPath);
    } else {
      drawingNo++;
      drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>`;
      const rel = addRel(sheetRels, "drawing", `../drawings/drawing${drawingNo}.xml`);
      sheetRels = rel.xml;
      sheetXml = insertDrawing(sheetXml, rel.id);
      addType(drawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml");
    }
    const drawingRelsPath = drawingPath.replace(/([^/]+)$/, "_rels/$1.rels");
    let drawingRels = zip.file(drawingRelsPath) ? await read(drawingRelsPath) : EMPTY_RELS;

    let anchors = "";
    // A chartEx chart (treemap, sunburst, funnel): its part, with Excel's style and colour
    // parts (Excel won't open the file without both); returns the drawing's link to it
    const addChartEx = (xml, colorsXml = CHARTEX_COLORS_XML, styleXml = CHARTEX_STYLE_XML) => {
      const exPath = `xl/charts/chartEx${exNo}.xml`;
      zip.file(exPath, xml);
      zip.file(`xl/charts/style${exNo}.xml`, styleXml);
      zip.file(`xl/charts/colors${exNo}.xml`, colorsXml);
      let exRels = addRel(EMPTY_RELS, "http://schemas.microsoft.com/office/2011/relationships/chartStyle", `style${exNo}.xml`).xml;
      exRels = addRel(exRels, "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle", `colors${exNo}.xml`).xml;
      zip.file(`xl/charts/_rels/chartEx${exNo}.xml.rels`, exRels);
      addType(exPath, CHARTEX_TYPE);
      addType(`xl/charts/style${exNo}.xml`, "application/vnd.ms-office.chartstyle+xml");
      addType(`xl/charts/colors${exNo}.xml`, "application/vnd.ms-office.chartcolorstyle+xml");
      const rel = addRel(drawingRels, CHARTEX_REL, `../charts/chartEx${exNo}.xml`);
      drawingRels = rel.xml;
      return rel.id;
    };
    for (const ch of charts) {
      if (ch.plan.panels) {
        // Small multiples: each panel a chart part, the panels and their title one grouped object
        const rids = ch.plan.panels.plans.map((p) => {
          chartNo++;
          const chartPath = `xl/charts/chart${chartNo}.xml`;
          zip.file(chartPath, chartXml(p).xml);
          addType(chartPath, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
          const rel = addRel(drawingRels, "chart", `../charts/chart${chartNo}.xml`);
          drawingRels = rel.xml;
          return rel.id;
        });
        anchors += panelGroupXml(ch, rids, 1000 + chartNo);
        continue;
      }
      if (ch.plan.spec.kind === "treemap" || ch.plan.spec.kind === "sunburst") {
        exNo++;
        anchors += chartExAnchorXml(ch, addChartEx(treemapXml(ch.plan, exNo)), 3000 + exNo);
        continue;
      }
      if (ch.plan.spec.kind === "boxWhisker") {
        exNo++;
        anchors += chartExAnchorXml(ch, addChartEx(boxWhiskerXml(ch.plan), CHARTEX_COLORS_XML, CHARTEX_STYLE_DARK_LABELS), 3000 + exNo);
        continue;
      }
      chartNo++;
      const chartPath = `xl/charts/chart${chartNo}.xml`;
      const part = chartXml(ch.plan);
      zip.file(chartPath, part.xml);
      addType(chartPath, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
      if (part.shapes) {
        // Shapes over the chart (a legend's title, a donut's hole): the chart's own drawing
        drawingNo++;
        const shapesPath = `xl/drawings/drawing${drawingNo}.xml`;
        zip.file(shapesPath, part.shapes);
        addType(shapesPath, "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml");
        zip.file(`xl/charts/_rels/chart${chartNo}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${USER_SHAPES_RID}" ` +
          `Type="${REL_TYPE}chartUserShapes" Target="../drawings/drawing${drawingNo}.xml"/></Relationships>`);
      }
      const rel = addRel(drawingRels, "chart", `../charts/chart${chartNo}.xml`);
      drawingRels = rel.xml;
            // A funnel: Excel's own funnel chart too, with the bar funnel as its fallback
      if (NATIVE_FUNNEL && ch.plan.funnel) {
        exNo++;
        anchors += chartAltAnchorXml(ch, addChartEx(funnelChartExXml(ch.plan)), rel.id, 1000 + chartNo);
      } else if (NATIVE_WATERFALL && ch.plan.spec.kind === "waterfall" && ch.plan.series[5]) {
        // A waterfall: Excel's own, in the report's increase, decrease and total colours
        exNo++;
        const p = ch.plan;
        anchors += chartAltAnchorXml(ch, addChartEx(waterfallChartExXml(p), chartExColorsXml([p.series[1].color, p.series[2].color, p.series[3].color]), CHARTEX_STYLE_DARK_LABELS), rel.id, 1000 + chartNo, "cx1");
      } else {
        anchors += chartAnchorXml(ch, rel.id, 1000 + chartNo);
      }
    }
    drawingXml = drawingXml.includes("</xdr:wsDr>")
      ? drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>")
      : drawingXml.replace(/<xdr:wsDr([^>]*)\/>/, `<xdr:wsDr$1>${anchors}</xdr:wsDr>`);

    zip.file(sheetPath, sheetXml);
    zip.file(sheetRelsPath, sheetRels);
    zip.file(drawingPath, drawingXml);
    zip.file(drawingRelsPath, drawingRels);
  }
  zip.file("[Content_Types].xml", types);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/*
 * Trend cards' sparklines. ExcelJS can't write them, so they go into each sheet's <extLst>
 * once the file is written: a group per sparkline (each has its own colour), a line over a
 * date axis (points spaced by date, as the visual's time scale spaces them), the last point
 * marked, gaps for blanks, and drawn from collapsed rows too.
 * sparkSheets: [{ sheetName, sparklines: [{ cell, values, dates, color, min, max }] }].
 */
const SPARKLINE_EXT = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
function sparklineGroupXml(s) {
  const color = (tag) => `<x14:${tag} rgb="${s.color}"/>`;
  const axis = (s.min !== null && s.min !== undefined ? ` manualMin="${s.min}"` : "") + (s.max !== null && s.max !== undefined ? ` manualMax="${s.max}"` : "");
  const types = (s.min !== null && s.min !== undefined ? ' minAxisType="custom"' : "") + (s.max !== null && s.max !== undefined ? ' maxAxisType="custom"' : "");
  return `<x14:sparklineGroup${axis} lineWeight="1.5"${s.dates ? ' dateAxis="1"' : ""} displayEmptyCellsAs="gap" last="1" displayHidden="1"${types}>` +
    color("colorSeries") + color("colorNegative") + `<x14:colorAxis rgb="FF000000"/>` + color("colorMarkers") +
    color("colorFirst") + color("colorLast") + color("colorHigh") + color("colorLow") +
    (s.dates ? `<xm:f>${xmlEsc(s.dates)}</xm:f>` : "") +
    `<x14:sparklines><x14:sparkline><xm:f>${xmlEsc(s.values)}</xm:f><xm:sqref>${s.cell}</xm:sqref></x14:sparkline></x14:sparklines>` +
    `</x14:sparklineGroup>`;
}

async function addSparklinesToFile(buffer, sparkSheets) {
  const zip = await JSZip.loadAsync(buffer);
  const sheetFiles = await sheetFileMap(zip);
  for (const { sheetName, sparklines } of sparkSheets) {
    const path = sheetFiles[sheetName];
    if (!path || !sparklines.length) continue;
    const xml = await zip.file(path).async("string");
    const ext = `<ext uri="${SPARKLINE_EXT}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
      `<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${sparklines.map(sparklineGroupXml).join("")}</x14:sparklineGroups></ext>`;
    // <extLst> is a worksheet's last element: join the one ExcelJS wrote (data bars), else add one
    const body = xml.slice(0, xml.lastIndexOf("</worksheet>")).trimEnd();
    const out = body.endsWith("</extLst>")
      ? body.slice(0, -"</extLst>".length) + ext + "</extLst></worksheet>"
      : body + "<extLst>" + ext + "</extLst></worksheet>";
    zip.file(path, out);
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/* ── Table sparklines ──────────────────────────────────────────────────────
 * A table's sparkline column (a measure over a field, e.g. Sales by Month) isn't in Power
 * BI's exported data. Its points are read with one DAX query per sparkline through the
 * Power BI executeQueries API, grouped by the table's row fields and the sparkline's field,
 * under the same report, page, slicer and visual filters. This needs the report file (to
 * know the fields) and the user's permission to query the semantic model (Build); otherwise
 * a note says why the trend lines are missing.
 */

/* A DAX literal for a filter value */
function daxLiteral(v) {
  if (v === null || v === undefined) return "BLANK()";
  if (typeof v === "boolean") return v ? "TRUE()" : "FALSE()";
  if (typeof v === "number") return String(v);
  const s = String(v);
  const d = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}))?/);
  if (d) return d[4] && (+d[4] || +d[5] || +d[6]) ? `DATE(${+d[1]},${+d[2]},${+d[3]})+TIME(${+d[4]},${+d[5]},${+d[6]})` : `DATE(${+d[1]},${+d[2]},${+d[3]})`;
  return `"${s.replace(/"/g, '""')}"`;
}

const daxColumn = (table, column) => `'${String(table).replace(/'/g, "''")}'[${String(column).replace(/]/g, "]]")}]`;

/* A filter model (report, page, slicer or visual filter) as a DAX filter table; throws on a
   kind of filter that can't be written (relative dates, Top N, hierarchy filters) */
const DAX_CONDITIONS = {
  Is: (c, v) => `${c} = ${v}`, IsNot: (c, v) => `${c} <> ${v}`, LessThan: (c, v) => `${c} < ${v}`, LessThanOrEqual: (c, v) => `${c} <= ${v}`,
  GreaterThan: (c, v) => `${c} > ${v}`, GreaterThanOrEqual: (c, v) => `${c} >= ${v}`, Contains: (c, v) => `CONTAINSSTRING(${c}, ${v})`,
  DoesNotContain: (c, v) => `NOT CONTAINSSTRING(${c}, ${v})`, StartsWith: (c, v) => `LEFT(${c}, LEN(${v})) = ${v}`,
  DoesNotStartWith: (c, v) => `LEFT(${c}, LEN(${v})) <> ${v}`, IsBlank: (c) => `ISBLANK(${c})`, IsNotBlank: (c) => `NOT ISBLANK(${c})`,
};
function daxFilter(f) {
  const t = f && f.target;
  if (!t || !t.table || !t.column || f.timeUnitType !== undefined || f.itemCount !== undefined || f.hierarchyData) {
    throw new Error(`filter on ${targetName(t)} can't be applied to the query`);
  }
  const c = daxColumn(t.table, t.column);
  if (f.values) {
    if (!f.values.length || f.operator === "All") return null;
    if (f.values.some(Array.isArray)) throw new Error(`filter on ${targetName(t)} can't be applied to the query`);
    const list = `{${f.values.map(daxLiteral).join(", ")}}`;
    return f.operator === "NotIn" || f.isExclude ? `FILTER(ALL(${c}), NOT ${c} IN ${list})` : `TREATAS(${list}, ${c})`;
  }
  if (f.conditions) {
    if (!f.conditions.length) return null;
    const parts = f.conditions.map((x) => {
      const fn = DAX_CONDITIONS[x.operator];
      if (!fn) throw new Error(`filter "${x.operator}" on ${t.column} can't be applied to the query`);
      return `(${fn(c, daxLiteral(x.value))})`;
    });
    return `FILTER(ALL(${c}), ${parts.join(f.logicalOperator === "Or" ? " || " : " && ")})`;
  }
  return null;
}

/* The DAX for a field's value: a measure, or a column's aggregation */
const DAX_AGGREGATIONS = { 0: "SUM", 1: "AVERAGE", 2: "DISTINCTCOUNT", 3: "MIN", 4: "MAX", 5: "COUNTA", 6: "MEDIAN", 7: "STDEV.P", 8: "VAR.P" };
function daxValue(f) {
  if (!f || !f.entity || !f.prop) return null;
  if (f.kind === "measure") return daxColumn(f.entity, f.prop);
  if (f.kind === "agg" && DAX_AGGREGATIONS[f.agg]) return `${DAX_AGGREGATIONS[f.agg]}(${daxColumn(f.entity, f.prop)})`;
  return null;
}

/* Text that matches a value however it was written: "2023-01-15 00:00:00", "1/15/2023", a JSON date */
function matchKey(v) {
  if (v === null || v === undefined) return "";
  const s = String(v).trim();
  const d = parseDate(s) || parseDate(s.replace(/T.*$/, ""));
  if (d) return d.date.toISOString().slice(0, 10);
  const n = toCellValue(s).value;
  return typeof n === "number" ? String(n) : s.toLowerCase();
}

/* Sparkline points in the order Power BI draws them: dates and numbers ascending, month
   and quarter names in calendar order, other text as the query returned it */
const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
function axisOrder(values) {
  const key = (v) => {
    if (v === null || v === undefined) return [3, ""];
    const s = String(v).trim();
    const d = parseDate(s.replace(/T(\d\d:\d\d:\d\d).*$/, " $1"));
    if (d) return [0, d.date.getTime()];
    if (typeof v === "number" || /^-?\d+(\.\d+)?$/.test(s)) return [0, Number(s)];
    const m = MONTHS.indexOf(s.slice(0, 3).toLowerCase());
    if (m >= 0) return [1, m];
    const q = s.match(/^(?:q|qtr|quarter)\s*(\d)$/i);
    if (q) return [1, Number(q[1])];
    return [2, 0];
  };
  return (a, b) => {
    const [ka, va] = key(a), [kb, vb] = key(b);
    return ka !== kb ? ka - kb : va < vb ? -1 : va > vb ? 1 : 0;
  };
}

/* The table's sparkline fields, with the table's grouping fields, from the report file */
function sparklineFields(item, info) {
  if (!info || !["tableEx", "pivotTable"].includes(item.type)) return null;
  const roles = info.roles || {};
  if (item.type === "pivotTable" && (roles.Columns || []).length) return null;   // a pivoted matrix: not supported
  const groupRole = item.type === "pivotTable" ? roles.Rows || [] : (roles.Values || []).filter((e) => e.field && e.field.kind === "column");
  const values = roles.Values || [];
  const sparks = values.filter((e) => e.field && e.field.kind === "spark");
  if (!sparks.length) return null;
  return { groups: groupRole, sparks, values };
}

async function querySparklines(item, info, filters) {
  const plan = sparklineFields(item, info);
  if (!plan) return null;
  const [header] = item.rows;
  if (plan.groups.some((g) => !g.field || g.field.kind !== "column" || !g.field.entity)) {
    return { note: `"${item.title}": its trend lines (sparklines) run over a hierarchy level, which the export can't query.` };
  }
  const found = fieldsByColumn(header, plan.groups);
  const where = plan.groups.map((g) => found.indexOf(g));
  if (where.some((ci) => ci < 0)) return null;
  let filterTables;
  try {
    filterTables = filters.map(daxFilter).filter(Boolean);
  } catch (e) {
    return { note: `"${item.title}": its trend lines (sparklines) weren't added: a ${e.message}.` };
  }
  const groupBy = plan.groups.map((g) => daxColumn(g.field.entity, g.field.prop));
  const out = [];
  for (const sp of plan.sparks) {
    const s = sp.field;
    const axis = (s.groupings || [])[0];
    const value = daxValue(s.measure);
    if (!value || !axis || axis.kind !== "column" || !axis.entity) {
      out.push({ entry: sp, error: "its field isn't a column the export can query" });
      continue;
    }
    const query = `EVALUATE SUMMARIZECOLUMNS(${[...groupBy, daxColumn(axis.entity, axis.prop), ...filterTables, `"v", ${value}`].join(", ")})`;
    const res = await executeDaxQuery(query);
    if (res.error) { out.push({ entry: sp, error: res.error }); continue; }
    // Rows come back as { "Table[Column]": …, "[v]": … }
    const keyName = (g) => `${g.field.entity}[${g.field.prop}]`;
    const axisName = `${axis.entity}[${axis.prop}]`;
    const byRow = new Map();
    const points = new Set();
    for (const row of res.rows) {
      const k = plan.groups.map((g) => matchKey(row[keyName(g)])).join("\u0001");
      if (!byRow.has(k)) byRow.set(k, new Map());
      byRow.get(k).set(row[axisName], typeof row["[v]"] === "number" ? row["[v]"] : null);
      points.add(row[axisName]);
    }
    const axisValues = [...points].sort(axisOrder([...points]));
    const name = `${s.measure.prop} by ${axis.prop}`;
    const series = item.rows.slice(1).map((r) => {
      const m = byRow.get(where.map((ci) => matchKey(r[ci])).join("\u0001"));
      return axisValues.map((a) => (m && m.has(a) ? m.get(a) : null));
    });
    out.push({ entry: sp, index: plan.values.indexOf(sp), keyCols: where, name, axis: axisValues.map((a) => (typeof a === "string" && /T\d\d:/.test(a) ? a.slice(0, 10) : a)), series });
  }
  return { sparks: out };
}

/* Runs a DAX query on the report's semantic model; { rows } or { error } */
async function executeDaxQuery(query) {
  try {
    const token = await getDesignToken({ quiet: true });
    if (!token) return { error: "the export doesn't have permission to query the data (Export options → Colours → Allow exact colours)" };
    let datasetId = currentReport.datasetId;
    const groupId = new URL(currentReport.embedUrl).searchParams.get("groupId");
    if (!datasetId) {
      const r = await fetch(`${PBI_API}${groupId ? `/groups/${groupId}` : ""}/reports/${currentReport.id}`, { headers: { Authorization: `Bearer ${token}` } });
      datasetId = r.ok ? (await r.json()).datasetId : null;
      currentReport.datasetId = datasetId;
    }
    if (!datasetId) return { error: "the report's semantic model couldn't be found" };
    const res = await fetch(`${PBI_API}${groupId ? `/groups/${groupId}` : ""}/datasets/${datasetId}/executeQueries`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ queries: [{ query }], serializerSettings: { includeNulls: true } }),
    });
    if (res.status === 401 || res.status === 403) return { error: "querying the semantic model needs Build permission on it" };
    const body = await res.json().catch(() => ({}));
    if (!res.ok) return { error: `the query failed (${res.status}${body.error ? ": " + (body.error.message || body.error.code) : ""})` };
    const result = body.results && body.results[0];
    if (result && result.error) return { error: `the query failed (${result.error.message || result.error.code})` };
    return { rows: ((result && result.tables && result.tables[0]) || {}).rows || [] };
  } catch (e) {
    return { error: `the query failed (${errorText(e)})` };
  }
}

/*
 * Power BI's own query service, as its website asks for a chart's data: Viewers can use it
 * where the semantic model's query API needs Build permission. fields: [{ kind, entity, prop,
 * agg }], the groups first; filters: the report's and page's (embed API models). Returns rows
 * as arrays in the fields' order. Internal and undocumented, so its answers are only used
 * after a check against the exported rows (websiteLevelValues).
 */
async function websiteQuery(fields, filters) {
  let modelId = websiteModelIds[currentReport.id];
  if (modelId === undefined) { await fetchWebsiteLayout(currentReport); modelId = websiteModelIds[currentReport.id]; }
  if (modelId === undefined) throw new Error("the report's model wasn't found");
  const from = [], alias = {};
  const source = (entity) => {
    if (!(entity in alias)) { alias[entity] = `t${from.length}`; from.push({ Name: alias[entity], Entity: entity, Type: 0 }); }
    return { SourceRef: { Source: alias[entity] } };
  };
  const column = (entity, prop) => ({ Column: { Expression: source(entity), Property: prop } });
  const expr = (f) => (f.kind === "measure" ? { Measure: { Expression: source(f.entity), Property: f.prop } }
    : f.kind === "agg" ? { Aggregation: { Expression: column(f.entity, f.prop), Function: f.agg } } : column(f.entity, f.prop));
  const select = fields.map((f, i) => ({ ...expr(f), Name: `f${i}` }));
  const where = semanticWhere(filters, column);
  const body = { version: "1.0.0", cancelQueries: [], modelId, queries: [{ QueryId: "",
    ApplicationContext: { ...(currentReport.datasetId ? { DatasetId: currentReport.datasetId } : {}), Sources: [{ ReportId: currentReport.id }] },
    Query: { Commands: [{ SemanticQueryDataShapeCommand: {
      Query: { Version: 2, From: from, Select: select, ...(where.length ? { Where: where } : {}) },
      Binding: { Primary: { Groupings: [{ Projections: fields.map((_, i) => i) }] }, DataReduction: { DataVolume: 4, Primary: { Window: { Count: 30000 } } }, Version: 1 },
    } }] } }] };
  const res = await postAsUser(`${await clusterUrl(currentReport)}/explore/querydata?synchronous=true`, body);
  if (!res.ok) throw new Error(`Power BI's query service answered ${res.status}`);
  const result = ((((await res.json()).results || [])[0]) || {}).result || {};
  const data = result.data || {};
  const ds = ((data.dsr || {}).DS || [])[0];
  if (!ds) throw new Error(`Power BI's query service returned no data${result.error ? ` (${JSON.stringify(result.error).slice(0, 200)})` : ""}`);
  if (ds.IC === false) throw new Error("Power BI's query service returned part of the data");
  return dsrRows(ds, data.descriptor, fields.length);
}

/* A query service answer (DSR: each row lists only the values that changed, with bit masks for
   values repeated from the row before (R) and empty ones (Ø), text as indexes into value
   dictionaries) as rows of values in the query's order (selects named f0, f1…) */
function dsrRows(ds, descriptor, count) {
  const nameOf = Object.fromEntries(((descriptor || {}).Select || []).map((s) => [s.Value, s.Name]));
  const dicts = ds.ValueDicts || {};
  const rows = [];
  let schema = null, prev = [];
  for (const ph of ds.PH || []) {
    for (const r of ph.DM0 || []) {
      if (r.S) schema = r.S;
      // A total with nothing to group by ("Readmission Rate" for all rows) comes as the value
      // under its own name ({ "M0": 0.108 }), with or without a schema, rather than in C
      const named = Object.keys(r).filter((k) => /^[MG]\d+$/.test(k));
      if (!schema && named.length) schema = named.map((N) => ({ N }));
      if (!schema) continue;
      const c = [...(r.C || [])], repeat = r.R || 0, empty = r["Ø"] || 0;
      const vals = schema.map((col, i) => {
        if (repeat & (1 << i)) return prev[i];
        if (empty & (1 << i)) return null;
        if (col.N in r && !Array.isArray(r[col.N])) return r[col.N];
        const v = c.shift();
        return col.DN && typeof v === "number" && dicts[col.DN] ? dicts[col.DN][v] : v;
      });
      prev = vals;
      const out = new Array(count).fill(null);
      schema.forEach((col, i) => {
        const k = Number(String(nameOf[col.N] || "").replace(/^f/, ""));
        if (Number.isInteger(k) && k >= 0 && k < count) out[k] = vals[i];
      });
      rows.push(out);
    }
  }
  return rows;
}

/* The report's and page's filters as the query service's conditions. A filter on a measure
   only drops rows, which the check against the exported rows covers; one that can't be
   written stops the query. */
function semanticWhere(filters, column) {
  const literal = (v) => ({ Literal: { Value: typeof v === "string" ? `'${v.replace(/'/g, "''")}'` : typeof v === "boolean" ? String(v)
    : v === null || v === undefined ? "null" : Number.isInteger(v) ? `${v}L` : `${v}D` } });
  const out = [];
  for (const f of filters || []) {
    const t = f && f.target;
    if (t && t.measure) continue;
    if (!t || !t.table || !t.column || f.timeUnitType !== undefined || f.itemCount !== undefined || f.hierarchyData) {
      throw new Error(`filter on ${targetName(t)} can't be applied to the query`);
    }
    if (f.conditions) {
      // "≥ -19 and is not blank": each condition as Power BI writes it (a comparison, 0 =, 1 >,
      // 2 ≥, 3 ≤, 4 <; blank as null), joined with And / Or
      if (!f.conditions.length) continue;
      const col = () => column(t.table, t.column);
      const cmp = (kind, v) => ({ Comparison: { ComparisonKind: kind, Left: col(), Right: literal(v) } });
      const not = (e) => ({ Not: { Expression: e } });
      const text = (op, v) => ({ [op]: { Left: col(), Right: literal(v) } });
      const SEMANTIC = { Is: (v) => cmp(0, v), IsNot: (v) => not(cmp(0, v)), GreaterThan: (v) => cmp(1, v), GreaterThanOrEqual: (v) => cmp(2, v),
        LessThanOrEqual: (v) => cmp(3, v), LessThan: (v) => cmp(4, v), IsBlank: () => cmp(0, null), IsNotBlank: () => not(cmp(0, null)),
        Contains: (v) => text("Contains", v), DoesNotContain: (v) => not(text("Contains", v)), StartsWith: (v) => text("StartsWith", v), DoesNotStartWith: (v) => not(text("StartsWith", v)) };
      const parts = f.conditions.map((x) => {
        const fn = SEMANTIC[x.operator];
        if (!fn) throw new Error(`filter "${x.operator}" on ${t.column} can't be applied to the query`);
        return fn(x.value);
      });
      const joined = parts.reduce((a, b) => ({ [f.logicalOperator === "Or" ? "Or" : "And"]: { Left: a, Right: b } }));
      out.push({ Condition: joined });
      continue;
    }
    if (!f.values || !f.values.length || f.operator === "All") continue;
    if (f.values.some(Array.isArray)) throw new Error(`filter on ${targetName(t)} can't be applied to the query`);
    const inList = { In: { Expressions: [column(t.table, t.column)], Values: f.values.map((v) => [literal(v)]) } };
    out.push({ Condition: f.operator === "NotIn" || f.isExclude ? { Not: { Expression: inList } } : inList });
  }
  return out;
}

/*
 * Values that don't add up, at the level a chart shows, from Power BI's query service (for
 * Viewers). First the same values at the exported rows' own level must match every exported
 * row exactly (same rows, same values to the digits exported), so filters, slicers and the
 * fields are known to be the chart's; else nothing is used. Returns { byKey: Map(group key →
 * values) } or { error }.
 */
async function websiteLevelValues(header, entries, keyCols, detailCols, open, data, cells, filters) {
  if (typeof websiteModelIds === "undefined" || typeof postAsUser !== "function") return { error: "not available" };
  const fieldOf = (ci) => entries[ci] && entries[ci].field;
  if (detailCols.some((ci) => { const f = fieldOf(ci); return !f || !f.entity || !["column", "level"].includes(f.kind); })) return { error: "its levels can't be queried" };
  const values = open.map(fieldOf);
  const keyOf = (row, cols) => cols.map((ci) => matchKey(row[ci])).join("\u0001");
  const at = async (cols) => (await websiteQuery([...cols.map(fieldOf), ...values], filters))
    .map((r) => ({ key: cols.map((_, i) => matchKey(r[i])).join("\u0001"), values: values.map((_, i) => r[cols.length + i]) }));
  // A value as exported (rounded to the digits shown) against the query's
  const near = (raw, v, q) => {
    if (q === null || q === undefined || q === "") return typeof v !== "number";
    const n = Number(q), s = String(raw ?? "");
    if (typeof v !== "number" || !Number.isFinite(n)) return false;
    const digits = ((s.match(/\.(\d+)/) || [])[1] || "").length;
    const unit = Math.pow(10, -digits) * (s.includes("%") ? 0.01 : 1);
    return Math.abs(n - v) <= unit * 0.51 + 1e-9 * Math.max(1, Math.abs(n));
  };
  try {
    const detail = await at(detailCols);
    const exported = new Map(data.map((row, ri) => [keyOf(row, detailCols), ri]));
    const seen = new Set();
    for (const r of detail) {
      if (r.values.every((v) => v === null || v === undefined || v === "")) continue;
      const ri = exported.get(r.key);
      if (ri === undefined) return { error: `the query has a row the export doesn't (${r.key.replace(/\u0001/g, " / ")})` };
      if (!open.every((ci, i) => near(data[ri][ci], cells[ri][ci].value, r.values[i]))) return { error: `its values differ from the export (${r.key.replace(/\u0001/g, " / ")})` };
      seen.add(ri);
    }
    const missing = data.findIndex((row, ri) => !seen.has(ri) && open.some((ci) => typeof cells[ri][ci].value === "number"));
    if (missing >= 0) return { error: `the export has a row the query doesn't (${keyOf(data[missing], detailCols).replace(/\u0001/g, " / ")})` };
    const level = await at(keyCols);
    // The service leaves out groups whose values are all empty (a VP with no separations last
    // year has no YoY change): with the rows checked above, a group missing is an empty value
    return { byKey: new Map(level.map((r) => [r.key, r.values])), missingIsEmpty: true };
  } catch (e) {
    return { error: errorText(e) };
  }
}

/* For every table with sparklines: their points (item.sparklines), or a note */
async function loadSparklines(collected, options) {
  for (const { page, data } of collected) {
    for (const item of data.items) {
      const info = visualInfo(options.reportColors, item.page, item.name);
      if (!sparklineFields(item, info)) continue;
      const res = await querySparklines(item, info, [...(data.queryFilters || []), ...(item.visualFilters || [])]);
      if (!res) continue;
      if (res.note) { data.notes.push(res.note); continue; }
      item.sparklines = res.sparks.filter((s) => !s.error);
      for (const s of res.sparks.filter((x) => x.error)) data.notes.push(`"${item.title}": a trend line (sparkline) wasn't added: ${s.error}.`);
      log(`Sparklines for "${item.title}" on ${page.displayName}: ${item.sparklines.length}`);
    }
  }
}

/*
 * Key influencers count what they analyse row by row (each opportunity), which the exported
 * rows can't give (Power BI merges identical rows). So the rows behind each explaining field
 * are counted with a query, per value and per value of the analysed field, under the same
 * report, page and visual filters: through the semantic model (Build permission), else Power
 * BI's own query service (Viewers). item.kiCounts: { fields: [{ ref, counts: [[value,
 * target, rows]] }] }, or item.kiError.
 */
async function loadKeyInfluencers(collected, options) {
  for (const { data } of collected) {
    for (const item of data.items) {
      if (item.type !== "keyDriversVisual") continue;
      const info = visualInfo(options.reportColors, item.page, item.name);
      const kd = info && info.keyDrivers;
      const target = kd && kd.target && kd.target.field;
      if (!target || target.kind !== "column" || !target.entity) { item.kiError = "its fields couldn't be read"; continue; }
      const filters = [...(data.queryFilters || []), ...(item.visualFilters || [])];
      const fields = [];
      for (const f of kd.explainBy) {
        if (!f.field || f.field.kind !== "column" || !f.field.entity) continue;   // number fields: explained in the report only
        const res = await keyInfluencerCounts(f.field, target, filters);
        if (res.error) { item.kiError = res.error; break; }
        fields.push({ ref: f.ref, label: f.label, prop: f.field.prop, counts: res.rows });
      }
      if (!item.kiError) item.kiCounts = { fields };
      log(`Key influencers "${item.title}": ${item.kiError ? item.kiError : `${fields.length} field(s) counted`}`);
    }
  }
}

/*
 * A decomposition tree whose measure doesn't add up from the exported rows (Var Plan, a
 * calculated measure): its value for the whole tree and for each level the report has open,
 * grouped by the levels above, asked of Power BI under the same report, page and visual
 * filters: through the semantic model (Build permission), else Power BI's own query service
 * (Viewers). item.treeValues: { total, byPath: { "europe": v, "europe\u00011": v } }; when
 * neither answers, item.treeError and the tree stays the table it was.
 */
async function loadTreeValues(collected, options) {
  for (const { data } of collected) {
    for (const item of data.items) {
      if (item.type !== "decompositionTreeVisual") continue;
      const info = visualInfo(options.reportColors, item.page, item.name);
      const measure = info && info.roles && (info.roles.Analyze || [])[0];
      const mf = measure && measure.field;
      if (!mf || (mf.kind === "agg" && [0, 5].includes(mf.agg))) continue;   // a sum or count adds up from the rows
      const { levelFields } = treeLevels(info);
      if (!levelFields.length || levelFields.some((f) => !f || !f.field || !["column", "level"].includes(f.field.kind) || !f.field.entity)) continue;
      const filters = [...(data.queryFilters || []), ...(item.visualFilters || [])];
      const byPath = {};
      let total = null;
      for (let d = 0; d <= levelFields.length && !item.treeError; d++) {
        const res = await groupedMeasure(levelFields.slice(0, d).map((f) => f.field), mf, filters);
        if (res.error) { item.treeError = res.error; break; }
        for (const r of res.rows) {
          if (!d) total = r[0];
          else byPath[r.slice(0, d).map(matchKey).join("\u0001")] = r[d];
        }
      }
      if (!item.treeError) item.treeValues = { total, byPath };
      log(`Decomposition tree "${item.title}": ${item.treeError ? `values not asked (${item.treeError}); kept as its table` : `${Object.keys(byPath).length} value(s) from Power BI`}`);
    }
  }
}

/*
 * A table's or matrix's Total row for the columns that don't add up from its rows (a rate, a
 * share, an average: Readmission Rate 10.8% for all surgeons together, not the rows' sum):
 * each asked of Power BI, the measure over all the rows under the same report, page and visual
 * filters (Build permission, else the query service, as loadTreeValues). item.reportTotals:
 * { column (normalized name): value }; a column Power BI doesn't answer for stays blank.
 */
async function loadTableTotals(collected, options) {
  for (const { data } of collected) {
    for (const item of data.items) {
      if (!TABLE_VISUAL_TYPES.includes(item.type) || item.isKpi || item.columnKinds || !item.rows || item.rows.length < 2) continue;
      const info = visualInfo(options.reportColors, item.page, item.name);
      if (!info || !info.roles) continue;
      const b = info.builtin || {};
      if ((item.type === "tableEx" && b.totals === false) || (item.type === "pivotTable" && b.rowSubtotals === false)) continue;
      const { header, data: rows, numeric, helper } = classifyColumns(item.rows);
      const cells = rows.map((row) => header.map((_, ci) => toCellValue(row[ci])));
      const formats = readableFormats(header, cells);
      const entries = fieldsByColumn(header, [...(info.roles.Values || []), ...(info.roles.Rows || [])]);
      const filters = [...(data.queryFilters || []), ...(item.visualFilters || [])];
      const totals = {};
      for (let ci = 0; ci < header.length; ci++) {
        const field = entries[ci] && entries[ci].field;
        if (!numeric[ci] || helper[ci] || !field || !["measure", "agg"].includes(field.kind)) continue;
        const fmt = formats[ci] || (cells.find((r) => typeof r[ci].value === "number") || [])[ci]?.numFmt;
        if (totalKind(field, header[ci], fmt)) continue;   // adds up: Excel's own total, as before
        const res = await groupedMeasure([], field, filters);
        let v = res.rows && res.rows[0] ? toCellValue(String(res.rows[0][0] ?? "")).value : null;
        if (typeof v === "number") v = onRowsScale(v, cells.map((r) => r[ci].value), fmt);
        if (typeof v === "number") totals[normalizeField(header[ci])] = v;
        else log(`"${item.title}": the total of ${header[ci]} wasn't given by Power BI${res.error ? ` (${res.error})` : ""}; left blank`);
      }
      if (Object.keys(totals).length) {
        item.reportTotals = totals;
        log(`"${item.title}": ${Object.keys(totals).length} total(s) from Power BI`);
      }
    }
  }
}

/*
 * Power BI's own value on the scale of the exported cells. A percentage column's cells hold
 * fractions ("10.4%" is 0.104), but a measure that returns 10.76 and only shows a "%" sign (a
 * format such as 0.0"%") comes back from Power BI as 10.76: when the value is a hundred times
 * the cells' size and a hundredth of it sits among them, it is that hundredth (0.1076, 10.8%).
 */
function onRowsScale(v, values, fmt) {
  if (!/%/.test(String(fmt || ""))) return v;
  const nums = values.filter((x) => typeof x === "number" && x !== 0).map(Math.abs);
  if (!nums.length || v === 0) return v;
  const lo = Math.min(...nums), hi = Math.max(...nums);
  const near = (x) => Math.abs(x) >= lo / 10 && Math.abs(x) <= hi * 10;
  return !near(v) && near(v / 100) ? v / 100 : v;
}

/* A measure grouped by some columns (none: one total) under the filters: { rows: [[...values, measure]] } or { error } */
async function groupedMeasure(columns, measure, filters) {
  let daxError = null;
  try {
    const value = daxValue(measure);
    if (!value) throw new Error("its measure can't be queried");
    const filterTables = (filters || []).filter((f) => !(f && f.target && f.target.measure)).map(daxFilter).filter(Boolean);
    const query = `EVALUATE SUMMARIZECOLUMNS(${[...columns.map((c) => daxColumn(c.entity, c.prop)), ...filterTables, `"v", ${value}`].join(", ")})`;
    const res = await executeDaxQuery(query);
    if (!res.error) return { rows: res.rows.map((r) => [...columns.map((c) => r[`${c.entity}[${c.prop}]`]), r["[v]"]]) };
    daxError = res.error;
  } catch (e) {
    daxError = errorText(e);
  }
  try {
    if (typeof websiteQuery !== "function") throw new Error("not available");
    return { rows: await websiteQuery([...columns, measure], filters) };
  } catch (e) {
    return { error: `${daxError}; ${errorText(e)}` };
  }
}

/* Rows per value of one field and of the analysed field: { rows: [[value, target, count]] } or { error } */
async function keyInfluencerCounts(field, target, filters) {
  let daxError = null;
  try {
    const filterTables = filters.map(daxFilter).filter(Boolean);
    const query = `EVALUATE SUMMARIZECOLUMNS(${[daxColumn(field.entity, field.prop), daxColumn(target.entity, target.prop), ...filterTables,
      `"n", COUNTROWS('${String(target.entity).replace(/'/g, "''")}')`].join(", ")})`;
    const res = await executeDaxQuery(query);
    if (!res.error) {
      const v = `${field.entity}[${field.prop}]`, t = `${target.entity}[${target.prop}]`;
      return { rows: res.rows.map((r) => [r[v], r[t], r["[n]"]]) };
    }
    daxError = res.error;
  } catch (e) {
    daxError = errorText(e);
  }
  try {
    if (typeof websiteQuery !== "function") throw new Error("not available");
    // Count (5): the analysed column's rows, grouped by it and the field
    const rows = await websiteQuery([field, target, { kind: "agg", entity: target.entity, prop: target.prop, agg: 5 }], filters);
    return { rows };
  } catch (e) {
    return { error: `the rows behind it couldn't be counted (${daxError}; ${errorText(e)})` };
  }
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
  if (typeof notePageTitles === "function") notePageTitles(pages);   // finds visuals by page title where page names differ

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
  // The drill levels as they are in the app now, before the export switches pages
  const liveDrill = options.matchColors ? await captureLiveDrill(embeddedReport) : null;
  // The page on screen is read first, while what the user clicked on it still holds (changing
  // page can clear it); the sheets keep the report's order
  const order = pages.map((_, i) => i).sort((a, b) => (pages[b].name === original.name) - (pages[a].name === original.name));
  try {
    for (let k = 0; k < order.length; k++) {
      const page = pages[order[k]];
      onProgress({ pageIndex: k + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: 0, visualCount: 0, visualTitle: "Opening page…" });
      await activatePage(page);
      const selection = page.name === original.name ? liveSelections[page.name] || null : null;
      const data = await collectPage(page, { ...options, selection },
        (i, n, title) => onProgress({ pageIndex: k + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: i, visualCount: n, visualTitle: title }),
        isCancelled);
      if (liveDrill && liveDrill[page.name]) data.items.forEach((item) => { item.liveDrill = liveDrill[page.name][item.name] || null; });
      collected[order[k]] = { page, data };
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

  // Each visual at the drill level the report shows, then table sparklines (both need the report file)
  if (options.reportColors) {
    onProgress({ pageIndex: pages.length, pageCount: pages.length, pageName: pages[pages.length - 1].displayName, visualIndex: 0, visualCount: 0, visualTitle: "Matching drill levels…" });
    await loadDrillLevels(collected, options);
    onProgress({ pageIndex: pages.length, pageCount: pages.length, pageName: pages[pages.length - 1].displayName, visualIndex: 0, visualCount: 0, visualTitle: "Reading sparklines…" });
    await loadSparklines(collected, options);
    await loadKeyInfluencers(collected, options);
    await loadTreeValues(collected, options);
    await loadTableTotals(collected, options);
  }

  // "Match the report" accent: the report theme's main colour, else a neutral navy
  if (branding.accentAuto) {
    if (!theme) theme = await readThemeColors();
    const main = options.reportColors ? options.reportColors.palette[0] : theme.fromReport ? theme.colors[0] : null;
    branding = { ...branding, accent: main || "#" + NEUTRAL_ACCENT.slice(2) };
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
  const chartSheets = [];   // charts mode: { sheetName, charts } for addChartsToFile
  const sparkSheets = [];   // trend cards: { sheetName, sparklines } for addSparklinesToFile
  const slicerSheets = [];  // Excel slicers: { sheetName, slicers } for addSlicersToFile
  // With several pages, one with nothing to show (a cover or info page of images and text)
  // isn't given an empty sheet; the first sheet's notes name it
  const filled = collected.filter(({ data }) => data.items.length || data.notes.length);
  const sheets = collected.length > 1 && filled.length ? filled : collected;
  const left = collected.filter((c) => !sheets.includes(c)).map(({ page }) => `"${page.displayName}"`);
  if (left.length) sheets[0].data.notes.push(`Not included: ${left.join(", ")} – no visuals with data.`);
  sheets.forEach(({ page, data }, i) => {
    const sheetName = nameSheet(page.displayName);
    // A hidden sheet for the page's table sparklines, named next to the page
    const trendSheet = data.items.some((it) => it.sparklines && it.sparklines.length) ? nameSheet(`Trends ${i + 1}`) : null;
    const ctx = { pageIndex: i + 1, pageCount: sheets.length, dateText, trendSheet };
    const counts = writePageSheet(wb, sheetName, page, data, options, branding, logo, nameTable, ctx);
    // Charts mode: this page's chart numbers go on its own "Data" sheet, right after it
    if (counts.charts.length) {
      writeChartDataSheet(wb, nameSheet(`${page.displayName} – Data`), counts.charts);
      chartSheets.push({ sheetName, charts: counts.charts });
    }
    if (counts.trendSheet) writeTrendSheet(wb, counts.trendSheet);
    if (counts.sparklines.length) sparkSheets.push({ sheetName, sparklines: counts.sparklines });
    if (counts.slicers.length) slicerSheets.push({ sheetName, slicers: counts.slicers });
    totals.tables += counts.tables;
    totals.kpis += counts.kpis;
    data.notes.forEach((n) => totals.notes.push(sheets.length > 1 ? `${page.displayName}: ${n}` : n));
  });

  let buffer = await wb.xlsx.writeBuffer();
  // Charts mode: add the real Excel charts to the finished file
  if (chartSheets.length) buffer = await addChartsToFile(buffer, chartSheets);
  if (sparkSheets.length) buffer = await addSparklinesToFile(buffer, sparkSheets);
  if (slicerSheets.length) buffer = await addSlicersToFile(buffer, slicerSheets);
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const baseName = sheets.length === 1 ? `${currentReportName} - ${sheets[0].page.displayName}` : currentReportName;
  const fileName = (baseName.replace(/[\\/:*?"<>|]/g, "").trim() || "Power BI export") + ".xlsx";
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(link.href), 10000);

  log(`Saved "${fileName}" – ${sheets.length} sheet(s), ${totals.tables} table(s), ${totals.kpis} KPI(s), ${totals.notes.length} note(s)`);
  return { fileName, sheets: sheets.length, tables: totals.tables, kpis: totals.kpis, notes: totals.notes };
}
