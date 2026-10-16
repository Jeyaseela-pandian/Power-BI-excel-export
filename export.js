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
const MATRIX_AS_TABLE = true;   // matrices as normal Excel tables (false: Power BI's stepped layout with groups and subtotals)
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
      const order = seriesOrder(data, legendCol, cols.filter((ci) => ci !== legendCol && !seriesCols.includes(ci) && !hex[ci]));
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
  const seriesColor = (ci, slot = indexOf[ci] ?? 0) => {
    const f = fieldOf[ci];
    const set = overrides["m:" + norm[ci]] || (f && (overrides["m:" + f.ref] || (f.label && overrides["m:" + f.label])));
    if (set) return toArgb(set);
    if (overrides.all) return toArgb(overrides.all);
    if (info && info.rule === "single") return color(0);
    return color(slot);
  };
  const keyedRows = (keyCol, colorOfValue, bySeries) => {
    const values = bySeries
      ? seriesOrder(data, keyCol, cols.filter((ci) => ci !== keyCol && !seriesCols.includes(ci) && !hex[ci]))
      : [...new Set(filled(keyCol))];   // order of first appearance
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

  // 4. Rows coloured by a field: the legend (Series) or, for pie / donut / treemap, the category
  let rowKey = null;
  if (info && (info.rule === "legend" || info.rule === "item")) {
    let keyCol = findCol(info.rule === "legend" ? info.series : info.category);
    if (keyCol < 0 && info.rule === "item" && !numeric[0] && !hex[0]) keyCol = 0;
    const oneColor = custom && custom.settings && custom.settings.all && overrides.all;   // e.g. a word cloud's single colour
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
const CHANGE_LIKE = /\b(yoy|mom|qoq|wow|growth|change|var|variance|index|score|share)\b/i;
const SUBTOTAL_FN = { sum: 9, min: 5, max: 4 };

/* How a column totals: "sum", "min", "max", or null (left blank) */
function totalKind(field, name, fmt) {
  if (/%/.test(fmt || "")) return null;
  const names = [name, field && field.prop].filter(Boolean).join(" ");
  if (field) {
    if (field.kind === "agg") return field.agg === 0 || field.agg === 5 ? "sum" : field.agg === 3 ? "min" : field.agg === 4 ? "max" : null;
    if (field.kind === "measure") return NON_ADDITIVE.test(names) || COUNT_LIKE.test(names) || CHANGE_LIKE.test(names) ? null : "sum";
    return null;   // a column shown as it is, a hierarchy level
  }
  return NON_ADDITIVE.test(names) || COUNT_LIKE.test(names) || CHANGE_LIKE.test(names) || RATE_LIKE.test(names) ? null : "sum";
}

/* Each column's field in the report: its renamed label first, then its name */
function fieldsByColumn(header, entries) {
  const lower = header.map((h) => String(h ?? "").trim().toLowerCase());
  const norm = header.map((h) => normalizeField(h));
  const out = header.map(() => null), taken = new Set();
  for (const e of entries || []) {
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
    return totalKind(info && entries[ci] ? entries[ci].field || null : null, h, fmt);
  });
  return kinds.some(Boolean) ? { kinds } : null;
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
    cell.value = { formula: `SUBTOTAL(${SUBTOTAL_FN[kind]},${colRange(ws, col + ci, hr + 1, hr + n)})`, result: subtotalOf(kind, values) };
    const fmt = table.formats[ci] || (table.data.find((r) => typeof r[ci].value === "number") || [])[ci]?.numFmt;
    if (fmt) cell.numFmt = fmt;
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
    let kept = (item.drillKept ? list : shownLevels(list)).filter(present);
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
    const list = roles[r] || [], shown = shownLevels(list);
    return list.filter((f) => !shown.includes(f));
  });
  if (!hiddenFields.length) return null;
  const header = item.rows[0], data = item.rows.slice(1);
  const keyFields = new Set(KEY_ROLES.flatMap((r) => shownLevels(roles[r] || [])));
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
      if (res.error) error = res.error;
      else {
        const byKey = new Map(res.rows.map((r) => [keyCols.map((ci) => matchKey(r[`${entries[ci].field.entity}[${entries[ci].field.prop}]`])).join("\u0001"), r]));
        const keys = [...groups.keys()];
        if (keys.some((k) => !byKey.has(k))) error = "its answer didn't match the exported rows";
        else keys.forEach((k, gi) => {
          const r = byKey.get(k);
          open.forEach((ci, i) => {
            const v = r[`[v${i}]`];
            out[gi][ci] = typeof v === "number" ? displayText(v, fmtOf(ci)) : "";
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
    const isLive = (f) => !!f.field && live.some((l) => key(l) === key(f.field) || (!f.field.entity && String(l.prop).toLowerCase() === String(f.field.prop || "").toLowerCase()));
    if (!list.some(isLive)) continue;   // the fields didn't match: keep the saved levels
    roles[role] = list.map((f) => ({ ...f, active: isLive(f) }));
  }
  return { ...info, roles, expansion: sv.expansionStates || info.expansion };
}

/* For every visual: its rows at the drill level the report shows (needs the report file) */
async function loadDrillLevels(collected, options) {
  for (const { data } of collected) {
    for (const item of data.items) {
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
    return { k, ci, name: header[ci], kind: totalKind(f.field, header[ci], fmt) };
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
  item.cellFormats = [null, ...columns.map((c) => formats[c.m.ci] || null)];

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
function splitCards(item, info) {
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
    const compact = !!compactFormat(hero.value, fmt);
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
  const tinted = (font, i) => (valueColor(i) ? { ...font, color: { argb: valueColor(i) } } : font);
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
    widths.want(col, Math.min(60, cellUnits(header[i], fonts.kpiSubLabel, 1)));
    widths.want(col + 1, cellUnits(displayText(cells[i].value, fmtOf(i)), fonts.kpiSubValue, 1));
  }

  const last = Math.max(next - 1, strip ? r + strip.height - 1 : 0);
  for (let rr = top; rr <= last; rr++) {
    for (let cc = col; cc <= col + 1; cc++) {
      const cell = ws.getCell(rr, cc);
      cell.fill = SURFACE_FILL;
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
    cellUnits(labelText, fonts.kpiLabel, 1),
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
    if (fnOf(ci) && n) cell.value = { formula: `SUBTOTAL(${fnOf(ci)},${colRange(ws, col + ci, hr + 1, hr + n)})`, result: resultOf(ci, 0, n - 1) };
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
          cfvo: [{ type: "min" }, { type: "max" }], color: { argb: toArgb(f.bar.color) } };
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
  for (const item of items) {
    const info = atDrillLevel(item, withLiveDrill(item, visualInfo(lay.reportColors, item.page, item.name)));
    if (item.isKpi) mainValueFirst(item, info);
    if ((customVisual(item.type) || {}).rule === "nested") nestedBarsRows(item, info, colorNotes);
    if ((customVisual(item.type) || {}).layout === "grid") comparisonRows(item, info);
    if ((customVisual(item.type) || {}).layout === "bins") histogramRows(item, info);
    if ((customVisual(item.type) || {}).layout === "trend") dualKpiTrend(item, info, !!lay.palette);
    if (item.type === "kpi" && !item.isKpi) kpiTrend(item, info, !!lay.palette);
    if (item.type === "pivotTable" && info && !MATRIX_AS_TABLE) matrixLayout(item, info);
    if (item.sparklines && item.sparklines.length && !item.stepped) addSparklineColumns(item);
    if (TABLE_VISUAL_TYPES.includes(item.type) && !item.stepped) item.totals = totalsPlan(item, info);
    // A title turned off in the report isn't shown either (it is often an internal name such
    // as "TYS Card"): a table is titled from its columns instead, a KPI tile by its label
    const hidden = !!(info && info.titleHidden);
    let title = hidden || isGenericTitle(item.title, item.type) ? (item.isKpi ? null : describeTitle(item)) : item.title;
    if (title) {
      const key = title.toLowerCase();
      usedTitles[key] = (usedTitles[key] || 0) + 1;
      if (usedTitles[key] > 1) title = `${title} (${usedTitles[key]})`;
    }
    item.displayTitle = title;
    // A trend visual's colours are on its cards; its table stays plain
    item.colors = lay.palette && !item.isKpi && !item.trend ? columnColors(item, lay.palette, info) : null;
    item.tableFormat = item.gridFormat || (info && info.tableFormats && info.tableFormats.length ? info : null);
    item.colorInfo = info;
    item.chart = options.visualsAs === "charts" && !item.trend ? chartPlan(item) : null;   // charts mode only
    if (item.chart && item.chart.note) colorNotes.push(item.chart.note);
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
        const { w, h } = chartSizePx(item);
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
          const table = writeDataTable(ws, r, col, item, nameTable(item.displayTitle), widths, lay, number);
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

  const filterChips = [];
  // The filters as models too: table sparklines are queried under them
  const queryFilters = [...await filtersOf(embeddedReport), ...await filtersOf(page)];
  if (options.includeFilters) {
    filterParts(queryFilters).forEach((p) => filterChips.push({ all: false, ...p }));
  }

  const items = [];
  const notes = [];
  for (let i = 0; i < visuals.length; i++) {
    if (isCancelled()) throw { cancelled: true };
    const v = visuals[i];
    const title = v.title || `${v.type} (${v.name})`;
    onProgress(i + 1, visuals.length, title);

    // Slicers become filter chips: "Field  selection", or "Field  All". A hidden slicer still
    // filters the page (bookmarks often hide them), so it gets a chip too.
    if (isSlicerType(v.type)) {
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
      items.push({ title, name: v.name, page: page.name, type: v.type, layout: v.layout, rows, isKpi, visualFilters });
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
  return { filterChips, items, notes, queryFilters };
}

/* Visuals whose data Power BI doesn't export, by type: what to call them in a note */
const NOT_EXPORTABLE = {
  keyDriversVisual: "key influencers visuals", qnaVisual: "Q&A visuals", scriptVisual: "R visuals", pythonVisual: "Python visuals",
  rdlVisual: "paginated report visuals", powerApps: "Power Apps visuals", PowerApps_PowerBI_Visual: "Power Apps visuals",
};

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
    fonts, accent, sheetName,
    trendSheet: ctx.trendSheet ? makeTrendSheet(ctx.trendSheet) : null,
    tableStyle: branding.tableStyle || EXECUTIVE,
    palette: options.matchColors ? options.palette : null,
    reportColors: options.matchColors ? options.reportColors || null : null,
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
};
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
  const { w, h } = chartSizePx(item);
  return { width: Math.max(2, Math.round(w / 90)), height: Math.ceil(h / ROW_PX) + 1, note: false };
}

/*
 * What a chart visual draws, from its exported rows: categories (the axis), one series
 * per measure (or per legend value), and colours when "Match the report's chart colours"
 * is on. Returns null when the visual can't be an Excel chart (it then stays a table).
 */
function chartPlan(item) {
  const spec = EXCEL_CHARTS[item.type];
  if (!spec || item.isKpi) return null;
  // Scatter, waterfall and funnel have their own builders; any problem keeps the visual as a table
  if (spec.kind === "scatter" || spec.kind === "waterfall" || spec.kind === "funnel") {
    try {
      return spec.kind === "scatter" ? scatterPlan(item) : spec.kind === "waterfall" ? waterfallPlan(item) : funnelPlan(item);
    } catch (e) {
      log(`Chart for "${item.title}" not drawn, kept as a table: ${errorText(e)}`);
      return null;
    }
  }
  const { header, data, numeric, hex, helper } = classifyColumns(item.rows);
  if (!data.length) return null;
  const cols = header.map((_, ci) => ci);
  const colors = item.colors || null;
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
      .filter((m) => m.ci >= 0 && numeric[m.ci] && !hex[m.ci]);
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
    axes: (info && info.axes) || null };

  // Pie / donut: one series, a colour per slice
  if (spec.kind === "pie" || spec.kind === "doughnut") {
    const vc = valueCols[0];
    const byKey = colors && colors.rowKey && colors.rowKey.col === catCol ? colors.rowKey.colorOf : null;
    return {
      ...base,
      series: [{ name: header[vc], values: valuesFor(vc, data), fmt: formatOf(vc), color: null }],
      pointColors: byKey ? categories.map((c) => byKey[c] || null) : null,
    };
  }

  // A legend splits one measure into one series per legend value
  let legendCol = -1;
  if (info && info.rule === "legend" && info.series) {
    legendCol = header.findIndex((h) => normalizeField(h) === info.series.ref || String(h).trim().toLowerCase() === info.series.label);
  } else if (!info && textCols.length === 2 && valueCols.length === 1) {
    legendCol = textCols[1];
  }
  if (legendCol >= 0 && legendCol !== catCol) {
    // A combo's legend splits only its column values: its line values (Y2) stay one line
    // each, on a second axis. A line value repeats on every legend row of a category, so it
    // is read once per category, not added up.
    const lines = spec.kind === "combo" && lineCols ? valueCols.filter((ci) => lineCols.includes(ci)) : [];
    const vc = valueCols.find((ci) => !lines.includes(ci)) ?? valueCols[0];
    const byKey = colors && colors.rowKey && colors.rowKey.col === legendCol ? colors.rowKey.colorOf : null;
    const legends = byKey ? Object.keys(byKey) : seriesOrder(data, legendCol, [catCol]);
    const once = (ci) => categories.map((c) => {
      const r = data.find((x) => catOf(x) === c && num(x[ci]) !== null);
      return r ? num(r[ci]) : null;
    });
    return {
      ...base,
      valueName: header[vc],
      series: [
        ...legends.map((name) => ({
          name,
          values: valuesFor(vc, data.filter((r) => key(r, legendCol) === name)),
          fmt: formatOf(vc),
          color: (byKey && byKey[name]) || null,
        })),
        ...lines.map((ci) => ({ name: header[ci], values: once(ci), fmt: formatOf(ci), color: (colors && colors.series && colors.series[ci]) || null, line: true })),
      ],
    };
  }

  // One series per measure; a combo draws its line values (its last measure, without the
  // report file) as lines on a second axis
  const series = valueCols.map((ci, i) => ({
    name: header[ci],
    values: valuesFor(ci, data),
    fmt: formatOf(ci),
    color: (colors && colors.series && colors.series[ci]) || null,
    line: spec.kind === "combo" && (lineCols ? lineCols.includes(ci) : valueCols.length > 1 && i === valueCols.length - 1),
  }));
  // A combo showing only line values is a line chart
  if (series.every((s) => s.line)) {
    series.forEach((s) => { s.line = false; });
    base.spec = { kind: "line" };
  }
  // One measure coloured per bar (colours set by hand, or a colour rule)
  let pointColors = null;
  if (series.length === 1 && colors) {
    const ci = valueCols[0];
    if (colors.perCol && colors.perCol[ci]) pointColors = categories.map((c) => colors.perCol[ci][firstRowOf(c)] || null);
    else if (colors.rowKey && colors.rowKey.col === catCol) pointColors = categories.map((c) => colors.rowKey.colorOf[c] || null);
    else if (colors.tabColors) pointColors = categories.map((c) => colors.tabColors[firstRowOf(c)] || null);
  }
  return { ...base, series, pointColors };
}

/* Shared helpers for the chart builders below */
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

/* One value per category, in the order they appear ("Total" rows are left out) */
function chartSteps(t, catCol, valueCol) {
  const additive = !NON_ADDITIVE.test(t.header[valueCol]);
  const steps = [];
  for (const r of t.data) {
    const label = t.key(r, catCol), v = t.num(r[valueCol]);
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
  const steps = chartSteps(t, catCol, valueCol);
  if (!steps.length) return null;
  const w = (item.colorInfo && item.colorInfo.waterfall) || {};
  const up = toArgb(w.increase || "#1AAB40"), down = toArgb(w.decrease || "#D64554"), total = toArgb(w.total || "#118DFF");
  const fmt = t.formatOf(valueCol);
  const negFmt = `"-"${fmt}`;   // a decrease is stored as its size and shown with a minus: -25.36%
  const starts = [], ends = [];
  let running = 0;
  steps.forEach((s) => { starts.push(running); running += s.value; ends.push(running); });
  // The axis like Power BI's (-40% to 20% in 20% steps), ending exactly on a step
  const size = chartSizePx(item);
  const axis = niceAxis([0, ...starts, ...ends], size.h - 100, 45, 0) || { min: Math.min(0, ...ends), max: Math.max(0, ...ends), step: 0 };
  const tidy = (x) => Number(x.toFixed(12));
  const d = item.colorInfo && item.colorInfo.display;
  const showLabels = !(d && d.labels && d.labels.show === false);   // labels on unless switched off in the report
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
    ],
    steps: [...steps.map((s) => s.value), running],
    axis, fmt,
    showLabels,
    labelInside: inside,   // per bar: label inside the bar's end (white) instead of beside it
    labelPrecision: d && d.labels ? d.labels.precision : undefined,
    pointColors: null,
  };
}

/* Funnel: centred bars (an invisible offset pushes each bar to the middle), largest on top */
function funnelPlan(item) {
  const t = chartColumns(item);
  const catCol = t.textCols.length ? t.textCols[0] : -1;
  const valueCol = t.measureCols.find((ci) => ci !== catCol);
  if (catCol < 0 || valueCol === undefined) return null;
  const stages = chartSteps(t, catCol, valueCol);
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
    catName: t.header[catCol],
    categories: stages.map((s) => s.label),
    series: [
      { name: "Offset", values: stages.map((s) => (max - s.value) / 2), fmt, color: null, hidden: true, noLegend: true },
      { name: t.header[valueCol], values: stages.map((s) => s.value), fmt, color, labels: true },
    ],
    pointColors,
  };
}

/* Value labels in the middle of each bar (used by the funnel) */
function dataLabelsXml(fmt, argb) {
  // Labels inside a bar: white or ink, whichever reads on the bar's colour (Power BI's white)
  const text = argb ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1">${fillXml(textOn(argb))}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` : "";
  return `<c:dLbls><c:numFmt formatCode="${xmlEsc(fmt || "General")}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${text}` +
    `<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>` +
    `<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

/* Scatter / bubble chart XML: each legend group reads its own block of rows on the Data sheet */
function scatterXml(plan) {
  const bubble = plan.spec.kind === "bubble";
  const [xs, ys, ss] = plan.series;
  const range = plan.refs.range;
  const labels = plan.perPoint ? nameLabelsXml() : "";
  const numRef = (i, g, s) => `<c:numRef><c:f>${xmlEsc(range(i, g.start, g.end))}</c:f>` +
    `${numCacheXml(s.values.slice(g.start, g.end + 1), s.fmt)}</c:numRef>`;
  const fill = (argb) => (argb ? `<a:solidFill><a:srgbClr val="${String(argb).slice(-6).toUpperCase()}"><a:alpha val="85000"/></a:srgbClr></a:solidFill>` : "");
  const sers = plan.groups.map((g, i) => {
    const head = `<c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xmlEsc(g.name)}</c:v></c:tx>`;
    const xv = `<c:xVal>${numRef(0, g, xs)}</c:xVal>`, yv = `<c:yVal>${numRef(1, g, ys)}</c:yVal>`;
    const look = `<c:spPr>${fill(g.color)}<a:ln w="9525">${fillXml(C.WHITE)}</a:ln></c:spPr>`;
    if (bubble) {
      return `<c:ser>${head}${look}<c:invertIfNegative val="0"/>${labels}${xv}${yv}<c:bubbleSize>${numRef(2, g, ss)}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
    }
    return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>` +
      `<c:marker><c:symbol val="circle"/><c:size val="8"/>${look}</c:marker>${labels}${xv}${yv}<c:smooth val="0"/></c:ser>`;
  }).join("");
  const A1 = 50110, A2 = 50120;
  const chart = bubble
    ? `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="40"/><c:showNegBubbles val="0"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:bubbleChart>`
    : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:scatterChart>`;
  // Axes along the bottom and left edges with their field names, as Power BI draws them
    // Axis range and steps like Power BI's automatic axes, sized to the visual (one step per
  // ~110 px across, ~45 px up), so the gridlines fall on the same values as in the report
  const size = plan.sizePx || { w: 0, h: 0 };
  const xAxis = niceAxis(xs.values, size.w - 90, 110), yAxis = niceAxis(ys.values, size.h - 100, 45);
  return chart +
    valAxXml(A1, A2, "b", stepFormat(xs.fmt, xAxis, xs.values), { crosses: "min", title: xs.name, axis: xAxis }) +
    valAxXml(A2, A1, "l", stepFormat(ys.fmt, yAxis, ys.values), { crosses: "min", title: ys.name, axis: yAxis });
}

/* Number format for axis labels, shortened the way Power BI does: "$4,000,000" → "$4.0M", "34.28%" → "34%" */
function axisFormat(fmt, values) {
  const f = String(fmt || "General");
  const nums = (values || []).filter((v) => typeof v === "number" && isFinite(v));
  const max = nums.length ? Math.max(...nums.map((v) => Math.abs(v))) : 0;
  if (f.includes("%")) return max >= 0.05 ? "0%" : "0.0%";
  const prefix = (f.match(/^"[^"]*"/) || [""])[0];
  if (max >= 1e9) return `${prefix}#,##0.0,,,"B"`;
  if (max >= 1e6) return `${prefix}#,##0.0,,"M"`;
  if (max >= 1e4) return `${prefix}#,##0.0,"K"`;
  return f;
}

/* A small grey axis title, like Power BI's */
function axisTitleXml(text) {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0"/></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="900" b="0">${fillXml(C.SLATE)}</a:rPr><a:t>${xmlEsc(text)}</a:t></a:r></a:p></c:rich></c:tx>` +
    `<c:overlay val="0"/></c:title>`;
}

/* Bubble labels: each bubble's name beside it ("050-Shoes") */
function nameLabelsXml() {
  return `<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800">${fillXml(C.SLATE)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` +
    `<c:dLblPos val="r"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="1"/>` +
    `<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

/* Pie labels: each slice's share (the value shows on hover) */
function pieLabelsXml(pos) {
  return `<c:dLbls><c:numFmt formatCode="0.00%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    (pos ? `<c:dLblPos val="${pos}"/>` : "") +
    `<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>` +
    `<c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>`;
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

/* Axis range and step like Power BI's automatic axes: the step is the first of 1, 2, 5, 10…
   that leaves room for each label (minGapPx between gridlines on a plot plotPx long) */
function niceAxis(values, plotPx, minGapPx, pad = 0.03) {
  const v = (values || []).filter((x) => typeof x === "number" && isFinite(x));
  if (!v.length || !plotPx) return null;
  let lo = Math.min(...v), hi = Math.max(...v);
  if (lo === hi) { const d = Math.abs(lo) || 1; lo -= d / 2; hi += d / 2; }
  const range = hi - lo, padded = range * 1.1;
  const most = Math.max(2, plotPx / minGapPx);
  const p = Math.pow(10, Math.floor(Math.log10(padded / most)));
  const step = [1, 2, 5, 10, 20, 50, 100].map((m) => m * p).find((s) => padded / s <= most) || 100 * p;
  const tidy = (x) => Number(x.toFixed(12));
  // pad: room beyond the data (bubbles need some; waterfall bars end exactly on a step)
  let min = tidy(Math.floor((lo - range * pad) / step + 1e-9) * step);
  const max = tidy(Math.ceil((hi + range * pad) / step - 1e-9) * step);
  if (lo >= 0 && min < 0) min = 0;
  return { min, max, step: tidy(step) };
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
  const A1 = 50210, A2 = 50220, A3 = 50230, A4 = 50240;
  const last = categories.length - 1;
  const noBox = `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`;
  const flags = (show) => `<c:showLegendKey val="0"/><c:showVal val="${show ? 1 : 0}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`;
  // Small labels like Power BI's: 8pt grey beside a bar, white bold inside one
  const label = (i, pos, argb, bold) => `<c:dLbl><c:idx val="${i}"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="800" b="${bold ? 1 : 0}">${fillXml(argb)}</a:rPr><a:t>${xmlEsc(waterfallLabel(plan.steps[i], plan.fmt, plan.labelPrecision))}</a:t></a:r></a:p></c:rich></c:tx>` +
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
    catAxXml(A1, A2, "b", { title: titles.cat, hidden: !titles.catShow, count: categories.length }) +
    valAxXml(A2, A1, "l", stepFormat(plan.fmt, axis), { axis, title: titles.val, hidden: !titles.valShow }) +
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
  for (const ch of charts) {
    const p = ch.plan;
    const t = ws.getCell(row, 2);
    t.value = p.title;
    t.font = { name: "Calibri", size: 12, bold: true, color: { argb: C.INK } };
    const hr = row + 1;
    const h0 = ws.getCell(hr, 2);
    h0.value = p.catName || "Category";
    h0.font = { name: "Calibri", size: 10, bold: true, color: { argb: C.INK } };
    h0.fill = SURFACE_FILL;
    p.series.forEach((s, i) => {
      const h = ws.getCell(hr, 3 + i);
      h.value = s.name;
      h.font = { name: "Calibri", size: 10, bold: true, color: { argb: C.INK } };
      h.fill = SURFACE_FILL;
      h.alignment = { horizontal: "right" };
    });
    p.categories.forEach((c, ri) => {
      ws.getCell(hr + 1 + ri, 2).value = axisLabel(c);   // a date reads "1/15/2023" on the axis, an empty one "(Blank)"
      p.series.forEach((s, i) => {
        const cell = ws.getCell(hr + 1 + ri, 3 + i);
        cell.value = s.values[ri];
        if (s.fmt && s.fmt !== "General") cell.numFmt = s.fmt;
      });
    });
    const first = hr + 1, last = hr + p.categories.length;
    p.refs = {
      catRef: `${sheetRef}!${abs(first, 2)}:${abs(last, 2)}`,
      series: p.series.map((s, i) => ({
        nameRef: `${sheetRef}!${abs(hr, 3 + i)}`,
        valRef: `${sheetRef}!${abs(first, 3 + i)}:${abs(last, 3 + i)}`,
      })),
    };
    // Any block of rows in one Data column (scatter charts read one block per legend value)
    p.refs.range = (i, from, to) => `${sheetRef}!${abs(first + from, 3 + i)}:${abs(first + to, 3 + i)}`;
    widest = Math.max(widest, p.series.length);
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

function chartSeriesXml(kind, s, i, plan, ref) {
  const head = `<c:idx val="${i}"/><c:order val="${i}"/>` +
    `<c:tx><c:strRef><c:f>${xmlEsc(ref.nameRef)}</c:f>${strCacheXml([s.name])}</c:strRef></c:tx>`;
  const cat = `<c:cat><c:strRef><c:f>${xmlEsc(plan.refs.catRef)}</c:f>${strCacheXml(plan.categories.map(axisLabel))}</c:strRef></c:cat>`;
  const val = `<c:val><c:numRef><c:f>${xmlEsc(ref.valRef)}</c:f>${numCacheXml(s.values, s.fmt)}</c:numRef></c:val>`;
  // Hidden helper series (waterfall base, funnel offset) take no bar colours
  const points = s.hidden ? "" : (plan.pointColors || []).map((c, pi) => (c
    ? `<c:dPt><c:idx val="${pi}"/>${kind === "bar" ? '<c:invertIfNegative val="0"/>' : ""}<c:bubble3D val="0"/><c:spPr>${fillXml(c)}</c:spPr></c:dPt>`
    : "")).join("");
  if (kind === "line") {
    const ln = s.color ? `<c:spPr><a:ln w="28575" cap="rnd">${fillXml(s.color)}<a:round/></a:ln></c:spPr>` : "";
    return `<c:ser>${head}${ln}<c:marker><c:symbol val="none"/></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
  }
  const sp = s.hidden ? `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`
    : s.color ? `<c:spPr>${fillXml(s.color)}</c:spPr>` : "";
  const labels = s.labels ? dataLabelsXml(s.fmt, s.color || (s.pointColors && s.pointColors.find(Boolean)) || null) : "";
  if (kind === "bar") return `<c:ser>${head}${sp}<c:invertIfNegative val="0"/>${points}${labels}${cat}${val}</c:ser>`;
  if (kind === "area") return `<c:ser>${head}${sp}${cat}${val}</c:ser>`;
  const pieLabels = s.pieLabels ? pieLabelsXml(s.pieLabelPos) : "";
  return `<c:ser>${head}${sp}${points}${pieLabels}${cat}${val}</c:ser>`;   // pie, doughnut
}
/* A category axis. Every label is drawn (Excel otherwise skips some when it thinks they are
   crowded, so labels came and went with the chart's size); up to 60 categories, beyond which
   Excel's own spacing reads better */
function catAxXml(id, crossId, pos, { reverse = false, hidden = false, title = null, count = 0 } = {}) {
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${reverse ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${hidden && !title ? 1 : 0}"/><c:axPos val="${pos}"/>` + (title ? axisTitleXml(title) : "") +
    `<c:numFmt formatCode="General" sourceLinked="1"/>` +
    `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="${hidden ? "none" : "low"}"/>` +
    `<c:spPr><a:ln w="9525">${hidden ? "<a:noFill/>" : fillXml(C.RULE)}</a:ln></c:spPr>` +
    `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>` +
    (count && count <= 60 && !hidden ? `<c:tickLblSkip val="1"/>` : "") + `<c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/* A chart's axis titles as Power BI draws them by default: the category field, and the value
   fields ("Last Year Sales and This Year Sales"); a title or axis the report turns off isn't
   drawn. Returns { cat, val, sec, catShow, valShow, secShow } */
function axisTitles(plan, main, lines) {
    // A funnel has no axis titles in Power BI: only its stage names and values
  if (plan.funnel) return { catShow: true, valShow: true, secShow: true, cat: null, val: null, sec: null };
  const a = plan.axes || {};
  const on = (x, key) => !x || x[key] !== false;
  const names = (list) => {
    const n = [...new Set(list.map(([s]) => s.name))];
    return n.length <= 2 ? n.join(" and ") : `${n.slice(0, -1).join(", ")} and ${n[n.length - 1]}`;
  };
  return {
    catShow: on(a.cat, "show"), valShow: on(a.val, "show"), secShow: on(a.sec, "show"),
    // An axis turned off loses its labels, not its title (Power BI still draws the title)
        // The author's own title text ("Revenue") when the report sets one, else the field names
    cat: on(a.cat, "title") ? (a.cat && a.cat.text) || plan.catName : null,
    val: on(a.val, "title") ? (a.val && a.val.text) || plan.valueName || names(main) : null,
    sec: lines.length && on(a.sec, "title") ? (a.sec && a.sec.text) || names(lines) : null,
  };
}

function valAxXml(id, crossId, pos, fmt, { crosses = "autoZero", grid = true, hidden = false, title = null, axis = null, labels = true, zero = false } = {}) {
  // zero: bars and columns start at 0 as Power BI's do; Excel's own minimum can cut close values short
  const bounds = axis ? `<c:max val="${axis.max}"/><c:min val="${axis.min}"/>` : zero ? `<c:min val="0"/>` : "";
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${bounds}</c:scaling><c:delete val="${hidden && !title ? 1 : 0}"/><c:axPos val="${pos}"/>` +
    (grid ? `<c:majorGridlines><c:spPr><a:ln w="6350">${fillXml(C.HAIR)}</a:ln></c:spPr></c:majorGridlines>` : "") +
    (title ? axisTitleXml(title) : "") +
    `<c:numFmt formatCode="${xmlEsc(fmt || "General")}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="${labels && !hidden ? "nextTo" : "none"}"/>` +
    `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="${crossId}"/><c:crosses val="${crosses}"/><c:crossBetween val="between"/>` +
    (axis ? `<c:majorUnit val="${axis.step}"/>` : "") + `</c:valAx>`;
}

function chartXml(plan) {
  const { spec, series } = plan;
  const kind = spec.kind;
  const refs = plan.refs;
  const scatter = kind === "scatter" || kind === "bubble";
  let body;
    if (scatter) {
    body = scatterXml(plan);
  } else if (kind === "waterfall") {
    body = waterfallXml(plan);
  } else if (kind === "pie" || kind === "doughnut") {
    const s0 = { ...series[0], pieLabels: true, pieLabelPos: kind === "pie" ? "outEnd" : null };
    const ser = chartSeriesXml("pie", s0, 0, plan, refs.series[0]);
    body = kind === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  } else {
    const indexed = series.map((s, i) => [s, i]);
    const main = indexed.filter(([s]) => !s.line);
    const lines = indexed.filter(([s]) => s.line);
    const horizontal = spec.dir === "bar";
    const pct = spec.grouping === "percentStacked";
    // Axis labels shortened like Power BI ($4.0M, 30%); stacked bars are sized by their totals
    const mainValues = main.flatMap(([s]) => s.values);
    const stackTotals = spec.grouping === "stacked"
      ? plan.categories.map((_, ci) => main.reduce((a, [s]) => a + Math.abs(s.values[ci] || 0), 0)) : [];
    const valFmt = pct ? "0%" : axisFormat(main[0] ? main[0][0].fmt : series[0].fmt, [...mainValues, ...stackTotals]);
    const A1 = 50010, A2 = 50020, A3 = 50030, A4 = 50040;
    let groups = "";
    if (kind === "bar" || kind === "combo") {
      const grouping = spec.grouping || "clustered";
      groups += `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>` +
        main.map(([s, i]) => chartSeriesXml("bar", s, i, plan, refs.series[i])).join("") +
        `<c:gapWidth val="60"/>${grouping !== "clustered" ? '<c:overlap val="100"/>' : ""}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:barChart>`;
    } else if (kind === "line") {
      groups += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
        main.map(([s, i]) => chartSeriesXml("line", s, i, plan, refs.series[i])).join("") +
        `<c:marker val="1"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:lineChart>`;
    } else if (kind === "area") {
      groups += `<c:areaChart><c:grouping val="${spec.grouping || "standard"}"/><c:varyColors val="0"/>` +
        main.map(([s, i]) => chartSeriesXml("area", s, i, plan, refs.series[i])).join("") +
        `<c:axId val="${A1}"/><c:axId val="${A2}"/></c:areaChart>`;
    }
    // A funnel hides its value axis and gridlines: the offsets would make the numbers misleading
    const titles = axisTitles(plan, main, lines);
    let axes = catAxXml(A1, A2, horizontal ? "l" : "b", { reverse: horizontal, title: titles.cat, hidden: !titles.catShow, count: plan.categories.length }) +
      valAxXml(A2, A1, horizontal ? "b" : "l", valFmt, { crosses: horizontal ? "max" : "autoZero", grid: !spec.hideValueAxis,
        hidden: !!spec.hideValueAxis || !titles.valShow, title: spec.hideValueAxis ? null : titles.val,
        zero: (kind === "bar" || kind === "combo") && main.every(([s]) => s.values.every((v) => v === null || v >= 0)) });
    if (lines.length) {
      groups += `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
        lines.map(([s, i]) => chartSeriesXml("line", s, i, plan, refs.series[i])).join("") +
        `<c:marker val="1"/><c:axId val="${A3}"/><c:axId val="${A4}"/></c:lineChart>`;
      const lineFmt = axisFormat(lines[0][0].fmt, lines.flatMap(([s]) => s.values));
      axes += catAxXml(A3, A4, "b", { hidden: true }) + valAxXml(A4, A3, "r", lineFmt, { crosses: "max", grid: false, hidden: !titles.secShow, title: titles.sec });
    }
    body = groups + axes;
  }
  const round = kind === "pie" || kind === "doughnut";
  // Legend: one entry per visible series (or per scatter group); helper series are left out
  const entries = scatter ? (plan.perPoint ? 0 : plan.groups.length) : series.filter((s) => !s.noLegend).length;
  const hiddenEntries = scatter ? "" : series.map((s, i) => (s.noLegend ? `<c:legendEntry><c:idx val="${i}"/><c:delete val="1"/></c:legendEntry>` : "")).join("");
  const legend = entries > 1 || round
    ? `<c:legend><c:legendPos val="${round ? "r" : "t"}"/>${hiddenEntries}<c:overlay val="0"/></c:legend>` : "";
  const title = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="1200" b="1">${fillXml(C.INK)}</a:rPr><a:t>${xmlEsc(plan.title)}</a:t></a:r></a:p></c:rich></c:tx>` +
    `<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${body}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${fillXml(C.WHITE)}<a:ln w="9525">${fillXml(C.HAIR)}</a:ln></c:spPr>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900">${fillXml(C.SLATE)}<a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` +
    `</c:chartSpace>`;
}

/* A chart's frame on the page sheet: its top-left cell and its size */
function chartAnchorXml(ch, rid, id) {
  const cx = Math.round(ch.wPx * CHART_PX_TO_EMU), cy = Math.round(ch.hPx * CHART_PX_TO_EMU);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
}

/* ── Adding the charts to the finished .xlsx ── */
/* ── Funnel charts: Excel's own funnel ───────────────────────────────────
 * Excel 2019 and Microsoft 365 have a real funnel chart ("chartEx", a newer chart format).
 * A funnel visual is written as one, with the centred-bar funnel as its fallback: Excel
 * versions without funnel charts, Excel on older Macs, LibreOffice and Google Sheets show the
 * bars instead, so every reader sees a funnel. NATIVE_FUNNEL = false keeps the bars only.
 */
const NATIVE_FUNNEL = false;
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${CHARTEX_NS}">` +
    `<cx:chartData><cx:data id="0">` +
    `<cx:strDim type="cat"><cx:f>${xmlEsc(plan.refs.catRef)}</cx:f><cx:lvl ptCount="${cats.length}">${pts(cats)}</cx:lvl></cx:strDim>` +
    `<cx:numDim type="val"><cx:f>${xmlEsc(ref.valRef)}</cx:f><cx:lvl ptCount="${s.values.length}" formatCode="${xmlEsc(s.fmt || "General")}">${pts(s.values)}</cx:lvl></cx:numDim>` +
    `</cx:data></cx:chartData>` +
    `<cx:chart>` +
    (plan.title ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${font(C.INK, 1200, true)}</a:pPr>` +
      `<a:r><a:rPr lang="en-US" sz="1200" b="1">${fillXml(C.INK)}<a:latin typeface="Calibri"/></a:rPr><a:t>${xmlEsc(plan.title)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>` : "") +
    `<cx:plotArea><cx:plotAreaRegion>` +
    `<cx:series layoutId="funnel" uniqueId="{6F7B2C1A-3E4D-4B5A-9C8D-1A2B3C4D5E6F}">` +
    `<cx:tx><cx:txData><cx:f>${xmlEsc(ref.nameRef)}</cx:f><cx:v>${xmlEsc(s.name)}</cx:v></cx:txData></cx:tx>` +
    `<cx:spPr>${fillXml(main)}</cx:spPr>${dataPts}` +
    `<cx:dataLabels pos="ctr"><cx:numFmt formatCode="${xmlEsc(s.fmt || "General")}" sourceLinked="0"/>${txPr(textOn(main), 900, true)}` +
    `<cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>` +
    `<cx:dataId val="0"/></cx:series>` +
    `</cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.06"/><cx:tickLabels/>${txPr(C.SLATE, 900, false)}</cx:axis></cx:plotArea>` +
    `</cx:chart>` +
    `<cx:spPr>${fillXml(C.WHITE)}<a:ln w="9525">${fillXml(C.HAIR)}</a:ln></cx:spPr>` +
    `</cx:chartSpace>`;
}

/* A chart frame that shows Excel's own funnel where Excel has it, else the bar funnel */
function chartAltAnchorXml(ch, exRid, rid, id) {
  const cx = Math.round(ch.wPx * CHART_PX_TO_EMU), cy = Math.round(ch.hPx * CHART_PX_TO_EMU);
  const frame = (uri, inner) => `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></xdr:graphicFrame>`;
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" Requires="cx2">` +
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
    for (const ch of charts) {
      chartNo++;
      const chartPath = `xl/charts/chart${chartNo}.xml`;
      zip.file(chartPath, chartXml(ch.plan));
      addType(chartPath, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
      const rel = addRel(drawingRels, "chart", `../charts/chart${chartNo}.xml`);
      drawingRels = rel.xml;
            // A funnel: Excel's own funnel chart too, with the bar funnel as its fallback
      if (NATIVE_FUNNEL && ch.plan.funnel) {
        exNo++;
        const exPath = `xl/charts/chartEx${exNo}.xml`;
        zip.file(exPath, funnelChartExXml(ch.plan));
        addType(exPath, CHARTEX_TYPE);
        const exRel = addRel(drawingRels, CHARTEX_REL, `../charts/chartEx${exNo}.xml`);
        drawingRels = exRel.xml;
        anchors += chartAltAnchorXml(ch, exRel.id, rel.id, 1000 + chartNo);
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
  // The drill levels as they are in the app now, before the export switches pages
  const liveDrill = options.matchColors ? await captureLiveDrill(embeddedReport) : null;
  try {
    for (let p = 0; p < pages.length; p++) {
      const page = pages[p];
      onProgress({ pageIndex: p + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: 0, visualCount: 0, visualTitle: "Opening page…" });
      await activatePage(page);
      const data = await collectPage(page, options,
        (i, n, title) => onProgress({ pageIndex: p + 1, pageCount: pages.length, pageName: page.displayName, visualIndex: i, visualCount: n, visualTitle: title }),
        isCancelled);
      if (liveDrill && liveDrill[page.name]) data.items.forEach((item) => { item.liveDrill = liveDrill[page.name][item.name] || null; });
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

  // Each visual at the drill level the report shows, then table sparklines (both need the report file)
  if (options.reportColors) {
    onProgress({ pageIndex: pages.length, pageCount: pages.length, pageName: pages[pages.length - 1].displayName, visualIndex: 0, visualCount: 0, visualTitle: "Matching drill levels…" });
    await loadDrillLevels(collected, options);
    onProgress({ pageIndex: pages.length, pageCount: pages.length, pageName: pages[pages.length - 1].displayName, visualIndex: 0, visualCount: 0, visualTitle: "Reading sparklines…" });
    await loadSparklines(collected, options);
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
    totals.tables += counts.tables;
    totals.kpis += counts.kpis;
    data.notes.forEach((n) => totals.notes.push(sheets.length > 1 ? `${page.displayName}: ${n}` : n));
  });

  let buffer = await wb.xlsx.writeBuffer();
  // Charts mode: add the real Excel charts to the finished file
  if (chartSheets.length) buffer = await addChartsToFile(buffer, chartSheets);
  if (sparkSheets.length) buffer = await addSparklinesToFile(buffer, sparkSheets);
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
