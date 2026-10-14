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
const CONTENT_COL = 2;     // everything starts at column B; column A is the margin
const MARGIN_W = 3;
const GUTTER_W = 3;
const MIN_PAGE_W = 100;    // minimum content width (Excel width units)
const SNAP_PX = 40;        // stacks starting within this many page pixels share a column
const BAND_GAP = 2;        // blank rows between bands
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

  // A double donut: both rings coloured as the visual's code does it
  const custom = customVisual(item.type);
  if (custom && custom.rule === "rings") {
    const rings = ringColors({ data, cols, numeric, hex, findCol, palette, info });
    if (rings) return { ...rings, hexCols, exact: !!info };
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
  const capSource = accentBar ? toArgb(accentBar) : valueColor(0);
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
  const titleUnits = cellUnits(item.displayTitle, fonts.exTitle) + (number ? cellUnits(String(number).padStart(2, "0") + "  ", fonts.exNum) : 0) - GUTTER_W;
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
  if (item.tableFormat) applyTableFormats(ws, hr, col, item, header, data, lay);
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
function applyTableFormats(ws, hr, col, item, header, data, lay) {
  const info = item.tableFormat;
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
  const columnOf = (name) => {
    const n = normalizeField(String(name || "").toLowerCase());
    const k = values.findIndex((m) => m.ref === n);
    if (k >= 0 && valueCol[k] >= 0) return valueCol[k];
    return norm.indexOf(n);
  };
  const first = hr + 1, last = hr + data.length;
  if (!data.length) return;

  // One colour per row for a column: a fixed colour, or its fx rule (null = left as is)
  const column = (name) => {
    const ci = columnOf(name);
    return ci >= 0 ? data.map((r) => r[ci].value ?? null) : null;
  };
  const rowColors = (f) => {
    if (f.color) return data.map(() => toArgb(f.color));
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
      : [columnOf(f.target.slice(2))].filter((ci) => ci >= 0);
    for (const ci of targets) {
      const c = col + ci;
      const L = ws.getColumn(c).letter;
      if (f.kind === "back" || f.kind === "font") {
        const colors = rowColors(f);
        colors.forEach((argb, ri) => {
          if (!argb) return;
          const cell = ws.getCell(first + ri, c);
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
        ws.addConditionalFormatting({ ref: `${L}${first}:${L}${last}`, rules: [bar] });
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
          ref: `${L}${first}:${L}${last}`,
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

  // Series drawn in one colour: the header naming the series
  for (const [ci, argb] of Object.entries(colors.series)) {
    const h = ws.getCell(hr, col + Number(ci));
    h.fill = styleOf(argb).fill;
    h.font = styleOf(argb).headFont;
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
    const info = visualInfo(lay.reportColors, item.page, item.name);
    item.colors = lay.palette && !item.isKpi ? columnColors(item, lay.palette, info) : null;
    item.tableFormat = info && info.tableFormats && info.tableFormats.length ? info : null;
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

    // Slicers become filter chips: "Field  selection", or "Field  All". A hidden slicer still
    // filters the page (bookmarks often hide them), so it gets a chip too.
    if (isSlicerType(v.type)) {
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
      items.push({ title, name: v.name, page: page.name, type: v.type, layout: v.layout, rows, isKpi });
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
  // With several pages, one with nothing to show (a cover or info page of images and text)
  // isn't given an empty sheet; the first sheet's notes name it
  const filled = collected.filter(({ data }) => data.items.length || data.notes.length);
  const sheets = collected.length > 1 && filled.length ? filled : collected;
  const left = collected.filter((c) => !sheets.includes(c)).map(({ page }) => `"${page.displayName}"`);
  if (left.length) sheets[0].data.notes.push(`Not included: ${left.join(", ")} – no visuals with data.`);
  sheets.forEach(({ page, data }, i) => {
    const ctx = { pageIndex: i + 1, pageCount: sheets.length, dateText };
    const counts = writePageSheet(wb, nameSheet(page.displayName), page, data, options, branding, logo, nameTable, ctx);
    totals.tables += counts.tables;
    totals.kpis += counts.kpis;
    data.notes.forEach((n) => totals.notes.push(sheets.length > 1 ? `${page.displayName}: ${n}` : n));
  });

  const buffer = await wb.xlsx.writeBuffer();
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
