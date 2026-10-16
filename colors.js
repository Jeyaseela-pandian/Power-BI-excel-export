"use strict";

/*
 * Exact report colours. Downloads the report's design file (needs the Dataset.Read.All
 * permission and download rights: Contributor, Member or Admin of the workspace), reads
 * the report's theme and the colours set by hand on each visual, and describes how each
 * visual colours its data (by item, by legend value, by measure, or one colour), plus
 * two display facts the embed API doesn't give: titles turned off, and the field a gauge
 * or KPI shows as its main value.
 * Both report formats are read: the classic one ("Report/Layout") and the newer PBIR
 * folders ("Report/definition/…"). Only this design information is used; nothing is
 * stored or sent anywhere. When the file can't be read the export falls back to the theme
 * colours Power BI reports through the embed API (export.js).
 * Uses globals from app.js: getDesignToken, PBI_API, log, errorText.
 */

/* ── Power BI rules ───────────────────────────────────────────────────────
 * Power BI doesn't store these in reports, so they live in the code.
 *  - FALLBACK_THEME: colours of Power BI's classic theme, for old reports with no theme.
 *  - COLOR_BY_ITEM_TYPES: visuals that give each category value its own colour.
 *  - SINGLE_COLOR_TYPES: visuals drawn in one colour unless a legend is used.
 *  - Funnels draw every bar in one colour unless colours are set per bar.
 *  - CUSTOM_VISUALS: AppSource visuals whose colouring is known (their code isn't in the file).
 */
const FALLBACK_THEME = {
  name: "Power BI classic",
  dataColors: ["#01B8AA", "#374649", "#FD625E", "#F2C80F", "#5F6B6D", "#8AD4EB", "#FE9666", "#A66999"],
  background: "#FFFFFF",
  foreground: "#252423",
};
// Named theme colours used by conditional formatting, with Power BI's defaults
const THEME_KEYWORDS = { maxColor: "maximum", minColor: "minimum", midColor: "center", centerColor: "center",
  good: "good", bad: "bad", neutral: "neutral", nullColor: "null", foreground: "foreground", background: "background" };
const KEYWORD_DEFAULTS = { maximum: "#118DFF", minimum: "#DEEFFF", center: "#D9B300", good: "#1AAB40", bad: "#D64554",
  neutral: "#D9B300", null: "#FF7F48", foreground: "#252423", background: "#FFFFFF" };
const COLOR_BY_ITEM_TYPES = ["pieChart", "donutChart", "treemap"];
const SINGLE_COLOR_TYPES = ["scatterChart", "map", "filledMap", "shapeMap", "azureMap", "funnel"];
const MEASURE_ROLES = ["Y", "Y2", "Values", "ColumnY", "LineY", "Size"];   // roles whose fields are drawn as series
const KPI_VALUE_ROLE = { gauge: "Y", kpi: "Indicator" };   // the main value; the other fields are its target, minimum…

/* Supported AppSource visuals, by type name (their defaults are taken from their code):
   - roles: each of the visual's own roles, as the built-in role it acts like
   - colors: the visual's object that holds colours set per category (a "fill")
   - settings: { key: [object, property, default] } other settings read from the report
   - rule: how the export shows it: "rings" (double donut), "nested" (outer / inner bars),
     "table" (a table, with the visual's cell bars as data bars), "kpi" (a KPI tile; mainRole
     is its main value), "slicer" (a filter chip), "none" (no data: skipped quietly)
   - layout: how its rows are rebuilt to read like the visual: "grid" (comparison table),
     "bins" (histogram), "trend" (trend cards with sparklines above the table: Dual KPI) */
const CUSTOM_VISUALS = {
  // NeenOpal's Double Donut Chart: categories on the inner ring, their subcategories on the outer ring
  doubleDonut2B2DE03D219E04DFF80684A45ECA4D25E: {
    name: "Double donut chart",
    rule: "rings",
    roles: { category: "Category", subcategory: "Category", value: "Values" },
    colors: "categoryColors",
  },
  // NeenOpal's Nested Bars: a bar per category with its selected subset's bar inside
  nestedBars578293FC1D324F5CA7440FB9344A25E3: {
    name: "Nested bars",
    rule: "nested",
    roles: { category: "Category", subset: "Subset", categoryValue: "Y", subsetValue: "Y2", tooltip: "Tooltips" },
    settings: {
      outer: ["columns", "outerColumnColor", "#12239E"],
      inner: ["columns", "innerColumnColor", "#118DFF"],
      autoInner: ["columns", "autoGenerateSubsetColor", true],   // the subset colour is the outer one's complement
      subset: ["subsetSelection", "selectedSubset", null],        // the subset drawn inside (else the first one)
    },
  },
  // NeenOpal's Comparison Table: metrics against regions; benchmark cells hold a blue bar,
  // value cells a green bar (red when negative)
  comparisonTable97CE315D158843019602B15CF4A4063F: {
    name: "Comparison table",
    rule: "table", layout: "grid",   // a row per metric, a column per region (comparisonRows in export.js)
    roles: { metrics: "Rows", regions: "Columns", benchmark: "Values", values: "Values" },
    bars: { benchmark: { color: "#2196F3" }, values: { color: "#4CAF50", negative: "#F44336" } },
    settings: {
      headerBack: ["columnHeaders", "headersBackgroundColor", null],
      headerText: ["columnHeaders", "headersTextColor", null],
    },
  },
  // NeenOpal's Progress Icon: a value against a benchmark, drawn as a filling icon
  iconwithgraph7E9A1A432F2E4C9099F64BB88C889CA4: {
    name: "Progress icon",
    rule: "kpi",
    roles: { Value: "Y", BenchmarkValue: "TargetValue" },
    mainRole: "Y",
    settings: { accent: ["iconSettings", "iconFillColor", "#2E7D32"] },
  },
  // NeenOpal's Date Range Picker: filters the page by a date range
  datePicker7A12AE28E48C4004852BA7D1536C378E: { name: "Date range picker", rule: "slicer", roles: { DateField: "Values" } },
  // ── Microsoft's open-source AppSource visuals (rules read from their GitHub source) ──
  // Bullet Chart: fixed band colours, the value bar black
  BulletChart1443347686880: {
    name: "Bullet chart", rule: "fixed",
    roles: { Category: "Category", Value: "Value", TargetValue: "TargetValue", Minimum: "Minimum", NeedsImprovement: "NeedsImprovement",
      Satisfactory: "Satisfactory", Good: "Good", VeryGood: "VeryGood", Maximum: "Maximum", TargetValue2: "TargetValue2" },
    seriesColors: { Value: ["bullet"], Minimum: ["min"], NeedsImprovement: ["needs"], Satisfactory: ["satisfactory"], Good: ["good"], VeryGood: ["veryGood"] },
    settings: { bullet: ["colors", "bulletColor", "#000000"], min: ["colors", "minColor", "#8B0000"], needs: ["colors", "needsImprovementColor", "#FF0000"],
      satisfactory: ["colors", "satisfactoryColor", "#FFFF00"], good: ["colors", "goodColor", "#008000"], veryGood: ["colors", "veryGoodColor", "#006400"] },
  },
  // Gantt: a theme colour per Legend value (colours set under Legend win); one task colour without a legend
  Gantt1448688115699: {
    name: "Gantt chart", rule: "auto",
    roles: { Legend: "Series", Task: "Category", Parent: "Parent", StartDate: "StartDate", EndDate: "Y", Duration: "Y", Completion: "Completion",
      Resource: "Resource", ExtraInformation: "Tooltips", Milestones: "Milestones" },
    colors: "legend",
    settings: { all: ["taskConfig", "fill", "#00B099"] },
  },
  // Mekko: a theme colour per Series value; one colour without a series
  MekkoChart1449744733038: {
    name: "Mekko chart", rule: "auto",
    roles: { Category: "Category", Series: "Series", Y: "Y", Width: "Width", CategoryCustomSorting: "Sorting" },
    settings: { all: ["dataPoint", "defaultColor", "#01B8AA"] },
  },
  // Stream Graph: a theme colour per group (stream), in order; per measure without a group
  StreamGraph1446659696222: {
    name: "Stream graph", rule: "auto",
    roles: { Category: "Category", Series: "Series", Y: "Y" },
    colors: "streams",
  },
  // Histogram: every bar one colour
  Histogram1445664487616: {
    name: "Histogram", rule: "fixed", layout: "bins",   // its bins, as the visual draws them (histogramRows in export.js)
    roles: { Values: "Category", Frequency: "Frequency" },
    seriesColors: { Frequency: ["fill"] },
    settings: { fill: ["dataPoint", "fill", "#01B8AA"], bins: ["general", "bins", null] },
  },
  // Dual KPI: top and bottom charts in fixed colours (the bottom matches the top by default)
  PBI_CV_3C80B1F2_09AF_4123_8E99_C3CBC46B23E0: {
    name: "Dual KPI", rule: "fixed", layout: "trend",   // two trend cards (dualKpiTrend in export.js)
    roles: { axis: "Category", topvalues: "topvalues", bottomvalues: "bottomvalues", warningstate: "warningstate",
      toppercentdate: "toppercentdate", bottompercentdate: "bottompercentdate" },
    seriesColors: { topvalues: ["top"], bottomvalues: ["bottom", "top"] },
    settings: {
      top: ["dualKpiColors", "dataColor", "#01B8AA"], bottom: ["dualKpiColorsBottom", "dataColor", null],
      matchBottom: ["dualKpiColorsBottom", "matchTopChartOptions", true],
      showTop: ["dualKpiProperties", "topChartShow", true], showBottom: ["dualKpiProperties", "bottomChartShow", true],
      titleText: ["dualKpiProperties", "titleText", ""],
      abbreviate: ["dualKpiProperties", "abbreviateValues", false],
      topStart: ["dualKpiProperties", "topPercentCalcDate", ""], bottomStart: ["dualKpiProperties", "bottomPercentCalcDate", ""],
      warningText: ["dualKpiProperties", "warningTooltipText", ""],
      staleWarning: ["dualKpiProperties", "showStaleDataWarning", true], staleDays: ["dualKpiProperties", "staleDataThreshold", 2],
      staleText: ["dualKpiProperties", "staleDataTooltipText", ""],
      showTopValue: ["dualKpiValues", "showKpiValuesTop", true], showBottomValue: ["dualKpiValues", "showKpiValuesBottom", true],
      topDefault: ["dualKpiValues", "topChartDefaultKpiValue", "N/A"], bottomDefault: ["dualKpiValues", "bottomChartDefaultKpiValue", "N/A"],
      units: ["dualKpiValueFormatting", "displayUnits", 1], precision: ["dualKpiValueFormatting", "precision", 0],
      topMin: ["dualKpiAxis", "topChartAxisMin", null], topMax: ["dualKpiAxis", "topChartAxisMax", null],
      bottomMin: ["dualKpiAxis", "bottomChartAxisMin", null], bottomMax: ["dualKpiAxis", "bottomChartAxisMax", null],
    },
  },
  // Sankey: a theme colour per node name (sources first, then destination-only nodes; the
  // visual's code skips the 2nd palette slot); links grey
  sankey02300D1BE6F5427989F3DE31CCA9E0F32020: {
    name: "Sankey diagram", rule: "nodes",
    roles: { Source: "Category", Destination: "Destination", Weight: "Y", SourceLabels: "SourceLabels" },
    colors: "nodes",
  },
  // Word Cloud: a theme colour per word (row order); one colour for all words when set
  WordCloud1447959067750: {
    name: "Word cloud", rule: "item",
    roles: { Category: "Category", Values: "Values", Excludes: "Excludes" },
    settings: { all: ["dataPoint", "defaultColor", null] },
  },
  // Sunburst: a theme colour per top-level group; inner rings inherit their group's colour
  Sunburst1445472000808: { name: "Sunburst", rule: "item", roles: { Nodes: "Category", Values: "Values" }, colors: "group" },
  // Aster Plot: a theme colour per slice (row order)
  AsterPlot1443303142064: { name: "Aster plot", rule: "item", roles: { Category: "Category", Y: "Y" }, colors: "pies" },
  // Radar Chart: a theme colour per measure (polygon)
  RadarChart1446119667547: { name: "Radar chart", rule: "measure", roles: { Category: "Category", Y: "Y" } },
  // Tornado: one fixed colour per side, purple then teal (the AppSource build; its newest source
  // switches to theme colours), whether the sides are legend values or two measures
  TornadoChart1452517688218: { name: "Tornado chart", rule: "auto", roles: { Category: "Category", Series: "Series", Values: "Values" },
    palette: ["#800080", "#008080"] },
  // Filter visuals: filter chips
  Timeline1447991079100: { name: "Timeline", rule: "slicer", roles: { Time: "Values" } },
  textFilter25A4896A83E0487089E2B90C9AE57C8A: { name: "Text filter", rule: "slicer", roles: { field: "Values" } },
  AttributeSlicer1652434005853: { name: "Attribute slicer", rule: "slicer", roles: { Category: "Values" } },
  // Microsoft's Power Automate button: no data
  FlowVisual_C29F1DCC_81F5_4973_94AD_0517D44CC06A: { name: "Power Automate button", rule: "none", roles: {} },
};

/* A supported visual's setting: from the report file when it sets one, else the visual's default */
function customSetting(type, info, key) {
  const c = customVisual(type);
  if (!c || !c.settings || !c.settings[key]) return null;
  const set = info && info.settings;
  return set && key in set ? set[key] : c.settings[key][2];
}

/* An AppSource (custom) visual's type name: a name followed by an id (32 hex digits, a GUID
   written with underscores, or a long number such as "SankeyDiagram1446463184954") */
function isCustomType(type) {
  const t = String(type || "");
  return /[0-9A-F]{32}$/i.test(t) || /_[0-9A-F]{8}_[0-9A-F]{4}_/i.test(t) || /[A-Za-z]\d{12,}$/.test(t);
}

/* A supported AppSource visual's description, or null */
function customVisual(type) {
  const t = String(type || "").toLowerCase();
  const key = Object.keys(CUSTOM_VISUALS).find((k) => k.toLowerCase() === t);
  return key ? CUSTOM_VISUALS[key] : null;
}

/* ── Small helpers ────────────────────────────────────────────────────── */

/* Mixes a colour towards another (amount 0–1): Power BI's lighter / darker theme shades */
function mixHex(hex, target, amount) {
  const a = parseInt(String(hex).replace("#", ""), 16);
  const b = parseInt(String(target).replace("#", ""), 16);
  const ch = (shift) => {
    const x = (a >> shift) & 255, y = (b >> shift) & 255;
    return Math.round(x + (y - x) * amount).toString(16).padStart(2, "0");
  };
  return "#" + (ch(16) + ch(8) + ch(0)).toUpperCase();
}

/* Field names: "Sum(Sales.This Year Sales)", "'Stores'[Store Type]", "Sum of Sales" → "this year sales", "store type", "sales".
   Only an aggregation around a table.field reference is unwrapped, so "Revenue(USD)" stays whole. */
function normalizeField(text) {
  let s = String(text ?? "").trim();
  const agg = s.match(/^[A-Za-z]+\(([^()]*\.[^()]*)\)$/);
  if (agg) s = agg[1];
  const col = s.match(/\[([^\]]+)\]\s*$/);
  if (col) s = col[1];
  else if (s.includes(".")) s = s.slice(s.lastIndexOf(".") + 1);
  return s.replace(/^(sum|average|count|min|max|count \(distinct\)) of /i, "").trim().toLowerCase();
}

/* Item values: "'Fashions Direct'" → "fashions direct", "'O''Brien'" → "o'brien", "2024L" → "2024",
   "datetime'2024-01-01T00:00:00'" → "2024-01-01t00:00:00" */
function normalizeValue(text) {
  let s = String(text ?? "").trim().replace(/^datetime(?=')/i, "");
  const quoted = s.match(/^'(.*)'$/);
  if (quoted) s = quoted[1].replace(/''/g, "'");
  return s.replace(/^(-?\d+(?:\.\d+)?)[LDM]$/, "$1").toLowerCase();
}

/* A colour set on a visual: "#01B8AA", or a theme colour with an optional lighter / darker
   shade. Theme colour ids: 0 = white and 1 = black (the picker's first two columns, the same
   in every theme: "white, 20% darker" is #CCCCCC even in a dark theme), 2 onwards = data colours.
   Rules (conditional formatting, "fx") return null here: see parseColorRule. */
function colorFromFill(fill, theme) {
  return colorFromExpr(fill && fill.solid && fill.solid.color && fill.solid.color.expr, theme);
}

function colorFromExpr(e, theme) {
  if (!e) return null;
  if (e.Literal) {
    const hex = String(e.Literal.Value).replace(/'/g, "");
    if (/^#[0-9A-F]{6}$/i.test(hex)) return hex.toUpperCase();
    const key = THEME_KEYWORDS[hex];   // "maxColor" → the theme's maximum colour …
    const named = key && ((theme.keywords && theme.keywords[key]) || KEYWORD_DEFAULTS[key]);
    return named ? String(named).toUpperCase() : null;
  }
  if (e.ThemeDataColor) {
    const id = e.ThemeDataColor.ColorId, pct = e.ThemeDataColor.Percent || 0;
    const base = id === 0 ? "#FFFFFF"
      : id === 1 ? "#000000"
        : theme.dataColors[(id - 2) % theme.dataColors.length];
    if (!base) return null;
    return pct ? mixHex(base, pct < 0 ? "#000000" : "#FFFFFF", Math.abs(pct)) : base.toUpperCase();
  }
  return null;
}

/* ── Conditional formatting ("fx") colour rules ───────────────────────────
 * Kept as data and evaluated per row by evaluateColorRule, because the colour depends on
 * each data point's values:
 *  { kind: "gradient", input, min, mid, max, nulls }   stops: { color, at: threshold|null }
 *  { kind: "conditional", input, cases: [{ test, color }], otherwise, nulls }
 *  { kind: "field", input }                             the field's values are colours
 * input is the name of the measure / column the rule reads, as the report spells it.
 * A threshold is { num }, { pct } (a fraction of the range from lowest to highest value),
 * { pctl } (a percentile) or { str } (a text value).
 */
function refName(x) {
  if (!x) return null;
  const inner = x.Measure || x.Column || (x.Aggregation && (x.Aggregation.Expression.Column || x.Aggregation.Expression.Measure));
  if (inner && inner.Property) return String(inner.Property);
  if (x.HierarchyLevel && x.HierarchyLevel.Level) return String(x.HierarchyLevel.Level);
  if (x.SelectRef && x.SelectRef.ExpressionName) return String(x.SelectRef.ExpressionName).split(".").pop();   // "Sales.Profit"
  return null;
}

function literalValue(x) {
  if (!x || !x.Literal) return null;
  const v = String(x.Literal.Value);
  const num = v.match(/^(-?\d+(?:\.\d+)?(?:E[+-]?\d+)?)[LDM]?$/i);
  if (num) return Number(num[1]);
  return normalizeValue(v);
}

function parseThreshold(x) {
  if (!x) return null;
  if (x.Literal) {
    const v = literalValue(x);
    return typeof v === "number" ? { num: v } : v === null ? null : { str: v };
  }
  if (x.RangePercent && x.RangePercent.Percent !== undefined) return { pct: Number(x.RangePercent.Percent) };
  if (x.Percentile && x.Percentile.Percent !== undefined) return { pctl: Number(x.Percentile.Percent) };
  return null;
}

/* ComparisonKind: 0 =, 1 >, 2 >=, 3 <=, 4 < */
function parseCondition(c) {
  if (!c) return null;
  if (c.Comparison) {
    const input = refName(c.Comparison.Left);
    const at = parseThreshold(c.Comparison.Right);
    return input && at ? { input, op: c.Comparison.ComparisonKind, at } : null;
  }
  for (const [k, kind] of [["And", "and"], ["Or", "or"]]) {
    if (c[k]) {
      const a = parseCondition(c[k].Left), b = parseCondition(c[k].Right);
      return a && b ? { [kind]: [a, b], input: a.input } : null;
    }
  }
  if (c.Not) {
    const a = parseCondition(c.Not.Expression);
    return a ? { not: a, input: a.input } : null;
  }
  return null;
}

/* How blanks are coloured: as zero (Power BI's default), not at all, or in a set colour */
function parseNulls(fr, theme) {
  const n = fr && fr.nullColoringStrategy;
  const mode = n && n.strategy && n.strategy.Literal ? String(n.strategy.Literal.Value).replace(/'/g, "") : "asZero";
  return { mode, color: n && n.color ? colorFromExpr(n.color.Literal ? n.color : n.color.solid ? n.color.solid.color.expr : n.color, theme) : null };
}

function parseColorRule(fill, theme) {
  const e = fill && fill.solid && fill.solid.color && fill.solid.color.expr;
  if (!e) return null;
  if (e.FillRule) {
    const fr = e.FillRule.FillRule || {};
    const g = fr.linearGradient3 || fr.linearGradient2;
    const input = refName(e.FillRule.Input);
    if (!g || !input) return null;
    const stop = (st) => (st ? { color: colorFromExpr(st.color, theme), at: parseThreshold(st.value) } : null);
    const rule = { kind: "gradient", input, min: stop(g.min), mid: stop(g.mid), max: stop(g.max), nulls: parseNulls(g, theme) };
    return rule.min && rule.min.color && rule.max && rule.max.color ? rule : null;
  }
  if (e.Conditional) {
    const cases = (e.Conditional.Cases || []).map((c) => ({ test: parseCondition(c.Condition), color: colorFromExpr(c.Value, theme) }))
      .filter((c) => c.test && c.color);
    if (!cases.length) return null;
    const otherwise = e.Conditional.DefaultValue ? colorFromExpr(e.Conditional.DefaultValue, theme) : null;
    return { kind: "conditional", input: cases[0].test.input, cases, otherwise };
  }
  // "Format by: Field value": a measure or column whose values are the colours
  const field = refName(e);
  return field ? { kind: "field", input: field } : null;
}

/*
 * Colours an fx rule gives each row. column(name) returns the field's value in every row
 * (numbers, text or null for blank), or null when the field isn't in the exported data.
 * Returns { colors: [hex | null] (null = no rule colour), missing: [field names not found] }.
 */
function evaluateColorRule(rule, column, rowCount) {
  const missing = [];
  const get = (name) => { const v = column(name); if (!v && !missing.includes(name)) missing.push(name); return v; };
  const none = () => ({ colors: new Array(rowCount).fill(null), missing });
  const numbers = (v) => v.map((x) => (typeof x === "number" ? x : null));
  const stats = (nums) => {
    const present = nums.filter((x) => x !== null).sort((a, b) => a - b);
    return { min: present[0] ?? 0, max: present[present.length - 1] ?? 0, sorted: present };
  };
  const level = (at, st) => {
    if (!at) return null;
    if (at.num !== undefined) return at.num;
    if (at.pct !== undefined) return st.min + (st.max - st.min) * at.pct;
    if (at.pctl !== undefined && st.sorted.length) return st.sorted[Math.min(st.sorted.length - 1, Math.floor(at.pctl * st.sorted.length))];
    return null;
  };

  if (rule.kind === "field") {
    const v = get(rule.input);
    if (!v) return none();
    return { colors: v.map((x) => { const t = String(x ?? "").trim(); return /^#[0-9A-F]{6}$/i.test(t) ? t.toUpperCase() : null; }), missing };
  }

  if (rule.kind === "gradient") {
    const v = get(rule.input);
    if (!v) return none();
    const nulls = rule.nulls || { mode: "asZero" };
    const nums = numbers(v).map((x) => (x === null && nulls.mode === "asZero" ? 0 : x));
    const st = stats(nums);
    const lo = level(rule.min.at, st) ?? st.min, hi = level(rule.max.at, st) ?? st.max;
    const hasMid = rule.mid && rule.mid.color;
    const mid = hasMid ? (level(rule.mid.at, st) ?? (lo + hi) / 2) : null;
    const part = (x, a, b) => (b === a ? 1 : Math.max(0, Math.min(1, (x - a) / (b - a))));
    return {
      colors: nums.map((x) => {
        if (x === null) return nulls.mode === "specificColor" ? nulls.color : null;
        if (hasMid) return x <= mid ? mixHex(rule.min.color, rule.mid.color, part(x, lo, mid)) : mixHex(rule.mid.color, rule.max.color, part(x, mid, hi));
        return mixHex(rule.min.color, rule.max.color, part(x, lo, hi));
      }),
      missing,
    };
  }

  // Conditional: the first case whose condition holds
  const cache = {};
  const fieldOf = (name) => {
    if (!(name in cache)) {
      const v = get(name);
      cache[name] = v ? { v, st: stats(numbers(v)) } : null;
    }
    return cache[name];
  };
  const test = (c, ri) => {
    if (c.and) return c.and.every((x) => test(x, ri));
    if (c.or) return c.or.some((x) => test(x, ri));
    if (c.not) return !test(c.not, ri);
    const f = fieldOf(c.input);
    if (!f) return false;
    const raw = f.v[ri];
    if (c.at.str !== undefined) {
      const eq = normalizeValue(raw) === c.at.str;
      return c.op === 0 ? eq : false;
    }
    const x = typeof raw === "number" ? raw : 0;   // blanks count as zero
    const y = level(c.at, f.st);
    if (y === null) return false;
    return [x === y, x > y, x >= y, x <= y, x < y][c.op] || false;
  };
  return {
    colors: Array.from({ length: rowCount }, (_, ri) => { const hit = rule.cases.find((c) => test(c.test, ri)); return hit ? hit.color : rule.otherwise || null; }),
    missing,
  };
}

/* Which data points a rule colours: every point / category ("all"), each legend value
   ("series"), or one measure ("m:<field>") */
function ruleScope(selector) {
  if (!selector) return "all";
  if (selector.metadata) return "m:" + normalizeField(selector.metadata);
  const roles = (selector.data || []).flatMap((d) => d.roles || []);
  return roles.includes("Series") ? "series" : "all";
}

/* Which data points a colour set by hand applies to:
   "all" (no selector, or every instance), "m:<measure>", or "i:<item value>[|…]" */
function selectorKey(selector) {
  if (!selector) return "all";
  if (selector.metadata) return "m:" + normalizeField(selector.metadata);
  const data = selector.data || [];
  if (data.length && data.every((d) => d.dataViewWildcard)) return "all";
  const values = [...JSON.stringify(data).matchAll(/"Literal":\{"Value":"((?:[^"\\]|\\.)*)"\}/g)]
    .map((m) => normalizeValue(JSON.parse('"' + m[1] + '"')));
  return values.length ? "i:" + values.join("|") : null;
}

/* ── Reading the design file ──────────────────────────────────────────── */
async function downloadDesignFile(report, signal) {
  const token = await getDesignToken();
  if (!token) throw new Error("NO_PERMISSION");
  const groupId = new URL(report.embedUrl).searchParams.get("groupId");
  const base = PBI_API + (groupId ? `/groups/${groupId}` : "") + `/reports/${report.id}/Export`;
  // "Design only" (no data) is small and quick; My workspace and some reports only offer the full file
  const urls = [base + "?downloadType=LiveConnect", base];
  let status = 0;
  for (const url of urls) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` }, signal });
    if (res.ok) return res.arrayBuffer();
    status = res.status;
    // A 401 after a redirect is the redirect losing the token, not a missing permission
    if (res.status === 401 && !res.redirected) throw new Error("NO_PERMISSION");
  }
  throw new Error(status === 403 ? "NO_ACCESS" : `DOWNLOAD_FAILED ${status}`);
}

async function readJson(zip, path, encoding) {
  const file = zip.file(path);
  if (!file) return null;
  const text = encoding === "utf-16le"
    ? new TextDecoder("utf-16le").decode(await file.async("uint8array"))
    : await file.async("string");
  return JSON.parse(text.replace(/^﻿/, ""));
}

/* The report's theme: the custom theme's values over the base theme's, as Power BI applies them */
async function readTheme(zip, themeCollection) {
  const files = Object.keys(zip.files).filter((n) => /^Report\/StaticResources\/.*\.json$/i.test(n));
  const load = async (t) => {
    const name = String((t && t.name) || "").toLowerCase();
    if (!name) return null;
    const path = files.find((n) => n.toLowerCase().endsWith("/" + name) || n.toLowerCase().endsWith("/" + name + ".json"));
    if (!path) return null;
    try { return await readJson(zip, path); } catch (e) { return null; }
  };
  const tc = themeCollection || {};
  const base = await load(tc.baseTheme);
  const custom = await load(tc.customTheme);
  const merged = { ...(base || {}), ...(custom || {}) };
  const dataColors = (merged.dataColors || []).filter((c) => /^#[0-9A-F]{6}$/i.test(c));
  if (!dataColors.length) return null;
  const keywords = {};
  for (const k of Object.values(THEME_KEYWORDS)) if (/^#[0-9A-F]{6}$/i.test(merged[k] || "")) keywords[k] = merged[k];
  return {
    name: (custom && custom.name) || (base && base.name) || "Report theme",
    dataColors,
    background: merged.background || FALLBACK_THEME.background,
    foreground: merged.foreground || FALLBACK_THEME.foreground,
    keywords,
    // Formatting defaults per visual type, base theme first, the custom theme over it
    visualStyles: [base && base.visualStyles, custom && custom.visualStyles].filter(Boolean),
  };
}

/*
 * One visual, in a shape shared by both formats:
 * { type, roles: { Category: [{ ref, label }], Series: [...], Y: [...] … }, dataPoint: [...], titleShown }
 * ref is the normalised field name; label the visual's own name for it when renamed.
 */
/*
 * A chart's axes as Power BI draws them: { cat, val, sec }, each with show (the axis) and title
 * (its title), true / false. Power BI resolves each setting in layers, and so does the export:
 * the visual's own default (axes shown, titles off), then the report's base theme, then its
 * custom theme (each for every visual, "*", then for this visual type), then the visual's own
 * setting. So a report on an older base theme shows axis titles unless a visual turns them off,
 * one on Fluent 2 or with no base theme doesn't unless a visual turns them on.
 */
const AXIS_DEFAULTS = { categoryAxis: { show: true, showAxisTitle: false }, valueAxis: { show: true, showAxisTitle: false, secShow: true } };
function axisSettings(objects, type, theme) {
  // The value set for this visual or in the themes (undefined when none sets it)
  const setting = (object, prop) => {
    const v = (((((objects || {})[object] || [])[0] || {}).properties || {})[prop] || {}).expr;
    const lit = v && v.Literal && String(v.Literal.Value);
    if (lit === "true" || lit === "false") return lit === "true";
    let value;
    for (const styles of (theme && theme.visualStyles) || []) {
      for (const key of ["*", type]) {
        const t = ((((styles || {})[key] || {})["*"] || {})[object] || [])[0];
        if (t && typeof t[prop] === "boolean") value = t[prop];
      }
    }
    return value;
  };
  const flag = (object, prop) => setting(object, prop) ?? (AXIS_DEFAULTS[object] || {})[prop] ?? null;
  const val = { show: flag("valueAxis", "show"), title: flag("valueAxis", "showAxisTitle") };
  return {
    cat: { show: flag("categoryAxis", "show"), title: flag("categoryAxis", "showAxisTitle") },
    val,
    // The line values' axis title follows the value axis's unless something sets it
    sec: { show: flag("valueAxis", "secShow"), title: setting("valueAxis", "secShowAxisTitle") ?? val.title },
  };
}

/* Whether a visual's title is shown, from its container settings: true / false when the
   report says, null when it leaves the default */
function titleShown(containerObjects) {
  const show = ((((containerObjects || {}).title || [])[0] || {}).properties || {}).show;
  const value = show && show.expr && show.expr.Literal && show.expr.Literal.Value;
  return value === "true" ? true : value === "false" ? false : null;
}

/* A query field's name: "Month", "Sum(Sales.Sales)" → "sales", a hierarchy level's name */
function fieldName(f) {
  if (!f) return null;
  if (f.Column && f.Column.Property) return f.Column.Property;
  if (f.Measure && f.Measure.Property) return f.Measure.Property;
  if (f.Aggregation && f.Aggregation.Expression) return fieldName(f.Aggregation.Expression);
  if (f.HierarchyLevel && f.HierarchyLevel.Level) return f.HierarchyLevel.Level;
  return null;
}

/*
 * What a query field is, for totals and queries: a column, a measure, an aggregated column
 * (agg: Power BI's aggregation code: 0 sum, 1 average, 2 distinct count, 3 min, 4 max,
 * 5 count, 6 median, 7 standard deviation, 8 variance), a hierarchy level or a table
 * sparkline (its measure and the fields it runs over); entity and prop name it in the model.
 * sources: the classic format's query aliases ("s" → "Sales").
 */
function fieldInfo(f, sources = {}) {
  if (!f) return null;
  const entityOf = (x) => {
    const ref = x && x.Expression && x.Expression.SourceRef;
    return ref ? ref.Entity || sources[ref.Source] || null : null;
  };
  if (f.Column) return { kind: "column", entity: entityOf(f.Column), prop: f.Column.Property };
  if (f.Measure) return { kind: "measure", entity: entityOf(f.Measure), prop: f.Measure.Property };
  if (f.Aggregation) {
    const inner = fieldInfo(f.Aggregation.Expression, sources);
    return inner ? { ...inner, kind: "agg", agg: Number(f.Aggregation.Function) } : null;
  }
  if (f.HierarchyLevel) {
    // A level of a hierarchy in a table: its column usually carries the level's name
    const h = f.HierarchyLevel.Expression && f.HierarchyLevel.Expression.Hierarchy;
    return { kind: "level", entity: (h && entityOf(h)) || null, prop: f.HierarchyLevel.Level };
  }
  if (f.SparklineData) {
    return { kind: "spark", measure: fieldInfo(f.SparklineData.Measure, sources),
      groupings: (f.SparklineData.Groupings || []).map((g) => fieldInfo(g, sources)) };
  }
  return null;
}

/*
 * The classic format names each projection by a query alias. Usually that is
 * "Table.Field", but it can be "select", "select1"… (hierarchy levels, fields added in
 * some ways), so the real field comes from the visual's query (prototypeQuery.Select).
 */
function legacyVisual(config) {
  const sv = config.singleVisual || {};
  const props = sv.columnProperties || {};
  const aliases = {}, fields = {};
  const sources = Object.fromEntries((((sv.prototypeQuery || {}).From) || []).map((f) => [f.Name, f.Entity]));
  for (const sel of ((sv.prototypeQuery || {}).Select || [])) {
    const name = fieldName(sel);
    if (sel.Name && name) aliases[sel.Name] = name;
    if (sel.Name) fields[sel.Name] = fieldInfo(sel, sources);
  }
  const roles = {};
  for (const [role, list] of Object.entries(sv.projections || {})) {
    roles[role] = (list || []).map((p) => ({
      ref: normalizeField(aliases[p.queryRef] || p.queryRef),
      label: props[p.queryRef] && props[p.queryRef].displayName ? String(props[p.queryRef].displayName).toLowerCase() : null,
      field: fields[p.queryRef] || null,
      queryRef: p.queryRef,
      active: p.active === true ? true : p.active === false ? false : null,   // the drill level shown
    }));
  }
  const o = sv.objects || {};
  return { type: sv.visualType, roles, objects: o, dataPoint: o.dataPoint || [], bubbles: o.bubbleLayer || [], values: [...(o.value || []), ...(o.labels || [])],
    titleShown: titleShown(sv.vcObjects), expansion: sv.expansionStates || null };
}

function pbirVisual(json) {
  const v = json.visual || {};
  const roles = {};
  for (const [role, state] of Object.entries((v.query && v.query.queryState) || {})) {
    roles[role] = ((state && state.projections) || []).map((p) => ({
      ref: normalizeField(fieldName(p.field) || p.queryRef || p.nativeQueryRef),
      label: p.displayName ? String(p.displayName).toLowerCase() : null,
      field: fieldInfo(p.field),
      queryRef: p.queryRef || null,
      active: p.active === true ? true : p.active === false ? false : null,   // the drill level shown
    }));
  }
  const o = v.objects || {};
  return { type: v.visualType, roles, objects: o, dataPoint: o.dataPoint || [], bubbles: o.bubbleLayer || [], values: [...(o.value || []), ...(o.labels || [])],
    titleShown: titleShown(v.visualContainerObjects), expansion: v.expansionStates || null };
}

/* Icon rules ("fx" icons): each case's lower bound (a threshold) and its icon */
function parseIconRule(icon) {
  const e = icon && icon.value && icon.value.expr && icon.value.expr.Conditional;
  if (!e) return null;
  const lower = (c) => {
    if (!c) return null;
    if (c.and) return lower(c.and[0]) || lower(c.and[1]);
    return c.op === 1 || c.op === 2 ? c.at : null;   // > or >=
  };
  const cases = (e.Cases || []).map((c) => ({ from: lower(parseCondition(c.Condition)), icon: String((c.Value && c.Value.Literal && c.Value.Literal.Value) || "").replace(/'/g, "") }))
    .map((c) => ({ ...c, from: c.from || { pct: 0 } }));
  return cases.length >= 2 ? { cases } : null;
}

/*
 * Formatting of a table or matrix visual: [{ target, kind, color | rule | bar | icons }]
 * target "m:<field>" (one column) or "all"; kind "back" | "font" (a colour or an fx rule),
 * "bar" (data bars) or "icon" (an icon rule).
 */
function tableFormats(objects, theme) {
  const out = [];
  for (const d of (objects && objects.values) || []) {
    const props = d.properties || {};
    const target = d.selector && d.selector.metadata ? "m:" + normalizeField(d.selector.metadata) : "all";
    for (const [prop, kind] of [["backColor", "back"], ["fontColor", "font"]]) {
      if (!props[prop]) continue;
      const color = colorFromFill(props[prop], theme);
      const rule = color ? null : parseColorRule(props[prop], theme);
      if (color || rule) out.push({ target, kind, color, rule });
    }
    if (props.icon) {
      const icons = parseIconRule(props.icon);
      if (icons) out.push({ target, kind: "icon", icons });
    }
  }
  // Column header colours set in the report
  for (const d of (objects && objects.columnHeaders) || []) {
    const props = d.properties || {};
    const back = colorFromFill(props.backColor, theme), font = colorFromFill(props.fontColor, theme);
    if (back) out.push({ target: "header", kind: "back", color: back });
    if (font) out.push({ target: "header", kind: "font", color: font });
  }
  // Formatting of one column ("Specific column"): its font / background colour, data bars
  const flag = (x) => !!(x && x.expr && x.expr.Literal && x.expr.Literal.Value === "true");
  for (const d of (objects && objects.columnFormatting) || []) {
    const props = d.properties || {};
    if (!d.selector || !d.selector.metadata) continue;
    const target = "m:" + normalizeField(d.selector.metadata);
    for (const [prop, kind] of [["backColor", "back"], ["fontColor", "font"]]) {
      if (!props[prop]) continue;
      const color = colorFromFill(props[prop], theme);
      const rule = color ? null : parseColorRule(props[prop], theme);
      if (color || rule) out.push({ target, kind, color, rule });
    }
    const bars = props.dataBars;
    const color = bars && colorFromFill(bars.positiveColor, theme);
    if (color) {
      out.push({ target, kind: "bar", bar: {
        color,
        negative: colorFromFill(bars.negativeColor, theme),
        axis: colorFromFill(bars.axisColor, theme),
        reverse: flag(bars.reverseDirection),
        showValue: !flag(bars.hideText),
      } });
    }
  }
  return out;
}

/* A custom table's formatting, as tableFormats describes a table's: the visual's cell bars
   as data bars on their columns, and header colours the report sets */
function customTableFormats(custom, ownRoles, settings) {
  const out = [];
  for (const [role, bar] of Object.entries(custom.bars || {})) {
    for (const f of ownRoles[role] || []) {
      out.push({ target: "m:" + f.ref, kind: "bar", bar: { color: bar.color, negative: bar.negative || null, axis: null, reverse: false, showValue: true } });
    }
  }
  if (settings.headerBack) out.push({ target: "header", kind: "back", color: settings.headerBack });
  if (settings.headerText) out.push({ target: "header", kind: "font", color: settings.headerText });
  return out;
}

/*
 * Settings of the built-in visuals the export rebuilds, at Power BI's defaults when the
 * report leaves them: a table's totals, a matrix's subtotals, a KPI's goal and status
 * colours (from the theme's good / neutral / bad), a new card's reference labels (each
 * belongs to one of the card's measures: owner, its index among the Data fields).
 */
function builtInSettings(visual, theme) {
  const o = visual.objects || {};
  const entry = (name) => (o[name] || []).find((e) => !e.selector) || (o[name] || [])[0] || {};
  const lit = (name, prop, dflt) => {
    const v = ((entry(name).properties || {})[prop] || {}).expr;
    if (!v || !v.Literal) return dflt;
    const text = String(v.Literal.Value);
    if (text === "true" || text === "false") return text === "true";
    const n = text.match(/^(-?\d+(?:\.\d+)?)[LDM]?$/i);
    return n ? Number(n[1]) : text.replace(/^'([\s\S]*)'$/, "$1").replace(/''/g, "'");
  };
  const color = (name, prop, dflt) => colorFromFill((entry(name).properties || {})[prop], theme) || dflt;
  const keyword = (k) => (theme && theme.keywords && theme.keywords[k]) || KEYWORD_DEFAULTS[k];
  switch (visual.type) {
    case "tableEx":
      return { totals: lit("total", "totals", true) };
    case "pivotTable":
      return { rowSubtotals: lit("subTotals", "rowSubtotals", true), columnSubtotals: lit("subTotals", "columnSubtotals", true) };
    case "kpi":
      return {
        direction: lit("status", "direction", "Positive"),
        good: color("status", "goodColor", keyword("good")),
        neutral: color("status", "neutralColor", keyword("neutral")),
        bad: color("status", "badColor", keyword("bad")),
        showGoal: lit("goals", "showGoal", true),
        goalText: lit("goals", "goalText", "Goal"),
        showDistance: lit("goals", "showDistance", true),
        distanceLabel: String(lit("goals", "distanceLabel", "Percent")),
        distanceDirection: lit("goals", "direction", null),
        showTrend: lit("trendline", "show", true),
        showLastDate: lit("lastDate", "show", false),
      };
    case "cardVisual": {
      const data = (visual.roles && visual.roles.Data) || [];
      const refLabels = (o.referenceLabel || []).map((e) => {
        const expr = e.properties && e.properties.value && e.properties.value.expr;
        const name = fieldName(expr);
        if (!name) return null;
        const sel = e.selector || {};
        let owner = data.findIndex((d) => d.queryRef && d.queryRef === sel.metadata);
        if (owner < 0 && data.length === 1) owner = 0;
        return { owner, ref: normalizeField(name), order: Number(sel.order) || 0 };
      }).filter(Boolean).sort((a, b) => a.owner - b.owner || a.order - b.order);
      return { refLabels };
    }
    default:
      return {};
  }
}

/* How a visual colours its data, with every colour set by hand resolved to a hex code */
function visualColorInfo(visual, theme) {
  const custom = customVisual(visual.type);
  const overrides = {};
  const rules = [];   // fx rules: { scope, rule }
  // Colours set per data point, or in a custom visual's own colour object
  const own = custom && custom.colors ? (visual.objects || {})[custom.colors] || [] : [];
  const points = [...visual.dataPoint, ...own].map((d) => [d, (d.properties || {}).fill]);
  const bubbles = visual.bubbles.map((d) => [d, (d.properties || {}).fillColor]);
  const strokes = ((visual.objects || {}).lineStyles || []).filter((d) => d.properties && d.properties.strokeColor)
    .map((d) => [d, d.properties.strokeColor]);
  // Where colours compete, the one drawn wins: a line's colour under "Lines" over its data
  // colour, an Azure Map's bubble colour over its leftover data point colour
  const entries = visual.type === "azureMap" ? [...bubbles, ...points] : [...strokes, ...points, ...bubbles];
  for (const [d, fill] of entries) {
    const props = d.properties || {};
    const key = selectorKey(d.selector);
    const color = colorFromFill(fill, theme);
    if (color && key && !(key in overrides)) overrides[key] = color;
    const rule = color ? null : parseColorRule(fill, theme);
    if (rule) rules.push({ scope: ruleScope(d.selector), rule });
    const fallback = colorFromFill(props.defaultColor, theme);
    if (fallback && !overrides.all) overrides.all = fallback;
  }
  // Card values: the new card's value font colour, or the classic card's label colour
  // (a fixed colour, or an fx rule); a gauge's value takes its arc colour
  const valueColors = {}, valueRules = {};
  for (const d of visual.values) {
    const props = d.properties || {};
    const fill = props.fontColor || props.color;
    const key = selectorKey(d.selector);
    if (!key) continue;
    const color = colorFromFill(fill, theme);
    if (color && !valueColors[key]) valueColors[key] = color;
    const rule = color ? null : parseColorRule(fill, theme);
    if (rule && !valueRules[key]) valueRules[key] = rule;
  }
  if (visual.type === "gauge" && overrides.all && !valueColors.all) valueColors.all = overrides.all;
  // The new card's accent bar and the multi-row card's bar: a fixed colour or a rule
  let accent = null;
  const o = visual.objects || {};
  for (const d of [...(o.accentBar || []), ...(o.card || [])]) {
    const props = d.properties || {};
    const fill = props.color || props.barColor;
    if (!fill || (props.show && props.show.expr && props.show.expr.Literal && props.show.expr.Literal.Value === "false")) continue;
    const color = colorFromFill(fill, theme);
    const rule = color ? null : parseColorRule(fill, theme);
    if (color || rule) { accent = { color, rule }; break; }
  }
  // Waterfall: increase / decrease / total colours (theme good / bad / first colour by default)
  const sentiment = ((visual.objects || {}).sentimentColors || [])[0];
  const sp = (sentiment && sentiment.properties) || {};
  const kw = (k) => (theme.keywords && theme.keywords[k]) || KEYWORD_DEFAULTS[k];
  const waterfall = visual.type === "waterfallChart" ? {
    increase: colorFromFill(sp.increaseFill, theme) || kw("good"),
    decrease: colorFromFill(sp.decreaseFill, theme) || kw("bad"),
    total: colorFromFill(sp.totalFill, theme) || theme.dataColors[0],
  } : null;
  // A custom visual's roles as the built-in roles they act like (category, then subcategory…)
  // An AppSource visual whose colouring isn't known: any colour it stores for one item
  // (whatever its object is called) counts as that item's colour
  const generic = !custom && isCustomType(visual.type) && !/^deneb/i.test(visual.type);
  if (generic) {
    for (const entries of Object.values(visual.objects || {})) {
      for (const d of entries || []) {
        const key = selectorKey(d.selector);
        if (!key || !key.startsWith("i:")) continue;
        for (const v of Object.values(d.properties || {})) {
          const c = colorFromFill(v, theme);
          if (c && !(key in overrides)) overrides[key] = c;
        }
      }
    }
  }
  // A custom visual's own settings, as the report sets them (colours resolved, true / false, text)
  const settings = {};
  for (const [k, [obj, prop]] of Object.entries((custom && custom.settings) || {})) {
    const v = ((((visual.objects || {})[obj] || [])[0] || {}).properties || {})[prop];
    if (!v) continue;
    const lit = v.expr && v.expr.Literal ? String(v.expr.Literal.Value) : null;
    const color = colorFromFill(v, theme);
    if (color) settings[k] = color;
    else if (lit !== null) settings[k] = lit === "true" ? true : lit === "false" ? false : lit.replace(/^'(.*)'$/, "$1");
  }
  const roles = custom ? {} : visual.roles;
  if (custom) for (const [name, role] of Object.entries(custom.roles)) roles[role] = [...(roles[role] || []), ...(visual.roles[name] || [])];
  const series = (roles.Series || [])[0] || null;
  const measures = MEASURE_ROLES.flatMap((r) => (roles[r] || []).map((m) => ({ ...m, role: r })));
  let rule = "measure";
  if (custom) rule = custom.rule === "auto" ? (series ? "legend" : "measure") : custom.rule;
  // One colour for every data point unless split by a legend (a custom visual's own default)
  if (custom && !series && !overrides.all && custom.settings && custom.settings.all) {
    const all = settings.all || custom.settings.all[2];
    if (all) overrides.all = all;
  }
  else if (generic) rule = "generic";
  else if (COLOR_BY_ITEM_TYPES.includes(visual.type)) rule = "item";
  else if (series) rule = "legend";
  else if (SINGLE_COLOR_TYPES.includes(visual.type)) rule = "single";
  return {
    rule,
    category: (roles.Category || roles.Group || [])[0] || null,   // treemaps use Group
    categories: roles.Category || roles.Group || [],               // a hierarchy: Group, then Region…
    series,
    measures,
    tooltips: (roles.Tooltips || []).map((t) => t.ref),
    overrides,
    rules,
    valueColors,
    valueRules,
    accent,
    waterfall,
    roles,   // every field well of the visual (X, Y, Size, Legend…), used by scatter charts
    tableFormats: ["tableEx", "pivotTable", "table", "matrix"].includes(visual.type) ? tableFormats(visual.objects, theme)
      : custom && custom.rule === "table" ? customTableFormats(custom, visual.roles, settings) : [],
    settings,
    titleHidden: visual.titleShown === false,
    axes: axisSettings(visual.objects, visual.type, theme),
    mainValue: (roles[(custom && custom.mainRole) || KPI_VALUE_ROLE[visual.type]] || [])[0] || null,
    builtin: builtInSettings(visual, theme),
    // A matrix's saved expand / collapse state: per role, each level's default and the groups toggled
    expansion: visual.expansion || null,
  };
}

/* Visuals are found by page and name: report authors often reuse names on other pages */
function visualKey(page, name) {
  return `${page || ""}\u0001${name}`;
}

/* Visuals whose name is used on only one page, for lookups without the page */
function uniqueByName(visuals) {
  const count = {}, out = {};
  for (const key of Object.keys(visuals)) { const name = key.split("\u0001")[1]; count[name] = (count[name] || 0) + 1; }
  for (const [key, info] of Object.entries(visuals)) { const name = key.split("\u0001")[1]; if (count[name] === 1) out[name] = info; }
  return out;
}

/* One visual's colour information in loaded report colours, or null */
function visualInfo(rc, pageName, visualName) {
  if (!rc || !visualName) return null;
  return rc.visuals[visualKey(pageName, visualName)] || (rc.byName && rc.byName[visualName]) || null;
}

/* The theme to use when the report names a theme the file doesn't hold: the colours
   Power BI reported through the embed API. Reports with no theme at all are classic. */
async function themeOrFallback(zip, themeCollection, embedTheme) {
  const theme = await readTheme(zip, themeCollection);
  if (theme) return theme;
  const named = themeCollection && (themeCollection.baseTheme || themeCollection.customTheme);
  return named && embedTheme ? { ...FALLBACK_THEME, ...embedTheme, name: "Report theme (from Power BI)" } : FALLBACK_THEME;
}

/* Reads the design file: { theme, visuals: { visualName: info }, format } */
async function readDesign(zip, embedTheme) {
  // Classic format: one UTF-16 JSON file with every page and visual
  const layout = await readJson(zip, "Report/Layout", "utf-16le").catch(() => null);
  if (layout) {
    let config = {};
    try { config = JSON.parse(layout.config || "{}"); } catch (e) { /* no theme */ }
    const theme = await themeOrFallback(zip, config.themeCollection, embedTheme);
    const visuals = {};
    for (const section of layout.sections || []) {
      for (const vc of section.visualContainers || []) {
        try {
          const c = JSON.parse(vc.config);
          if (c.name && c.singleVisual) visuals[visualKey(section.name, c.name)] = visualColorInfo(legacyVisual(c), theme);
        } catch (e) { /* skip an unreadable visual */ }
      }
    }
    return { theme, visuals, byName: uniqueByName(visuals), format: "classic" };
  }

  // PBIR format: report.json, then one visual.json per visual (its folder is the visual's name)
  const report = await readJson(zip, "Report/definition/report.json").catch(() => null);
  if (!report) throw new Error("UNKNOWN_FORMAT");
  const theme = await themeOrFallback(zip, report.themeCollection, embedTheme);
  const visuals = {};
  const pageNames = {};   // page folder → the page's name (what the embed API calls it)
  for (const path of Object.keys(zip.files).filter((n) => /^Report\/definition\/pages\/[^/]+\/page\.json$/.test(n))) {
    const page = await readJson(zip, path).catch(() => null);
    if (page && page.name) pageNames[path.split("/")[3]] = page.name;
  }
  const paths = Object.keys(zip.files).filter((n) => /^Report\/definition\/pages\/[^/]+\/visuals\/[^/]+\/visual\.json$/.test(n));
  for (const path of paths) {
    try {
      const json = await readJson(zip, path);
      const folder = path.split("/")[3];
      const name = json.name || path.split("/").slice(-2)[0];
      if (json.visual) visuals[visualKey(pageNames[folder] || folder, name)] = visualColorInfo(pbirVisual(json), theme);
    } catch (e) { /* skip an unreadable visual */ }
  }
  return { theme, visuals, byName: uniqueByName(visuals), format: "PBIR" };
}

/*
 * Main entry: a promise of { palette, theme, visuals, format }, or a rejection with a code.
 * Cached per report while it stays open (openReport calls forgetReportColors), so every
 * export of that report reuses one download. Failures that will happen again (no access,
 * unknown format, a 4xx) stay cached too; others (no permission yet, network, timeout)
 * are retried next time. The promise has abort() to stop the download.
 * embedTheme: { dataColors } from the embed API, used only if the file's theme is missing.
 */
const reportColorCache = new Map();
const LASTING_FAILURE = /^(NO_ACCESS|UNKNOWN_FORMAT|DOWNLOAD_FAILED 4\d\d)$/;
function loadReportColors(report, embedTheme) {
  if (!report || !report.id) return Promise.reject(new Error("NO_REPORT"));
  if (!reportColorCache.has(report.id)) {
    const controller = new AbortController();
    const job = (async () => {
      const started = performance.now();
      const zip = await JSZip.loadAsync(await downloadDesignFile(report, controller.signal));
      const design = await readDesign(zip, embedTheme);
      log(`Exact report colours: theme "${design.theme.name}" (${design.theme.dataColors.slice(0, 3).join(", ")}…), ` +
        `${Object.keys(design.visuals).length} visual(s), ${design.format} format, ${Math.round(performance.now() - started)} ms`);
      return { palette: design.theme.dataColors, theme: design.theme, visuals: design.visuals, byName: design.byName, format: design.format };
    })();
    job.abort = () => controller.abort();
    reportColorCache.set(report.id, job);
    job.catch((e) => {
      if (!LASTING_FAILURE.test(String((e && e.message) || "")) && reportColorCache.get(report.id) === job) reportColorCache.delete(report.id);
    });
  }
  return reportColorCache.get(report.id);
}

/* Drops a report's cached colours, so reopening it picks up colour changes made since */
function forgetReportColors(reportId) {
  const job = reportColorCache.get(reportId);
  if (job && job.abort) job.abort();
  reportColorCache.delete(reportId);
}

/* A note for the export when the exact colours couldn't be read */
function reportColorsMessage(e) {
  const code = String((e && e.message) || "");
  if (code === "NO_PERMISSION") return "Chart colours follow the report theme. Colours set by hand on a visual need one more permission (Dataset.Read.All): Export options → Colours → Allow exact colours.";
  if (code === "NO_ACCESS") return "Chart colours follow the report theme. Colours set by hand on a visual are read from the report file, which needs Contributor (or higher) access to the workspace and report downloads allowed.";
  return "Chart colours follow the report theme: the report file, which holds colours set by hand on a visual, couldn't be read.";
}
