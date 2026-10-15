"use strict";

/*
 * Exact report colours. Downloads the report's design file (needs the Dataset.Read.All
 * permission and download rights: Contributor, Member or Admin of the workspace), reads
 * the report's theme and the colours set by hand on each visual, and describes how each
 * visual colours its data (by item, by legend value, by measure, or one colour).
 * Both report formats are read: the classic one ("Report/Layout") and the newer PBIR
 * folders ("Report/definition/…"). Only colour information is used; nothing is stored
 * or sent anywhere. When the file can't be read the export falls back to the theme
 * colours Power BI reports through the embed API (export.js).
 * Uses globals from app.js: getDesignToken, PBI_API, log, errorText.
 */

/* ── Power BI rules ───────────────────────────────────────────────────────
 * Power BI doesn't store these in reports, so they live in the code.
 *  - FALLBACK_THEME: colours of Power BI's classic theme, for old reports with no theme.
 *  - COLOR_BY_ITEM_TYPES: visuals that give each category value its own colour.
 *  - SINGLE_COLOR_TYPES: visuals drawn in one colour unless a legend is used.
 *  - Funnels draw every bar in one colour unless colours are set per bar.
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
  };
}

/*
 * One visual, in a shape shared by both formats:
 * { type, roles: { Category: [{ ref, label }], Series: [...], Y: [...] … }, dataPoint: [...] }
 * ref is the normalised field name; label the visual's own name for it when renamed.
 */
function legacyVisual(config) {
  const sv = config.singleVisual || {};
  const props = sv.columnProperties || {};
  const roles = {};
  for (const [role, list] of Object.entries(sv.projections || {})) {
    roles[role] = (list || []).map((p) => ({
      ref: normalizeField(p.queryRef),
      label: props[p.queryRef] && props[p.queryRef].displayName ? String(props[p.queryRef].displayName).toLowerCase() : null,
    }));
  }
  const o = sv.objects || {};
  return { type: sv.visualType, roles, objects: o, dataPoint: o.dataPoint || [], bubbles: o.bubbleLayer || [], values: [...(o.value || []), ...(o.labels || [])] };
}

function pbirVisual(json) {
  const v = json.visual || {};
  const roles = {};
  for (const [role, state] of Object.entries((v.query && v.query.queryState) || {})) {
    roles[role] = ((state && state.projections) || []).map((p) => ({
      ref: normalizeField(p.queryRef || p.nativeQueryRef),
      label: p.displayName ? String(p.displayName).toLowerCase() : null,
    }));
  }
  const o = v.objects || {};
  return { type: v.visualType, roles, objects: o, dataPoint: o.dataPoint || [], bubbles: o.bubbleLayer || [], values: [...(o.value || []), ...(o.labels || [])] };
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
  for (const d of (objects && objects.columnFormatting) || []) {
    const bars = (d.properties || {}).dataBars;
    if (!bars || !d.selector || !d.selector.metadata) continue;
    const color = colorFromFill(bars.positiveColor, theme);
    const hide = bars.hideText && bars.hideText.expr && bars.hideText.expr.Literal && bars.hideText.expr.Literal.Value === "true";
    if (color) out.push({ target: "m:" + normalizeField(d.selector.metadata), kind: "bar", bar: { color, showValue: !hide } });
  }
  return out;
}

/* How a visual colours its data, with every colour set by hand resolved to a hex code */
function visualColorInfo(visual, theme) {
  const overrides = {};
  const rules = [];   // fx rules: { scope, rule }
  const points = visual.dataPoint.map((d) => [d, (d.properties || {}).fill]);
  const bubbles = visual.bubbles.map((d) => [d, (d.properties || {}).fillColor]);
  // An Azure Map draws its bubble layer's colour; its data point colour is a leftover
  const entries = visual.type === "azureMap" ? [...bubbles, ...points] : [...points, ...bubbles];
  for (const [d, fill] of entries) {
    const props = d.properties || {};
    const key = selectorKey(d.selector);
    const color = colorFromFill(fill, theme);
    if (color && key) overrides[key] = color;
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
  // Waterfall: increase / decrease / total colours (theme good / bad / first colour by default)
  const sentiment = ((visual.objects || {}).sentimentColors || [])[0];
  const sp = (sentiment && sentiment.properties) || {};
  const kw = (k) => (theme.keywords && theme.keywords[k]) || KEYWORD_DEFAULTS[k];
  const waterfall = visual.type === "waterfallChart" ? {
    increase: colorFromFill(sp.increaseFill, theme) || kw("good"),
    decrease: colorFromFill(sp.decreaseFill, theme) || kw("bad"),
    total: colorFromFill(sp.totalFill, theme) || theme.dataColors[0],
  } : null;
  const roles = visual.roles;
  const series = (roles.Series || [])[0] || null;
  const measures = MEASURE_ROLES.flatMap((r) => roles[r] || []);
  let rule = "measure";
  if (COLOR_BY_ITEM_TYPES.includes(visual.type)) rule = "item";
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
    waterfall,
    roles,   // every field well of the visual (X, Y, Size, Legend…), used by scatter charts
    tableFormats: ["tableEx", "pivotTable", "table", "matrix"].includes(visual.type) ? tableFormats(visual.objects, theme) : [],
  };
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
          if (c.name && c.singleVisual) visuals[c.name] = visualColorInfo(legacyVisual(c), theme);
        } catch (e) { /* skip an unreadable visual */ }
      }
    }
    return { theme, visuals, format: "classic" };
  }

  // PBIR format: report.json, then one visual.json per visual (its folder is the visual's name)
  const report = await readJson(zip, "Report/definition/report.json").catch(() => null);
  if (!report) throw new Error("UNKNOWN_FORMAT");
  const theme = await themeOrFallback(zip, report.themeCollection, embedTheme);
  const visuals = {};
  const paths = Object.keys(zip.files).filter((n) => /^Report\/definition\/pages\/[^/]+\/visuals\/[^/]+\/visual\.json$/.test(n));
  for (const path of paths) {
    try {
      const json = await readJson(zip, path);
      const name = json.name || path.split("/").slice(-2)[0];
      if (json.visual) visuals[name] = visualColorInfo(pbirVisual(json), theme);
    } catch (e) { /* skip an unreadable visual */ }
  }
  return { theme, visuals, format: "PBIR" };
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
      return { palette: design.theme.dataColors, theme: design.theme, visuals: design.visuals, format: design.format };
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
