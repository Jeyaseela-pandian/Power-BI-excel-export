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
   shade. Theme colour ids: 0 = background, 1 = text (foreground), 2 onwards = data colours.
   Rules (conditional formatting, "fx") return null: their colours depend on the data. */
function colorFromFill(fill, theme) {
  const e = fill && fill.solid && fill.solid.color && fill.solid.color.expr;
  if (!e) return null;
  if (e.Literal) {
    const hex = String(e.Literal.Value).replace(/'/g, "");
    return /^#[0-9A-F]{6}$/i.test(hex) ? hex.toUpperCase() : null;
  }
  if (e.ThemeDataColor) {
    const id = e.ThemeDataColor.ColorId, pct = e.ThemeDataColor.Percent || 0;
    const base = id === 0 ? theme.background
      : id === 1 ? theme.foreground
        : theme.dataColors[(id - 2) % theme.dataColors.length];
    if (!base) return null;
    return pct ? mixHex(base, pct < 0 ? "#000000" : "#FFFFFF", Math.abs(pct)) : base.toUpperCase();
  }
  return null;
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
  return {
    name: (custom && custom.name) || (base && base.name) || "Report theme",
    dataColors,
    background: merged.background || FALLBACK_THEME.background,
    foreground: merged.foreground || FALLBACK_THEME.foreground,
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
  return { type: sv.visualType, roles, dataPoint: (sv.objects || {}).dataPoint || [] };
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
  return { type: v.visualType, roles, dataPoint: (v.objects || {}).dataPoint || [] };
}

/* How a visual colours its data, with every colour set by hand resolved to a hex code */
function visualColorInfo(visual, theme) {
  const overrides = {};
  for (const d of visual.dataPoint) {
    const props = d.properties || {};
    const key = selectorKey(d.selector);
    const color = colorFromFill(props.fill, theme);
    if (color && key) overrides[key] = color;
    const fallback = colorFromFill(props.defaultColor, theme);
    if (fallback && !overrides.all) overrides.all = fallback;
  }
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
    series,
    measures,
    tooltips: (roles.Tooltips || []).map((t) => t.ref),
    overrides,
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
