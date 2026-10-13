"use strict";

/*
 * Report colours – used when Branding → Table style is "Same colours as the report".
 * Downloads the report's design file (possible for Contributor, Member or Admin of the
 * workspace), reads the report's theme and the colours set on each visual, and works
 * out the colour Power BI gives each item, legend value or measure.
 * Only colour information is used; nothing is sent anywhere.
 * Uses globals from app.js (getToken, PBI_API, log) and export.js (toCellValue).
 */

/* ── Defaults and Power BI rules ──────────────────────────────────────────
 * Everything below is read from the report whenever the report provides it.
 * These values are used only when it doesn't:
 *  - FALLBACK_THEME: Power BI's classic default theme. Old reports with no theme
 *    (e.g. the Store Sales sample) store no colours, so this is the only source.
 *  - The visual type lists describe how Power BI colours each visual type. Power BI
 *    doesn't store these rules in reports, so they must live in the code.
 */
const FALLBACK_THEME = {
  name: "Power BI classic default",
  dataColors: ["#01B8AA", "#374649", "#FD625E", "#F2C80F", "#5F6B6D", "#8AD4EB", "#FE9666", "#A66999"],
  background: "#FFFFFF",
  foreground: "#252423",
};
const COLOR_BY_ITEM_TYPES = ["pieChart", "donutChart", "treemap", "funnel"];     // one colour per category
const SINGLE_COLOR_TYPES = ["scatterChart", "map", "filledMap", "shapeMap"];    // one colour without a legend

/* ── Small helpers ────────────────────────────────────────────────────── */
function hexToArgb(hex) {
  return "FF" + String(hex).replace("#", "").toUpperCase();
}

/* Mixes a colour towards another (amount 0–1), e.g. towards white to make it lighter */
function mixHex(hex, target, amount) {
  const a = parseInt(String(hex).replace("#", ""), 16);
  const b = parseInt(String(target).replace("#", ""), 16);
  const ch = (shift) => {
    const x = (a >> shift) & 255, y = (b >> shift) & 255;
    return Math.round(x + (y - x) * amount).toString(16).padStart(2, "0");
  };
  return "#" + (ch(16) + ch(8) + ch(0)).toUpperCase();
}

function isDarkColor(hex) {
  const n = parseInt(String(hex).replace("#", ""), 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 150;
}

/* Field names: "Sum(Sales.This Year Sales)", "'Stores'[Store Type]", "Sum of Sales" → "this year sales", "store type", "sales" */
function normalizeField(text) {
  let s = String(text ?? "").trim();
  const agg = s.match(/^[A-Za-z]+\((.*)\)$/);
  if (agg) s = agg[1];
  const col = s.match(/\[([^\]]+)\]\s*$/);
  if (col) s = col[1];
  else if (s.includes(".")) s = s.slice(s.lastIndexOf(".") + 1);
  return s.replace(/^(sum|average|count|min|max) of /i, "").trim().toLowerCase();
}

/* Item values: "'Fashions Direct'" → "fashions direct", "2024L" → "2024" */
function normalizeValue(text) {
  return String(text ?? "").trim().replace(/^'(.*)'$/, "$1").replace(/^(-?\d+(?:\.\d+)?)[LDM]$/, "$1").toLowerCase();
}

/* A colour set on a visual: "#01B8AA", or a theme colour (with optional lighter/darker shade).
   Theme colour ids: 0 = theme background, 1 = theme text, 2 onwards = theme data colours. */
function colorFromFill(fill, theme) {
  const e = fill && fill.solid && fill.solid.color && fill.solid.color.expr;
  if (!e) return null;
  if (e.Literal) return String(e.Literal.Value).replace(/'/g, "");
  if (e.ThemeDataColor) {
    const id = e.ThemeDataColor.ColorId, pct = e.ThemeDataColor.Percent || 0;
    const base = id === 0 ? theme.background
      : id === 1 ? theme.foreground
      : theme.dataColors[(id - 2) % theme.dataColors.length];
    if (!base) return null;
    return pct ? mixHex(base, pct < 0 ? "#000000" : "#FFFFFF", Math.abs(pct)) : base;
  }
  return null;
}

/* Which item or measure a manually set colour belongs to */
function selectorKey(selector) {
  if (!selector) return "all";
  if (selector.metadata) return "m:" + normalizeField(selector.metadata);
  const values = [...JSON.stringify(selector).matchAll(/"Literal":\{"Value":"((?:[^"\\]|\\.)*)"\}/g)]
    .map((m) => normalizeValue(JSON.parse('"' + m[1] + '"')));
  return values.length ? "i:" + values.join("|") : null;
}

/* ── Reading the design file ──────────────────────────────────────────── */
async function downloadDesignFile(report) {
  const groupId = new URL(report.embedUrl).searchParams.get("groupId");
  const base = PBI_API + (groupId ? `/groups/${groupId}` : "") + `/reports/${report.id}/Export`;
  // Shared workspace: try "design only" first; My workspace and sample reports need the full file
  const urls = groupId ? [base + "?downloadType=LiveConnect", base] : [base];
  const token = await getToken();
  for (const url of urls) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (res.ok) return res.arrayBuffer();
    if (res.status === 401 || res.status === 403) throw new Error("NO_ACCESS");
  }
  throw new Error("DOWNLOAD_FAILED");
}

/* The report's theme (custom theme first, then base theme), read from the design file */
async function readTheme(zip, layout) {
  let themes = {};
  try { themes = JSON.parse(layout.config || "{}").themeCollection || {}; } catch (e) { /* no theme */ }
  for (const t of [themes.customTheme, themes.baseTheme].filter(Boolean)) {
    const name = String(t.name || "").toLowerCase();
    if (!name) continue;
    const path = Object.keys(zip.files).find((n) => n.toLowerCase().endsWith(".json") && n.toLowerCase().includes(name));
    if (!path) continue;
    try {
      const json = JSON.parse((await zip.file(path).async("string")).replace(/^\uFEFF/, ""));
      if (Array.isArray(json.dataColors) && json.dataColors.length) {
        return {
          name: json.name || t.name,
          dataColors: json.dataColors,
          background: json.background || FALLBACK_THEME.background,
          foreground: json.foreground || FALLBACK_THEME.foreground,
        };
      }
    } catch (e) { /* try the next theme */ }
  }
  return FALLBACK_THEME;
}

/* How one visual is coloured: by item, by legend value, by measure, or one colour */
function visualColorInfo(config, theme) {
  const sv = config.singleVisual || {};
  const series = ((sv.projections || {}).Series || [])[0];
  const overrides = {};
  for (const d of (sv.objects || {}).dataPoint || []) {
    const color = colorFromFill(d.properties && d.properties.fill, theme);
    const key = selectorKey(d.selector);
    if (color && key) overrides[key] = color;
  }
  let rule = "measure";
  if (COLOR_BY_ITEM_TYPES.includes(sv.visualType)) rule = "item";
  else if (series) rule = "legend";
  else if (SINGLE_COLOR_TYPES.includes(sv.visualType)) rule = "single";
  return { rule, seriesField: series ? normalizeField(series.queryRef) : null, overrides };
}

/* Main entry: returns { palette, visuals: { visualName: info }, kpi } or throws */
async function loadReportColors(report) {
  const zip = await JSZip.loadAsync(await downloadDesignFile(report));
  const layoutFile = zip.file("Report/Layout");
  if (!layoutFile) throw new Error("UNSUPPORTED_FORMAT");   // newer report format (PBIR)
  const layout = JSON.parse(new TextDecoder("utf-16le").decode(await layoutFile.async("uint8array")));
  const theme = await readTheme(zip, layout);

  const visuals = {};
  for (const section of layout.sections || []) {
    for (const vc of section.visualContainers || []) {
      try {
        const config = JSON.parse(vc.config);
        if (config.name) visuals[config.name] = visualColorInfo(config, theme);
      } catch (e) { /* skip unreadable visual */ }
    }
  }
  const palette = theme.dataColors;
  log(`Report colours loaded: theme "${theme.name}", ${Object.keys(visuals).length} visual(s)`);
  return {
    palette,
    visuals,
    kpi: { header: hexToArgb(palette[0]), stripe: hexToArgb(mixHex(palette[0], theme.background, 0.8)) },
  };
}

function reportColorsMessage(e) {
  const code = e && e.message;
  if (code === "NO_ACCESS") return "Report colours need edit access (Contributor or higher) to this report, so standard colours were used.";
  if (code === "UNSUPPORTED_FORMAT") return "This report uses Power BI's newer report format, which report colours don't support yet, so standard colours were used.";
  return "Report colours couldn't be read, so standard colours were used.";
}

/* ── Painting an exported table ───────────────────────────────────────── */
function applyReportColors(ws, row, col, item, header, rc) {
  const info = rc && item.name && rc.visuals[item.name];
  if (!info) return;
  const data = item.rows.slice(1);
  const paint = (cell, hex) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: hexToArgb(hex) } };
    cell.font = { ...(cell.font || {}), bold: true, color: { argb: isDarkColor(hex) ? "FFFFFFFF" : "FF000000" } };
  };

  // By item (pie…) or by legend value (map…): colour the item / legend cells
  if (info.rule === "item" || info.rule === "legend") {
    const ci = info.rule === "item" ? 0 : header.findIndex((h) => normalizeField(h) === info.seriesField);
    if (ci < 0) return;
    let order = [];
    data.forEach((r) => { if (!order.includes(r[ci])) order.push(r[ci]); });
    if (info.rule === "legend") order = order.sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
    data.forEach((r, ri) => {
      const hex = info.overrides["i:" + normalizeValue(r[ci])] || rc.palette[order.indexOf(r[ci]) % rc.palette.length];
      paint(ws.getCell(row + 2 + ri, col + ci), hex);
    });
    return;
  }

  // By measure (bar, column, line, combo…) or one colour: colour the measure headers
  let m = 0;
  header.forEach((h, ci) => {
    if (ci === 0 || !data.length) return;
    const numeric = data.every((r) => (r[ci] ?? "").trim() === "" || typeof toCellValue(r[ci]).value === "number");
    if (!numeric) return;
    const hex = info.overrides["m:" + normalizeField(h)] || (info.rule === "single" ? rc.palette[0] : rc.palette[m % rc.palette.length]);
    m++;
    paint(ws.getCell(row + 1, col + ci), hex);
  });
}
