"use strict";

/*
 * Builds the Excel workbook – one sheet per exported report page:
 *  1. Header – optional logo and company name, report / page name, export time,
 *     active filters and slicer selections.
 *  2. Visuals, laid out either
 *     - "report": like the Power BI page. Visuals whose vertical extents overlap form
 *       a "band" (a row of the page) and start on the same Excel row; inside a band,
 *       visuals sharing horizontal space form a "stack" placed one under another;
 *       each stack starts at the Excel column matching its x position. Or
 *     - "stacked": one visual under another in reading order, all from column A.
 *     Cards / KPIs are label / value pairs; every other visual is a real Excel Table
 *     with its own filter and sort buttons.
 *  3. Long tables collapse rows past `visibleRows` behind [+] / [-], but only rows that
 *     hold nothing except those extra table rows – Excel hides whole rows, so a row
 *     shared with any other content is never collapsed. A small note line sits at the
 *     bottom of each collapsed table, next to its [+] / [-] button. Note lines of
 *     side-by-side tables are lined up so no table is split.
 *     Tables can end in a totals row (SUBTOTAL formulas, so it follows Excel filters);
 *     it shares the note row, so it takes part in the same alignment.
 *  4. Export notes and footer.
 * Values from Power BI's CSV are turned into real numbers and dates (with formats
 * matching the report), matrices can be rebuilt as cross-tabs, and number columns can
 * get Excel colour scales or data bars.
 * Uses globals from app.js: embeddedReport, currentReportName, models, log, errorText,
 * describeFilters, NON_DATA_TYPES, cfg.
 */

const KPI_TYPES = ["card", "multiRowCard", "kpi", "cardVisual", "gauge"];
const PX_PER_COL = 110;        // Power BI page pixels per Excel column (for x alignment)
const PX_TOLERANCE = 10;       // overlap smaller than this does not count as overlapping
const PAGE_RENDER_TIMEOUT_MS = 20000;
const MAX_EXPORT_ROWS = 30000; // Power BI's limit for exportData; reaching it means rows may be missing
const EXPORT_ATTEMPTS = 3;     // one try + two retries for errors that may be temporary
const MAX_CROSSTAB_COLUMNS = 60;
const CF_COLORS = {
  scale: ["FFF8696B", "FFFFEB84", "FF63BE7B"],   // red – yellow – green, as Excel's default 3-colour scale
  bar: "FF638EC6",
};


const COLORS = {
  kpiLabel: "FFF5F5F5",
  note: "FF888888",
};
const THIN = { style: "thin", color: { argb: "FFD0D0D0" } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const TABLE_STYLE_COLORS = {
  TableStyleMedium2: { header: "FF4472C4", stripe: "FFD9E1F2" },  // Blue
  TableStyleMedium7: { header: "FF70AD47", stripe: "FFE2EFDA" },  // Green
  TableStyleMedium3: { header: "FFED7D31", stripe: "FFFCE4D6" },  // Orange
  TableStyleMedium4: { header: "FFA5A5A5", stripe: "FFEDEDED" },  // Gray
  TableStyleLight9:  { header: "FF4472C4", stripe: "FFDDEBF7" },  // Light blue
};

/* "#1a73e8" → "FF1A73E8" */
function toArgb(hex) {
  const h = (hex || "#1A73E8").replace("#", "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(h) ? "FF" + h : "FF1A73E8";
}

/* ── CSV → cells ──────────────────────────────────────────────────────── */
function parseCsv(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  text = text.replace(/^\uFEFF/, "");
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
 * Rows where every column that is sometimes filled is blank, and at least one of those
 * columns holds numbers – e.g. months outside a date slicer that Power BI still lists
 * because some measures (colours, ranks, labels) return a constant. Optional.
 */
function removeEmptyRows(rows) {
  const [header, ...data] = rows;
  const isBlank = (v) => (v ?? "").trim() === "";
  const sparse = header.map((_, ci) => ci).filter((ci) => {
    const blanks = data.filter((r) => isBlank(r[ci])).length;
    return blanks > 0 && blanks < data.length;
  });
  const hasNumbers = sparse.some((ci) => data.some((r) => typeof toCellValue(r[ci]).value === "number"));
  if (sparse.length === 0 || !hasNumbers) return { rows, removed: 0 };
  const kept = data.filter((r) => !sparse.every((ci) => isBlank(r[ci])));
  return { rows: [header, ...kept], removed: data.length - kept.length };
}

/* ── Values: numbers and dates ────────────────────────────────────────── */

/* The browser's conventions decide ambiguous values such as "1,234" or "6/7/2024" */
const LOCALE = (() => {
  try {
    const decimal = new Intl.NumberFormat().formatToParts(1234.5).find((p) => p.type === "decimal").value;
    const order = new Intl.DateTimeFormat().formatToParts(new Date(2001, 10, 22))
      .filter((p) => p.type === "day" || p.type === "month" || p.type === "year")
      .map((p) => p.type[0]).join("");
    return { decimal, dateOrder: /^(dmy|mdy|ymd)$/.test(order) ? order : "mdy" };
  } catch (e) {
    return { decimal: ".", dateOrder: "mdy" };
  }
})();

const CURRENCY = "[$€£₹¥]|R\\$|Rs\\.?|USD|EUR|GBP|INR|AUD|CAD";
const SCALES = { K: 1e3, M: 1e6, bn: 1e9, B: 1e9, T: 1e12 };
const SCALE_COMMAS = { K: ",", M: ",,", bn: ",,,", B: ",,,", T: ",,,," };

/*
 * "$1,234.50", "1.234,56 €", "(300)", "12.5%", "1.2K", "3,4 M" → a number plus an Excel
 * format that shows it the way Power BI did. Returns null when it is not a number.
 */
function parseNumber(s) {
  let t = s, negative = false, parens = false;
  if (/^\(.*\)$/.test(t)) { negative = true; parens = true; t = t.slice(1, -1).trim(); }
  if (t.startsWith("-")) { negative = !negative; t = t.slice(1).trim(); }
  const isPercent = t.endsWith("%");
  if (isPercent) t = t.slice(0, -1).trim();

  let currency = null, currencyAfter = false, m;
  if ((m = t.match(new RegExp(`^(${CURRENCY})\\s?`, "i")))) { currency = m[1]; t = t.slice(m[0].length); }
  else if ((m = t.match(new RegExp(`\\s?(${CURRENCY})$`, "i")))) { currency = m[1]; currencyAfter = true; t = t.slice(0, -m[0].length); }
  if (t.startsWith("-")) { negative = !negative; t = t.slice(1); }

  let scale = "";
  if ((m = t.match(/^(.*\d)\s?(K|M|bn|B|T)$/))) { t = m[1]; scale = m[2]; }

  // Which of "." and "," is the decimal point?
  const dot = t.lastIndexOf("."), comma = t.lastIndexOf(",");
  let decimalSep = ".";
  if (dot >= 0 && comma >= 0) decimalSep = dot > comma ? "." : ",";
  else if (dot >= 0 || comma >= 0) {
    const sep = dot >= 0 ? "." : ",";
    const repeated = t.split(sep).length > 2;
    const digitsAfter = t.length - t.lastIndexOf(sep) - 1;
    if (repeated) decimalSep = sep === "." ? "," : ".";          // 1.234.567 – a thousands separator
    else if (digitsAfter !== 3) decimalSep = sep;                // 12.5 or 1,25 – a decimal point
    else decimalSep = LOCALE.decimal === sep ? sep : (sep === "." ? "," : "."); // 1,234 – ask the locale
  }
  const groupSep = decimalSep === "." ? "," : ".";
  // Thousands groups are exactly three digits: "1.2.3" or "12,34,5" is not a number
  const whole = t.split(decimalSep)[0].split(/[\s  ']/).join(groupSep).split(groupSep);
  if (whole.length > 1 && (!/^\d{1,3}$/.test(whole[0]) || whole.slice(1).some((g) => !/^\d{3}$/.test(g)))) return null;
  const hadGroups = t.includes(groupSep) || /[\s  ']\d/.test(t);
  let cleaned = t.split(groupSep).join("").replace(/[\s  ']/g, "");
  if (decimalSep === ",") cleaned = cleaned.replace(",", ".");
  if (!/^\d+(\.\d+)?$/.test(cleaned)) return null;

  const decimals = Math.min((cleaned.split(".")[1] || "").length, 4);
  const value = Number(cleaned) * (scale ? SCALES[scale] : 1) * (negative ? -1 : 1);
  const dec = decimals ? "." + "0".repeat(decimals) : "";
  let fmt = `${hadGroups || scale ? "#,##0" : "0"}${dec}`;    // no separator for plain values like years
  if (scale) fmt += `${SCALE_COMMAS[scale]}"${scale}"`;
  if (isPercent) return { value: value / 100, numFmt: `${fmt}%` };
  if (currency) fmt = currencyAfter ? `${fmt} "${currency}"` : `"${currency}"${fmt}`;
  if (parens) fmt = `${fmt};(${fmt})`;
  return { value, numFmt: fmt };
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
// Whole month names or their abbreviations only, so words like "Marketing" stay text
const MONTH_NAME = "(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept|sep|oct|nov|dec)\\.?";

function makeDate(y, mo, d, h = 0, mi = 0, sec = 0) {
  if (y < 100) y += y < 50 ? 2000 : 1900;
  if (mo < 1 || mo > 12 || d < 1 || d > new Date(Date.UTC(y, mo, 0)).getUTCDate()) return null;
  if (h > 23 || mi > 59 || sec > 59) return null;
  // ExcelJS stores dates as UTC, so build the wall-clock time in UTC
  return new Date(Date.UTC(y, mo - 1, d, h, mi, sec));
}

function monthIndex(word) {
  return MONTHS.indexOf(word.slice(0, 3).toLowerCase()) + 1;
}

/*
 * Dates as Power BI writes them in its CSV – "2024-06-07", "6/7/2024 12:00:00 AM",
 * "07/06/2024", "Jun 2024", "June 7, 2024", "07-Jun-24" – become real Excel dates with a
 * matching format, so Excel's date filters and sorting work. Returns null otherwise.
 */
function parseDate(s) {
  let m;
  // Optional time part shared by the numeric forms
  const TIME = "(?:[ T](\\d{1,2}):(\\d{2})(?::(\\d{2}))?(?:\\.\\d+)?\\s*([AP]M)?)?";
  const withTime = (date, h, mi, sec, ampm, dateFmt) => {
    if (h === undefined) return { value: date, numFmt: dateFmt };
    let hour = Number(h);
    if (ampm) hour = (hour % 12) + (ampm.toUpperCase() === "PM" ? 12 : 0);
    const full = makeDate(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate(), hour, Number(mi), Number(sec || 0));
    if (!full) return null;
    const midnight = hour === 0 && Number(mi) === 0 && Number(sec || 0) === 0;
    return { value: full, numFmt: midnight ? dateFmt : `${dateFmt} ${ampm ? "h:mm AM/PM" : "hh:mm"}` };
  };

  // ISO: 2024-06-07, 2024-06-07T13:45:00
  if ((m = s.match(new RegExp(`^(\\d{4})-(\\d{2})-(\\d{2})${TIME}$`, "i")))) {
    const d = makeDate(+m[1], +m[2], +m[3]);
    return d && withTime(d, m[4], m[5], m[6], m[7], "yyyy-mm-dd");
  }
  // Numeric: 6/7/2024, 07.06.2024, 7-6-24, 2024/06/07 (+ time)
  if ((m = s.match(new RegExp(`^(\\d{1,4})([/.-])(\\d{1,2})\\2(\\d{2,4})${TIME}$`, "i")))) {
    const [a, sep, b, c] = [+m[1], m[2], +m[3], +m[4]];
    let y, mo, d, fmt;
    if (m[1].length === 4) { y = a; mo = b; d = c; fmt = `yyyy${sep}mm${sep}dd`; }
    else if (m[4].length < 2) return null;
    else {
      // Day or month first? A value above 12 settles it; dotted dates (07.06.2024) are
      // day-first by convention; otherwise the locale decides
      const dayFirst = a > 12 ? true : b > 12 ? false : sep === "." || LOCALE.dateOrder === "dmy";
      [d, mo] = dayFirst ? [a, b] : [b, a];
      y = c;
      fmt = dayFirst ? `dd${sep}mm${sep}yyyy` : `m${sep}d${sep}yyyy`;
    }
    const date = makeDate(y, mo, d);
    return date && withTime(date, m[5], m[6], m[7], m[8], fmt);
  }
  // Month and year: Jan 2024, January 2024, Jan-24
  if ((m = s.match(new RegExp(`^${MONTH_NAME}[\\s-]+(\\d{4}|\\d{2})$`, "i")))) {
    const date = makeDate(+m[2], monthIndex(m[1]), 1);
    return date && { value: date, numFmt: s.length > 8 && !/^\w{3}[\s-]/.test(s) ? "mmmm yyyy" : "mmm yyyy" };
  }
  // June 7, 2024 · Friday, June 7, 2024 · Jun 7 2024
  if ((m = s.match(new RegExp(`^(?:([a-z]+),\\s*)?${MONTH_NAME}\\s+(\\d{1,2}),?\\s+(\\d{4})$`, "i")))) {
    const date = makeDate(+m[4], monthIndex(m[2]), +m[3]);
    const long = m[1] ? "dddd, mmmm d, yyyy" : /^[a-z]+\s/i.test(s) && s.split(" ")[0].replace(/\.$/, "").length > 3 ? "mmmm d, yyyy" : "mmm d, yyyy";
    return date && { value: date, numFmt: long };
  }
  // 7 June 2024 · 07-Jun-24 · 7 Jun 2024
  if ((m = s.match(new RegExp(`^(\\d{1,2})[\\s-]${MONTH_NAME}[\\s-,]+(\\d{2,4})$`, "i")))) {
    const date = makeDate(+m[3], monthIndex(m[2]), +m[1]);
    const dash = s.includes("-");
    return date && { value: date, numFmt: dash ? (m[3].length === 2 ? "dd-mmm-yy" : "dd-mmm-yyyy") : "d mmm yyyy" };
  }
  return null;
}

/*
 * Power BI's CSV carries values as formatted text. Turn them into real Excel numbers
 * and dates with a matching format, so they can be filtered, sorted and summed.
 * Anything that is not clearly a number or a date stays text.
 */
function toCellValue(raw) {
  const s = (raw ?? "").trim();
  if (s === "") return { value: null };
  if (/^0\d+$/.test(s)) return { value: s };             // codes with leading zeros (ZIP, IDs)
  return parseDate(s) || parseNumber(s) || { value: s };
}

/* ── Helpers ──────────────────────────────────────────────────────────── */
function makeWidthTracker() {
  const widths = {};
  return {
    note(col, text) {
      const len = String(text ?? "").length;
      widths[col] = Math.max(widths[col] || 0, len);
    },
    apply(ws, min = 12, max = 40) {
      for (const [col, len] of Object.entries(widths)) {
        ws.getColumn(Number(col)).width = Math.max(min, Math.min(max, len + 4)); // +4 leaves room for filter buttons
      }
    },
  };
}

function styleTitle(cell, text, accent) {
  cell.value = text;
  cell.font = { bold: true, size: 12, color: { argb: accent } };
  cell.alignment = { vertical: "middle", horizontal: "left" }; // long titles overflow into empty cells to the right
}

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

/* Sheet names: unique, at most 31 characters, none of \ / * ? : [ ] */
function makeSheetNamer() {
  const used = new Set();
  return (title) => {
    const base = (title || "Page").replace(/[\\/*?:[\]]/g, "").trim().slice(0, 31) || "Page";
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

/* ── Block writers ────────────────────────────────────────────────────── */

/*
 * Size of a visual's block in Excel cells. Tables get one extra row at the bottom –
 * the note row – when they have a totals row or, with collapsing on, more rows than
 * stay visible ("x of y rows shown" sits there).
 */
function blockSize(item, options) {
  if (item.isKpi) return { width: 2, height: 1 + item.rows[0].length, note: false };
  const n = item.rows.length - 1;
  const note = !!(options && ((options.collapse && n > options.visibleRows) || item.totals));
  return { width: item.rows[0].length, height: n + 2 + (note ? 1 : 0), note }; // title + header + data (+ note)
}

/*
 * What the totals row shows per column: "sum", "average" or null (no total).
 * Only columns that are all numbers get one; ranks, IDs, years and sort keys don't;
 * percentages, averages, rates and prices are averaged rather than added up.
 */
function totalFunction(name, cells) {
  const filled = cells.filter((c) => c.value !== null);
  if (!filled.length || filled.some((c) => typeof c.value !== "number")) return null;
  if (/\b(rank|id|key|code|year|sort|index|no|number|zip|postal)\b|#/i.test(name)) return null;
  if (filled.every((c) => Number.isInteger(c.value) && c.value >= 1900 && c.value <= 2100 && !(c.numFmt || "").includes(","))) return null;
  if (filled.some((c) => (c.numFmt || "").includes("%")) ||
      /avg|average|mean|median|rate|ratio|percent|%|margin|price|per\b|score/i.test(name)) return "average";
  return "sum";
}

/*
 * Converts a table's values once, ahead of layout, and decides its totals:
 * item.cells (rows of {value, numFmt}) and item.totals (per column, or null if none).
 */
function prepareTable(item, options) {
  item.header = uniqueHeaders(item.rows[0]);
  item.cells = item.rows.slice(1).map((r) => item.header.map((_, i) => toCellValue(r[i])));
  const fns = item.header.map((h, ci) => totalFunction(h, item.cells.map((r) => r[ci])));
  item.measureColumns = fns.map((f) => f !== null);
  item.totals = options.totalsRow && fns.some(Boolean) ? fns : null;
}

/* KPI / card: title, then one label / value pair per field. */
function writeKpiBlock(ws, row, col, item, widths, style) {
  const [header, values = []] = item.rows;
  const colors = TABLE_STYLE_COLORS[style.tableStyle] || TABLE_STYLE_COLORS.TableStyleMedium2;
  styleTitle(ws.getCell(row, col), item.title, style.accent);
  header.forEach((label, i) => {
    const l = ws.getCell(row + 1 + i, col);
    l.value = label;
    l.font = { bold: true, color: { argb: "FFFFFFFF" } };
    l.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header } };
    l.border = BORDER;
    const v = ws.getCell(row + 1 + i, col + 1);
    const { value, numFmt } = toCellValue(values[i]);
    v.value = value;
    if (numFmt) v.numFmt = numFmt;
    v.font = { bold: true, size: 12 };
    v.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.stripe } };
    v.alignment = { horizontal: "right" };
    v.border = BORDER;
    widths.note(col, label);
    widths.note(col + 1, values[i]);
  });
}

/* Data visual: title row, then an Excel Table (header + data) with its own filter buttons. */
function writeDataTable(ws, row, col, item, tableName, widths, style, options) {
  const header = item.header;
  const data = item.cells;

  styleTitle(ws.getCell(row, col), item.displayTitle || item.title, style.accent);

  ws.addTable({
    name: tableName,
    ref: ws.getCell(row + 1, col).address,
    headerRow: true,
    style: { theme: style.tableStyle, showRowStripes: true },
    columns: header.map((name) => ({ name, filterButton: true })),
    rows: data.map((r) => r.map((c) => c.value)),
  });

  // Number formats per cell (the table only carries values)
  data.forEach((r, ri) => r.forEach((c, ci) => {
    if (c.numFmt) ws.getCell(row + 2 + ri, col + ci).numFmt = c.numFmt;
  }));

  // Live Excel rules on measure columns, so colours follow filtering and edits
  if (options.conditionalFormat !== "none" && data.length > 1) {
    item.measureColumns.forEach((isMeasure, ci) => {
      if (!isMeasure) return;
      const ref = `${ws.getCell(row + 2, col + ci).address}:${ws.getCell(row + 1 + data.length, col + ci).address}`;
      const rule = options.conditionalFormat === "dataBar"
        ? { type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: CF_COLORS.bar }, gradient: true }
        : { type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: CF_COLORS.scale.map((argb) => ({ argb })) };
      ws.addConditionalFormatting({ ref, rules: [rule] });
    });
  }

  header.forEach((h, ci) => widths.note(col + ci, h));
  item.rows.slice(1).forEach((r) => r.forEach((v, ci) => widths.note(col + ci, v)));
}

/*
 * Totals in the table's note row: SUBTOTAL formulas over the data rows, so they follow
 * Excel filters (filtered-out rows drop out) while rows collapsed behind [+] still count.
 * The cached result is filled in too, for viewers that don't recalculate.
 */
function writeTotalsRow(ws, noteRow, t) {
  const { item, top, col, n } = t;
  item.totals.forEach((fn, ci) => {
    if (!fn) return;
    const first = ws.getCell(top + 2, col + ci).address, last = ws.getCell(top + 1 + n, col + ci).address;
    const nums = item.cells.map((r) => r[ci].value).filter((v) => typeof v === "number");
    const sum = nums.reduce((a, b) => a + b, 0);
    const cell = ws.getCell(noteRow, col + ci);
    cell.value = { formula: `SUBTOTAL(${fn === "sum" ? 9 : 1},${first}:${last})`, result: fn === "sum" ? sum : (nums.length ? sum / nums.length : 0) };
    const fmt = (item.cells.find((r) => r[ci].numFmt) || [])[ci];
    if (fmt && fmt.numFmt) cell.numFmt = fmt.numFmt;
    cell.font = { bold: true };
    cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
  });
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

/* Rows of a planned table that may collapse: data rows past the threshold, plus padding before its note */
function collapsibleRange(p, threshold) {
  if (p.item.isKpi) return null;
  const n = p.item.rows.length - 1;
  if (n <= threshold) return null;
  const end = p.noteRow ? p.noteRow - 1 : p.r + 1 + n;
  return [p.r + 2 + threshold, end];
}

/* Rows of a planned block that must always stay visible: title, header and the first
   rows of a table, or the whole block for cards and short tables */
function topRange(p, threshold) {
  if (p.item.isKpi) return [p.r, p.r + p.size.height - 1];
  const n = p.item.rows.length - 1;
  return [p.r, p.r + 1 + Math.min(n, threshold)];
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
 *  1. A table's collapsible rows must not sit next to a neighbour's visible rows,
 *     otherwise those rows can't be hidden and the table shows too many rows.
 *     The table is moved down so its collapsing starts after the neighbour's
 *     visible rows (side-by-side tables end up starting on the same row).
 *  2. A table's note row must not sit inside a neighbour's collapsible rows,
 *     otherwise it splits the neighbour. The note moves down below them, so both
 *     notes share one row and one [+] / [-] button.
 */
function alignNoteRows(plan, threshold) {
  for (let guard = 0; guard < 200; guard++) {
    let moved = false;
    for (const p of plan) {
      const myRange = collapsibleRange(p, threshold);
      for (const o of plan) {
        if (o === p || o.stackIndex === p.stackIndex) continue;

        // 1. My collapsible rows next to a neighbour's visible rows: move me down
        const top = topRange(o, threshold);
        if (myRange && myRange[0] <= top[1] && top[0] <= myRange[1]) {
          shiftStack(plan, p, top[1] + 1 - myRange[0]);
          moved = true;
          break;
        }

        // 2. My note row inside a neighbour's collapsible rows: move my note down
        const range = collapsibleRange(o, threshold);
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
 * layout, the number of tables written and whether any rows were collapsed.
 */
function writePageLayout(ws, items, startRow, widths, options, style, nameTable) {
  const stacked = options.layout === "stacked";
  const threshold = options.visibleRows;
  const tables = [];
  const rowUse = [];  // per Excel row: does it hold collapsible table rows ("deep") / anything else ("other")?
  const mark = (r, kind) => { (rowUse[r] = rowUse[r] || { deep: false, other: false })[kind] = true; };

  const bands = stacked
    ? groupIntoBands(items).flatMap((b) => [...b.items].sort((a, c) => a.layout.x - c.layout.x).map((it) => ({ items: [it] })))
    : groupIntoBands(items);

  let bandTop = startRow;
  for (const band of bands) {
    // 1. Plan where every visual of this band goes
    const plan = [];
    let cursorCol = 1;
    const stacks = stacked ? [{ items: band.items, left: 0 }] : groupIntoStacks(band.items);
    stacks.forEach((stack, stackIndex) => {
      const col = stacked ? 1 : Math.max(cursorCol, 1 + Math.round(stack.left / PX_PER_COL));
      let r = bandTop, stackWidth = 0;
      for (const item of stack.items) {
        const size = blockSize(item, options);
        plan.push({ item, col, r, size, stackIndex, noteRow: size.note ? r + size.height - 1 : null });
        stackWidth = Math.max(stackWidth, size.width);
        r += size.height + 1;
      }
      cursorCol = col + stackWidth + 1;
    });

    // 2. Line up note rows of side-by-side tables so none is split
    if (options.collapse) alignNoteRows(plan, threshold);

    // 3. Write the band
    let bandBottom = bandTop;
    for (const p of plan) {
      const { item, col, r, size } = p;
      if (item.isKpi) {
        writeKpiBlock(ws, r, col, item, widths, style);
        for (let i = 0; i < size.height; i++) mark(r + i, "other");
      } else {
        writeDataTable(ws, r, col, item, nameTable(item.title), widths, style, options);
        const n = item.rows.length - 1;
        for (let i = 0; i < 2 + n; i++) mark(r + i, i >= 2 + threshold ? "deep" : "other");
        if (p.noteRow) {
          for (let x = r + 2 + n; x < p.noteRow; x++) mark(x, "deep");  // padding lined up with a neighbour
          mark(p.noteRow, "other");
        }
        tables.push({ item, top: r, col, n, noteRow: p.noteRow });
      }
      bandBottom = Math.max(bandBottom, r + size.height + 1);
    }
    bandTop = bandBottom;
  }

  let grouped = false;
  if (options.collapse) {
    // Collapse only rows that hold nothing but a long table's extra rows
    rowUse.forEach((use, r) => {
      if (use && use.deep && !use.other) {
        const excelRow = ws.getRow(r);
        excelRow.outlineLevel = 1;
        excelRow.hidden = true;
        grouped = true;
      }
    });

  }

  // Note row at the bottom of each table: totals and / or the "x of y rows shown" line
  // next to its [+] / [-] button
  for (const t of tables) {
    if (!t.noteRow) continue;
    let hidden = 0;
    if (options.collapse) {
      for (let i = threshold; i < t.n; i++) if (ws.getRow(t.top + 2 + i).hidden) hidden++;
    }
    if (t.item.totals) writeTotalsRow(ws, t.noteRow, t);
    const parts = [];
    if (t.item.totals) parts.push({ text: "Total", font: { bold: true } });
    if (hidden) {
      parts.push({
        text: `${parts.length ? " · " : ""}${t.n - hidden} of ${t.n} rows shown – use [+] / [–] on the left`,
        font: { italic: true, size: 9, color: { argb: COLORS.note } },
      });
    }
    if (!parts.length) continue;
    // The label sits in the first column, or just right of the table when that column has a total
    const labelCol = t.item.totals && t.item.totals[0] ? t.col + t.item.header.length : t.col;
    ws.getCell(t.noteRow, labelCol).value = { richText: parts };
  }

  return { nextRow: bandTop, tableCount: tables.length, grouped };
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

/*
 * exportData with retries: throttling, timeouts and gateway hiccups often pass on a
 * second try. Errors that say export isn't possible are not retried.
 */
async function exportWithRetry(visual, isCancelled) {
  for (let attempt = 1; ; attempt++) {
    try {
      return await visual.exportData(models.ExportDataType.Summarized, MAX_EXPORT_ROWS);
    } catch (e) {
      const permanent = NON_DATA_TYPES.includes(visual.type) ||
        /not supported|unsupported|disabled|not allowed|not permitted|permission/i.test(errorText(e));
      if (permanent || attempt >= EXPORT_ATTEMPTS) throw e;
      if (isCancelled()) throw { cancelled: true };
      log(`Retrying "${visual.title || visual.name}" (attempt ${attempt + 1} of ${EXPORT_ATTEMPTS}) after: ${errorText(e)}`);
      await new Promise((resolve) => setTimeout(resolve, 1000 * attempt * attempt));
    }
  }
}

/* ── Matrices as cross-tabs ───────────────────────────────────────────── */

/* A CSV header matches a field when it is the field's name, or an aggregation of it ("Sum of Sales") */
function matchesField(header, name) {
  const h = header.trim().toLowerCase(), n = name.trim().toLowerCase();
  return h === n || h.endsWith(` of ${n}`);
}

/*
 * Which fields of a matrix are on rows, on columns and in values. Needs the Power BI
 * report-authoring library (loaded in index.html); returns null when it isn't there
 * or the visual doesn't expose those roles.
 */
async function matrixFields(visual) {
  if (typeof visual.getCapabilities !== "function" || typeof visual.getDataFields !== "function") return null;
  const capabilities = await visual.getCapabilities();
  const roles = (capabilities.dataRoles || []).map((r) => r.name);
  const role = (re) => roles.find((name) => re.test(name));
  const rowsRole = role(/^rows?$/i), columnsRole = role(/^columns?$/i), valuesRole = role(/^values?$/i);
  if (!rowsRole || !columnsRole || !valuesRole) return null;
  const names = async (r) => (await visual.getDataFields(r))
    .map((f) => f.column || f.measure || f.hierarchyLevel || f.hierarchy || "");
  return { rows: await names(rowsRole), columns: await names(columnsRole), values: await names(valuesRole) };
}

/*
 * Turns a matrix's flat export (one row per row-item × column-item) back into the grid
 * Power BI shows: one row per row-item, one column per column-item (× value).
 * Returns the new rows, or null to keep the flat table (a note says why when useful).
 */
async function crossTab(visual, rows, title, notes) {
  let fields;
  try {
    fields = await matrixFields(visual);
  } catch (e) {
    log(`Matrix layout of "${title}" unavailable: ${errorText(e)}`);
    return null;
  }
  if (!fields || !fields.columns.length || !fields.values.length) return null;   // no column fields: already a table

  const header = rows[0];
  const indexOf = (name) => header.findIndex((h) => matchesField(h, name));
  const rowIdx = fields.rows.map(indexOf), colIdx = fields.columns.map(indexOf), valIdx = fields.values.map(indexOf);
  if ([...rowIdx, ...colIdx, ...valIdx].some((i) => i < 0)) {
    notes.push(`"${title}": kept as a flat table – some matrix fields are renamed in the visual`);
    return null;
  }

  const columnKeys = [];
  const seenKeys = new Set();
  const byRow = new Map();
  for (const r of rows.slice(1)) {
    const ck = colIdx.map((i) => r[i]).join(" · ");
    if (!seenKeys.has(ck)) { seenKeys.add(ck); columnKeys.push(ck); }
    const rk = rowIdx.map((i) => r[i]).join("\u0001");
    if (!byRow.has(rk)) byRow.set(rk, { dims: rowIdx.map((i) => r[i]), cells: new Map() });
    valIdx.forEach((vi, k) => byRow.get(rk).cells.set(`${ck}\u0002${k}`, r[vi]));
  }
  if (columnKeys.length * valIdx.length > MAX_CROSSTAB_COLUMNS) {
    notes.push(`"${title}": kept as a flat table – as a grid it would need ${columnKeys.length * valIdx.length} columns`);
    return null;
  }

  const newHeader = rowIdx.map((i) => header[i]);
  columnKeys.forEach((ck) => valIdx.forEach((vi) => newHeader.push(valIdx.length === 1 ? ck : `${ck} · ${header[vi]}`)));
  const data = [...byRow.values()].map(({ dims, cells }) => {
    const out = [...dims];
    columnKeys.forEach((ck) => valIdx.forEach((_, k) => out.push(cells.get(`${ck}\u0002${k}`) ?? "")));
    return out;
  });
  return [newHeader, ...data];
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
 * Reads one page: filter lines, exportable items and notes. Visuals are exported one
 * at a time (Power BI rejects concurrent exportData calls).
 * options.excludedVisuals: Set of visual names to leave out (current-page mode only).
 */
async function collectPage(page, options, onProgress, isCancelled) {
  const visuals = (await page.getVisuals()).sort((a, b) =>
    (a.layout.y - b.layout.y) || (a.layout.x - b.layout.x));

  const filterLines = [];
  if (options.includeFilters) {
    const reportFilters = describeFilters(await embeddedReport.getFilters());
    const pageFilters = describeFilters(await page.getFilters());
    if (reportFilters !== "none") filterLines.push(`Report filters: ${reportFilters}`);
    if (pageFilters !== "none") filterLines.push(`Page filters: ${pageFilters}`);
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

    // Slicers go into the filter lines as "Field: selection"
    if (v.type === "slicer") {
      if (!options.includeFilters) continue;
      try {
        const state = await v.getSlicerState();
        const selection = describeFilters(state.filters);
        if (selection !== "none") {
          filterLines.push(selection);
        } else {
          // Nothing selected: name the slicer's field (e.g. "Region: All")
          const target = (state.targets && state.targets[0]) || {};
          const field = target.column || target.hierarchyLevel || target.measure || target.hierarchy ||
            await slicerFieldNames(v, title);
          filterLines.push(`${field}: All`);
        }
      } catch (e) {
        notes.push(`Could not read the selection of "${title}": ${errorText(e)}`);
      }
      continue;
    }

    // Every other visual: try to export. Decorative ones (text box, shape, button…)
    // usually have no data; they are included only if Power BI returns some.
    try {
      const result = await exportWithRetry(v, isCancelled);
      const parsed = parseCsv(result.data);
      // Power BI stops at its row limit without saying so
      const truncated = parsed.length - 1 >= MAX_EXPORT_ROWS;
      let rows = cleanRows(parsed);
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

      if (!isKpi && options.hideEmptyRows) {
        const cleaned = removeEmptyRows(rows);
        if (cleaned.removed) notes.push(`"${title}": left out ${cleaned.removed} row(s) with no values`);
        rows = cleaned.rows;
      }

      // Matrices: rebuild the rows × columns layout Power BI shows (export comes flat)
      if (!isKpi && options.crossTab && v.type === "pivotTable") {
        const pivoted = await crossTab(v, rows, title, notes);
        if (pivoted) rows = pivoted;
      }

      // Visual-level filters (e.g. "Top 10 by Sales") belong to this visual only
      const titleParts = [title];
      if (options.includeFilters) {
        try {
          const own = describeFilters(await v.getFilters());
          if (own !== "none") titleParts.push(`Filter: ${own}`);
        } catch (e) {
          // not every visual type exposes its filters
        }
      }
      if (truncated) {
        titleParts.push(`first ${MAX_EXPORT_ROWS.toLocaleString()} rows only`);
        notes.push(`"${title}": reached Power BI's limit of ${MAX_EXPORT_ROWS.toLocaleString()} rows – the visual may have more data than was exported`);
      }

      items.push({ title, displayTitle: titleParts.join("  ·  "), type: v.type, layout: v.layout, rows, isKpi, truncated });
      log(`Exported "${title}" (${rows.length - 1} rows${truncated ? ", row limit reached" : ""})`);
    } catch (e) {
      if (e && e.cancelled) throw e;
      if (NON_DATA_TYPES.includes(v.type)) continue;
      notes.push(`Export failed for "${title}" (${v.type}): ${errorText(e)}`);
      log(`Export failed for "${title}": ${errorText(e)}`);
    }
  }
  return { filterLines, items, notes };
}

/* ── Writing a sheet ──────────────────────────────────────────────────── */
function writePageSheet(wb, sheetName, page, data, options, branding, logo, nameTable) {
  const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
  const widths = makeWidthTracker();
  const style = { accent: toArgb(branding.accent), tableStyle: branding.tableStyle || "TableStyleMedium2" };
  let row = 1;

  // Logo across the first rows, then the text header below it
  if (logo) {
    const height = 48;
    const width = Math.round(height * (logo.width / logo.height || 1));
    ws.addImage(logo.id, { tl: { col: 0, row: 0 }, ext: { width, height }, editAs: "oneCell" });
    row = 4;
  }
  if (branding.companyName) {
    const c = ws.getCell(row++, 1);
    c.value = branding.companyName;
    c.font = { bold: true, size: 11, color: { argb: "FF5F6B7A" } };
  }
  const head = ws.getCell(row++, 1);
  head.value = `${currentReportName} – ${page.displayName}`;
  head.font = { bold: true, size: 18, color: { argb: style.accent } };
  const sub = ws.getCell(row++, 1);
  sub.value = `Exported on ${new Date().toLocaleString()}`;
  sub.font = { size: 10, color: { argb: COLORS.note } };
  if (options.includeFilters) {
    (data.filterLines.length ? data.filterLines : ["Filters: none"]).forEach((line) => {
      const c = ws.getCell(row++, 1);
      c.value = line;
      c.font = { size: 10, italic: true };
    });
  }
  row++;

  let layout = { nextRow: row, tableCount: 0, grouped: false };
  if (data.items.length) {
    // Values and totals are settled first: a totals row changes a table's height
    data.items.filter((it) => !it.isKpi).forEach((it) => prepareTable(it, options));
    layout = writePageLayout(ws, data.items, row, widths, options, style, nameTable);
  } else {
    ws.getCell(row, 1).value = "No visual on this page could be exported.";
    ws.getCell(row, 1).font = { italic: true, color: { argb: COLORS.note } };
    layout.nextRow = row + 2;
  }
  row = layout.nextRow;

  // [+] / [-] buttons sit below each collapsed group, next to the table's note line
  if (layout.grouped) {
    ws.properties.outlineLevelRow = 1;
    ws.properties.outlineProperties = { summaryBelow: true, summaryRight: false };
  }

  if (options.includeNotes && data.notes.length) {
    ws.getCell(row, 1).value = "Export notes";
    ws.getCell(row++, 1).font = { bold: true };
    data.notes.forEach((n) => {
      const c = ws.getCell(row++, 1);
      c.value = n;
      c.font = { size: 9, color: { argb: COLORS.note } };
    });
    row++;
  }

  // Footer: only the user's own footer text from branding settings (optional)
  const footer = [branding.footer].filter(Boolean);
  footer.forEach((text) => {
    const c = ws.getCell(row++, 1);
    c.value = text;
    c.font = { size: 9, italic: true, color: { argb: COLORS.note } };
  });

  widths.apply(ws);
  return { tables: layout.tableCount, kpis: data.items.length - layout.tableCount };
}

/* ── Orchestrator ─────────────────────────────────────────────────────── */
/*
 * pages: Power BI Page objects to export, in order.
 * onProgress({ pageIndex, pageCount, pageName, visualIndex, visualCount, visualTitle })
 * Returns { fileName, sheets, tables, kpis, notes } or throws { cancelled: true }.
 */
async function exportToExcel(pages, options, branding, onProgress, isCancelled) {
  await refreshEmbedToken();   // a multi-page export can outlast an old token
  const original = await embeddedReport.getActivePage();
  const collected = [];
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

  const wb = new ExcelJS.Workbook();
  wb.creator = cfg.productName || "NeenOpal Export for Power BI";
  wb.calcProperties.fullCalcOnLoad = true;   // totals rows are formulas
  let logo = null;
  if (branding.logo) {
    const extension = /^data:image\/png/i.test(branding.logo) ? "png" : "jpeg";
    logo = { id: wb.addImage({ base64: branding.logo, extension }), width: branding.logoWidth, height: branding.logoHeight };
  }

  const nameSheet = makeSheetNamer();
  const nameTable = makeTableNamer();
  const totals = { tables: 0, kpis: 0, notes: [] };
  for (const { page, data } of collected) {
    const sheetName = collected.length === 1 ? nameSheet(page.displayName || "Page") : nameSheet(page.displayName);
    const counts = writePageSheet(wb, sheetName, page, data, options, branding, logo, nameTable);
    totals.tables += counts.tables;
    totals.kpis += counts.kpis;
    data.notes.forEach((n) => totals.notes.push(collected.length > 1 ? `${page.displayName}: ${n}` : n));
  }

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