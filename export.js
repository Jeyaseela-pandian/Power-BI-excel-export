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
 *  4. Export notes and footer.
 * Uses globals from app.js: embeddedReport, currentReportName, models, log, errorText,
 * describeFilters, NON_DATA_TYPES, cfg.
 */

const KPI_TYPES = ["card", "multiRowCard", "kpi", "cardVisual", "gauge"];
const PX_PER_COL = 110;        // Power BI page pixels per Excel column (for x alignment)
const PX_TOLERANCE = 10;       // overlap smaller than this does not count as overlapping
const PAGE_RENDER_TIMEOUT_MS = 20000;


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
 * Size of a visual's block in Excel cells. Long tables get one extra row at the
 * bottom for the "x of y rows shown" note when collapsing is on.
 */
function blockSize(item, options) {
  if (item.isKpi) return { width: 2, height: 1 + item.rows[0].length, note: false };
  const n = item.rows.length - 1;
  const note = !!(options && options.collapse && n > options.visibleRows);
  return { width: item.rows[0].length, height: n + 2 + (note ? 1 : 0), note }; // title + header + data (+ note)
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
function writeDataTable(ws, row, col, item, tableName, widths, style) {
  const header = uniqueHeaders(item.rows[0]);
  const data = item.rows.slice(1).map((r) => header.map((_, i) => toCellValue(r[i])));

  styleTitle(ws.getCell(row, col), item.title, style.accent);

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

  header.forEach((h, ci) => widths.note(col + ci, h));
  item.rows.slice(1).forEach((r) => r.forEach((v, ci) => widths.note(col + ci, v)));
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
        writeDataTable(ws, r, col, item, nameTable(item.title), widths, style);
        const n = item.rows.length - 1;
        for (let i = 0; i < 2 + n; i++) mark(r + i, i >= 2 + threshold ? "deep" : "other");
        if (p.noteRow) {
          for (let x = r + 2 + n; x < p.noteRow; x++) mark(x, "deep");  // padding lined up with a neighbour
          mark(p.noteRow, "other");
        }
        tables.push({ top: r, col, n, noteRow: p.noteRow });
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

    // Note line at the bottom of each collapsed table, next to its [+] / [-] button
    for (const t of tables) {
      if (!t.noteRow) continue;
      let hidden = 0;
      for (let i = threshold; i < t.n; i++) if (ws.getRow(t.top + 2 + i).hidden) hidden++;
      if (!hidden) continue;
      const c = ws.getCell(t.noteRow, t.col);
      c.value = `${t.n - hidden} of ${t.n} rows shown – use [+] / [–] on the left`;
      c.font = { italic: true, size: 9, color: { argb: COLORS.note } };
    }
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
      items.push({ title, type: v.type, layout: v.layout, rows, isKpi });
      log(`Exported "${title}" (${rows.length - 1} rows)`);
    } catch (e) {
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