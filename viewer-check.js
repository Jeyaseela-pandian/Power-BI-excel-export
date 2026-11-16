"use strict";

/*
 * Viewer check (support): can people who can't download the report file (workspace Viewers)
 * still get the report's design, its colours and drill hierarchies? Compares, for the open
 * report:
 *  1. the report file, as exports read it today (colors.js; Contributor or higher),
 *  2. the report's current view, captured like a bookmark (any role),
 *  3. the layout Power BI's own website loads to draw the report (any role; colors.js reads
 *     it for exports when the report file can't be read).
 * Only counts and field-well names are shown: no data, titles or values.
 * Uses globals from app.js, colors.js, export.js and ui.js: $, log, errorText, toast, models,
 * embeddedReport, currentReport, loadReportColors, reportColorsMessage, fetchWebsiteLayout,
 * classicVisuals, visualInfo, fieldInfo, decodeBookmarkState, captureLiveDrill,
 * readThemeColors, FALLBACK_THEME, NON_DATA_TYPES, HIERARCHY_ROLES, parseCsv, cleanRows,
 * withLiveDrill, atDrillLevel, levelsShown, expandedLevelCols, columnColors.
 */

const DRILL_ROLES = ["Category", "Group", "Rows", "Columns", "Axis"];

/* Counts of what a set of visual infos (colors.js visualColorInfo) holds */
function designStats(visuals) {
  const list = Object.values(visuals || {});
  return {
    visuals: list.length,
    handColours: list.filter((v) => Object.keys(v.overrides || {}).length).length,
    colourRules: list.filter((v) => (v.rules || []).length || (Array.isArray(v.valueRules) && v.valueRules.length)).length,
    hierarchies: list.filter((v) => DRILL_ROLES.some((r) => ((v.roles || {})[r] || []).length > 1)).length,
    matrixStates: list.filter((v) => v.expansion).length,
  };
}

function statsLine(s) {
  return `${s.visuals} visuals · ${s.handColours} with colours set by hand · ${s.colourRules} with colour rules · ` +
    `${s.hierarchies} with drill hierarchies · ${s.matrixStates} with matrix expand/collapse state`;
}

/* 1. The report file, as exports read it today */
async function checkReportFile() {
  try {
    const rc = await loadReportColors(currentReport, null);
    return { ok: true, stats: designStats(rc.visuals), format: rc.format };
  } catch (e) {
    return { ok: false, reason: reportColorsMessage(e), code: errorText(e) };
  }
}

/* 2. The current view, captured like a bookmark */
async function checkCurrentView() {
  try {
    const bm = await embeddedReport.bookmarksManager.capture({ allPages: true });
    const es = await decodeBookmarkState(bm && bm.state);
    const sections = es && (es.sections || (es.explorationState && es.explorationState.sections));
    if (!sections) return { ok: false, reason: "the captured view had no pages" };
    const counts = { visuals: 0, fieldWells: 0, formatting: 0, drillLevel: 0, matrixStates: 0 };
    const keys = new Set();
    for (const s of Object.values(sections)) {
      for (const vc of Object.values((s && s.visualContainers) || {})) {
        const sv = vc && vc.singleVisual;
        if (!sv) continue;
        counts.visuals++;
        Object.keys(sv).forEach((k) => keys.add(k));
        if (sv.projections || sv.prototypeQuery) counts.fieldWells++;
        if (sv.objects && Object.keys(sv.objects).length) counts.formatting++;
        if (sv.activeProjections) counts.drillLevel++;
        if (sv.expansionStates) counts.matrixStates++;
      }
    }
    return { ok: true, counts, keys: [...keys].sort() };
  } catch (e) {
    return { ok: false, reason: errorText(e) };
  }
}

/* 3. The layout Power BI's website loads to draw the report (colors.js) */
async function checkWebsiteLayout() {
  try {
    const { layout, route } = await fetchWebsiteLayout(currentReport);
    let hasTheme = false;
    try { hasTheme = !!JSON.parse(layout.config || "{}").themeCollection; } catch (e) { /* no theme */ }
    return { ok: true, route, pages: layout.sections.length, stats: designStats(classicVisuals(layout, FALLBACK_THEME)), hasTheme };
  } catch (e) {
    return { ok: false, reason: errorText(e) };
  }
}

/* Runs the three checks and writes a plain-language result */
async function runViewerCheck() {
  const out = $("viewerCheckOut");
  if (!embeddedReport || !currentReport) { out.hidden = false; out.textContent = "Open a report first."; return; }
  $("viewerCheckRun").disabled = true;
  $("viewerCheckCopy").hidden = true;
  out.hidden = false;
  out.textContent = "Checking…";

  const [file, view, site] = [await checkReportFile(), await checkCurrentView(), await checkWebsiteLayout()];
  const lines = [`Viewer check · ${new Date().toISOString()}`, ""];

  lines.push("1. Report design (what exports use)");
  if (!file.ok) lines.push(`   Not readable: ${file.reason} [${file.code}]`);
  else if (file.format === "website") lines.push(`   Report file not readable; exports use the website layout (3): ${statsLine(file.stats)}`);
  else lines.push(`   Report file readable (${file.format} format): ${statsLine(file.stats)}`);
  lines.push("");

  lines.push("2. Current view (works for every role)");
  if (view.ok) {
    const c = view.counts;
    lines.push(`   ${c.visuals} visuals · ${c.fieldWells} with field wells · ${c.formatting} with formatting · ` +
      `${c.drillLevel} with drill level · ${c.matrixStates} with matrix expand/collapse state`);
    lines.push(`   Holds: ${view.keys.join(", ") || "nothing"}`);
  } else {
    lines.push(`   Not readable: ${view.reason}`);
  }
  lines.push("");

  lines.push("3. Power BI website layout (works for every role, if reachable)");
  if (site.ok) {
    lines.push(`   Readable ${site.route}: ${site.pages} pages · ${statsLine(site.stats)}`);
    lines.push(`   Theme named: ${site.hasTheme ? "yes" : "no"}`);
  } else {
    lines.push(`   Not readable${site.route ? ` (${site.route})` : ""}: ${site.reason}`);
  }
  lines.push("");

  // The verdict
  const covers = site.ok && site.stats.visuals > 0;
  const fromFile = file.ok && file.format !== "website";
  const same = covers && fromFile && JSON.stringify(site.stats) === JSON.stringify(file.stats);
  lines.push("Result");
  if (same) lines.push("   The website layout matches the report file exactly: Viewers get the same colours and drill levels.");
  else if (covers && fromFile) lines.push("   The website layout is readable but differs from the report file: share this result so it can be compared.");
  else if (covers) lines.push("   The website layout is readable for this account: exports use it for colours and drill levels.");
  else lines.push("   The website layout isn't reachable: exports use the report theme's colours. Share this result.");

  out.textContent = lines.join("\n");
  log(`Viewer check done – report file: ${file.ok ? "yes" : "no"}, current view: ${view.ok ? "yes" : "no"}, website layout: ${site.ok ? "yes" : "no"}`);
  $("viewerCheckCopy").hidden = false;
  $("viewerCheckRun").disabled = false;
}

/* The colours the export gives a visual, in words: "by Region: East #118DFF, …" */
function colourSummary(colors, header) {
  if (!colors) return "none (cards, tables, or no palette)";
  const hex = (argb) => "#" + String(argb || "").slice(-6);
  const list = (entries) => entries.slice(0, 6).map((e) => `${e.label} ${hex(e.argb)}`).join(", ") +
    (entries.length > 6 ? `, +${entries.length - 6} more` : "");
  const pairs = (map) => Object.entries(map).map(([label, argb]) => ({ label, argb }));
  if (colors.tiles) return `each tile / slice: ${list(pairs(colors.tiles.colorOf))}`;
  if (colors.rowKey) return `by ${header[colors.rowKey.col]}: ${list(pairs(colors.rowKey.colorOf))}`;
  if (colors.legend && colors.legend.entries) return `per row: ${list(colors.legend.entries)}`;
  if (colors.perCol) return "per row, from a colour rule";
  const series = Object.entries(colors.series || {}).map(([ci, argb]) => ({ label: header[ci], argb }));
  return series.length ? `per series: ${list(series)}` : "none";
}

/*
 * Visual check (support): for each visual on the current page, the field-well levels the
 * export reads (marked shown / hidden as the layout says), the drill state of the report as it
 * is now, the columns Power BI exports, and what the export does with them: the levels it
 * shows and the colours it gives (the first few category names with their colours). Runs the
 * export's own steps on a copy; nothing is changed or downloaded.
 */
async function runDrillCheck() {
  const out = $("viewerCheckOut");
  if (!embeddedReport || !currentReport) { out.hidden = false; out.textContent = "Open a report first."; return; }
  $("drillCheckRun").disabled = true;
  $("viewerCheckCopy").hidden = true;
  out.hidden = false;
  out.textContent = "Checking…";

  const page = await embeddedReport.getActivePage();
  const visuals = await page.getVisuals();
  const theme = await readThemeColors();
  let rc = null, rcError = "";
  try {
    rc = await loadReportColors(currentReport, theme.fromReport ? { dataColors: theme.colors } : null);
  } catch (e) {
    rcError = errorText(e);
  }
  const live = (await captureLiveDrill(embeddedReport)) || {};
  const flag = (f) => (f.active === true ? " [shown]" : f.active === false ? " [hidden]" : "") + (f.field ? "" : " [no query field]");
  const name = (e) => { const f = fieldInfo(e); return f ? `${f.entity || "?"}.${f.prop}` : "?"; };

  notePageTitles([page]);
  notePageVisuals(page.name, visuals);
  const palette = rc ? rc.palette : theme.colors;
  const lines = [`Visual check · ${new Date().toISOString()} · page ${page.name} (${page.displayName})`,
    `Layout source: ${rc ? rc.format : `none (${rcError})`} · palette: ${palette.slice(0, 6).join(", ")}`];
  if (rc && rc.pages) lines.push(`Layout pages: ${rc.pages.map((p) => `${p.name} "${p.title}" (${p.visuals})`).join(", ")}`);
  if (rc) {
    // The theme the export used, and the base / custom theme the report names (axis title defaults come from them)
    const tc = rc.themeCollection || {};
    const nameOf = (t) => (t && t.name) || "-";
    lines.push(`Theme used: ${(rc.theme && rc.theme.name) || "-"} · report's base theme: ${nameOf(tc.baseTheme)} · custom theme: ${nameOf(tc.customTheme)}`);
    if (rc.format === "website" && typeof websiteThemeTries !== "undefined") {
      lines.push(`Theme files: ${websiteThemeTries.length ? websiteThemeTries.join(" | ") : "none named"}`);
    }
  }
  lines.push("");
  // Sequential on purpose: Power BI does not support concurrent exportData calls
  for (const v of visuals) {
    if (NON_DATA_TYPES.includes(v.type)) continue;
    const saved = visualInfo(rc, page.name, v.name);
    const sv = (live[page.name] || {})[v.name];
    lines.push(`${v.title || "(untitled)"} · ${v.type}`);
    if (rc && !saved) {
      const where = Object.keys(rc.visuals).filter((k) => k.split("\u0001")[1] === v.name).map((k) => k.split("\u0001")[0] || "(no page name)");
      lines.push(`   Not matched to the layout (visual ${v.name}; ${where.length ? `found on ${where.join(", ")}` : "not in the layout"})`);
      const l = v.layout || {}, here = String(page.displayName || "").trim().toLowerCase();
      const near = (rc.placed || []).filter((p) => p.title === here).map((p) => `${p.type} ${Math.round(p.x)},${Math.round(p.y)} ${Math.round(p.w)}x${Math.round(p.h)}`);
      lines.push(`   Here: ${v.type} ${Math.round(l.x)},${Math.round(l.y)} ${Math.round(l.width)}x${Math.round(l.height)} · layout has: ${near.join("; ") || "nothing on this page"}`);
    }
    for (const [role, list] of Object.entries((saved && saved.roles) || {})) {
      if ((list || []).length) lines.push(`   ${role}: ${list.map((f) => f.ref + flag(f)).join(" > ")}`);
    }
    // The colours the layout holds for this visual: set by hand (per item / measure / all) and rules
    if (saved) {
      const set = Object.entries(saved.overrides || {});
      const rules = (saved.rules || []).map((r) => `${r.scope} (${r.rule.kind}${r.rule.input ? ` on ${r.rule.input}` : ""})`);
      lines.push(`   Colours in the layout: coloured by ${saved.rule || "-"}${saved.series ? ` (legend ${saved.series.ref})` : ""} · ` +
        `set by hand: ${set.length ? set.slice(0, 10).map(([k, c]) => `${k} ${c}`).join(", ") : "none"}` +
        `${rules.length ? ` · colour rules: ${rules.join(", ")}` : ""}`);
      const f = (saved.look && saved.look.frame) || {};
      lines.push(`   Background: ${f.background ? `${f.background}${f.transparency ? ` (${f.transparency}% transparent)` : ""}` : "none (transparent: the page shows through)"}` +
        ` · title colour ${(saved.look && saved.look.title.color) || "-"}`);
    }
    if (sv && sv.activeProjections) {
      const ap = Object.entries(sv.activeProjections).map(([role, exprs]) => `${role} = ${(exprs || []).map(name).join(" + ")}`);
      lines.push(`   Now in the report: ${ap.join("; ")}${sv.expansionStates ? " · matrix expand state saved" : ""}`);
    }
    let rows = null;
    try {
      rows = cleanRows(parseCsv((await v.exportData(models.ExportDataType.Summarized, 30000)).data));
      lines.push(`   Exported columns: ${(rows[0] || []).join(", ")}`);
    } catch (e) {
      lines.push(`   Not exported: ${errorText(e)}`, "");
      continue;
    }
    // What the export does with them: the same steps, without changing anything
    const item = { type: v.type, title: v.title, name: v.name, page: page.name, rows, liveDrill: sv || null };
    const info = withLiveDrill(item, saved);
    // Settings as Power BI shows them on screen, and the title the export gives the chart
    if (chartSpecOf(v.type)) {
      const ls = await liveVisualSettings(v);
      const say = (x) => (x === undefined ? "not given" : JSON.stringify(x));
      const ax = (info && info.axes) || {};
      lines.push(`   On screen: title text ${say(ls && ls.titleText)} · export's axis titles: category ${say((ax.cat || {}).title)}, value ${say((ax.val || {}).title)}`);
      if (!v.title) lines.push(`   Automatic title: ${autoTitle({ ...item, rows }, atDrillLevel(item, info)) || "couldn't be built (kept as before)"}`);
    }
    if (info && info.roles) {
      const shown = HIERARCHY_ROLES.filter((r) => ((info.roles[r] || []).length > 1))
        .map((r) => `${r} ${levelsShown(item, info, r).map((f) => f.ref).join(" + ") || "?"}`);
      const together = expandedLevelCols(item, atDrillLevel(item, info), rows[0] || []);
      if (shown.length) lines.push(`   Export shows: ${shown.join("; ")}${together ? " (levels together, e.g. \"East New York\")" : ""}`);
      // The roll-up to that level, as the export runs it (on a copy)
      if (shown.length) {
        const copy = { ...item, rows: rows.map((r) => [...r]) };
        try {
          const filters = [...await filtersOf(embeddedReport), ...await filtersOf(page), ...await filtersOf(v)];
          const note = await rollUpToDrillLevel(copy, info, filters);
          lines.push(note || copy.drillKept
            ? `   Roll-up: every level kept${copy.siteQueryError ? ` · Power BI's query service not used: ${copy.siteQueryError}` : ""}${note ? ` · ${note}` : ""}`
            : `   Roll-up: ${rows.length - 1} → ${copy.rows.length - 1} rows, columns ${(copy.rows[0] || []).join(", ")}`);
        } catch (e) {
          lines.push(`   Roll-up: failed (${errorText(e)})`);
        }
      }
    }
    try {
      lines.push(`   Export colours: ${colourSummary(columnColors(item, palette, atDrillLevel(item, info)), rows[0] || [])}`);
    } catch (e) {
      lines.push(`   Export colours: couldn't be worked out (${errorText(e)})`);
    }
    lines.push("");
  }
  if (lines.length === 3) lines.push("No visual with data on this page.");

  // How the website layout holds this page's visuals (their shape, not their data), to see why
  // a visual isn't matched
  if (rc && rc.format === "website") {
    try {
      const { layout } = await fetchWebsiteLayout(currentReport);
      // The exploration content's outline (keys three levels down, first items of lists) and
      // what the export reads from it
      const outline = (o, depth) => {
        o = parsedJson(o) || o;
        if (Array.isArray(o)) return `[${o.length}${depth > 0 && o.length ? `: ${outline(o[0], depth - 1)}` : ""}]`;
        if (!o || typeof o !== "object") return typeof o;
        return `{${Object.keys(o).slice(0, 25).map((k) => (depth > 0 ? `${k} ${outline(o[k], depth - 1)}` : k)).join(", ")}}`;
      };
      lines.push(`Exploration content: ${typeof layout.explorationContent} · ${outline(layout.explorationContent, 4).slice(0, 1500)}`);
      const read = contentVisuals(layout.explorationContent, rc.theme, layout.sections);
      lines.push(`Read from it: ${read.length} visual(s)${read.length ? ": " + read.slice(0, 12).map((v) => `${v.title || "?"}/${v.type} ${v.name} ${Math.round(v.x)},${Math.round(v.y)}`).join("; ") : ""}`);
      const vc0 = (((layout.sections || [])[0] || {}).visualContainers || [])[0];
      if (vc0) lines.push(`First visual: ${JSON.stringify(vc0).slice(0, 300)}`);
      const here = String(page.displayName || "").trim().toLowerCase();
      const section = (layout.sections || []).find((s) => String(s.displayName || "").trim().toLowerCase() === here);
      lines.push(`Layout details: ${(rc.placed || []).length} placed in all; page keys: ${section ? Object.keys(section).join(", ") : "page not found"}`);
      for (const vc of ((section && section.visualContainers) || []).slice(0, 8)) {
        let c = vc.config, parsed = "";
        try { if (typeof c === "string") c = JSON.parse(c); } catch (e) { parsed = " (config unreadable)"; }
        const sv = (c && c.singleVisual) || {};
        lines.push(`   container keys: ${Object.keys(vc).join(", ")} · config: ${typeof vc.config}${parsed} · config keys: ${c && typeof c === "object" ? Object.keys(c).join(", ") : "-"}` +
          ` · name ${c && c.name} · type ${sv.visualType} · at ${vc.x},${vc.y} ${vc.width}x${vc.height}`);
      }
    } catch (e) {
      lines.push(`Layout details: couldn't be read (${errorText(e)})`);
    }
  }

  out.textContent = lines.join("\n");
  log("Drill check done");
  $("viewerCheckCopy").hidden = false;
  $("drillCheckRun").disabled = false;
}

$("drillCheckRun").addEventListener("click", () => runDrillCheck().catch((e) => {
  log("Drill check error: " + errorText(e));
  $("viewerCheckOut").textContent = "The check failed: " + errorText(e);
  $("drillCheckRun").disabled = false;
}));
$("viewerCheckRun").addEventListener("click", () => runViewerCheck().catch((e) => {
  log("Viewer check error: " + errorText(e));
  $("viewerCheckOut").textContent = "The check failed: " + errorText(e);
  $("viewerCheckRun").disabled = false;
}));
$("viewerCheckCopy").addEventListener("click", () => {
  navigator.clipboard.writeText($("viewerCheckOut").textContent)
    .then(() => toast("Copied", "Paste the result into the chat.", "success"))
    .catch(() => toast("Couldn't copy", "Select the text and copy it by hand.", "error"));
});
