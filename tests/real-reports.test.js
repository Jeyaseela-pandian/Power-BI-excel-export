/*
 * Real reports (.pbix files in REPORTS_DIR, Downloads by default; skipped when there are none):
 *  - every chart's axis titles as a Viewer's export decides them (theme names only) match the
 *    report's own theme files
 *  - the Artificial Intelligence Sample's key influencers and decomposition tree settings are read
 * Set REPORTS_DIR to another folder to check other reports.
 */
const test = require("node:test");
const assert = require("node:assert");
const path = require("path");
const { loadApp, readZip, designZip, reportFiles, reportFile, REPORTS_DIR } = require("./harness");

const files = reportFiles();
const none = files.length ? false : `no .pbix files in ${REPORTS_DIR}`;

/* A report's theme collection, theme files and charts (type and own settings) */
function readReport(file) {
  const z = readZip(file);
  const json = (name, enc = "utf8") => { const b = z.read(name); return b ? JSON.parse(b.toString(enc).replace(/^﻿/, "")) : null; };
  let tc = null;
  const visuals = [];
  if (z.names.includes("Report/Layout")) {
    const lay = json("Report/Layout", "utf16le");
    tc = JSON.parse(lay.config || "{}").themeCollection || null;
    for (const s of lay.sections || []) {
      for (const vc of s.visualContainers || []) {
        const sv = (JSON.parse(vc.config || "{}").singleVisual) || {};
        if (sv.visualType) visuals.push({ type: sv.visualType, objects: sv.objects || {} });
      }
    }
  } else if (z.names.includes("Report/definition/report.json")) {
    tc = json("Report/definition/report.json").themeCollection || null;
    for (const n of z.names.filter((x) => x.endsWith("/visual.json"))) {
      const v = (json(n) || {}).visual || {};
      if (v.visualType) visuals.push({ type: v.visualType, objects: v.objects || {} });
    }
  } else {
    return null;
  }
  const theme = (t) => {
    const name = String((t && t.name) || "").toLowerCase().replace(/\.json$/, "");
    if (!name) return null;
    const n = z.names.find((x) => x.startsWith("Report/StaticResources") && x.toLowerCase().split("/").pop().replace(/\.json$/, "") === name);
    try { return n ? json(n) : null; } catch (e) { return null; }
  };
  return { tc: tc || {}, base: theme((tc || {}).baseTheme), custom: theme((tc || {}).customTheme), visuals };
}

test("every chart's axis titles match the report's themes, for Admins and for Viewers", { skip: none }, () => {
  const { ctx, json } = loadApp();
  let charts = 0;
  const wrong = [];
  for (const file of files) {
    let r;
    try { r = readReport(file); } catch (e) { continue; }   // not a report file the export reads
    if (!r) continue;
    ctx.R = r;
    const found = json(`(() => {
      const fromFiles = themeFromJson(R.base, R.custom) || namedThemeFallback(R.tc, null);
      const forViewers = withBaseThemeAxes(namedThemeFallback(R.tc, null), R.tc);
      const s = (v, t) => { const a = axisSettings(v.objects, v.type, t); return a.cat.title + "/" + a.val.title; };
      return R.visuals.filter((v) => { const sp = chartSpecOf(v.type); return sp && ["bar", "line", "area", "combo", "scatter", "waterfall"].includes(sp.kind); })
        .map((v) => [v.type, s(v, fromFiles), s(v, forViewers)]);
    })()`);
    charts += found.length;
    for (const [type, files_, viewers] of found) if (files_ !== viewers) wrong.push(`${path.basename(file)}: ${type} (theme files ${files_}, Viewers ${viewers})`);
  }
  assert.ok(charts > 0, "charts were checked");
  assert.deepStrictEqual(wrong, [], `${wrong.length} of ${charts} charts differ`);
});

const aiSample = reportFile("Artificial Intelligence Sample.pbix");
test("the AI sample's key influencers and decomposition tree settings are read", { skip: aiSample ? false : "Artificial Intelligence Sample.pbix not found" }, async () => {
  const app = loadApp();
  app.ctx.ZIP = designZip(aiSample);
  const settings = await app.run(`(async () => {
    const d = await readDesign(ZIP, null);
    const ki = Object.values(d.visuals).find((v) => v.keyDrivers);
    const tree = Object.values(d.visuals).find((v) => v.barsPerLevel !== null && v.barsPerLevel !== undefined);
    return JSON.stringify({ target: ki.keyDrivers.target.ref, value: ki.keyDrivers.targetValue, explain: ki.keyDrivers.explainBy.length,
      bars: tree.barsPerLevel, levels: tree.roles.ExplainBy.filter((f) => f.active).length,
      opened: tree.expansion[0].root.children.filter((c) => c.isToggled).length });
  })()`);
  assert.deepStrictEqual(JSON.parse(settings), { target: "status", value: "Won", explain: 9, bars: 8, levels: 3, opened: 1 });
});

test("the AI sample's 'Revenue Won by Product': the report's blue, labels past the bar's end", { skip: aiSample ? false : "Artificial Intelligence Sample.pbix not found" }, async () => {
  const app = loadApp();
  app.ctx.ZIP = designZip(aiSample);
  const xml = await app.run(`(async () => {
    const d = await readDesign(ZIP, null);
    const info = Object.values(d.visuals).find((v) => v.roles && ((v.roles.Category || [])[0] || {}).ref === "product" && v.look.labels.show === true && v.overrides.all);
    const T = { type: "barChart", title: "Revenue Won by Product", rows: [["Product", "Revenue Won"], ["Design app", "$5700000"], ["Scanners", "$300000"]], colorInfo: info, themePalette: d.theme.dataColors };
    T.colors = columnColors(T, d.theme.dataColors, info);
    const p = chartPlan(T);
    Object.assign(p, { look: info.look, sizePx: { w: 360, h: 296 }, title: T.title, showTitle: true });
    p.refs = { catRef: "D!A1:A2", series: p.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B3" })), range: () => "D!B2" };
    return chartXml(p).xml;
  })()`);
  assert.match(xml, /<c:ser>[\s\S]*?<a:srgbClr val="0072BF"/, "bars in the report's default colour #0072BF");
  assert.match(xml, /<c:grouping val="clustered"\/>/);
  assert.match(xml, /<c:dLblPos val="outEnd"\/>/);
});
