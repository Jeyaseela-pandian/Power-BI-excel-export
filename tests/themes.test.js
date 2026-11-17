/* Axis titles from the report's themes: every CY base theme turns them on, Fluent2 themes and
   reports without a theme don't, a chart's own setting wins, and newer-format reports' theme
   names are found in the website layout even when the theme files can't be read */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { run, json } = loadApp();
const OFF = { categoryAxis: [{ properties: { showAxisTitle: { expr: { Literal: { Value: "false" } } } } }] };

/* The axis titles a column chart gets with this report's themes (when the theme files can't be read) */
const titles = (themeCollection, objects = {}) =>
  json(`(() => { const a = axisSettings(${JSON.stringify(objects)}, "clusteredColumnChart", namedThemeFallback(${JSON.stringify(themeCollection)}, null)); return [a.cat.title, a.val.title]; })()`);

test("CY base themes, old and new, turn axis titles on", () => {
  for (const name of ["CY18SU07", "CY23SU08", "CY24SU02", "CY24SU10", "CY26SU05"]) {
    assert.deepStrictEqual(titles({ baseTheme: { name } }), [true, true], name);
  }
});

test("Fluent2 themes, a custom theme alone, or no theme leave them off", () => {
  assert.deepStrictEqual(titles({ baseTheme: { name: "Fluent2-CY26SU08" } }), [false, false]);
  assert.deepStrictEqual(titles({ customTheme: { name: "x.json" } }), [false, false]);
  assert.deepStrictEqual(titles(null), [false, false]);
});

test("a chart that turns its title off keeps it off", () => {
  assert.deepStrictEqual(titles({ baseTheme: { name: "CY26SU05" } }, OFF), [false, true]);
});

test("Fluent2 base themes turn subtitles on (and axis titles off); CY themes the other way round", () => {
  const look = (name) => json(`(() => { const tc = ${JSON.stringify(name ? { baseTheme: { name } } : null)};
    const t = withBaseThemeAxes(namedThemeFallback(tc, null), tc);
    const l = chartLook({ type: "lineChart", objects: {}, container: {} }, t); const a = axisSettings({}, "lineChart", t);
    return [!!l.subtitle, a.cat.title]; })()`);
  assert.deepStrictEqual(look("Fluent2-CY26SU08"), [true, false]);
  assert.deepStrictEqual(look("CY26SU05"), [false, true]);
  assert.deepStrictEqual(look(null), [false, false]);
});

test("with subtitles on, an automatic chart title splits as Power BI shows it: 'Orders' / 'by Mth and Year'", () => {
  const titles = (subtitleOn, title) => json(`(() => {
    const ws = { getCell: () => ({ font: {}, alignment: {}, border: {} }), getRow: () => ({ getCell: () => ({}) }), getColumn: () => ({ width: 10, letter: "A" }),
      mergeCells() {}, mergeCellsWithoutStyle() {}, addTable() {}, addConditionalFormatting() {}, addImage() {}, properties: {}, views: [], pageSetup: {}, headerFooter: {} };
    const look = chartLook({ type: "lineChart", objects: {}, container: {} }, ${subtitleOn ? `{ ...FALLBACK_THEME, visualStyles: [{ "*": { "*": { subTitle: [{ show: true }] } } }] }` : "null"});
    const info = { rule: "measure", category: { ref: "month", label: "month" }, categories: [{ ref: "month", label: "month" }],
      measures: [{ ref: "orders", label: "orders", role: "Y" }], overrides: {}, rules: [], tooltips: [], look,
      roles: { Category: [{ ref: "month", label: "month", field: { kind: "column", entity: "Date", prop: "Month" } }], Y: [{ ref: "orders", label: "orders", field: { kind: "measure", entity: "Sales", prop: "Orders" } }] } };
    const items = [{ type: "lineChart", title: ${JSON.stringify(title)}, name: "l1", page: "p", layout: { x: 0, y: 0, width: 500, height: 250 },
      rows: [["Month", "Orders"], ["Jan", "2100"], ["Feb", "2000"]] }];
    const lay = { fonts: makeFonts(), accent: { tab: "FF000000", cap: "FF000000" }, sheetName: "P", palette: ["#118DFF"], slicerFields: [], tableStyle: "TableStyleMedium2",
      reportColors: { palette: ["#118DFF"], theme: FALLBACK_THEME, visuals: { ["p" + String.fromCharCode(1) + "l1"]: info }, byName: {} } };
    visualInfo = () => info;
    writePageLayout(ws, items, 10, makeWidthTracker(), { visualsAs: "charts", layout: "report", collapse: true, visibleRows: 15, chartBackground: "white", tableNumbers: "report" }, lay, (t) => t);
    const p = lay.charts[0].plan;
    return [p.title, p.look.subtitle && p.look.subtitle.text];
  })()`);
  assert.deepStrictEqual(titles(true, "lineChart (l1)"), ["Orders", "by Month"]);
  assert.deepStrictEqual(titles(false, "lineChart (l1)"), ["Orders by Month", null], "subtitles off: one line, as before");
  assert.deepStrictEqual(titles(true, "Monthly orders"), ["Monthly orders", null], "the author's own title: unchanged, no subtitle");
});

/* The website layout (Viewers' route) with Power BI answering as given */
async function websiteAxes(config, explorationContent) {
  const app = loadApp();
  const layout = { sections: [{ name: "s1", displayName: "Overview", visualContainers: [{ config: JSON.stringify({ name: "v1", singleVisual: { visualType: "clusteredBarChart", objects: {} } }) }] }],
    config, explorationContent, resourcePackages: [] };
  app.ctx.fetch = async (u) => (decodeURIComponent(String(u)).includes("modelsAndExploration")
    ? { ok: true, status: 200, json: async () => ({ models: [], exploration: layout }) }
    : { ok: false, status: 404, text: async () => "" });
  Object.assign(app.ctx, { getToken: async () => "t", errorText: (e) => String(e && e.message), log() {} });
  const report = { id: "r", embedUrl: "https://x/?config=" + Buffer.from(JSON.stringify({ clusterUrl: "https://wabi.analysis.windows.net" })).toString("base64") };
  app.ctx.REPORT = report;
  const d = await app.run("readWebsiteDesign(REPORT, null)");
  const v = Object.values(d.visuals)[0];
  return [v.axes.cat.title, v.axes.val.title];
}
const content = (tc) => JSON.stringify({ explorationDocument: { report: { filePath: "definition/report.json", content: { themeCollection: tc } } } });

test("Viewers: theme names in a newer-format layout are found (an empty theme entry is ignored)", async () => {
  assert.deepStrictEqual(await websiteAxes(JSON.stringify({ themeCollection: {} }), content({ baseTheme: { name: "CY26SU05" } })), [true, true]);
  assert.deepStrictEqual(await websiteAxes("{}", JSON.stringify({ explorationDocument: { report: { content: JSON.stringify({ themeCollection: { baseTheme: { name: "CY25SU12" } } }) } } })), [true, true]);
  assert.deepStrictEqual(await websiteAxes(JSON.stringify({ themeCollection: { baseTheme: { name: "CY24SU02" } } }), undefined), [true, true]);
  assert.deepStrictEqual(await websiteAxes(JSON.stringify({ themeCollection: {} }), content({ baseTheme: { name: "Fluent2-CY26SU08" } })), [false, false]);
  assert.deepStrictEqual(await websiteAxes("{}", content({})), [false, false]);
});
