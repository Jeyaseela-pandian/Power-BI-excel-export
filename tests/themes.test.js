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

test("the page's title text box is found ('Warehouse Fulfillment Dashboard'); other text boxes aren't", () => {
  const { json } = loadApp();
  const box = (text, size, pos) => ({ visualType: "textbox", textBox: { text, lines: 1, size }, position: pos });
  const chart = { visualType: "clusteredColumnChart", position: { x: 0, y: 142, width: 768, height: 365 } };
  const title = (visuals) => json(`pageTitleText({ visuals: ${JSON.stringify(Object.fromEntries(visuals.map((v, i) => ["p" + String.fromCharCode(1) + "v" + i, v])))} }, { name: "p" })`);
  assert.strictEqual(title([chart, box("Warehouse Fulfillment Dashboard", 28, { x: 384, y: 0, width: 616, height: 77 })]), "Warehouse Fulfillment Dashboard");
  // not a title: small type, a sentence, a small badge, below a chart, or a page with no charts
  assert.strictEqual(title([chart, box("Warehouse Fulfillment Dashboard", 11, { x: 384, y: 0, width: 616, height: 77 })]), null);
  assert.strictEqual(title([chart, box("Upgrade to get more out of this report.", 20, { x: 0, y: 0, width: 900, height: 60 })]), null);
  assert.strictEqual(title([chart, box("Free", 20, { x: 1200, y: 0, width: 60, height: 40 })]), null);
  assert.strictEqual(title([chart, box("Notes", 20, { x: 0, y: 600, width: 700, height: 40 })]), null);
  assert.strictEqual(title([box("Welcome", 28, { x: 0, y: 0, width: 700, height: 60 })]), null);
});

test("a text box's text and type size are read from the report (classic layout)", () => {
  const { json } = loadApp();
  const objects = { general: [{ properties: { paragraphs: [{ textRuns: [{ value: "Warehouse Fulfillment Dashboard", textStyle: { fontWeight: "bold", fontSize: "28pt" } }] }] } }] };
  const one = json(`textBoxText(${JSON.stringify(objects)})`);
  assert.deepStrictEqual([one.text, one.lines, one.size, one.bold, one.runs.length], ["Warehouse Fulfillment Dashboard", 1, 28, true, 1]);
  // two parts in two looks (Corporate Spend): 18pt in "Segoe (Bold)", then the default type
  const two = json(`textBoxText(${JSON.stringify({ general: [{ properties: { paragraphs: [{ textRuns: [
    { value: "CORPORATE SPEND SAMPLE PBIX ", textStyle: { fontFamily: "Segoe (Bold)", fontSize: "18pt" } },
    { value: "YTD IT Spend Trend Analysis", textStyle: { fontFamily: "Segoe UI" } }] }] } }] })})`);
  assert.deepStrictEqual([two.text, two.size, two.bold, two.face], ["CORPORATE SPEND SAMPLE PBIX YTD IT Spend Trend Analysis", 18, true, "Segoe UI"]);
  assert.deepStrictEqual(two.runs.map((r) => [r.text, r.size, r.bold, r.face]),
    [["CORPORATE SPEND SAMPLE PBIX ", 18, true, "Segoe UI"], ["YTD IT Spend Trend Analysis", null, false, "Segoe UI"]]);
  assert.strictEqual(json(`textBoxText({})`), null);
});

test("the title text box gets its own row above the visuals, centred as in the report; the heading keeps the page's name", () => {
  const { ctx, run, json } = loadApp();
  const box = { visualType: "textbox", textBox: { text: "Warehouse Fulfillment Dashboard", lines: 1, size: 28, bold: true, color: null, face: null }, position: { x: 384, y: 0, width: 616, height: 77 } };
  const chart = { visualType: "clusteredColumnChart", position: { x: 0, y: 142, width: 1440, height: 365 } };
  assert.strictEqual(json(`pageTitleBox({ visuals: { ["p" + String.fromCharCode(1) + "t"]: ${JSON.stringify(box)}, ["p" + String.fromCharCode(1) + "c"]: ${JSON.stringify(chart)} } }, { name: "p" }).align`), "center");
  // written: one merged row, the text in its own bold type
  const cells = {};
  ctx.WS = { mergeCellsWithoutStyle: (r1, c1, r2, c2) => { cells.merged = [r1, c1, r2, c2]; }, getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}) };
  run(`writeTitleBox(WS, 9, 2, 20, ${JSON.stringify({ ...box.textBox, align: "center" })})`);
  assert.deepStrictEqual(cells.merged, [9, 2, 9, 20]);
  assert.strictEqual(cells["9,2"].value, "Warehouse Fulfillment Dashboard");
  assert.deepStrictEqual([cells["9,2"].font.size, cells["9,2"].font.bold, cells["9,2"].alignment.horizontal], [26, true, "center"]);
});

test("a title text box in two looks is written in both: big and bold, then small", () => {
  const { ctx, run } = loadApp();
  const cells = {};
  ctx.WS = { mergeCellsWithoutStyle() {}, getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}) };
  ctx.BOX = { text: "CORPORATE SPEND SAMPLE PBIX YTD IT Spend Trend Analysis", lines: 1, size: 18, bold: true, face: "Segoe UI", align: "left",
    runs: [{ text: "CORPORATE SPEND SAMPLE PBIX ", size: 18, bold: true, face: "Segoe UI" }, { text: "YTD IT Spend Trend Analysis", size: null, bold: false, face: "Segoe UI" }] };
  run("writeTitleBox(WS, 9, 2, 20, BOX)");
  const rich = cells["9,2"].value.richText;
  assert.deepStrictEqual(rich.map((r) => [r.text, r.font.size, r.font.bold]), [["CORPORATE SPEND SAMPLE PBIX ", 18, true], ["YTD IT Spend Trend Analysis", 10, false]]);
});

test("pie labels: the theme's style (value and share from CY19SU12, Fluent2), else Power BI's default, the category", () => {
  const { json } = loadApp();
  const style = (tc) => json(`(() => { const tc = ${JSON.stringify(tc)}; const t = withBaseThemeAxes(namedThemeFallback(tc, null), tc);
    return chartLook({ type: "pieChart", objects: {}, container: {} }, t).labels.style; })()`);
  // no theme (the HR sample) or the oldest base theme: the category ("Female", "Male")
  assert.strictEqual(style(null), "Category");
  assert.strictEqual(style({ baseTheme: { name: "CY18SU07" } }), "Category");
  // newer base themes, by their name when the theme file can't be read (Viewers)
  for (const name of ["CY19SU12", "CY24SU10", "CY26SU05", "Fluent2-CY26SU08"]) assert.strictEqual(style({ baseTheme: { name } }), "Data value, percent of total", name);
  // the visual's own style wins
  const own = { labels: [{ properties: { labelStyle: { expr: { Literal: { Value: "'Percent of total'" } } } } }] };
  assert.strictEqual(json(`chartLook({ type: "pieChart", objects: ${JSON.stringify(own)}, container: {} }, withBaseThemeAxes(namedThemeFallback(null, null), null)).labels.style`), "Percent of total");
  // the report's design not read: as before (value and share)
  assert.strictEqual(json(`chartLook({ type: "pieChart", objects: {}, container: {} }, null).labels.style`), null);
});

test("a pie labelled by category shows the names, as Power BI does (Female, Male)", () => {
  const { ctx, run, json } = loadApp();
  ctx.T = { type: "pieChart", title: "t", rows: [["Gender", "BadHires"], ["Female", "15000"], ["Male", "17000"]] };
  run(`T.colors = columnColors(T, ["#01B8AA", "#374649"], null); T.themePalette = ["#01B8AA", "#374649"]; var pp = chartPlan(T);
    var pl = chartLook({ type: "pieChart" }, null); pl.labels.show = true; pl.labels.style = "Category"; pp.look = pl; pp.sizePx = { w: 300, h: 250 };`);
  assert.deepStrictEqual(json("pieLabelPoints(pp, pl).map((p) => p.text)"), ["Female", "Male"]);
});
