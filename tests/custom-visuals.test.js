/* AppSource visuals drawn as they draw themselves in charts mode: the Tornado (no value axis,
   values inside the bars) and Dual KPI (two area charts, not trend cards) */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const TORNADO = "TornadoChart1452517688218";
const DUAL_KPI = "PBI_CV_3C80B1F2_09AF_4123_8E99_C3CBC46B23E0";
const OPTIONS = (visualsAs) => ({ visualsAs, layout: "report", collapse: true, visibleRows: 15, chartBackground: "white", tableNumbers: "report" });
const sheet = () => ({ getCell: () => ({ font: {}, alignment: {}, border: {} }), getRow: () => ({ getCell: () => ({}) }), getColumn: () => ({ width: 10, letter: "A" }),
  mergeCells() {}, mergeCellsWithoutStyle() {}, addTable() {}, addConditionalFormatting() {}, addImage() {}, properties: {}, views: [], pageSetup: {}, headerFooter: {} });

/* The visual through the page layout: lay (its charts, sparklines…) */
function layout(item, info, visualsAs = "charts") {
  const { ctx, run } = loadApp();
  Object.assign(ctx, { ITEM: { page: "p", name: "v", layout: { x: 0, y: 0, width: 500, height: 280 }, ...item }, INFO: info, WS: sheet(), OPTS: OPTIONS(visualsAs) });
  run(`LAY = { fonts: makeFonts(), accent: { tab: "FF000000", cap: "FF000000" }, sheetName: "P", palette: ["#118DFF"], slicerFields: [], tableStyle: "TableStyleMedium2",
      reportColors: { palette: ["#118DFF"], theme: FALLBACK_THEME, visuals: { ["p" + String.fromCharCode(1) + "v"]: INFO }, byName: {} } };
    visualInfo = () => INFO;
    writePageLayout(WS, [ITEM], 10, makeWidthTracker(), OPTS, LAY, (t) => t);`);
  return { run, ctx };
}

test("a Tornado: no value axis, gridlines or axis titles, values inside the bars, no legend title", () => {
  // the visual's settings as the report reader gives them (its defaults: nothing set)
  const f = (kind, entity, prop) => ({ ref: prop.toLowerCase(), label: null, field: { kind, entity, prop } });
  const info = loadApp().json(`visualColorInfo(${JSON.stringify({ type: TORNADO, objects: {}, container: {}, dataPoint: [], bubbles: [], values: [],
    roles: { Category: [f("column", "Sales", "Region")], Series: [f("column", "Sales", "Chain")], Values: [f("measure", "Sales", "Total Sales")] } })}, null)`);
  const { run } = layout({ type: TORNADO, title: "TC1 Tornado, Chain on each side",
    rows: [["Region", "Chain", "Total Sales"], ["East", "Fashions Direct", "$4599"], ["East", "Lindseys", "$3810"], ["West", "Fashions Direct", "$2892"], ["West", "Lindseys", "$3166"]] }, info);
  run(`P = LAY.charts[0].plan; P.refs = { catRef: "D!A2:A3", series: P.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B3" })), range: () => "D!B2" };`);
  const xml = run("chartXml(P).xml");
  const val = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)[0];
  assert.match(val, /<c:delete val="1"\/>/, "no value axis");
  assert.ok(!/<c:majorGridlines>/.test(val), "no gridlines");
  assert.ok(!/<c:title>/.test(xml.replace(/^[\s\S]*?<c:plotArea>/, "")), "no axis titles");
  assert.strictEqual((xml.match(/<c:dLblPos val="ctr"\/>/g) || []).length, 2, "values inside each side's bars");
  assert.ok(!/>Chain</.test(xml), "no legend title");
});

const DK_ROWS = [["Date", "Total Sales", "Total Units"], ["1/15/2023", "$870", "6"], ["2/15/2023", "$1120", "7"], ["3/15/2023", "$1250", "11"], ["4/15/2023", "$1192", "7"]];
const DK_INFO = { rule: "fixed", roles: { axis: [{ ref: "date", label: "date" }], topvalues: [{ ref: "total sales", label: "total sales" }], bottomvalues: [{ ref: "total units", label: "total units" }] },
  overrides: {}, rules: [], tooltips: [], settings: {} };

test("a Dual KPI in charts mode: its two area charts, the value and change over each, the axis at its ends", () => {
  const { run, ctx } = layout({ type: DUAL_KPI, title: "DK1 Dual KPI, sales on top, units below", rows: DK_ROWS }, DK_INFO);
  const g = JSON.parse(run("JSON.stringify(LAY.charts[0].plan.panels, (k, v) => (k === 'look' ? undefined : v))"));
  assert.strictEqual(g.plans.length, 2, "top and bottom");
  assert.deepStrictEqual(g.plans.map((p) => p.overlay.lines.map((l) => l.text)), [["Total Sales (+37%)", "$1,192"], ["Total Units (+16.6%)", "7"]]);
  assert.deepStrictEqual(g.plans.map((p) => [p.axisEnds.min, p.axisEnds.max]), [[870, 1250], [6, 11]]);
  assert.match(g.footer.text, /^Last \d+ days$/);
  // each panel an area chart, its axis labelled at the two ends only
  run(`G = LAY.charts[0].plan.panels; G.plans.forEach((p) => { p.refs = { catRef: "D!A2:A5", series: p.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B5" })), range: () => "D!B2" }; });`);
  const top = run("chartXml(G.plans[0]).xml");
  assert.match(top, /<c:areaChart>/);
  assert.match(top, /<c:max val="1250"\/><c:min val="870"\/>[\s\S]*?<c:majorUnit val="380"\/>/);
  const group = run(`panelGroupXml({ plan: LAY.charts[0].plan, col: 2, row: 10, wPx: 500, hPx: 280 }, ["rId1", "rId2"], 1000)`);
  assert.match(group, /<a:t>\$1,192<\/a:t>/);
  assert.ok(!run("(LAY.sparklines || []).length"), "no sparklines in charts mode");
});

test("a Dual KPI in tables mode too: its two area charts, no table; other visuals still tables", () => {
  const { run } = layout({ type: DUAL_KPI, title: "DK1", rows: DK_ROWS }, DK_INFO, "tables");
  assert.strictEqual(run("LAY.charts.length"), 1);
  assert.strictEqual(run("LAY.charts[0].plan.panels.plans.length"), 2);
  const other = layout({ type: "clusteredColumnChart", title: "t", rows: [["A", "B"], ["x", "1"], ["y", "2"]] },
    { rule: "measure", category: { ref: "a", label: "a" }, categories: [{ ref: "a", label: "a" }], measures: [{ ref: "b", label: "b", role: "Y" }], overrides: {}, rules: [], tooltips: [], roles: {} }, "tables");
  assert.ok(!other.run("(LAY.charts || []).length"), "a column chart in tables mode: a table, as before");
});
