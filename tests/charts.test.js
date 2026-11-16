/* Charts: every chart type is drawn, a chart that can't be drawn stays a table, and the
   chart fixes keep working (empty legends, number axes in order, second axis on lines) */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { ctx, run, json } = loadApp();
const PALETTE = ["#118DFF", "#12239E", "#E66C37"];

/* A visual's chart plan, as the export makes it without the report file */
function plan(type, rows) {
  ctx.T = { type, title: type, rows, layout: { width: 400, height: 300 } };
  ctx.P = PALETTE;
  run(`T.colors = columnColors(T, P, null); T.themePalette = P; var p = chartPlan(T);
    if (p) { Object.assign(p, { look: chartLook({ type: T.type }, null), sizePx: { w: 400, h: 300 }, showTitle: true, title: "t" }); p = withDateAxis(withNumberOrder(p)); }`);
  return run("p");
}

const CHARTS = [
  ["clusteredColumnChart", [["Month", "Sales"], ["Jan", "10"], ["Feb", "20"]]],
  ["columnChart", [["Month", "Region", "Sales"], ["Jan", "East", "10"], ["Jan", "West", "5"], ["Feb", "East", "20"], ["Feb", "West", "7"]]],
  ["lineChart", [["Month", "Sales", "Profit"], ["Jan", "10", "3"], ["Feb", "20", "4"]]],
  ["stackedAreaChart", [["Period", "Area", "Actual"], ["1", "BU", "1"], ["1", "RD", "2"], ["2", "BU", "3"], ["2", "RD", "4"]]],
  ["lineClusteredColumnComboChart", [["Month", "Sales", "Margin"], ["Jan", "10", "0.3"], ["Feb", "20", "0.4"]]],
  ["pieChart", [["Cat", "V"], ["A", "5"], ["B", "3"]]],
  ["donutChart", [["Cat", "V"], ["A", "5"], ["B", "3"], ["C", "1"]]],
  ["scatterChart", [["Name", "X", "Y"], ["A", "1", "5"], ["B", "2", "3"], ["C", "4", "1"]]],
  ["waterfallChart", [["Cat", "V"], ["A", "5"], ["B", "-3"], ["C", "4"]]],
  ["funnel", [["Stage", "V"], ["Lead", "100"], ["Won", "40"]]],
  ["treemap", [["Cat", "V"], ["A", "5"], ["B", "3"]]],
  ["RadarChart1446119667547", [["Cat", "V"], ["A", "5"], ["B", "3"], ["C", "4"]]],
  ["Histogram1445664487616", [["Bin", "Count"], ["0-10", "5"], ["10-20", "3"]]],
  ["TornadoChart1452517688218", [["Cat", "Side", "V"], ["A", "Left", "5"], ["A", "Right", "3"], ["B", "Left", "2"], ["B", "Right", "4"]]],
  ["BoxWhiskerChart1455240051538", [["Group", "Sample", "Value"], ["A", "1", "5"], ["A", "2", "7"], ["A", "3", "9"], ["B", "1", "2"], ["B", "2", "4"], ["B", "3", "8"]]],
  ["Sunburst1445472000808", [["Region", "Country", "V"], ["EU", "FR", "5"], ["EU", "DE", "3"], ["NA", "US", "9"]]],
];

test("every chart type is drawn as an Excel chart", () => {
  for (const [type, rows] of CHARTS) {
    const p = plan(type, rows);
    assert.ok(p, `${type}: a chart is planned`);
    ctx.p = p;
    assert.strictEqual(run("chartBuildError(p)"), null, `${type}: its chart part is written`);
  }
});

test("a legend whose values are all blank is drawn without it (no 'reading fmt' error)", () => {
  const p = plan("clusteredBarChart", [["Month", "Region", "Sales"], ["Jan", "", "10"], ["Feb", "", "20"]]);
  assert.strictEqual(p.series.length, 1);
  ctx.p = p;
  assert.strictEqual(run("chartBuildError(p)"), null);
});

test("a chart that can't be written is caught, so it stays a table", () => {
  const err = run(`chartBuildError({ spec: { kind: "bar", dir: "col", grouping: "clustered" }, categories: ["A"], series: [],
    look: chartLook({ type: "x" }, null), sizePx: { w: 400, h: 300 } })`);
  assert.ok(err instanceof Error || (err && err.message), "the error is returned, not thrown");
});

test("numbers on a line or area axis run in ascending order (Period 1…12)", () => {
  ctx.q = { spec: { kind: "area", grouping: "stacked" }, categories: ["12", "11", "2", "1"], catText: null, pointColors: null,
    look: run(`chartLook({ type: "stackedAreaChart" }, null)`), series: [{ name: "BU", values: [5, 4, 1, 0.5] }, { name: "R&D", values: [9, 8, 2, 1] }] };
  const r = json("withNumberOrder(q)");
  assert.deepStrictEqual(r.categories, ["1", "2", "11", "12"]);
  assert.deepStrictEqual(r.series[0].values, [0.5, 1, 4, 5]);
  // text, and an axis the report sets to Categorical, keep their order
  ctx.q2 = { ...ctx.q, categories: ["Zed", "Amy"], series: [{ name: "V", values: [1, 5] }] };
  assert.deepStrictEqual(json("withNumberOrder(q2)").categories, ["Zed", "Amy"]);
  ctx.q3 = { ...ctx.q, look: { ...ctx.q.look, cat: { ...ctx.q.look.cat, continuous: false } } };
  assert.deepStrictEqual(json("withNumberOrder(q3)").categories, ["12", "11", "2", "1"]);
});

test("a line chart's secondary y-axis values go on a second axis, with their own title", () => {
  ctx.T = { type: "lineChart", title: "t", rows: [["RecoveryDay", "Avg Mobility Score", "Avg Pain Score"], ["1", "2", "7.5"], ["3", "3.5", "6.5"], ["7", "5", "5"]] };
  ctx.I = { rule: "measure", category: { ref: "recoveryday", label: "recoveryday" }, categories: [{ ref: "recoveryday", label: "recoveryday" }],
    measures: [{ ref: "avg mobility score", label: "avg mobility score", role: "Y" }, { ref: "avg pain score", label: "avg pain score", role: "Y2" }],
    overrides: {}, rules: [], roles: {}, tooltips: [], axes: { cat: { show: true, title: true }, val: { show: true, title: true }, sec: { show: true, title: true } } };
  ctx.P = PALETTE;
  run(`T.colorInfo = I; T.colors = columnColors(T, P, I); T.themePalette = P; var p2 = chartPlan(T); p2.axes = I.axes;
    Object.assign(p2, { look: chartLook({ type: T.type }, null), sizePx: { w: 400, h: 300 }, title: "t" });
    p2.refs = { catRef: "D!A1:A2", series: p2.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B3" })), range: () => "D!B2" };`);
  assert.deepStrictEqual(json("p2.series.map((s) => [s.name, !!s.line])"), [["Avg Mobility Score", false], ["Avg Pain Score", true]]);
  const xml = run("chartXml(p2).xml");
  const axes = [...xml.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/g)].filter((m) => !/<c:delete val="1"\/>/.test(m[1]))
    .map((m) => [(m[1].match(/axPos val="(\w)"/) || [])[1], (m[1].match(/<c:title>[\s\S]*?<a:t>([^<]*)<\/a:t>/) || [])[1]]);
  assert.deepStrictEqual(axes, [["l", "Avg Mobility Score"], ["r", "Avg Pain Score"]]);
});

test("ribbon charts stay tables", () => {
  assert.strictEqual(run(`chartSpecOf("ribbonChart")`), null);
  assert.ok(run(`chartSpecOf("columnChart")`));
});
