/* The differences found on Microsoft's Corporate Spend sample: a combo whose Y-axis is off, its
   stepped Plan line, labels of small and negative bars, a decomposition tree of a calculated
   measure, the anomaly detection note */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const PALETTE = ["#F2A900", "#9B0065"];
const IT_AREAS = [["IT Area", "Actual", "Plan"], ["Infrastructure", "$15160941", "$13901853"], ["BU Support", "$9761517", "$9034120"],
  ["Functional IT Area", "$5747716", "$6332988"], ["Governance", "$2013686", "$1633377"], ["Enablement", "$843681", "$495443"]];
const comboInfo = (look) => ({ rule: "measure", category: { ref: "it area", label: "it area" }, categories: [{ ref: "it area", label: "it area" }],
  measures: [{ ref: "actual", label: "actual", role: "Y" }, { ref: "plan", label: "plan", role: "Y2" }], overrides: {}, rules: [], tooltips: [],
  roles: { Category: [{ ref: "it area", label: "it area" }], Y: [{ ref: "actual", label: "actual" }], Y2: [{ ref: "plan", label: "plan" }] }, look });

/* The combo's chart XML with the given settings: { yAxisOff, stepped, labels, position } */
function combo({ yAxisOff = false, stepped = false, labels = true, position = "Auto" } = {}) {
  const { ctx, run } = loadApp();
  ctx.T = { type: "lineStackedColumnComboChart", title: "Actual and Plan by IT Area", rows: IT_AREAS };
  ctx.P = PALETTE;
  ctx.I = comboInfo(null);
  run(`T.colorInfo = I; T.colors = columnColors(T, P, I); T.themePalette = P; var c = chartPlan(T);
    var lk = chartLook({ type: T.type }, null); lk.labels.show = ${labels}; lk.labels.position = ${JSON.stringify(position)};
    lk.lines = { ...lk.lines, stepped: ${stepped}, style: "dotted", width: 2, series: {} };
    c.axes = { ...(c.axes || {}), val: { show: ${!yAxisOff} } };
    Object.assign(c, { look: lk, sizePx: { w: 600, h: 300 }, title: "t", showTitle: true });
    c.refs = { catRef: "D!A2:A6", series: c.series.map((s, i) => ({ nameRef: "D!" + "BC"[i] + "1", valRef: "D!" + "BC"[i] + "2:" + "BC"[i] + "6" })), range: () => "D!B2" };`);
  return run("chartXml(c).xml");
}
const valAxes = (xml) => (xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g) || []).filter((a) => !/<c:delete val="1"\/>/.test(a));

test("a combo whose Y-axis the report turns off: no second axis, the line on the columns' scale", () => {
  const off = combo({ yAxisOff: true });
  assert.ok(!valAxes(off).some((a) => /<c:axPos val="r"\/>/.test(a)), "no Plan axis on the right");
  assert.ok(!/<c:axId val="50030"\/>/.test(off), "the line isn't on a second axis");
  // the axis shown as before when the report shows it
  const on = combo({ yAxisOff: false });
  assert.ok(valAxes(on).some((a) => /<c:axPos val="r"\/>/.test(a)), "Plan's own axis, as before");
});

test("a stepped line is drawn as steps: flat across each category, a drop at each boundary, one legend entry", () => {
  const xml = combo({ yAxisOff: true, stepped: true });
  const scatter = (xml.match(/<c:scatterChart>[\s\S]*?<\/c:scatterChart>/) || [""])[0];
  const pieces = scatter.match(/<c:ser>[\s\S]*?<\/c:ser>/g) || [];
  // 5 flats and 4 drops; Excel draws no more points of one series than the chart has categories
  assert.strictEqual(pieces.length, 9);
  assert.ok(pieces.every((p) => (p.match(/<c:pt /g) || []).length / 2 <= 5));
  assert.ok(/<c:f>D!C1<\/c:f>/.test(pieces[0]), "the first piece carries the line's name");
  const legendOff = (xml.match(/<c:legendEntry><c:idx val="\d+"\/><c:delete val="1"\/>/g) || []).length;
  assert.ok(legendOff >= 8, "the other pieces are left out of the legend");
  // not stepped: one line, as before
  assert.ok(!/<c:scatterChart>/.test(combo({ yAxisOff: true, stepped: false })));
});

test("labels inside the bars: left out where a bar is too small to hold them, as Power BI does", () => {
  const xml = combo({ yAxisOff: true, position: "InsideBase" });
  const bars = (xml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/) || [""])[0];
  const off = [...bars.matchAll(/<c:dLbl><c:idx val="(\d+)"\/><c:delete val="1"\/><\/c:dLbl>/g)].map((m) => Number(m[1]));
  assert.ok(off.includes(4), "Enablement ($0.8M, a sliver): left out");
  assert.ok(!off.includes(0) && !off.includes(1) && !off.includes(2), "Infrastructure, BU Support, Functional IT Area: shown");
  // labels outside the bars (a clustered column chart's, by default): all shown
  const { ctx, run } = loadApp();
  ctx.T = { type: "clusteredColumnChart", title: "t", rows: IT_AREAS.map((r) => r.slice(0, 2)) };
  ctx.P = PALETTE;
  run(`T.colors = columnColors(T, P, null); T.themePalette = P; var o = chartPlan(T); var ol = chartLook({ type: T.type }, null); ol.labels.show = true;
    Object.assign(o, { look: ol, sizePx: { w: 600, h: 300 }, title: "t", showTitle: true });
    o.refs = { catRef: "D!A2:A6", series: o.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B6" })), range: () => "D!B2" };`);
  assert.ok(!/<c:delete val="1"\/><\/c:dLbl>/.test(run("chartXml(o).xml")));
});

test("negative bars' labels get room before the names (the axis runs on below zero)", () => {
  const { ctx, run } = loadApp();
  ctx.T = { type: "clusteredBarChart", title: "Var Plan by Country/Region", rows: [["Country/Region", "Var Plan"], ["Germany", "$1900000"], ["Ireland", "-$380000"], ["France", "-$460000"], ["Italy", "-$610000"]] };
  ctx.P = PALETTE;
  run(`T.colors = columnColors(T, P, null); T.themePalette = P; var b = chartPlan(T); var bl = chartLook({ type: T.type }, null); bl.labels.show = true;
    Object.assign(b, { look: bl, sizePx: { w: 420, h: 360 }, title: "t", showTitle: true });
    b.refs = { catRef: "D!A2:A5", series: b.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B5" })), range: () => "D!B2" };`);
  const xml = run("chartXml(b).xml");
  const min = Number((xml.match(/<c:valAx>[\s\S]*?<c:min val="([^"]+)"/) || [])[1]);
  const max = Number((xml.match(/<c:valAx>[\s\S]*?<c:max val="([^"]+)"/) || [])[1]);
  const plotPx = 420 * 0.6, labelPx = 40;
  assert.ok(((-610000 - min) / (max - min)) * plotPx >= labelPx * 0.9, `room left of Italy's bar: axis from ${min}`);
});

test("a decomposition tree of a calculated measure: Power BI's values when it gives them, else the table as before", () => {
  const { ctx, json } = loadApp();
  ctx.INFO = {
    roles: {
      Analyze: [{ ref: "var plan", label: "var plan", field: { kind: "measure", entity: "Fact", prop: "Var Plan" } }],
      ExplainBy: [{ ref: "sales region", label: "sales region", queryRef: "Country Region.Sales Region", field: { kind: "column", entity: "Country Region", prop: "Sales Region" } },
        { ref: "period", label: "period", queryRef: "Date.Period", field: { kind: "column", entity: "Date", prop: "Period" } }],
    },
    expansion: [{ roles: ["ExplainBy"], levels: [{ queryRefs: ["Country Region.Sales Region"], isPinned: true }, { queryRefs: ["Date.Period"], isPinned: true }],
      root: { children: [{ identityValues: [{ Literal: { Value: "'Europe'" } }], isToggled: true, children: [] }] } }],
    barsPerLevel: 2,
  };
  ctx.ROWS = [["Var Plan", "Sales Region", "Period"], ["1000.50", "Europe", "12"], ["500.25", "Europe", "11"], ["250.00", "Europe", "10"], ["-90.00", "North America", "12"]];
  const sep = "\u0001";
  ctx.VALUES = { total: 2000000.5, byPath: { europe: 2129759.32, "north america": -129758.82, [`europe${sep}12`]: 2129759.32, [`europe${sep}11`]: 1161729.7, [`europe${sep}10`]: 607987.94 } };
  const tree = json("decompositionTreeItem({ type: 'decompositionTreeVisual', title: 't', rows: ROWS, treeValues: VALUES }, INFO)");
  assert.strictEqual(tree.type, "decompositionTree");
  assert.deepStrictEqual(tree.rows.slice(1).map((r) => [r[1], r[2]]),
    [["", "2000000.50"], ["Europe", "2129759.32"], ["12", "2129759.32"], ["11", "1161729.70"], ["10", "607987.94"], ["North America", "-129758.82"]],
    "every period listed (values that don't add up aren't put together)");
  assert.match(tree.captionText, /values from Power BI/);
  // no values from Power BI: the table it was, as before
  assert.strictEqual(json("decompositionTreeItem({ type: 'decompositionTreeVisual', title: 't', rows: ROWS }, INFO).type"), "decompositionTreeVisual");
});

test("the anomaly detection on a line chart is read from the report, for its note", () => {
  const { json } = loadApp();
  const visual = (show) => ({ type: "lineChart", roles: {}, objects: { anomalyDetection: [{ properties: { show: { expr: { Literal: { Value: show } } } } }] },
    container: {}, dataPoint: [], bubbles: [], values: [] });
  assert.strictEqual(json(`visualColorInfo(${JSON.stringify(visual("true"))}, null).anomalies`), true);
  assert.strictEqual(json(`visualColorInfo(${JSON.stringify(visual("false"))}, null).anomalies`), false);
});
