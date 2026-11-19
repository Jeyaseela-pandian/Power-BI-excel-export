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

/* A table of rates (Readmission & Complication by Surgeon) with Power BI answering the totals as given */
async function rateTable(answer) {
  const { ctx, run, json } = loadApp();
  const f = (kind, entity, prop) => ({ ref: prop.toLowerCase(), label: null, field: { kind, entity, prop } });
  ctx.INFO = json(`visualColorInfo(${JSON.stringify({ type: "tableEx", objects: {}, container: {}, dataPoint: [], bubbles: [], values: [],
    roles: { Values: [f("column", "Surgeons", "SurgeonName"), f("measure", "Metrics", "Readmission Rate"), f("measure", "Metrics", "Complication Rate"), f("measure", "Metrics", "Revenue")] } })}, null)`);
  ctx.ITEM = { type: "tableEx", title: "Readmission & Complication by Surgeon", page: "p", name: "t",
    rows: [["SurgeonName", "Readmission Rate", "Complication Rate", "Revenue"], ["Dr. A. Sharma", "10.4%", "14.6%", "$1000"], ["Dr. C. Patel", "8.3%", "14.6%", "$500"], ["Dr. E. Johnson", "12.5%", "16.7%", "$250"]] };
  ctx.ANSWER = answer;
  ctx.ASKED = [];
  run(`visualInfo = () => INFO;
    executeDaxQuery = async (q) => { ASKED.push(q); const m = q.match(/(Readmission Rate|Complication Rate)/); return m && ANSWER[m[1]] !== undefined ? { rows: [{ "[v]": ANSWER[m[1]] }] } : { error: "no answer" }; };
    websiteQuery = async () => { throw new Error("not reached"); };`);
  await run("loadTableTotals([{ data: { items: [ITEM], queryFilters: [] } }], { reportColors: {} })");
  run("ITEM.totals = totalsPlan(ITEM, INFO)");
  return { kinds: json("ITEM.totals && ITEM.totals.kinds"), report: json("ITEM.reportTotals || null"), asked: json("ASKED.length") };
}

test("a table's Total row: Power BI's own totals for rates (10.8%, 15.3%), sums as Excel formulas as before", async () => {
  const t = await rateTable({ "Readmission Rate": 0.108, "Complication Rate": 0.153 });
  assert.strictEqual(t.asked, 2, "asked only for the two rates (Revenue adds up)");
  assert.deepStrictEqual(t.kinds, [null, "report", "report", "sum"]);
  assert.deepStrictEqual(t.report, { "readmission rate": 0.108, "complication rate": 0.153 });
  // Power BI doesn't answer: blank, as before (never a guess)
  const none = await rateTable({});
  assert.deepStrictEqual(none.kinds, [null, null, null, "sum"]);
});

test("the totals row writes Power BI's value in the column's format, and Excel's formula where it adds up", () => {
  const { ctx, run } = loadApp();
  const cells = {};
  ctx.WS = { getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}), getColumn: (c) => ({ letter: "ABCDEFG"[c - 1] }) };
  ctx.ITEM = { rows: [["SurgeonName", "Readmission Rate", "Revenue"]], totals: { kinds: [null, "report", "sum"], report: { "readmission rate": 0.108 } } };
  ctx.TABLE = { data: [[{ value: "A" }, { value: 0.104, numFmt: "0.0%" }, { value: 1000, numFmt: "#,##0" }]], formats: [null, "0.0%", "#,##0"] };
  run("writeTotalsRow(WS, 9, 2, ITEM, 3, 1, { fonts: makeFonts() }, TABLE)");
  assert.strictEqual(cells["9,3"].value, 0.108);
  assert.strictEqual(cells["9,3"].numFmt, "0.0%");
  assert.match(cells["9,4"].value.formula, /^SUBTOTAL\(9,/);
  assert.strictEqual(cells["9,2"].value, "Total");
});

test("Power BI's query service answers read right: a single total ({ M0 }) as well as rows (C)", () => {
  const { json } = loadApp();
  const descriptor = { Select: [{ Value: "M0", Name: "f0" }] };
  // one total, nothing to group by: the value under its name, with or without a schema
  assert.deepStrictEqual(json(`dsrRows({ PH: [{ DM0: [{ S: [{ N: "M0", T: 3 }], M0: 0.108 }] }] }, ${JSON.stringify(descriptor)}, 1)`), [[0.108]]);
  assert.deepStrictEqual(json(`dsrRows({ PH: [{ DM0: [{ M0: 0.153 }] }] }, ${JSON.stringify(descriptor)}, 1)`), [[0.153]]);
  // rows grouped by a column: values in C, repeats and dictionaries as before
  const rows = json(`dsrRows({ ValueDicts: { D0: ["Europe", "Asia"] }, PH: [{ DM0: [{ S: [{ N: "G0", T: 1, DN: "D0" }, { N: "M0", T: 3 }], C: [0, 5] }, { C: [1, 7] }] }] },
    { Select: [{ Value: "G0", Name: "f0" }, { Value: "M0", Name: "f1" }] }, 2)`);
  assert.deepStrictEqual(rows, [["Europe", 5], ["Asia", 7]]);
});

test("a Viewer's table totals: asked of the query service, its single-total answer written in the Total row", async () => {
  const { ctx, run, json } = loadApp();
  const f = (kind, entity, prop) => ({ ref: prop.toLowerCase(), label: null, field: { kind, entity, prop } });
  ctx.INFO = json(`visualColorInfo(${JSON.stringify({ type: "tableEx", objects: {}, container: {}, dataPoint: [], bubbles: [], values: [],
    roles: { Values: [f("column", "Surgeons", "SurgeonName"), f("measure", "Metrics", "Readmission Rate")] } })}, null)`);
  ctx.ITEM = { type: "tableEx", title: "t", page: "p", name: "t", rows: [["SurgeonName", "Readmission Rate"], ["Dr. A. Sharma", "10.4%"], ["Dr. C. Patel", "8.3%"]] };
  // no model permission (a Viewer); the query service answers { M0: 0.108 }
  run(`visualInfo = () => INFO;
    executeDaxQuery = async () => ({ error: "no Build permission" });
    websiteQuery = async (fields) => dsrRows({ PH: [{ DM0: [{ S: [{ N: "M0", T: 3 }], M0: 0.108 }] }] }, { Select: [{ Value: "M0", Name: "f0" }] }, fields.length);`);
  await run("loadTableTotals([{ data: { items: [ITEM], queryFilters: [] } }], { reportColors: {} })");
  assert.deepStrictEqual(json("ITEM.reportTotals"), { "readmission rate": 0.108 });
  assert.deepStrictEqual(json("totalsPlan(ITEM, INFO).kinds"), [null, "report"]);
});

test("a total from Power BI on the rows' scale: a measure returning 10.76 with a '%' sign is 10.8%, not 1076%", () => {
  const { json } = loadApp();
  const rows = [0.104, 0.104, 0.083, 0.125, 0.125, 0.104];
  assert.strictEqual(json(`onRowsScale(10.764, ${JSON.stringify(rows)}, "0.0%")`), 0.10764);
  // a measure returning a fraction: unchanged
  assert.strictEqual(json(`onRowsScale(0.108, ${JSON.stringify(rows)}, "0.0%")`), 0.108);
  // a share well outside the rows (-42% … 201%), a total 6.78%: unchanged
  assert.strictEqual(json(`onRowsScale(0.0678, [-0.4291, 2.0163, 0.0531], "0.00%")`), 0.0678);
  // not a percentage: unchanged
  assert.strictEqual(json(`onRowsScale(1076, [10, 12], "#,##0")`), 1076);
});

test("a condition filter for the query service in Power BI's own words ('RELATIVE MONTH ≥ -19 and is not blank')", () => {
  const { json } = loadApp();
  const filter = { target: { table: "Opportunity Calendar", column: "RELATIVE MONTH" }, logicalOperator: "And",
    conditions: [{ operator: "GreaterThanOrEqual", value: -19 }, { operator: "IsNotBlank" }] };
  const where = json(`semanticWhere(${JSON.stringify([filter])}, (entity, prop) => ({ Column: { Expression: { SourceRef: { Source: "o11" } }, Property: prop } }))`);
  // exactly as the Regional Sales report writes this filter in its own queries
  const col = { Column: { Expression: { SourceRef: { Source: "o11" } }, Property: "RELATIVE MONTH" } };
  assert.deepStrictEqual(where, [{ Condition: { And: {
    Left: { Comparison: { ComparisonKind: 2, Left: col, Right: { Literal: { Value: "-19L" } } } },
    Right: { Not: { Expression: { Comparison: { ComparisonKind: 0, Left: col, Right: { Literal: { Value: "null" } } } } } } } } }]);
  // an operator Power BI's query can't take: the query isn't made (as before)
  assert.throws(() => json(`semanticWhere(${JSON.stringify([{ target: filter.target, conditions: [{ operator: "Unknown", value: 1 }] }])}, () => ({}))`));
});

test("a renamed matrix column takes the report's units by its field ('Qualified Pipeline' in K), and a fixed unit stays on its total", () => {
  const { ctx, json } = loadApp();
  const f = (kind, entity, prop, label) => ({ ref: prop.toLowerCase(), label, queryRef: `${entity}.${prop}`, field: { kind, entity, prop } });
  ctx.INFO = { roles: { Values: [f("measure", "Opportunities", "Revenue Won", null), f("measure", "Opportunities", "Revenue In Pipeline", "qualified pipeline")] },
    columnUnits: { "m:revenue won": { units: 1000, precision: null }, "m:revenue in pipeline": { units: 1000, precision: null } } };
  ctx.ITEM = { type: "pivotTable", rows: [["Product", "Revenue Won", "Qualified Pipeline"], ["Black cover 7\"", "$1512090.00", "$2259504.00"], ["E-reader pen", "$1509290.00", "$2717228.00"]] };
  const fmts = json("reportFormats(ITEM, INFO)");
  assert.ok(/,"K"/.test(fmts[1]) && /,"K"/.test(fmts[2]), JSON.stringify(fmts));
  assert.deepStrictEqual(json("ITEM.fixedUnits"), [false, true, true]);
});

test("a combo split by a legend: its line is the category's total (Seps SPLY per month), on the columns' scale with no second axis", () => {
  const { ctx, run, json } = loadApp();
  const f = (kind, entity, prop) => ({ ref: prop.toLowerCase(), label: null, field: { kind, entity, prop } });
  ctx.INFO = json(`visualColorInfo(${JSON.stringify({ type: "lineStackedColumnComboChart", objects: {}, container: {}, dataPoint: [], bubbles: [], values: [],
    roles: { Category: [f("column", "Date", "Month")], Series: [f("column", "Separation Reasons", "SeparationReason")], Y: [f("measure", "Employee", "Seps")], Y2: [f("measure", "Employee", "Seps SPLY")] } })}, null)`);
  ctx.T = { type: "lineStackedColumnComboChart", title: "t", rows: [["Month", "SeparationReason", "Seps", "Seps SPLY"],
    ["Jan", "Involuntary", "374", "280"], ["Jan", "Voluntary", "690", "620"], ["Feb", "Involuntary", "278", "208"], ["Feb", "Voluntary", "710", "600"]] };
  run(`T.colorInfo = INFO; T.colors = columnColors(T, ["#118DFF", "#12239E"], INFO); T.themePalette = ["#118DFF"]; var cl = chartPlan(T);
    Object.assign(cl, { look: chartLook({ type: T.type }, null), sizePx: { w: 600, h: 300 }, title: "t", showTitle: true });
    cl.refs = { catRef: "D!A2:A3", series: cl.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B3" })), range: () => "D!B2" };`);
  assert.deepStrictEqual(json("cl.series.find((s) => s.line).values"), [900, 808], "Jan 280 + 620, Feb 208 + 600");
  const xml = run("chartXml(cl).xml");
  assert.ok(!/<c:axId val="50030"\/>/.test(xml), "no second axis");
  // a line value repeated on every legend row: read once, as before
  ctx.T.rows = [["Month", "SeparationReason", "Seps", "Seps SPLY"], ["Jan", "Involuntary", "374", "900"], ["Jan", "Voluntary", "690", "900"]];
  run("var c2 = chartPlan(T);");
  assert.deepStrictEqual(json("c2.series.find((s) => s.line).values"), [900]);
});

test("two fields shown under one name (Seps YoY and Actives YoY, both 'YoY'): their columns aren't matched on a guess", () => {
  const { json } = loadApp();
  const e = (ref, label, queryRef) => ({ ref, label, queryRef, field: { kind: "measure", entity: "Employee", prop: queryRef.split(".")[1] } });
  const entries = [e("seps", null, "Employee.Seps"), e("seps yoy % change", "yoy", "Employee.Seps YoY % Change"), e("actives", null, "Employee.Actives"),
    e("actives yoy % change", "yoy", "Employee.Actives YoY % Change")];
  for (const header of [["Seps", "YoY", "Actives", "YoY"], ["Seps", "YoY", "Actives", "YoY (2)"]]) {
    const m = json(`fieldsByColumn(${JSON.stringify(header)}, ${JSON.stringify(entries)}).map((x) => x && x.queryRef)`);
    assert.deepStrictEqual(m, ["Employee.Seps", null, "Employee.Actives", null], header.join(", "));
  }
  // one field under its own new name: matched, as before
  assert.deepStrictEqual(json(`fieldsByColumn(["Act SPLY"], [${JSON.stringify(e("actives sply", "act sply", "Employee.Actives SPLY"))}]).map((x) => x && x.queryRef)`), ["Employee.Actives SPLY"]);
});
