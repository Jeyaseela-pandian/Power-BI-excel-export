/* Colours: hand-set and rule colours, legend colour order, the theme the export falls back on,
   and the chart background option */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { ctx, run, json } = loadApp();

test("funnels and pies keep their stages and hand-set colours; a two-row measure funnel stays a table", () => {
  ctx.P = ["#118DFF", "#12239E", "#E66C37"];
  ctx.T = { type: "funnel", title: "F", rows: [["Stage", "Rev"], ["Qualify", "100"], ["Close", "40"]] };
  run("T.colors = columnColors(T, P, null); T.themePalette = P; var f = chartPlan(T);");
  assert.deepStrictEqual(json("[f.categories, f.series[1].values]"), [["Qualify", "Close"], [100, 40]]);
  ctx.T = { type: "pieChart", title: "P", rows: [["Cat", "V"], ["A", "5"], ["B", "3"]] };
  run(`T.colorInfo = { rule: "item", category: { ref: "cat", label: "cat" }, measures: [{ ref: "v", label: "v", role: "Y" }],
    overrides: { "i:b": "#00FF00" }, roles: {}, rules: [], tooltips: [] }; T.colors = columnColors(T, P, T.colorInfo); T.themePalette = P;`);
  assert.deepStrictEqual(json("chartPlan(T).pointColors"), ["FF118DFF", "FF00FF00"]);
  ctx.T = { type: "funnel", title: "F", rows: [["A", "B"], ["1", "2"], ["3", "4"]] };
  run("T.colors = null; T.themePalette = P;");
  assert.strictEqual(run("chartPlan(T)"), null);
});

test("a line chart's series colours follow its colour rule, or the colours set by hand", () => {
  run(`P4 = ["#4DA6FF", "#45D5F5", "#7F8CFF", "#C77DFF"];
    rows4 = [["Mth", "Year", "Orders"]]; for (const [y, b] of [["2017", 300], ["2018", 350], ["2019", 600], ["2020", 2000]]) for (const m of ["Jan", "Feb"]) rows4.push([m, y, String(b)]);
    info4 = { rule: "legend", series: { ref: "year", label: "year" }, category: { ref: "mth", label: "mth" }, categories: [{ ref: "mth", label: "mth" }],
      measures: [{ ref: "orders", label: "orders", role: "Y" }], overrides: {}, tooltips: [], roles: {},
      rules: [{ scope: "series", rule: { kind: "gradient", input: "orders", min: { color: "#555555", at: null }, max: { color: "#4DA6FF", at: null }, nulls: null } }] };
    L4 = { type: "lineChart", title: "Orders", rows: rows4, colorInfo: info4, themePalette: P4 }; L4.colors = columnColors(L4, P4, info4);
    I5 = { ...info4, rules: [], overrides: { "i:2019": "#AACCFF" } }; L5 = { type: "lineChart", title: "O", rows: rows4, colorInfo: I5, themePalette: P4 }; L5.colors = columnColors(L5, P4, I5);`);
  assert.deepStrictEqual(json("chartPlan(L4).series.map((s) => [s.name, s.color])"),
    [["2017", "FF555555"], ["2018", "FF55575A"], ["2019", "FF546373"], ["2020", "FF4DA6FF"]]);
  assert.deepStrictEqual(json("chartPlan(L5).series.map((s) => [s.name, s.color])"),
    [["2017", "FF4DA6FF"], ["2018", "FF45D5F5"], ["2019", "FFAACCFF"], ["2020", "FFC77DFF"]]);
});

test("legend colours follow Power BI's order, not the exported row order", () => {
  run(`P8 = ["#118DFF", "#12239E", "#E66C37", "#6B007B"];
    i8 = { rule: "legend", series: { ref: "category", label: "category" }, category: { ref: "country", label: "country" }, categories: [{ ref: "country", label: "country" }],
      measures: [{ ref: "sales", label: "sales", role: "Y" }], overrides: {}, rules: [], tooltips: [], roles: {} };
    S8 = { type: "columnChart", title: "S", themePalette: P8, colorInfo: i8, rows: [["Country", "Category", "Sales"],
      ["United States", "Bikes", "15600000"], ["United States", "Components", "2700000"], ["United States", "Clothing", "500000"], ["United States", "Accessories", "300000"],
      ["Canada", "Bikes", "4000000"], ["Canada", "Components", "900000"], ["Canada", "Clothing", "200000"], ["Canada", "Accessories", "100000"]] };
    S8.colors = columnColors(S8, P8, i8);
    i9 = { ...i8, series: { ref: "priority", label: "priority" }, category: { ref: "region", label: "region" }, categories: [{ ref: "region", label: "region" }], measures: [{ ref: "n", label: "n", role: "Y" }] };
    S9 = { type: "columnChart", title: "C", themePalette: P8, colorInfo: i9, rows: [["Region", "Priority", "N"], ["East", "Low", "5"], ["East", "Medium", "9"], ["East", "High", "2"], ["West", "Low", "7"], ["West", "Medium", "1"], ["West", "High", "4"]] };
    S9.colors = columnColors(S9, P8, i9);`);
  assert.deepStrictEqual(json("chartPlan(S8).series.map((s) => [s.name, s.color])"),
    [["Accessories", "FF118DFF"], ["Bikes", "FF12239E"], ["Clothing", "FFE66C37"], ["Components", "FF6B007B"]]);
  assert.deepStrictEqual(json("chartPlan(S9).series.map((s) => s.name)"), ["Low", "Medium", "High"]);
  assert.deepStrictEqual(json(`legendOrder([["A", "Mar", "9"], ["A", "Jan", "5"], ["A", "Feb", "1"], ["B", "Jan", "8"], ["B", "Mar", "3"], ["B", "Feb", "2"]], 1, [0], 2)`), ["Jan", "Feb", "Mar"]);
});

test("the theme the export falls back on: Power BI's current colours, classic only without a theme", () => {
  const colors = (code) => json(`(${code}).dataColors.slice(0, 4)`);
  assert.deepStrictEqual(colors(`namedThemeFallback({ baseTheme: { name: "CY24SU10" } }, null)`), ["#118DFF", "#12239E", "#E66C37", "#6B007B"]);
  assert.deepStrictEqual(colors(`namedThemeFallback({ baseTheme: { name: "CY24SU10" } }, { dataColors: ["#4DA6FF", "#45D5F5", "#7F8CFF", "#C77DFF"] })`), ["#4DA6FF", "#45D5F5", "#7F8CFF", "#C77DFF"]);
  assert.deepStrictEqual(colors(`namedThemeFallback(null, null)`), ["#01B8AA", "#374649", "#FD625E", "#F2C80F"]);
  assert.deepStrictEqual(colors(`namedThemeFallback({}, null)`), ["#01B8AA", "#374649", "#FD625E", "#F2C80F"]);
});

test("treemaps and box & whisker charts take the report's background; white stays as before", () => {
  const frameOf = (type, rows, dark) => {
    ctx.T = { type, title: "t", rows };
    ctx.P = ["#118DFF", "#12239E"];
    run(`T.colors = columnColors(T, P, null); T.themePalette = P; var cx = chartPlan(T); var cl = chartLook({ type: T.type }, null);
      ${dark ? 'cl.frame.background = "#252423";' : "cl = lookOnWhite(cl);"}
      Object.assign(cx, { look: cl, sizePx: { w: 400, h: 300 }, title: "t", showTitle: true });
      cx.refs = { catRef: "D!A1:A3", series: cx.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B4" })), range: () => "D!B2", levels: [] };`);
    const xml = run("cx.spec.kind") === "boxWhisker" ? run("boxWhiskerXml(cx)") : run("treemapXml(cx, 1)");
    return ((xml.match(/<\/cx:chart>(?:<cx:spPr><a:solidFill><a:srgbClr val="(\w+)")?/) || [])[1]) || "none";
  };
  const treemap = [["Cat", "V"], ["A", "5"], ["B", "3"]];
  const box = [["Group", "Sample", "Value"], ["A", "1", "5"], ["A", "2", "7"], ["A", "3", "9"], ["B", "1", "2"], ["B", "2", "4"], ["B", "3", "8"]];
  assert.strictEqual(frameOf("treemap", treemap, true), "252423");
  assert.strictEqual(frameOf("treemap", treemap, false), "none", "white: Excel's own frame, as before");
  assert.strictEqual(frameOf("BoxWhiskerChart1455240051538", box, true), "252423");
  assert.strictEqual(frameOf("BoxWhiskerChart1455240051538", box, false), "FFFFFF");
});

test("chart background: the report's dark background, or white with readable text", () => {
  run(`LF = chartLook({ type: "funnel" }, null); LF.frame.background = "#252423"; LF.title.color = "#FFFFFF"; LF.cat.text.color = "#C8C6C4";
    FP = { funnel: true, title: "Sales, Cost and Profit", categories: ["Sales", "Cost", "Profit"],
      series: [{ name: "Offset", values: [0, 6, 48], fmt: "$#,##0" }, { name: "Value", values: [110, 97, 13], fmt: "$#,##0", color: null }],
      pointColors: ["FFCCCCCC", "FF808080", "FF7F8CFF"], refs: { catRef: "S!A1:A3", series: [{}, { nameRef: "S!B1", valRef: "S!B2:B4" }] } };`);
  const bg = (x) => (x.match(/<\/cx:chart><cx:spPr><a:solidFill><a:srgbClr val="(\w+)"/) || [])[1];
  const names = (x) => (x.match(/<cx:axis id="0">.*?srgbClr val="(\w+)"/) || [])[1];
  const asReport = run("funnelChartExXml({ ...FP, look: LF })");
  const white = run("funnelChartExXml({ ...FP, look: lookOnWhite(LF) })");
  assert.deepStrictEqual([bg(asReport), names(asReport)], ["252423", "C8C6C4"]);
  assert.deepStrictEqual([bg(white), names(white)], ["FFFFFF", "5B6B7F"]);
});
