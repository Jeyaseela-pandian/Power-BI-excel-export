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
