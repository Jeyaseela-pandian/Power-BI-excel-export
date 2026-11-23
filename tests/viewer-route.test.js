/* The Viewer route switch (config.js viewerRoute): on (the default), Viewers' exports read the
   report's design and Power BI's own values from the services Power BI's website uses; off,
   nothing is asked of those services and the export falls back exactly as when they fail */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const REPORT = { id: "r1", embedUrl: "https://app.powerbi.com/reportEmbed?reportId=r1&groupId=g1" };

/* The app as a Viewer sees it: the report file can't be downloaded (403), and every website
   service call is recorded (and fails, so nothing further depends on its answer) */
function viewerApp(viewerRoute) {
  const app = loadApp();
  const calls = [];
  app.ctx.AbortController = AbortController;
  app.ctx.JSZip = { loadAsync: async () => ({}) };   // never reached: the download fails first
  app.ctx.getToken = async () => "token";
  app.ctx.getDesignToken = async () => "token";
  app.ctx.fetch = async (url) => {
    const u = String(url);
    if (/\/reports\/r1\/Export/.test(u)) return { ok: false, status: 403, redirected: false };
    calls.push(u);
    return { ok: false, status: 500, json: async () => ({}) };
  };
  if (viewerRoute !== undefined) app.ctx.window.APP_CONFIG.viewerRoute = viewerRoute;
  app.run(`currentReport = ${JSON.stringify(REPORT)}`);
  return { app, calls };
}

async function viewerExport(viewerRoute) {
  const { app, calls } = viewerApp(viewerRoute);
  const code = await app.ctx.loadReportColors(REPORT, null).then(() => "read", (e) => e.message);
  return { code, calls };
}

test("Viewer route: on by default, as before the switch", async () => {
  const { app } = viewerApp(undefined);
  assert.strictEqual(app.ctx.websiteRouteAllowed(), true);
  const { code, calls } = await viewerExport(undefined);
  assert.ok(calls.length > 0, "the website's services are asked");
  assert.strictEqual(code, "NO_ACCESS", "the export's note is the report file's, as before");
});

test("Viewer route: viewerRoute true behaves as the default", async () => {
  const on = await viewerExport(true);
  const unset = await viewerExport(undefined);
  assert.deepStrictEqual(on, unset);
});

test("Viewer route off: no website service is called, and the export falls back with the same note", async () => {
  const { code, calls } = await viewerExport(false);
  assert.deepStrictEqual(calls, [], "nothing is asked of the website's services");
  assert.strictEqual(code, "NO_ACCESS", "the same fallback and note as when the website layout can't be read");
});

test("Viewer route off: the query service isn't used either (totals, tree values, key influencers)", async () => {
  const { app, calls } = viewerApp(false);
  await assert.rejects(app.ctx.websiteQuery([{ kind: "column", entity: "Sales", prop: "Region" }], []), /WEBSITE_OFF/);
  await assert.rejects(app.ctx.postAsUser("https://example/explore/querydata", {}), /WEBSITE_OFF/);
  await assert.rejects(app.ctx.fetchAsUser("https://example/explore/reports/r1/modelsAndExploration"), /WEBSITE_OFF/);
  assert.deepStrictEqual(calls, []);
});
