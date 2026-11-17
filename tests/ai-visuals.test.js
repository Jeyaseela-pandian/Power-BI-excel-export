/* AI and special visuals: key influencers as exact shares, decomposition trees as indented
   tables, small multiples as one grouped chart, and notes for visuals with nothing to export */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const col = (entity, prop) => ({ kind: "column", entity, prop });

/* ── Key influencers ── */
const KI_INFO = {
  keyDrivers: {
    target: { ref: "status", label: "status", field: col("Opportunities", "Status") },
    targetValue: "Won",
    explainBy: [
      { ref: "campaign name", label: "campaign name", field: col("Opportunities", "Campaign Name") },
      { ref: "value", label: null, field: { kind: "agg", entity: "Opportunities", prop: "Value", agg: 0 } },
    ],
    analysis: "keyInfluencers",
  },
};
// Customer Follow-up: 84 of 100 rows Won; the other campaigns 545 of 905
const COUNTS = [["Customer Follow-up", "Won", 84], ["Customer Follow-up", "Lost", 16], ["Twitter", "Won", 300], ["Twitter", "Lost", 300],
  ["Monthly Newsletter", "Won", 240], ["Monthly Newsletter", "Lost", 60], [null, "Won", 5]];

async function keyInfluencers(route) {
  const app = loadApp();
  const queries = [];
  app.ctx.executeDaxQuery = async (q) => {
    queries.push(q);
    if (route !== "semantic model") return { error: "querying the semantic model needs Build permission on it" };
    return { rows: COUNTS.map(([v, t, n]) => ({ "Opportunities[Campaign Name]": v, "Opportunities[Status]": t, "[n]": n })) };
  };
  app.ctx.websiteQuery = async () => { if (route === "query service") return COUNTS; throw new Error("Power BI's query service answered 403"); };
  app.ctx.INFO = KI_INFO;
  app.run(`visualInfo = () => INFO; KI = { type: "keyDriversVisual", title: "Key Influencers", name: "ki", page: "p", rows: [["Status", "Campaign Name"], ["Won", "x"]], layout: { x: 0, y: 0, width: 800, height: 400 }, visualFilters: [] };`);
  await app.run(`loadKeyInfluencers([{ data: { items: [KI], queryFilters: [] } }], { reportColors: {} })`);
  return { out: app.json(`(() => { const notes = []; const out = keyInfluencerItems(KI, INFO, notes); return { tables: out.map((i) => [i.type, i.title, i.rows]), notes }; })()`), queries };
}

test("key influencers: exact shares from counted rows, through the semantic model or Power BI's query service", async () => {
  for (const route of ["semantic model", "query service"]) {
    const { out } = await keyInfluencers(route);
    assert.strictEqual(out.tables.length, 1, route);
    const [type, title, rows] = out.tables[0];
    assert.strictEqual(type, "keyInfluencers");
    assert.strictEqual(title, "What influences Status to be Won?");
    assert.deepStrictEqual(rows[0], ["When", "Is", "Status is Won", "Status is Won (others)"]);
    assert.deepStrictEqual(rows.find((r) => r[1] === "Customer Follow-up"), ["Campaign Name", "Customer Follow-up", "84.0%", "60.2%"]);
    assert.ok(out.notes.some((n) => /Value .*number fields/.test(n)), "number fields are noted");
  }
});

test("key influencers: no table, only a note, when the rows can't be counted", async () => {
  const { out, queries } = await keyInfluencers("neither");
  assert.strictEqual(out.tables.length, 0);
  assert.ok(out.notes.some((n) => /couldn't be counted/.test(n)));
  assert.match(queries[0], /^EVALUATE SUMMARIZECOLUMNS\('Opportunities'\[Campaign Name\], 'Opportunities'\[Status\], "n", COUNTROWS\('Opportunities'\)\)$/);
});

/* ── Decomposition tree ── */
const lit = (v) => [{ Literal: { Value: `'${v}'` } }];
const TREE_INFO = {
  roles: {
    Analyze: [{ ref: "value", label: "sales opportunity", field: { kind: "agg", entity: "Opportunities", prop: "Value", agg: 0 } }],
    ExplainBy: [{ queryRef: "P.Category", ref: "product category", label: null }, { queryRef: "P.Product", ref: "product", label: null }, { queryRef: "A.Region", ref: "region", label: null }],
  },
  expansion: [{ roles: ["ExplainBy"], levels: [{ queryRefs: ["P.Category"], isPinned: true }, { queryRefs: ["P.Product"], isPinned: true }, { queryRefs: ["A.Region"], isPinned: true }],
    root: { children: [{ identityValues: lit("Software"), isToggled: true, children: [{ identityValues: lit("Design app"), isToggled: true }] }] } }],
  barsPerLevel: 8,
};
const TREE_ROWS = [["Product category", "Product", "Region", "Sales Opportunity"],
  ["Software", "Design app", "North America", "700"], ["Software", "Design app", "EMEA", "100"], ["Software", "Mobile app", "APAC", "200"],
  ["Computers", "Laptops", "EMEA", "500"], ["Furniture", "Desks", "LATAM", "300"]];

test("decomposition tree: the report's open branches as an indented table with exact totals", () => {
  const { ctx, json } = loadApp();
  ctx.INFO = TREE_INFO; ctx.ROWS = TREE_ROWS;
  const t = json(`decompositionTreeItem({ type: "decompositionTreeVisual", title: "decompositionTreeVisual (t)", name: "t", rows: ROWS }, INFO)`);
  assert.strictEqual(t.type, "decompositionTree");
  assert.strictEqual(t.title, "Sales Opportunity by Product category, Product and Region");
  assert.match(t.captionText, /Software › Design app/);
  const plain = t.rows.map((r) => [r[0].trim(), r[1], r[2], r[3]]);
  assert.deepStrictEqual(plain, [
    ["Level", "Name", "Sales Opportunity", "Share of parent"],
    ["Total", "", "1800", "100.0%"],
    ["Product category", "Software", "1000", "55.6%"],
    ["Product", "Design app", "800", "80.0%"],
    ["Region", "North America", "700", "87.5%"],
    ["Region", "EMEA", "100", "12.5%"],
    ["Product", "Mobile app", "200", "20.0%"],
    ["Product category", "Computers", "500", "27.8%"],
    ["Product category", "Furniture", "300", "16.7%"],
  ]);
});

test("decomposition tree: unchanged when its value isn't a sum or count, or without the layout", () => {
  const { ctx, run } = loadApp();
  ctx.INFO = { ...TREE_INFO, roles: { ...TREE_INFO.roles, Analyze: [{ ref: "value", label: "sales opportunity", field: { kind: "measure", entity: "O", prop: "Avg" } }] } };
  ctx.ROWS = TREE_ROWS;
  assert.strictEqual(run(`decompositionTreeItem({ type: "decompositionTreeVisual", rows: ROWS }, INFO).type`), "decompositionTreeVisual");
  assert.strictEqual(run(`decompositionTreeItem({ type: "decompositionTreeVisual", rows: ROWS }, null).type`), "decompositionTreeVisual");
  assert.strictEqual(run(`decompositionTreeItem({ type: "barChart", rows: ROWS }, INFO).type`), "barChart");
});

/* ── Small multiples ── */
const SM_INFO = {
  rule: "measure", category: { ref: "manager", label: "manager" }, categories: [{ ref: "manager", label: "manager" }],
  measures: [{ ref: "revenue won", label: "revenue won", role: "Y" }], overrides: {}, rules: [], tooltips: [],
  roles: { Rows: [{ ref: "product category", label: "product category" }], Category: [{ ref: "manager", label: "manager" }], Y: [{ ref: "revenue won", label: "revenue won" }] },
};
const SM_ROWS = [["Product category", "Manager", "Revenue Won"],
  ["Computers", "Low, Spencer", "$1200000"], ["Computers", "Weiler, Anne", "$800000"], ["Electronics", "Low, Spencer", "$300000"], ["Electronics", "Weiler, Anne", "$500000"],
  ["Furniture", "Low, Spencer", "$600000"], ["Furniture", "Weiler, Anne", "$900000"], ["Software", "Low, Spencer", "$2400000"], ["Software", "Weiler, Anne", "$3100000"]];

test("small multiples: one chart per value in a 2 x 2 grid, one shared axis, joined into one object", () => {
  const { ctx, run, json } = loadApp();
  ctx.INFO = SM_INFO; ctx.ROWS = SM_ROWS; ctx.P = ["#118DFF"];
  run(`[SM] = smallMultipleItems({ type: "clusteredBarChart", title: "clusteredBarChart (sm)", name: "sm", page: "p", rows: ROWS, layout: { x: 0, y: 0, width: 304, height: 312 } }, INFO);
    SM.colors = columnColors(SM, P, INFO); SM.colorInfo = INFO; SM.themePalette = P; SM.chart = chartPlan(SM);
    LK = chartLook({ type: SM.type }, null); S = valueScale([SM.sharedMax], 200, 75, { zero: true });
    LK = { ...LK, val: { ...LK.val, start: 0, end: Math.ceil(SM.sharedMax / S.step) * S.step } };
    Object.assign(SM.chart, { look: LK, title: "Revenue Won by Manager and Product category", sizePx: chartSizePx(SM), showTitle: true });
    G = panelPlans(SM, INFO, LK, P);
    G.plans.forEach((p, i) => { p.refs = { catRef: "'D'!$B$3:$B$4", series: p.series.map(() => ({ nameRef: "'D'!$C$2", valRef: "'D'!$C$3:$C$4" })), range: () => "'D'!$C$3" }; });
    SM.chart.panels = G;`);
  assert.deepStrictEqual(json("[G.cols, G.rows, G.plans.length]"), [2, 2, 4]);
  assert.deepStrictEqual(json("G.plans.map((p) => p.title)"), ["Computers", "Electronics", "Furniture", "Software"]);
  assert.deepStrictEqual(json("G.plans[3].series[0].values"), [2400000, 3100000]);
  // every panel on the same axis, rounded (3.1M → 4M)
  assert.deepStrictEqual(json(`G.plans.map((p) => (chartXml(p).xml.match(/<c:max val="([^"]+)"/) || [])[1])`), ["4000000", "4000000", "4000000", "4000000"]);
  const group = run(`panelGroupXml({ plan: SM.chart, col: 3, row: 12, wPx: SM.chart.sizePx.w, hPx: SM.chart.sizePx.h }, ["rId1", "rId2", "rId3", "rId4"], 1004)`);
  assert.strictEqual((group.match(/<xdr:grpSp>/g) || []).length, 1, "one group");
  assert.strictEqual((group.match(/<xdr:graphicFrame /g) || []).length, 4, "four charts inside it");
  assert.match(group, /<a:t>Revenue Won by Manager and Product category<\/a:t>/);
});

test("small multiples: a chart without them, with one value, or without the layout is unchanged", () => {
  const { ctx, json } = loadApp();
  ctx.INFO = SM_INFO; ctx.ROWS = SM_ROWS;
  assert.deepStrictEqual(json(`smallMultipleItems({ type: "clusteredBarChart", rows: [["a", "b"], ["x", "1"]] }, { roles: { Category: [{ ref: "a" }] } }).map((i) => !!i.panels)`), [false]);
  assert.deepStrictEqual(json(`smallMultipleItems({ type: "clusteredBarChart", rows: ROWS.slice(0, 3) }, INFO).map((i) => !!i.panels)`), [false]);
  assert.deepStrictEqual(json(`smallMultipleItems({ type: "clusteredBarChart", rows: ROWS }, null).map((i) => !!i.panels)`), [false]);
  assert.deepStrictEqual(json(`smallMultipleItems({ type: "pivotTable", rows: ROWS }, INFO).map((i) => !!i.panels)`), [false]);
});

/* ── Visuals with nothing to export ── */
test("smart narratives and Q&A are listed in the notes; other visuals are exported", async () => {
  const app = loadApp();
  const vis = (type, title, name, y) => ({ type, title, name, layout: { x: 0, y, width: 300, height: 200 }, isHidden: false,
    exportData: async () => ({ data: "Status,Value\nWon,10\nLost,5" }), getFilters: async () => [], getSlicerState: async () => ({ filters: [] }),
    getProperty: async () => { throw new Error("n/a"); } });
  app.ctx.PAGE = { name: "p1", getFilters: async () => [],
    getVisuals: async () => [vis("aiNarratives", "Revenue Summary", "sn", 2), vis("qnaVisual", "", "qa", 3), vis("clusteredBarChart", "Revenue Won by Product", "bar", 4)] };
  const r = await app.run(`embeddedReport = { getFilters: async () => [] }; collectPage(PAGE, { includeFilters: true }, () => {}, () => false)`);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r.items.map((i) => i.title))), ["Revenue Won by Product"]);
  assert.ok(r.notes.some((n) => /Revenue Summary.*smart narrative/.test(n)));
  assert.ok(r.notes.some((n) => /Q&A/.test(n)));
});

/* The visual's sort, from each place it's kept: the classic query, the newer format, a captured view */
test("a visual's sort is read from the report (classic and newer formats) and from the session's view", () => {
  const { json } = loadApp();
  const measure = { Measure: { Expression: { SourceRef: { Source: "o" } }, Property: "Revenue Won" } };
  assert.deepStrictEqual(json(`sortOf(${JSON.stringify([{ Direction: 2, Expression: measure }])}, { o: "Opportunities" })`),
    [{ field: { kind: "measure", entity: "Opportunities", prop: "Revenue Won" }, desc: true }]);
  assert.deepStrictEqual(json(`sortOf(${JSON.stringify([{ field: { Column: { Expression: { SourceRef: { Entity: "Owners" } }, Property: "Manager" } }, direction: "Ascending" }])})`),
    [{ field: { kind: "column", entity: "Owners", prop: "Manager" }, desc: false }]);
  assert.strictEqual(json("sortOf(undefined)"), null);
});

/* Small multiples built from SM_ROWS (exported largest first, as Power BI exports them) with a sort */
function smallMultiples(sort, live) {
  const { ctx, run, json } = loadApp();
  const field = (kind, entity, prop) => ({ kind, entity, prop });
  ctx.INFO = { ...SM_INFO, sort,
    roles: { Rows: [{ ref: "product category", label: "product category", field: field("column", "Products", "Product category") }],
      Category: [{ ref: "manager", label: "manager", field: field("column", "Owners", "Manager") }],
      Y: [{ ref: "revenue won", label: "revenue won", field: field("measure", "Opportunities", "Revenue Won") }] } };
  // exported by revenue, largest first: Software, Computers, Furniture, Electronics
  ctx.ROWS = [SM_ROWS[0], SM_ROWS[8], SM_ROWS[7], SM_ROWS[1], SM_ROWS[2], SM_ROWS[6], SM_ROWS[5], SM_ROWS[4], SM_ROWS[3]];
  ctx.LIVE = live || null;
  run(`[SM] = smallMultipleItems({ type: "clusteredBarChart", title: "t", name: "sm", page: "p", rows: ROWS, liveDrill: LIVE ? { orderBy: LIVE } : null, layout: { x: 0, y: 0, width: 304, height: 312 } }, INFO);`);
  return json("SM.panels.map((p) => p.value + ': ' + p.rows.slice(1).map((r) => r[0]).join(' / '))");
}
const BY_REVENUE_DESC = [{ field: { kind: "measure", entity: "Opportunities", prop: "Revenue Won" }, desc: true }];

test("small multiples: panels A to Z, and the bars in one order in every panel, from the visual's sort", () => {
  // saved sort: Revenue Won, high to low, by each manager's total (Weiler 5.3M, Low 4.5M)
  assert.deepStrictEqual(smallMultiples(BY_REVENUE_DESC), [
    "Computers: Weiler, Anne / Low, Spencer", "Electronics: Weiler, Anne / Low, Spencer",
    "Furniture: Weiler, Anne / Low, Spencer", "Software: Weiler, Anne / Low, Spencer"]);
  // the session's own sort wins: Manager, A to Z
  const byManager = [{ Direction: 1, Expression: { Column: { Expression: { SourceRef: { Entity: "Owners" } }, Property: "Manager" } } }];
  assert.deepStrictEqual(smallMultiples(BY_REVENUE_DESC, byManager), [
    "Computers: Low, Spencer / Weiler, Anne", "Electronics: Low, Spencer / Weiler, Anne",
    "Furniture: Low, Spencer / Weiler, Anne", "Software: Low, Spencer / Weiler, Anne"]);
  // sorted by the panels' own field, Z to A: panels Z to A
  assert.deepStrictEqual(smallMultiples([{ field: { kind: "column", entity: "Products", prop: "Product category" }, desc: true }]).map((p) => p.split(":")[0]),
    ["Software", "Furniture", "Electronics", "Computers"]);
  // no sort known: panels A to Z, bars as exported
  assert.deepStrictEqual(smallMultiples(null), [
    "Computers: Low, Spencer / Weiler, Anne", "Electronics: Weiler, Anne / Low, Spencer",
    "Furniture: Weiler, Anne / Low, Spencer", "Software: Weiler, Anne / Low, Spencer"]);
});

test("small multiples: axes only on the outer panels, the same plot in every panel, names on one line", () => {
  const { ctx, run, json } = loadApp();
  ctx.INFO = SM_INFO; ctx.ROWS = SM_ROWS; ctx.P = ["#118DFF"];
  run(`[SM] = smallMultipleItems({ type: "clusteredBarChart", title: "t", name: "sm", page: "p", rows: ROWS, layout: { x: 0, y: 0, width: 304, height: 312 } }, INFO);
    SM.colors = columnColors(SM, P, INFO); SM.colorInfo = INFO; SM.themePalette = P; SM.chart = chartPlan(SM);
    Object.assign(SM.chart, { look: chartLook({ type: SM.type }, null), title: "t", sizePx: chartSizePx(SM), showTitle: true });
    G = panelPlans(SM, INFO, SM.chart.look, P);
    G.plans.forEach((p) => { p.refs = { catRef: "'D'!$B$3:$B$4", series: p.series.map(() => ({ nameRef: "'D'!$C$2", valRef: "'D'!$C$3:$C$4" })), range: () => "'D'!$C$3" }; });
    X = G.plans.map((p) => chartXml(p).xml);`);
  const X = json("X");
  const axes = X.map((x) => {
    const cat = x.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)[0], val = x.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)[0];
    const noNames = /<c:delete val="1"\/>/.test(cat);
    return [noNames ? "no names" : "names", /<c:tickLblPos val="none"\/>/.test(val) ? "no numbers" : "numbers", /<c:majorGridlines>/.test(val), noNames || /wrap="none"/.test(cat)];
  });
  // Computers (top left), Electronics (top right), Furniture (bottom left), Software (bottom right)
  assert.deepStrictEqual(axes, [["names", "no numbers", true, true], ["no names", "no numbers", true, true], ["names", "numbers", true, true], ["no names", "numbers", true, true]]);
  // one plot size for all, one place per column
  const boxes = X.map((x) => {
    const m = x.match(/<c:plotArea><c:layout><c:manualLayout>[\s\S]*?<\/c:manualLayout>/)[0];
    return ["x", "y", "w", "h"].map((k) => Number((m.match(new RegExp(`<c:${k} val="([^"]+)"`)) || [])[1]));
  });
  assert.ok(boxes.every((b) => b[2] === boxes[0][2] && b[3] === boxes[0][3]), JSON.stringify(boxes));
  assert.ok(boxes[0][0] === boxes[2][0] && boxes[1][0] === boxes[3][0] && boxes[0][0] > boxes[1][0], JSON.stringify(boxes));
  // an ordinary bar chart: as before (no fixed layout, names wrap as Excel decides)
  ctx.T = { type: "clusteredBarChart", title: "t", rows: SM_ROWS.map((r) => r.slice(1)) };
  run(`T.colors = columnColors(T, P, null); T.themePalette = P; var one = chartPlan(T); Object.assign(one, { look: chartLook({ type: T.type }, null), sizePx: { w: 400, h: 300 }, title: "t", showTitle: true });
    one.refs = { catRef: "D!A1:A2", series: one.series.map(() => ({ nameRef: "D!B1", valRef: "D!B2:B3" })), range: () => "D!B2" }; ONE = chartXml(one).xml;`);
  const one = json("ONE");
  assert.ok(one.includes("<c:plotArea><c:layout/>") && !/wrap="none"/.test(one.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)[0]));
});
