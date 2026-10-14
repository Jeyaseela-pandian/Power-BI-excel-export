/*
 * Dev-only regression suite for report colours (not loaded by the app). Builds report
 * files (PBIR format) covering every way Power BI colours data, runs the colour code on
 * matching exported data and compares with colours worked out by hand.
 * Open http://localhost:5500, load harness.js and this file, then in the console:
 *   const r = await __colourCases();   // { passed, failed, results }
 */
(() => {
  const THEME = {
    name: "Case theme",
    dataColors: ["#E41A1C", "#377EB8", "#4DAF4A", "#984EA3", "#FF7F00", "#A65628", "#F781BF", "#999999"],
    background: "#FFFFFF", foreground: "#252423", good: "#00AA00", bad: "#CC0000", maximum: "#0000FF", minimum: "#FFFFFF",
  };
  const P = THEME.dataColors;

  // ── Report-file building blocks ─────────────────────────────────────
  const fieldRef = (f) => {
    const [table, prop] = f.split(".");
    return /^[A-Z][a-z]+ /.test(prop) || /Sales|Cost|Margin|Price|Orders|Value|Year/.test(prop) && !/Category|Group|Region|Segment|Channel|Month/.test(prop)
      ? { Measure: { Expression: { SourceRef: { Entity: table } }, Property: prop } }
      : { Column: { Expression: { SourceRef: { Entity: table } }, Property: prop } };
  };
  const roles = (r) => Object.fromEntries(Object.entries(r).map(([role, fields]) => [role, {
    projections: fields.map((f) => (typeof f === "string" ? { field: fieldRef(f), queryRef: f, nativeQueryRef: f.split(".")[1] }
      : { field: fieldRef(f.ref), queryRef: f.ref, nativeQueryRef: f.ref.split(".")[1], displayName: f.label })),
  }]));
  const lit = (hex) => ({ Literal: { Value: `'${hex}'` } });
  const num = (n) => ({ Literal: { Value: `${n}D` } });
  const themeColor = (id, pct = 0) => ({ ThemeDataColor: { ColorId: id, Percent: pct } });
  const solid = (expr) => ({ solid: { color: { expr } } });
  const measure = (f) => fieldRef(f);
  const grad = (input, min, max, mid, opts = {}) => ({
    FillRule: { Input: measure(input), FillRule: { [mid ? "linearGradient3" : "linearGradient2"]: {
      min: { color: lit(min), ...(opts.minAt !== undefined ? { value: num(opts.minAt) } : {}) },
      ...(mid ? { mid: { color: lit(mid) } } : {}),
      max: { color: lit(max), ...(opts.maxAt !== undefined ? { value: num(opts.maxAt) } : {}) },
      nullColoringStrategy: { strategy: { Literal: { Value: `'${opts.nulls || "asZero"}'` } }, ...(opts.nullColor ? { color: lit(opts.nullColor) } : {}) },
    } } },
  });
  const cmp = (input, kind, right) => ({ Comparison: { ComparisonKind: kind, Left: measure(input), Right: right } });
  const cond = (cases) => ({ Conditional: { Cases: cases.map(([c, color]) => ({ Condition: c, Value: typeof color === "string" ? lit(color) : color })) } });
  const pct = (p) => ({ RangePercent: { Min: {}, Max: {}, Percent: p } });
  const all = { data: [{ dataViewWildcard: { matchingOption: 1 } }] };
  const meta = (f) => ({ metadata: f });
  const item = (col, v) => ({ data: [{ scopeId: { Comparison: { ComparisonKind: 0, Left: fieldRef(col), Right: lit(v) } } }] });
  const fill = (expr, selector) => ({ properties: { fill: solid(expr) }, ...(selector ? { selector } : {}) });

  const visual = (name, visualType, r, objects = {}) => ({ name, visual: { visualType, query: { queryState: roles(r) }, objects } });

  async function buildReport(visuals) {
    const zip = new JSZip();
    zip.file("Report/definition/report.json", JSON.stringify({ themeCollection: { baseTheme: { name: "CaseTheme", type: "SharedResources" } } }));
    zip.file("Report/StaticResources/SharedResources/BaseThemes/CaseTheme.json", JSON.stringify(THEME));
    for (const v of visuals) zip.file(`Report/definition/pages/page1/visuals/${v.name}/visual.json`, JSON.stringify(v));
    return readDesign(await JSZip.loadAsync(await zip.generateAsync({ type: "arraybuffer" })), null);
  }

  // ── Cases: report visual, exported rows, expected colours ───────────
  // expect: { col: header name, rows: [hex per row] } or { series: { header: hex } } or { note: true }
  const cases = [
    { id: "theme series colour", v: visual("c1", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] }),
      rows: [["Category", "Sales"], ["A", "1"], ["B", "2"]], expect: { series: { Sales: P[0] } } },
    { id: "measure colours, one set by hand", v: visual("c2", "lineChart", { Category: ["S.Month"], Y: ["S.Sales", "S.Cost"] }, { dataPoint: [fill(lit("#123456"), meta("S.Cost"))] }),
      rows: [["Month", "Sales", "Cost"], ["Jan", "1", "2"]], expect: { series: { Sales: P[0], Cost: "#123456" } } },
    { id: "legend order with a series missing from the first group, one set by hand",
      v: visual("c3", "clusteredColumnChart", { Category: ["S.Year"], Series: ["S.Segment"], Y: ["S.Sales"] }, { dataPoint: [fill(lit("#ABCDEF"), item("S.Segment", "C"))] }),
      rows: [["Year", "Segment", "Sales"], ["2023", "B", "1"], ["2023", "C", "2"], ["2024", "A", "3"], ["2024", "B", "4"], ["2024", "C", "5"]],
      expect: { col: "Sales", rows: [P[1], "#ABCDEF", P[0], P[1], "#ABCDEF"] } },
    { id: "pie slice set by hand", v: visual("c4", "pieChart", { Category: ["S.Category"], Y: ["S.Sales"] }, { dataPoint: [fill(lit("#00FFFF"), item("S.Category", "B"))] }),
      rows: [["Category", "Sales"], ["A", "1"], ["B", "2"], ["C", "3"]], expect: { col: "Sales", rows: [P[0], "#00FFFF", P[2]] } },
    { id: "theme colour shade for all points", v: visual("c5", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] }, { dataPoint: [fill(themeColor(3, 0.4))] }),
      rows: [["Category", "Sales"], ["A", "1"]], expect: { series: { Sales: mixHex(P[1], "#FFFFFF", 0.4) } } },
    { id: "white and black theme colours", v: visual("c6", "funnel", { Y: ["S.Sales", "S.Cost"] }, { dataPoint: [fill(themeColor(0, -0.2), meta("S.Sales")), fill(themeColor(1), meta("S.Cost"))] }),
      rows: [["Sales", "Cost"], ["10", "5"]], expect: { series: { Sales: "#CCCCCC", Cost: "#000000" } } },
    { id: "colour scale with set min / max values", v: visual("c7", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] }, { dataPoint: [fill(grad("S.Sales", "#000000", "#FFFFFF", null, { minAt: 0, maxAt: 100 }), all)] }),
      rows: [["Category", "Sales"], ["A", "0"], ["B", "50"], ["C", "100"]], expect: { col: "Sales", rows: ["#000000", "#808080", "#FFFFFF"] } },
    { id: "3-stop colour scale, middle by default", v: visual("c8", "clusteredColumnChart", { Category: ["S.Category"], Y: ["S.Sales"] }, { dataPoint: [fill(grad("S.Sales", "#FF0000", "#00FF00", "#FFFFFF"), all)] }),
      rows: [["Category", "Sales"], ["A", "0"], ["B", "50"], ["C", "100"]], expect: { col: "Sales", rows: ["#FF0000", "#FFFFFF", "#00FF00"] } },
    { id: "rules on % of range", v: visual("c9", "clusteredColumnChart", { Category: ["S.Category"], Y: ["S.Sales"] },
      { dataPoint: [fill(cond([[cmp("S.Sales", 2, pct(0.5)), "#00FF00"], [cmp("S.Sales", 4, pct(0.5)), "#FF0000"]]), all)] }),
      rows: [["Category", "Sales"], ["A", "10"], ["B", "40"], ["C", "90"]], expect: { col: "Sales", rows: ["#FF0000", "#FF0000", "#00FF00"] } },
    { id: "rule on a text value, others keep the series colour", v: visual("c10", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] },
      { dataPoint: [fill(cond([[{ Comparison: { ComparisonKind: 0, Left: fieldRef("S.Category"), Right: lit("B") } }, "#800080"]]), all)] }),
      rows: [["Category", "Sales"], ["A", "1"], ["B", "2"]], expect: { col: "Sales", rows: [P[0], "#800080"] } },
    { id: "rule on one measure, the other keeps its colour", v: visual("c11", "lineChart", { Category: ["S.Month"], Y: ["S.Sales", "S.Cost"] },
      { dataPoint: [fill(cond([[cmp("S.Cost", 1, num(5)), "#00FF00"]]), meta("S.Cost"))] }),
      rows: [["Month", "Sales", "Cost"], ["Jan", "1", "9"], ["Feb", "2", "3"]], expect: { series: { Sales: P[0] }, col: "Cost", rows: ["#00FF00", P[1]] } },
    { id: "theme keyword colours in rules", v: visual("c12", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] },
      { dataPoint: [fill(cond([[cmp("S.Sales", 2, num(0)), { Literal: { Value: "'good'" } }], [cmp("S.Sales", 4, num(0)), { Literal: { Value: "'bad'" } }]]), all)] }),
      rows: [["Category", "Sales"], ["A", "5"], ["B", "-5"]], expect: { col: "Sales", rows: [THEME.good, THEME.bad] } },
    { id: "hierarchy: coloured by the group total", v: visual("c13", "clusteredBarChart", { Category: ["S.Group", "S.Region"], Y: ["S.Sales"] }, { dataPoint: [fill(grad("S.Sales", "#000000", "#FFFFFF"), all)] }),
      rows: [["Group", "Region", "Sales"], ["G1", "R1", "10"], ["G1", "R2", "30"], ["G2", "R3", "20"]], expect: { col: "Sales", rows: ["#FFFFFF", "#FFFFFF", "#000000"] } },
    { id: "hierarchy with an average: row by row", v: visual("c14", "clusteredBarChart", { Category: ["S.Group", "S.Region"], Y: ["S.Avg Price"] }, { dataPoint: [fill(grad("S.Avg Price", "#000000", "#FFFFFF"), all)] }),
      rows: [["Group", "Region", "Avg Price"], ["G1", "R1", "10"], ["G1", "R2", "30"], ["G2", "R3", "20"]], expect: { col: "Avg Price", rows: ["#000000", "#FFFFFF", "#808080"] } },
    { id: "scale per legend value on a field not exported: legend order + note",
      v: visual("c15", "lineChart", { Category: ["S.Month"], Series: ["S.Year"], Y: ["S.Orders"] }, { dataPoint: [fill(grad("S.Latest year", "#000000", "#FFFFFF"), { data: [{ roles: ["Series"] }] })] }),
      rows: [["Month", "Year", "Orders"], ["Jan", "2018", "1"], ["Jan", "2019", "2"], ["Jan", "2020", "3"]], expect: { col: "Orders", rows: ["#000000", "#808080", "#FFFFFF"], note: true } },
    { id: "blanks: as zero, or in a set colour", v: visual("c16", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] },
      { dataPoint: [fill(grad("S.Sales", "#000000", "#FFFFFF", null, { minAt: 0, maxAt: 100, nulls: "specificColor", nullColor: "#FF00FF" }), all)] }),
      rows: [["Category", "Sales"], ["A", ""], ["B", "100"]], expect: { col: "Sales", rows: ["#FF00FF", "#FFFFFF"] } },
    { id: "colour from a field in the data", v: visual("c17", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"], Tooltips: ["S.Bar Color"] }, { dataPoint: [fill(measure("S.Bar Color"), all)] }),
      rows: [["Category", "Sales", "Bar Color"], ["A", "1", "#111111"], ["B", "2", "#222222"]], expect: { col: "Sales", rows: ["#111111", "#222222"] } },
    { id: "colour from a field not exported: series colour", v: visual("c18", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"] }, { dataPoint: [fill(measure("S.Hidden Color"), all)] }),
      rows: [["Category", "Sales"], ["A", "1"]], expect: { series: { Sales: P[0] } } },
    { id: "waterfall default colours", v: visual("c19", "waterfallChart", { Category: ["S.Month"], Y: ["S.Sales"] }),
      rows: [["Month", "Sales"], ["Jan", "5"], ["Feb", "-3"]], expect: { col: "Sales", rows: [THEME.good, THEME.bad] } },
    { id: "waterfall colours set by hand", v: visual("c20", "waterfallChart", { Category: ["S.Month"], Y: ["S.Sales"] }, { sentimentColors: [{ properties: { increaseFill: solid(lit("#0000AA")), decreaseFill: solid(lit("#AA0000")) } }] }),
      rows: [["Month", "Sales"], ["Jan", "5"], ["Feb", "-3"]], expect: { col: "Sales", rows: ["#0000AA", "#AA0000"] } },
    { id: "renamed measure in a visual", v: visual("c21", "lineChart", { Category: ["S.Month"], Y: [{ ref: "S.Total Sales", label: "Revenue" }] }, { dataPoint: [fill(lit("#0F0F0F"), meta("S.Total Sales"))] }),
      rows: [["Month", "Revenue"], ["Jan", "1"]], expect: { series: { Revenue: "#0F0F0F" } } },
    { id: "Azure Map: the bubble colour wins over a leftover data point rule",
      v: visual("c23", "azureMap", { Category: ["S.Country"], Size: ["S.Customers"] }, {
        bubbleLayer: [{ properties: { fillColor: solid(grad("S.Customers", "#000000", "#FFFFFF")) }, selector: all }],
        dataPoint: [fill(grad("S.Customers", "#FF0000", "#0000FF"), all)] }),
      rows: [["Country", "Customers"], ["A", "0"], ["B", "10"]], expect: { col: "Customers", rows: ["#000000", "#FFFFFF"] } },
    { id: "tooltip measures aren't series", v: visual("c22", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"], Tooltips: ["S.Cost"] }),
      rows: [["Category", "Sales", "Cost"], ["A", "1", "2"]], expect: { series: { Sales: P[0] }, uncoloured: ["Cost"] } },
  ];

  const hexOf = (argb) => (argb ? "#" + String(argb).slice(2).toUpperCase() : null);

  window.__colourCases = async () => {
    const design = await buildReport(cases.map((c) => c.v));
    const results = [];
    for (const c of cases) {
      const info = design.visuals[c.v.name];
      const item = { rows: c.rows, type: c.v.visual.visualType, name: c.v.name, displayTitle: c.id };
      const colors = columnColors(item, design.theme.dataColors, info);
      const header = c.rows[0];
      const problems = [];
      if (!colors) problems.push("no colours");
      else {
        for (const [h, hex] of Object.entries(c.expect.series || {})) {
          const got = hexOf(colors.series[header.indexOf(h)]);
          if (got !== hex.toUpperCase()) problems.push(`${h} series ${got} ≠ ${hex}`);
        }
        if (c.expect.col) {
          const ci = header.indexOf(c.expect.col);
          c.rows.slice(1).forEach((row, ri) => {
            const got = colors.perCol && colors.perCol[ci] ? hexOf(colors.perCol[ci][ri])
              : colors.dot ? hexOf(colors.dot.colorOf(String(row[colors.dot.keyCol]).trim())) : hexOf(colors.series[ci]);
            if (got !== c.expect.rows[ri].toUpperCase()) problems.push(`row ${ri + 1} ${got} ≠ ${c.expect.rows[ri]}`);
          });
        }
        if (c.expect.note && !colors.note) problems.push("no approximation note");
        for (const h of c.expect.uncoloured || []) if (colors.series[header.indexOf(h)]) problems.push(`${h} shouldn't be coloured`);
      }
      results.push({ id: c.id, ok: !problems.length, problems });
    }

    // Tables, matrices and cards: a full export, read back
    results.push(...await tableCases());
    return { passed: results.filter((r) => r.ok).length, failed: results.filter((r) => !r.ok).length, results: results.filter((r) => !r.ok).length ? results : results.map((r) => r.id) };
  };

  async function tableCases() {
    const visuals = [
      visual("t1", "tableEx", { Values: ["S.Category", "S.Sales", "S.Cost", "S.Margin"] }, {
        values: [
          { properties: { backColor: solid(grad("S.Sales", "#FFFFFF", "#0000FF")) }, selector: { ...all, metadata: "S.Sales" } },
          { properties: { fontColor: solid(cond([[cmp("S.Cost", 1, num(50)), "#FF0000"]])) }, selector: { ...all, metadata: "S.Cost" } },
          { properties: { icon: { kind: "Icon", value: { expr: cond([[cmp("S.Cost", 2, num(66)), "'ArrowUp'"], [cmp("S.Cost", 2, num(33)), "'ArrowRight'"], [cmp("S.Cost", 2, num(0)), "'ArrowDown'"]].map(([c, v]) => [c, { Literal: { Value: v } }])) } } }, selector: { metadata: "S.Cost" } },
        ],
        columnFormatting: [{ properties: { dataBars: { positiveColor: solid(lit("#00AA00")) } }, selector: { metadata: "S.Margin" } }],
        columnHeaders: [{ properties: { backColor: solid(lit("#222222")) } }],
      }),
      visual("t2", "pivotTable", { Rows: ["S.Category"], Values: [{ ref: "S.Total Sales", label: "Revenue" }] }, {
        values: [{ properties: { backColor: solid(lit("#FFFF00")) }, selector: { metadata: "S.Total Sales" } }],
      }),
      visual("k1", "cardVisual", { Data: ["S.Sales", "S.Cost"] }, {
        value: [{ properties: { fontColor: solid(lit("#123456")) }, selector: meta("S.Sales") }, { properties: { fontColor: solid(cond([[cmp("S.Cost", 1, num(0)), "#00AA00"]])) }, selector: meta("S.Cost") }],
      }),
    ];
    const zip = new JSZip();
    zip.file("Report/definition/report.json", JSON.stringify({ themeCollection: { baseTheme: { name: "CaseTheme" } } }));
    zip.file("Report/StaticResources/SharedResources/BaseThemes/CaseTheme.json", JSON.stringify(THEME));
    for (const v of visuals) zip.file(`Report/definition/pages/page1/visuals/${v.name}/visual.json`, JSON.stringify(v));
    const buf = await zip.generateAsync({ type: "arraybuffer" });

    const csv = (rows) => rows.map((r) => r.join(",")).join("\n");
    const data = {
      t1: [["Category", "Sales", "Cost", "Margin"], ["A", "0", "90", "10"], ["B", "100", "20", "20"]],
      t2: [["Category", "Revenue"], ["A", "5"], ["B", "6"]],
      k1: [["Sales", "Cost"], ["100", "7"]],
    };
    const page = { name: "page1", displayName: "Cases", visibility: 0, getFilters: async () => [], setActive: async () => {},
      getVisuals: async () => [
        { name: "t1", title: "Table case", type: "tableEx", layout: { x: 0, y: 200, width: 500, height: 200 }, exportData: async () => ({ data: csv(data.t1) }) },
        { name: "t2", title: "Matrix case", type: "pivotTable", layout: { x: 600, y: 200, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.t2) }) },
        { name: "k1", title: "Card", type: "cardVisual", layout: { x: 0, y: 0, width: 300, height: 100 }, exportData: async () => ({ data: csv(data.k1) }) },
      ] };
    embeddedReport = { getActivePage: async () => page, getPages: async () => [page], getFilters: async () => [], getTheme: async () => ({ themeJson: THEME }), on() {}, off() {} };
    currentReport = { id: "cases", embedUrl: "https://x/?groupId=g" };
    currentReportName = "Cases";
    reportColorCache.clear();
    downloadDesignFile = async () => buf;
    let href; const orig = HTMLAnchorElement.prototype.click; HTMLAnchorElement.prototype.click = function () { href = this.href; };
    try {
      await exportToExcel([page], { pages: "current", layout: "report", collapse: true, visibleRows: 8, includeFilters: true, includeNotes: true, matchColors: true, colorCells: true },
        { accent: "#0B7F56", tableStyle: "Executive", titleFont: "Calibri Light" }, () => {}, () => false);
    } finally { HTMLAnchorElement.prototype.click = orig; }
    const out = await (await fetch(href)).arrayBuffer();
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(out);
    const ws = wb.worksheets[0];
    const check = await __check({ wb, buf: out });
    const xml = (await __unzip(out))["xl/worksheets/sheet1.xml"];
    // A header cell of an exported table (not a KPI tile label with the same text)
    const colN = (x) => x.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    const find = (text) => {
      for (const x of Object.values(ws.tables)) {
        const m = x.table.tableRef.match(/([A-Z]+)(\d+):/);
        const k = x.table.columns.findIndex((c) => c.name === text);
        if (k >= 0) return { r: +m[2], c: colN(m[1]) + k };
      }
      return null;
    };
    const fillAt = (r, c) => hexOf(ws.getCell(r, c).fill && ws.getCell(r, c).fill.fgColor && ws.getCell(r, c).fill.fgColor.argb);
    const fontAt = (r, c) => hexOf(ws.getCell(r, c).font && ws.getCell(r, c).font.color && ws.getCell(r, c).font.color.argb);
    const res = [];
    const t = (id, ok, detail) => res.push({ id, ok, problems: ok ? [] : [detail] });
    const sales = find("Sales"), cost = find("Cost"), revenue = find("Revenue");
    t("table: file is valid", !check.errors.length, check.errors.join("; "));
    t("table: colour-scale background", sales && fillAt(sales.r + 1, sales.c) === "#FFFFFF" && fillAt(sales.r + 2, sales.c) === "#0000FF", `${fillAt(sales.r + 1, sales.c)} ${fillAt(sales.r + 2, sales.c)}`);
    t("table: readable text on dark fills", sales && fontAt(sales.r + 2, sales.c) === "#FFFFFF", fontAt(sales.r + 2, sales.c));
    t("table: rule font colour", cost && fontAt(cost.r + 1, cost.c) === "#FF0000" && fontAt(cost.r + 2, cost.c) !== "#FF0000", `${fontAt(cost.r + 1, cost.c)} ${fontAt(cost.r + 2, cost.c)}`);
    t("table: header colour", sales && fillAt(sales.r, sales.c) === "#222222", fillAt(sales.r, sales.c));
    t("table: data bars", /<dataBar>[\s\S]*?FF00AA00/.test(xml), "no green data bar");
    t("table: icon set with number thresholds", /iconSet iconSet="3Arrows"[^>]*>(<cfvo type="min"\/>)<cfvo type="num" val="33"\/><cfvo type="num" val="66"\/>/.test(xml), (xml.match(/<iconSet[\s\S]*?<\/iconSet>/) || ["none"])[0]);
    t("matrix: renamed value found by position", revenue && fillAt(revenue.r + 1, revenue.c) === "#FFFF00", revenue && fillAt(revenue.r + 1, revenue.c));
    let salesValue = null, costValue = null;
    ws.eachRow((row) => row.eachCell((cell) => {
      if (cell.value === 100 && cell.font && cell.font.size >= 20) salesValue = hexOf(cell.font.color.argb);
      if (cell.value === 7 && cell.font && cell.font.bold) costValue = hexOf(cell.font.color.argb);
    }));
    t("card: value colour set by hand", salesValue === "#123456", salesValue);
    t("card: value colour from a rule", costValue === "#00AA00", costValue);
    return res;
  }
})();
