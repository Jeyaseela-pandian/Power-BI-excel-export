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
  const DONUT = "doubleDonut2B2DE03D219E04DFF80684A45ECA4D25E";   // NeenOpal's Double Donut Chart (AppSource)

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
      rows: [["Category", "Sales"], ["A", "1"]], expect: { series: { Sales: P[0] }, note: true } },
    { id: "map: the rule on a tooltip measure colours that measure",
      v: visual("c26", "azureMap", { Category: ["S.Country"], Tooltips: ["S.Revenue"] }, { dataPoint: [fill(grad("S.Revenue", "#000000", "#FFFFFF"), all)] }),
      rows: [["Country", "Revenue"], ["A", "0"], ["B", "10"]], expect: { col: "Revenue", rows: ["#000000", "#FFFFFF"] } },
    { id: "rule per legend value reads each legend value's total",
      v: visual("c27", "stackedAreaChart", { Category: ["S.Month"], Series: ["S.Area"], Y: ["S.Actual"] }, { dataPoint: [fill(grad("S.Actual", "#000000", "#FFFFFF"), { data: [{ roles: ["Series"] }] })] }),
      rows: [["Month", "Area", "Actual"], ["Jan", "A", "1"], ["Feb", "A", "1"], ["Jan", "B", "5"], ["Feb", "B", "5"]], expect: { col: "Actual", rows: ["#000000", "#000000", "#FFFFFF", "#FFFFFF"] } },
    { id: "rule on a measure named alike apart from little words (+ note)",
      v: visual("c28", "clusteredBarChart", { Category: ["S.Region"], Y: ["S.Var Plan %"] }, { dataPoint: [fill(cond([[cmp("S.Var to Plan %", 1, num(0)), "#00FF00"]]), all)] }),
      rows: [["Region", "Var Plan %"], ["A", "5%"], ["B", "-3%"]], expect: { col: "Var Plan %", rows: ["#00FF00", P[0]], note: true } },
    { id: "a line's colour under Lines wins over its data colour",
      v: visual("c29", "lineChart", { Category: ["S.Month"], Y: ["S.Sales"] }, {
        dataPoint: [fill(lit("#111111"), meta("S.Sales"))], lineStyles: [{ properties: { strokeColor: solid(lit("#222222")) }, selector: meta("S.Sales") }] }),
      rows: [["Month", "Sales"], ["Jan", "1"]], expect: { series: { Sales: "#222222" } } },
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
    { id: "combo chart: line values aren't split by the legend",
      v: visual("c24", "lineStackedColumnComboChart", { Category: ["S.Month"], Y: ["S.Seps"], Y2: ["S.Seps SPLY"], Series: ["S.Reason"] },
        { dataPoint: [fill(lit("#AA0000"), item("S.Reason", "Voluntary")), fill(themeColor(1, 0.1), meta("S.Seps SPLY"))] }),
      rows: [["Month", "Seps", "Seps SPLY", "Reason"], ["Jan", "1", "2", "Involuntary"], ["Jan", "3", "4", "Voluntary"]],
      expect: { col: "Seps", rows: [P[0], "#AA0000"], series: { "Seps SPLY": "#1A1A1A" } } },
    { id: "legend chart ignores a leftover rule for every data point",
      v: visual("c25", "clusteredColumnChart", { Category: ["S.Month"], Series: ["S.Gender"], Y: ["S.Sales"] },
        { dataPoint: [fill(cond([[cmp("S.Sales", 1, num(0)), "#00FF00"]]), all), fill(lit("#F5C869"), item("S.Gender", "Female"))] }),
      rows: [["Month", "Gender", "Sales"], ["Jan", "Male", "5"], ["Jan", "Female", "6"]], expect: { col: "Sales", rows: [P[0], "#F5C869"] } },
    { id: "tooltip measures aren't series", v: visual("c22", "clusteredBarChart", { Category: ["S.Category"], Y: ["S.Sales"], Tooltips: ["S.Cost"] }),
      rows: [["Category", "Sales", "Cost"], ["A", "1", "2"]], expect: { series: { Sales: P[0] }, uncoloured: ["Cost"] } },
    // Double donut (NeenOpal): inner ring by category, outer ring in lighter shades of it
    { id: "double donut: blank category, outer ring in shades (CAPSTONE report)",
      v: visual("d1", DONUT, { value: ["S.Supplied CBM"], category: ["S.Year"], subcategory: ["S.Quarter"] }),
      rows: [["Supplied CBM", "Year", "Quarter"], ["181.26", "", ""], ["11126.49", "2021", "Qtr 2"], ["4995.39", "2021", "Qtr 3"], ["19875.18", "2021", "Qtr 4"]],
      expect: { keyCol: "Year", col: "Year", rows: [P[0], P[1], P[1], P[1]],
        rings: { col: "Quarter", rows: [mixHex(P[2], "#FFFFFF", 0.3), mixHex(P[1], "#FFFFFF", 0.3), mixHex(P[1], "#FFFFFF", 0.55), mixHex(P[1], "#FFFFFF", 0.8)] } } },
    { id: "double donut: colour set for a category still uses up its theme colour",
      v: visual("d2", DONUT, { value: ["S.Sales"], category: ["S.Region"], subcategory: ["S.Segment"] }, { categoryColors: [fill(lit("#00FFFF"), item("S.Region", "South"))] }),
      rows: [["Region", "Segment", "Sales"], ["North", "A", "1"], ["North", "B", "2"], ["South", "A", "3"], ["West", "A", "4"], ["West", "B", "5"], ["West", "C", "6"]],
      expect: { keyCol: "Region", col: "Region", rows: [P[0], P[0], "#00FFFF", P[2], P[2], P[2]],
        rings: { col: "Segment", rows: [mixHex(P[0], "#FFFFFF", 0.3), mixHex(P[0], "#FFFFFF", 0.8), mixHex("#00FFFF", "#FFFFFF", 0.3),
          mixHex(P[2], "#FFFFFF", 0.3), mixHex(P[2], "#FFFFFF", 0.55), mixHex(P[2], "#FFFFFF", 0.8)] } } },
    { id: "double donut without the report file: the grouped column is the category", noFile: true,
      v: visual("d3", DONUT, { value: ["S.Supplied CBM"], category: ["S.Year"], subcategory: ["S.Quarter"] }),
      rows: [["Supplied CBM", "Quarter", "Year"], ["181.26", "", ""], ["11126.49", "Qtr 2", "2021"], ["4995.39", "Qtr 3", "2021"], ["19875.18", "Qtr 4", "2021"]],
      expect: { keyCol: "Year", col: "Year", rows: [P[0], P[1], P[1], P[1]],
        rings: { col: "Quarter", rows: [mixHex(P[2], "#FFFFFF", 0.3), mixHex(P[1], "#FFFFFF", 0.3), mixHex(P[1], "#FFFFFF", 0.55), mixHex(P[1], "#FFFFFF", 0.8)] } } },
    { id: "double donut: date categories take new theme colours on the outer ring",
      v: visual("d4", DONUT, { value: ["S.Sales"], category: ["S.Date"], subcategory: ["S.Segment"] }),
      rows: [["Date", "Segment", "Sales"], ["1/1/2024", "A", "1"], ["1/1/2024", "B", "2"], ["2/1/2024", "A", "3"]],
      expect: { keyCol: "Date", col: "Date", rows: [P[0], P[0], P[1]],
        rings: { col: "Segment", rows: [mixHex(P[2], "#FFFFFF", 0.3), mixHex(P[2], "#FFFFFF", 0.8), mixHex(P[3], "#FFFFFF", 0.3)] } } },
  ];

  const hexOf = (argb) => (argb ? "#" + String(argb).slice(2).toUpperCase() : null);

  window.__colourCases = async () => {
    const design = await buildReport(cases.map((c) => c.v));
    const results = [];
    for (const c of cases) {
      const info = c.noFile ? null : visualInfo(design, "page1", c.v.name);
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
        if (c.expect.keyCol && (!colors.dot || header[colors.dot.keyCol] !== c.expect.keyCol)) problems.push(`colour key ${colors.dot ? header[colors.dot.keyCol] : "none"} ≠ ${c.expect.keyCol}`);
        if (c.expect.rings) {
          const rings = colors.rings;
          if (!rings || header[rings.col] !== c.expect.rings.col) problems.push(`outer ring on ${rings ? header[rings.col] : "nothing"} ≠ ${c.expect.rings.col}`);
          else c.expect.rings.rows.forEach((hex, ri) => { if (hexOf(rings.colors[ri]) !== hex.toUpperCase()) problems.push(`outer row ${ri + 1} ${hexOf(rings.colors[ri])} ≠ ${hex}`); });
        }
        if (c.expect.note && !colors.note) problems.push("no approximation note");
        for (const h of c.expect.uncoloured || []) if (colors.series[header.indexOf(h)]) problems.push(`${h} shouldn't be coloured`);
      }
      results.push({ id: c.id, ok: !problems.length, problems });
    }

    // Tables, matrices and cards: a full export, read back
    results.push(...await pageCases());
    results.push(...await tableCases());
    return { passed: results.filter((r) => r.ok).length, failed: results.filter((r) => !r.ok).length, results: results.filter((r) => !r.ok).length ? results : results.map((r) => r.id) };
  };

  /*
   * Visual names repeat across pages (VisualContainer1 on several pages), in both formats;
   * the classic format can name projections by query alias ("select") instead of the field.
   */
  async function pageCases() {
    const res = [];
    const t = (id, ok, detail) => res.push({ id, ok, problems: ok ? [] : [detail] });
    const rows = [["Month", "Segment", "Sales"], ["Jan", "A", "1"], ["Jan", "B", "2"]];

    // PBIR: the same visual name on two pages, coloured differently
    const zip = new JSZip();
    zip.file("Report/definition/report.json", JSON.stringify({ themeCollection: { baseTheme: { name: "CaseTheme" } } }));
    zip.file("Report/StaticResources/SharedResources/BaseThemes/CaseTheme.json", JSON.stringify(THEME));
    const sameName = (color) => visual("VisualContainer1", "clusteredColumnChart", { Category: ["S.Month"], Series: ["S.Segment"], Y: ["S.Sales"] }, { dataPoint: [fill(lit(color), item("S.Segment", "B"))] });
    zip.file("Report/definition/pages/aaa/page.json", JSON.stringify({ name: "ReportSectionA" }));
    zip.file("Report/definition/pages/aaa/visuals/VisualContainer1/visual.json", JSON.stringify(sameName("#111111")));
    zip.file("Report/definition/pages/bbb/page.json", JSON.stringify({ name: "ReportSectionB" }));
    zip.file("Report/definition/pages/bbb/visuals/VisualContainer1/visual.json", JSON.stringify(sameName("#222222")));
    const pbir = await readDesign(await JSZip.loadAsync(await zip.generateAsync({ type: "arraybuffer" })), null);
    const colourOn = (design, page) => {
      const c = columnColors({ rows, type: "clusteredColumnChart" }, design.theme.dataColors, visualInfo(design, page, "VisualContainer1"));
      return c && c.dot ? hexOf(c.dot.colorOf("B")) : null;
    };
    t("PBIR: same visual name on two pages", colourOn(pbir, "ReportSectionA") === "#111111" && colourOn(pbir, "ReportSectionB") === "#222222",
      `${colourOn(pbir, "ReportSectionA")} / ${colourOn(pbir, "ReportSectionB")}`);

    // Classic: the same name on two sections, and a pie whose category is the alias "select"
    const col = (table, prop) => ({ Column: { Expression: { SourceRef: { Source: "s" } }, Property: prop } });
    const vc = (name, visualType, select, projections, dataPoint) => ({ config: JSON.stringify({ name, singleVisual: {
      visualType, projections, prototypeQuery: { Select: select }, objects: { dataPoint } } }) });
    const legacySel = [{ ...col("S", "Month"), Name: "S.Month" }, { ...col("S", "Segment"), Name: "S.Segment" }, { Measure: { Expression: { SourceRef: { Source: "s" } }, Property: "Sales" }, Name: "S.Sales" }];
    const itemFill = (v, color) => fill(lit(color), item("S.Segment", v));
    const layout = {
      config: JSON.stringify({ themeCollection: { baseTheme: { name: "CaseTheme" } } }),
      sections: [
        { name: "ReportSection1", visualContainers: [vc("VisualContainer1", "clusteredColumnChart", legacySel, { Category: [{ queryRef: "S.Month" }], Series: [{ queryRef: "S.Segment" }], Y: [{ queryRef: "S.Sales" }] }, [itemFill("B", "#333333")])] },
        { name: "ReportSection2", visualContainers: [
          vc("VisualContainer1", "clusteredColumnChart", legacySel, { Category: [{ queryRef: "S.Month" }], Series: [{ queryRef: "S.Segment" }], Y: [{ queryRef: "S.Sales" }] }, [itemFill("B", "#444444")]),
          vc("VisualContainer9", "pieChart", [{ ...col("S", "Segment"), Name: "select" }, { Measure: { Expression: { SourceRef: { Source: "s" } }, Property: "Sales" }, Name: "S.Sales" }],
            { Category: [{ queryRef: "select" }], Y: [{ queryRef: "S.Sales" }] }, [itemFill("B", "#555555")]),
        ] },
      ],
    };
    const text = JSON.stringify(layout);
    const utf16 = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) { const ch = text.charCodeAt(i); utf16[2 * i] = ch & 255; utf16[2 * i + 1] = ch >> 8; }
    const lz = new JSZip();
    lz.file("Report/Layout", utf16);
    lz.file("Report/StaticResources/SharedResources/BaseThemes/CaseTheme.json", JSON.stringify(THEME));
    const classic = await readDesign(await JSZip.loadAsync(await lz.generateAsync({ type: "arraybuffer" })), null);
    t("classic: same visual name on two pages", colourOn(classic, "ReportSection1") === "#333333" && colourOn(classic, "ReportSection2") === "#444444",
      `${colourOn(classic, "ReportSection1")} / ${colourOn(classic, "ReportSection2")}`);
    // The pie's category is the alias "select": it must still find the Segment column, even when it isn't first
    const pieRows = [["Sales", "Segment"], ["1", "A"], ["2", "B"]];
    const pie = columnColors({ rows: pieRows, type: "pieChart" }, classic.theme.dataColors, visualInfo(classic, "ReportSection2", "VisualContainer9"));
    t("classic: field named by query alias", pie && pie.dot && pie.dot.keyCol === 1 && hexOf(pie.dot.colorOf("B")) === "#555555" && hexOf(pie.dot.colorOf("A")) === P[0],
      pie && pie.dot ? `key ${pie.dot.keyCol} A ${hexOf(pie.dot.colorOf("A"))} B ${hexOf(pie.dot.colorOf("B"))}` : "no colours");
    return res;
  }

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
      visual("t3", "pivotTable", { Rows: ["S.Region"], Values: [{ ref: "S.Seps YoY", label: "YoY" }, { ref: "S.Actives YoY", label: "YoY" }], Columns: ["S.Reason"] }, {
        values: [{ properties: { backColor: solid(lit("#00FFFF")) }, selector: { metadata: "S.Actives YoY" } }],
      }),
      // Where colours land: legend → category cells, measures → headers, scales → values
      visual("p1", "clusteredColumnChart", { Category: ["S.Week"], Series: ["S.FP"], Y: ["S.Hires"] }, { dataPoint: [fill(lit("#F5C869"), item("S.FP", "Part-Time"))] }),
      visual("p2", "lineChart", { Category: ["S.Day"], Y: ["S.Revenue2", "S.Spend"] }),
      visual("p3", "clusteredBarChart", { Category: ["S.Store"], Y: ["S.Amount"] }, { dataPoint: [fill(grad("S.Amount", "#000000", "#FFFFFF"), all)] }),
      visual("k2", "cardVisual", { Data: ["S.Profit"] }, { accentBar: [{ properties: { color: solid(themeColor(3)) }, selector: { id: "default" } }] }),
      visual("t4", "tableEx", { Values: ["S.Item", "S.Delta", "S.Score"] }, { columnFormatting: [
        { properties: { fontColor: solid(lit("#AA00AA")) }, selector: meta("S.Item") },
        { properties: { dataBars: { positiveColor: solid(lit("#00AA00")), negativeColor: solid(lit("#EA1313")), axisColor: solid(lit("#333333")), reverseDirection: { expr: { Literal: { Value: "true" } } } } }, selector: meta("S.Delta") },
      ] }),
      visual("k1", "cardVisual", { Data: ["S.Sales", "S.Cost"] }, {
        value: [{ properties: { fontColor: solid(lit("#123456")) }, selector: meta("S.Sales") }, { properties: { fontColor: solid(cond([[cmp("S.Cost", 1, num(0)), "#00AA00"]])) }, selector: meta("S.Cost") }],
      }),
      visual("dd", DONUT, { value: ["S.Supplied CBM"], category: ["S.Year"], subcategory: ["S.Quarter"] }),
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
      t3: [["Region", "YoY", "YoY", "Reason"], ["North", "1", "2", "Voluntary"]],
      k1: [["Sales", "Cost"], ["100", "7"]],
      p1: [["Week", "FP", "Hires"], ["W1", "Full-Time", "3"], ["W1", "Part-Time", "4"]],
      p2: [["Day", "Revenue2", "Spend"], ["D1", "1", "2"]],
      p3: [["Store", "Amount"], ["A", "0"], ["B", "10"]],
      k2: [["Profit"], ["42"]],
      t4: [["Item", "Delta", "Score"], ["One", "-5", "1"], ["Two", "8", "2"]],
      s1: [["Region"], ["North"], ["South"]],
      c1: [["Units Sold", "Channel"], ["10", "Online"], ["20", "Retail"]],
      dd: [["Supplied CBM", "Year", "Quarter"], ["181.26", "", ""], ["11126.49", "2021", "Qtr 2"], ["4995.39", "2021", "Qtr 3"], ["19875.18", "2021", "Qtr 4"]],
    };
    const page = { name: "page1", displayName: "Cases", visibility: 0, getFilters: async () => [], setActive: async () => {},
      getVisuals: async () => [
        { name: "t1", title: "Table case", type: "tableEx", layout: { x: 0, y: 200, width: 500, height: 200 }, exportData: async () => ({ data: csv(data.t1) }) },
        { name: "t2", title: "Matrix case", type: "pivotTable", layout: { x: 600, y: 200, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.t2) }) },
        { name: "t3", title: "Matrix with repeated names", type: "pivotTable", layout: { x: 0, y: 500, width: 500, height: 200 }, exportData: async () => ({ data: csv(data.t3) }) },
        { name: "k1", title: "Card", type: "cardVisual", layout: { x: 0, y: 0, width: 300, height: 100 }, exportData: async () => ({ data: csv(data.k1) }) },
        { name: "p1", title: "Legend chart", type: "clusteredColumnChart", layout: { x: 0, y: 800, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.p1) }) },
        { name: "p2", title: "Measure chart", type: "lineChart", layout: { x: 400, y: 800, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.p2) }) },
        { name: "p3", title: "Scale chart", type: "clusteredBarChart", layout: { x: 800, y: 800, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.p3) }) },
        { name: "k2", title: "Card", type: "cardVisual", layout: { x: 320, y: 0, width: 200, height: 100 }, exportData: async () => ({ data: csv(data.k2) }) },
        { name: "t4", title: "Column formats", type: "tableEx", layout: { x: 0, y: 1100, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.t4) }) },
        // a button slicer: a filter chip, not a table
        { name: "s1", title: "Region", type: "advancedSlicerVisual", layout: { x: 900, y: 0, width: 200, height: 80 },
          getSlicerState: async () => ({ filters: [{ target: { table: "S", column: "Region" }, filterType: 1, operator: "In", values: ["North"] }], targets: [{ column: "Region" }] }),
          exportData: async () => ({ data: csv(data.s1) }) },
        // hidden: a shape (no note) and a chart (a note)
        { name: "h1", title: "Toggle shape", type: "shape", layout: { x: 0, y: 1400, width: 50, height: 50, displayState: { mode: 1 } }, exportData: async () => { throw { message: "no data" }; } },
        { name: "h2", title: "Hidden chart", type: "lineChart", layout: { x: 0, y: 1500, width: 50, height: 50, displayState: { mode: 1 } }, exportData: async () => ({ data: csv(data.p2) }) },
        // a card split by category, untitled: its title comes from its fields
        { name: "c1", title: "", type: "cardVisual", layout: { x: 600, y: 1100, width: 300, height: 200 }, exportData: async () => ({ data: csv(data.c1) }) },
        { name: "dd", title: "Supplied CBM by Year and Quarter", type: DONUT, layout: { x: 0, y: 1700, width: 600, height: 400 }, exportData: async () => ({ data: csv(data.dd) }) },
      ] };
    embeddedReport = { getActivePage: async () => page, getPages: async () => [page], getFilters: async () => [], getTheme: async () => ({ themeJson: THEME }), on() {}, off() {} };
    currentReport = { id: "cases", embedUrl: "https://x/?groupId=g" };
    currentReportName = "Cases";
    reportColorCache.clear();
    downloadDesignFile = async () => buf;
    let href; const orig = HTMLAnchorElement.prototype.click; HTMLAnchorElement.prototype.click = function () { href = this.href; };
    let result;
    try {
      result = await exportToExcel([page], { pages: "current", layout: "report", collapse: true, visibleRows: 8, includeFilters: true, includeNotes: true, matchColors: true },
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
    const yoy2 = find("YoY (2)"), yoy1 = find("YoY");
    t("matrix: second of two values renamed alike, column field last", yoy2 && fillAt(yoy2.r + 1, yoy2.c) === "#00FFFF" && fillAt(yoy1.r + 1, yoy1.c) !== "#00FFFF",
      yoy2 ? `${fillAt(yoy1.r + 1, yoy1.c)} ${fillAt(yoy2.r + 1, yoy2.c)}` : "no YoY (2) column");
    const fp = find("FP"), hires = find("Hires"), rev = find("Revenue2"), spend = find("Spend"), store = find("Store"), amount = find("Amount");
    t("legend colours on the category cells", fp && fillAt(fp.r + 1, fp.c) === P[0] && fillAt(fp.r + 2, fp.c) === "#F5C869",
      fp && `${fillAt(fp.r + 1, fp.c)} ${fillAt(fp.r + 2, fp.c)}`);
    t("legend charts keep the numbers neutral", hires && !fillAt(hires.r + 1, hires.c) && !fillAt(hires.r + 2, hires.c), hires && `${fillAt(hires.r + 1, hires.c)}`);
    t("measure series colour their headers, numbers neutral",
      rev && fillAt(rev.r, rev.c) === P[0] && fillAt(spend.r, spend.c) === P[1] && !fillAt(rev.r + 1, rev.c) && !fillAt(spend.r + 1, spend.c),
      rev && `${fillAt(rev.r, rev.c)} ${fillAt(spend.r, spend.c)} values ${fillAt(rev.r + 1, rev.c)} ${fillAt(spend.r + 1, spend.c)}`);
    t("colour scales stay on the values", amount && fillAt(amount.r + 1, amount.c) === "#000000" && fillAt(amount.r + 2, amount.c) === "#FFFFFF" && !fillAt(store.r + 1, store.c),
      amount && `${fillAt(amount.r + 1, amount.c)} ${fillAt(amount.r + 2, amount.c)} store ${fillAt(store.r + 1, store.c)}`);
    let salesValue = null, costValue = null;
    ws.eachRow((row) => row.eachCell((cell) => {
      if (cell.value === 100 && cell.font && cell.font.size >= 20) salesValue = hexOf(cell.font.color.argb);
      if (cell.value === 7 && cell.font && cell.font.bold) costValue = hexOf(cell.font.color.argb);
    }));
    t("card: value colour set by hand", salesValue === "#123456", salesValue);
    // The card's tile cap takes its value colour (a top border on the pale tile fill)
    let capColor = null;
    ws.eachRow((row) => row.eachCell((cell) => {
      const b = cell.border && cell.border.top;
      if (!capColor && b && b.style === "medium" && cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb === "FFF3F5F8" && cell.value === "SALES") capColor = hexOf(b.color.argb);
    }));
    t("card: tile cap in the card's colour", capColor === "#123456", capColor);
    t("card: value colour from a rule", costValue === "#00AA00", costValue);
    // The new card's accent bar colours its tile's cap
    let profitCap = null;
    ws.eachRow((row) => row.eachCell((cell) => {
      const b = cell.border && cell.border.top;
      if (cell.value === "PROFIT" && b && b.style === "medium") profitCap = hexOf(b.color.argb);
    }));
    t("card: accent bar colours the tile cap", profitCap === P[1], profitCap);
    // Column formats: a column's font colour; data bars with negative colour, axis, direction
    const itemCol = find("Item");
    t("table: a column's own font colour", itemCol && ws.getCell(itemCol.r + 1, itemCol.c).font.color.argb === "FFAA00AA", itemCol && ws.getCell(itemCol.r + 1, itemCol.c).font.color.argb);
    t("table: data bars keep negative colour, axis and direction",
      /negativeBarColorSameAsPositive="0"[^>]*direction="rightToLeft"[\s\S]*?FFEA1313[\s\S]*?axisColor rgb="FF333333"/.test(xml), (xml.match(/<x14:dataBar[\s\S]*?<\/x14:dataBar>/g) || []).join(" ").slice(0, 300));
    // Slicers of every kind become filter chips; hidden decorative visuals add no note
    const allText = []; ws.eachRow((row) => row.eachCell((cell) => allText.push(cell.value && cell.value.richText ? cell.value.richText.map((x) => x.text).join("") : String(cell.value ?? ""))));
    t("button slicer becomes a filter chip", allText.some((x) => /Region\s+North/.test(x)) && !Object.values(ws.tables).some((x) => x.table.columns.length === 1 && x.table.columns[0].name === "Region"),
      allText.filter((x) => /FILTERS/.test(x)).join(" / "));
    t("hidden shape adds no note, hidden chart does", !result.notes.some((n) => /Toggle shape/.test(n)) && result.notes.some((n) => /Hidden chart/.test(n)), JSON.stringify(result.notes));
    t("untitled card split by category is titled from its fields", allText.includes("Units Sold by Channel"), allText.filter((x) => /Units Sold/.test(x)).join(" / "));
    // Double donut: the category cells in the inner ring's colours, the subcategory cells in
    // the outer ring's shades, the values neutral, the inner ring as the caption legend
    const year = find("Year"), quarter = find("Quarter"), cbm = find("Supplied CBM");
    const shade = (hex, amount) => mixHex(hex, "#FFFFFF", amount);
    const yearFills = year ? [1, 2, 3, 4].map((k) => fillAt(year.r + k, year.c)) : [];
    const quarterFills = quarter ? [1, 2, 3, 4].map((k) => fillAt(quarter.r + k, quarter.c)) : [];
    t("double donut: category cells in the inner ring's colours", yearFills.join() === [P[0], P[1], P[1], P[1]].join(), yearFills.join(" "));
    t("double donut: subcategory cells in the outer ring's shades",
      quarterFills.join() === [shade(P[2], 0.3), shade(P[1], 0.3), shade(P[1], 0.55), shade(P[1], 0.8)].join(), quarterFills.join(" "));
    t("double donut: values and their header stay neutral", cbm && !fillAt(cbm.r + 1, cbm.c) && fillAt(cbm.r, cbm.c) === "#F3F5F8", cbm && `${fillAt(cbm.r, cbm.c)} ${fillAt(cbm.r + 1, cbm.c)}`);
    const donutCap = cbm ? ws.getCell(cbm.r - 1, cbm.c).value : null;
    const capRuns = (donutCap && donutCap.richText) || [];
    const keyColor = (label) => {
      const at = capRuns.findIndex((x) => x.text.trim() === label);
      return at > 0 ? hexOf(capRuns[at - 1].font.color.argb) : null;
    };
    t("double donut: caption legend lists the inner ring, blank as (Blank)", keyColor("(Blank)") === P[0] && keyColor("2021") === P[1],
      capRuns.map((x) => x.text).join("|"));
    return res;
  }
})();
