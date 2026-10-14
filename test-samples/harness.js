/*
 * Dev-only test harness for the Excel export (not loaded by the app).
 * Open http://localhost:5500, then in the browser console:
 *   await import("./test-samples/harness.js?v=" + Date.now())   // or add a <script> tag
 *   const run = await __run([0], {}, {});      // pages by index, branding, options
 *   await __check(run);                        // structural assertions
 *   __render(run.wb.worksheets[0]);            // draws the sheet into the page
 *   __colors(run.wb.worksheets[0]);            // each table's header / cell / gutter colours
 * Fixtures are shaped like the test report in test-samples/README.md. The report file
 * for exact colours is stubbed: opts.__design = "pbix" (test sample.pbix, PBIR format),
 * "legacy" (a generated classic-format file), "noaccess", or omitted (no permission).
 */
(() => {
  const months = Array.from({ length: 24 }, (_, i) => new Date(2024, i, 1).toLocaleString("en-US", { month: "short", year: "numeric" }));
  const cats = ["Apparel", "Electronics", "Furniture", "Office Supplies"];
  const segs = ["Consumer", "Corporate", "Home Office"];
  const csv = (rows) => rows.map((r) => r.map((v) => (/[",]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v)).join(",")).join("\n");

  const fx = {
    card: csv([["Total Sales"], ["2517775.17"]]),
    card2: csv([["Total Orders"], ["63214"]]),
    card3: csv([["Margin %"], ["43.97%"]]),
    line: csv([["Month", "Total Sales", "Total Cost"], ...months.map((m, i) => [m, (93687.3 + i * 1234.5).toFixed(2), (57797.13 + i * 800).toFixed(2)])]),
    seg: csv([["Month", "Segment", "Total Sales"], ...months.flatMap((m, i) => segs.map((s, j) => [m, s, (45000 - j * 12000 + i * 300).toFixed(2)]))]),
    tree: csv([["Category", "Total Sales"], ...cats.map((c, i) => [c, (900000 - i * 150000).toFixed(2)])]),
    deneb: csv([["Month", "Month Sort", "Total Sales", "Bar Color"], ...months.map((m, i) => [m, 202401 + (i % 12) + Math.floor(i / 12) * 100, (93687.3 + i * 1234.5).toFixed(2), "#0B8E37"])]),
    table: csv([["Category", "Total Sales", "Total Cost", "Margin %"], ...cats.map((c, i) => [c, (900000 - i * 150000).toFixed(2), (500000 - i * 90000).toFixed(2), `${(40 + i).toFixed(1)}%`])]),
    matrix: csv([["Category", "2024", "2025"], ...cats.map((c, i) => [c, 400000 - i * 50000, 450000 - i * 50000])]),
    long: csv([["Date", "Category", "Segment", "Region", "Sales"], ...Array.from({ length: 576 }, (_, i) => [`2024-${String((i % 12) + 1).padStart(2, "0")}-01`, cats[i % 4], segs[i % 3], i % 2 ? "North" : "South", (1000 + i * 3.5).toFixed(2)])]),
  };

  const mkV = (name, title, type, x, y, w, h, data, filters) => ({
    name, title, type, layout: { x, y, width: w, height: h },
    exportData: async (t, n) => (data === null ? Promise.reject({ message: "no data" }) : { data: n === 1 ? data.split("\n")[0] + "\n" : data }),
    getSlicerState: async () => ({ filters: filters || [], targets: [{ column: "Segment" }] }),
  });

  // Page changes fire "rendered" like Power BI does, so activatePage doesn't wait for its timeout
  const rendered = new Set();
  const setActive = async () => { setTimeout(() => [...rendered].forEach((fn) => fn()), 0); };
  const pages = () => [
    { name: "p1", displayName: "Colour test", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [
        mkV("s1", "Slicer", "slicer", 900, 0, 200, 80, "Region\nNorth", [{ target: { column: "Region" }, values: ["North"] }]),
        mkV("s2", "Slicer", "slicer", 1120, 0, 200, 80, "Date\nx", [{ target: { column: "Date" }, logicalOperator: "And", conditions: [{ operator: "GreaterThanOrEqual", value: "2024-05-31T18:30:00.000Z" }, { operator: "LessThan", value: "2025-08-31T18:30:00.000Z" }] }]),
        mkV("s3", "Slicer", "slicer", 1340, 0, 120, 80, "Segment\nx", []),
        mkV("k1", "Card", "card", 0, 0, 200, 100, fx.card),
        mkV("k2", "Card", "card", 210, 0, 200, 100, fx.card2),
        mkV("k3", "Card", "card", 420, 0, 200, 100, fx.card3),
        mkV("l1", "Sales and cost by month", "lineChart", 0, 120, 600, 300, fx.line),
        mkV("t1", "Sales by category tiles", "treemap", 620, 120, 400, 300, fx.tree),
        mkV("g1", "Sales by segment", "clusteredColumnChart", 0, 440, 600, 300, fx.seg),
        mkV("d1", "Deneb", "deneb7E15AEF80B9E4D4F8E12924291ECE89A", 620, 440, 600, 300, fx.deneb),
        mkV("c1", "Category summary", "tableEx", 0, 760, 600, 200, fx.table),
        mkV("m1", "", "pivotTable", 620, 760, 400, 200, fx.matrix),
        mkV("x1", "Text box", "textbox", 0, 0, 100, 30, null),
      ] },
    { name: "p2", displayName: "Long table", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [mkV("a1", "All rows", "tableEx", 0, 0, 800, 600, fx.long), mkV("a2", "Category totals", "tableEx", 820, 0, 400, 200, fx.tree)] },
    { name: "p3", displayName: "O'Brien", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [mkV("o1", "Orders by category", "clusteredColumnChart", 0, 0, 600, 300, fx.tree)] },
    // The "Colour test" page of test-samples/test sample.pbix: same visual names and positions,
    // so the colours read from that PBIR file apply (pass { __design: "pbix" })
    { name: "e25f46948c1a0a8e9564", displayName: "Colour test (pbix)", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [
        mkV("ff7e126096feeb677190", "Card", "cardVisual", 0, 0, 216, 168, fx.card),
        mkV("56a2092fd1f3fab4ebcf", "Sales and cost by month", "lineChart", 216, 0, 408, 456, fx.line),
        mkV("915cd3320ac20ff8360c", "Sales by segment", "clusteredColumnChart", 624, 0, 696, 456, fx.seg),
        mkV("a1e5436cc4b9b7d45203", "Sales by category (custom colour)", "clusteredBarChart", 1320, 0, 600, 456, fx.tree),
        mkV("062e0c7ae2af004ddb40", "Deneb", "deneb7E15AEF80B9E4D4F8E12924291ECE89A", 0, 456, 1920, 624, fx.deneb),
      ] },
    // Classic-format report (pass { __design: "legacy" }): colours set per slice, per measure, per funnel bar
    { name: "lg", displayName: "Legacy format", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [
        mkV("lp1", "Sales by category pie", "pieChart", 0, 0, 400, 300, fx.tree),
        mkV("ll1", "Sales and cost", "lineChart", 420, 0, 600, 300, fx.line),
        mkV("lf1", "Funnel", "funnel", 0, 320, 400, 300, fx.tree),
      ] },
  ];

  // Layout edge case: a pie with row tabs snaps under a bar chart with none, so its tab
  // column is the bar chart's (wide) first column – tabs must become swatches, not blocks
  const extraPages = () => [
    { name: "edge", displayName: "Layout edge", visibility: 0, getFilters: async () => [], setActive,
      getVisuals: async () => [
        mkV("e1", "Sales and cost", "lineChart", 0, 0, 560, 300, fx.line),
        mkV("e2", "Sales by category bars", "clusteredBarChart", 600, 0, 500, 300, fx.tree),
        mkV("e3", "Category summary", "tableEx", 0, 320, 560, 300, fx.table),
        mkV("e4", "Sales by category pie", "pieChart", 600, 320, 500, 300, fx.tree),
      ] },
  ];

  /* A classic-format design file: Report/Layout (UTF-16) and a base theme */
  const legacyDesign = async () => {
    const theme = { name: "CY24SU10", dataColors: ["#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2", "#D9B300", "#D64550"], background: "#FFFFFF", foreground: "#252423" };
    const col = (p) => ({ Column: { Expression: { SourceRef: { Entity: "Sales" } }, Property: p } });
    const item = (v, color) => ({ properties: { fill: { solid: { color: { expr: color } } } },
      selector: { data: [{ scopeId: { Comparison: { ComparisonKind: 0, Left: col("Category"), Right: { Literal: { Value: `'${v}'` } } } } }] } });
    const vc = (name, visualType, projections, dataPoint) => ({ config: JSON.stringify({ name, singleVisual: { visualType, projections, objects: dataPoint ? { dataPoint } : {} } }) });
    const layout = {
      config: JSON.stringify({ themeCollection: { baseTheme: { name: "CY24SU10" } } }),
      sections: [{ displayName: "Legacy format", visualContainers: [
        vc("lp1", "pieChart", { Category: [{ queryRef: "Sales.Category" }], Y: [{ queryRef: "Sales.Total Sales" }] },
          [item("Furniture", { Literal: { Value: "'#123456'" } })]),
        vc("ll1", "lineChart", { Category: [{ queryRef: "Sales.Month" }], Y: [{ queryRef: "Sales.Total Sales" }, { queryRef: "Sales.Total Cost" }] },
          [{ properties: { fill: { solid: { color: { expr: { ThemeDataColor: { ColorId: 5, Percent: -0.25 } } } } } }, selector: { metadata: "Sales.Total Cost" } }]),
        vc("lf1", "funnel", { Category: [{ queryRef: "Sales.Category" }], Y: [{ queryRef: "Sales.Total Sales" }] },
          [item("Electronics", { Literal: { Value: "'#FF00AA'" } })]),
      ] }],
    };
    const text = JSON.stringify(layout);
    const utf16 = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); utf16[2 * i] = c & 255; utf16[2 * i + 1] = c >> 8; }
    const zip = new JSZip();
    zip.file("Report/Layout", utf16);
    zip.file("Report/StaticResources/SharedResources/BaseThemes/CY24SU10.json", JSON.stringify(theme));
    return zip.generateAsync({ type: "arraybuffer" });
  };

  /* Where the report file comes from: the sample pbix, a classic-format file, or a refusal */
  const stubDesign = async (mode) => {
    reportColorCache.clear();
    if (mode === "pbix") {
      const buf = await (await fetch("test-samples/test%20sample.pbix", { cache: "reload" })).arrayBuffer();
      downloadDesignFile = async () => buf;
    } else if (mode === "legacy") {
      const buf = await legacyDesign();
      downloadDesignFile = async () => buf;
    } else {
      downloadDesignFile = async () => { throw new Error(mode === "noaccess" ? "NO_ACCESS" : "NO_PERMISSION"); };
    }
  };

  /* Per table: header fills, first data row fills (and text colours), gutter fill, CF rule count */
  window.__colors = (ws) => Object.values(ws.tables || {}).map((x) => {
    const t = x.table, m = t.tableRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    const r0 = +m[2], c0 = colN(m[1]);
    const fillOf = (cell) => (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb ? "#" + cell.fill.fgColor.argb.slice(2) : "-");
    const fontOf = (cell) => (cell.font && cell.font.color && cell.font.color.argb ? "#" + cell.font.color.argb.slice(2) : "-");
    const cf = (ws.conditionalFormattings || []).find((f) => f.ref.startsWith(ws.getColumn(c0 - 1).letter + (r0 + 1)));
    return {
      table: t.name,
      header: t.columns.map((c, i) => `${c.name}=${fillOf(ws.getCell(r0, c0 + i))}`).join(" | "),
      rows: [1, 2, 3].map((k) => t.columns.map((c, i) => `${fillOf(ws.getCell(r0 + k, c0 + i))}/${fontOf(ws.getCell(r0 + k, c0 + i))}`).join(" | ")),
      gutter: [1, 2, 3].map((k) => fillOf(ws.getCell(r0 + k, c0 - 1))).join(" "),
      cfRules: cf ? cf.rules.length : 0,
    };
  });

  /* Runs exportToExcel against the fixtures and reads the file back */
  window.__run = async (pageIdx, branding = {}, opts = {}) => {
    const all = [...pages(), ...extraPages()];
    const { __design, ...rest } = opts;
    opts = rest;
    await stubDesign(__design);
    currentReport = { id: "harness", embedUrl: "https://app.powerbi.com/reportEmbed?reportId=harness&groupId=g1" };
    embeddedReport = {
      getActivePage: async () => all[0], getPages: async () => all, getFilters: async () => [],
      getTheme: async () => ({ themeJson: { dataColors: ["#E41A1C", "#377EB8", "#4DAF4A", "#984EA3", "#FF7F00", "#A65628", "#F781BF", "#999999"] } }),
      on(event, fn) { if (event === "rendered") rendered.add(fn); }, off(event, fn) { rendered.delete(fn); },
    };
    currentReportName = "Export colour test";
    let href, name;
    const original = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () { href = this.href; name = this.download; };
    const options = { pages: "current", layout: "report", collapse: true, visibleRows: 8, includeFilters: true, includeNotes: true, matchColors: true, exhibitNumbers: false, ...opts };
    try {
      const res = await exportToExcel(pageIdx.map((i) => all[i]), options,
        { companyName: "", accent: "#0B7F56", tableStyle: "Executive", titleFont: "Calibri Light", footer: "", ...branding }, () => {}, () => false);
      const buf = await (await fetch(href)).arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      return { res, wb, buf, name };
    } finally {
      HTMLAnchorElement.prototype.click = original;
    }
  };

  /* Reads the raw XML parts of an .xlsx (central directory + deflate-raw) */
  window.__unzip = async (buf) => {
    const dv = new DataView(buf), u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let k = 0; k < count; k++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.slice(p + 46, p + 46 + nlen));
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const data = u8.slice(start, start + csize);
      files[name] = method === 0 ? new TextDecoder().decode(data)
        : await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
      p += 46 + nlen + elen + clen;
    }
    return files;
  };

  const colN = (s) => s.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const textOf = (v) => (v && v.richText ? v.richText.map((x) => x.text).join("") : String(v ?? ""));

  /* Structural assertions from the design plan */
  window.__check = async ({ wb, buf }) => {
    const errors = [];
    const files = await window.__unzip(buf);
    const allXml = Object.entries(files).filter(([n]) => /\.(xml|rels)$/.test(n)).map(([, t]) => t).join("\n");
    if (/neenopal/i.test(allXml)) errors.push("product name found in XML");
    if (/outlinePr/.test(allXml)) errors.push("outlinePr written");
    for (const ws of wb.worksheets) {
      const ranges = Object.values(ws.tables || {}).map((x) => {
        const t = x.table, m = t.tableRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        return { t, r0: +m[2], r1: +m[4], c0: colN(m[1]), c1: colN(m[3]) };
      });
      for (const g of ranges) {
        g.t.columns.forEach((c, i) => {
          const v = ws.getCell(g.r0, g.c0 + i).value;
          if (v !== c.name) errors.push(`${ws.name}: header mismatch in ${g.t.name}: ${JSON.stringify(v)} vs "${c.name}"`);
        });
      }
      for (const m of ws.model.merges || []) {
        const [a, b] = m.split(":"), ma = a.match(/([A-Z]+)(\d+)/), mb = b.match(/([A-Z]+)(\d+)/);
        const mr0 = +ma[2], mr1 = +mb[2], mc0 = colN(ma[1]), mc1 = colN(mb[1]);
        for (const g of ranges) if (mr0 <= g.r1 && g.r0 <= mr1 && mc0 <= g.c1 && g.c0 <= mc1) errors.push(`${ws.name}: merge ${m} overlaps ${g.t.name}`);
        for (let r = mr0; r <= mr1; r++) if (ws.getRow(r).hidden) errors.push(`${ws.name}: merge ${m} on hidden row ${r}`);
      }
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        if (!row.hidden) return;
        row.eachCell((cell, c) => {
          if (cell.value == null) return;
          const inTable = ranges.some((g) => r > g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1);
          if (!inTable && cell.value !== "●" && cell.value !== "■") errors.push(`${ws.name}: hidden row ${r} col ${c} holds "${textOf(cell.value).slice(0, 40)}"`);
        });
      });
    }
    return { errors, sheets: wb.worksheets.map((w) => w.name) };
  };

  /* Draws a worksheet as an HTML table (fonts, fills, borders, merges, widths, heights) */
  window.__render = (ws, { maxRows = 70, maxCols = 16, scale = 1 } = {}) => {
    const px = (pt) => Math.round(pt * 1.333);
    const hex = (a) => (typeof a === "string" ? "#" + a.slice(2) : null);
    const bw = { thin: 1, hair: 1, medium: 2, thick: 3 };
    const merges = {};
    const covered = new Set();
    for (const m of ws.model.merges || []) {
      const [a, b] = m.split(":"), ma = a.match(/([A-Z]+)(\d+)/), mb = b.match(/([A-Z]+)(\d+)/);
      const r0 = +ma[2], r1 = +mb[2], c0 = colN(ma[1]), c1 = colN(mb[1]);
      merges[`${r0},${c0}`] = { rs: r1 - r0 + 1, cs: c1 - c0 + 1 };
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (r !== r0 || c !== c0) covered.add(`${r},${c}`);
    }
    const fontCss = (f) => !f ? "" : `font-family:'${f.name || "Calibri"}',Calibri,sans-serif;font-size:${px(f.size || 11)}px;` +
      `${f.bold ? "font-weight:700;" : ""}${f.italic ? "font-style:italic;" : ""}${f.color && f.color.argb ? `color:${hex(f.color.argb)};` : ""}`;
    const fmt = (v, nf) => (typeof v === "number" ? displayText(v, nf) : textOf(v));
    const colPx = Array.from({ length: maxCols }, (_, i) => Math.round((ws.getColumn(i + 1).width || 9) * 7));
    // Chrome only honours fixed column widths when the table itself has a width
    let html = `<table style="border-collapse:collapse;table-layout:fixed;width:${colPx.reduce((a, b) => a + b, 0)}px;background:#fff;transform:scale(${scale});transform-origin:0 0">`;
    html += "<colgroup>" + colPx.map((w) => `<col style="width:${w}px">`).join("") + "</colgroup>";
    let shown = 0;
    for (let r = 1; shown < maxRows && r <= ws.rowCount + 2; r++) {
      const row = ws.getRow(r);
      if (row.hidden) continue;
      shown++;
      html += `<tr style="height:${px(row.height || 15)}px">`;
      for (let c = 1; c <= maxCols; c++) {
        if (covered.has(`${r},${c}`)) continue;
        const cell = ws.getCell(r, c);
        const m = merges[`${r},${c}`];
        const st = ["padding:0 3px", "white-space:nowrap", "overflow:visible", "vertical-align:" + ({ top: "top", middle: "middle", bottom: "bottom" }[cell.alignment && cell.alignment.vertical] || "bottom")];
        const fill = cell.fill && cell.fill.fgColor && hex(cell.fill.fgColor.argb);
        if (fill) st.push(`background:${fill}`);
        const b = cell.border || {};
        for (const side of ["top", "bottom", "left", "right"]) if (b[side]) st.push(`border-${side}:${bw[b[side].style] || 1}px solid ${hex(b[side].color && b[side].color.argb) || "#000"}`);
        const al = cell.alignment || {};
        if (al.horizontal) st.push(`text-align:${al.horizontal}`);
        else if (typeof cell.value === "number") st.push("text-align:right");
        if (al.indent) st.push(al.horizontal === "right" ? `padding-right:${al.indent * 9}px` : `padding-left:${al.indent * 9}px`);
        let inner;
        if (cell.value && cell.value.richText) inner = cell.value.richText.map((x) => `<span style="${fontCss(x.font)}">${x.text.replace(/</g, "&lt;")}</span>`).join("");
        else inner = `<span style="${fontCss(cell.font)}">${fmt(cell.value, cell.numFmt).replace(/</g, "&lt;")}</span>`;
        html += `<td${m ? ` rowspan="${m.rs}" colspan="${m.cs}"` : ""} style="${st.join(";")}">${inner}</td>`;
      }
      html += "</tr>";
    }
    html += "</table>";
    document.body.innerHTML = `<div id="__pv" style="padding:10px;background:#fff">${html}</div>`;
    document.body.style.background = "#fff";
    document.body.style.overflow = "auto";
  };
})();
