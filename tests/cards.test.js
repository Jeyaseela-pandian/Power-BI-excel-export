/* KPI tiles: an empty value reads "(Blank)", as in Power BI */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { run, json } = loadApp();

/* A card's tile as written: its label cell, main value cell and fill */
function tile(info, reportBackground) {
  run(`cells2 = {}; ws2 = { getCell: (r, c) => (cells2[r + "," + c] = cells2[r + "," + c] || {}), mergeCellsWithoutStyle() {} };
    visualInfo = () => (${JSON.stringify(info)});
    lay2 = { fonts: makeFonts(), reportColors: null, accent: { cap: "FF000000" }, reportBackground: ${reportBackground} };
    writeKpiBlock(ws2, 1, 1, { rows: [["Patient Satisfaction Score"], ["83.3%"]], page: "p", name: "n", type: "card", displayTitle: null }, { span() {}, want() {} }, lay2, false, null);`);
  return json(`({ label: cells2["1,1"].alignment, labelColor: cells2["1,1"].font.color.argb, valueColor: cells2["2,1"].font.color.argb, fill: cells2["1,1"].fill.fgColor.argb })`);
}

test("a card's label always fits its tile (shrunk when Excel draws it wider)", () => {
  assert.strictEqual(tile(null, false).label.shrinkToFit, true);
});

test("cards keep their usual tile by default; with the report's background, the card's own colour and readable text", () => {
  const usual = tile({ look: { frame: { background: "#1B1B3A", transparency: 0 } } }, false);
  assert.strictEqual(usual.fill, "FFF3F5F8", "default: the usual light tile");
  const dark = tile({ look: { frame: { background: "#1B1B3A", transparency: 0 } } }, true);
  assert.deepStrictEqual([dark.fill, dark.labelColor, dark.valueColor], ["FF1B1B3A", "FFFFFFFF", "FFFFFFFF"]);
  const light = tile({ look: { frame: { background: "#E8E4F8", transparency: 0 } } }, true);
  assert.strictEqual(light.fill, "FFE8E4F8");
  assert.notStrictEqual(light.labelColor, "FFFFFFFF", "dark text on a light card");
  const clear = tile({ look: { frame: { background: "#1B1B3A", transparency: 100 } } }, true);
  assert.strictEqual(clear.fill, "FFF3F5F8", "a transparent card background: the usual tile");
});

test("a card's empty values show (Blank); values that exist are unchanged", () => {
  run(`cells = {}; ws = { getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}), mergeCellsWithoutStyle() {} };
    widths = { span() {}, want() {} }; lay = { fonts: new Proxy({}, { get: () => ({}) }), reportColors: null, accent: { cap: "FF000000" } };
    card = { rows: [["Sum of Sales", "Target", "Variance"], ["$3,937,628", "", ""]], page: "p", name: "n", type: "multiRowCard", displayTitle: null };
    writeKpiBlock(ws, 1, 1, card, widths, lay, false, null);`);
  const written = json(`Object.entries(cells).filter(([k, v]) => v.value !== undefined).map(([k, v]) => k + " = " + v.value)`);
  assert.deepStrictEqual(written, ["1,1 = SUM OF SALES", "2,1 = 3937628", "4,1 = Target", "4,2 = (Blank)", "5,1 = Variance", "5,2 = (Blank)"]);
});
