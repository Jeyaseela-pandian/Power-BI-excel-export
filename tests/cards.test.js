/* KPI tiles: an empty value reads "(Blank)", as in Power BI */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { run, json } = loadApp();

test("a card's empty values show (Blank); values that exist are unchanged", () => {
  run(`cells = {}; ws = { getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}), mergeCellsWithoutStyle() {} };
    widths = { span() {}, want() {} }; lay = { fonts: new Proxy({}, { get: () => ({}) }), reportColors: null, accent: { cap: "FF000000" } };
    card = { rows: [["Sum of Sales", "Target", "Variance"], ["$3,937,628", "", ""]], page: "p", name: "n", type: "multiRowCard", displayTitle: null };
    writeKpiBlock(ws, 1, 1, card, widths, lay, false, null);`);
  const written = json(`Object.entries(cells).filter(([k, v]) => v.value !== undefined).map(([k, v]) => k + " = " + v.value)`);
  assert.deepStrictEqual(written, ["1,1 = SUM OF SALES", "2,1 = 3937628", "4,1 = Target", "4,2 = (Blank)", "5,1 = Variance", "5,2 = (Blank)"]);
});
