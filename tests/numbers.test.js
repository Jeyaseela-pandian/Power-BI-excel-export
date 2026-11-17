/* "Numbers: As in the report": cards and table columns in the report's own display units and
   decimals; Exact (the default) and Short unchanged */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const { ctx, run, json } = loadApp();

/* A card's main value as written on the sheet: { value, numFmt, rows } */
function card(type, rows, info, reportNumbers) {
  ctx.CARD = { rows, page: "p", name: "n", type, displayTitle: null };
  ctx.INFO = info;
  return json(`(() => {
    const cells = {};
    const ws = { getCell: (r, c) => (cells[r + "," + c] = cells[r + "," + c] || {}), mergeCellsWithoutStyle() {} };
    const lay = { fonts: new Proxy({}, { get: () => ({}) }), reportColors: null, accent: { cap: "FF000000" }, reportNumbers: ${reportNumbers} };
    visualInfo = () => INFO;
    writeKpiBlock(ws, 1, 1, CARD, { span() {}, want() {} }, lay, false, null);
    const v = cells["2,1"];
    return { value: v.value, numFmt: v.numFmt || null, rows: Object.keys(cells).length };
  })()`);
}

test("a card shows its value in the report's display units: Auto picks them, as Power BI does", () => {
  const auto = card("card", [["Sum of Sales"], ["$3937628"]], { valueUnits: { units: 0, precision: null } }, true);
  assert.strictEqual(auto.value, 3937628, "the cell keeps the exact value");
  assert.strictEqual(auto.numFmt, '[>=9950000]"$"#,##0,,"M";"$"#,##0.0,,"M"', "shown as $3.9M: one decimal below ten of the unit");
  // the Adventure Works cards: 18K, 31K, $110M, $97M, $13M, as Power BI shows them
  for (const [raw, fmt] of [["18490", '[>=9950]#,##0,"K";#,##0.0,"K"'], ["$109810000", '[>=9950000]"$"#,##0,,"M";"$"#,##0.0,,"M"'], ["$12550000", '[>=9950000]"$"#,##0,,"M";"$"#,##0.0,,"M"']]) {
    assert.strictEqual(card("cardVisual", [["Value"], [raw]], { valueUnits: { units: 0, precision: null } }, true).numFmt, fmt, raw);
  }
  const thousands = card("card", [["Orders"], ["52345"]], { valueUnits: { units: 1000, precision: 0 } }, true);
  assert.strictEqual(thousands.numFmt, '#,##0,"K"', "the report's units and decimals: 52K");
  const none = card("card", [["Orders"], ["52345"]], { valueUnits: { units: 1, precision: null } }, true);
  assert.strictEqual(none.numFmt, "#,##0", "units 'None': the full number");
  const pct = card("card", [["Margin"], ["89.2%"]], { valueUnits: { units: 0, precision: null } }, true);
  assert.strictEqual(pct.numFmt, "0.0%", "percentages unchanged");
});

test("with Exact (the default), cards are unchanged", () => {
  const exact = card("card", [["Sum of Sales"], ["$3937628"]], { valueUnits: { units: 0, precision: null } }, false);
  assert.notStrictEqual(exact.numFmt, '"$"#,##0.00,,"M"');
  assert.strictEqual(exact.value, 3937628);
});

test("a table column the report gives units or decimals shows them; other columns unchanged", () => {
  ctx.T = { type: "tableEx", rows: [["Product", "Revenue", "Orders"], ["A", "$2471203", "1200"], ["B", "$812000", "640"]] };
  ctx.I = { columnUnits: { "m:revenue": { units: 1000000, precision: 1 } } };
  assert.deepStrictEqual(json("reportFormats(T, I)"), [null, '"$"#,##0.0,,"M"', null]);
  // decimals only, units Auto: the full number with those decimals
  ctx.I2 = { columnUnits: { "m:orders": { units: 0, precision: 2 } } };
  assert.deepStrictEqual(json("reportFormats(T, I2)")[2], "#,##0.00");
  // nothing set: no change
  assert.strictEqual(run("reportFormats(T, { columnUnits: null })"), null);
});
