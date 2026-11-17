# NeenOpal Export for Power BI

A browser app that exports Power BI report pages to a clean, branded Excel workbook. You sign in with your work account and open a report from your workspaces. When you choose Export, each report page becomes a sheet laid out like the report, with the report's colours, number formats, filters and drill level.

## What it does

- **Whole pages, one file.** Each report page becomes a sheet. Visuals keep their place on the page, or can be stacked one under another.
- **Two ways to show visuals:**
  - **Data tables:** each visual's data as an Excel table.
  - **Charts:** bar, column, line, area, combo, pie, donut, scatter, waterfall, funnel and treemap visuals become real Excel charts. So do common AppSource charts: radar, sunburst, histogram, tornado and box & whisker. Every chart carries alt text for screen readers.
- **Charts that look like the report.**
  - Series and legend colours in Power BI's order, the visual's background (or white), titles and axis titles as Power BI shows them.
  - Line values on a second axis, and whole-number axes (days, periods, years) spaced by value.
  - Data labels in Power BI's display units and decimals (1,037 · 6.8K · 33K · $11M), placed past the bar's end with room before the axis ends, and total labels on stacked charts.
  - Long legends get the room they need. Pie and donut labels appear only where they fit.
- **Numbers as in the report.** Cards, KPIs and table columns use the report's display units and decimals ($3.94M on a card). Every cell keeps its exact value for sums and formulas.
- **Report formatting carried over.** Colour scales, rules, data bars, icons, theme and hand-set colours, live totals rows and sparklines.
- **Cards and KPIs as tiles.** KPI tiles show their goal, status and trend. A one-row multi-row card shows its fields side by side, and an empty value reads "(Blank)".
- **AI and special visuals:**
  - Key influencers become a table of the exact shares behind them ("Status is Won: 91.4% for Customer Follow-up, 64.4% for the others").
  - A decomposition tree becomes an indented table of the branches opened in the report.
  - Small multiples become one grouped Excel object: a chart per panel in a grid, on a shared axis.
- **Filters on record.** Slicers, active filters and what you clicked in the report (a bar, a slice) appear in each sheet's filter bar.
- **Drill level kept.** Each visual is exported at the level shown in the report, or in the app at the moment of export.
- **AppSource visuals.** Common marketplace visuals have their own rules; any other visual still exports as a table.
- **Every role.** Admins, Members, Contributors and Viewers all get the report's colours and settings (see [How roles are handled](#how-roles-are-handled)).
- **Branding.** Company name, logo and colours are set once in the app and applied to every export.

## Requirements

- A Power BI account that can open the reports you want to export (normally Power BI Pro or Premium Per User; a free licence is enough for Viewers of a workspace on Premium or Fabric F64 or larger capacity).
- A Microsoft Entra **app registration**:
  - Platform: Single-page application.
  - Redirect URI: `http://localhost:5500`.
  - Delegated Power BI permissions: `Report.Read.All` and `Workspace.Read.All`.
  - Optional: `Dataset.Read.All`. With it, the app reads the report's design file for exact hand-set colours, and can query the semantic model where the user has Build permission. It may need admin consent.
- Python 3 (to serve the app locally), and a current version of Edge or Chrome.
- Excel 2016 or later to open the exports. Treemap, sunburst, waterfall, funnel and box & whisker charts use Excel's newer chart types. Older Excel shows a stand-in: a classic chart for waterfalls and funnels, a note for the rest.

## Setup

1. Clone the repository.
2. Open `config.js` and fill in:
   - `clientId`: the app registration's Application (client) ID.
   - `tenantId`: your organization's tenant ID, or `"organizations"` to allow any work account.
   - `redirectUri`: must match the redirect URI on the app registration.
   - `edition`: `"trial"` adds a "Created with" footer to every sheet; `"pro"` removes it.
3. Start the app:

   ```bash
   python serve.py
   ```

4. Open http://localhost:5500 and sign in.

Always start the app with `serve.py`, not another static server (such as VS Code's Live Server). `serve.py` turns off browser caching, so the browser always loads the latest files, and passes the app's requests for a report's layout on to Power BI, which browsers won't send directly. Without it, Viewers' exports fall back to the report theme's colours. If you host the app elsewhere, that host needs the same pass-through.

## Using it

1. Pick a workspace and a report. The report opens in the app.
2. Choose **Export to Excel**, then set the options:
   - **Pages:** the current page, all visible pages, or a chosen set. Hidden pages are listed but unticked.
   - **Visuals on this page:** leave out any visual you don't want.
   - **Layout:** like the report, or stacked.
   - **Visuals as:** data tables or charts.
   - **Chart background:** white (clean and printable), or each visual's own background as in the report.
   - **Long tables:** fold extra rows behind [+], or keep every row visible.
   - **Colours:** match the report's chart colours.
   - **Include:** filters and slicer selections, export notes (skipped or changed visuals), numbered visuals.
3. Choose **Export**. The workbook downloads when it's ready.

## How roles are handled

The app has one export for every role. Only the way it reads a report's design (colours, layout and settings) depends on what the user may access, and the app finds out by trying:

1. **The report file**, for Admins, Members and Contributors with `Dataset.Read.All`.
2. **The layout Power BI's website uses to draw the report**, for everyone else, Viewers included. This goes through `serve.py`.
3. **The report theme's colours only**, when neither can be read.

Both routes give the export the same settings, so Viewers' exports match. The few differences:

- Settings that only a custom theme file sets (for example axis titles turned on by a custom theme) need the theme file. Base-theme settings work for every role.
- Table sparklines need Build permission on the dataset.
- The Viewer route relies on Power BI's website layout, which Microsoft doesn't document. If it changes, Viewers' exports fall back to the theme's colours, with a note.

The support panel has **Check viewer access** and **Check visuals**, which show which route the app used and where each visual's settings came from.

## Limitations

- Maps and most AppSource charts export as tables, because Excel has no matching chart. Ribbon charts export as tables too.
- Excel arranges treemap tiles and sunburst segments itself: the same data, sizes and colours, in Excel's own layout.
- Excel charts show Excel's own hover (series, category and value), not Power BI tooltips, and have no legend scrolling or play axis.
- Key influencers show exact shares; Power BI's "× more likely" figures come from its own model and aren't exported.
- Anomaly bands, smart narratives and Q&A visuals are worked out by Power BI on screen and aren't exported. The sheet's notes name them.
- R, Python, paginated report and Power Apps visuals give no data to export. The sheet's notes name them.
- Totals of percentages, averages and distinct counts are left blank, because they can't be added up from the rows.
- Table sparklines need Build permission on the dataset.

## Checks

The export's logic has automatic checks: every chart type, colours, axis titles from themes, cards, AI visuals, and the `.pbix` files in your Downloads folder. They run in Node 18 or later, without a browser or sign-in:

```bash
node --test tests/*.test.js
```

Run them after every change. See [tests/README.md](tests/README.md) for details.

## Project files

| File | Purpose |
|---|---|
| `index.html` | The app page and export options |
| `styles.css` | App styles |
| `config.js` | App registration settings and edition |
| `app.js` | Sign-in, workspaces and the embedded report |
| `ui.js` | Export dialog, page and visual selection, branding |
| `colors.js` | Reads the report's design: theme, colours and visual settings, from the report file or the website layout |
| `export.js` | Builds the Excel workbook: tables, tiles, charts, filters and formatting |
| `viewer-check.js` | The support panel's checks: which route the app used and each visual's settings |
| `serve.py` | Local server on port 5500, with the pass-through to Power BI |
| `tests/` | Automatic checks (`node --test tests/*.test.js`) |

Libraries are loaded from public CDNs: MSAL for sign-in, the Power BI JavaScript client, ExcelJS and JSZip. There is no build step.
