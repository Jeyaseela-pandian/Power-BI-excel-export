# NeenOpal Export for Power BI

A browser app that exports Power BI report pages to a clean, branded Excel workbook. You sign in with your work account and open a report from your workspaces. When you choose Export, each report page becomes a sheet laid out like the report, with the report's colours, number formats, filters and drill level.

## What it does

- **Whole pages, one file.** Each report page becomes a sheet. Visuals keep their place on the page, or can be stacked one under another.
- **Two ways to show visuals:**
  - **Data tables:** each visual's data as an Excel table.
  - **Charts:** bar, column, line, area, combo, pie, donut, scatter, waterfall, funnel and treemap visuals become real Excel charts. So do common AppSource charts: radar, sunburst, histogram, tornado and box & whisker. Every chart carries alt text for screen readers.
- **Report formatting carried over.** Colour scales, rules, data bars, icons, theme and hand-set colours, axis titles, live totals rows and sparklines.
- **Cards and KPIs as tiles.** KPI tiles show their goal, status and trend.
- **Filters on record.** Slicers, active filters and what you clicked in the report (a bar, a slice) appear in each sheet's filter bar.
- **Excel slicers (optional).** The page's slicer fields become real Excel slicers on the tables that have them.
- **Drill level kept.** Each visual is exported at the level shown in the report, or in the app at the moment of export.
- **AppSource visuals.** Common marketplace visuals have their own rules; any other visual still exports as a table.
- **Branding.** Company name, logo and colours are set once in the app and applied to every export.

## Requirements

- A Power BI account with access to the reports you want to export.
- A Microsoft Entra **app registration**:
  - Platform: Single-page application.
  - Redirect URI: `http://localhost:5500`.
  - Delegated Power BI permissions: `Report.Read.All` and `Workspace.Read.All`.
  - Optional: `Dataset.Read.All`. With it, the app reads the report's design file for exact hand-set colours. It may need admin consent.
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

`serve.py` is a small static server that turns off browser caching, so the browser always loads the latest files.

## Using it

1. Pick a workspace and a report. The report opens in the app.
2. Choose **Export to Excel**, then set the options:
   - **Pages:** the current page, all visible pages, or a chosen set. Hidden pages are listed but unticked.
   - **Layout:** like the report, or stacked.
   - **Visuals as:** data tables or charts.
   - **Numbers in tables:** exact (2,471,203) or short (2.47M). Short numbers still hold the exact value.
   - **Matrices:** a flat table, or Power BI's grouped layout with subtotals.
   - **Long tables:** fold extra rows behind [+], or keep every row visible.
   - **Colours:** match the report's chart colours.
   - **Include:** filters and slicer selections, Excel slicers, export notes, numbered visuals.
3. Choose **Export**. The workbook downloads when it's ready.

## Limitations

- Maps, ribbon charts and most AppSource charts export as tables, because Excel has no matching chart.
- Excel arranges treemap tiles and sunburst segments itself: the same data, sizes and colours, in Excel's own layout.
- An Excel slicer filters one table. When several tables share a field, each gets its own slicer.
- Excel charts show Excel's own hover (series, category and value), not Power BI tooltips.
- Totals of percentages, averages and distinct counts are left blank, because they can't be added up from the rows.
- Key influencers, R, Python, paginated report and Power Apps visuals give no data to export. The sheet's notes name them.
- Table sparklines need Build permission on the dataset.

## Project files

| File | Purpose |
|---|---|
| `index.html` | The app page and export options |
| `styles.css` | App styles |
| `config.js` | App registration settings and edition |
| `app.js` | Sign-in, workspaces and the embedded report |
| `ui.js` | Export dialog, page and visual selection, branding |
| `colors.js` | Reads the report's theme, colours and visual settings |
| `export.js` | Builds the Excel workbook: tables, tiles, charts, filters and formatting |
| `serve.py` | Local server on port 5500 |

Libraries are loaded from public CDNs: MSAL for sign-in, the Power BI JavaScript client, ExcelJS and JSZip. There is no build step.
