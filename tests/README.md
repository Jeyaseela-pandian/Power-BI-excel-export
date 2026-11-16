# Checks

Checks for the export's logic: every chart type, colours, axis titles from themes, KPI tiles,
key influencers, decomposition trees, small multiples, and real report files. They run in Node
(version 18 or later) without a browser or sign-in, and change nothing in the app.

```bash
node --test tests/*.test.js
```

Run them after every change. A failing check names what stopped working.

## Real reports
`real-reports.test.js` reads the `.pbix` files in your Downloads folder and checks every chart's
axis titles against the report's own themes. To use another folder:

```bash
REPORTS_DIR="C:/path/to/reports" node --test tests/*.test.js
```

Without any `.pbix` files those checks are skipped; the rest still run.

## Files
| File | Checks |
|---|---|
| `harness.js` | Loads the app's scripts for the checks; reads `.pbix` files |
| `charts.test.js` | Every chart type, the chart safety net, empty legends, number axes, second axis, ribbon charts |
| `colors.test.js` | Hand-set and rule colours, legend order, fallback theme, chart background |
| `themes.test.js` | Axis titles from base themes, Fluent2, chart settings, Viewers' website layout |
| `cards.test.js` | KPI tiles' "(Blank)" values |
| `ai-visuals.test.js` | Key influencers, decomposition tree, small multiples, notes for skipped visuals |
| `real-reports.test.js` | Axis titles on real reports; the AI sample's settings |
