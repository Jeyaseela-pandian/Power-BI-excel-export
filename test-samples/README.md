# Test report for the Excel export

Build one small Power BI report from these files, publish it, and export it with the app.
Every visual has a known, expected result, so you can see straight away whether the export
(layout, filters, collapsing and especially **colour matching**) works.

| File | What it is |
| --- | --- |
| `sample-sales.csv` | 576 rows: 24 months (Jan 2024 – Dec 2025) × 4 categories × 3 segments × 2 regions |
| `export-test-theme.json` | Colour theme with 8 very different colours, so a match or mismatch is obvious |
| `measures.dax` | The 5 measures to create (Total Sales, Total Cost, Total Orders, Margin %, Bar Color) |
| `deneb-spec.json` | Chart definition for the Deneb visual |

The theme colours, in order: **1 red `#E41A1C`**, **2 blue `#377EB8`**, **3 green `#4DAF4A`**,
4 purple `#984EA3`, 5 orange `#FF7F00`, 6 brown `#A65628`, 7 pink `#F781BF`, 8 grey `#999999`.

Time needed: about 20 minutes. You need Power BI Desktop and a Pro licence or trial to publish.

---

## 1. Load the data and theme

1. Power BI Desktop → **Get data → Text/CSV** → `sample-sales.csv` → **Transform Data**
   (not Load).
2. In Power Query, click the type icon on the **Month** column header → **Text** → choose
   **Replace current**. Power BI otherwise reads "Jan 2024" as a date, and Deneb and the Excel
   export show a date-time such as `2023-12-31T18:30:00.000Z` instead of `Jan 2024`.
   (*Add new step* would give "1/1/2024" instead – use *Replace current*.)
3. Rename the query from `sample-sales` to **Sales**, then **Close & Apply**.
4. Select the **Month** column → Column tools → **Sort by column → Month Sort**.
   (Otherwise months sort alphabetically: Apr, Aug, Dec…)
5. Create the 5 measures from `measures.dax` (Modeling → New measure, one at a time).
   Format **Margin %** as a percentage.
6. **View → Themes → Browse for themes** → `export-test-theme.json`.

## 2. Page 1 – "Colour test"

Rename the first page to **Colour test** and add these visuals. The title in **bold** is the
visual's title (Format → General → Title); the export uses it to name each table.

| # | Visual type | Title | Fields |
| --- | --- | --- | --- |
| 1 | Card | **Total Sales** | Total Sales |
| 2 | Line chart | **Sales and cost by month** | X-axis: Month · Y-axis: Total Sales, then Total Cost |
| 3 | Clustered column chart | **Sales by segment** | X-axis: Month · Y-axis: Total Sales · Legend: Segment |
| 4 | Clustered bar chart | **Sales by category (custom colour)** | Y-axis: Category · X-axis: Total Sales. Then Format visual → Visual → Bars → Colors → pick **black** (the theme's text colour, top row of the picker) or More colors → **#000000**. Don't add **Bar Color** to this visual or use it for the colour (**fx**): that's the Deneb test, and this one checks a colour picked by hand. |
| 5 | Deneb (see below) | **Deneb sales** | Month, Month Sort, Total Sales, Bar Color |
| 6 | Table | **Category summary** | Category, Total Sales, Total Cost, Margin % |
| 7 | Matrix | **Sales by category and year** | Rows: Category · Columns: Year · Values: Total Sales |
| 8 | Slicer | Region | Region – select **North** |
| 9 | Slicer (Between) | Date | Date – drag to roughly **Jun 2024 – Aug 2025** |
| 10 | Treemap | **Sales by category tiles** | Category: Category · Values: Total Sales |

Place them in a rough grid: cards and slicers along the top, charts side by side below,
tables at the bottom. This also exercises the "Like the report" layout.

**Deneb visual** (visual 5):
1. Visualizations pane → **…** → *Get more visuals* → search **Deneb** → Add (it's a certified visual).
2. Add the fields Month, Month Sort, Total Sales and Bar Color. For **Month Sort**, click its arrow
   in the field well → **Don't summarize**.
3. Click the visual's **…** → *Edit* → choose **Vega-Lite**, empty template → replace the
   specification with the contents of `deneb-spec.json` → *Apply* → close the editor.
4. The bars should be green (`#0B8E37`).

## 3. Page 2 – "Long table"

This page tests long tables: collapsing rows behind **[+]**, the "x of y rows shown" note, and a
short table sitting beside a long one.

**Create the page**

1. Click **+** next to the page tabs at the bottom to add a page.
2. Double-click the new tab (*Page 2*), type **Long table**, press Enter.

**Visual 1 – "All rows" (the long table)**

1. Click an empty spot on the page, then click **Table** in the Visualizations pane.
2. Drag it to fill the **left two-thirds** of the page and most of its height.
3. In the Data pane, tick these fields **in this order**: Date, Month, Category, Segment, Region,
   Sales, Cost. They appear under **Columns** in the Visualizations pane.
4. **Date shows as Year / Quarter / Month / Day?** Power BI added a date hierarchy. Under Columns,
   click the arrow next to **Date** and choose **Date** (not *Date Hierarchy*).
5. Check that **Sales** and **Cost** show as *Sum of Sales* / *Sum of Cost*. That's fine: every
   row is a unique Date × Category × Segment × Region combination, so the sum is the row's own value.
6. Format visual (paint-brush icon) → **General → Title** → On → Text: **All rows**.
7. Scroll the table: it should hold **576 rows** plus a Total row. The Total row is not exported;
   Power BI's export never includes totals.

**Visual 2 – "Category totals" (the short table beside it)**

1. Click an empty spot to the **right** of "All rows", then click **Table**.
2. Place it in the right third, with its **top edge level with the top of "All rows"**.
3. Tick **Category**, then **Total Sales**.
4. Title: **Category totals**. It shows 4 rows.

**Check before moving on**

- Page 1's slicers don't apply here (no Region or Date filter on this page), so "All rows" shows
  all 576 rows. If it shows fewer, open **View → Sync slicers** and make sure the Region and Date
  slicers are **not** synced to *Long table*.

**Expected in Excel** (Layout: *Like the report*, Collapse on, Rows always visible: 8)

| What | Expected |
| --- | --- |
| "All rows" | Starts in column B (column A is the margin); title, caption `Table · 576 rows`, header + first **8** rows visible; the other 568 rows collapsed behind **[+]** in the left margin |
| Note under it | On the closing rule: **`576 rows`** `· 568 grouped under [+] / [–] in the left margin · expand before sorting or filtering` (shorter wording on narrow tables) |
| Click **[+]** | All 576 rows appear; **[–]** collapses them again |
| "Category totals" | To the right of "All rows", tops level; all 4 rows **always visible** (never hidden by the collapse) |
| Colours | Neither table gets chart colours (table visuals keep the plain table look) |
| Collapse off | Re-export with *Collapse extra rows* unticked: all 576 rows visible, no [+], no note |
| Rows always visible = 20 | Re-export: header + 20 rows visible, note says `556 grouped` |

## 4. Page 3 – "Hidden page"

This page tests page selection: hidden pages are skipped by *All visible pages*, but can be
exported when picked on purpose.

**Create the page**

1. Click **+** next to the page tabs to add a page. Rename it **Hidden page**.

**Visual – "Orders by category"**

1. Click **Clustered column chart** in the Visualizations pane and size it to about half the page.
2. Drag **Category** to **X-axis** and **Total Orders** to **Y-axis**.
3. Title: **Orders by category**.
4. The columns are **red** (theme colour 1), since there's one series and no legend.

**Hide the page**

1. Right-click the **Hidden page** tab → **Hide page**. An eye-with-a-slash icon appears on the tab.
2. The page stays visible to you in Desktop; it disappears for viewers once published.
3. Click back on the **Colour test** tab before publishing, so the report opens on page 1.

**Expected in the app and in Excel**

| Test | Expected |
| --- | --- |
| Export options → **Choose pages** | Lists *Colour test*, *Long table*, and *Hidden page* marked **hidden page** |
| Export options → **All visible pages** | 2 sheets: *Colour test*, *Long table*. No *Hidden page* sheet |
| **Choose pages** → tick *Hidden page* only | 1 sheet *Hidden page* with "Orders by category": caption key `■ Total Orders` in **red**, and the Total Orders header and values filled red (blue if the theme couldn't be read) |
| After any multi-page export | The report in the app is back on the page you were viewing |

## 5. Publish

**File → Publish → My workspace** (or a test workspace). Save the file as
`Export colour test.pbix` when asked.

---

## 6. Run the tests in the app

1. Open http://localhost:5500, sign in, open **Export colour test**.
2. Click the **▾** next to Export → **Colours**. If it shows **Allow exact colours**, click it and
   accept the permission (Dataset.Read.All). Without it the export still works, but colours set by
   hand on a visual (test 4, the black bar) can't be read and come out in a theme colour.
   The report file is read only for colours, and only with Contributor (or higher) access to the
   workspace; nothing is stored.
3. Gear → **Diagnostics** → **Run checks** → open the **Activity log**. Expect the lines:
   `Colour theme: read from the report – #E41A1C, #377EB8, #4DAF4A, …` and
   `Exact colours: available – PBIR report file, 1 visual(s) with colours set by hand`
   (`classic report file` for reports saved in the older format).
   If the theme line says *not readable, using Power BI defaults* and exact colours aren't
   available, the colour tests below show Power BI's default colours (blue `#118DFF`, navy
   `#12239E`, orange `#E66C37`…) instead of red/blue/green.
4. Export **Colour test** (current page, Layout: Like the report, both colour options on).
5. Export again with **All visible pages**, then with **Choose pages** including the hidden page.

### Expected results – page "Colour test"

| Visual | Expected in Excel |
| --- | --- |
| Masthead | Thin accent rule across the top, `EXPORT COLOUR TEST` eyebrow, page title `Colour test` (Calibri Light 24), meta line with export time and `2 filters applied`, closed by a thin ink rule |
| Filter bar | Pale grey bar: `FILTERS  Region North │ Date 6/1/2024 – 8/31/2025` (untouched slicers show `All` in grey) |
| Total Sales (card) | A scorecard tile: accent cap on top, `TOTAL SALES` label, large value `2,517,775.17`; no "Card" title. Cards on one row join into one strip |
| All tables | Same frame everywhere: pale header with an ink rule above it, hairlines between rows, a closing ink rule; numbers with thousands separators (e.g. `93,687.30`); Month Sort stays `202401`. Only chart values are coloured, never text columns |
| Sales and cost by month | Caption key **■ red** Total Sales, **■ blue** Total Cost. The *Total Sales* header and every value under it are filled **red** with white text; *Total Cost* **blue** |
| Sales by segment | Caption key **■ red** Consumer, **■ blue** Corporate, **■ green** Home Office. Each *Total Sales* value is filled in its segment's colour, and a colour tab fills the cell left of each row. Sort the table: the fills move with their rows and the tabs re-colour to match |
| Sales by category (custom colour) | With exact colours: the *Total Sales* header and values filled **black** (`#000000`) with white text, as in the report. Without the permission: **red** (theme colour 1) and a note at the bottom: *Chart colours follow the report theme…* |
| Deneb sales | Title built from the data (e.g. *Total Sales by Month*, as "Deneb" says nothing); caption `■ Total Sales · row colours from Bar Color` in **green**; the *Total Sales* header and values filled green. No row tabs: every row has the same colour |
| Sales by category tiles (treemap) | Each *Total Sales* value filled in its tile's colour, with a matching colour tab beside the row: **Apparel red, Electronics blue, Furniture green, Office Supplies purple** (in the order the categories appear in the export) |
| Category summary (table) | Caption `Table · 4 rows`, no colour keys |
| Sales by category and year (matrix) | Caption `Matrix · 4 rows`, flat rows (matrix layout isn't rebuilt) |
| Footer and print | Footer line `Source: Power BI · Export colour test`; File → Print shows landscape, one page wide, page numbers in the footer |

### Expected results – other pages and options

- **Long table** page: see the expected-results table in section 3.
- **Hidden page** and page selection: see the expected-results table in section 4.

| Test | Expected |
| --- | --- |
| *Match the report's chart colours* off | No colours at all: no keys, fills or tabs; captions read `Line chart · 24 rows` etc. |
| *Fill value cells with the chart colour* off | Keys, coloured series headers and row tabs only; value cells stay white |
| Branding → Table style *Classic blue* | Tables use Excel's blue table style (blue header, stripes) with no ink rules; series headers, value fills and tabs still in the report colours |
| Branding → a very pale accent (e.g. `#FFF200`) | The top rule and KPI caps turn ink-dark so they stay visible; the sheet tab keeps the pale accent |
| Export options → *Number visuals* | Titles read `01  Sales and cost by month`, `02  …` |
| Open in desktop Excel | No repair prompt; every table's filter and sort buttons work |

If a result differs, note the visual and what you saw, and copy the Diagnostics table and
Activity log into the bug report.
