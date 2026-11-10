# Pivot tables and charts (steps)

Menu names are the English interface; give the translated names if the user's interface is in another language, or describe the icon and position.

## Before you start

- One header row, no merged cells, no blank rows, one type per column.
- Select one cell inside the data (Excel) or the whole range (Sheets).

## Pivot table: Excel (Windows / web)

1. Click inside the data → **Insert → PivotTable** → From Table/Range.
2. Choose **New Worksheet** → OK.
3. In the PivotTable Fields panel, drag:
   - the category (e.g. "Category") to **Rows**,
   - the time (e.g. "Month") to **Columns** if you want a cross-table,
   - the amount to **Values** (it shows "Sum of Amount").
   - a field to **Filters** to look at one group at a time.
4. To change Sum to Count or Average: click the field in Values → **Value Field Settings**.
5. Dates: right-click a date in the pivot → **Group** → Months / Years.
6. New data added? Right-click the pivot → **Refresh**. If the data is a Table (Ctrl+T), new rows are included automatically.

## Pivot table: Google Sheets

1. Select the data → **Insert → Pivot table** → New sheet → Create.
2. In the editor on the right, click **Add** next to Rows, Columns, Values, Filters.
3. Values: choose **Summarise by** SUM, COUNTA, AVERAGE.
4. Dates: right-click a date in the pivot → **Create pivot date group** → Month / Year.
5. Sheets pivots update automatically when the source range includes the new rows (use `A1:D` style open ranges).

## Which chart

| Question | Chart |
|---|---|
| How does it change over time? | Line (or column for few periods) |
| Which is bigger? | Bar (long names) or column |
| What share is each part? (≤ 5 parts) | Pie / doughnut, or 100% stacked bar |
| How are two numbers related? | Scatter |
| How are values spread? | Histogram |
| Plan vs actual | Clustered column, or line + column combo |

## Chart: Excel

1. Select the data including headers (or click a pivot) → **Insert → Recommended Charts** or pick a type.
2. Click the chart → the **+** button (Windows) or **Chart Design → Add Chart Element**: Chart Title, Axis Titles, Data Labels.
3. Right-click an axis → **Format Axis** to set the minimum to 0 for bar/column charts.
4. Move it: **Chart Design → Move Chart** → New sheet.

## Chart: Google Sheets

1. Select the data → **Insert → Chart**.
2. In the Chart editor: **Setup** → Chart type; check "Use row 1 as headers".
3. **Customise** → Chart & axis titles, Series (colours, data labels), Legend.
4. Three-dot menu on the chart → **Download** or **Publish** if needed.

## Good chart habits

- A title that says the finding ("Food is 40% of my spending"), not just "Chart 1".
- Bars start at zero. Sort bars from largest to smallest unless order matters (months).
- Few colours; highlight the one thing that matters.
- Label units (AZN, hours, %).
