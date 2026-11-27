# Group Header Matrix for Power BI

**An Excel-style matrix for Power BI reports** — built for people who prepare MIS, P&L and management reports and want the formatting control they are used to in Excel: highlighted total rows, green/red variance columns, column separators, aligned headers, and groups you can open and close on both rows and columns.

## Why this visual?

The native Power BI matrix is powerful, but some everyday Excel formatting is hard or impossible to reproduce with it. Group Header Matrix adds:

| You want to… | In Group Header Matrix |
| --- | --- |
| Open and close individual **column groups** (for example one region or one year), not only whole levels | Every column group header has its own **+ / −** button |
| Highlight **whole rows** such as Total Revenue, Gross Margin or EBITDA | Add one measure to **Row Background Color** (and **Row Bold**) — the full row, including its label, is formatted |
| Show **green / red** only in the variance columns (vs Budget %, vs Last Year %) while Actual and Budget stay plain | Choose, per column, which color measure colors it — or none |
| Separate the month block from the YTD block | A colored **separator band** after any column, no dummy "|" measure needed |
| Style each column header differently | Font, size, bold/italic/underline, text and background color **per measure** |
| Align headers and numbers top, middle or bottom | **Horizontal and vertical alignment** for row headers, column headers and values |
| Switch between an indented (compact) and an Excel tabular layout while reading the report | **Compact** and **Repeat Labels** right on the visual's toolbar |
| Show amounts and percentages in the same column | Each value keeps its own **dynamic format** (`#,##0.0` on amount rows, `0.0%` on ratio rows) |
| Use it as a plain table | Leave **Columns** empty: one column per measure, rows still expandable |

## Features at a glance

### Reading the report
- **Expand / collapse anywhere** — each row group and each column group has its own + / − button, including groups with a single member.
- **Toolbar** — expand all, collapse all, open or close one row level or one column level at a time, compact or tabular layout, repeat labels, and paging. The toolbar can be hidden.
- **Sorting** — click a column header to sort by its values, or the top-left header to sort by row labels. Numbers, dates and formatted numbers sort numerically; blanks go last; a third click restores the original order.
- **Filtering both ways** — click any value to filter the other visuals on the page by its row (an account, a group or a subtotal) and its column (for example a region or a month). Ctrl+click to add more. Slicers and other visuals filter the matrix as usual.
- **Tooltips and right-click menu** like any Power BI visual.
- **Large reports** — renders quickly, and offers **Pagination** or **Virtualization** for very long tables.
- **Keyboard and high-contrast** support.

### Formatting
- **Whole-row formatting** from measures: background, text color and bold — with a choice of whether row or cell formatting wins where both apply, and whether the row label is colored too.
- **Cell formatting** from measures: background and text color, assigned per column.
- **Per column (measure) settings**: header font and colors, separator after the column, and which color measure colors its cells.
- **Alignment**: horizontal and vertical, separately for row headers, column headers and values.
- **Grid**: horizontal/vertical lines, thickness, color, cell padding, header row height and row height.
- **Zebra striping** for rows and columns, **color presets** (Default, Light Blue, Dark, Excel-like, Modern).
- **Totals**: row subtotals at the top or bottom, column subtotals, grand total row (top or bottom) and grand total column, each on or off.
- **Numbers and labels**: each value uses its measure's format (including dynamic format strings); measures that return formatted text are shown as they are; dates and numbers in row and column headers use their field's format, and blank members show as `(Blank)`.

## Getting started

1. Import the visual (`.pbiviz`) into your report and add it to a page.
2. Drag fields into the field wells:

   | Field well | Put here | Example |
   | --- | --- | --- |
   | **Rows** | one or more row fields (a hierarchy works too) | Line → Detail → GL Code |
   | **Columns** | optional column fields; leave empty for a table | Region → Property, or Year → Month |
   | **Values** | your measures | Actual, Budget, Last Year, vs Budget %, YTD Actual … |
   | **Row Background Color / Row Font Color / Row Bold** | optional measures that style whole rows | a measure returning `"#D9E1F2"` for total rows |
   | **Cell Background Color / Cell Font Color** | optional measures that color cells | a measure returning `"#1E7B34"` (green) or `"#C00000"` (red) |

3. Open the **Format** pane to fine-tune (see below).

The color measures are never shown as columns. Note that, like every field on a visual, they appear in **Show as a table** and **Export data**.

## How to…

### Highlight total rows (blue band, bold)

Create a measure that returns a color only for the rows to highlight:

```DAX
Row Highlight =
IF (
    SELECTEDVALUE ( 'P&L'[Line] ) IN { "Total Revenue", "Gross Margin", "EBITDA" },
    "#D9E1F2"
)
```

```DAX
Row Highlight Bold =
IF ( SELECTEDVALUE ( 'P&L'[Line] ) IN { "Total Revenue", "Gross Margin", "EBITDA" }, 1 )
```

Put them in **Row Background Color** and **Row Bold**. In **Format → Row formatting**:
- **Apply to row headers** — also color the row label (on by default).
- **When row and cell formatting overlap** — *Cell formatting wins* (default) keeps, for example, green/red variance cells visible on a blue total row; *Row formatting wins* paints the whole row.

### Green / red variance columns only

Create one measure per comparison, for example "better or worse than budget":

```DAX
Font vs Budget =
VAR A = [Actual]
VAR B = [Budget]
RETURN IF ( NOT ISBLANK ( A ) && NOT ISBLANK ( B ), IF ( A >= B, "#1E7B34", "#C00000" ) )
```

(Store costs as negative numbers so "Actual ≥ Budget" means better for revenue, profit and costs alike.)

Add the measures to **Cell Font Color**. Then in **Format → Measure columns**, open each measure and set **Cell font color from**:
- *None* for Actual, Budget, Last Year,
- *Font vs Budget* for vs Budget %, *Font vs LY* for vs LY %, and so on.

Left on *Automatic*, color measures pair with values by position (1st with 1st, …).

### Separate the month block from the YTD block

**Format → Measure columns →** (the last month column, for example *vs LY %*) **→ Separator after column**: width 6–10 px, color white or your page background. Works in every column group.

### Make headers line up nicely

- **Format → Alignment** — horizontal and vertical alignment for row headers, column headers and values. Group labels set to *Left* sit at the start of their span, like merged cells in Excel.
- Vertical alignment shows when a row is taller than its text: set **Format → Grid → Header row height** or **Row height**.
- **Format → Grid** also has cell padding.

### Use it as a simple table

Leave **Columns** empty. Each measure becomes a column; row fields still group and expand.

## Format pane reference

| Card | What you can set |
| --- | --- |
| **Behavior** | Show toolbar · sticky headers · page size · show diagnostics |
| **Labels** | Font, size and bold for row headers, column headers and values |
| **Alignment** | Horizontal and vertical alignment for row headers, column headers and values |
| **Measure columns** | Per measure: header font and colors · separator after the column · cell font / background color source |
| **Row formatting** | Apply to row headers · row vs cell precedence |
| **Colors** | Row header, column header and cell colors |
| **Grid** | Lines, thickness, color · cell padding · header row height · row height |
| **Zebra striping / Column striping** | Alternating row / column colors |
| **Subtotals / Grand Total** | Row and column subtotals · row subtotal position · grand total row and column · grand total row position |
| **Theme** | Color presets |

## Tips and troubleshooting

- **Nothing turns red/green?** Check the measure returns a color text such as `"#C00000"` for the cells you expect, and that the column's **Cell font color from** points to it.
- **Values show as 0.65 instead of 65%?** The value's own format is used when Power BI provides one; otherwise the measure's format string applies.
- **Something looks wrong?** Turn on **Format → Behavior → Show diagnostics** to see exactly what data the visual received (useful when reporting an issue), then turn it off again.

## Sample report

The [sample](sample) folder has a Superstore demo (grouped headers and two-way filtering with charts) and scripts that build a synthetic hotel P&L model with GL-code drill-down, budgets and the formatting shown above. See [sample/README.md](sample/README.md).

## Privacy

The visual works entirely inside Power BI: no external services, no downloads, no data leaves the report. See [PRIVACY.md](PRIVACY.md). Support: [SUPPORT.md](SUPPORT.md).

## For developers

```powershell
npm install
```

```powershell
npm run eslint
```

```powershell
npm run package
```

`npm run package` builds `dist/groupHeaderMatrix19740818.<version>.pbiviz` with the certification audit. Bump `version` in `pbiviz.json` for each build you import so Power BI Desktop does not reuse a cached copy.

## License

MIT — see [LICENSE](LICENSE). Repository: https://github.com/SRathinaGiri/PowerBIGroupHeaderMatrix
