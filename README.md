# Group Header Matrix Custom Visual

An Excel-like matrix visual for Microsoft Power BI: grouped row and column headers, expand/collapse at every level, subtotals, a table layout for measures-only reports, whole-row and per-column conditional formatting, dynamic format strings and fast rendering for large matrices. Built for MIS and P&L-style reports.

## Key Features

**Layout and navigation**
- Grouped row and column headers for pivot-table layouts, or a plain table when only Rows and Values are used.
- Expand/collapse per group, Expand All / Collapse All, and one-level-at-a-time row and column expand/collapse. Groups with a single member can collapse too.
- Compact (indented) or tabular row headers, with optional repeated labels.
- Row subtotals (top or bottom) and column subtotals, with separate grand-total row and column controls.
- Sorting by values or by row labels: numbers, dates and formatted numbers sort numerically, blanks last, and clearing a sort restores the original order.
- Normal, pagination and virtualization modes.
- Optional toolbar (can be hidden).

**Formatting**
- Horizontal and vertical alignment for row headers, column headers and values.
- Per-measure column settings: header font, size, bold/italic/underline, header text and background colors, a separator band after the column, and which conditional-color measure colors its cells.
- Whole-row conditional formatting (background, font color, bold) from measures, applied to the row header too if you want, with a choice of whether row or cell formatting wins where both apply.
- Cell-level conditional formatting (background and font color) from measures.
- Grid lines, cell padding, header and row heights, zebra and column striping, color presets (themes).
- Dynamic format strings: each value uses its own format (for example amounts as `#,##0.0` and ratio rows as `0.0%` in the same column). Measures that return formatted text are shown as returned.
- Header labels use each field's format string (dates, numbers) and show `(Blank)` for blank members.

**Interaction**
- Clicking a value filters other visuals by its row (leaf or group/subtotal row) and its column group. Ctrl+click adds to the selection; non-selected values are dimmed.
- Other visuals and slicers filter the matrix as usual.
- Power BI context menu, tooltips, keyboard navigation and high contrast support.

**Performance**
- Renders only when data, formatting or view state actually change: resizes and the visual's own saved state don't rebuild the table.
- Selection updates highlighting in place; one set of event listeners per table instead of per cell; shared cell styles via CSS variables; cached value formatters.

**Privacy**
- No external service calls or external resource downloads.

## Data Roles

| Data Role | Type | Description |
| --- | --- | --- |
| Rows | Grouping | Row hierarchy shown down the left side. |
| Columns | Grouping | Column hierarchy shown across the top. Leave empty for a table layout. |
| Values | Measure | Values shown in the cells. |
| Cell Background Color | Measure | Optional measures returning colors (for example `"#FFF2CC"`) for cell backgrounds. |
| Cell Font Color | Measure | Optional measures returning colors for cell text. |
| Row Background Color | Measure | Optional measure returning a color for the whole row. |
| Row Font Color | Measure | Optional measure returning a text color for the whole row. |
| Row Bold | Measure | Optional measure returning `1`, `TRUE` or `"bold"` to make the whole row bold. |

The color and row-style measures are never shown as columns. Like any field in a visual, they are included in **Show as a table** and **Export data**.

## How to Use

1. Add the visual to a report.
2. Put fields on Rows (and optionally Columns), and measures on Values.
3. Optionally add color measures to the Cell and Row formatting wells.
4. Use the toolbar to expand/collapse, switch compact/tabular layout and choose a rendering mode.
5. Use the Format pane (below) for alignment, per-column styles, row formatting, grid and totals.

When the visual has no data yet, it shows a landing page with these steps.

## Conditional Formatting

### Whole rows

Add measures to **Row Background Color**, **Row Font Color** and/or **Row Bold**. Each row reads them from its grand-total column (pure row context), so a measure only needs to look at the row. For example:

```DAX
Row Fill =
IF ( SELECTEDVALUE ( Lines[Line] ) IN { "Total Revenue", "Contribution", "Operating Profit" }, "#D9E1F2" )
```

Format pane → **Row formatting**:
- **Apply to row headers**: color the row header cells too (default on).
- **When row and cell formatting overlap**: *Cell formatting wins* (default) or *Row formatting wins*.

### Individual cells

Add measures returning colors to **Cell Background Color** / **Cell Font Color**. By default the Nth color measure applies to the Nth value (extra values use the last one). To choose explicitly, open Format pane → **Measure columns** → *measure* and set **Cell font color from** / **Cell background from** to *None* or a specific color measure. A typical MIS setup colors only the comparison columns:

| Column | Cell font color from |
| --- | --- |
| Actual, Plan, Last Year | None |
| vs Plan % | a "better/worse than plan" measure |
| vs LY % | a "better/worse than last year" measure |

Order applied: base style → zebra stripes → row style and cell style (in the chosen precedence).

## Format Pane

| Card | Settings |
| --- | --- |
| Behavior | Show toolbar, sticky headers, page size, show diagnostics |
| Labels | Fonts, sizes and bold for row headers, column headers and values |
| Alignment | Horizontal and vertical alignment for row headers, column headers and values |
| Measure columns | Per measure: header font and colors, separator after the column (width, color), cell font/background color source |
| Row formatting | Apply to row headers, row vs cell precedence |
| Colors | Row header, column header and cell colors |
| Grid | Horizontal/vertical lines, thickness, color, cell padding, header row height, row height |
| Zebra striping / Column striping | Alternating row and column colors |
| Subtotals / Grand Total | Row and column subtotals, subtotal position, grand-total row/column and position |
| Theme | Presets: Default, Light Blue, Dark, Excel-like, Modern |

Vertical alignment is visible when a row is taller than its text, for example with a **Header row height** or **Row height** set in the Grid card.

## Rendering Modes

- **Normal** renders all prepared rows in one table.
- **Pagination** renders one page of rows at a time using the configured page size.
- **Virtualization** renders only the visible scroll window plus a buffer, measuring the real row height. It reduces DOM work for large row counts but does not reduce the data loaded.

## Troubleshooting

Format pane → **Behavior** → **Show diagnostics** displays the data layout Power BI sends to the visual: measures and their roles, column and row levels, how each column maps to values and color measures, and the raw values (with per-value format strings) of the rendered rows. Turn it off when done.

## Sample Report and Test Data

The `sample` folder contains the sample report and scripts that build a synthetic hotel P&L test model and test pages (GL-code drill-down, properties and regions, dynamic formats, row and cell formatting, slicers, cross-filter charts). See [sample/README.md](sample/README.md).

## Accessibility

- Keyboard users can tab to matrix cells and use Enter or Space to select values.
- Expand/collapse controls have accessible labels.
- Selection state is exposed with ARIA attributes.
- Power BI high contrast mode uses host foreground, background, selected and separator colors.
- The visual declares `supportsKeyboardFocus` and `supportsMultiVisualSelection`.

## Certification and Privacy

- `capabilities.json` declares no external privileges.
- The visual does not use `fetch`, `XMLHttpRequest`, WebSocket or external HTTP requests.
- User and report data are rendered locally inside Power BI.
- The build runs `pbiviz package --certification-audit`.

See [PRIVACY.md](PRIVACY.md) and [SUPPORT.md](SUPPORT.md).

## Development

```powershell
npm install
```

```powershell
npm start
```

```powershell
npm run lint
```

```powershell
npm run package
```

`npm run package` writes `dist/groupHeaderMatrix19740818.<version>.pbiviz`. Bump `version` in `pbiviz.json` for each build you import, so Power BI Desktop does not reuse a cached copy of the visual.

## Repository

https://github.com/SRathinaGiri/PowerBIGroupHeaderMatrix

## License

MIT. See [LICENSE](LICENSE).
