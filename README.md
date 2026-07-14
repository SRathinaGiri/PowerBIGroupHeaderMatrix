# Group Header Matrix Custom Visual

An Excel-like matrix visual for Microsoft Power BI with grouped row and column headers, expand/collapse controls, subtotals, sorting, pagination, virtualization, and formatting options for table-style analysis.

## Key Features

- Row and column grouped headers for pivot-table-style layouts.
- Expand all, collapse all, and level-wise row/column expand-collapse controls.
- Compact and repeated-label row header modes.
- Row and column subtotals with separate grand total row and grand total column controls.
- Value sorting and row label sorting.
- Pagination and virtualization modes for larger matrices.
- Grid lines, zebra striping, column striping, color presets, and per-measure color settings.
- Power BI selection, context menu, cross-filtering, tooltips, keyboard navigation, and high contrast support.
- No external service calls or external resource downloads.

## Data Roles

| Data Role | Type | Description |
| --- | --- | --- |
| Rows | Grouping | Row hierarchy fields shown down the left side of the matrix. |
| Columns | Grouping | Column hierarchy fields shown across the top of the matrix. |
| Values | Measure | Numeric or formatted values rendered in the matrix cells. |
| Cell Background Color | Measure | Optional measure-driven cell background color values. |
| Cell Font Color | Measure | Optional measure-driven cell text color values. |

## How to Use

1. Add the visual to a Power BI report.
2. Assign one or more fields to Rows and Columns.
3. Assign one or more measures to Values.
4. Use the toolbar to expand/collapse groups, switch between compact and repeated labels, and choose normal, pagination, or virtualization mode.
5. Use the formatting pane to configure labels, colors, grid lines, zebra striping, subtotals, grand totals, page size, and theme presets.

When the visual has not been configured yet, it displays an in-visual landing page with the same setup sequence and data role guidance.

## Sample Report

The `sample` folder contains `GroupedHeaderMatrixSample.pbix`, an offline Power BI report that demonstrates grouped row headers, grouped column headers, expand/collapse behavior, subtotals, and cross-filtering.

See [sample/README.md](sample/README.md).

## Accessibility

The visual is designed for accessible report consumption:

- Keyboard users can tab to matrix cells and use Enter or Space to select values.
- Expand/collapse controls have accessible labels.
- Selection state is exposed with ARIA attributes.
- Power BI high contrast mode uses host foreground, background, selected, and separator colors.
- The visual declares `supportsKeyboardFocus` and `supportsMultiVisualSelection`.

## Certification and Privacy

This visual is designed for Microsoft Power BI certification:

- `capabilities.json` declares no external privileges.
- The visual does not use `fetch`, `XMLHttpRequest`, WebSocket, or external HTTP requests.
- User and report data are rendered locally inside Power BI.
- The build supports `pbiviz package --certification-audit`.

See [PRIVACY.md](PRIVACY.md) and [SUPPORT.md](SUPPORT.md).

## Development

Install dependencies:

```powershell
npm install
```

Run the visual:

```powershell
npm start
```

Run lint:

```powershell
npm run lint
```

Build a certification-audited package:

```powershell
npm run package
```

## Repository

https://github.com/SRathinaGiri/PowerBIGroupHeaderMatrix

## License

MIT. See [LICENSE](LICENSE).
