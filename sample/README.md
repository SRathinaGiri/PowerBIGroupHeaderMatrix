# Group Header Matrix Sample

This folder contains `GroupedHeaderMatrixSample.pbix`, an offline Power BI report that showcases the Group Header Matrix visual, plus scripts that add a synthetic P&L test model and test pages to the report.

## Superstore demo (Demo page)

- Grouped row headers (Category > Sub-Category) and column headers (Region > State > City)
- Expand/collapse, subtotals and grand totals
- Cross-filtering in both directions with two small native charts (Total Sales by Region, Total Sales by Category): clicking a matrix value filters the charts by its row and column; clicking a bar or slice filters the matrix

## Hotel P&L test model (synthetic data)

All names and numbers are made up. The model has:

- **PL Lines**: a `P&L` hierarchy Line > Detail > GL Code (48 GL accounts) with calculated lines (Total Revenue, Contribution, Operating Profit) and ratio lines (margins, staff cost share)
- **PL Property**: 12 properties in 3 regions, with a `Location` hierarchy Region > City > Property
- **PL Calendar**: months Jan 2024 to Dec 2026 with a `Period` hierarchy; Actual runs to Aug 2026, Plan to Dec 2026
- **Measures**: Actual, Plan, Last Year, vs Plan %, vs LY % and their YTD versions, Full Year Plan (values use a dynamic format string: `0.0%` on ratio rows), an `Actual (text)` version, and style measures for the formatting wells (row fill and bold, Plan fill, green/red better-or-worse fonts for each comparison)

Test pages:

- **P&L Test**: Year, Month, Region and Property slicers; the matrix with row formatting, Plan fill, green/red comparison columns and column separators
- **P&L by Property**: Region > Property on columns for a wide matrix

### Building it into a report

The scripts need Power BI Desktop, Python 3.12 and [pbi-cli](https://pypi.org/project/pbi-cli-tool/) (for its bundled Analysis Services libraries).

1. Save the report as a Power BI Project (`.pbip`) with the PBIR format enabled, and keep it open in Desktop.
2. Build the model in the open report (`pbi connect` prints the port):

   ```powershell
   py -3.12 sample/build_pl_model.py localhost:<port>
   ```

3. Save and close Desktop, then add the pages and the demo charts:

   ```powershell
   py -3.12 sample/build_pl_pages.py sample/GroupedHeaderMatrixSample.Report
   ```

   ```powershell
   py -3.12 sample/add_demo_charts.py sample/GroupedHeaderMatrixSample.Report
   ```

4. Open the `.pbip` again.

Re-running a script replaces only the objects it created; formatting already set on the test-page matrices is kept.

## Your own data

Use any dataset with at least two descriptive dimensions and one numeric measure.

| Visual Role | Example Field |
| --- | --- |
| Rows | Region, Category (or a Line > Detail > Account hierarchy) |
| Columns | Year, Quarter (or empty for a table layout) |
| Values | Sales Amount, Profit |
| Cell Background / Font Color | Optional measures returning colors such as `"#F7F7F7"` |
| Row Background / Font Color, Row Bold | Optional measures that style whole rows |

For AppSource submission, use this PBIX or a smaller derivative with the same scenarios.
