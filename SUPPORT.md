# Support Document for Group Header Matrix

Developer: S. Rathinagiri

Website: https://www.rathinagiri.in

Support Email: srg@rathinagiri.in

## 1. Introduction

Group Header Matrix is a Microsoft Power BI custom visual for Excel-like matrix analysis with grouped row and column headers, expand/collapse behavior, subtotals, sorting, pagination, virtualization, and table formatting.

## 2. Getting Started

To use Group Header Matrix in Power BI:

1. Import the visual from AppSource or from a `.pbiviz` package.
2. Assign fields to the data roles:
   - Rows: row hierarchy fields.
   - Columns: column hierarchy fields.
   - Values: measures displayed in the matrix.
   - Cell Background Color: optional measure values that contain color strings.
   - Cell Font Color: optional measure values that contain color strings.
3. Use the toolbar to expand or collapse groups, switch label display, and choose normal, pagination, or virtualization mode.
4. Use the format pane to configure labels, colors, grid lines, zebra striping, subtotals, grand total behavior, page size, and theme presets.

## 3. Formatting

The format pane includes settings for sticky headers, page size, label fonts, header and cell colors, grid lines, row and column striping, subtotals, grand totals, and preset themes.

## 4. Troubleshooting

Visual is blank: Confirm that Rows, Columns, and Values fields are assigned. At least one column hierarchy and one value measure are recommended.

Values look blank after collapse: Collapsed row or column groups show host-provided subtotal values only. Enable row and column subtotals in Power BI if aggregate values are required.

Performance is slow: Use Pagination or Virtualization mode, reduce page size, or reduce the number of row and column hierarchy members.

Selection does not affect other visuals: Confirm report visual interactions are enabled and that the target visuals are configured to receive filtering or highlighting from this visual.

## 5. Technical Support

Issue Tracker: https://github.com/SRathinaGiri/PowerBIGroupHeaderMatrix/issues

Direct Contact: email srg@rathinagiri.in with a description of the issue, Power BI version, sample data structure if possible, and a screenshot.

## 6. Version History

v1.0.0.0: Certification-readiness update with keyboard navigation, multi-visual selection, high contrast support, rendering events, documentation, and certification-audited packaging.
