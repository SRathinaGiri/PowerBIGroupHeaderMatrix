# Group Header Matrix Sample

This folder contains `GroupedHeaderMatrixSample.pbix`, an offline Power BI report that showcases the Group Header Matrix visual.

The sample demonstrates:

- grouped row headers
- grouped column headers
- expand and collapse controls
- subtotals and grand total placement
- keyboard selection and cross-filtering with another native Power BI visual

For your own tests, use any tabular dataset that has at least two descriptive dimensions and one numeric measure.

Example field mapping:

| Visual Role | Example Field |
| --- | --- |
| Rows | Region, Category |
| Columns | Year, Quarter |
| Values | Sales Amount, Profit |
| Cell Background Color | Optional DAX measure returning color strings such as `#f7f7f7` |
| Cell Font Color | Optional DAX measure returning color strings such as `#333333` |

For AppSource submission, use this PBIX or a smaller derivative with the same scenarios.
