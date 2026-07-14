# Group Header Matrix Sample Guidance

Use any tabular dataset that has at least two descriptive dimensions and one numeric measure.

Example field mapping:

| Visual Role | Example Field |
| --- | --- |
| Rows | Region, Category |
| Columns | Year, Quarter |
| Values | Sales Amount, Profit |
| Cell Background Color | Optional DAX measure returning color strings such as `#f7f7f7` |
| Cell Font Color | Optional DAX measure returning color strings such as `#333333` |

For AppSource submission, include an offline `.pbix` sample report that demonstrates:

- grouped row headers
- grouped column headers
- expand and collapse controls
- subtotals and grand total placement
- pagination or virtualization mode
- keyboard selection and cross-filtering with another native Power BI visual
