# Power BI Custom Visual - Matrix with Group Headers

## Overview
This visual is a custom Matrix visualization for Power BI that supports grouping of row and column headers, similar to Excel Pivot Tables. It allows users to expand and collapse hierarchy levels dynamically.

### Key Features
*   **Data Binding**: Accepts Rows, Columns, and Values (Measures) via the standard Power BI data view mapping.
*   **Hierarchy Support**: Supports multiple levels of row and column hierarchies with expand/collapse functionality.
*   **Layouts**:
    *   **Tabular (Default)**: Standard matrix layout.
    *   **Compact**: Stepped layout for row headers, similar to the "Compact" form in Pivot Tables.
*   **Interactivity**:
    *   **Expand/Collapse**: Buttons for individual groups and global "Expand All" / "Collapse All".
    *   **Column Resizing**: Interactive column width adjustment.
    *   **Sticky Headers**: Headers remain visible while scrolling.
*   **Formatting**:
    *   Extensive customization for fonts (headers, data), colors (backgrounds, text), and borders.
    *   **Conditional Formatting**: Supports cell background and font color based on measure values.
    *   **Zebra Striping**: Options for row and column striping.
    *   **Grid Lines**: Configurable horizontal and vertical grid lines.
*   **Totals**:
    *   **Subtotals**: Configurable visibility and position (Top/Bottom) for rows and columns.
    *   **Grand Totals**: Configurable visibility and position.

## Code Analysis

### Structure
*   `src/visual.ts`: The core logic of the visual.
    *   `update(options)`: Main entry point for rendering. Handles data processing, DOM construction, and applying settings.
    *   `renderMatrix(matrix)`: The primary rendering function. Builds the HTML table structure.
    *   `collectOutlineRowsWithTotals(...)`: Traverses the data hierarchy to flatten it into renderable rows, handling totals and layout logic.
*   `src/settings.ts`: Defines the formatting model using `powerbi-visuals-utils-formattingmodel`.
*   `capabilities.json`: Defines data roles and object properties exposed to the Power BI pane.

### Logic Flow
1.  **Data Update**: `update` receives `VisualUpdateOptions` containing the `DataView`.
2.  **Settings Parsing**: Formatting settings are extracted using `FormattingSettingsService`.
3.  **Matrix Processing**: The `DataViewMatrix` is processed to determine visible columns (`computeDisplayColumns`) and rows (`collectOutlineRowsWithTotals`).
4.  **Rendering**: An HTML `<table>` is constructed.
    *   Headers are built based on column hierarchy depth.
    *   Body rows are iterated, applying styles and formatting values.
    *   Sticky headers are applied using JavaScript-based positioning.

## Identified Bugs & Issues

### 1. Dead Code
The file `src/visual.ts` contains a significant amount of unused code, specifically the `renderRowGroup` function and its recursive helpers (`collectLeaves`, `collectRowLeaves`, `countLeaves`). These functions appear to be from an older or alternative rendering implementation and are never called by `renderMatrix`.

### 2. Duplicate Grand Total Logic
In `renderMatrix`, the code filters out rows to remove potential duplicates of the Grand Total:
```typescript
const prunedRows = outlineRows.filter(r => { ... });
```
This suggests that the row generation logic (`collectOutlineRowsWithTotals`) might be traversing the data in a way that produces redundant total rows, which is then patched by filtering. This is inefficient and indicates a flaw in the traversal logic.

### 3. Sticky Header Implementation
Sticky headers are implemented using a manual JavaScript approach with `requestAnimationFrame` (`applyStickyOffsets`). This updates the `top` style of header cells on scroll/resize.
*   **Issue**: This can lead to jittery performance and synchronization issues compared to the native CSS `position: sticky`.
*   **Recommendation**: Refactor to use CSS `position: sticky` on `<th>` elements and the `<thead>` container.

### 4. Performance (Rendering)
The visual clears the entire container and rebuilds the DOM from scratch on every update:
```typescript
while (this.contentHost.firstChild) this.contentHost.removeChild(this.contentHost.firstChild);
```
*   **Issue**: This is inefficient for large datasets and can cause flickering or sluggishness.
*   **Recommendation**: Implement virtualization (rendering only visible rows) or incremental DOM updates.

### 5. Layout Logic Complexity
The `collectOutlineRowsWithTotals` function handles both "Compact" and "Tabular" layouts, as well as total generation logic. This makes the function complex and harder to maintain or extend.
*   **Recommendation**: Split the row generation logic into separate strategies or helper functions for each layout mode.

## Recommendations

1.  **Clean Up Codebase**: Remove the identified dead code (`renderRowGroup` and associated unused methods) to improve maintainability.
2.  **Refactor Sticky Headers**: Replace the manual JS implementation with CSS `position: sticky`. This is a low-effort, high-impact improvement for user experience.
3.  **Optimize Total Generation**: Refactor `collectOutlineRowsWithTotals` to correctly generate the Grand Total once, removing the need for the post-generation filter.
4.  **Implement Virtualization**: For better performance with large matrices, implement row virtualization.
5.  **Separate Layout Strategies**: Refactor the row generation logic to separate "Compact" and "Tabular" layout handling.
6.  **Accessibility**: Add `aria-label` attributes to toolbar buttons for better accessibility.
