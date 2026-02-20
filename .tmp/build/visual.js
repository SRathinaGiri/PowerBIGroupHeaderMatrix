/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import "./../style/visual.less";
import { VisualFormattingSettingsModel } from "./settings";
export class Visual {
    container;
    toolbar;
    contentHost;
    debugEl;
    debugEnabled = false;
    table = null;
    formattingSettings;
    formattingSettingsService;
    // Collapse/expand state
    collapsedRowKeys = new Set();
    collapsedColKeys = new Set();
    lastMatrix = null;
    collapseInitialized = false;
    // Column sizing
    columnWidthPx = new Map();
    colElsByKey = new Map();
    repeatLabels = false;
    rowHeaderMinWidth = 160;
    rowHeaderFontSize = 12;
    rowHeaderFontFamily = "";
    rowHeaderBold = false;
    colHeaderFontSize = 11;
    colHeaderFontFamily = "";
    colHeaderBold = false;
    compactLayout = false;
    measureFormats = [];
    displayMeasureIndices = [];
    rowBackColorMeasureIndex = -1;
    rowFontColorMeasureIndex = -1;
    dataFontSize = 11;
    dataFontFamily = "";
    dataBold = false;
    rowHeaderColor;
    rowHeaderBg;
    colHeaderColor;
    colHeaderBg;
    cellColor;
    cellBg;
    rowLevelStyles = [];
    colLevelStyles = [];
    host;
    // Totals behavior
    showGrandTotal = true;
    grandTotalFallback = false;
    constructor(options) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.host = options.host;
        this.container = document.createElement("div");
        this.container.className = "ghm-container";
        // Toolbar
        this.toolbar = document.createElement("div");
        this.toolbar.className = "ghm-toolbar";
        const btnExpand = document.createElement("button");
        btnExpand.className = "ghm-btn";
        btnExpand.textContent = "➕ Expand All";
        btnExpand.addEventListener("click", () => { this.expandAll(); this.persistState(); this.refresh(); });
        const btnCollapse = document.createElement("button");
        btnCollapse.className = "ghm-btn";
        btnCollapse.textContent = "➖ Collapse All";
        btnCollapse.addEventListener("click", () => { this.collapseAll(); this.persistState(); this.refresh(); });
        const btnRowExpandLevel = document.createElement("button");
        btnRowExpandLevel.className = "ghm-btn";
        btnRowExpandLevel.textContent = "+𝄘";
        btnRowExpandLevel.title = "Expand Row Level";
        btnRowExpandLevel.addEventListener("click", () => { this.expandRowLevel(); this.persistState(); this.refresh(); });
        // Use requested glyph-only label
        btnRowExpandLevel.textContent = "+ 𝄘";
        const btnRowCollapseLevel = document.createElement("button");
        btnRowCollapseLevel.className = "ghm-btn";
        btnRowCollapseLevel.textContent = "−𝄘";
        btnRowCollapseLevel.title = "Collapse Row Level";
        btnRowCollapseLevel.addEventListener("click", () => { this.collapseRowLevel(); this.persistState(); this.refresh(); });
        btnRowCollapseLevel.textContent = "- 𝄘";
        const btnColExpandLevel = document.createElement("button");
        btnColExpandLevel.className = "ghm-btn";
        btnColExpandLevel.textContent = "+⦀";
        btnColExpandLevel.title = "Expand Column Level";
        btnColExpandLevel.addEventListener("click", () => { this.expandColLevel(); this.persistState(); this.refresh(); });
        btnColExpandLevel.textContent = "+ ⦀";
        const btnColCollapseLevel = document.createElement("button");
        btnColCollapseLevel.className = "ghm-btn";
        btnColCollapseLevel.textContent = "−⦀";
        btnColCollapseLevel.title = "Collapse Column Level";
        btnColCollapseLevel.addEventListener("click", () => { this.collapseColLevel(); this.persistState(); this.refresh(); });
        btnColCollapseLevel.textContent = "- ⦀";
        const lblRepeat = document.createElement("label");
        lblRepeat.style.fontSize = "12px";
        const chkRepeat = document.createElement("input");
        chkRepeat.type = "checkbox";
        chkRepeat.id = "ghm-repeat";
        chkRepeat.addEventListener("change", () => { this.repeatLabels = chkRepeat.checked; this.persistState(); this.refresh(); });
        lblRepeat.appendChild(chkRepeat);
        lblRepeat.appendChild(document.createTextNode(" Repeat Labels"));
        // Compact layout toolbar toggle
        const lblCompact = document.createElement("label");
        lblCompact.style.fontSize = "12px";
        const chkCompact = document.createElement("input");
        chkCompact.type = "checkbox";
        chkCompact.id = "ghm-compact";
        chkCompact.addEventListener("change", () => {
            this.compactLayout = chkCompact.checked;
            if (this.compactLayout) {
                this.repeatLabels = false;
                chkRepeat.checked = false;
                chkRepeat.disabled = true;
            }
            else {
                chkRepeat.disabled = false;
            }
            this.persistState();
            this.refresh();
        });
        lblCompact.appendChild(chkCompact);
        lblCompact.appendChild(document.createTextNode(" Compact"));
        // Add buttons
        this.toolbar.appendChild(btnExpand);
        const btnCollapseAll = document.createElement("button");
        btnCollapseAll.className = "ghm-btn";
        btnCollapseAll.textContent = "- All";
        btnCollapseAll.title = "Collapse All";
        btnCollapseAll.addEventListener("click", () => { this.collapseAll(); this.persistState(); this.refresh(); });
        btnExpand.textContent = "+ All";
        btnExpand.title = "Expand All";
        this.toolbar.appendChild(btnRowExpandLevel);
        this.toolbar.appendChild(btnRowCollapseLevel);
        this.toolbar.appendChild(btnColExpandLevel);
        this.toolbar.appendChild(btnColCollapseLevel);
        this.toolbar.appendChild(btnCollapseAll);
        this.toolbar.appendChild(lblRepeat);
        this.toolbar.appendChild(lblCompact);
        const btnDebug = document.createElement("button");
        btnDebug.className = "ghm-btn";
        btnDebug.textContent = "🧪 Debug";
        btnDebug.addEventListener("click", () => { this.debugEnabled = !this.debugEnabled; this.updateDebugOverlay(); });
        // Hide debug button in production
        try {
            btnDebug.style.display = "none";
        }
        catch { }
        this.toolbar.appendChild(btnDebug);
        this.container.appendChild(this.toolbar);
        // content host
        this.contentHost = document.createElement("div");
        this.container.appendChild(this.contentHost);
        // debug overlay
        this.debugEl = document.createElement("div");
        this.debugEl.className = "ghm-debug";
        this.debugEl.style.display = "none";
        this.container.appendChild(this.debugEl);
        options.element.appendChild(this.container);
    }
    update(options) {
        const dataView = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        // Apply sticky preference if present
        const sticky = this.getObjectValue(dataView?.metadata?.objects, "state", "stickyHeaders", true);
        this.container.classList.toggle("ghm-sticky", !!sticky);
        this.repeatLabels = this.getObjectValue(dataView?.metadata?.objects, "state", "repeatLabels", false);
        this.compactLayout = this.getObjectValue(dataView?.metadata?.objects, "state", "compactLayout", false);
        if (this.compactLayout)
            this.repeatLabels = false;
        // Grand total formatting options
        this.showGrandTotal = this.getObjectValue(dataView?.metadata?.objects, "grandTotal", "show", true);
        this.grandTotalFallback = this.getObjectValue(dataView?.metadata?.objects, "grandTotal", "fallbackToRootValues", false);
        this.rowHeaderMinWidth = this.getObjectValue(dataView?.metadata?.objects, "state", "rowHeaderMinWidth", 160) || 160;
        this.rowHeaderFontSize = this.getObjectValue(dataView?.metadata?.objects, "labels", "rowHeaderFontSize", 12) || 12;
        this.rowHeaderFontFamily = this.getObjectValue(dataView?.metadata?.objects, "labels", "rowHeaderFontFamily", { value: "" }) || this.getObjectValue(dataView?.metadata?.objects, "labels", "rowHeaderFontFamily", "");
        this.rowHeaderBold = this.getObjectValue(dataView?.metadata?.objects, "labels", "rowHeaderBold", false);
        this.colHeaderFontSize = this.getObjectValue(dataView?.metadata?.objects, "labels", "colHeaderFontSize", 11) || 11;
        this.colHeaderFontFamily = this.getObjectValue(dataView?.metadata?.objects, "labels", "colHeaderFontFamily", { value: "" }) || this.getObjectValue(dataView?.metadata?.objects, "labels", "colHeaderFontFamily", "");
        this.colHeaderBold = this.getObjectValue(dataView?.metadata?.objects, "labels", "colHeaderBold", false);
        this.dataFontSize = this.getObjectValue(dataView?.metadata?.objects, "labels", "dataFontSize", 11) || 11;
        this.dataFontFamily = this.getObjectValue(dataView?.metadata?.objects, "labels", "dataFontFamily", { value: "" }) || this.getObjectValue(dataView?.metadata?.objects, "labels", "dataFontFamily", "");
        this.dataBold = this.getObjectValue(dataView?.metadata?.objects, "labels", "dataBold", false);
        this.rowHeaderColor = this.getObjectValue(dataView?.metadata?.objects, "colors", "rowHeaderColor", { value: "" })?.value ?? "";
        this.rowHeaderBg = this.getObjectValue(dataView?.metadata?.objects, "colors", "rowHeaderBg", { value: "" })?.value ?? "";
        this.colHeaderColor = this.getObjectValue(dataView?.metadata?.objects, "colors", "colHeaderColor", { value: "" })?.value ?? "";
        this.colHeaderBg = this.getObjectValue(dataView?.metadata?.objects, "colors", "colHeaderBg", { value: "" })?.value ?? "";
        this.cellColor = this.getObjectValue(dataView?.metadata?.objects, "colors", "cellColor", { value: "" })?.value ?? "";
        this.cellBg = this.getObjectValue(dataView?.metadata?.objects, "colors", "cellBg", { value: "" })?.value ?? "";
        // Reflect toolbar checkbox states
        const chkRepeatEl = this.toolbar.querySelector('#ghm-repeat');
        if (chkRepeatEl) {
            chkRepeatEl.checked = this.repeatLabels && !this.compactLayout;
            chkRepeatEl.disabled = !!this.compactLayout;
        }
        const chkCompactEl = this.toolbar.querySelector('#ghm-compact');
        if (chkCompactEl)
            chkCompactEl.checked = !!this.compactLayout;
        // Load persisted widths and collapsed sets
        const widthsJson = this.getObjectValue(dataView?.metadata?.objects, "state", "columnWidths", "");
        if (widthsJson) {
            try {
                const w = JSON.parse(widthsJson);
                this.columnWidthPx = new Map(Object.entries(w));
            }
            catch { }
        }
        const colsJson = this.getObjectValue(dataView?.metadata?.objects, "state", "collapsedCols", "");
        if (colsJson) {
            try {
                this.collapsedColKeys = new Set(JSON.parse(colsJson));
                this.collapseInitialized = true;
            }
            catch { }
        }
        const rowsJson = this.getObjectValue(dataView?.metadata?.objects, "state", "collapsedRows", "");
        if (rowsJson) {
            try {
                this.collapsedRowKeys = new Set(JSON.parse(rowsJson));
                this.collapseInitialized = true;
            }
            catch { }
        }
        // If no measure is bound, force all expanded to show full structure
        const noMeasures = !(dataView && dataView.matrix && dataView.matrix.valueSources && dataView.matrix.valueSources.length > 0);
        if (noMeasures) {
            this.collapsedColKeys.clear();
            this.collapsedRowKeys.clear();
            this.collapseInitialized = true;
        }
        // Clear only content (preserve toolbar)
        while (this.contentHost.firstChild)
            this.contentHost.removeChild(this.contentHost.firstChild);
        if (!dataView || !dataView.matrix) {
            this.renderPlaceholder("Add Columns hierarchy and a measure");
            return;
        }
        try {
            this.lastMatrix = dataView.matrix;
            if (!this.collapseInitialized) {
                const mc = (this.lastMatrix.valueSources && this.lastMatrix.valueSources.length) ? this.lastMatrix.valueSources.length : 0;
                const enableDefault = mc > 0; // don't default-collapse when no measures bound
                this.initializeDefaultCollapsed(this.lastMatrix, enableDefault);
                this.collapseInitialized = true;
            }
            this.loadLevelStyles(this.lastMatrix);
            this.renderMatrix(dataView.matrix);
        }
        catch (e) {
            this.renderPlaceholder("Unable to render matrix");
            console.error(e);
        }
    }
    getFormattingModel() {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
    // Enumerates per-level subtotal toggles so the host can attach properties to individual row/column levels
    enumerateObjectInstances(options) {
        const enumeration = [];
        if (!this.lastMatrix)
            return enumeration;
        if (options.objectName === "subtotalPerLevel") {
            const rows = this.lastMatrix.rows;
            const cols = this.lastMatrix.columns;
            // Row levels
            if (rows && rows.levels && rows.levels.length) {
                rows.levels.forEach((lvl, i) => {
                    const src = lvl.sources && lvl.sources[0];
                    if (!src)
                        return;
                    const displayName = src.displayName || `Row level ${i + 1}`;
                    const selector = src.queryName ? { metadata: src.queryName } : null;
                    enumeration.push({
                        objectName: options.objectName,
                        displayName: `Row: ${displayName}`,
                        properties: { levelSubtotalEnabled: true },
                        selector
                    });
                });
            }
            // Column levels
            if (cols && cols.levels && cols.levels.length) {
                cols.levels.forEach((lvl, i) => {
                    const src = lvl.sources && lvl.sources[0];
                    if (!src)
                        return;
                    const displayName = src.displayName || `Column level ${i + 1}`;
                    const selector = src.queryName ? { metadata: src.queryName } : null;
                    enumeration.push({
                        objectName: options.objectName,
                        displayName: `Column: ${displayName}`,
                        properties: { levelSubtotalEnabled: true },
                        selector
                    });
                });
            }
        }
        return enumeration;
    }
    renderPlaceholder(text) {
        const el = document.createElement("div");
        el.className = "ghm-placeholder";
        el.textContent = text;
        this.contentHost.appendChild(el);
    }
    renderMatrix(matrix) {
        const columns = matrix.columns;
        const rows = matrix.rows;
        if (!columns || !columns.root || !columns.root.children || columns.root.children.length === 0) {
            this.renderPlaceholder("No column hierarchy provided");
            return;
        }
        const table = document.createElement("table");
        table.className = "ghm-table";
        // Apply a global font if only one family provided
        if (this.rowHeaderFontFamily && !this.colHeaderFontFamily)
            table.style.fontFamily = this.rowHeaderFontFamily;
        if (this.colHeaderFontFamily && !this.rowHeaderFontFamily)
            table.style.fontFamily = this.colHeaderFontFamily;
        const colgroup = document.createElement("colgroup");
        const thead = document.createElement("thead");
        const tbody = document.createElement("tbody");
        // Determine columns to display (compress collapsed groups to a single column)
        const colDepth = this.getColumnDepth(columns);
        const displayCols = this.computeDisplayColumns(columns.root, colDepth);
        // Build column header rows from the display list
        const headerRows = this.buildHeaderRowsFromDisplay(displayCols, colDepth);
        // Measures handling: compute display measures vs style measures
        const totalMeasureCount = (matrix.valueSources && matrix.valueSources.length) ? matrix.valueSources.length : 0;
        this.displayMeasureIndices = [];
        this.rowBackColorMeasureIndex = -1;
        this.rowFontColorMeasureIndex = -1;
        const vsArr = matrix.valueSources || [];
        for (let i = 0; i < vsArr.length; i++) {
            const roles = vsArr[i].roles || {};
            if (roles.measure)
                this.displayMeasureIndices.push(i);
            if (roles.rowBackColor)
                this.rowBackColorMeasureIndex = i;
            if (roles.rowFontColor)
                this.rowFontColorMeasureIndex = i;
        }
        const measureCount = this.displayMeasureIndices.length;
        const measureLabels = this.displayMeasureIndices.map((i, idx) => String(vsArr[i].displayName || vsArr[i].queryName || `Measure ${idx + 1}`));
        this.measureFormats = (vsArr || []).map(m => String(m.format || ""));
        if (measureCount > 1) {
            for (const row of headerRows) {
                for (const cell of row)
                    cell.span = (cell.span || 1) * measureCount;
            }
        }
        // Determine number of row header levels (depth)
        const rowDepth = this.getRowDepth(rows);
        // Hide deeper row header columns when levels are collapsed, even in Repeat Labels mode
        const visibleRowDepth = this.getVisibleRowDepth(this.lastMatrix?.rows, rowDepth);
        const rowHeaderCols = (rowDepth > 0) ? (this.compactLayout ? 1 : visibleRowDepth) : 0;
        // Include the thin resizer row as part of the sticky header block
        const headerDepth = headerRows.length + (measureCount > 1 ? 1 : 0) + 1;
        headerRows.forEach((rowCells, i) => {
            const tr = document.createElement("tr");
            if (i === 0 && rowHeaderCols > 0) {
                const corner = document.createElement("th");
                corner.className = "ghm-corner";
                corner.rowSpan = headerDepth;
                if (rowHeaderCols > 1)
                    corner.colSpan = rowHeaderCols;
                tr.appendChild(corner);
            }
            for (const cell of rowCells) {
                const th = document.createElement("th");
                th.className = "ghm-colheader";
                th.title = cell.label;
                // Toggle indicator and label
                if (cell.togglable) {
                    const toggle = document.createElement("span");
                    toggle.className = "ghm-toggle";
                    toggle.textContent = cell.collapsed ? "+" : "−";
                    toggle.addEventListener("click", (ev) => {
                        ev.stopPropagation();
                        if (cell.collapsed)
                            this.collapsedColKeys.delete(cell.key);
                        else
                            this.collapsedColKeys.add(cell.key);
                        this.persistState();
                        this.refresh();
                    });
                    th.appendChild(toggle);
                }
                const txt = document.createElement("span");
                txt.textContent = cell.label;
                const ch = this.colLevelStyles[i] || {};
                const fs = ch.fontSize ?? this.colHeaderFontSize;
                if (fs)
                    txt.style.fontSize = `${fs}px`;
                txt.style.fontFamily = (ch.fontFamily || this.colHeaderFontFamily) || "";
                if (this.colHeaderBold || ch.fontWeight)
                    txt.style.fontWeight = (ch.fontWeight || "bold");
                if (ch.italic)
                    txt.style.fontStyle = "italic";
                if (ch.underline)
                    txt.style.textDecoration = "underline";
                if (this.colHeaderColor)
                    th.style.color = this.colHeaderColor;
                if (this.colHeaderBg)
                    th.style.backgroundColor = this.colHeaderBg;
                th.appendChild(txt);
                if (cell.span && cell.span > 1)
                    th.colSpan = cell.span;
                tr.appendChild(th);
            }
            thead.appendChild(tr);
        });
        if (measureCount > 1) {
            const tr = document.createElement("tr");
            for (const ref of displayCols) {
                for (let m = 0; m < measureCount; m++) {
                    const th = document.createElement("th");
                    th.className = "ghm-colheader";
                    th.textContent = measureLabels[m] ?? `M${m + 1}`;
                    tr.appendChild(th);
                }
            }
            thead.appendChild(tr);
        }
        // Resizer row (always present so handles align with columns)
        const resizerRow = document.createElement("tr");
        resizerRow.className = "ghm-resizers-row";
        // Prepend row-header placeholders so cell count matches total columns
        for (let r = 0; r < rowHeaderCols; r++) {
            const th = document.createElement("th");
            th.className = "ghm-resizer-cell";
            resizerRow.appendChild(th);
        }
        const resizerRefs = displayCols;
        for (const ref of resizerRefs) {
            for (let m = 0; m < measureCount; m++) {
                const th = document.createElement("th");
                th.className = "ghm-resizer-cell";
                const key = (ref.kind === "leaf") ? ref.keys.join("||") : ref.key;
                const handle = document.createElement("div");
                handle.className = "ghm-resizer";
                handle.title = "Drag to resize column";
                handle.addEventListener("mousedown", (e) => this.beginResize(e, key));
                th.appendChild(handle);
                resizerRow.appendChild(th);
            }
        }
        thead.appendChild(resizerRow);
        // Determine column leaf count (for cell generation) and order
        const columnLeaves = displayCols;
        // Build colgroup to control widths (reserve all row-header columns; hide non-visible with width=0)
        this.colElsByKey.clear();
        for (let r = 0; r < rowHeaderCols; r++) {
            const col = document.createElement("col");
            if (!this.compactLayout && r < visibleRowDepth)
                col.style.width = (r === 0 ? `${this.rowHeaderMinWidth}px` : "120px");
            else if (this.compactLayout && r === 0)
                col.style.width = `${this.rowHeaderMinWidth}px`;
            else
                col.style.width = "0px";
            colgroup.appendChild(col);
        }
        for (const ref of columnLeaves) {
            const key = (ref.kind === "leaf") ? ref.keys.join("||") : ref.key;
            const width = this.columnWidthPx.get(key) ?? 120;
            for (let m = 0; m < measureCount; m++) {
                const col = document.createElement("col");
                col.style.width = `${width}px`;
                colgroup.appendChild(col);
                if (!this.colElsByKey.has(key))
                    this.colElsByKey.set(key, []);
                this.colElsByKey.get(key).push(col);
            }
        }
        const colLeafCount = columnLeaves.length;
        // Do not inject a separate "Grand Total" row here.
        // Body rows (including the root grand total when applicable) are generated
        // exclusively by collectOutlineRowsWithTotals() below. This avoids any
        // chance of duplicating the total row.
        // Build body rows
        if (rows && rows.root && rows.root.children && rows.root.children.length) {
            const outlineRows = this.collectOutlineRowsWithTotals(rows.root, rowDepth);
            // Deduplicate any duplicate Grand Total rows just in case
            let sawGT = false;
            const filteredRows = outlineRows.filter(r => {
                const isGT = r.isTotal && r.labels && r.labels[0] === "Grand Total";
                if (isGT) {
                    if (sawGT)
                        return false;
                    sawGT = true;
                }
                return true;
            });
            for (const rowInfoRaw of filteredRows) {
                // When Repeat Labels is off, keep only the leaf label visible
                // (the last populated level) for non-total rows.
                const rowInfo = { ...rowInfoRaw };
                // Precompute row-level style from style measures, if present
                const valuesMapForRow = rowInfo.valuesMap || {};
                const rowStyle = this.getRowStyleFromValuesMap(valuesMapForRow, totalMeasureCount);
                let labelsToUse = rowInfo.labels;
                if (!this.compactLayout && !this.repeatLabels && !rowInfo.isTotal) {
                    const newLabels = new Array(rowDepth).fill("");
                    let lastIdx = -1;
                    for (let i = rowDepth - 1; i >= 0; i--) {
                        if (rowInfo.labels && rowInfo.labels[i]) {
                            lastIdx = i;
                            break;
                        }
                    }
                    if (lastIdx >= 0)
                        newLabels[lastIdx] = rowInfo.labels[lastIdx];
                    labelsToUse = newLabels;
                }
                const tr = document.createElement("tr");
                if (rowInfo.isTotal)
                    tr.className = "ghm-totalrow";
                if (rowHeaderCols > 0) {
                    if (this.compactLayout) {
                        const th = document.createElement("th");
                        th.className = "ghm-rowheader";
                        // deepest available label
                        let txt = "";
                        let lastIdx = -1;
                        for (let i = rowDepth - 1; i >= 0; i--) {
                            if (labelsToUse && labelsToUse[i]) {
                                txt = labelsToUse[i];
                                lastIdx = i;
                                break;
                            }
                        }
                        th.textContent = txt || "";
                        this.applyRowHeaderStyle(th, 0);
                        if (rowStyle.bg)
                            th.style.backgroundColor = rowStyle.bg;
                        if (rowStyle.color)
                            th.style.color = rowStyle.color;
                        // Indentation according to depth in compact layout
                        const depthIndent = rowInfo.isTotal && rowInfo.depth !== undefined
                            ? Math.max(0, Math.min(rowDepth - 1, rowInfo.depth))
                            : Math.max(0, lastIdx);
                        const basePad = 8, step = 14;
                        th.style.paddingLeft = `${basePad + step * depthIndent}px`;
                        if (rowInfo.isTotal && rowInfo.toggleKey) {
                            const toggle = document.createElement("span");
                            toggle.className = "ghm-toggle";
                            const collapsed = !!rowInfo.collapsed;
                            toggle.textContent = collapsed ? "+" : "-";
                            toggle.title = collapsed ? "Expand group" : "Collapse group";
                            toggle.style.marginRight = "6px";
                            toggle.addEventListener("click", (ev) => {
                                ev.stopPropagation();
                                const key = String(rowInfo.toggleKey);
                                if (collapsed)
                                    this.collapsedRowKeys.delete(key);
                                else
                                    this.collapsedRowKeys.add(key);
                                this.persistState();
                                this.refresh();
                            });
                            th.prepend(toggle);
                        }
                        if (this.rowHeaderBg)
                            th.style.backgroundColor = this.rowHeaderBg;
                        tr.appendChild(th);
                    }
                    else {
                        const toggleLevel = Math.max(0, Math.min(rowHeaderCols - 1, rowInfo.depth ?? 0));
                        for (let lvl = 0; lvl < rowHeaderCols; lvl++) {
                            const th = document.createElement("th");
                            th.className = "ghm-rowheader";
                            th.textContent = (labelsToUse && labelsToUse[lvl]) || "";
                            this.applyRowHeaderStyle(th, lvl);
                            if (rowStyle.bg)
                                th.style.backgroundColor = rowStyle.bg;
                            if (rowStyle.color)
                                th.style.color = rowStyle.color;
                            if (this.rowHeaderBg)
                                th.style.backgroundColor = this.rowHeaderBg;
                            // Place +/- toggle on the appropriate visible level
                            if (rowInfo.isTotal && rowInfo.toggleKey && lvl === toggleLevel) {
                                const toggle = document.createElement("span");
                                toggle.className = "ghm-toggle";
                                const collapsed = !!rowInfo.collapsed;
                                toggle.textContent = collapsed ? "+" : "-";
                                toggle.title = collapsed ? "Expand group" : "Collapse group";
                                toggle.addEventListener("click", (ev) => {
                                    ev.stopPropagation();
                                    const key = String(rowInfo.toggleKey);
                                    if (collapsed)
                                        this.collapsedRowKeys.delete(key);
                                    else
                                        this.collapsedRowKeys.add(key);
                                    this.persistState();
                                    this.refresh();
                                });
                                th.prepend(toggle);
                            }
                            tr.appendChild(th);
                        }
                    }
                }
                const valuesMap = rowInfo.valuesMap || {};
                const numericKeys = Object.keys(valuesMap).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
                for (let c = 0; c < colLeafCount; c++) {
                    for (let m = 0; m < measureCount; m++) {
                        const globalM = this.displayMeasureIndices[m] ?? m;
                        const td = document.createElement("td");
                        const v = this.getCellValueForDisplayCol(valuesMap, numericKeys, columnLeaves[c], globalM, totalMeasureCount);
                        td.textContent = this.formatValueByMeasure(v, globalM);
                        this.applyCellConditionalStyle(td, valuesMap, columnLeaves[c], globalM, totalMeasureCount);
                        if (this.dataFontSize)
                            td.style.fontSize = `${this.dataFontSize}px`;
                        if (this.dataFontFamily)
                            td.style.fontFamily = this.dataFontFamily;
                        if (this.dataBold)
                            td.style.fontWeight = "bold";
                        if (this.cellColor)
                            td.style.color = this.cellColor;
                        if (this.cellBg)
                            td.style.backgroundColor = this.cellBg;
                        tr.appendChild(td);
                    }
                }
                tbody.appendChild(tr);
            }
        }
        else {
            // No row groups – single total row
            const tr = document.createElement("tr");
            // Data cells
            const valuesMap = rows && rows.root ? rows.root.values || {} : {};
            const numericKeys = Object.keys(valuesMap).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
            for (let c = 0; c < colLeafCount; c++) {
                for (let m = 0; m < measureCount; m++) {
                    const globalM = this.displayMeasureIndices[m] ?? m;
                    const td = document.createElement("td");
                    const v = this.getCellValueForDisplayCol(valuesMap, numericKeys, columnLeaves[c], globalM, totalMeasureCount);
                    td.textContent = this.formatValueByMeasure(v, globalM);
                    this.applyCellConditionalStyle(td, valuesMap, columnLeaves[c], globalM, totalMeasureCount);
                    if (this.dataFontSize)
                        td.style.fontSize = `${this.dataFontSize}px`;
                    if (this.dataFontFamily)
                        td.style.fontFamily = this.dataFontFamily;
                    if (this.dataBold)
                        td.style.fontWeight = "bold";
                    if (this.cellColor)
                        td.style.color = this.cellColor;
                    if (this.cellBg)
                        td.style.backgroundColor = this.cellBg;
                    tr.appendChild(td);
                }
            }
            tbody.appendChild(tr);
        }
        table.appendChild(colgroup);
        table.appendChild(thead);
        table.appendChild(tbody);
        this.contentHost.appendChild(table);
        this.applyStickyOffsets(thead);
        this.updateDebugOverlay({ table, headerRows, measureCount, columnLeaves, rowDepth, colDepth });
    }
    computeDisplayColumns(root, depth) {
        const leaves = [];
        const ranges = new Map();
        const subtotalOffsetByKey = new Map();
        let offset = 0;
        const walk = (node, labels, keys, parentKey, underSubtotal) => {
            const label = this.nodeLabel(node);
            const key = [...keys, label].filter(Boolean).join("||");
            const newLabels = [...labels, label];
            const newKeys = [...keys, label];
            const isSubtotal = node.isSubtotal === true;
            if (!node.children || node.children.length === 0) {
                // Detect subtotal leaf: record parent group key -> offset
                if (underSubtotal || isSubtotal) {
                    if (parentKey)
                        subtotalOffsetByKey.set(parentKey, offset);
                }
                let collapsedAt = null;
                for (let i = 0; i < newLabels.length; i++) {
                    const k = newKeys.slice(0, i + 1).filter(Boolean).join("||");
                    if (this.collapsedColKeys.has(k)) {
                        collapsedAt = i;
                        break;
                    }
                }
                while (newLabels.length < depth)
                    newLabels.push("");
                while (newKeys.length < depth)
                    newKeys.push(newKeys[newKeys.length - 1] || "");
                leaves.push({ offset, labels: newLabels, keys: newKeys, collapsedAt });
                if (collapsedAt !== null) {
                    const gkey = newKeys.slice(0, collapsedAt + 1).filter(Boolean).join("||");
                    const r = ranges.get(gkey);
                    if (!r)
                        ranges.set(gkey, { start: offset, end: offset, level: collapsedAt });
                    else
                        r.end = offset;
                }
                offset++;
                return;
            }
            for (const ch of node.children) {
                const chLabel = this.nodeLabel(ch);
                const chKey = [...newKeys, chLabel].filter(Boolean).join("||");
                walk(ch, newLabels, newKeys, key, underSubtotal || isSubtotal);
            }
        };
        if (root.children)
            for (const ch of root.children)
                walk(ch, [], [], "", false);
        const result = [];
        let i = 0;
        while (i < leaves.length) {
            const leaf = leaves[i];
            if (leaf.collapsedAt !== null) {
                const gkey = leaf.keys.slice(0, leaf.collapsedAt + 1).filter(Boolean).join("||");
                const r = ranges.get(gkey);
                const subtotalOffset = subtotalOffsetByKey.get(gkey);
                result.push({ kind: "collapsed", start: r.start, end: r.end, labels: leaf.labels, keys: leaf.keys, collapsedLevel: r.level, key: gkey, subtotalOffset });
                i = r.end + 1;
            }
            else {
                result.push({ kind: "leaf", offset: leaf.offset, labels: leaf.labels, keys: leaf.keys });
                i++;
            }
        }
        return result;
    }
    isAnyAncestorCollapsed(path) {
        // path is array of labels up to, but not including, the leaf
        let accum = [];
        for (const p of path) {
            accum.push(p);
            const k = accum.join("||");
            if (this.collapsedColKeys.has(k))
                return true;
        }
        return false;
    }
    getRowDepth(rows) {
        if (!rows || !rows.root)
            return 0;
        if (rows.levels && rows.levels.length)
            return rows.levels.length;
        const depthFrom = (node) => {
            if (!node.children || node.children.length === 0)
                return 0;
            let max = 0;
            for (const c of node.children)
                max = Math.max(max, depthFrom(c));
            return 1 + max;
        };
        if (!rows.root.children || rows.root.children.length === 0)
            return 0;
        return depthFrom(rows.root);
    }
    getColumnDepth(columns) {
        if (!columns || !columns.root)
            return 0;
        if (columns.levels && columns.levels.length)
            return columns.levels.length;
        const depthFrom = (node) => {
            if (!node.children || node.children.length === 0)
                return 1;
            let max = 0;
            for (const c of node.children)
                max = Math.max(max, depthFrom(c));
            return 1 + max;
        };
        if (!columns.root.children || columns.root.children.length === 0)
            return 1;
        return depthFrom(columns.root);
    }
    renderRowGroup(node, depth, parentKey, columnLeaves, measureCount) {
        // Skip the artificial root and render its children
        if (node.level === undefined && node.children && node.children.length) {
            let all = [];
            for (const child of node.children) {
                all = all.concat(this.renderRowGroup(child, 0, parentKey, columnLeaves, measureCount));
            }
            return all;
        }
        if (node.children && node.children.length) {
            let rows = [];
            const label = this.nodeLabel(node) || "Total";
            const thisKey = [parentKey, label].filter(Boolean).join("||");
            const collapsed = this.collapsedRowKeys.has(thisKey);
            if (collapsed) {
                // Render a single group row with group aggregates
                const tr = document.createElement("tr");
                const th = document.createElement("th");
                th.className = "ghm-rowheader";
                th.title = label;
                const toggle = document.createElement("span");
                toggle.className = "ghm-toggle";
                toggle.textContent = "+";
                toggle.addEventListener("click", (ev) => {
                    ev.stopPropagation();
                    this.collapsedRowKeys.delete(thisKey);
                    this.persistState();
                    this.refresh();
                });
                th.appendChild(toggle);
                const txt = document.createElement("span");
                txt.textContent = label;
                if (this.rowHeaderFontSize)
                    txt.style.fontSize = `${this.rowHeaderFontSize}px`;
                if (this.rowHeaderFontFamily)
                    txt.style.fontFamily = this.rowHeaderFontFamily;
                th.appendChild(txt);
                tr.appendChild(th);
                // fill remaining row header columns (collapsed depths)
                const visDepth = this.repeatLabels ? this.getRowDepth(this.lastMatrix?.rows) : this.getDisplayedRowDepth(this.lastMatrix?.rows);
                for (let k = 1; k < visDepth; k++) {
                    const thFill = document.createElement("th");
                    thFill.className = "ghm-rowheader";
                    this.applyRowHeaderStyle(thFill, k);
                    tr.appendChild(thFill);
                }
                const totalCountLocal = (this.lastMatrix?.valueSources?.length) || measureCount;
                for (let c = 0; c < columnLeaves.length; c++) {
                    for (let m = 0; m < measureCount; m++) {
                        const globalM = this.displayMeasureIndices[m] ?? m;
                        const td = document.createElement("td");
                        const v = this.getCollapsedRowGroupValue(node, columnLeaves[c], globalM, totalCountLocal);
                        td.textContent = this.formatValueByMeasure(v, globalM);
                        const valuesMapNode = (node.values || {});
                        this.applyCellConditionalStyle(td, valuesMapNode, columnLeaves[c], globalM, totalCountLocal);
                        if (this.dataFontSize)
                            td.style.fontSize = `${this.dataFontSize}px`;
                        if (this.dataFontFamily)
                            td.style.fontFamily = this.dataFontFamily;
                        if (this.dataBold)
                            td.style.fontWeight = "bold";
                        if (this.cellColor)
                            td.style.color = this.cellColor;
                        if (this.cellBg)
                            td.style.backgroundColor = this.cellBg;
                        tr.appendChild(td);
                    }
                }
                return [tr];
            }
            else {
                const normalChildren = node.children.filter(ch => !ch.isSubtotal);
                const subtotalChild = node.children.find(ch => ch.isSubtotal);
                for (const child of normalChildren) {
                    rows = rows.concat(this.renderRowGroup(child, depth + 1, thisKey, columnLeaves, measureCount));
                }
                if (subtotalChild) {
                    const totalDepth = this.getRowDepth(this.lastMatrix?.rows);
                    const trTotal = document.createElement("tr");
                    trTotal.className = "ghm-totalrow";
                    // Do not emit TH at current group depth since the row-spanning group header occupies that column across this block
                    for (let i = 0; i < totalDepth; i++) {
                        if (i === depth)
                            continue;
                        const thFill = document.createElement("th");
                        thFill.className = "ghm-rowheader";
                        this.applyRowHeaderStyle(thFill, i);
                        trTotal.appendChild(thFill);
                    }
                    const valuesMap = (subtotalChild.values || {});
                    const totalCountLocal2 = (this.lastMatrix?.valueSources?.length) || measureCount;
                    for (let c = 0; c < columnLeaves.length; c++) {
                        for (let m = 0; m < measureCount; m++) {
                            const globalM = this.displayMeasureIndices[m] ?? m;
                            const td = document.createElement("td");
                            td.textContent = this.formatValueByMeasure(this.getValueFromMapForDisplayCol(valuesMap, columnLeaves[c], globalM, totalCountLocal2), globalM);
                            this.applyCellConditionalStyle(td, valuesMap, columnLeaves[c], globalM, totalCountLocal2);
                            if (this.dataFontSize)
                                td.style.fontSize = `${this.dataFontSize}px`;
                            if (this.dataFontFamily)
                                td.style.fontFamily = this.dataFontFamily;
                            if (this.dataBold)
                                td.style.fontWeight = "bold";
                            if (this.cellColor)
                                td.style.color = this.cellColor;
                            if (this.cellBg)
                                td.style.backgroundColor = this.cellBg;
                            trTotal.appendChild(td);
                        }
                    }
                    // Place subtotal row at the top of the group's block so totals are shown first
                    rows.unshift(trTotal);
                }
                if (rows.length > 0) {
                    const th = document.createElement("th");
                    th.className = "ghm-rowheader";
                    th.title = label;
                    const toggle = document.createElement("span");
                    toggle.className = "ghm-toggle";
                    toggle.textContent = "−";
                    toggle.addEventListener("click", (ev) => {
                        ev.stopPropagation();
                        this.collapsedRowKeys.add(thisKey);
                        this.persistState();
                        this.refresh();
                    });
                    th.appendChild(toggle);
                    const txt = document.createElement("span");
                    txt.textContent = label;
                    this.applyRowHeaderStyle(txt, depth);
                    th.appendChild(txt);
                    th.rowSpan = rows.length;
                    // Ensure the first row in the group has TH placeholders up to the target depth
                    const totalDepth = this.getRowDepth(this.lastMatrix?.rows);
                    let headerCells = Array.from(rows[0].querySelectorAll('th'));
                    // If there are fewer than 'depth' header cells, prepend blanks until we can insert at the correct index
                    while (headerCells.length < Math.min(depth, totalDepth)) {
                        const pad = document.createElement("th");
                        pad.className = "ghm-rowheader";
                        this.applyRowHeaderStyle(pad, headerCells.length);
                        rows[0].insertBefore(pad, headerCells[0] || null);
                        headerCells = Array.from(rows[0].querySelectorAll('th'));
                    }
                    const targetIndex = Math.min(depth, headerCells.length);
                    const refNode = headerCells[targetIndex] || null;
                    rows[0].insertBefore(th, refNode);
                }
                return rows;
            }
        }
        // Leaf row: create a row, add deepest-level header, then data cells
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.className = "ghm-rowheader";
        th.textContent = this.nodeLabel(node) || "";
        th.rowSpan = 1;
        this.applyRowHeaderStyle(th, depth);
        tr.appendChild(th);
        // Add placeholder THs for any deeper hidden levels to keep column alignment (total columns = rowDepth)
        const totalDepth = this.getRowDepth(this.lastMatrix?.rows);
        for (let d = depth + 1; d < totalDepth; d++) {
            const ph = document.createElement("th");
            ph.className = "ghm-rowheader";
            this.applyRowHeaderStyle(ph, d);
            tr.appendChild(ph);
        }
        const valuesMap = node.values || {};
        const numericKeys = Object.keys(valuesMap).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
        for (let c = 0; c < columnLeaves.length; c++) {
            for (let m = 0; m < measureCount; m++) {
                const td = document.createElement("td");
                const v = this.getCellValueForDisplayCol(valuesMap, numericKeys, columnLeaves[c], m, measureCount);
                td.textContent = this.formatValue(v);
                tr.appendChild(td);
            }
        }
        return [tr];
    }
    refresh() {
        // Re-render using the last received matrix
        while (this.contentHost.firstChild)
            this.contentHost.removeChild(this.contentHost.firstChild);
        if (this.lastMatrix)
            this.renderMatrix(this.lastMatrix);
    }
    applyStickyOffsets(thead) {
        if (!this.container.classList.contains("ghm-sticky"))
            return;
        const compute = () => {
            const rows = Array.from(thead.rows);
            // If first row has no height yet, try again on next frame
            const firstH = rows[0] ? (rows[0].offsetHeight || rows[0].getBoundingClientRect().height) : 0;
            if (!firstH) {
                requestAnimationFrame(compute);
                return;
            }
            let top = 0;
            for (const tr of rows) {
                const h = tr.offsetHeight || tr.getBoundingClientRect().height || 0;
                const cells = Array.from(tr.cells);
                for (const cell of cells)
                    cell.style.top = `${top}px`;
                top += h;
            }
        };
        requestAnimationFrame(compute);
        window.addEventListener("resize", () => requestAnimationFrame(compute), { once: true });
    }
    persistState() {
        const widthsObj = {};
        for (const [k, v] of this.columnWidthPx.entries())
            widthsObj[k] = v;
        this.host.persistProperties({
            merge: [
                {
                    objectName: "state",
                    properties: {
                        columnWidths: JSON.stringify(widthsObj),
                        collapsedCols: JSON.stringify(Array.from(this.collapsedColKeys)),
                        collapsedRows: JSON.stringify(Array.from(this.collapsedRowKeys)),
                        repeatLabels: this.repeatLabels,
                        compactLayout: this.compactLayout
                    },
                    selector: null
                }
            ]
        });
    }
    getObjectValue(objects, objectName, propertyName, defaultValue) {
        const obj = objects && objects[objectName];
        const v = obj && obj[propertyName];
        return (v !== undefined) ? v : defaultValue;
    }
    getValueFromMapForDisplayCol(valuesMap, ref, measureIndex, measureCount) {
        if (ref.kind === "leaf") {
            const key = ref.offset * measureCount + measureIndex;
            const cell = valuesMap[key];
            return cell && cell.value != null ? cell.value : "";
        }
        else {
            // Strict host-only: for collapsed columns, render a value only when
            // the host emitted a dedicated subtotal leaf for that column group.
            const coll = ref;
            if (coll.subtotalOffset !== undefined) {
                const key = coll.subtotalOffset * measureCount + measureIndex;
                const cell = valuesMap[key];
                return (cell && cell.value != null) ? cell.value : "";
            }
            return "";
        }
    }
    getRowStyleFromValuesMap(valuesMap, totalMeasureCount) {
        const style = {};
        const keys = Object.keys(valuesMap).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
        const fetchByIndex = (mi) => {
            if (mi == null || mi < 0)
                return undefined;
            for (const k of keys) {
                if ((k % totalMeasureCount) === mi) {
                    const cell = valuesMap[k];
                    return cell && (cell.value != null ? cell.value : undefined);
                }
            }
            return undefined;
        };
        const bgVal = fetchByIndex(this.rowBackColorMeasureIndex);
        const fgVal = fetchByIndex(this.rowFontColorMeasureIndex);
        if (bgVal !== undefined)
            style.bg = String(bgVal);
        if (fgVal !== undefined)
            style.color = String(fgVal);
        return style;
    }
    applyCellConditionalStyle(td, valuesMap, ref, measureIndex, totalMeasureCount) {
        let key = null;
        if (ref.kind === 'leaf')
            key = ref.offset * totalMeasureCount + measureIndex;
        else {
            const coll = ref;
            if (coll.subtotalOffset !== undefined)
                key = coll.subtotalOffset * totalMeasureCount + measureIndex;
        }
        if (key == null)
            return;
        const cell = valuesMap[key];
        const obj = cell && cell.objects;
        const colorsObj = obj && obj.colors;
        if (colorsObj) {
            const font = colorsObj.cellColor && (colorsObj.cellColor.solid?.color || colorsObj.cellColor.value);
            const bg = colorsObj.cellBg && (colorsObj.cellBg.solid?.color || colorsObj.cellBg.value);
            if (font)
                td.style.color = font;
            if (bg)
                td.style.backgroundColor = bg;
        }
    }
    updateDebugOverlay(_ctx) { /* debug disabled */ return; }
    loadLevelStyles(matrix) {
        this.rowLevelStyles = [];
        this.colLevelStyles = [];
        const rows = matrix.rows;
        const cols = matrix.columns;
        if (rows && rows.levels) {
            rows.levels.forEach((lvl, i) => {
                const src = lvl.sources && lvl.sources[0];
                const obj = (src && src.objects && src.objects.labelStylePerLevel) || {};
                this.rowLevelStyles[i] = {
                    fontFamily: obj.fontFamily,
                    fontSize: obj.fontSize,
                    fontWeight: obj.fontWeight,
                    italic: obj.italic,
                    underline: obj.underline
                };
            });
        }
        if (cols && cols.levels) {
            cols.levels.forEach((lvl, i) => {
                const src = lvl.sources && lvl.sources[0];
                const obj = (src && src.objects && src.objects.labelStylePerLevel) || {};
                this.colLevelStyles[i] = {
                    fontFamily: obj.fontFamily,
                    fontSize: obj.fontSize,
                    fontWeight: obj.fontWeight,
                    italic: obj.italic,
                    underline: obj.underline
                };
            });
        }
    }
    computeRowInfo(rows, rowDepth) {
        if (!rows || !rows.root)
            return { count: 0 };
        if (this.repeatLabels) {
            const flat = this.collectDisplayRowsRepeat(rows.root, rowDepth);
            return { count: flat.length };
        }
        const countFrom = (node, parentKey) => {
            if (node.level === undefined && node.children && node.children.length) {
                let total = 0;
                for (const ch of node.children)
                    total += countFrom(ch, '');
                return total;
            }
            const label = this.nodeLabel(node);
            const key = [parentKey, label].filter(Boolean).join('||');
            if (node.children && node.children.length) {
                if (this.collapsedRowKeys.has(key))
                    return 1;
                let total = 0;
                for (const ch of node.children)
                    total += countFrom(ch, key);
                return total;
            }
            return 1;
        };
        return { count: countFrom(rows.root, '') };
    }
    getCollapsedRowGroupValue(node, ref, measureIndex, measureCount) {
        // Prefer host-provided row subtotal on the collapsed node itself
        const tryFromNodeValues = () => {
            const valuesMap = (node.values || {});
            if (ref.kind === "leaf") {
                const key = ref.offset * measureCount + measureIndex;
                const cell = valuesMap[key];
                if (cell && cell.value != null)
                    return cell.value;
                return undefined;
            }
            else {
                // Strict host-only for collapsed column groups: require a dedicated subtotal leaf
                const coll = ref;
                if (coll.subtotalOffset !== undefined) {
                    const key = coll.subtotalOffset * measureCount + measureIndex;
                    const cell = valuesMap[key];
                    if (cell && cell.value != null)
                        return cell.value;
                }
                return undefined;
            }
        };
        const direct = tryFromNodeValues();
        if (direct !== undefined)
            return direct;
        // Try subtotal child provided by host (isSubtotal)
        const subtotal = this.findSubtotalChild(node);
        if (subtotal) {
            const map = (subtotal.values || {});
            if (ref.kind === "leaf") {
                const key = ref.offset * measureCount + measureIndex;
                const cell = map[key];
                if (cell && cell.value != null)
                    return cell.value;
            }
            else {
                const coll = ref;
                if (coll.subtotalOffset !== undefined) {
                    const key = coll.subtotalOffset * measureCount + measureIndex;
                    const cell = map[key];
                    if (cell && cell.value != null)
                        return cell.value;
                }
            }
        }
        // Otherwise leave empty by design
        return "";
    }
    findSubtotalChild(node) {
        if (!node || !node.children)
            return null;
        for (const ch of node.children) {
            if (ch.isSubtotal)
                return ch;
        }
        return null;
    }
    aggregateAcrossRowLeaves(node, ref, measureIndex, measureCount) {
        let sum = null;
        let first = null;
        const visit = (n) => {
            if (!n.children || n.children.length === 0) {
                const valuesMap = (n.values || {});
                if (ref.kind === "leaf") {
                    const key = ref.offset * measureCount + measureIndex;
                    const cell = valuesMap[key];
                    if (cell && cell.value != null) {
                        if (typeof cell.value === "number")
                            sum = (sum ?? 0) + cell.value;
                        else if (first === null)
                            first = cell.value;
                    }
                }
                else {
                    for (let off = ref.start; off <= ref.end; off++) {
                        const key = off * measureCount + measureIndex;
                        const cell = valuesMap[key];
                        if (cell && cell.value != null) {
                            if (typeof cell.value === "number")
                                sum = (sum ?? 0) + cell.value;
                            else if (first === null)
                                first = cell.value;
                        }
                    }
                }
                return;
            }
            for (const ch of n.children)
                visit(ch);
        };
        visit(node);
        return { sum, first };
    }
    getVisibleRowDepth(rows, maxDepth) {
        if (!rows || !rows.root)
            return maxDepth;
        const depthFrom = (node, depth, parentKey) => {
            const label = this.nodeLabel(node);
            const key = [parentKey, label].filter(Boolean).join("||");
            if (!node.children || node.children.length === 0)
                return depth + 1;
            if (this.collapsedRowKeys.has(key))
                return depth + 1;
            let max = depth + 1;
            for (const ch of node.children)
                max = Math.max(max, depthFrom(ch, depth + 1, key));
            return max;
        };
        if (!rows.root.children || rows.root.children.length === 0)
            return 1;
        let m = 1;
        for (const ch of rows.root.children)
            m = Math.max(m, depthFrom(ch, 0, ""));
        return Math.min(maxDepth, m);
    }
    applyRowHeaderStyle(th, level) {
        const style = (level !== undefined && this.rowLevelStyles[level]) ? this.rowLevelStyles[level] : {};
        const fs = style.fontSize ?? this.rowHeaderFontSize;
        if (fs)
            th.style.fontSize = `${fs}px`;
        th.style.fontFamily = (style.fontFamily || this.rowHeaderFontFamily) || "";
        if (this.rowHeaderBold || style.fontWeight)
            th.style.fontWeight = style.fontWeight || "bold";
        if (style.italic)
            th.style.fontStyle = "italic";
        if (style.underline)
            th.style.textDecoration = "underline";
        if (this.rowHeaderColor)
            th.style.color = this.rowHeaderColor;
    }
    getDisplayedRowDepth(rows) {
        if (!rows || !rows.root || !rows.root.children)
            return 1;
        const depthFrom = (node, depth, parentKey) => {
            const label = this.nodeLabel(node);
            const key = [parentKey, label].filter(Boolean).join("||");
            if (this.collapsedRowKeys.has(key))
                return depth + 1;
            if (!node.children || node.children.length === 0)
                return depth + 1;
            let max = depth + 1;
            for (const ch of node.children)
                max = Math.max(max, depthFrom(ch, depth + 1, key));
            return max;
        };
        let m = 1;
        for (const ch of rows.root.children)
            m = Math.max(m, depthFrom(ch, 0, ""));
        return m;
    }
    getDisplayedColDepth(cols) {
        if (!cols || !cols.root || !cols.root.children)
            return 1;
        const depthFrom = (node, depth, path) => {
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (this.collapsedColKeys.has(key))
                return depth + 1;
            if (!node.children || node.children.length === 0)
                return depth + 1;
            let max = depth + 1;
            for (const ch of node.children)
                max = Math.max(max, depthFrom(ch, depth + 1, [...path, label]));
            return max;
        };
        let m = 1;
        for (const ch of cols.root.children)
            m = Math.max(m, depthFrom(ch, 0, []));
        return m;
    }
    collapseAll() {
        if (this.lastMatrix?.columns?.root?.children) {
            const add = (node, path) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key)
                        this.collapsedColKeys.add(key);
                    for (const ch of node.children)
                        add(ch, [...path, label]);
                }
            };
            for (const ch of this.lastMatrix.columns.root.children)
                add(ch, []);
        }
        if (this.lastMatrix?.rows?.root?.children) {
            const add = (node, path) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key)
                        this.collapsedRowKeys.add(key);
                    for (const ch of node.children)
                        add(ch, [...path, label]);
                }
            };
            for (const ch of this.lastMatrix.rows.root.children)
                add(ch, []);
        }
    }
    expandAll() {
        this.collapsedColKeys.clear();
        this.collapsedRowKeys.clear();
    }
    collectDisplayRowsRepeat(node, rowDepth, parentLabels = [], parentKey = "") {
        const rows = [];
        const label = this.nodeLabel(node);
        const thisKey = [parentKey, label].filter(Boolean).join("||");
        const labels = label ? [...parentLabels, label] : [...parentLabels];
        if (node.level === undefined && node.children && node.children.length) {
            for (const ch of node.children)
                rows.push(...this.collectDisplayRowsRepeat(ch, rowDepth, labels, thisKey));
            return rows;
        }
        if (node.children && node.children.length) {
            if (this.collapsedRowKeys.has(thisKey)) {
                const padded = [...labels];
                while (padded.length < rowDepth)
                    padded.push("");
                // Strict host-only subtotal for collapsed groups
                const subtotalChild = node.children.find(ch => ch.isSubtotal);
                const map = subtotalChild ? subtotalChild.values : {};
                const numericKeys = Object.keys(map).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
                rows.push({ labels: padded, valuesMap: map, numericKeys });
                return rows;
            }
            for (const ch of node.children)
                rows.push(...this.collectDisplayRowsRepeat(ch, rowDepth, labels, thisKey));
            return rows;
        }
        const padded = [...labels];
        while (padded.length < rowDepth)
            padded.push("");
        const valuesMap = node.values || {};
        const numericKeys = Object.keys(valuesMap).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
        rows.push({ labels: padded, valuesMap: valuesMap, numericKeys });
        return rows;
    }
    collectOutlineRowsWithTotals(root, rowDepth) {
        const rows = [];
        const tryRootTotal = () => {
            if (!this.showGrandTotal)
                return;
            const subtotalChild = root.children && root.children.find(ch => ch.isSubtotal);
            if (subtotalChild && subtotalChild.values) {
                const map = subtotalChild.values;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                rows.push({ labels, valuesMap: map, isTotal: true });
                return;
            }
            if (this.grandTotalFallback && root.values) {
                const map = root.values;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                rows.push({ labels, valuesMap: map, isTotal: true });
            }
        };
        tryRootTotal();
        const traverse = (node, depth, path, keyPath) => {
            const label = this.nodeLabel(node);
            const hasChildren = !!(node.children && node.children.length);
            const subtotalChild = hasChildren ? node.children.find(ch => ch.isSubtotal) : null;
            const hasGroupValues = hasChildren && node.values && Object.keys(node.values).length > 0;
            const map = subtotalChild ? subtotalChild.values : (hasGroupValues ? node.values : null);
            const gKey = [...keyPath, label].filter(Boolean).join("||");
            const isCollapsed = gKey && this.collapsedRowKeys.has(gKey);
            // Add totals for every group node (including top-level groups), never for leaves
            if (hasChildren && map && depth < rowDepth) {
                const labels = new Array(rowDepth).fill("");
                for (let i = 0; i < path.length; i++)
                    labels[i] = path[i];
                // When the group is collapsed, show just the group label;
                // the " Total" suffix is only meaningful when expanded.
                labels[depth] = isCollapsed ? label : `${label} Total`;
                rows.push({ labels, valuesMap: map, isTotal: true, toggleKey: gKey || undefined, depth, collapsed: !!isCollapsed });
                if (isCollapsed)
                    return; // do not emit descendants when collapsed
            }
            if (hasChildren) {
                for (const ch of node.children) {
                    if (ch.isSubtotal)
                        continue;
                    const chLabel = this.nodeLabel(ch);
                    // build next path correctly: fill parent positions, put current node's label at its depth
                    const nextPath = new Array(rowDepth).fill("");
                    for (let i = 0; i < path.length; i++)
                        nextPath[i] = path[i];
                    nextPath[depth] = label;
                    traverse(ch, depth + 1, nextPath.slice(0, depth + 1), [...keyPath, label]);
                }
                return;
            }
            // leaf
            const labels = new Array(rowDepth).fill("");
            for (let i = 0; i < path.length; i++)
                labels[i] = path[i];
            labels[depth] = label;
            rows.push({ labels, valuesMap: node.values, isTotal: false });
        };
        if (root.children) {
            for (const ch of root.children)
                traverse(ch, 0, [], []);
        }
        return rows;
    }
    // Level-wise expand/collapse helpers
    collapseRowLevel() {
        if (!this.lastMatrix?.rows?.root?.children)
            return;
        const currentDepth = this.getDisplayedRowDepth(this.lastMatrix.rows);
        if (currentDepth <= 1)
            return;
        const targetDepth = currentDepth - 1; // collapse one level globally
        const addKeysAtDepth = (node, depth, path) => {
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (!node.children || node.children.length === 0)
                return;
            if (depth === targetDepth) {
                if (!this.collapsedRowKeys.has(key))
                    this.collapsedRowKeys.add(key);
                return;
            }
            for (const ch of node.children)
                addKeysAtDepth(ch, depth + 1, [...path, label]);
        };
        for (const ch of this.lastMatrix.rows.root.children)
            addKeysAtDepth(ch, 1, []);
    }
    expandRowLevel() {
        if (this.collapsedRowKeys.size === 0)
            return;
        let minDepth = Number.MAX_SAFE_INTEGER;
        for (const key of this.collapsedRowKeys) {
            const depth = key ? key.split("||").length : 1;
            if (depth < minDepth)
                minDepth = depth;
        }
        for (const key of Array.from(this.collapsedRowKeys)) {
            const depth = key ? key.split("||").length : 1;
            if (depth === minDepth)
                this.collapsedRowKeys.delete(key);
        }
    }
    collapseColLevel() {
        if (!this.lastMatrix?.columns?.root?.children)
            return;
        const current = this.getDisplayedColDepth(this.lastMatrix.columns);
        if (current <= 1)
            return;
        const target = current - 1;
        const addAtDepth = (node, depth, path) => {
            if (!node.children || node.children.length === 0)
                return;
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (depth === target) {
                if (!this.collapsedColKeys.has(key))
                    this.collapsedColKeys.add(key);
                return;
            }
            for (const ch of node.children)
                addAtDepth(ch, depth + 1, [...path, label]);
        };
        for (const ch of this.lastMatrix.columns.root.children)
            addAtDepth(ch, 1, []);
    }
    expandColLevel() {
        if (this.collapsedColKeys.size === 0)
            return;
        let minDepth = Number.MAX_SAFE_INTEGER;
        for (const key of this.collapsedColKeys) {
            const depth = key ? key.split("||").length : 1;
            if (depth < minDepth)
                minDepth = depth;
        }
        for (const key of Array.from(this.collapsedColKeys)) {
            const depth = key ? key.split("||").length : 1;
            if (depth === minDepth)
                this.collapsedColKeys.delete(key);
        }
    }
    initializeDefaultCollapsed(matrix, enableCollapse = true) {
        if (!enableCollapse)
            return;
        // Collapse all column groups by default
        if (matrix.columns && matrix.columns.root) {
            const addColKeys = (node, path) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key)
                        this.collapsedColKeys.add(key);
                    for (const ch of node.children)
                        addColKeys(ch, [...path, label]);
                }
            };
            if (matrix.columns.root.children) {
                for (const ch of matrix.columns.root.children)
                    addColKeys(ch, []);
            }
        }
        // Collapse all row groups by default
        if (matrix.rows && matrix.rows.root) {
            const addRowKeys = (node, path) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key)
                        this.collapsedRowKeys.add(key);
                    for (const ch of node.children)
                        addRowKeys(ch, [...path, label]);
                }
            };
            if (matrix.rows.root.children) {
                for (const ch of matrix.rows.root.children)
                    addRowKeys(ch, []);
            }
        }
    }
    // no custom aggregation: collapsed groups show blank cells to avoid misrepresenting measure semantics
    buildHeaderRowsFromDisplay(displayCols, depth) {
        const rows = [];
        for (let level = 0; level < depth; level++) {
            const row = [];
            let i = 0;
            while (i < displayCols.length) {
                const col = displayCols[i];
                const label = col.kind === "leaf" ? col.labels[level] : (level < col.collapsedLevel ? col.labels[level] : (level === col.collapsedLevel ? col.labels[level] : ""));
                const key = col.keys[level] || "";
                const collapsed = col.kind === "collapsed" && level === col.collapsedLevel ? true : this.collapsedColKeys.has(key) && level < depth - 1;
                // Avoid showing a bare +/- toggle with no label when an upper level
                // header is collapsed. Only show toggles when the label is visible.
                const togglable = !!key && level < depth - 1 && !!label;
                let span = 1;
                let j = i + 1;
                while (j < displayCols.length) {
                    const nxt = displayCols[j];
                    const nLabel = nxt.kind === "leaf" ? nxt.labels[level] : (level < (nxt.collapsedLevel ?? 0) ? nxt.labels[level] : (level === (nxt.collapsedLevel ?? -1) ? nxt.labels[level] : ""));
                    const nKey = nxt.keys[level] || "";
                    const nCollapsed = nxt.kind === "collapsed" && level === (nxt.collapsedLevel ?? -1) ? true : this.collapsedColKeys.has(nKey) && level < depth - 1;
                    if (nLabel !== label || nKey !== key || nCollapsed !== collapsed)
                        break;
                    span++;
                    j++;
                }
                row.push({ label, span, key, togglable, collapsed });
                i = j;
            }
            // If a full header row yields no labels at all (e.g., due to a higher
            // level being collapsed), drop the row entirely to avoid an empty
            // header line containing only +/- icons or blanks.
            const hasAnyLabel = row.some(c => !!(c.label && c.label.trim().length > 0));
            if (hasAnyLabel)
                rows.push(row);
        }
        return rows;
    }
    beginResize(e, key) {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = this.columnWidthPx.get(key) ?? 120;
        const onMove = (ev) => {
            const delta = ev.clientX - startX;
            const newW = Math.max(40, startWidth + delta);
            this.columnWidthPx.set(key, newW);
            const cols = this.colElsByKey.get(key) || [];
            for (const c of cols)
                c.style.width = `${newW}px`;
        };
        const onUp = () => {
            window.removeEventListener("mousemove", onMove);
            window.removeEventListener("mouseup", onUp);
            this.persistState();
        };
        window.addEventListener("mousemove", onMove);
        window.addEventListener("mouseup", onUp);
    }
    getCellValueForDisplayCol(valuesMap, numericKeys, ref, measureIndex, measureCount) {
        if (ref.kind === "leaf") {
            const prefKey = ref.offset * measureCount + measureIndex;
            const preferredCell = valuesMap[prefKey];
            if (preferredCell && preferredCell.value != null)
                return preferredCell.value;
            if (numericKeys.length) {
                const idx = ref.offset * measureCount + measureIndex;
                const altKey = numericKeys[idx] ?? numericKeys[idx % numericKeys.length];
                const altCell = valuesMap[altKey];
                if (altCell && altCell.value != null)
                    return altCell.value;
            }
            return "";
        }
        else {
            // Strict host-only for collapsed columns: show value only if host supplied a dedicated
            // subtotal leaf for this column group (subtotalOffset). Otherwise blank.
            const coll = ref;
            if (coll.subtotalOffset !== undefined) {
                const key = coll.subtotalOffset * measureCount + measureIndex;
                const cell = valuesMap[key];
                return (cell && cell.value != null) ? cell.value : "";
            }
            return "";
        }
    }
    collectLeaves(node, out) {
        if (!node.children || node.children.length === 0) {
            out.push(node);
            return;
        }
        for (const child of node.children)
            this.collectLeaves(child, out);
    }
    collectRowLeaves(node, path, out) {
        const label = this.nodeLabel(node);
        const nextPath = label ? [...path, label] : [...path];
        if (!node.children || node.children.length === 0) {
            out.push({ node, label: nextPath.join(" / ") || "Total" });
            return;
        }
        for (const child of node.children)
            this.collectRowLeaves(child, nextPath, out);
    }
    countLeaves(node) {
        if (!node.children || node.children.length === 0)
            return 1;
        let n = 0;
        for (const child of node.children)
            n += this.countLeaves(child);
        return n;
    }
    nodeLabel(node) {
        if (node.levelValues && node.levelValues.length) {
            // Prefer levelValues for matrix nodes
            return String(node.levelValues.map(v => v.value).filter(v => v != null)[0] ?? "");
        }
        if (node.value != null)
            return String(node.value);
        return "";
    }
    formatValue(v) {
        if (v == null)
            return "";
        if (typeof v === "number")
            return v.toLocaleString();
        return String(v);
    }
    formatValueByMeasure(v, measureIndex) {
        if (v == null || v === "")
            return "";
        const fmt = (this.measureFormats && this.measureFormats[measureIndex]) ? this.measureFormats[measureIndex] : undefined;
        try {
            const f = valueFormatter.create({ format: fmt });
            return f.format(v);
        }
        catch {
            return String(v);
        }
    }
}
//# sourceMappingURL=visual.js.map