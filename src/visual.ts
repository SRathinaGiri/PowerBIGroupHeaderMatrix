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

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import { createTooltipServiceWrapper, ITooltipServiceWrapper, TooltipEventArgs, TooltipEnabledDataPoint } from "powerbi-visuals-utils-tooltiputils";
import * as d3 from "d3-selection";
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;

import { VisualFormattingSettingsModel } from "./settings";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

type DataView = powerbi.DataView;
type DataViewMatrix = powerbi.DataViewMatrix;
type DataViewMatrixNode = powerbi.DataViewMatrixNode;

type DisplayCol =
    | { kind: "leaf"; offset: number; labels: string[]; keys: string[]; measureIndex?: number }
    | { kind: "collapsed"; start: number; end: number; labels: string[]; keys: string[]; collapsedLevel: number; key: string; subtotalOffset?: number; measureIndex?: number };

interface SortState {
    type: "value" | "label";
    queryKeys: string; // for value sort
    measureIndex: number; // for value sort
    level: number; // for label sort
    direction: "ASC" | "DESC";
}

export class Visual implements IVisual {
    private root: HTMLElement;
    private container: HTMLElement; // scroll container
    private toolbar: HTMLElement;
    private contentHost: HTMLElement;
    private debugEl: HTMLElement;
    private debugEnabled: boolean = false;
    private table: HTMLTableElement | null = null;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    // Collapse/expand state
    private collapsedRowKeys: Set<string> = new Set();
    private collapsedColKeys: Set<string> = new Set();
    private lastMatrix: DataViewMatrix | null = null;
    private collapseInitialized: boolean = false;
    // Column sizing
    private columnWidthPx: Map<string, number> = new Map();
    private colElsByKey: Map<string, HTMLTableColElement[]> = new Map();
    private repeatLabels: boolean = false;
    private rowHeaderMinWidth: number = 160;
    private rowHeaderFontSize: number = 12;
    private rowHeaderFontFamily: string = "";
    private rowHeaderBold: boolean = false;
    private colHeaderFontSize: number = 11;
    private colHeaderFontFamily: string = "";
    private colHeaderBold: boolean = false;
    private compactLayout: boolean = true;
    private measureFormats: string[] = [];
    private displayMeasureIndices: number[] = [];
    private cellBgColorMeasureIndices: number[] = [];
    private cellFontColorMeasureIndices: number[] = [];
    private measureFontColorSettings: string[] = [];
    private measureBgColorSettings: string[] = [];
    private gridShowHorizontal: boolean = true;
    private gridShowVertical: boolean = true;
    private gridThickness: number = 1;
    private gridColor: string = "#d0d0d0";
    private zebraEnabled: boolean = false;
    private zebraOddColor: string = "#f7f7f7";
    private zebraEvenColor: string = "#ffffff";
    private zebraColEnabled: boolean = false;
    private zebraColOddColor: string = "#f7f7f7";
    private zebraColEvenColor: string = "#ffffff";
    private dataFontSize: number = 11;
    private dataFontFamily: string = "";
    private dataBold: boolean = false;
    private rowHeaderColor?: string;
    private rowHeaderBg?: string;
    private colHeaderColor?: string;
    private colHeaderBg?: string;
    private cellColor?: string;
    private cellBg?: string;
    private rowLevelStyles: Array<{fontFamily?: string; fontSize?: number; fontWeight?: string; italic?: boolean; underline?: boolean}> = [];
    private colLevelStyles: Array<{fontFamily?: string; fontSize?: number; fontWeight?: string; italic?: boolean; underline?: boolean}> = [];
    private rowLevelNames: string[] = [];
    private sortState: SortState | null = null;

    private viewMode: string = "normal";
    private currentPage: number = 1;
    private pageSize: number = 100;
    private scrollListener: ((e: Event) => void) | null = null;

    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private resizeObserver: ResizeObserver | null = null;
    // Totals behavior
    private showGrandTotal: boolean = true;
    private grandTotalFallback: boolean = false;
    private rowSubtotalsEnabled: boolean = true;
    private rowSubtotalPosition: "Top" | "Bottom" = "Bottom";
    private grandTotalPosition: "Top" | "Bottom" = "Bottom";

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipServiceWrapper = createTooltipServiceWrapper(options.host.tooltipService, options.element);

        // Root flex container
        this.root = document.createElement("div");
        this.root.className = "ghm-root";

        // Toolbar (fixed at top)
        this.toolbar = document.createElement("div");
        this.toolbar.className = "ghm-toolbar";

        // Scroll container (occupies remaining space)
        this.container = document.createElement("div");
        this.container.className = "ghm-container"; // keeps existing styles for scroll behavior

        const btnExpand = document.createElement("button");
        btnExpand.className = "ghm-btn";
        btnExpand.textContent = "➕ Expand All";
        btnExpand.setAttribute("aria-label", "Expand all groups");
        btnExpand.addEventListener("click", () => { this.expandAll(); this.persistState(); this.refresh(); });
        const btnCollapse = document.createElement("button");
        btnCollapse.className = "ghm-btn";
        btnCollapse.textContent = "➖ Collapse All";
        btnCollapse.setAttribute("aria-label", "Collapse all groups");
        btnCollapse.addEventListener("click", () => { this.collapseAll(); this.persistState(); this.refresh(); });

        const btnRowExpandLevel = document.createElement("button");
        btnRowExpandLevel.className = "ghm-btn";
        btnRowExpandLevel.title = "Expand Row Level";
        btnRowExpandLevel.setAttribute("aria-label", "Expand row hierarchy one level");
        btnRowExpandLevel.addEventListener("click", () => { this.expandRowLevel(); this.persistState(); this.refresh(); });
        // Use requested glyph-only label
        btnRowExpandLevel.textContent = "+ 𝄘";
        const btnRowCollapseLevel = document.createElement("button");
        btnRowCollapseLevel.className = "ghm-btn";
        btnRowCollapseLevel.title = "Collapse Row Level";
        btnRowCollapseLevel.setAttribute("aria-label", "Collapse row hierarchy one level");
        btnRowCollapseLevel.addEventListener("click", () => { this.collapseRowLevel(); this.persistState(); this.refresh(); });
        btnRowCollapseLevel.textContent = "- 𝄘";
        const btnColExpandLevel = document.createElement("button");
        btnColExpandLevel.className = "ghm-btn";
        btnColExpandLevel.title = "Expand Column Level";
        btnColExpandLevel.setAttribute("aria-label", "Expand column hierarchy one level");
        btnColExpandLevel.addEventListener("click", () => { this.expandColLevel(); this.persistState(); this.refresh(); });
        btnColExpandLevel.textContent = "+ ⦀";
        const btnColCollapseLevel = document.createElement("button");
        btnColCollapseLevel.className = "ghm-btn";
        btnColCollapseLevel.title = "Collapse Column Level";
        btnColCollapseLevel.setAttribute("aria-label", "Collapse column hierarchy one level");
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
            } else {
                chkRepeat.disabled = false;
            }
            this.persistState();
            this.refresh();
        });
        lblCompact.appendChild(chkCompact);
        lblCompact.appendChild(document.createTextNode(" Compact"));

        // View Mode select
        const viewModeSelect = document.createElement("select");
        viewModeSelect.id = "ghm-viewmode";
        viewModeSelect.className = "ghm-btn";
        const optNormal = document.createElement("option"); optNormal.value = "normal"; optNormal.text = "Normal";
        const optPag = document.createElement("option"); optPag.value = "pagination"; optPag.text = "Pagination";
        const optVirt = document.createElement("option"); optVirt.value = "virtualization"; optVirt.text = "Virtualization";
        viewModeSelect.appendChild(optNormal);
        viewModeSelect.appendChild(optPag);
        viewModeSelect.appendChild(optVirt);
        viewModeSelect.addEventListener("change", () => {
            this.viewMode = viewModeSelect.value;
            this.currentPage = 1;
            this.persistState();
            this.refresh();
        });

        const btnPrevPage = document.createElement("button");
        btnPrevPage.id = "ghm-btn-prevpage";
        btnPrevPage.className = "ghm-btn";
        btnPrevPage.textContent = "<";
        btnPrevPage.title = "Previous Page";
        btnPrevPage.addEventListener("click", () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.persistState();
                this.refresh();
            }
        });

        const btnNextPage = document.createElement("button");
        btnNextPage.id = "ghm-btn-nextpage";
        btnNextPage.className = "ghm-btn";
        btnNextPage.textContent = ">";
        btnNextPage.title = "Next Page";
        btnNextPage.addEventListener("click", () => {
            const rowCount = this.lastMatrix && this.lastMatrix.rows && this.lastMatrix.rows.root && this.lastMatrix.rows.root.children ? this.collectOutlineRowsWithTotals(this.lastMatrix.rows.root, this.getRowDepth(this.lastMatrix.rows)).filter((r: any) => !r.isTotal).length : 0;
            // A more robust max page relies on preprocessed rows length, but calculating it correctly via preprocessedRows is done in renderMatrix. We approximate or trust the user. For safety, we increment and allow the empty view to trigger bounds check on next render, or compute here.
            // But we can approximate total pages easily. We will do a generic bounds check in update() instead or accept unbounded.
            // Let's implement unbounded with visual feedback.
            this.currentPage++;
            this.persistState();
            this.refresh();
        });

        const lblPage = document.createElement("span");
        lblPage.id = "ghm-lbl-page";
        lblPage.style.fontSize = "12px";
        lblPage.style.margin = "0 4px";
        lblPage.textContent = "Page 1";

        // Add buttons
        this.toolbar.appendChild(btnExpand);
        const btnCollapseAll = document.createElement("button");
        btnCollapseAll.className = "ghm-btn";
        btnCollapseAll.textContent = "- All";
        btnCollapseAll.title = "Collapse All";
        btnCollapseAll.setAttribute("aria-label", "Collapse all groups");
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

        this.toolbar.appendChild(viewModeSelect);
        this.toolbar.appendChild(btnPrevPage);
        this.toolbar.appendChild(lblPage);
        this.toolbar.appendChild(btnNextPage);

        const btnDebug = document.createElement("button");
        btnDebug.className = "ghm-btn";
        btnDebug.textContent = "🧪 Debug";
        btnDebug.addEventListener("click", () => { this.debugEnabled = !this.debugEnabled; this.updateDebugOverlay(); });
        // Hide debug button in production
        try { (btnDebug as any).style.display = "none"; } catch {}
        this.toolbar.appendChild(btnDebug);

        // content host inside scroll container
        this.contentHost = document.createElement("div");
        this.container.appendChild(this.contentHost);

        // Assemble DOM
        this.root.appendChild(this.toolbar);
        this.root.appendChild(this.container);

        // debug overlay
        this.debugEl = document.createElement("div");
        this.debugEl.className = "ghm-debug";
        this.debugEl.style.display = "none";
        this.root.appendChild(this.debugEl);

        options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions) {
        const dataView: DataView | undefined = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        // Apply sticky preference if present
        const sticky = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "stickyHeaders", true);
        this.container.classList.toggle("ghm-sticky", !!sticky);
        this.repeatLabels = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "repeatLabels", false);
        this.compactLayout = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "compactLayout", true);
        if (this.compactLayout) this.repeatLabels = false;

        this.viewMode = this.getObjectValue<string>(dataView?.metadata?.objects, "state", "viewMode", "normal") || "normal";
        this.currentPage = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "currentPage", 1) || 1;
        this.pageSize = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "pageSize", 100) || 100;

        // Grand total formatting options
        this.showGrandTotal = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grandTotal", "show", true);
        this.grandTotalFallback = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grandTotal", "fallbackToRootValues", false);
        this.rowSubtotalsEnabled = this.getObjectValue<boolean>(dataView?.metadata?.objects, "subtotal", "rowSubtotals", true);
        const rowSubtotalPosRaw = this.getObjectValue<any>(dataView?.metadata?.objects, "subtotal", "rowSubtotalsType", { value: "Bottom" } as any);
        const rowSubtotalPos = typeof rowSubtotalPosRaw === "string" ? rowSubtotalPosRaw : (rowSubtotalPosRaw && (rowSubtotalPosRaw as any).value);
        this.rowSubtotalPosition = rowSubtotalPos === "Top" ? "Top" : "Bottom";
        const grandTotalPosRaw = this.getObjectValue<any>(dataView?.metadata?.objects, "grandTotal", "position", { value: "Bottom" } as any);
        const grandTotalPos = typeof grandTotalPosRaw === "string" ? grandTotalPosRaw : (grandTotalPosRaw && (grandTotalPosRaw as any).value);
        this.grandTotalPosition = grandTotalPos === "Top" ? "Top" : "Bottom";
        this.rowHeaderMinWidth = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "rowHeaderMinWidth", 160) || 160;
        this.rowHeaderFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "rowHeaderFontSize", 12) || 12;
        this.rowHeaderFontFamily = this.getObjectValue<any>(dataView?.metadata?.objects, "labels", "rowHeaderFontFamily", { value: "" } as any as string) as any as string || this.getObjectValue<string>(dataView?.metadata?.objects, "labels", "rowHeaderFontFamily", "");
        this.rowHeaderBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "rowHeaderBold", false);
        this.colHeaderFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "colHeaderFontSize", 11) || 11;
        this.colHeaderFontFamily = this.getObjectValue<any>(dataView?.metadata?.objects, "labels", "colHeaderFontFamily", { value: "" } as any as string) as any as string || this.getObjectValue<string>(dataView?.metadata?.objects, "labels", "colHeaderFontFamily", "");
        this.colHeaderBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "colHeaderBold", false);
        this.dataFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "dataFontSize", 11) || 11;
        this.dataFontFamily = this.getObjectValue<any>(dataView?.metadata?.objects, "labels", "dataFontFamily", { value: "" } as any as string) as any as string || this.getObjectValue<string>(dataView?.metadata?.objects, "labels", "dataFontFamily", "");
        this.dataBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "dataBold", false);

        const themePresetObj = this.getObjectValue<any>(dataView?.metadata?.objects, "theme", "preset", { value: "Default" } as any);
        const themePreset = typeof themePresetObj === "string" ? themePresetObj : (themePresetObj && themePresetObj.value) || "Default";

        let fallbackRowHeaderColor = "";
        let fallbackRowHeaderBg = "";
        let fallbackColHeaderColor = "";
        let fallbackColHeaderBg = "";
        let fallbackCellColor = "";
        let fallbackCellBg = "";
        let fallbackZebraOdd = "#f7f7f7";
        let fallbackZebraEven = "#ffffff";
        let fallbackZebraColOdd = "#f7f7f7";
        let fallbackZebraColEven = "#ffffff";
        let fallbackGridColor = "#d0d0d0";

        if (themePreset === "Light Blue") {
            fallbackRowHeaderColor = "#ffffff"; fallbackRowHeaderBg = "#3498db";
            fallbackColHeaderColor = "#ffffff"; fallbackColHeaderBg = "#3498db";
            fallbackCellColor = "#333333"; fallbackCellBg = "#ffffff";
            fallbackZebraOdd = "#eaf2f8"; fallbackZebraEven = "#ffffff";
            fallbackZebraColOdd = "#eaf2f8"; fallbackZebraColEven = "#ffffff";
            fallbackGridColor = "#bdc3c7";
        } else if (themePreset === "Dark") {
            fallbackRowHeaderColor = "#ffffff"; fallbackRowHeaderBg = "#2c3e50";
            fallbackColHeaderColor = "#ffffff"; fallbackColHeaderBg = "#2c3e50";
            fallbackCellColor = "#ecf0f1"; fallbackCellBg = "#34495e";
            fallbackZebraOdd = "#3d566e"; fallbackZebraEven = "#34495e";
            fallbackZebraColOdd = "#3d566e"; fallbackZebraColEven = "#34495e";
            fallbackGridColor = "#7f8c8d";
        } else if (themePreset === "Excel-like") {
            fallbackRowHeaderColor = "#000000"; fallbackRowHeaderBg = "#f3f2f1";
            fallbackColHeaderColor = "#000000"; fallbackColHeaderBg = "#f3f2f1";
            fallbackCellColor = "#000000"; fallbackCellBg = "#ffffff";
            fallbackZebraOdd = "#ffffff"; fallbackZebraEven = "#ffffff";
            fallbackZebraColOdd = "#ffffff"; fallbackZebraColEven = "#ffffff";
            fallbackGridColor = "#d4d4d4";
        } else if (themePreset === "Modern") {
            fallbackRowHeaderColor = "#2c3e50"; fallbackRowHeaderBg = "#ecf0f1";
            fallbackColHeaderColor = "#2c3e50"; fallbackColHeaderBg = "#ecf0f1";
            fallbackCellColor = "#2c3e50"; fallbackCellBg = "#ffffff";
            fallbackZebraOdd = "#fafafa"; fallbackZebraEven = "#ffffff";
            fallbackZebraColOdd = "#fafafa"; fallbackZebraColEven = "#ffffff";
            fallbackGridColor = "#e0e0e0";
        }

        this.rowHeaderColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "rowHeaderColor", { value: "" } as any)) || fallbackRowHeaderColor;
        this.rowHeaderBg = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "rowHeaderBg", { value: "" } as any)) || fallbackRowHeaderBg;
        this.colHeaderColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "colHeaderColor", { value: "" } as any)) || fallbackColHeaderColor;
        this.colHeaderBg = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "colHeaderBg", { value: "" } as any)) || fallbackColHeaderBg;
        this.cellColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "cellColor", { value: "" } as any)) || fallbackCellColor;
        this.cellBg = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "colors", "cellBg", { value: "" } as any)) || fallbackCellBg;
        this.gridShowHorizontal = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grid", "showHorizontal", true);
        this.gridShowVertical = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grid", "showVertical", true);
        this.gridThickness = this.getObjectValue<number>(dataView?.metadata?.objects, "grid", "thickness", 1) || 0;
        this.gridColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "grid", "color", { value: fallbackGridColor } as any)) || fallbackGridColor;
        this.zebraEnabled = this.getObjectValue<boolean>(dataView?.metadata?.objects, "zebra", "enabled", false);
        this.zebraOddColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "zebra", "oddColor", { value: fallbackZebraOdd } as any)) || fallbackZebraOdd;
        this.zebraEvenColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "zebra", "evenColor", { value: fallbackZebraEven } as any)) || fallbackZebraEven;
        this.zebraColEnabled = this.getObjectValue<boolean>(dataView?.metadata?.objects, "zebraColumns", "enabled", false);
        this.zebraColOddColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "zebraColumns", "oddColor", { value: fallbackZebraColOdd } as any)) || fallbackZebraColOdd;
        this.zebraColEvenColor = this.parseColor(this.getObjectValue<any>(dataView?.metadata?.objects, "zebraColumns", "evenColor", { value: fallbackZebraColEven } as any)) || fallbackZebraColEven;
        // Reflect toolbar checkbox states
        const chkRepeatEl = this.toolbar.querySelector('#ghm-repeat') as HTMLInputElement | null;
        if (chkRepeatEl) {
            chkRepeatEl.checked = this.repeatLabels && !this.compactLayout;
            chkRepeatEl.disabled = !!this.compactLayout;
        }
        const chkCompactEl = this.toolbar.querySelector('#ghm-compact') as HTMLInputElement | null;
        if (chkCompactEl) chkCompactEl.checked = !!this.compactLayout;

        const viewModeSelectEl = this.toolbar.querySelector('#ghm-viewmode') as HTMLSelectElement | null;
        if (viewModeSelectEl) viewModeSelectEl.value = this.viewMode;

        const btnPrevPage = this.toolbar.querySelector('#ghm-btn-prevpage') as HTMLButtonElement | null;
        const btnNextPage = this.toolbar.querySelector('#ghm-btn-nextpage') as HTMLButtonElement | null;
        const lblPage = this.toolbar.querySelector('#ghm-lbl-page') as HTMLElement | null;

        if (btnPrevPage) btnPrevPage.style.display = this.viewMode === "pagination" ? "inline-block" : "none";
        if (btnNextPage) btnNextPage.style.display = this.viewMode === "pagination" ? "inline-block" : "none";
        if (lblPage) {
             lblPage.style.display = this.viewMode === "pagination" ? "inline-block" : "none";
             lblPage.textContent = `Page ${this.currentPage}`;
        }

        // Load persisted widths and collapsed sets
        const widthsJson = this.getObjectValue<string>(dataView?.metadata?.objects, "state", "columnWidths", "");
        if (widthsJson) {
            try {
                const w = JSON.parse(widthsJson) as { [k: string]: number };
                this.columnWidthPx = new Map(Object.entries(w));
            } catch {}
        }
        const colsJson = this.getObjectValue<string>(dataView?.metadata?.objects, "state", "collapsedCols", "");
        if (colsJson) {
            try { this.collapsedColKeys = new Set(JSON.parse(colsJson)); this.collapseInitialized = true; } catch {}
        }
        const rowsJson = this.getObjectValue<string>(dataView?.metadata?.objects, "state", "collapsedRows", "");
        if (rowsJson) {
            try { this.collapsedRowKeys = new Set(JSON.parse(rowsJson)); this.collapseInitialized = true; } catch {}
        }
        // If no measure is bound, force all expanded to show full structure
        const noMeasures = !(dataView && dataView.matrix && dataView.matrix.valueSources && dataView.matrix.valueSources.length > 0);
        if (noMeasures) {
            this.collapsedColKeys.clear();
            this.collapsedRowKeys.clear();
            this.collapseInitialized = true;
        }

        // Clear only content (preserve toolbar)
        while (this.contentHost.firstChild) this.contentHost.removeChild(this.contentHost.firstChild);

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
        } catch (e) {
            this.renderPlaceholder("Unable to render matrix");
            console.error(e);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // Enumerates per-level subtotal toggles so the host can attach properties to individual row/column levels
    public enumerateObjectInstances(options: powerbi.EnumerateVisualObjectInstancesOptions): powerbi.VisualObjectInstanceEnumeration {
        const enumeration: powerbi.VisualObjectInstance[] = [];
        if (!this.lastMatrix) return enumeration;

        if (options.objectName === "subtotalPerLevel") {
            const rows = this.lastMatrix.rows;
            const cols = this.lastMatrix.columns;
            // Row levels
            if (rows && rows.levels && rows.levels.length) {
                rows.levels.forEach((lvl, i) => {
                    const src = lvl.sources && lvl.sources[0];
                    if (!src) return;
                    const displayName = src.displayName || `Row level ${i+1}`;
                    const selector: any = src.queryName ? { metadata: src.queryName } : null;
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
                    if (!src) return;
                    const displayName = src.displayName || `Column level ${i+1}`;
                    const selector: any = src.queryName ? { metadata: src.queryName } : null;
                    enumeration.push({
                        objectName: options.objectName,
                        displayName: `Column: ${displayName}`,
                        properties: { levelSubtotalEnabled: true },
                        selector
                    });
                });
            }
        }
        if (options.objectName === "measureColors") {
            const vsArr = (this.lastMatrix.valueSources || []) as any[];
            vsArr.forEach((vs, i) => {
                const roles = (vs as any).roles || {};
                if (!roles.measure) return;
                const displayName = (vs as any).displayName || (vs as any).queryName || `Measure ${i + 1}`;
                const selector: any = (vs as any).queryName ? { metadata: (vs as any).queryName } : null;
                const obj = (vs as any).objects || {};
                const font = this.parseColor(obj.measureColors && (obj.measureColors as any).cellFontColor) || "";
                const bg = this.parseColor(obj.measureColors && (obj.measureColors as any).cellBgColor) || "";
                enumeration.push({
                    objectName: options.objectName,
                    displayName,
                    properties: {
                        cellFontColor: font,
                        cellBgColor: bg
                    },
                    selector
                });
            });
        }
        return enumeration;
    }

    private renderPlaceholder(text: string) {
        const el = document.createElement("div");
        el.className = "ghm-placeholder";
        el.textContent = text;
        this.contentHost.appendChild(el);
    }

    private renderMatrix(matrix: DataViewMatrix) {
        const columns = matrix.columns;
        const rows = matrix.rows;

        if (!columns || !columns.root || !columns.root.children || columns.root.children.length === 0) {
            this.renderPlaceholder("No column hierarchy provided");
            return;
        }

        const table = document.createElement("table");
        table.className = "ghm-table";
        table.style.borderCollapse = "collapse";
        // Apply a global font if only one family provided
        if (this.rowHeaderFontFamily && !this.colHeaderFontFamily) table.style.fontFamily = this.rowHeaderFontFamily;
        if (this.colHeaderFontFamily && !this.rowHeaderFontFamily) table.style.fontFamily = this.colHeaderFontFamily;
        const colgroup = document.createElement("colgroup");
        const thead = document.createElement("thead");
        const tbody = document.createElement("tbody");

        // Measures handling: compute display measures vs style measures
        const totalMeasureCount = (matrix.valueSources && matrix.valueSources.length) ? matrix.valueSources.length : 0;
        this.displayMeasureIndices = [];
        this.cellBgColorMeasureIndices = [];
        this.cellFontColorMeasureIndices = [];
        this.measureFontColorSettings = [];
        this.measureBgColorSettings = [];
        const vsArr = matrix.valueSources || [];
        for (let i = 0; i < vsArr.length; i++) {
            const roles: any = (vsArr[i] as any).roles || {};
            if (roles.measure) this.displayMeasureIndices.push(i);
            if (roles.cellBgColor) this.cellBgColorMeasureIndices.push(i);
            if (roles.cellFontColor) this.cellFontColorMeasureIndices.push(i);
            const objects = (vsArr[i] as any).objects || {};
            this.measureFontColorSettings[i] = this.parseColor(objects.measureColors && (objects.measureColors as any).cellFontColor) || "";
            this.measureBgColorSettings[i] = this.parseColor(objects.measureColors && (objects.measureColors as any).cellBgColor) || "";
        }
        if (!this.displayMeasureIndices.length && vsArr.length) {
            // Default to all non-style measures if the role is missing
            this.displayMeasureIndices = vsArr
                .map((_, idx) => idx)
                .filter(idx => {
                    const roles: any = (vsArr[idx] as any).roles || {};
                    return roles.measure || (!roles.cellBgColor && !roles.cellFontColor);
                });
            if (!this.displayMeasureIndices.length) {
                this.displayMeasureIndices = vsArr.map((_, idx) => idx);
            }
        }
        const columnsHaveLevels = !!(columns.levels && columns.levels.length);
        // Check if ANY level in columns is a measure. Power BI can place measures at any level (e.g. nested).
        const measuresOnColumns = !!(columns.levels && columns.levels.some(lvl => lvl.sources && lvl.sources.some((s: any) => s && (s as any).isMeasure)));

        // Determine columns to display (compress collapsed groups to a single column)
        const colDepth = this.getColumnDepth(columns);
        // If measuresOnColumns, displayMeasureCount is 1, so totalMeasureCount is relevant for expansion
        const displayCols = this.computeDisplayColumns(columns.root, colDepth, measuresOnColumns, totalMeasureCount);
        // Build column header rows from the display list
        const headerRows = this.buildHeaderRowsFromDisplay(displayCols, colDepth);
        // If measures are already on the column axis, skip multiplying by measureCount
        const displayMeasureCount = measuresOnColumns ? 1 : Math.max(1, this.displayMeasureIndices.length);
        const resolveMeasureIndex = (ref: DisplayCol, displayIdx: number): number => {
            if (measuresOnColumns) {
                if (ref.measureIndex !== undefined) return ref.measureIndex;
                // Fallback (risky if collapsed group without measureIndex)
                if (ref.kind === "leaf") return ref.offset % totalMeasureCount;
                return 0;
            }
            return this.displayMeasureIndices[displayIdx] ?? displayIdx;
        };
        const measureLabels = measuresOnColumns
            ? displayCols.map((c, idx) => c.labels[c.labels.length - 1] || `Measure ${idx + 1}`)
            : this.displayMeasureIndices.map((i, idx) => String((vsArr[i] as any).displayName || (vsArr[i] as any).queryName || `Measure ${idx + 1}`));
        this.measureFormats = (vsArr || []).map(m => String((m as any).format || ""));
        const totalCountForKeys = Math.max(1, totalMeasureCount);
        if (columnsHaveLevels && displayMeasureCount > 1) {
            for (const row of headerRows) {
                for (const cell of row) cell.span = (cell.span || 1) * displayMeasureCount;
            }
        }

        // Determine number of row header levels (depth)
        const rowDepth = this.getRowDepth(rows);
        // Hide deeper row header columns when levels are collapsed, even in Repeat Labels mode
        const visibleRowDepth = this.getVisibleRowDepth(this.lastMatrix?.rows, rowDepth);
        const rowHeaderCols = (rowDepth > 0) ? (this.compactLayout ? 1 : visibleRowDepth) : 0;
        // Include the thin resizer row as part of the sticky header block
        const headerDepth = headerRows.length + (displayMeasureCount > 1 ? 1 : 0) + 1;

        // Calculate cumulative widths for row header columns to sticky-position them correctly
        const rowHeaderWidths: number[] = [];
        if (rowHeaderCols > 0) {
             if (this.compactLayout) {
                rowHeaderWidths.push(this.rowHeaderMinWidth);
             } else {
                rowHeaderWidths.push(this.rowHeaderMinWidth);
                for (let k = 1; k < rowHeaderCols; k++) rowHeaderWidths.push(120);
             }
        }

        headerRows.forEach((rowCells, i) => {
            const tr = document.createElement("tr");
            if (i === 0 && rowHeaderCols > 0) {
                let accumulatedCornerLeft = 0;
                for (let r = 0; r < rowHeaderCols; r++) {
                    const corner = document.createElement("th");
                    corner.className = "ghm-corner";
                    corner.rowSpan = headerDepth;

                    const width = rowHeaderWidths[r] || 120;
                    corner.style.width = `${width}px`;
                    corner.style.minWidth = `${width}px`;
                    corner.style.maxWidth = `${width}px`;

                    if (!this.compactLayout) {
                        corner.style.left = `${accumulatedCornerLeft}px`;
                        accumulatedCornerLeft += width;
                    } else {
                        corner.style.left = "0px";
                    }

                    // Text content: Compact layout uses "Rows" or first level; Tabular uses specific level names
                    const levelName = this.compactLayout ? (this.rowLevelNames[0] || "Rows") : (this.rowLevelNames[r] || "");
                    corner.textContent = levelName;

                    this.attachSortHandler(corner, "", -1, r, "label");

                    this.applyGridBorder(corner, true);
                    tr.appendChild(corner);
                }
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
                    toggle.setAttribute("role", "button");
                    toggle.setAttribute("aria-label", cell.collapsed ? "Expand column group" : "Collapse column group");
                    toggle.tabIndex = 0;
                    const handler = (ev: Event) => {
                        ev.stopPropagation();
                        if (cell.collapsed) this.collapsedColKeys.delete(cell.key);
                        else this.collapsedColKeys.add(cell.key);
                        this.persistState();
                        this.refresh();
                    };
                    toggle.addEventListener("click", handler);
                    toggle.addEventListener("keydown", (e) => {
                        if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            handler(e);
                        }
                    });
                    th.appendChild(toggle);
                }
                const txt = document.createElement("span");
                txt.textContent = cell.label;
                const ch = this.colLevelStyles[i] || {};
                const fs = ch.fontSize ?? this.colHeaderFontSize; if (fs) txt.style.fontSize = `${fs}px`;
                txt.style.fontFamily = (ch.fontFamily || this.colHeaderFontFamily) || "";
                if (this.colHeaderBold || ch.fontWeight) txt.style.fontWeight = (ch.fontWeight || "bold");
                if (ch.italic) txt.style.fontStyle = "italic";
                if (ch.underline) txt.style.textDecoration = "underline";
                if (this.colHeaderColor) th.style.color = this.colHeaderColor;
                if (this.colHeaderBg) th.style.backgroundColor = this.colHeaderBg;
                this.applyGridBorder(th, true);
                th.appendChild(txt);
                if (cell.span && cell.span > 1) th.colSpan = cell.span;

                // Add sort handler
                if (displayMeasureCount <= 1 && cell.isLeafHeader) {
                    this.attachSortHandler(th, cell.queryKeys, 0);
                }

                tr.appendChild(th);
            }
            thead.appendChild(tr);
        });

        if (displayMeasureCount > 1) {
            const tr = document.createElement("tr");
            for (const ref of displayCols) {
                for (let m = 0; m < displayMeasureCount; m++) {
                    const th = document.createElement("th");
                    th.className = "ghm-colheader";
                    th.textContent = measureLabels[m] ?? `M${m + 1}`;
                    if (this.colHeaderColor) th.style.color = this.colHeaderColor;
                    if (this.colHeaderBg) th.style.backgroundColor = this.colHeaderBg;
                    this.applyGridBorder(th, true);

                    // Add sort handler
                    const queryKeys = ref.kind === "leaf" ? ref.keys.join("||") : ref.key;
                    this.attachSortHandler(th, queryKeys, m);

                    tr.appendChild(th);
                }
            }
            thead.appendChild(tr);
        }

        // Resizer row (always present so handles align with columns)
        const resizerRow = document.createElement("tr");
        resizerRow.className = "ghm-resizers-row";
        // Prepend row-header placeholders so cell count matches total columns
        let accumulatedResizerLeft = 0;
        for (let r = 0; r < rowHeaderCols; r++) {
            const th = document.createElement("th");
            th.className = "ghm-resizer-cell";
            // In tabular layout, sticky row headers must stack horizontally
            if (!this.compactLayout) {
                th.style.left = `${accumulatedResizerLeft}px`;
                accumulatedResizerLeft += (r === 0 ? this.rowHeaderMinWidth : 120);
            } else {
                th.style.left = "0px";
            }
            resizerRow.appendChild(th);
        }
        const resizerRefs = displayCols;
        for (const ref of resizerRefs) {
            for (let m = 0; m < displayMeasureCount; m++) {
                const th = document.createElement("th");
                th.className = "ghm-resizer-cell";
                const key = (ref.kind === "leaf") ? ref.keys.join("||") : ref.key;
                const handle = document.createElement("div");
                handle.className = "ghm-resizer";
                handle.title = "Drag to resize column";
                handle.addEventListener("mousedown", (e) => this.beginResize(e as MouseEvent, key));
                th.appendChild(handle);
                this.applyGridBorder(th, true);
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
            if (!this.compactLayout && r < visibleRowDepth) col.style.width = (r === 0 ? `${this.rowHeaderMinWidth}px` : "120px");
            else if (this.compactLayout && r === 0) col.style.width = `${this.rowHeaderMinWidth}px`;
            else col.style.width = "0px";
            colgroup.appendChild(col);
        }
        for (const ref of columnLeaves) {
            const key = (ref.kind === "leaf") ? ref.keys.join("||") : ref.key;
            const width = this.columnWidthPx.get(key) ?? 120;
            for (let m = 0; m < displayMeasureCount; m++) {
                const col = document.createElement("col");
                col.style.width = `${width}px`;
                colgroup.appendChild(col);
                if (!this.colElsByKey.has(key)) this.colElsByKey.set(key, []);
                this.colElsByKey.get(key)!.push(col);
            }
        }

        const colLeafCount = columnLeaves.length;

        // Apply sorting if active
        if (this.sortState && rows && rows.root && rows.root.children) {
            if (this.sortState.type === "label") {
                 this.sortRowsRecursive(rows.root.children, -1, this.sortState.direction, "label", this.sortState.level, 0);
            } else {
                const sortCol = columnLeaves.find(c => {
                    const k = c.kind === "leaf" ? c.keys.join("||") : c.key;
                    return k === this.sortState!.queryKeys;
                });
                if (sortCol) {
                    let valueKey = -1;
                    if (sortCol.kind === "leaf") {
                        valueKey = sortCol.offset * totalCountForKeys + this.sortState.measureIndex;
                    } else {
                         const coll = sortCol as any as { subtotalOffset?: number };
                         if (coll.subtotalOffset !== undefined) {
                            valueKey = coll.subtotalOffset * totalCountForKeys + this.sortState.measureIndex;
                         }
                    }
                    if (valueKey !== -1) {
                        this.sortRowsRecursive(rows.root.children, valueKey, this.sortState.direction, "value", -1, 0);
                    }
                }
            }
        }

        // Do not inject a separate "Grand Total" row here.
        // Body rows (including the root grand total when applicable) are generated
        // exclusively by collectOutlineRowsWithTotals() below. This avoids any
        // chance of duplicating the total row.

        // Build body rows
        let rowsToRender: any[] = [];
        if (rows && rows.root && rows.root.children && rows.root.children.length) {
            rowsToRender = this.collectOutlineRowsWithTotals(rows.root, rowDepth);
        } else {
            // Single row mode
            const tr = document.createElement("tr");
            const valuesMap = rows && rows.root ? rows.root.values || {} : {};
            const numericKeys = Object.keys(valuesMap as any).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
            for (let c = 0; c < colLeafCount; c++) {
                for (let m = 0; m < displayMeasureCount; m++) {
                    const globalM = resolveMeasureIndex(columnLeaves[c], m);
                    const td = document.createElement("td");
                    const v = this.getCellValueForDisplayCol(valuesMap as any, numericKeys, columnLeaves[c], globalM, totalCountForKeys, measuresOnColumns);
                    td.textContent = this.formatValueByMeasure(v, globalM);
                    this.applyCellConditionalStyle(td, valuesMap as any, columnLeaves[c], globalM, totalCountForKeys, measuresOnColumns);
                    if (this.dataFontSize) td.style.fontSize = `${this.dataFontSize}px`;
                    if (this.dataFontFamily) td.style.fontFamily = this.dataFontFamily;
                    if (this.dataBold) td.style.fontWeight = "bold";
                    if (this.cellColor) td.style.color = this.cellColor;
                    let baseBg = "";
                    if (this.zebraColEnabled) {
                        const colIdx = c * displayMeasureCount + m;
                        baseBg = (colIdx % 2 === 0) ? this.zebraColEvenColor : this.zebraColOddColor;
                    }
                    if (!baseBg && this.cellBg) baseBg = this.cellBg;
                    if (baseBg) td.style.backgroundColor = baseBg;
                    this.applyGridBorder(td, false);
                    tr.appendChild(td);
                }
            }
            tbody.appendChild(tr);
        }

        const lastShownRowLabels: Array<string | null> = new Array(rowDepth).fill(null);

        // Pre-process rows for sliceability (pagination and virtualization)
        let bodyRowIndex = 0;
        const preprocessedRows = rowsToRender.map((rowInfoRaw) => {
             const rowInfo = { ...rowInfoRaw } as any;
             const valuesMapForRow = rowInfo.valuesMap || {} as any;
             const rowMeasureBg = this.getRowMeasureBg(valuesMapForRow as any, totalCountForKeys);
             let labelsToUse: string[] = rowInfo.labels as string[];

             if (!this.compactLayout && !this.repeatLabels && !rowInfo.isTotal && this.rowSubtotalsEnabled) {
                 if (this.rowSubtotalPosition === "Bottom") {
                     const newLabels = new Array(rowDepth).fill("");
                     for (let lvl = 0; lvl < rowDepth; lvl++) {
                         const lbl = (rowInfo.labels && rowInfo.labels[lvl]) ? rowInfo.labels[lvl] : "";
                         if (!lbl) { newLabels[lvl] = ""; continue; }
                         if (lastShownRowLabels[lvl] === lbl) {
                             newLabels[lvl] = "";
                         } else {
                             newLabels[lvl] = lbl;
                             lastShownRowLabels[lvl] = lbl;
                             for (let deeper = lvl + 1; deeper < rowDepth; deeper++) lastShownRowLabels[deeper] = null;
                         }
                     }
                     labelsToUse = newLabels;
                 } else {
                     const newLabels = new Array(rowDepth).fill("");
                     let lastIdx = -1;
                     for (let k = rowDepth - 1; k >= 0; k--) { if (rowInfo.labels && rowInfo.labels[k]) { lastIdx = k; break; } }
                     if (lastIdx >= 0) newLabels[lastIdx] = rowInfo.labels[lastIdx];
                     labelsToUse = newLabels;
                 }
             }
             if (rowInfo.isTotal) {
                 for (let k = 0; k < lastShownRowLabels.length; k++) lastShownRowLabels[k] = null;
             }

             rowInfo.labelsToUse = labelsToUse;
             rowInfo.rowMeasureBg = rowMeasureBg;
             if (!rowInfo.isTotal) {
                 rowInfo.bodyRowIndex = bodyRowIndex;
                 bodyRowIndex++;
             }
             return rowInfo;
        });

        const renderBody = (rowsToRenderSlice: any[]) => {
            while (tbody.firstChild) tbody.removeChild(tbody.firstChild);

            for (const rowInfo of rowsToRenderSlice) {
                const labelsToUse = rowInfo.labelsToUse;
                const rowMeasureBg = rowInfo.rowMeasureBg;

                const tr = document.createElement("tr");
                if (rowInfo.isTotal) tr.className = "ghm-totalrow";
                if (this.zebraEnabled && !rowInfo.isTotal) {
                    const zebraColor = (rowInfo.bodyRowIndex % 2 === 0) ? this.zebraEvenColor : this.zebraOddColor;
                    if (!rowMeasureBg && zebraColor) tr.style.backgroundColor = zebraColor;
                }
            const rowStyleBg = rowMeasureBg || this.rowHeaderBg || "";
            const rowStyleColor = this.rowHeaderColor || "";

            if (rowHeaderCols > 0) {
                if (this.compactLayout) {
                    const th = document.createElement("th");
                    th.className = "ghm-rowheader";
                    let txt = "";
                    let lastIdx = -1;
                    for (let k = rowDepth - 1; k >= 0; k--) { if (labelsToUse && labelsToUse[k]) { txt = labelsToUse[k]; lastIdx = k; break; } }
                    th.textContent = txt || "";
                    this.applyRowHeaderStyle(th, 0);
                    if (rowStyleBg) th.style.backgroundColor = rowStyleBg;
                    if (rowStyleColor) th.style.color = rowStyleColor;
                    this.applyGridBorder(th, true);
                    const depthIndent = (rowInfo as any).isTotal && (rowInfo as any).depth !== undefined
                        ? Math.max(0, Math.min(rowDepth - 1, (rowInfo as any).depth))
                        : Math.max(0, lastIdx);
                    const basePad = 8, step = 14;
                    th.style.paddingLeft = `${basePad + step * depthIndent}px`;

                    if ((rowInfo as any).toggleKey) {
                        const toggle = document.createElement("span");
                        toggle.className = "ghm-toggle";
                        const collapsed = !!(rowInfo as any).collapsed;
                        toggle.textContent = collapsed ? "+" : "-";
                        toggle.title = collapsed ? "Expand group" : "Collapse group";
                        toggle.setAttribute("role", "button");
                        toggle.setAttribute("aria-label", collapsed ? "Expand row group" : "Collapse row group");
                        toggle.tabIndex = 0;
                        (toggle as any).style.marginRight = "6px";
                        const handler = (ev: Event) => {
                            ev.stopPropagation();
                            const key = String((rowInfo as any).toggleKey);
                            if (collapsed) this.collapsedRowKeys.delete(key); else this.collapsedRowKeys.add(key);
                            this.persistState();
                            this.refresh();
                        };
                        toggle.addEventListener("click", handler);
                        toggle.addEventListener("keydown", (e) => {
                            if (e.key === "Enter" || e.key === " ") {
                                e.preventDefault();
                                handler(e);
                            }
                        });
                        th.prepend(toggle);
                    }
                    if (this.rowHeaderBg) th.style.backgroundColor = this.rowHeaderBg;
                    tr.appendChild(th);
                } else {
                    const toggleLevel = Math.max(0, Math.min(rowHeaderCols - 1, (rowInfo as any).depth ?? 0));
                    let accumulatedLeft = 0;
                    const firstWidth = (this as any).rowHeaderMinWidth || 160;
                    const otherWidth = 120;

                    for (let lvl = 0; lvl < rowHeaderCols; lvl++) {
                        const th = document.createElement("th");
                        th.className = "ghm-rowheader";
                        th.textContent = (labelsToUse && labelsToUse[lvl]) || "";

                        const currentWidth = (lvl === 0 ? firstWidth : otherWidth);
                        th.style.left = `${accumulatedLeft}px`;
                        th.style.width = `${currentWidth}px`;
                        th.style.minWidth = `${currentWidth}px`;
                        th.style.maxWidth = `${currentWidth}px`;
                        accumulatedLeft += currentWidth;

                        this.applyRowHeaderStyle(th, lvl);
                        if (rowStyleBg) th.style.backgroundColor = rowStyleBg;
                        if (rowStyleColor) th.style.color = rowStyleColor;
                        if (this.rowHeaderBg) th.style.backgroundColor = this.rowHeaderBg;
                        this.applyGridBorder(th, true);

                        const toggleKeyAtLevel = (rowInfo as any).toggles ? (rowInfo as any).toggles[lvl] : ((rowInfo as any).toggleKey && lvl === toggleLevel ? (rowInfo as any).toggleKey : null);

                        if (toggleKeyAtLevel) {
                            const toggle = document.createElement("span");
                            toggle.className = "ghm-toggle";
                            const isKeyCollapsed = this.collapsedRowKeys.has(toggleKeyAtLevel);

                            toggle.textContent = isKeyCollapsed ? "+" : "-";
                            toggle.title = isKeyCollapsed ? "Expand group" : "Collapse group";
                            toggle.setAttribute("role", "button");
                            toggle.setAttribute("aria-label", isKeyCollapsed ? "Expand row group" : "Collapse row group");
                            toggle.tabIndex = 0;
                            const handler = (ev: Event) => {
                                ev.stopPropagation();
                                const key = String(toggleKeyAtLevel);
                                if (isKeyCollapsed) this.collapsedRowKeys.delete(key); else this.collapsedRowKeys.add(key);
                                this.persistState();
                                this.refresh();
                            };
                            toggle.addEventListener("click", handler);
                            toggle.addEventListener("keydown", (e) => {
                                if (e.key === "Enter" || e.key === " ") {
                                    e.preventDefault();
                                    handler(e);
                                }
                            });
                            th.prepend(toggle);
                        }
                        tr.appendChild(th);
                    }
                }
            }

            const valuesMap = (rowInfo as any).valuesMap || {} as any;
            const numericKeys = Object.keys(valuesMap as any).map(k => +k).filter(k => !Number.isNaN(k)).sort((a,b)=>a-b);
            for (let c = 0; c < colLeafCount; c++) {
                for (let m = 0; m < displayMeasureCount; m++) {
                    const globalM = resolveMeasureIndex(columnLeaves[c], m);
                    const td = document.createElement("td");
                    const v = this.getCellValueForDisplayCol(valuesMap as any, numericKeys, columnLeaves[c], globalM, totalCountForKeys, measuresOnColumns);
                    td.textContent = this.formatValueByMeasure(v, globalM);
                    this.applyCellConditionalStyle(td, valuesMap as any, columnLeaves[c], globalM, totalCountForKeys, measuresOnColumns);
                    if (this.dataFontSize) td.style.fontSize = `${this.dataFontSize}px`;
                    if (this.dataFontFamily) td.style.fontFamily = this.dataFontFamily;
                    if (this.dataBold) td.style.fontWeight = "bold";
                    if (this.cellColor) td.style.color = this.cellColor;
                    let baseBg = "";
                    if (this.zebraColEnabled) {
                        const colIdx = c * displayMeasureCount + m;
                        baseBg = (colIdx % 2 === 0) ? this.zebraColEvenColor : this.zebraColOddColor;
                    }
                    if (!baseBg && this.cellBg) baseBg = this.cellBg;
                    if (baseBg) td.style.backgroundColor = baseBg;
                    this.applyGridBorder(td, false);

                    // Interaction (Multi-row mode)
                    if (!rowInfo.isTotal) {
                        const rowNode = (rowInfo as any).node;
                        if (rowNode) {
                            const selectionId = this.createSelectionId(rowNode, undefined, undefined);

                            // Selection state opacity
                            const hasSelection = this.selectionManager.hasSelection();
                            const isSelected = this.selectionManager.getSelectionIds().some(id => (id as any).equals(selectionId));
                            if (hasSelection && !isSelected) {
                                td.style.opacity = "0.5";
                            } else {
                                td.style.opacity = "1";
                            }

                            td.addEventListener("click", (e) => {
                                this.selectionManager.select(selectionId, e.ctrlKey || e.metaKey).then(() => {
                                    this.refresh();
                                });
                                e.stopPropagation();
                            });
                            td.addEventListener("contextmenu", (e) => {
                                this.selectionManager.showContextMenu(selectionId, {x: e.clientX, y: e.clientY});
                                e.preventDefault();
                            });

                            // Tooltip
                            this.tooltipServiceWrapper.addTooltip(d3.select(td) as any, (tooltipEvent: TooltipEventArgs<TooltipEnabledDataPoint>) => {
                                return this.getTooltipData(v, rowNode, undefined, globalM);
                            }, () => selectionId);
                        }
                    }

                    tr.appendChild(td);
                }
            }
            tbody.appendChild(tr);
        }
    }; // end of renderBody

    if (this.scrollListener) {
        this.container.removeEventListener("scroll", this.scrollListener);
        this.scrollListener = null;
    }

    if (this.viewMode === "pagination") {
        const start = (this.currentPage - 1) * this.pageSize;
        const end = start + this.pageSize;
        renderBody(preprocessedRows.slice(start, end));
        const lblPage = this.toolbar.querySelector('#ghm-lbl-page') as HTMLElement | null;
        if (lblPage) {
            const totalPages = Math.ceil(preprocessedRows.length / this.pageSize) || 1;
            if (this.currentPage > totalPages && totalPages > 0) {
                this.currentPage = totalPages;
                const newStart = (this.currentPage - 1) * this.pageSize;
                const newEnd = newStart + this.pageSize;
                renderBody(preprocessedRows.slice(newStart, newEnd));
            }
            lblPage.textContent = `Page ${this.currentPage} of ${totalPages}`;

            const btnNextPage = this.toolbar.querySelector('#ghm-btn-nextpage') as HTMLButtonElement | null;
            if (btnNextPage) btnNextPage.disabled = this.currentPage >= totalPages;
            const btnPrevPage = this.toolbar.querySelector('#ghm-btn-prevpage') as HTMLButtonElement | null;
            if (btnPrevPage) btnPrevPage.disabled = this.currentPage <= 1;
        }
    } else if (this.viewMode === "virtualization") {
        const ROW_HEIGHT = 25; // Estimated
        const renderVirtual = () => {
            const scrollTop = this.container.scrollTop;
            const clientHeight = this.container.clientHeight;

            // Add buffer rows to avoid rapid re-rendering
            const buffer = 10;
            const startIdx = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - buffer);
            const visibleCount = Math.ceil(clientHeight / ROW_HEIGHT) + (buffer * 2);
            const endIdx = Math.min(preprocessedRows.length, startIdx + visibleCount);

            const sliced = preprocessedRows.slice(startIdx, endIdx);
            renderBody(sliced);

            const topSpacerHeight = startIdx * ROW_HEIGHT;
            const bottomSpacerHeight = (preprocessedRows.length - endIdx) * ROW_HEIGHT;

            if (topSpacerHeight > 0) {
                const trTop = document.createElement("tr");
                trTop.style.height = `${topSpacerHeight}px`;
                trTop.className = "ghm-virtual-spacer-top";
                tbody.insertBefore(trTop, tbody.firstChild);
            }
            if (bottomSpacerHeight > 0) {
                const trBottom = document.createElement("tr");
                trBottom.style.height = `${bottomSpacerHeight}px`;
                trBottom.className = "ghm-virtual-spacer-bottom";
                tbody.appendChild(trBottom);
            }
        };

        let isTicking = false;
        this.scrollListener = () => {
            if (!isTicking) {
                requestAnimationFrame(() => {
                    renderVirtual();
                    isTicking = false;
                });
                isTicking = true;
            }
        };
        this.container.addEventListener("scroll", this.scrollListener, { passive: true });
        // Initial render
        renderVirtual();
    } else {
        renderBody(preprocessedRows);
    }

        table.appendChild(colgroup);
        table.appendChild(thead);
        table.appendChild(tbody);
        this.contentHost.appendChild(table);

        this.applyStickyOffsets(thead);
        this.updateDebugOverlay({ table, headerRows, measureCount: displayMeasureCount, columnLeaves, rowDepth, colDepth });
    }

    private computeDisplayColumns(root: DataViewMatrixNode, depth: number, measuresOnColumns: boolean, totalMeasureCount: number): DisplayCol[] {
        type Leaf = { offset: number; labels: string[]; keys: string[]; collapsedAt: number | null; measureIndex?: number };
        const leaves: Leaf[] = [];
        const ranges = new Map<string, { start: number; end: number; level: number }>();
        const subtotalOffsetByKey = new Map<string, number>();
        let offset = 0;
        // In this walk, if measuresOnColumns, the leaf nodes ARE measures.
        // We can try to infer the measure index by their position among siblings?
        // Or strict modulo?
        // Since we iterate children in order, and 'offset' increments, offset % measureCount is one way.
        // But collapsed groups disrupt linear offset.
        // We need to capture the measure index at the leaf.

        const walk = (node: DataViewMatrixNode, labels: string[], keys: string[], parentKey: string, underSubtotal: boolean) => {
            let label = this.nodeLabel(node);
            if (!label && (node as any).isSubtotal) {
                label = (labels.length === 0) ? "Grand Total" : "Total";
            }
            const key = [...keys, label].filter(Boolean).join("||");
            const newLabels = [...labels, label];
            const newKeys = [...keys, label];
            const isSubtotal = (node as any).isSubtotal === true;
            if (!node.children || node.children.length === 0) {
                // Detect subtotal leaf: record parent group key -> offset
                if (underSubtotal || isSubtotal) {
                    if (parentKey) subtotalOffsetByKey.set(parentKey, offset);
                }
                let collapsedAt: number | null = null;
                for (let i = 0; i < newLabels.length; i++) {
                    const k = newKeys.slice(0, i + 1).filter(Boolean).join("||");
                    if (this.collapsedColKeys.has(k)) { collapsedAt = i; break; }
                }
                while (newLabels.length < depth) newLabels.push("");
                while (newKeys.length < depth) newKeys.push(newKeys[newKeys.length - 1] || "");

                // Identify measure index for this leaf
                // If measuresOnColumns, the leaf is the measure.
                // We assume leaves come in blocks of 'totalMeasureCount' for a full group, but they are visited sequentially.
                // Simpler: Just rely on mod logic later if we can't find it here.
                // But we CAN find it if we assume 'node' corresponds to a specific measure.
                // DataViewMatrixNode doesn't explicitly state "I am measure index 0".
                // But we can assign based on traversal if the hierarchy is uniform.
                // Let's rely on global offset for leaf measure index for now?
                // Actually, if we collapsed, we skip leaves.

                let measureIndex: number | undefined = undefined;
                if (measuresOnColumns) {
                    measureIndex = offset % totalMeasureCount;
                    // Always use the measure name for the label when measures are on columns
                    if (this.lastMatrix && this.lastMatrix.valueSources && this.lastMatrix.valueSources[measureIndex]) {
                        newLabels[depth - 1] = this.lastMatrix.valueSources[measureIndex].displayName;
                    }
                }

                leaves.push({ offset, labels: newLabels, keys: newKeys, collapsedAt, measureIndex });
                if (collapsedAt !== null) {
                    const gkey = newKeys.slice(0, collapsedAt + 1).filter(Boolean).join("||");
                    const r = ranges.get(gkey);
                    if (!r) ranges.set(gkey, { start: offset, end: offset, level: collapsedAt });
                    else r.end = offset;
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
        if (root.children) for (const ch of root.children) walk(ch, [], [], "", false);

        const result: DisplayCol[] = [];
        let i = 0;
        while (i < leaves.length) {
            const leaf = leaves[i];
            if (leaf.collapsedAt !== null) {
                const gkey = leaf.keys.slice(0, leaf.collapsedAt + 1).filter(Boolean).join("||");
                const r = ranges.get(gkey)!;
                // Try specific subtotal keys first (standard "Total" or "Grand Total" suffix)
                let subtotalOffset = subtotalOffsetByKey.get(gkey + "||Total");
                if (subtotalOffset === undefined) subtotalOffset = subtotalOffsetByKey.get(gkey + "||Grand Total");
                if (subtotalOffset === undefined) subtotalOffset = subtotalOffsetByKey.get(gkey); // Fallback to direct key if subtotal was mapped directly

                // If measures are on columns, we must emit a column for EACH measure for the collapsed group.
                if (measuresOnColumns && totalMeasureCount > 0) {
                    // Use the group's OWN children range if subtotalOffset is missing.
                    // If subtotalOffset is valid, it points to the LAST measure of the subtotal block.
                    // We calculate baseOffset to point to the FIRST measure.
                    const baseOffset = (subtotalOffset !== undefined) ? (subtotalOffset - (totalMeasureCount - 1)) : r.start;

                    for (let m = 0; m < totalMeasureCount; m++) {
                        // Attempt to retrieve measure name for the header
                        // We can't easily get it here without looking up valueSources.
                        // We will inject it into the labels at the last level.
                        const measureName = (this.lastMatrix && this.lastMatrix.valueSources && this.lastMatrix.valueSources[m])
                            ? (this.lastMatrix.valueSources[m].displayName)
                            : `Measure ${m+1}`;

                        const newLabels = [...leaf.labels];
                        // If we are collapsed at r.level, the label at r.level is visible (Group Name).
                        // Labels below are empty.
                        // We should put the measure name at the LAST level (depth - 1).
                        if (depth > 0) newLabels[depth - 1] = measureName;

                        result.push({
                            kind: "collapsed",
                            start: r.start,
                            end: r.end,
                            labels: newLabels,
                            keys: leaf.keys,
                            collapsedLevel: r.level,
                            key: gkey,
                            subtotalOffset: baseOffset + m,
                            measureIndex: m
                        });
                    }
                } else {
                    result.push({ kind: "collapsed", start: r.start, end: r.end, labels: leaf.labels, keys: leaf.keys, collapsedLevel: r.level, key: gkey, subtotalOffset });
                }

                i = r.end + 1;
            } else {
                result.push({ kind: "leaf", offset: leaf.offset, labels: leaf.labels, keys: leaf.keys, measureIndex: leaf.measureIndex });
                i++;
            }
        }
        return result;
    }

    private isAnyAncestorCollapsed(path: string[]): boolean {
        // path is array of labels up to, but not including, the leaf
        let accum: string[] = [];
        for (const p of path) {
            accum.push(p);
            const k = accum.join("||");
            if (this.collapsedColKeys.has(k)) return true;
        }
        return false;
    }

    private getRowDepth(rows?: powerbi.DataViewHierarchy): number {
        if (!rows || !rows.root) return 0;
        if (rows.levels && rows.levels.length) return rows.levels.length;
        const depthFrom = (node: DataViewMatrixNode): number => {
            if (!node.children || node.children.length === 0) return 0;
            let max = 0;
            for (const c of node.children) max = Math.max(max, depthFrom(c));
            return 1 + max;
        };
        if (!rows.root.children || rows.root.children.length === 0) return 0;
        return depthFrom(rows.root);
    }

    private getColumnDepth(columns?: powerbi.DataViewHierarchy): number {
        if (!columns || !columns.root) return 0;
        if (columns.levels && columns.levels.length) return columns.levels.length;
        const depthFrom = (node: DataViewMatrixNode): number => {
            if (!node.children || node.children.length === 0) return 1;
            let max = 0;
            for (const c of node.children) max = Math.max(max, depthFrom(c));
            return 1 + max;
        };
        if (!columns.root.children || columns.root.children.length === 0) return 1;
        return depthFrom(columns.root);
    }

    private refresh() {
        // Re-render using the last received matrix
        while (this.contentHost.firstChild) this.contentHost.removeChild(this.contentHost.firstChild);
        if (this.lastMatrix) this.renderMatrix(this.lastMatrix);
    }

    private applyStickyOffsets(thead: HTMLTableSectionElement) {
        if (this.resizeObserver) {
            this.resizeObserver.disconnect();
            this.resizeObserver = null;
        }

        if (!this.container.classList.contains("ghm-sticky")) return;

        const compute = () => {
            const rows = Array.from(thead.rows);
            if (rows.length === 0) return;

            let top = 0;
            for (const tr of rows) {
                const rect = tr.getBoundingClientRect();
                const h = rect.height;
                const cells = Array.from(tr.cells) as HTMLElement[];
                for (const cell of cells) {
                    cell.style.top = `${top}px`;
                }
                top += h;
            }
        };

        // Compute initially
        requestAnimationFrame(compute);

        // Compute on resize
        this.resizeObserver = new ResizeObserver(() => requestAnimationFrame(compute));
        this.resizeObserver.observe(thead);
        for (const r of Array.from(thead.rows)) this.resizeObserver.observe(r);
    }

    private persistState() {
        const widthsObj: any = {};
        for (const [k, v] of this.columnWidthPx.entries()) widthsObj[k] = v;
        this.host.persistProperties({
            merge: [
                {
                    objectName: "state",
                    properties: {
                        columnWidths: JSON.stringify(widthsObj),
                        collapsedCols: JSON.stringify(Array.from(this.collapsedColKeys)),
                        collapsedRows: JSON.stringify(Array.from(this.collapsedRowKeys)),
                        repeatLabels: this.repeatLabels,
                        compactLayout: this.compactLayout,
                        viewMode: this.viewMode,
                        currentPage: this.currentPage
                    },
                    selector: null
                }
            ]
        });
    }

    private createSelectionId(rowNode: DataViewMatrixNode, colNode?: DataViewMatrixNode, measureIndex?: number): ISelectionId {
        const builder = this.host.createSelectionIdBuilder();
        if (rowNode) builder.withMatrixNode(rowNode, this.lastMatrix!.rows.levels!);
        if (colNode && this.lastMatrix!.columns.levels) builder.withMatrixNode(colNode, this.lastMatrix!.columns.levels);
        if (measureIndex !== undefined) builder.withMeasure(this.lastMatrix!.valueSources![measureIndex].queryName);
        return builder.createSelectionId();
    }

    private getObjectValue<T>(objects: powerbi.DataViewObjects | undefined, objectName: string, propertyName: string, defaultValue: T): T {
        const obj = objects && (objects as any)[objectName];
        const v = obj && obj[propertyName];
        return (v !== undefined) ? (v as T) : defaultValue;
    }

    private parseColor(input: any): string {
        if (input == null) return "";
        if (typeof input === "string") return input;
        const maybeVal = (input as any).value;
        if (typeof maybeVal === "string") return maybeVal;
        const solid = (input as any).solid;
        if (solid && typeof solid.color === "string") return solid.color;
        return "";
    }


    private getValueFromMapForDisplayCol(valuesMap: { [key: number]: powerbi.DataViewMatrixNodeValue }, ref: DisplayCol, measureIndex: number, totalMeasureCount: number): any {
        if (ref.kind === "leaf") {
            const key = ref.offset * totalMeasureCount + measureIndex;
            const cell = valuesMap[key];
            return cell && cell.value != null ? cell.value : "";
        } else {
            // Strict host-only: for collapsed columns, render a value only when
            // the host emitted a dedicated subtotal leaf for that column group.
            const coll = ref as any as { subtotalOffset?: number };
            if (coll.subtotalOffset !== undefined) {
                const key = (coll.subtotalOffset as number) * totalMeasureCount + measureIndex;
                const cell = valuesMap[key];
                return (cell && cell.value != null) ? cell.value : "";
            }
            return "";
        }
    }

    private applyCellConditionalStyle(td: HTMLTableCellElement, valuesMap: { [key:number]: powerbi.DataViewMatrixNodeValue }, ref: DisplayCol, measureIndex: number, totalMeasureCount: number, measuresOnColumns: boolean) {
        const calcKey = (offset: number) => measuresOnColumns ? offset : (offset * totalMeasureCount + measureIndex);

        let key: number | null = null;
        if (ref.kind === 'leaf') key = calcKey(ref.offset);
        else {
            const coll = ref as any as { subtotalOffset?: number };
            if (coll.subtotalOffset !== undefined) {
                 key = measuresOnColumns ? coll.subtotalOffset : (coll.subtotalOffset * totalMeasureCount + measureIndex);
            }
        }
        if (key == null) return;
        const cell = valuesMap[key];
        const obj = cell && (cell.objects as any);
        const colorsObj = obj && (obj.colors as any);
        const explicitFont = colorsObj && (colorsObj.cellColor as any);
        const explicitBg = colorsObj && (colorsObj.cellBg as any);
        const font = explicitFont && ((explicitFont as any).solid?.color || (explicitFont as any).value);
        const bg = explicitBg && ((explicitBg as any).solid?.color || (explicitBg as any).value);

        // Measure-driven colors (per value)
        const tryMeasureColor = (mi: number): string | null => {
            if (mi < 0) return null;

            if (!measuresOnColumns) {
                if (ref.kind === "leaf") {
                    const colorKey = ref.offset * totalMeasureCount + mi;
                    const c = (valuesMap as any)[colorKey];
                    if (c && c.value != null) return String(c.value);
                } else {
                    const coll = ref as any as { subtotalOffset?: number };
                    if (coll.subtotalOffset !== undefined) {
                        const colorKey = (coll.subtotalOffset as number) * totalMeasureCount + mi;
                        const c = (valuesMap as any)[colorKey];
                        if (c && c.value != null) return String(c.value);
                    }
                }
            }
            return null;
        };
        const selectColorIndex = (arr: number[]): number => {
            if (!arr.length) return -1;
            const displayPos = this.displayMeasureIndices.indexOf(measureIndex);
            if (displayPos >= 0 && displayPos < arr.length) return arr[displayPos];
            return arr[arr.length - 1];
        };
        const measureBg = tryMeasureColor(selectColorIndex(this.cellBgColorMeasureIndices));
        const measureFont = tryMeasureColor(selectColorIndex(this.cellFontColorMeasureIndices));
        const measureDefaultFont = this.measureFontColorSettings[measureIndex] || "";
        const measureDefaultBg = this.measureBgColorSettings[measureIndex] || "";

        if (measureFont) (td as any).style.color = measureFont;
        if (measureBg) (td as any).style.backgroundColor = measureBg;
        if (!measureFont && font) (td as any).style.color = font;
        if (!measureBg && bg) (td as any).style.backgroundColor = bg;
        if (!measureFont && !font && measureDefaultFont) (td as any).style.color = measureDefaultFont;
        if (!measureBg && !bg && measureDefaultBg) (td as any).style.backgroundColor = measureDefaultBg;
    }

    private getRowMeasureBg(valuesMap: { [key:number]: powerbi.DataViewMatrixNodeValue }, totalMeasureCount: number): string | null {
        if (!this.cellBgColorMeasureIndices.length || totalMeasureCount <= 0) return null;
        const bgIndex = this.cellBgColorMeasureIndices[0];
        const keys = Object.keys(valuesMap as any).map(k => +k).filter(k => !Number.isNaN(k));
        for (const k of keys) {
            if ((k % totalMeasureCount) === bgIndex) {
                const c = (valuesMap as any)[k];
                if (c && c.value != null) return String(c.value);
            }
        }
        return null;
    }

    private updateDebugOverlay(_ctx?: { table: HTMLTableElement; headerRows: Array<any>; measureCount: number; columnLeaves: DisplayCol[]; rowDepth: number; colDepth: number }) { /* debug disabled */ return; }

    private loadLevelStyles(matrix: DataViewMatrix) {
        this.rowLevelStyles = [];
        this.colLevelStyles = [];
        this.rowLevelNames = [];
        const rows = matrix.rows; const cols = matrix.columns;
        if (rows && rows.levels) {
            rows.levels.forEach((lvl, i) => {
                const src = lvl.sources && lvl.sources[0];
                this.rowLevelNames.push((src && src.displayName) || `Row level ${i + 1}`);
                const obj = (src && (src as any).objects && (src as any).objects.labelStylePerLevel) || {};
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
                const obj = (src && (src as any).objects && (src as any).objects.labelStylePerLevel) || {};
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

    private getCollapsedRowGroupValue(node: DataViewMatrixNode, ref: DisplayCol, measureIndex: number, measureCount: number): any {
        // Prefer host-provided row subtotal on the collapsed node itself
        const tryFromNodeValues = (): any => {
            const valuesMap = (node.values || {}) as any;
            if (ref.kind === "leaf") {
                const key = ref.offset * measureCount + measureIndex;
                const cell = valuesMap[key];
                if (cell && cell.value != null) return cell.value;
                return undefined;
            } else {
                // Strict host-only for collapsed column groups: require a dedicated subtotal leaf
                const coll = ref as any as { subtotalOffset?: number };
                if (coll.subtotalOffset !== undefined) {
                    const key = (coll.subtotalOffset as number) * measureCount + measureIndex;
                    const cell = valuesMap[key];
                    if (cell && cell.value != null) return cell.value;
                }
                return undefined;
            }
        };

        const direct = tryFromNodeValues();
        if (direct !== undefined) return direct;
        // Try subtotal child provided by host (isSubtotal)
        const subtotal = this.findSubtotalChild(node);
        if (subtotal) {
            const map = (subtotal.values || {}) as any;
            if (ref.kind === "leaf") {
                const key = ref.offset * measureCount + measureIndex;
                const cell = map[key];
                if (cell && cell.value != null) return cell.value;
            } else {
                const coll = ref as any as { subtotalOffset?: number };
                if (coll.subtotalOffset !== undefined) {
                    const key = (coll.subtotalOffset as number) * measureCount + measureIndex;
                    const cell = map[key];
                    if (cell && cell.value != null) return cell.value;
                }
            }
        }
        // Otherwise leave empty by design
        return "";
    }

    private findSubtotalChild(node: DataViewMatrixNode): DataViewMatrixNode | null {
        if (!node || !node.children) return null;
        for (const ch of node.children) {
            if ((ch as any).isSubtotal) return ch;
        }
        return null;
    }

    private getVisibleRowDepth(rows: powerbi.DataViewHierarchy | undefined, maxDepth: number): number {
        if (!rows || !rows.root) return maxDepth;
        const depthFrom = (node: DataViewMatrixNode, depth: number, parentKey: string): number => {
            const label = this.nodeLabel(node);
            const key = [parentKey, label].filter(Boolean).join("||");
            if (!node.children || node.children.length === 0) return depth + 1;
            if (this.collapsedRowKeys.has(key)) return depth + 1;
            let max = depth + 1;
            for (const ch of node.children) max = Math.max(max, depthFrom(ch, depth + 1, key));
            return max;
        };
        if (!rows.root.children || rows.root.children.length === 0) return 1;
        let m = 1; for (const ch of rows.root.children) m = Math.max(m, depthFrom(ch, 0, ""));
        return Math.min(maxDepth, m);
    }

    private applyRowHeaderStyle(th: HTMLElement, level?: number) {
        const style = (level !== undefined && this.rowLevelStyles[level]) ? this.rowLevelStyles[level] : {};
        const fs = style.fontSize ?? this.rowHeaderFontSize; if (fs) th.style.fontSize = `${fs}px`;
        th.style.fontFamily = (style.fontFamily || this.rowHeaderFontFamily) || "";
        if (this.rowHeaderBold || style.fontWeight) th.style.fontWeight = style.fontWeight || "bold";
        if (style.italic) th.style.fontStyle = "italic";
        if (style.underline) th.style.textDecoration = "underline";
        if (this.rowHeaderColor) th.style.color = this.rowHeaderColor;
    }

    private applyGridBorder(el: HTMLElement, isHeader: boolean) {
        const thick = Math.max(0, this.gridThickness || 0);
        if (!thick) return;
        const color = this.gridColor || "#d0d0d0";
        const horiz = this.gridShowHorizontal;
        const vert = this.gridShowVertical;
        if (horiz) {
            el.style.borderTop = `${thick}px solid ${color}`;
            el.style.borderBottom = `${thick}px solid ${color}`;
        }
        if (vert) {
            el.style.borderLeft = `${thick}px solid ${color}`;
            el.style.borderRight = `${thick}px solid ${color}`;
        }
        // Ensure borders collapse visually
        (el as any).style.borderCollapse = "collapse";
    }

    private getDisplayedRowDepth(rows?: powerbi.DataViewHierarchy): number {
        if (!rows || !rows.root || !rows.root.children) return 1;
        const depthFrom = (node: DataViewMatrixNode, depth: number, parentKey: string): number => {
            const label = this.nodeLabel(node);
            const key = [parentKey, label].filter(Boolean).join("||");
            if (this.collapsedRowKeys.has(key)) return depth + 1;
            if (!node.children || node.children.length === 0) return depth + 1;
            let max = depth + 1;
            for (const ch of node.children) max = Math.max(max, depthFrom(ch, depth + 1, key));
            return max;
        };
        let m = 1; for (const ch of rows.root.children) m = Math.max(m, depthFrom(ch, 0, ""));
        return m;
    }

    private getDisplayedColDepth(cols?: powerbi.DataViewHierarchy): number {
        if (!cols || !cols.root || !cols.root.children) return 1;
        const depthFrom = (node: DataViewMatrixNode, depth: number, path: string[]): number => {
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (this.collapsedColKeys.has(key)) return depth + 1;
            if (!node.children || node.children.length === 0) return depth + 1;
            let max = depth + 1;
            for (const ch of node.children) max = Math.max(max, depthFrom(ch, depth + 1, [...path, label]));
            return max;
        };
        let m = 1; for (const ch of cols.root.children) m = Math.max(m, depthFrom(ch, 0, []));
        return m;
    }

    private collapseAll() {
        if (this.lastMatrix?.columns?.root?.children) {
            const add = (node: DataViewMatrixNode, path: string[]) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key) this.collapsedColKeys.add(key);
                    for (const ch of node.children) add(ch, [...path, label]);
                }
            };
            for (const ch of this.lastMatrix.columns.root.children) add(ch, []);
        }
        if (this.lastMatrix?.rows?.root?.children) {
            const add = (node: DataViewMatrixNode, path: string[]) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key) this.collapsedRowKeys.add(key);
                    for (const ch of node.children) add(ch, [...path, label]);
                }
            };
            for (const ch of this.lastMatrix.rows.root.children) add(ch, []);
        }
    }

    private expandAll() {
        this.collapsedColKeys.clear();
        this.collapsedRowKeys.clear();
    }

    private collectDisplayRowsRepeat(node: DataViewMatrixNode, rowDepth: number, parentLabels: string[] = [], parentKey: string = ""): Array<{ labels: string[]; valuesMap?: { [key: number]: powerbi.DataViewMatrixNodeValue }; numericKeys?: number[] }> {
        const rows: Array<{ labels: string[]; valuesMap?: { [key: number]: powerbi.DataViewMatrixNodeValue }; numericKeys?: number[] }> = [];
        const label = this.nodeLabel(node);
        const thisKey = [parentKey, label].filter(Boolean).join("||");
        const labels = label ? [...parentLabels, label] : [...parentLabels];
        if (node.level === undefined && node.children && node.children.length) {
            for (const ch of node.children) rows.push(...this.collectDisplayRowsRepeat(ch, rowDepth, labels, thisKey));
            return rows;
        }
        if (node.children && node.children.length) {
            if (this.collapsedRowKeys.has(thisKey)) {
                const padded = [...labels]; while (padded.length < rowDepth) padded.push("");
                // Strict host-only subtotal for collapsed groups
                const subtotalChild = (node.children as any[]).find(ch => (ch as any).isSubtotal);
                const map = subtotalChild ? (subtotalChild.values as any) : {};
                const numericKeys = Object.keys(map as any).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
                rows.push({ labels: padded, valuesMap: map, numericKeys });
                return rows;
            }
            for (const ch of node.children) rows.push(...this.collectDisplayRowsRepeat(ch, rowDepth, labels, thisKey));
            return rows;
        }
        const padded = [...labels]; while (padded.length < rowDepth) padded.push("");
        const valuesMap = node.values || {};
        const numericKeys = Object.keys(valuesMap as any).map(k => +k).filter(k => !Number.isNaN(k)).sort((a, b) => a - b);
        rows.push({ labels: padded, valuesMap: valuesMap as any, numericKeys });
        return rows;
    }

    private collectOutlineRowsWithTotals(root: DataViewMatrixNode, rowDepth: number): Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> {
        const rows: Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> = [];
        const grandTotalRows: Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> = [];
        const tryRootTotal = () => {
            if (!this.showGrandTotal) return;
            const subtotalChild = root.children && (root.children as any[]).find(ch => (ch as any).isSubtotal);
            if (subtotalChild && subtotalChild.values) {
                const map = subtotalChild.values as any;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                grandTotalRows.push({ labels, valuesMap: map, isTotal: true });
                return;
            }
            if (this.grandTotalFallback && root.values) {
                const map = root.values as any;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                grandTotalRows.push({ labels, valuesMap: map, isTotal: true });
            }
        };
        tryRootTotal();

        const traverse = (node: DataViewMatrixNode, depth: number, path: string[], keyPath: string[]) => {
            const label = this.nodeLabel(node);
            const hasChildren = !!(node.children && node.children.length);
            const subtotalChild = hasChildren ? (node.children as any[]).find(ch => (ch as any).isSubtotal) : null;
            const hasGroupValues = hasChildren && node.values && Object.keys(node.values as any).length > 0;
            const map = subtotalChild ? (subtotalChild.values as any) : (hasGroupValues ? (node.values as any) : null);
            const gKey = [...keyPath, label].filter(Boolean).join("||");
            const isCollapsed = gKey && this.collapsedRowKeys.has(gKey);
            const isRootNode = depth === 0 && (!label || label === "") && keyPath.length === 0;

            if (isRootNode) {
                 if (hasChildren) {
                    for (const ch of node.children) {
                        if ((ch as any).isSubtotal) continue;
                        traverse(ch, 0, [], []);
                    }
                }
                return;
            }

            // Compact mode logic
            if (this.compactLayout) {
                const labels = new Array(rowDepth).fill("");
                for (let i = 0; i < path.length; i++) labels[i] = path[i];
                labels[depth] = label;

                // Collapsed State: Always show one row with values (acting as total)
                if (isCollapsed) {
                    rows.push({
                        labels,
                        valuesMap: map || {},
                        isTotal: true, // Treat as total for styling/behavior
                        toggleKey: gKey,
                        depth,
                        collapsed: true
                    });
                    return;
                }

                // Expanded State
                if (hasChildren) {
                    const subtotalAtBottom = this.rowSubtotalPosition === "Bottom";

                    // 1. Header Row
                    // If subtotals at bottom, this is "empty line" (label only).
                    // If subtotals at top, this IS the subtotal row.
                    if (subtotalAtBottom) {
                        rows.push({
                            labels,
                            valuesMap: {}, // Empty values
                            isTotal: true, // Mark as total/group header
                            toggleKey: gKey,
                            depth,
                            collapsed: false
                        });
                    } else {
                        // Top Subtotal
                        if (this.rowSubtotalsEnabled && map) {
                             rows.push({
                                labels,
                                valuesMap: map,
                                isTotal: true,
                                toggleKey: gKey,
                                depth,
                                collapsed: false
                            });
                        } else {
                            // Header only (no subtotal values if disabled)
                            rows.push({
                                labels,
                                valuesMap: {},
                                isTotal: true,
                                toggleKey: gKey,
                                depth,
                                collapsed: false
                            });
                        }
                    }

                    // 2. Children
                    for (const ch of node.children) {
                        if ((ch as any).isSubtotal) continue;
                        // build next path correctly: fill parent positions
                        const nextPath = new Array(rowDepth).fill("");
                        for (let i = 0; i < path.length; i++) nextPath[i] = path[i];
                        nextPath[depth] = label;
                        traverse(ch, depth + 1, nextPath.slice(0, depth + 1), [...keyPath, label]);
                    }

                    // 3. Bottom Subtotal
                    if (subtotalAtBottom && this.rowSubtotalsEnabled && map) {
                        const totalLabels = [...labels];
                        totalLabels[depth] = `${label} Total`;
                        rows.push({
                            labels: totalLabels,
                            valuesMap: map,
                            isTotal: true,
                            depth,
                            // No toggle on bottom total usually, but indentation matches group
                        });
                    }
                    return;
                } else {
                    // Leaf in compact mode
                    rows.push({ labels, valuesMap: node.values as any, isTotal: false, node });
                    return;
                }
            }

            // Tabular (Original) Logic
            const includeSubtotal = (this.rowSubtotalsEnabled || isCollapsed) && hasChildren && map && depth < rowDepth;
            const subtotalAtBottom = this.rowSubtotalPosition === "Bottom";
            const pushSubtotalRow = (collapsedFlag: boolean) => {
                const labels = new Array(rowDepth).fill("");
                for (let i = 0; i < path.length; i++) labels[i] = path[i];
                labels[depth] = collapsedFlag ? label : `${label} Total`;
                const toggles: { [key: number]: string } = {};
                if (gKey) toggles[depth] = gKey;
                rows.push({ labels, valuesMap: map!, isTotal: true, toggleKey: gKey || undefined, depth, collapsed: collapsedFlag, toggles });
            };

            if (includeSubtotal && (!subtotalAtBottom || isCollapsed)) {
                pushSubtotalRow(!!isCollapsed);
                if (isCollapsed) return;
            }

            if (hasChildren) {
                let firstChild = true;
                for (const ch of node.children) {
                    if ((ch as any).isSubtotal) continue;
                    const nextPath = new Array(rowDepth).fill("");
                    for (let i = 0; i < path.length; i++) nextPath[i] = path[i];
                    nextPath[depth] = label;

                    // In tabular mode, the first row of the expanded group must carry the toggleKey
                    // so that the button appears on the parent's label.
                    // We pass it down as a special property to the recursive call if it's the first child.
                    // However, traverse() pushes rows directly. We need a way to attach the toggleKey to the *first row emitted* by the traversal of the first child.

                    // Actually, simpler: we can just attach the toggleKey to the first child's traversal call.
                    // But traverse doesn't accept an "override toggle key".
                    // Let's modify traverse signature or logic slightly?
                    // No, traverse is recursive.

                    // Alternative: The row with the label is emitted deep inside traverse.
                    // If we are at 'depth', the row emitted for 'ch' (at depth+1) will display 'label' at 'depth'.
                    // Wait, the row generation logic in traverse puts 'path' into labels.
                    // So every child row has the parent label at 'depth'.
                    // We only want the toggle on the *first* row that shows this label.
                    // The rendering logic handles "first row of group" visual toggle placement?
                    // In renderBody: "if ((rowInfo as any).isTotal && (rowInfo as any).toggleKey && lvl === toggleLevel)"
                    // This only puts toggle on TOTAL rows.
                    // We need it on the HEADER row for tabular.

                    // In tabular mode without top subtotals, the "header" is just the first data row of the first child.
                    // We need to mark that specific row.

                    const isFirstChildOfGroup = firstChild;
                    firstChild = false;

                    // Recurse
                    const childStartIdx = rows.length;
                    traverse(ch, depth + 1, nextPath.slice(0, depth + 1), [...keyPath, label]);

                    // If this was the first child, and we didn't emit a top subtotal, we must attach the toggleKey to the first row generated by this child.
                    if (isFirstChildOfGroup && !isCollapsed && !(includeSubtotal && !subtotalAtBottom)) {
                        if (rows.length > childStartIdx) {
                            const firstRow = rows[childStartIdx];
                            // Ensure toggles map exists and add the toggle for this depth
                            if (!firstRow.toggles) firstRow.toggles = {};
                            firstRow.toggles[depth] = gKey;

                            // Keep toggleKey for backward compatibility (used by compact mode logic)
                            // but in tabular mode we prefer `toggles`
                            if (!firstRow.toggleKey) {
                                firstRow.toggleKey = gKey;
                                firstRow.depth = depth;
                                firstRow.isTotal = firstRow.isTotal || false; // preserve existing
                            }
                        }
                    }
                }
                if (includeSubtotal && subtotalAtBottom) pushSubtotalRow(false);
                return;
            }
            // leaf
            const labels = new Array(rowDepth).fill("");
            for (let i = 0; i < path.length; i++) labels[i] = path[i];
            labels[depth] = label;
            rows.push({ labels, valuesMap: node.values as any, isTotal: false, node });
        };

        traverse(root, 0, [], []);

        if (this.grandTotalPosition === "Top") return grandTotalRows.concat(rows);
        return rows.concat(grandTotalRows);
    }

    // Level-wise expand/collapse helpers
    private collapseRowLevel() {
        if (!this.lastMatrix?.rows?.root?.children) return;
        const currentDepth = this.getDisplayedRowDepth(this.lastMatrix.rows);
        if (currentDepth <= 1) return;
        const targetDepth = currentDepth - 1; // collapse one level globally
        const addKeysAtDepth = (node: DataViewMatrixNode, depth: number, path: string[]) => {
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (!node.children || node.children.length === 0) return;
            if (depth === targetDepth) {
                if (!this.collapsedRowKeys.has(key)) this.collapsedRowKeys.add(key);
                return;
            }
            for (const ch of node.children) addKeysAtDepth(ch, depth + 1, [...path, label]);
        };
        for (const ch of this.lastMatrix.rows.root.children) addKeysAtDepth(ch, 1, []);
    }

    private expandRowLevel() {
        if (this.collapsedRowKeys.size === 0) return;
        let minDepth = Number.MAX_SAFE_INTEGER;
        for (const key of this.collapsedRowKeys) {
            const depth = key ? key.split("||").length : 1;
            if (depth < minDepth) minDepth = depth;
        }
        for (const key of Array.from(this.collapsedRowKeys)) {
            const depth = key ? key.split("||").length : 1;
            if (depth === minDepth) this.collapsedRowKeys.delete(key);
        }
    }

    private collapseColLevel() {
        if (!this.lastMatrix?.columns?.root?.children) return;
        const current = this.getDisplayedColDepth(this.lastMatrix.columns);
        if (current <= 1) return;
        const target = current - 1;
        const addAtDepth = (node: DataViewMatrixNode, depth: number, path: string[]) => {
            if (!node.children || node.children.length === 0) return;
            const label = this.nodeLabel(node);
            const key = [...path, label].filter(Boolean).join("||");
            if (depth === target) {
                if (!this.collapsedColKeys.has(key)) this.collapsedColKeys.add(key);
                return;
            }
            for (const ch of node.children) addAtDepth(ch, depth + 1, [...path, label]);
        };
        for (const ch of this.lastMatrix.columns.root.children) addAtDepth(ch, 1, []);
    }

    private expandColLevel() {
        if (this.collapsedColKeys.size === 0) return;
        let minDepth = Number.MAX_SAFE_INTEGER;
        for (const key of this.collapsedColKeys) {
            const depth = key ? key.split("||").length : 1;
            if (depth < minDepth) minDepth = depth;
        }
        for (const key of Array.from(this.collapsedColKeys)) {
            const depth = key ? key.split("||").length : 1;
            if (depth === minDepth) this.collapsedColKeys.delete(key);
        }
    }

    private initializeDefaultCollapsed(matrix: DataViewMatrix, enableCollapse: boolean = true) {
        if (!enableCollapse) return;
        // Collapse all column groups by default
        if (matrix.columns && matrix.columns.root) {
            const addColKeys = (node: DataViewMatrixNode, path: string[]) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key) this.collapsedColKeys.add(key);
                    for (const ch of node.children) addColKeys(ch, [...path, label]);
                }
            };
            if (matrix.columns.root.children) {
                for (const ch of matrix.columns.root.children) addColKeys(ch, []);
            }
        }

        // Collapse all row groups by default
        if (matrix.rows && matrix.rows.root) {
            const addRowKeys = (node: DataViewMatrixNode, path: string[]) => {
                const label = this.nodeLabel(node);
                const key = [...path, label].filter(Boolean).join("||");
                if (node.children && node.children.length) {
                    if (key) this.collapsedRowKeys.add(key);
                    for (const ch of node.children) addRowKeys(ch, [...path, label]);
                }
            };
            if (matrix.rows.root.children) {
                for (const ch of matrix.rows.root.children) addRowKeys(ch, []);
            }
        }
    }

    // no custom aggregation: collapsed groups show blank cells to avoid misrepresenting measure semantics

    private buildHeaderRowsFromDisplay(displayCols: DisplayCol[], depth: number): Array<Array<{ label: string; span: number; key: string; togglable: boolean; collapsed: boolean; isLeafHeader: boolean; queryKeys: string }>> {
        const rows: Array<Array<{ label: string; span: number; key: string; togglable: boolean; collapsed: boolean; isLeafHeader: boolean; queryKeys: string }>> = [];
        for (let level = 0; level < depth; level++) {
            const row: Array<{ label: string; span: number; key: string; togglable: boolean; collapsed: boolean; isLeafHeader: boolean; queryKeys: string }> = [];
            let i = 0;
            while (i < displayCols.length) {
                const col = displayCols[i];
                const label = col.kind === "leaf" ? col.labels[level] : (
                    (level <= col.collapsedLevel || (col.measureIndex !== undefined && level === depth - 1))
                        ? col.labels[level] : "");
                const key = col.keys[level] || "";
                const collapsed = col.kind === "collapsed" && level === col.collapsedLevel ? true : this.collapsedColKeys.has(key) && level < depth - 1;
                // Avoid showing a bare +/- toggle with no label when an upper level
                // header is collapsed. Only show toggles when the label is visible.
                const togglable = !!key && level < depth - 1 && !!label;

                // Determine if this is the "effective leaf" for sorting
                const isLeafHeader = (col.kind === "leaf" && level === depth - 1) || (col.kind === "collapsed" && level === col.collapsedLevel);
                const queryKeys = col.kind === "leaf" ? col.keys.join("||") : col.key;

                let span = 1;
                let j = i + 1;
                while (j < displayCols.length) {
                    const nxt = displayCols[j];
                    const nLabel = nxt.kind === "leaf" ? nxt.labels[level] : (level < (nxt.collapsedLevel ?? 0) ? nxt.labels[level] : (level === (nxt.collapsedLevel ?? -1) ? nxt.labels[level] : ""));
                    const nKey = nxt.keys[level] || "";
                    const nCollapsed = nxt.kind === "collapsed" && level === (nxt.collapsedLevel ?? -1) ? true : this.collapsedColKeys.has(nKey) && level < depth - 1;
                    if (nLabel !== label || nKey !== key || nCollapsed !== collapsed) break;
                    span++; j++;
                }
                row.push({ label, span, key, togglable, collapsed, isLeafHeader, queryKeys });
                i = j;
            }
            // If a full header row yields no labels at all (e.g., due to a higher
            // level being collapsed), drop the row entirely to avoid an empty
            // header line containing only +/- icons or blanks.
            const hasAnyLabel = row.some(c => !!(c.label && c.label.trim().length > 0));
            if (hasAnyLabel) rows.push(row);
        }
        return rows;
    }

    private beginResize(e: MouseEvent, key: string) {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = this.columnWidthPx.get(key) ?? 120;
        const onMove = (ev: MouseEvent) => {
            const delta = ev.clientX - startX;
            const newW = Math.max(40, startWidth + delta);
            this.columnWidthPx.set(key, newW);
            const cols = this.colElsByKey.get(key) || [];
            for (const c of cols) c.style.width = `${newW}px`;
        };
        const onUp = () => {
            window.removeEventListener("mousemove", onMove);
            window.removeEventListener("mouseup", onUp);
            this.persistState();
        };
        window.addEventListener("mousemove", onMove);
        window.addEventListener("mouseup", onUp);
    }

    private getCellValueForDisplayCol(valuesMap: { [key: number]: powerbi.DataViewMatrixNodeValue }, numericKeys: number[], ref: DisplayCol, measureIndex: number, totalMeasureCount: number, measuresOnColumns: boolean): any {
        // If measures on columns, ref.offset is the leaf index.
        // BUT if collapsed, ref.subtotalOffset is base offset.
        // If collapsed group has multiple measures, ref.measureIndex differentiates.
        // We adjusted subtotalOffset in computeDisplayColumns to point to specific measure offset if possible.
        // Let's trust ref.subtotalOffset.

        const calcKey = (offset: number) => measuresOnColumns ? offset : (offset * totalMeasureCount + measureIndex);

        if (ref.kind === "leaf") {
            const prefKey = calcKey(ref.offset);
            const preferredCell = valuesMap[prefKey];
            if (preferredCell && preferredCell.value != null) return preferredCell.value;
            return "";
        } else {
            const coll = ref as any as { subtotalOffset?: number };
            if (coll.subtotalOffset !== undefined) {
                // In computeDisplayColumns, we already added +m to subtotalOffset if needed.
                const key = measuresOnColumns ? coll.subtotalOffset : (coll.subtotalOffset * totalMeasureCount + measureIndex);
                const cell = valuesMap[key];
                return (cell && cell.value != null) ? cell.value : "";
            }
            return "";
        }
    }

    private nodeLabel(node: DataViewMatrixNode): string {
        if (node.levelValues && node.levelValues.length) {
            // Prefer levelValues for matrix nodes
            return String(node.levelValues.map(v => v.value).filter(v => v != null)[0] ?? "");
        }
        if (node.value != null) return String(node.value);
        return "";
    }

    private formatValueByMeasure(v: any, measureIndex: number): string {
        if (v == null || v === "") return "";
        const fmt = (this.measureFormats && this.measureFormats[measureIndex]) ? this.measureFormats[measureIndex] : undefined;
        try {
            const f = valueFormatter.create({ format: fmt });
            return f.format(v);
        } catch { return String(v); }
    }

    private sortRowsRecursive(nodes: DataViewMatrixNode[], valueKey: number, direction: "ASC" | "DESC", type: "value" | "label", level: number, currentDepth: number) {
        if (!nodes) return;

        if (type === "value") {
            nodes.sort((a, b) => {
                const valA = (a.values && a.values[valueKey]) ? a.values[valueKey].value : null;
                const valB = (b.values && b.values[valueKey]) ? b.values[valueKey].value : null;
                if (valA === valB) return 0;
                if (valA == null) return 1;
                if (valB == null) return -1;
                if (valA < valB) return direction === "ASC" ? -1 : 1;
                return direction === "ASC" ? 1 : -1;
            });
        } else {
            // Label sort
            // In compact layout (level=-1 or similar implicit), or if current depth matches target level
            if (level === -1 || currentDepth === level) {
                 nodes.sort((a, b) => {
                    const labelA = this.nodeLabel(a);
                    const labelB = this.nodeLabel(b);
                    if (labelA === labelB) return 0;
                    if (labelA == null) return 1;
                    if (labelB == null) return -1;
                    if (labelA < labelB) return direction === "ASC" ? -1 : 1;
                    return direction === "ASC" ? 1 : -1;
                 });
            }
        }

        for (const child of nodes) {
            if (child.children) this.sortRowsRecursive(child.children, valueKey, direction, type, level, currentDepth + 1);
        }
    }

    private attachSortHandler(th: HTMLElement, queryKeys: string, measureIndex: number, level: number = 0, type: "value" | "label" = "value") {
        th.style.cursor = "pointer";
        const isSorted = this.sortState && this.sortState.type === type && (
            (type === "value" && this.sortState.queryKeys === queryKeys && this.sortState.measureIndex === measureIndex) ||
            (type === "label" && this.sortState.level === level)
        );

        if (isSorted) {
            const arrow = document.createElement("span");
            arrow.className = "ghm-sort-icon";
            arrow.textContent = this.sortState!.direction === "ASC" ? "▲" : "▼";
            th.appendChild(arrow);
        }
        th.addEventListener("click", (e) => {
            e.stopPropagation();
            if (isSorted) {
                if (this.sortState!.direction === "ASC") {
                    this.sortState!.direction = "DESC";
                } else {
                    this.sortState = null;
                }
            } else {
                this.sortState = { type, queryKeys, measureIndex, level, direction: "ASC" };
            }
            this.refresh();
        });
    }

    private getTooltipData(value: any, rowNode: DataViewMatrixNode, colNode: DataViewMatrixNode | undefined, measureIndex: number): VisualTooltipDataItem[] {
        const res: VisualTooltipDataItem[] = [];

        // Rows
        if (rowNode) {
            // Ancestors?
            // The node structure doesn't easily link back to parents.
            // But we can use levelValues if available or simple node value.
            // For a flat list of headers, we need to reconstruct the path or iterate ancestors if available.
            // DataViewMatrixNode doesn't have 'parent'.
            // However, we can just show the current node's label and its level name.
            if (this.lastMatrix && this.lastMatrix.rows && this.lastMatrix.rows.levels && rowNode.level !== undefined) {
                const levelName = this.lastMatrix.rows.levels[rowNode.level].sources[0].displayName;
                res.push({
                    displayName: levelName,
                    value: this.nodeLabel(rowNode)
                });
            } else {
                 res.push({
                    displayName: "Row",
                    value: this.nodeLabel(rowNode)
                });
            }
        }

        // Columns
        // If we had the column node... (not passed yet, need to enhance later if crucial)
        // For now, simple measure
        if (this.lastMatrix && this.lastMatrix.valueSources && this.lastMatrix.valueSources[measureIndex]) {
            res.push({
                displayName: this.lastMatrix.valueSources[measureIndex].displayName,
                value: this.formatValueByMeasure(value, measureIndex)
            });
        } else {
             res.push({
                displayName: "Value",
                value: String(value)
            });
        }

        return res;
    }
}
