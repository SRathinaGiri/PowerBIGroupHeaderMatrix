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
import { createTooltipServiceWrapper, ITooltipServiceWrapper } from "powerbi-visuals-utils-tooltiputils";
import * as d3 from "d3-selection";
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.extensibility.ISelectionId;
import IVisualEventService = powerbi.extensibility.IVisualEventService;

import { VisualFormattingSettingsModel, MeasureHeaderInfo } from "./settings";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

type DataView = powerbi.DataView;
type DataViewMatrix = powerbi.DataViewMatrix;
type DataViewMatrixNode = powerbi.DataViewMatrixNode;

type DataViewHierarchy = powerbi.DataViewHierarchy;
type ValuesMap = { [key: number]: powerbi.DataViewMatrixNodeValue };

const KEY_SEP = "||";
const BLANK_LABEL = "(Blank)";

/** One rendered data column: exactly one measure for one (possibly collapsed) column group. */
interface DisplayCol {
    kind: "leaf" | "collapsed";
    /** Index into a row node's values map; undefined renders a blank cell. */
    offset?: number;
    /** Index into matrix.valueSources. */
    measureIndex: number;
    /** Header label per header level. */
    labels: string[];
    /** Header merge/toggle key per header level. */
    keys: string[];
    /** Stable identity of this column (widths, sorting). */
    uniqueKey: string;
    /** valueSource index -> values-map offset for every measure of the same column group. */
    colorOffsets: Map<number, number>;
    collapsedLevel?: number;
}

interface HeaderCell {
    label: string;
    span: number;
    key: string;
    togglable: boolean;
    collapsed: boolean;
    /** Set when this header cell identifies exactly one data column that can be sorted. */
    sortKey?: string;
    /** Set on measure-name cells: index into matrix.valueSources. */
    measureIndex?: number;
}

interface HeaderRow {
    level: number;
    cells: HeaderCell[];
}

interface ColumnLayout {
    displayCols: DisplayCol[];
    /** Number of header levels (group levels plus the measure-name level when shown). */
    depth: number;
    groupDepth: number;
    /** valueSource index -> every values-map offset of that measure, grand-total column first. */
    measureOffsets: Map<number, number[]>;
}

interface RowStyle {
    bg: string;
    color: string;
    bold: boolean;
}

interface SortState {
    type: "value" | "label";
    colKey: string; // DisplayCol.uniqueKey for value sort
    level: number; // row level for label sort (-1 = every level)
    direction: "ASC" | "DESC";
}

export class Visual implements IVisual {
    private root: HTMLElement;
    private container: HTMLElement; // scroll container
    private toolbar: HTMLElement;
    private contentHost: HTMLElement;
    private debugEl: HTMLElement;
    private debugEnabled: boolean = false;
    private debugFromSetting: boolean = false;
    private table: HTMLTableElement | null = null;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    // Collapse/expand state
    private collapsedRowKeys: Set<string> = new Set();
    private collapsedColKeys: Set<string> = new Set();
    private lastMatrix: DataViewMatrix | null = null;
    /** Last rendered layout and body rows, kept for the Ctrl+Alt+D diagnostics panel. */
    private lastDebug: { cols: DisplayCol[]; rows: any[] } | null = null;
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
    /** Measures that are never displayed: they only style cells or rows. */
    private static readonly STYLE_ROLES = ["cellBgColor", "cellFontColor", "rowBgColor", "rowFontColor", "rowBold"];
    private rowBgMeasureIndex: number = -1;
    private rowFontMeasureIndex: number = -1;
    private rowBoldMeasureIndex: number = -1;
    private rowFormatApplyToHeaders: boolean = true;
    /** True when the row style is applied after (over) the cell conditional style. */
    private rowFormatWins: boolean = false;
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
    /** Values-map offset of the column currently used for value sorting (undefined = none). */
    private activeSortOffset: number | undefined = undefined;
    private formatterCache: Map<string, ReturnType<typeof valueFormatter.create>> = new Map();
    private virtualRowHeight: number = 25;
    /** Stored "measureHeader" object per valueSource index (only user-set properties). */
    private measureHeaderStyles: any[] = [];
    private align = {
        rowHeaderH: "left", rowHeaderV: "middle",
        colHeaderH: "center", colHeaderV: "middle",
        valuesH: "right", valuesV: "middle"
    };

    private viewMode: string = "normal";
    private currentPage: number = 1;
    private pageSize: number = 100;
    private scrollListener: ((e: Event) => void) | null = null;

    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private events: IVisualEventService;
    private resizeObserver: ResizeObserver | null = null;
    private selectedIds: ISelectionId[] = [];
    private isHighContrast: boolean = false;
    private highContrastForeground: string = "";
    private highContrastBackground: string = "";
    private highContrastSelected: string = "";
    private highContrastSeparator: string = "";
    private allowInteractions: boolean = true;
    // Totals behavior
    private showGrandTotalRows: boolean = true;
    private showGrandTotalColumns: boolean = true;
    private rowSubtotalsEnabled: boolean = true;
    private columnSubtotalsEnabled: boolean = true;
    private rowSubtotalPosition: "Top" | "Bottom" = "Top";
    private grandTotalPosition: "Top" | "Bottom" = "Bottom";

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipServiceWrapper = createTooltipServiceWrapper(options.host.tooltipService, options.element);
        this.events = options.host.eventService;
        this.allowInteractions = options.host.hostCapabilities?.allowInteractions !== false;
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            this.selectedIds = ids || [];
            this.refresh();
        });

        // Root flex container
        this.root = document.createElement("div");
        this.root.className = "ghm-root";
        this.root.tabIndex = 0;
        this.root.setAttribute("role", "region");
        this.root.setAttribute("aria-label", "Group Header Matrix visual");
        this.root.addEventListener("contextmenu", (e) => this.showVisualContextMenu(e));
        // Hidden diagnostics: Ctrl+Alt+D shows the data shape Power BI sent
        this.root.addEventListener("keydown", (e) => {
            if (e.ctrlKey && e.altKey && (e.key === "d" || e.key === "D")) {
                e.preventDefault();
                this.debugEnabled = !this.debugEnabled;
                this.updateDebugOverlay();
            }
        });

        // Toolbar (fixed at top)
        this.toolbar = document.createElement("div");
        this.toolbar.className = "ghm-toolbar";
        this.toolbar.setAttribute("role", "toolbar");
        this.toolbar.setAttribute("aria-label", "Matrix controls");

        // Scroll container (occupies remaining space)
        this.container = document.createElement("div");
        this.container.className = "ghm-container"; // keeps existing styles for scroll behavior

        const btnExpand = document.createElement("button");
        btnExpand.className = "ghm-btn";
        btnExpand.type = "button";
        btnExpand.textContent = "➕ Expand All";
        btnExpand.setAttribute("aria-label", "Expand all groups");
        btnExpand.addEventListener("click", () => { this.expandAll(); this.persistState(); this.refresh(); });
        const btnCollapse = document.createElement("button");
        btnCollapse.className = "ghm-btn";
        btnCollapse.type = "button";
        btnCollapse.textContent = "➖ Collapse All";
        btnCollapse.setAttribute("aria-label", "Collapse all groups");
        btnCollapse.addEventListener("click", () => { this.collapseAll(); this.persistState(); this.refresh(); });

        const btnRowExpandLevel = document.createElement("button");
        btnRowExpandLevel.className = "ghm-btn";
        btnRowExpandLevel.type = "button";
        btnRowExpandLevel.title = "Expand Row Level";
        btnRowExpandLevel.setAttribute("aria-label", "Expand row hierarchy one level");
        btnRowExpandLevel.addEventListener("click", () => { this.expandRowLevel(); this.persistState(); this.refresh(); });
        // Use requested glyph-only label
        btnRowExpandLevel.textContent = "+ 𝄘";
        const btnRowCollapseLevel = document.createElement("button");
        btnRowCollapseLevel.className = "ghm-btn";
        btnRowCollapseLevel.type = "button";
        btnRowCollapseLevel.title = "Collapse Row Level";
        btnRowCollapseLevel.setAttribute("aria-label", "Collapse row hierarchy one level");
        btnRowCollapseLevel.addEventListener("click", () => { this.collapseRowLevel(); this.persistState(); this.refresh(); });
        btnRowCollapseLevel.textContent = "- 𝄘";
        const btnColExpandLevel = document.createElement("button");
        btnColExpandLevel.className = "ghm-btn";
        btnColExpandLevel.type = "button";
        btnColExpandLevel.title = "Expand Column Level";
        btnColExpandLevel.setAttribute("aria-label", "Expand column hierarchy one level");
        btnColExpandLevel.addEventListener("click", () => { this.expandColLevel(); this.persistState(); this.refresh(); });
        btnColExpandLevel.textContent = "+ ⦀";
        const btnColCollapseLevel = document.createElement("button");
        btnColCollapseLevel.className = "ghm-btn";
        btnColCollapseLevel.type = "button";
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
        btnPrevPage.type = "button";
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
        btnNextPage.type = "button";
        btnNextPage.textContent = ">";
        btnNextPage.title = "Next Page";
        btnNextPage.addEventListener("click", () => {
            // renderMatrix clamps the page and disables this button on the last page.
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
        btnCollapseAll.type = "button";
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
        btnDebug.type = "button";
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
        this.events.renderingStarted(options);
        this.allowInteractions = this.host.hostCapabilities?.allowInteractions !== false;
        const dataView: DataView | undefined = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        this.updateHighContrastState();
        // Apply sticky preference if present
        const sticky = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "stickyHeaders", true);
        this.container.classList.toggle("ghm-sticky", !!sticky);
        this.root.classList.toggle("ghm-high-contrast", this.isHighContrast);
        this.repeatLabels = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "repeatLabels", false);
        this.compactLayout = this.getObjectValue<boolean>(dataView?.metadata?.objects, "state", "compactLayout", true);
        if (this.compactLayout) this.repeatLabels = false;

        this.viewMode = this.getObjectValue<string>(dataView?.metadata?.objects, "state", "viewMode", "normal") || "normal";
        this.currentPage = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "currentPage", 1) || 1;
        this.pageSize = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "pageSize", 100) || 100;

        // Grand total formatting options
        this.showGrandTotalRows = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grandTotal", "showRows", true);
        this.showGrandTotalColumns = this.getObjectValue<boolean>(dataView?.metadata?.objects, "grandTotal", "showColumns", true);
        this.rowSubtotalsEnabled = this.getObjectValue<boolean>(dataView?.metadata?.objects, "subtotal", "showRowSubtotals", true);
        this.columnSubtotalsEnabled = this.getObjectValue<boolean>(dataView?.metadata?.objects, "subtotal", "showColumnSubtotals", true);
        const rowSubtotalPosRaw = this.getObjectValue<any>(dataView?.metadata?.objects, "subtotal", "rowSubtotalsType", { value: "Top" } as any);
        const rowSubtotalPos = typeof rowSubtotalPosRaw === "string" ? rowSubtotalPosRaw : (rowSubtotalPosRaw && (rowSubtotalPosRaw as any).value);
        this.rowSubtotalPosition = rowSubtotalPos === "Bottom" ? "Bottom" : "Top";
        const grandTotalPosRaw = this.getObjectValue<any>(dataView?.metadata?.objects, "grandTotal", "position", { value: "Bottom" } as any);
        const grandTotalPos = typeof grandTotalPosRaw === "string" ? grandTotalPosRaw : (grandTotalPosRaw && (grandTotalPosRaw as any).value);
        this.grandTotalPosition = grandTotalPos === "Top" ? "Top" : "Bottom";
        this.rowHeaderMinWidth = this.getObjectValue<number>(dataView?.metadata?.objects, "state", "rowHeaderMinWidth", 160) || 160;
        this.rowHeaderFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "rowHeaderFontSize", 12) || 12;
        this.rowHeaderFontFamily = this.readText(dataView?.metadata?.objects, "labels", "rowHeaderFontFamily");
        this.rowHeaderBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "rowHeaderBold", false);
        this.colHeaderFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "colHeaderFontSize", 11) || 11;
        this.colHeaderFontFamily = this.readText(dataView?.metadata?.objects, "labels", "colHeaderFontFamily");
        this.colHeaderBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "colHeaderBold", false);
        this.dataFontSize = this.getObjectValue<number>(dataView?.metadata?.objects, "labels", "dataFontSize", 11) || 11;
        this.dataFontFamily = this.readText(dataView?.metadata?.objects, "labels", "dataFontFamily");
        this.dataBold = this.getObjectValue<boolean>(dataView?.metadata?.objects, "labels", "dataBold", false);

        // Per-measure header styles: one Format pane group per displayed measure
        this.measureHeaderStyles = [];
        const headerMeasures: MeasureHeaderInfo[] = [];
        ((dataView && dataView.matrix && dataView.matrix.valueSources) || []).forEach((vs: any, i: number) => {
            const stored = (vs.objects && vs.objects.measureHeader) || {};
            this.measureHeaderStyles[i] = stored;
            const roles = vs.roles || {};
            if (roles.measure && vs.queryName) {
                headerMeasures.push({ queryName: vs.queryName, displayName: vs.displayName || vs.queryName, stored });
            }
        });
        this.formattingSettings.measureHeaderCard.build(headerMeasures, {
            fontFamily: this.colHeaderFontFamily || "Segoe UI",
            fontSize: this.colHeaderFontSize || 11,
            bold: !!this.colHeaderBold
        }, (v) => this.parseColor(v));

        const objs = dataView?.metadata?.objects;
        this.debugFromSetting = this.getObjectValue<boolean>(objs, "state", "showDiagnostics", false) === true;
        this.rowFormatApplyToHeaders = this.getObjectValue<boolean>(objs, "rowFormatting", "applyToRowHeaders", true) !== false;
        this.rowFormatWins = this.pickEnum(this.getObjectValue<any>(objs, "rowFormatting", "precedence", "cell"), ["cell", "row"], "cell") === "row";
        const hAlign = (prop: string, def: string) => this.pickEnum(this.getObjectValue<any>(objs, "alignment", prop, def), ["left", "center", "right"], def);
        const vAlign = (prop: string) => this.pickEnum(this.getObjectValue<any>(objs, "alignment", prop, "middle"), ["top", "middle", "bottom"], "middle");
        this.align = {
            rowHeaderH: hAlign("rowHeaderHAlign", "left"), rowHeaderV: vAlign("rowHeaderVAlign"),
            colHeaderH: hAlign("colHeaderHAlign", "center"), colHeaderV: vAlign("colHeaderVAlign"),
            valuesH: hAlign("valuesHAlign", "right"), valuesV: vAlign("valuesVAlign")
        };

        const themePresetObj =this.getObjectValue<any>(dataView?.metadata?.objects, "theme", "preset", { value: "Default" } as any);
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
            this.renderLandingPage();
            this.events.renderingFinished(options);
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
            if (this.pruneInvalidCollapsedState(this.lastMatrix)) {
                this.persistState();
            }
            this.loadLevelStyles(this.lastMatrix);
            this.renderMatrix(dataView.matrix);
            this.events.renderingFinished(options);
        } catch (e) {
            this.renderPlaceholder("Unable to render matrix");
            this.events.renderingFailed(options, e instanceof Error ? e.message : "Unable to render matrix");
            console.error(e);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const model = this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
        // Per-measure groups repeat the same property names; the service only de-duplicates slice
        // uids within one group, so make them unique across groups (values are saved by selector).
        const seen = new Set<string>();
        for (const card of model.cards || []) {
            for (const group of (card as any).groups || []) {
                for (const slice of group.slices || []) {
                    let uid = slice.uid;
                    for (let n = 1; seen.has(uid); n++) uid = `${slice.uid}-g${n}`;
                    slice.uid = uid;
                    seen.add(uid);
                }
            }
        }
        return model;
    }

    public enumerateObjectInstances(options: powerbi.EnumerateVisualObjectInstancesOptions): powerbi.VisualObjectInstanceEnumeration {
        const enumeration: powerbi.VisualObjectInstance[] = [];
        if (!this.lastMatrix) return enumeration;

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

    private renderLandingPage() {
        const page = document.createElement("div");
        page.className = "ghm-landing";
        page.setAttribute("role", "group");
        page.setAttribute("aria-label", "Group Header Matrix setup instructions");

        const title = document.createElement("h2");
        title.textContent = "Group Header Matrix";
        page.appendChild(title);

        const intro = document.createElement("p");
        intro.textContent = "Build an Excel-like pivot matrix with grouped row headers, grouped column headers, expand and collapse controls, subtotals, sorting, and pagination.";
        page.appendChild(intro);

        const stepsTitle = document.createElement("h3");
        stepsTitle.textContent = "Setup";
        page.appendChild(stepsTitle);

        const steps = document.createElement("ol");
        [
            "Add one or more fields to Rows.",
            "Add one or more fields to Columns.",
            "Add one or more measures to Values.",
            "Optional: add measures that return color strings to Cell Background Color or Cell Font Color.",
            "Optional: style whole rows with Row Background Color, Row Font Color and Row Bold measures (cell colors are applied on top).",
            "Use the visual toolbar to expand, collapse, switch compact layout, and enable pagination or virtualization."
        ].forEach(stepText => {
            const li = document.createElement("li");
            li.textContent = stepText;
            steps.appendChild(li);
        });
        page.appendChild(steps);

        const roles = document.createElement("div");
        roles.className = "ghm-landing-roles";
        [
            ["Rows", "Row hierarchy displayed down the left side."],
            ["Columns", "Column hierarchy displayed across the top."],
            ["Values", "Measures rendered in the matrix cells."],
            ["Cell colors", "Optional measure-driven color values."]
        ].forEach(([name, description]) => {
            const item = document.createElement("div");
            const heading = document.createElement("strong");
            heading.textContent = name;
            const text = document.createElement("span");
            text.textContent = description;
            item.appendChild(heading);
            item.appendChild(text);
            roles.appendChild(item);
        });
        page.appendChild(roles);

        this.contentHost.appendChild(page);
    }

    private renderMatrix(matrix: DataViewMatrix) {
        const columns = matrix.columns;
        const rows = matrix.rows;

        if (!columns || !columns.root || !columns.root.children || columns.root.children.length === 0) {
            this.renderLandingPage();
            return;
        }

        const table = document.createElement("table");
        table.className = "ghm-table";
        table.setAttribute("role", "grid");
        table.setAttribute("aria-label", "Group Header Matrix");
        table.style.borderCollapse = "collapse";
        // Apply a global font if only one family provided
        if (this.rowHeaderFontFamily && !this.colHeaderFontFamily) table.style.fontFamily = this.rowHeaderFontFamily;
        if (this.colHeaderFontFamily && !this.rowHeaderFontFamily) table.style.fontFamily = this.colHeaderFontFamily;
        const colgroup = document.createElement("colgroup");
        const thead = document.createElement("thead");
        const tbody = document.createElement("tbody");

        // Measures handling: compute display measures vs style measures
        this.displayMeasureIndices = [];
        this.cellBgColorMeasureIndices = [];
        this.cellFontColorMeasureIndices = [];
        this.rowBgMeasureIndex = -1;
        this.rowFontMeasureIndex = -1;
        this.rowBoldMeasureIndex = -1;
        this.measureFontColorSettings = [];
        this.measureBgColorSettings = [];
        const vsArr = matrix.valueSources || [];
        for (let i = 0; i < vsArr.length; i++) {
            const roles: any = (vsArr[i] as any).roles || {};
            if (roles.measure) this.displayMeasureIndices.push(i);
            if (roles.cellBgColor) this.cellBgColorMeasureIndices.push(i);
            if (roles.cellFontColor) this.cellFontColorMeasureIndices.push(i);
            if (roles.rowBgColor && this.rowBgMeasureIndex < 0) this.rowBgMeasureIndex = i;
            if (roles.rowFontColor && this.rowFontMeasureIndex < 0) this.rowFontMeasureIndex = i;
            if (roles.rowBold && this.rowBoldMeasureIndex < 0) this.rowBoldMeasureIndex = i;
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
                    return roles.measure || !Visual.STYLE_ROLES.some(r => roles[r]);
                });
            // Only when the host sent no role information at all: show everything
            const anyRoles = vsArr.some(vs => Object.keys((vs as any).roles || {}).length > 0);
            if (!this.displayMeasureIndices.length && !anyRoles) {
                this.displayMeasureIndices = vsArr.map((_, idx) => idx);
            }
        }
        this.measureFormats = vsArr.map(m => String((m as any).format || ""));

        // Column layout: one display column per (column group, display measure)
        const layout = this.computeColumnLayout(matrix);
        const displayCols = layout.displayCols;
        const colDepth = layout.depth;
        const headerRows = this.buildHeaderRows(displayCols, layout);

        // Determine number of row header levels (depth)
        const rowDepth = this.getRowDepth(rows);
        // Hide deeper row header columns when levels are collapsed, even in Repeat Labels mode
        const visibleRowDepth = Math.max(1, Math.min(rowDepth, this.getDisplayedDepth("rows")));
        const rowHeaderCols = (rowDepth > 0) ? (this.compactLayout ? 1 : visibleRowDepth) : 0;
        // Include the thin resizer row as part of the sticky header block
        const headerDepth = headerRows.length + 1;

        // Calculate cumulative widths for row header columns to sticky-position them correctly
        const rowHeaderWidths: number[] = [];
        if (rowHeaderCols > 0) {
            rowHeaderWidths.push(this.rowHeaderMinWidth);
            if (!this.compactLayout) {
                for (let k = 1; k < rowHeaderCols; k++) rowHeaderWidths.push(120);
            }
        }

        const appendCorner = (tr: HTMLTableRowElement) => {
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
                // The corner sits above the row headers and beside the column headers
                this.setAlignment(corner, this.align.rowHeaderH, this.align.colHeaderV);

                if (this.rowHeaderBg) corner.style.backgroundColor = this.rowHeaderBg;
                if (this.rowHeaderColor) corner.style.color = this.rowHeaderColor;

                // Compact layout shows every level in one column, so its label sort applies to all levels
                this.attachSortHandler(corner, "", this.compactLayout ? -1 : r, "label");

                this.applyGridBorder(corner, true);
                tr.appendChild(corner);
            }
        };

        headerRows.forEach((headerRow, i) => {
            const tr = document.createElement("tr");
            if (i === 0) appendCorner(tr);
            for (const cell of headerRow.cells) {
                const th = document.createElement("th");
                th.className = "ghm-colheader";
                th.title = cell.label;
                this.setAlignment(th, this.align.colHeaderH, this.align.colHeaderV);
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
                const ch = this.colLevelStyles[headerRow.level] || {};
                const fs = ch.fontSize ?? this.colHeaderFontSize; if (fs) txt.style.fontSize = `${fs}px`;
                txt.style.fontFamily = (ch.fontFamily || this.colHeaderFontFamily) || "";
                if (this.colHeaderBold || ch.fontWeight) txt.style.fontWeight = (ch.fontWeight || "bold");
                if (ch.italic) txt.style.fontStyle = "italic";
                if (ch.underline) txt.style.textDecoration = "underline";
                if (this.colHeaderColor) th.style.color = this.colHeaderColor;
                if (this.colHeaderBg) th.style.backgroundColor = this.colHeaderBg;
                if (cell.measureIndex !== undefined) this.applyMeasureHeaderStyle(th, txt, cell.measureIndex);
                this.applyGridBorder(th, true);
                th.appendChild(txt);
                if (cell.span > 1) th.colSpan = cell.span;

                if (cell.sortKey) this.attachSortHandler(th, cell.sortKey);

                tr.appendChild(th);
            }
            thead.appendChild(tr);
        });

        // Resizer row (always present so handles align with columns)
        const resizerRow = document.createElement("tr");
        resizerRow.className = "ghm-resizers-row";
        if (headerRows.length === 0) appendCorner(resizerRow);
        else {
            // Prepend row-header placeholders so cell count matches total columns
            let accumulatedResizerLeft = 0;
            for (let r = 0; r < rowHeaderCols; r++) {
                const th = document.createElement("th");
                th.className = "ghm-resizer-cell";
                // In tabular layout, sticky row headers must stack horizontally
                if (!this.compactLayout) {
                    th.style.left = `${accumulatedResizerLeft}px`;
                    accumulatedResizerLeft += (rowHeaderWidths[r] || 120);
                } else {
                    th.style.left = "0px";
                }
                resizerRow.appendChild(th);
            }
        }
        for (const ref of displayCols) {
            const th = document.createElement("th");
            th.className = "ghm-resizer-cell";
            const handle = document.createElement("div");
            handle.className = "ghm-resizer";
            handle.title = "Drag to resize column";
            handle.addEventListener("mousedown", (e) => this.beginResize(e as MouseEvent, ref.uniqueKey));
            th.appendChild(handle);
            this.applyGridBorder(th, true);
            resizerRow.appendChild(th);
        }
        thead.appendChild(resizerRow);

        // Build colgroup to control widths
        this.colElsByKey.clear();
        for (let r = 0; r < rowHeaderCols; r++) {
            const col = document.createElement("col");
            col.style.width = `${rowHeaderWidths[r] || 120}px`;
            colgroup.appendChild(col);
        }
        for (const ref of displayCols) {
            const width = this.columnWidthPx.get(ref.uniqueKey) ?? 120;
            const col = document.createElement("col");
            col.style.width = `${width}px`;
            colgroup.appendChild(col);
            if (!this.colElsByKey.has(ref.uniqueKey)) this.colElsByKey.set(ref.uniqueKey, []);
            this.colElsByKey.get(ref.uniqueKey)!.push(col);
        }

        const colLeafCount = displayCols.length;
        const columnLeaves = displayCols;

        // Resolve the values-map offset used for value sorting. Sorting never mutates the
        // host data view; collectOutlineRowsWithTotals orders children on the fly.
        this.activeSortOffset = undefined;
        if (this.sortState && this.sortState.type === "value") {
            const sortCol = displayCols.find(c => c.uniqueKey === this.sortState!.colKey);
            if (sortCol && sortCol.offset !== undefined) this.activeSortOffset = sortCol.offset;
        }

        // Build body rows. Body rows (including the grand total) are generated
        // exclusively by collectOutlineRowsWithTotals() to avoid duplicated totals.
        let rowsToRender: any[] = [];
        if (rowDepth === 0) {
            // No row fields: a single row holding the column totals
            const root = rows && rows.root;
            const firstChild = root && root.children && root.children[0];
            const valuesMap = (firstChild && firstChild.values) || (root && root.values) || {};
            rowsToRender = [{ labels: [], valuesMap, isTotal: false }];
        } else if (rows && rows.root && rows.root.children && rows.root.children.length) {
            rowsToRender = this.collectOutlineRowsWithTotals(rows.root, rowDepth);
        }

        const lastShownRowLabels: Array<string | null> = new Array(rowDepth).fill(null);

        // Pre-process rows for sliceability (pagination and virtualization)
        let bodyRowIndex = 0;
        this.lastDebug = { cols: displayCols, rows: rowsToRender };
        const preprocessedRows = rowsToRender.map((rowInfoRaw) => {
             const rowInfo = { ...rowInfoRaw } as any;
             const valuesMapForRow = rowInfo.valuesMap || {} as any;
             const rowMeasureBg = this.getRowMeasureBg(valuesMapForRow as ValuesMap, displayCols);
             let labelsToUse: string[] = rowInfo.labels as string[];

             if (!this.compactLayout && !this.repeatLabels && !rowInfo.isTotal && this.rowSubtotalsEnabled) {
                 if (this.rowSubtotalPosition === "Bottom") {
                     const newLabels = new Array(rowDepth).fill("");
                     for (let lvl = 0; lvl < rowDepth; lvl++) {
                         const lbl = (rowInfo.labels && rowInfo.labels[lvl]) ? rowInfo.labels[lvl] : "";
                         if (!lbl) { newLabels[lvl] = ""; continue; }
                         const hasToggleAtLevel = !!(rowInfo.toggles && rowInfo.toggles[lvl]) || (rowInfo.toggleKey && rowInfo.depth === lvl);
                         if (hasToggleAtLevel) {
                             newLabels[lvl] = lbl;
                             lastShownRowLabels[lvl] = lbl;
                             for (let deeper = lvl + 1; deeper < rowDepth; deeper++) lastShownRowLabels[deeper] = null;
                         } else if (lastShownRowLabels[lvl] === lbl) {
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
             rowInfo.rowStyle = this.getRowStyle(valuesMapForRow as ValuesMap, layout);
             if (!rowInfo.isTotal) {
                 rowInfo.bodyRowIndex = bodyRowIndex;
                 bodyRowIndex++;
             }
             return rowInfo;
        });

        const renderBody = (rowsToRenderSlice: any[]) => {
            while (tbody.firstChild) tbody.removeChild(tbody.firstChild);

            // Selection state is the same for every cell: compute it once per render
            const currentSelection = this.selectedIds.length ? this.selectedIds : this.selectionManager.getSelectionIds();
            const hasSelection = currentSelection.length > 0 || this.selectionManager.hasSelection();

            for (const rowInfo of rowsToRenderSlice) {
                const labelsToUse = rowInfo.labelsToUse;
                const rowMeasureBg = rowInfo.rowMeasureBg;
                const rowStyle: RowStyle | null = rowInfo.rowStyle;

                const tr = document.createElement("tr");
                if (rowInfo.isTotal) tr.className = "ghm-totalrow";
                let zebraRowColor = "";
                if (this.zebraEnabled && !rowInfo.isTotal && !rowMeasureBg && !(rowStyle && rowStyle.bg)) {
                    zebraRowColor = (rowInfo.bodyRowIndex % 2 === 0) ? this.zebraEvenColor : this.zebraOddColor;
                    if (zebraRowColor) tr.style.backgroundColor = zebraRowColor;
                }

                let rowSelectionId: ISelectionId | undefined;
                let rowIsSelected = false;
                if (!rowInfo.isTotal && this.allowInteractions && rowInfo.node) {
                    rowSelectionId = this.createSelectionId(rowInfo.node);
                    rowIsSelected = currentSelection.some(id => this.selectionIdsMatch(id, rowSelectionId!));
                }
            // Row header: base colors, then the measure colors, then the whole-row style (if enabled)
            const headerRowStyle = (this.rowFormatApplyToHeaders && rowStyle) ? rowStyle : null;
            const rowStyleBg = (headerRowStyle && headerRowStyle.bg) || rowMeasureBg || this.rowHeaderBg || "";
            const rowStyleColor = (headerRowStyle && headerRowStyle.color) || this.rowHeaderColor || "";

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
                    if (headerRowStyle && headerRowStyle.bold) th.style.fontWeight = "bold";
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
                        if (headerRowStyle && headerRowStyle.bold) th.style.fontWeight = "bold";
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

            const valuesMap: ValuesMap = (rowInfo as any).valuesMap || {};
            const rowAriaLabel = rowSelectionId ? this.getAccessibleRowLabel(rowInfo) : "";
            for (let c = 0; c < colLeafCount; c++) {
                const col = columnLeaves[c];
                const globalM = col.measureIndex;
                const td = document.createElement("td");
                this.setAlignment(td, this.align.valuesH, this.align.valuesV);
                const v = this.getCellValue(valuesMap, col);
                td.textContent = this.formatValueByMeasure(v, globalM, this.getCellFormat(valuesMap, col));
                if (this.dataFontSize) td.style.fontSize = `${this.dataFontSize}px`;
                if (this.dataFontFamily) td.style.fontFamily = this.dataFontFamily;
                if (this.dataBold) td.style.fontWeight = "bold";
                if (this.cellColor) td.style.color = this.cellColor;
                // Base background: column stripe, then row stripe, then cell background
                let baseBg = "";
                if (this.zebraColEnabled) baseBg = (c % 2 === 0) ? this.zebraColEvenColor : this.zebraColOddColor;
                if (!baseBg) baseBg = zebraRowColor;
                if (!baseBg && this.cellBg) baseBg = this.cellBg;
                if (baseBg) td.style.backgroundColor = baseBg;
                // Row and cell conditional styles sit on the base styles; whichever is applied
                // last wins where both set the same property (Row formatting > precedence).
                const applyRow = () => {
                    if (!rowStyle) return;
                    if (rowStyle.bg) td.style.backgroundColor = rowStyle.bg;
                    if (rowStyle.color) td.style.color = rowStyle.color;
                    if (rowStyle.bold) td.style.fontWeight = "bold";
                };
                if (this.rowFormatWins) {
                    this.applyCellConditionalStyle(td, valuesMap, col);
                    applyRow();
                } else {
                    applyRow();
                    this.applyCellConditionalStyle(td, valuesMap, col);
                }
                this.applyGridBorder(td, false);

                // Interaction (Multi-row mode)
                if (rowSelectionId) {
                    const selectionId = rowSelectionId;
                    td.style.opacity = (hasSelection && !rowIsSelected) ? "0.5" : "1";
                    td.tabIndex = 0;
                    td.setAttribute("aria-selected", String(rowIsSelected));
                    td.setAttribute("aria-label", `${rowAriaLabel} ${td.textContent || ""}`.trim());

                    td.addEventListener("click", (e) => {
                        this.selectionManager.select(selectionId, e.ctrlKey || e.metaKey).then(() => {
                            this.selectedIds = this.selectionManager.getSelectionIds();
                            this.refresh();
                        });
                        e.stopPropagation();
                    });
                    td.addEventListener("keydown", (e) => {
                        if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            this.selectionManager.select(selectionId, e.ctrlKey || e.metaKey).then(() => {
                                this.selectedIds = this.selectionManager.getSelectionIds();
                                this.refresh();
                            });
                        }
                    });
                    td.addEventListener("contextmenu", (e) => {
                        this.selectionManager.showContextMenu(selectionId, {x: e.clientX, y: e.clientY});
                        e.preventDefault();
                        e.stopPropagation();
                    });
                }

                this.tooltipServiceWrapper.addTooltip(d3.select(td) as any, () => {
                    return this.getTooltipData(v, rowInfo, col, globalM);
                }, () => rowSelectionId || this.createEmptySelectionId());

                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
    }; // end of renderBody

    if (this.scrollListener) {
        this.container.removeEventListener("scroll", this.scrollListener);
        this.scrollListener = null;
    }

    let renderAfterAttach: (() => void) | null = null;
    if (this.viewMode === "pagination") {
        const totalPages = Math.ceil(preprocessedRows.length / this.pageSize) || 1;
        if (this.currentPage > totalPages) this.currentPage = totalPages;
        if (this.currentPage < 1) this.currentPage = 1;
        const start = (this.currentPage - 1) * this.pageSize;
        renderBody(preprocessedRows.slice(start, start + this.pageSize));
        const lblPage = this.toolbar.querySelector('#ghm-lbl-page') as HTMLElement | null;
        if (lblPage) lblPage.textContent = `Page ${this.currentPage} of ${totalPages}`;
        const btnNextPage = this.toolbar.querySelector('#ghm-btn-nextpage') as HTMLButtonElement | null;
        if (btnNextPage) btnNextPage.disabled = this.currentPage >= totalPages;
        const btnPrevPage = this.toolbar.querySelector('#ghm-btn-prevpage') as HTMLButtonElement | null;
        if (btnPrevPage) btnPrevPage.disabled = this.currentPage <= 1;
    } else if (this.viewMode === "virtualization") {
        const totalColumns = rowHeaderCols + colLeafCount;
        const makeSpacer = (height: number, className: string): HTMLTableRowElement => {
            // A row without cells has no height in some browsers, so the spacer carries one empty cell
            const tr = document.createElement("tr");
            tr.className = className;
            const td = document.createElement("td");
            td.colSpan = Math.max(1, totalColumns);
            td.style.height = `${height}px`;
            td.style.padding = "0";
            td.style.border = "none";
            tr.appendChild(td);
            return tr;
        };
        let measuredOnce = false;
        const renderVirtual = () => {
            const rowHeight = this.virtualRowHeight;
            const headerHeight = thead.getBoundingClientRect().height || 0;
            const scrollTop = Math.max(0, this.container.scrollTop - headerHeight);
            const clientHeight = this.container.clientHeight;

            // Add buffer rows to avoid rapid re-rendering
            const buffer = 10;
            const startIdx = Math.max(0, Math.floor(scrollTop / rowHeight) - buffer);
            const visibleCount = Math.ceil(clientHeight / rowHeight) + (buffer * 2);
            const endIdx = Math.min(preprocessedRows.length, startIdx + visibleCount);

            renderBody(preprocessedRows.slice(startIdx, endIdx));
            this.applyHighContrastStyles(tbody);

            // Measure the real row height (fonts and padding change it) and re-render once if the estimate was off
            const rendered = Array.from(tbody.rows);
            if (rendered.length && !measuredOnce) {
                measuredOnce = true;
                const measured = rendered.reduce((sum, r) => sum + r.getBoundingClientRect().height, 0) / rendered.length;
                if (measured > 0 && Math.abs(measured - rowHeight) > 0.5) {
                    this.virtualRowHeight = measured;
                    requestAnimationFrame(renderVirtual);
                }
            }

            const topSpacerHeight = startIdx * rowHeight;
            const bottomSpacerHeight = (preprocessedRows.length - endIdx) * rowHeight;
            if (topSpacerHeight > 0) tbody.insertBefore(makeSpacer(topSpacerHeight, "ghm-virtual-spacer-top"), tbody.firstChild);
            if (bottomSpacerHeight > 0) tbody.appendChild(makeSpacer(bottomSpacerHeight, "ghm-virtual-spacer-bottom"));
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
        // Rows can only be measured once the table is in the document
        renderAfterAttach = renderVirtual;
    } else {
        renderBody(preprocessedRows);
    }

        table.appendChild(colgroup);
        table.appendChild(thead);
        table.appendChild(tbody);
        this.contentHost.appendChild(table);
        this.table = table;
        if (renderAfterAttach) renderAfterAttach();
        this.applyHighContrastStyles(table);

        this.applyStickyOffsets(thead);
        this.updateDebugOverlay({ table, headerRows, measureCount: this.displayMeasureIndices.length, columnLeaves, rowDepth, colDepth });
    }

    /**
     * Builds the list of rendered data columns from the column hierarchy.
     *
     * Power BI keys every row node's values map by column-leaf position, and when more than one
     * measure is bound it adds a measure level as the last column level (one leaf per measure).
     * Each display column therefore points at exactly one leaf offset. Measures bound to the
     * colour roles are never displayed; their offsets are kept in colorOffsets so cells can look
     * up the colour value of the same column group.
     */
    private computeColumnLayout(matrix: DataViewMatrix): ColumnLayout {
        const columns = matrix.columns;
        const vsArr = matrix.valueSources || [];
        const levels = (columns && columns.levels) || [];
        const measureLevelIdx = levels.findIndex(lvl => this.isMeasureLevel(lvl));
        const hasMeasureLevel = measureLevelIdx >= 0;
        const groupDepth = levels.filter(lvl => !this.isMeasureLevel(lvl)).length;
        // Fallback for hosts that send several measures without a measure level
        const expandMeasuresPerLeaf = !hasMeasureLevel && vsArr.length > 1;

        type Leaf = { offset: number; measureIndex: number; path: string[]; labelPath: string[]; subtotalOf: string | null };
        const leaves: Leaf[] = [];
        let leafCounter = 0;
        const walk = (node: DataViewMatrixNode, path: string[], labelPath: string[], subtotalOf: string | null, measureIndex: number, measureOrdinal: number) => {
            let p = path, lp = labelPath, st = subtotalOf, mi = measureIndex;
            const isSubtotal = !!(node as any).isSubtotal;
            const isMeasure = this.isColumnMeasureNode(node);
            // A total node can itself sit on the measure level, so both checks apply independently
            if (isSubtotal && st === null) st = path.join(KEY_SEP); // subtotal of the parent group ("" = grand total)
            if (isMeasure) mi = this.measureIndexOfColumnNode(node, vsArr, measureOrdinal);
            if (!isSubtotal && !isMeasure) {
                p = [...path, this.nodeKey(node)];
                lp = [...labelPath, this.nodeLabel(node, columns)];
            }
            if (!node.children || node.children.length === 0) {
                const offset = leafCounter++;
                if (expandMeasuresPerLeaf) {
                    for (let m = 0; m < vsArr.length; m++) {
                        leaves.push({ offset: offset * vsArr.length + m, measureIndex: m, path: p, labelPath: lp, subtotalOf: st });
                    }
                } else {
                    leaves.push({ offset, measureIndex: mi, path: p, labelPath: lp, subtotalOf: st });
                }
                return;
            }
            this.walkColumnChildren(node.children, (ch, ordinal) => walk(ch, p, lp, st, mi, ordinal));
        };
        if (columns && columns.root && columns.root.children) {
            this.walkColumnChildren(columns.root.children, (ch, ordinal) => walk(ch, [], [], null, 0, ordinal));
        }

        // Every value in a row node records which measure it belongs to (valueSourceIndex, omitted
        // when 0). That is authoritative, so it overrides what the column tree implied.
        if (hasMeasureLevel) {
            const sample = this.sampleRowValues(matrix, leafCounter);
            // Only trust it when the host actually populates it (some non-zero index is present)
            const hostSetsIndex = Object.keys(sample).some(k => !!((sample as any)[k] && (sample as any)[k].valueSourceIndex));
            for (const leaf of hostSetsIndex ? leaves : []) {
                const cell = sample[leaf.offset];
                if (cell) leaf.measureIndex = (cell as any).valueSourceIndex || 0;
            }
        }

        // Offsets of every measure per column group, used for colour lookups
        const groupOffsets = new Map<string, Map<number, number>>();
        const subtotalOffsets = new Map<string, Map<number, number>>();
        for (const leaf of leaves) {
            const target = leaf.subtotalOf !== null ? subtotalOffsets : groupOffsets;
            const gk = leaf.subtotalOf !== null ? leaf.subtotalOf : leaf.path.join(KEY_SEP);
            if (!target.has(gk)) target.set(gk, new Map());
            const map = target.get(gk)!;
            if (!map.has(leaf.measureIndex)) map.set(leaf.measureIndex, leaf.offset);
        }

        const displaySet = new Set(this.displayMeasureIndices);
        // A single displayed measure under column groups does not need its own header row
        const showMeasureRow = groupDepth === 0 || (hasMeasureLevel && this.displayMeasureIndices.length > 1) || expandMeasuresPerLeaf;
        const depth = groupDepth + (showMeasureRow ? 1 : 0);
        const measureName = (mi: number) => {
            const vs = vsArr[mi] as any;
            return String((vs && (vs.displayName || vs.queryName)) || `Measure ${mi + 1}`);
        };
        const withMeasureLevel = (labels: string[], keys: string[], baseKey: string, mi: number) => {
            if (showMeasureRow) {
                labels.push(measureName(mi));
                keys.push(`${baseKey}${KEY_SEP}#m${mi}`);
            }
        };

        const result: DisplayCol[] = [];
        const emittedCollapsed = new Set<string>();
        for (const leaf of leaves) {
            let collapsedAt = -1;
            for (let d = 0; d < leaf.path.length; d++) {
                if (this.collapsedColKeys.has(leaf.path.slice(0, d + 1).join(KEY_SEP))) { collapsedAt = d; break; }
            }

            if (collapsedAt >= 0) {
                const gkey = leaf.path.slice(0, collapsedAt + 1).join(KEY_SEP);
                if (emittedCollapsed.has(gkey)) continue;
                emittedCollapsed.add(gkey);
                // A collapsed group shows the host subtotal of that group, one column per displayed measure
                const totals = subtotalOffsets.get(gkey) || this.onlyDescendantOffsets(gkey, leaves, groupOffsets, subtotalOffsets);
                for (const mi of this.displayMeasureIndices) {
                    const labels: string[] = [];
                    const keys: string[] = [];
                    for (let d = 0; d < groupDepth; d++) {
                        if (d <= collapsedAt) {
                            labels.push(leaf.labelPath[d] || "");
                            keys.push(leaf.path.slice(0, d + 1).join(KEY_SEP));
                        } else {
                            labels.push("");
                            keys.push(`${gkey}${KEY_SEP}#collapsed`);
                        }
                    }
                    withMeasureLevel(labels, keys, gkey, mi);
                    result.push({
                        kind: "collapsed",
                        offset: totals.get(mi),
                        measureIndex: mi,
                        labels,
                        keys,
                        uniqueKey: `${gkey}${KEY_SEP}#m${mi}`,
                        colorOffsets: totals,
                        collapsedLevel: collapsedAt
                    });
                }
                continue;
            }

            if (!displaySet.has(leaf.measureIndex)) continue;
            const isGrandTotal = leaf.subtotalOf === "";
            if (leaf.subtotalOf !== null) {
                if (isGrandTotal && !this.showGrandTotalColumns) continue;
                if (!isGrandTotal && !this.columnSubtotalsEnabled) continue;
            }

            const labels: string[] = [];
            const keys: string[] = [];
            let baseKey: string;
            if (leaf.subtotalOf !== null) {
                baseKey = `${leaf.subtotalOf}${KEY_SEP}#total`;
                const totalLabel = isGrandTotal ? "Grand Total" : "Total";
                for (let d = 0; d < groupDepth; d++) {
                    if (d < leaf.path.length) {
                        labels.push(leaf.labelPath[d] || "");
                        keys.push(leaf.path.slice(0, d + 1).join(KEY_SEP));
                    } else {
                        labels.push(d === leaf.path.length ? totalLabel : "");
                        keys.push(baseKey);
                    }
                }
            } else {
                baseKey = leaf.path.join(KEY_SEP);
                for (let d = 0; d < groupDepth; d++) {
                    if (d < leaf.path.length) {
                        labels.push(leaf.labelPath[d] || "");
                        keys.push(leaf.path.slice(0, d + 1).join(KEY_SEP));
                    } else {
                        labels.push("");
                        keys.push(`${baseKey}${KEY_SEP}#pad`);
                    }
                }
            }
            withMeasureLevel(labels, keys, baseKey, leaf.measureIndex);
            result.push({
                kind: "leaf",
                offset: leaf.offset,
                measureIndex: leaf.measureIndex,
                labels,
                keys,
                uniqueKey: `${baseKey}${KEY_SEP}#m${leaf.measureIndex}`,
                colorOffsets: (leaf.subtotalOf !== null ? subtotalOffsets.get(leaf.subtotalOf) : groupOffsets.get(baseKey)) || new Map()
            });
        }
        // Row-level style measures are read per row: the grand-total column first (pure row
        // context), then any other column that has a value.
        const measureOffsets = new Map<number, number[]>();
        const ordered = [...leaves.filter(l => l.subtotalOf === ""), ...leaves.filter(l => l.subtotalOf !== "")];
        for (const leaf of ordered) {
            if (!measureOffsets.has(leaf.measureIndex)) measureOffsets.set(leaf.measureIndex, []);
            measureOffsets.get(leaf.measureIndex)!.push(leaf.offset);
        }
        return { displayCols: result, depth, groupDepth, measureOffsets };
    }

    /**
     * Total offsets for a collapsed column group the host sent no subtotal for. That only happens
     * when the group has a single member, whose own (or deepest available) values are the total.
     */
    private onlyDescendantOffsets(
        gkey: string,
        leaves: Array<{ path: string[]; subtotalOf: string | null }>,
        groupOffsets: Map<string, Map<number, number>>,
        subtotalOffsets: Map<string, Map<number, number>>
    ): Map<number, number> {
        const prefix = gkey + KEY_SEP;
        const under = (k: string) => k === gkey || k.startsWith(prefix);
        // Prefer the shallowest descendant subtotal, then the single leaf group
        const subKeys = Array.from(subtotalOffsets.keys()).filter(k => k !== gkey && under(k))
            .sort((a, b) => a.split(KEY_SEP).length - b.split(KEY_SEP).length);
        if (subKeys.length) {
            const shallow = subKeys.filter(k => k.split(KEY_SEP).length === subKeys[0].split(KEY_SEP).length);
            if (shallow.length === 1) return subtotalOffsets.get(shallow[0])!;
        }
        const leafGroups = new Set(leaves.filter(l => l.subtotalOf === null && under(l.path.join(KEY_SEP))).map(l => l.path.join(KEY_SEP)));
        if (leafGroups.size === 1) return groupOffsets.get(Array.from(leafGroups)[0]) || new Map();
        return new Map();
    }

    private isMeasureLevel(level: powerbi.DataViewHierarchyLevel | undefined): boolean {
        return !!(level && level.sources && level.sources.some(s => !!(s as any).isMeasure));
    }

    private isColumnMeasureNode(node: DataViewMatrixNode): boolean {
        if (node.level === undefined || node.level === null) return false;
        const levels = this.lastMatrix && this.lastMatrix.columns && this.lastMatrix.columns.levels;
        return this.isMeasureLevel(levels ? levels[node.level] : undefined);
    }

    /** Merges the values maps of the first row nodes until every column offset has been seen. */
    private sampleRowValues(matrix: DataViewMatrix, leafCount: number): ValuesMap {
        const sample: ValuesMap = {};
        let seen = 0;
        let visited = 0;
        const visit = (node: DataViewMatrixNode | undefined): boolean => {
            if (!node || seen >= leafCount || visited > 500) return true;
            visited++;
            if (node.values) {
                for (const k of Object.keys(node.values)) {
                    const key = +k;
                    if (!(key in sample)) { sample[key] = (node.values as any)[k]; seen++; }
                }
            }
            for (const ch of node.children || []) if (visit(ch)) return true;
            return false;
        };
        visit(matrix.rows && matrix.rows.root);
        return sample;
    }

    /** Calls cb for each child with its position among the measure-level siblings (-1 for other nodes). */
    private walkColumnChildren(children: DataViewMatrixNode[], cb: (child: DataViewMatrixNode, measureOrdinal: number) => void) {
        let ordinal = 0;
        for (const ch of children) cb(ch, this.isColumnMeasureNode(ch) ? ordinal++ : -1);
    }

    /** Ctrl+Alt+D overlay: the column/row hierarchy shape Power BI sent, for diagnosing layouts. */
    private describeMatrixShape(): string {
        const m = this.lastMatrix;
        if (!m) return "No data";
        const out: string[] = [];
        const src = (s: any) => `${s.displayName}${s.isMeasure ? "*" : ""}${s.roles ? "{" + Object.keys(s.roles).join(",") + "}" : ""}`;
        out.push("valueSources: " + (m.valueSources || []).map((s, i) => `${i}:${src(s)}`).join(" | "));
        out.push("column levels: " + ((m.columns && m.columns.levels) || []).map((l, i) => `${i}:[${l.sources.map(src).join(",")}]`).join(" "));
        out.push("row levels: " + ((m.rows && m.rows.levels) || []).map((l, i) => `${i}:[${l.sources.map(src).join(",")}]`).join(" "));
        let lines = 0;
        let leaf = 0;
        const dump = (n: DataViewMatrixNode, indent: string) => {
            if (lines++ > 120) return;
            const bits = [`lvl=${n.level}`];
            if (n.levelSourceIndex !== undefined) bits.push(`lsi=${n.levelSourceIndex}`);
            if ((n as any).isSubtotal) bits.push("SUBTOTAL");
            const raw = this.nodeRawValue(n);
            if (raw !== undefined) bits.push(`"${raw}"`);
            if (!n.children || !n.children.length) bits.push(`leaf#${leaf++}`);
            out.push(indent + bits.join(" "));
            for (const ch of n.children || []) dump(ch, indent + "  ");
        };
        const vsName = (i: number) => {
            const vs = m.valueSources && m.valueSources[i];
            return vs ? vs.displayName : `#${i}`;
        };
        out.push(`style roles: cellBg=[${this.cellBgColorMeasureIndices.join(",")}] cellFont=[${this.cellFontColorMeasureIndices.join(",")}]`
            + ` rowBg=${this.rowBgMeasureIndex} rowFont=${this.rowFontMeasureIndex} rowBold=${this.rowBoldMeasureIndex} display=[${this.displayMeasureIndices.join(",")}]`);
        if (this.lastDebug) {
            out.push("display columns (offset:measure | colour offsets):");
            for (const c of this.lastDebug.cols.slice(0, 30)) {
                const colours = Array.from(c.colorOffsets.entries()).map(([mi, off]) => `${mi}@${off}`).join(",");
                out.push(`  ${c.kind} ${c.offset}:${c.measureIndex} ${vsName(c.measureIndex)} | ${colours}`);
            }
            out.push("rows (label: offset=value):");
            for (const r of this.lastDebug.rows.slice(0, 25)) {
                const label = (r.labels || []).filter((l: string) => !!l).join(" / ") || "(row)";
                const vm = r.valuesMap || {};
                const cells = Object.keys(vm).slice(0, 30).map(k => `${k}=${vm[k] && vm[k].value}`).join(" ");
                out.push(`  ${label}${r.isTotal ? " [total]" : ""}${r.collapsed ? " [collapsed]" : ""}: ${cells}`);
            }
        }
        out.push("column tree:");
        for (const ch of (m.columns && m.columns.root && m.columns.root.children) || []) dump(ch, "  ");
        const firstRow = m.rows && m.rows.root && m.rows.root.children && m.rows.root.children[0];
        if (firstRow && firstRow.values) {
            out.push("first row values: " + Object.keys(firstRow.values).slice(0, 40).map(k => {
                const c = (firstRow.values as any)[k];
                const fmt = c && c.objects && c.objects.general && c.objects.general.formatString;
                return `${k}=${c && c.value}${c && c.valueSourceIndex !== undefined ? "@" + c.valueSourceIndex : ""}${fmt ? ` [${fmt}]` : ""}`;
            }).join(" "));
        }
        return out.join("\n");
    }

    /**
     * Maps a measure-level column node to its index in matrix.valueSources.
     * Power BI omits levelSourceIndex when it is 0, so a missing value falls back to the
     * node's position among its measure siblings.
     */
    private measureIndexOfColumnNode(node: DataViewMatrixNode, valueSources: powerbi.DataViewMetadataColumn[], measureOrdinal: number = -1): number {
        const levels = this.lastMatrix && this.lastMatrix.columns && this.lastMatrix.columns.levels;
        const level = levels && node.level !== undefined ? levels[node.level] : undefined;
        const sourceCount = level && level.sources ? level.sources.length : 0;
        let lsi = node.levelSourceIndex;
        if (lsi === undefined || lsi === null) lsi = (measureOrdinal >= 0 && measureOrdinal < sourceCount) ? measureOrdinal : 0;
        const src = level && level.sources[lsi];
        if (src) {
            let idx = valueSources.findIndex(vs => vs.queryName === src.queryName);
            if (idx < 0) idx = valueSources.findIndex(vs => vs.index === src.index);
            if (idx >= 0) return idx;
        }
        return lsi;
    }

    private getRowDepth(rows?: DataViewHierarchy): number {
        if (!rows || !rows.root) return 0;
        // An empty levels array means no row fields are bound
        if (rows.levels) return rows.levels.length;
        const depthFrom = (node: DataViewMatrixNode): number => {
            if (!node.children || node.children.length === 0) return 0;
            let max = 0;
            for (const c of node.children) max = Math.max(max, depthFrom(c));
            return 1 + max;
        };
        if (!rows.root.children || rows.root.children.length === 0) return 0;
        return depthFrom(rows.root);
    }

    /** Stable identity of a hierarchy member (raw value, not the formatted label). */
    private nodeKey(node: DataViewMatrixNode): string {
        const raw = this.nodeRawValue(node);
        if (raw === undefined) return "";
        if (raw === null || raw === "") return BLANK_LABEL;
        return String(raw);
    }

    /** Raw member value: undefined for structural nodes (root, measures, totals), null for blank members. */
    private nodeRawValue(node: DataViewMatrixNode): any {
        if (node.levelValues && node.levelValues.length) {
            for (const lv of node.levelValues) {
                if (lv && lv.value != null) return lv.value;
            }
            return null;
        }
        if (node.value !== undefined) return node.value;
        return undefined;
    }

    /** Member label formatted with the field's format string. */
    private nodeLabel(node: DataViewMatrixNode, hierarchy?: DataViewHierarchy): string {
        const raw = this.nodeRawValue(node);
        if (raw === undefined) return "";
        if (raw === null || raw === "") return BLANK_LABEL;
        if (typeof raw === "string") return raw;
        const level = hierarchy && hierarchy.levels && node.level !== undefined ? hierarchy.levels[node.level] : undefined;
        const src = level && level.sources && level.sources[node.levelSourceIndex || 0];
        const fmt = src && src.format;
        try {
            return this.getFormatter(fmt).format(raw);
        } catch {
            return String(raw);
        }
    }

    private childKey(parentKey: string, node: DataViewMatrixNode): string {
        const part = this.nodeKey(node);
        return parentKey ? `${parentKey}${KEY_SEP}${part}` : part;
    }

    /** Children that form the next group level (no subtotal nodes, no measure-level column nodes). */
    private groupChildren(node: DataViewMatrixNode, axis: "rows" | "columns"): DataViewMatrixNode[] {
        return ((node.children || []) as DataViewMatrixNode[]).filter(ch =>
            !(ch as any).isSubtotal && !(axis === "columns" && this.isColumnMeasureNode(ch)));
    }

    /**
     * Visits every group member of an axis with its key and 1-based depth.
     * Return false from the callback to skip that member's descendants.
     */
    private walkGroups(axis: "rows" | "columns", cb: (node: DataViewMatrixNode, key: string, depth: number, children: DataViewMatrixNode[]) => boolean | void) {
        const hierarchy = this.lastMatrix && this.lastMatrix[axis];
        if (!hierarchy || !hierarchy.root) return;
        if (axis === "rows" && this.getRowDepth(hierarchy) === 0) return;
        const visit = (node: DataViewMatrixNode, parentKey: string, depth: number) => {
            const key = this.childKey(parentKey, node);
            const children = this.groupChildren(node, axis);
            if (cb(node, key, depth, children) === false) return;
            for (const ch of children) visit(ch, key, depth + 1);
        };
        for (const ch of this.groupChildren(hierarchy.root, axis)) visit(ch, "", 1);
    }

    private getCollapsibleKeysFor(axis: "rows" | "columns"): Set<string> {
        const keys = new Set<string>();
        this.walkGroups(axis, (_node, key, _depth, children) => { if (children.length > 0) keys.add(key); });
        return keys;
    }

    /** Number of levels currently visible on an axis, given its collapsed groups. */
    private getDisplayedDepth(axis: "rows" | "columns"): number {
        const collapsed = axis === "rows" ? this.collapsedRowKeys : this.collapsedColKeys;
        let max = 0;
        this.walkGroups(axis, (_node, key, depth, children) => {
            max = Math.max(max, depth);
            if (!children.length || collapsed.has(key)) return false;
        });
        return Math.max(1, max);
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

    private showVisualContextMenu(event: MouseEvent): void {
        if (!this.allowInteractions) return;
        const selectionId = this.createEmptySelectionId();
        this.selectionManager.showContextMenu(selectionId, { x: event.clientX, y: event.clientY });
        event.preventDefault();
    }

    private createEmptySelectionId(): ISelectionId {
        return this.host.createSelectionIdBuilder().createSelectionId();
    }

    private selectionIdsMatch(left: ISelectionId, right: ISelectionId): boolean {
        const leftAny = left as any;
        const rightAny = right as any;
        if (typeof leftAny.equals === "function" && leftAny.equals(right)) return true;
        if (typeof leftAny.includes === "function" && leftAny.includes(right, true)) return true;
        if (typeof rightAny.includes === "function" && rightAny.includes(left, true)) return true;
        return left === right;
    }

    private getAccessibleRowLabel(rowInfo: any): string {
        const labels = (rowInfo && (rowInfo.labelsToUse || rowInfo.labels)) || [];
        return labels.filter((label: string) => !!label).join(" / ") || "Matrix value";
    }

    private getObjectValue<T>(objects: powerbi.DataViewObjects | undefined, objectName: string, propertyName: string, defaultValue: T): T {
        const obj = objects && (objects as any)[objectName];
        const v = obj && obj[propertyName];
        return (v !== undefined) ? (v as T) : defaultValue;
    }

    /** Text/enumeration property as a string ("" when unset); tolerates {value} wrappers. */
    private readText(objects: powerbi.DataViewObjects | undefined, objectName: string, propertyName: string): string {
        const v: any = this.getObjectValue<any>(objects, objectName, propertyName, "");
        if (typeof v === "string") return v;
        return v && typeof v.value === "string" ? v.value : "";
    }

    /** Returns the enumeration value (plain or {value}) when it is one of the allowed values, else the default. */
    private pickEnum(input: any, allowed: string[], def: string): string {
        const v = typeof input === "string" ? input : (input && typeof input.value === "string" ? input.value : "");
        const lower = v.toLowerCase();
        return allowed.indexOf(lower) >= 0 ? lower : def;
    }

    /** Applies the per-measure header overrides; properties the user never set keep the column-header style. */
    private applyMeasureHeaderStyle(th: HTMLElement, txt: HTMLElement, measureIndex: number) {
        const s = this.measureHeaderStyles[measureIndex];
        if (!s) return;
        if (s.fontFamily) txt.style.fontFamily = String(s.fontFamily);
        if (typeof s.fontSize === "number" && s.fontSize > 0) txt.style.fontSize = `${s.fontSize}px`;
        if (typeof s.bold === "boolean") txt.style.fontWeight = s.bold ? "bold" : "normal";
        if (typeof s.italic === "boolean") txt.style.fontStyle = s.italic ? "italic" : "normal";
        if (typeof s.underline === "boolean") txt.style.textDecoration = s.underline ? "underline" : "none";
        const fore = this.parseColor(s.fontColor);
        const back = this.parseColor(s.backColor);
        if (fore) th.style.color = fore;
        if (back) th.style.backgroundColor = back;
    }

    private setAlignment(el: HTMLElement, horizontal: string, vertical: string) {
        el.style.textAlign = horizontal;
        el.style.verticalAlign = vertical;
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


    private getCellValue(valuesMap: ValuesMap, col: DisplayCol): any {
        if (col.offset === undefined) return "";
        const cell = valuesMap[col.offset];
        return (cell && cell.value != null) ? cell.value : "";
    }

    /** Value of a colour measure for the same column group as col, or null. */
    private getColorMeasureValue(valuesMap: ValuesMap, col: DisplayCol, colorMeasureIndex: number): string | null {
        if (colorMeasureIndex < 0) return null;
        const offset = col.colorOffsets.get(colorMeasureIndex);
        if (offset === undefined) return null;
        const c = valuesMap[offset];
        if (!c || c.value == null || c.value === "") return null;
        return String(c.value);
    }

    private applyCellConditionalStyle(td: HTMLTableCellElement, valuesMap: ValuesMap, col: DisplayCol) {
        const measureIndex = col.measureIndex;
        const cell = col.offset !== undefined ? valuesMap[col.offset] : undefined;
        const obj = cell && (cell.objects as any);
        const colorsObj = obj && (obj.colors as any);
        const explicitFont = colorsObj && (colorsObj.cellColor as any);
        const explicitBg = colorsObj && (colorsObj.cellBg as any);
        const font = explicitFont && ((explicitFont as any).solid?.color || (explicitFont as any).value);
        const bg = explicitBg && ((explicitBg as any).solid?.color || (explicitBg as any).value);

        // Measure-driven colors (per value)
        const tryMeasureColor = (mi: number): string | null => this.getColorMeasureValue(valuesMap, col, mi);
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

    /** Whole-row style from the Row Background Color / Row Font Color / Row Bold measures. */
    private getRowStyle(valuesMap: ValuesMap, layout: ColumnLayout): RowStyle | null {
        const first = (mi: number): any => {
            if (mi < 0) return null;
            for (const off of layout.measureOffsets.get(mi) || []) {
                const c = valuesMap[off];
                if (c && c.value != null && c.value !== "") return c.value;
            }
            return null;
        };
        const bg = first(this.rowBgMeasureIndex);
        const color = first(this.rowFontMeasureIndex);
        const boldRaw = first(this.rowBoldMeasureIndex);
        const bold = boldRaw === true || (typeof boldRaw === "number" && boldRaw !== 0)
            || (typeof boldRaw === "string" && ["1", "true", "yes", "bold"].indexOf(boldRaw.trim().toLowerCase()) >= 0);
        if (bg == null && color == null && !bold) return null;
        return { bg: bg != null ? String(bg) : "", color: color != null ? String(color) : "", bold };
    }

    /** Row header background taken from the first non-empty Cell Background Color value of the row. */
    private getRowMeasureBg(valuesMap: ValuesMap, displayCols: DisplayCol[]): string | null {
        if (!this.cellBgColorMeasureIndices.length) return null;
        // Legacy behaviour; once the Row formatting wells are used, they alone style the row header
        if (this.rowBgMeasureIndex >= 0 || this.rowFontMeasureIndex >= 0 || this.rowBoldMeasureIndex >= 0) return null;
        const bgIndex = this.cellBgColorMeasureIndices[0];
        for (const col of displayCols) {
            const v = this.getColorMeasureValue(valuesMap, col, bgIndex);
            if (v) return v;
        }
        return null;
    }

    private updateDebugOverlay(_ctx?: { table: HTMLTableElement; headerRows: Array<any>; measureCount: number; columnLeaves: DisplayCol[]; rowDepth: number; colDepth: number }) {
        // Shown by Format pane > Behavior > Show diagnostics, or Ctrl+Alt+D where the host passes keys through
        const show = this.debugEnabled || this.debugFromSetting;
        this.debugEl.style.display = show ? "block" : "none";
        if (show) this.debugEl.textContent = this.describeMatrixShape();
    }

    private updateHighContrastState() {
        const palette = this.host.colorPalette;
        this.isHighContrast = !!(palette && palette.isHighContrast);
        this.highContrastForeground = palette?.foreground?.value || "#000000";
        this.highContrastBackground = palette?.background?.value || "#ffffff";
        this.highContrastSelected = palette?.foregroundSelected?.value || this.highContrastForeground;
        this.highContrastSeparator = palette?.separator?.value || this.highContrastForeground;
        this.root.style.setProperty("--ghm-hc-foreground", this.highContrastForeground);
        this.root.style.setProperty("--ghm-hc-background", this.highContrastBackground);
        this.root.style.setProperty("--ghm-hc-selected", this.highContrastSelected);
        this.root.style.setProperty("--ghm-hc-separator", this.highContrastSeparator);
    }

    private applyHighContrastStyles(scope: HTMLElement) {
        if (!this.isHighContrast) return;
        scope.style.color = this.highContrastForeground;
        scope.style.backgroundColor = this.highContrastBackground;
        const cells = Array.from(scope.querySelectorAll("th, td")) as HTMLElement[];
        for (const cell of cells) {
            cell.style.color = this.highContrastForeground;
            cell.style.backgroundColor = this.highContrastBackground;
            cell.style.borderColor = this.highContrastSeparator;
        }
    }

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

    private applyRowHeaderStyle(th: HTMLElement, level?: number) {
        const style = (level !== undefined && this.rowLevelStyles[level]) ? this.rowLevelStyles[level] : {};
        this.setAlignment(th, this.align.rowHeaderH, this.align.rowHeaderV);
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

    private collapseAll() {
        for (const key of this.getCollapsibleKeysFor("columns")) this.collapsedColKeys.add(key);
        for (const key of this.getCollapsibleKeysFor("rows")) this.collapsedRowKeys.add(key);
    }

    private expandAll() {
        this.collapsedColKeys.clear();
        this.collapsedRowKeys.clear();
    }

    private collectOutlineRowsWithTotals(root: DataViewMatrixNode, rowDepth: number): Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> {
        const rows: Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> = [];
        const grandTotalRows: Array<{ labels: string[]; valuesMap?: { [key:number]: powerbi.DataViewMatrixNodeValue }; isTotal?: boolean; toggleKey?: string; depth?: number; collapsed?: boolean; toggles?: { [key: number]: string }; node?: DataViewMatrixNode }> = [];
        const tryRootTotal = () => {
            if (!this.showGrandTotalRows) return;
            const subtotalChild = root.children && (root.children as any[]).find(ch => (ch as any).isSubtotal);
            if (subtotalChild && subtotalChild.values) {
                const map = subtotalChild.values as any;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                grandTotalRows.push({ labels, valuesMap: map, isTotal: true });
                return;
            }
            if (root.values) {
                const map = root.values as any;
                const labels = new Array(rowDepth).fill("");
                labels[0] = "Grand Total";
                grandTotalRows.push({ labels, valuesMap: map, isTotal: true });
            }
        };
        tryRootTotal();

        const rowsHierarchy = this.lastMatrix ? this.lastMatrix.rows : undefined;
        const traverse = (node: DataViewMatrixNode, depth: number, path: string[], keyPath: string[]) => {
            const label = this.nodeLabel(node, rowsHierarchy);
            const keyPart = this.nodeKey(node);
            const hasChildren = !!(node.children && node.children.length);
            const isRootNode = depth === 0 && !keyPart && keyPath.length === 0;
            // Children are ordered here (never in place) so clearing a sort restores the host order
            const realChildren = this.getSortedChildren(node, isRootNode ? 0 : depth + 1);
            const hasCollapsibleChildren = realChildren.length > 0;
            const map: any = hasChildren ? this.rowGroupValues(node) : null;
            const gKey = [...keyPath, keyPart].join(KEY_SEP);
            const isCollapsed = hasCollapsibleChildren && gKey && this.collapsedRowKeys.has(gKey);

            if (isRootNode) {
                 if (hasChildren) {
                    for (const ch of realChildren) traverse(ch, 0, [], []);
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
                        toggleKey: hasCollapsibleChildren ? gKey : undefined,
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
                            toggleKey: hasCollapsibleChildren ? gKey : undefined,
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
                                toggleKey: hasCollapsibleChildren ? gKey : undefined,
                                depth,
                                collapsed: false
                            });
                        } else {
                            // Header only (no subtotal values if disabled)
                            rows.push({
                                labels,
                                valuesMap: {},
                                isTotal: true,
                                toggleKey: hasCollapsibleChildren ? gKey : undefined,
                                depth,
                                collapsed: false
                            });
                        }
                    }

                    // 2. Children
                    for (const ch of realChildren) {
                        // build next path correctly: fill parent positions
                        const nextPath = new Array(rowDepth).fill("");
                        for (let i = 0; i < path.length; i++) nextPath[i] = path[i];
                        nextPath[depth] = label;
                        traverse(ch, depth + 1, nextPath.slice(0, depth + 1), [...keyPath, keyPart]);
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
                if (hasCollapsibleChildren && gKey) toggles[depth] = gKey;
                rows.push({ labels, valuesMap: map!, isTotal: true, toggleKey: hasCollapsibleChildren ? gKey || undefined : undefined, depth, collapsed: collapsedFlag, toggles });
            };

            if (includeSubtotal && (!subtotalAtBottom || isCollapsed)) {
                pushSubtotalRow(!!isCollapsed);
                if (isCollapsed) return;
            }

            if (hasChildren) {
                let firstChild = true;
                for (const ch of realChildren) {
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
                    traverse(ch, depth + 1, nextPath.slice(0, depth + 1), [...keyPath, keyPart]);

                    // If this was the first child, and we didn't emit a top subtotal, we must attach the toggleKey to the first row generated by this child.
                    if (hasCollapsibleChildren && isFirstChildOfGroup && !isCollapsed && !(includeSubtotal && !subtotalAtBottom)) {
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
        this.collapseOneLevel("rows", this.collapsedRowKeys);
    }

    private expandRowLevel() {
        this.expandOneLevel(this.collapsedRowKeys);
    }

    private collapseColLevel() {
        this.collapseOneLevel("columns", this.collapsedColKeys);
    }

    private expandColLevel() {
        this.expandOneLevel(this.collapsedColKeys);
    }

    /**
     * Collapses every visible, still-expanded group at the deepest depth that has one. The depth is
     * taken from groups that can collapse, not from the deepest visible member.
     */
    private collapseOneLevel(axis: "rows" | "columns", target: Set<string>) {
        let deepest = 0;
        this.walkGroups(axis, (_node, key, depth, children) => {
            if (target.has(key)) return false;
            if (children.length > 0) deepest = Math.max(deepest, depth);
        });
        if (!deepest) return;
        this.walkGroups(axis, (_node, key, depth, children) => {
            if (target.has(key)) return false;
            if (depth === deepest) {
                if (children.length > 0) target.add(key);
                return false;
            }
        });
    }

    /** Expands the shallowest collapsed level. */
    private expandOneLevel(target: Set<string>) {
        if (target.size === 0) return;
        let minDepth = Number.MAX_SAFE_INTEGER;
        for (const key of target) minDepth = Math.min(minDepth, key.split(KEY_SEP).length);
        for (const key of Array.from(target)) {
            if (key.split(KEY_SEP).length === minDepth) target.delete(key);
        }
    }

    private initializeDefaultCollapsed(_matrix: DataViewMatrix, enableCollapse: boolean = true) {
        if (!enableCollapse) return;
        // Collapse all column and row groups by default
        this.collapseAll();
    }

    private pruneInvalidCollapsedState(_matrix: DataViewMatrix): boolean {
        const rowsChanged = this.pruneKeySet(this.collapsedRowKeys, this.getCollapsibleKeysFor("rows"));
        const colsChanged = this.pruneKeySet(this.collapsedColKeys, this.getCollapsibleKeysFor("columns"));
        return rowsChanged || colsChanged;
    }

    private pruneKeySet(target: Set<string>, validKeys: Set<string>): boolean {
        let changed = false;
        for (const key of Array.from(target)) {
            if (!validKeys.has(key)) {
                target.delete(key);
                changed = true;
            }
        }
        return changed;
    }

    /** Row children without subtotal nodes, ordered by the active sort (a copy; the data view is never mutated). */
    /** Values of a row group: the host subtotal, else its own values, else those of an only child. */
    private rowGroupValues(node: DataViewMatrixNode): ValuesMap | null {
        const hasValues = (n: DataViewMatrixNode) => !!(n.values && Object.keys(n.values).length > 0);
        if (!node.children || !node.children.length) return hasValues(node) ? (node.values as ValuesMap) : null;
        const subtotalChild = (node.children as any[]).find(ch => ch.isSubtotal);
        if (subtotalChild && subtotalChild.values) return subtotalChild.values as ValuesMap;
        if (hasValues(node)) return node.values as ValuesMap;
        // A group with one member has the same totals as that member
        const real = this.groupChildren(node, "rows");
        return real.length === 1 ? this.rowGroupValues(real[0]) : null;
    }

    private getSortedChildren(node: DataViewMatrixNode, childDepth: number): DataViewMatrixNode[] {
        const children = this.groupChildren(node, "rows");
        const sort = this.sortState;
        if (!sort || children.length < 2) return children;
        const dir = sort.direction === "ASC" ? 1 : -1;
        if (sort.type === "label") {
            if (sort.level !== -1 && sort.level !== childDepth) return children;
            return children
                .map((n, i) => ({ n, i, v: this.nodeRawValue(n) }))
                .sort((a, b) => this.compareSortValues(a.v, b.v, dir) || a.i - b.i)
                .map(x => x.n);
        }
        const offset = this.activeSortOffset;
        if (offset === undefined) return children;
        const valueOf = (n: DataViewMatrixNode): any => {
            // Groups sort by their subtotal, leaves by their own value
            let map: any = n.values;
            const st = n.children && (n.children as any[]).find(ch => (ch as any).isSubtotal);
            if (st && st.values) map = st.values;
            const cell = map && map[offset];
            return cell ? cell.value : null;
        };
        return children
            .map((n, i) => ({ n, i, v: valueOf(n) }))
            .sort((a, b) => this.compareSortValues(a.v, b.v, dir) || a.i - b.i)
            .map(x => x.n);
    }

    /** Number inside formatted text (currency, separators, %, parentheses for negatives), or null. */
    private parseFormattedNumber(v: any): number | null {
        if (typeof v === "number") return v;
        if (typeof v !== "string") return null;
        let s = v.trim();
        // Text with letters ("Q1 2024", "N/A") is a label, not a number
        if (!s || /\p{L}/u.test(s)) return null;
        let negative = false;
        if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
        s = s.replace(/[^\d.\-+]/g, "");
        if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
        const n = parseFloat(s);
        return Number.isFinite(n) ? (negative ? -n : n) : null;
    }

    /** Typed comparison; blanks always sort last. */
    private compareSortValues(a: any, b: any, dir: number): number {
        const aBlank = a == null || a === "";
        const bBlank = b == null || b === "";
        if (aBlank || bBlank) return aBlank === bBlank ? 0 : (aBlank ? 1 : -1);
        if (a instanceof Date && b instanceof Date) return (a.getTime() - b.getTime()) * dir;
        if (typeof a === "number" && typeof b === "number") return (a - b) * dir;
        if (typeof a === "boolean" && typeof b === "boolean") return (Number(a) - Number(b)) * dir;
        // Pre-formatted text such as "1,021.69", "(1,805.80)", "-22.7%" or "₹4,326" sorts by its number
        const na = this.parseFormattedNumber(a);
        const nb = this.parseFormattedNumber(b);
        if (na !== null && nb !== null) return (na - nb) * dir;
        return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" }) * dir;
    }

    private buildHeaderRows(displayCols: DisplayCol[], layout: ColumnLayout): HeaderRow[] {
        const depth = layout.depth;
        const collapsible = this.getCollapsibleKeysFor("columns");
        // The deepest labelled level of each column is where it can be sorted
        const deepestLabel = displayCols.map(col => {
            for (let d = depth - 1; d >= 0; d--) if (col.labels[d]) return d;
            return -1;
        });
        const rows: HeaderRow[] = [];
        for (let level = 0; level < depth; level++) {
            const cells: HeaderCell[] = [];
            let i = 0;
            while (i < displayCols.length) {
                const col = displayCols[i];
                const label = col.labels[level] || "";
                const key = col.keys[level] || "";
                let j = i + 1;
                while (j < displayCols.length
                    && (displayCols[j].keys[level] || "") === key
                    && (displayCols[j].labels[level] || "") === label) j++;
                const span = j - i;
                const isGroupLevel = level < layout.groupDepth;
                cells.push({
                    label,
                    span,
                    key,
                    togglable: isGroupLevel && !!label && collapsible.has(key),
                    collapsed: isGroupLevel && this.collapsedColKeys.has(key),
                    sortKey: (span === 1 && deepestLabel[i] === level) ? col.uniqueKey : undefined,
                    // Levels below the group levels hold the measure names
                    measureIndex: (level >= layout.groupDepth && span === 1) ? col.measureIndex : undefined
                });
                i = j;
            }
            // Drop header rows without any label (e.g. levels hidden by a collapsed parent)
            if (cells.some(c => !!(c.label && c.label.trim()))) rows.push({ level, cells });
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

    /** Formatters are expensive to build, so one is kept per format string. */
    private getFormatter(format: string | undefined): ReturnType<typeof valueFormatter.create> {
        const cacheKey = format || "";
        let f = this.formatterCache.get(cacheKey);
        if (!f) {
            f = valueFormatter.create({ format: format || undefined });
            this.formatterCache.set(cacheKey, f);
        }
        return f;
    }

    /**
     * Dynamic format string the host attached to this cell's value (measures with a dynamic
     * format string, calculation groups). Requires general.formatString in capabilities.
     */
    private getCellFormat(valuesMap: ValuesMap | undefined, col: DisplayCol | undefined): string | undefined {
        if (!valuesMap || !col || col.offset === undefined) return undefined;
        const cell: any = valuesMap[col.offset];
        const fmt = cell && cell.objects && cell.objects.general && cell.objects.general.formatString;
        return typeof fmt === "string" && fmt ? fmt : undefined;
    }

    /** Formats with the value's own dynamic format when present, else the measure's format string. */
    private formatValueByMeasure(v: any, measureIndex: number, cellFormat?: string): string {
        if (v == null || v === "") return "";
        // Measures that return text (e.g. FORMAT(...)) are already formatted: show them as-is
        if (typeof v === "string") return v;
        const fmt = cellFormat || ((this.measureFormats && this.measureFormats[measureIndex]) ? this.measureFormats[measureIndex] : undefined);
        try {
            return this.getFormatter(fmt).format(v);
        } catch { return String(v); }
    }

    /** Clicking cycles ascending -> descending -> unsorted. colKey is DisplayCol.uniqueKey for value sorts. */
    private attachSortHandler(th: HTMLElement, colKey: string, level: number = 0, type: "value" | "label" = "value") {
        th.style.cursor = "pointer";
        const isSorted = !!this.sortState && this.sortState.type === type && (
            (type === "value" && this.sortState.colKey === colKey) ||
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
                this.sortState = { type, colKey, level, direction: "ASC" };
            }
            this.refresh();
        });
    }

    private getTooltipData(value: any, rowInfo: any | undefined, displayCol: DisplayCol | undefined, measureIndex: number): VisualTooltipDataItem[] {
        const res: VisualTooltipDataItem[] = [];

        const addHierarchyItems = (axis: "rows" | "columns", labels: string[] | undefined) => {
            if (!labels || !labels.length) return;
            const levels = axis === "rows"
                ? this.lastMatrix?.rows?.levels
                : this.lastMatrix?.columns?.levels;

            labels.forEach((label, index) => {
                const text = String(label || "").trim();
                if (!text) return;
                const level = levels && levels[index];
                const source = level && level.sources && level.sources[0];
                if (axis === "columns" && source && (source as any).isMeasure) return;
                const displayName = source && source.displayName
                    ? source.displayName
                    : (axis === "rows" ? `Row level ${index + 1}` : `Column level ${index + 1}`);

                res.push({
                    displayName,
                    value: text
                });
            });
        };

        addHierarchyItems("rows", rowInfo && (rowInfo.labels || rowInfo.labelsToUse));
        addHierarchyItems("columns", this.getTooltipColumnLabels(displayCol));

        if (this.lastMatrix && this.lastMatrix.valueSources && this.lastMatrix.valueSources[measureIndex]) {
            res.push({
                displayName: this.lastMatrix.valueSources[measureIndex].displayName,
                value: this.formatValueByMeasure(value, measureIndex, this.getCellFormat(rowInfo && rowInfo.valuesMap, displayCol))
            });
        } else {
             res.push({
                displayName: "Value",
                value: String(value)
            });
        }

        return res;
    }

    private getTooltipColumnLabels(displayCol: DisplayCol | undefined): string[] | undefined {
        // Collapsed columns already carry blank labels below the collapsed level
        return displayCol ? displayCol.labels : undefined;
    }
}
