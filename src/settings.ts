/*
 *  Power BI Visualizations
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

import type powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import TextInput = formattingSettings.TextInput;
import ItemDropdown = formattingSettings.ItemDropdown;
import ColorPicker = formattingSettings.ColorPicker;

const verticalAlignItems = [
    { value: "top", displayName: "Top" },
    { value: "middle", displayName: "Middle" },
    { value: "bottom", displayName: "Bottom" }
];

/** Horizontal + vertical alignment for one area; properties are <prefix>HAlign / <prefix>VAlign on the "alignment" object. */
function alignmentGroup(prefix: string, displayName: string, defaultHorizontal: string): formattingSettings.Group {
    return new formattingSettings.Group({
        name: `${prefix}Alignment`,
        displayName,
        slices: [
            new formattingSettings.AlignmentGroup({
                name: `${prefix}HAlign`,
                displayName: "Horizontal",
                // Literal of the const enum AlignmentGroupMode.Horizonal (const enums are not emitted at runtime)
                mode: "horizontalAlignment" as powerbi.visuals.AlignmentGroupMode,
                value: defaultHorizontal
            }),
            new ItemDropdown({
                name: `${prefix}VAlign`,
                displayName: "Vertical",
                items: verticalAlignItems,
                value: verticalAlignItems[1]
            })
        ]
    });
}

export interface MeasureHeaderInfo {
    queryName: string;
    displayName: string;
    /** Raw "measureHeader" object stored on the measure (only properties the user changed). */
    stored: any;
}

export interface MeasureHeaderDefaults {
    fontFamily: string;
    fontSize: number;
    bold: boolean;
}

/** A measure in the Cell Font Color / Cell Background Color wells that a column can take its colour from. */
export interface ColorSourceOption {
    queryName: string;
    displayName: string;
}

/** Stored values of the "Font color from" / "Background from" dropdowns. */
export const COLOR_SOURCE_AUTO = "auto";
export const COLOR_SOURCE_NONE = "none";

function colorSourceItems(options: ColorSourceOption[]) {
    return [
        { value: COLOR_SOURCE_AUTO, displayName: "Automatic (by position)" },
        { value: COLOR_SOURCE_NONE, displayName: "None" },
        ...options.map(o => ({ value: o.queryName, displayName: o.displayName }))
    ];
}

function pickItem(items: Array<{ value: string; displayName: string }>, stored: any) {
    const v = typeof stored === "string" ? stored : (stored && stored.value);
    return items.find(i => i.value === v) || items[0];
}

/** One Format pane group per measure; every slice is bound to that measure by a metadata selector. */
export class MeasureHeaderCard extends formattingSettings.CompositeCard {
    name: string = "measureHeader";
    displayName: string = "Measure columns";
    groups: Array<formattingSettings.Group> = [];

    public build(measures: MeasureHeaderInfo[], defaults: MeasureHeaderDefaults, colorOf: (v: any) => string,
                 fontSources: ColorSourceOption[] = [], backSources: ColorSourceOption[] = []) {
        this.visible = measures.length > 0;
        const fontItems = colorSourceItems(fontSources);
        const backItems = colorSourceItems(backSources);
        this.groups = measures.map((m, i) => {
            const selector = { metadata: m.queryName };
            const s = m.stored || {};
            const color = (v: any) => ({ value: colorOf(v) || "" });
            // Which conditional-colour measure colours this column's cells (Automatic = pair by position)
            const sourceSlices: FormattingSettingsSlice[] = [];
            if (fontSources.length) {
                sourceSlices.push(new ItemDropdown({ name: "fontColorSource", displayName: "Cell font color from", items: fontItems, value: pickItem(fontItems, s.fontColorSource), selector }));
            }
            if (backSources.length) {
                sourceSlices.push(new ItemDropdown({ name: "backColorSource", displayName: "Cell background from", items: backItems, value: pickItem(backItems, s.backColorSource), selector }));
            }
            return new formattingSettings.Group({
                name: `measureHeader${i}`,
                displayName: m.displayName,
                slices: [
                    new formattingSettings.FontControl({
                        name: "font",
                        displayName: "Font",
                        fontFamily: new formattingSettings.FontPicker({ name: "fontFamily", value: s.fontFamily || defaults.fontFamily, selector }),
                        fontSize: new formattingSettings.NumUpDown({ name: "fontSize", value: s.fontSize ?? defaults.fontSize, selector }),
                        bold: new ToggleSwitch({ name: "bold", value: s.bold ?? defaults.bold, selector }),
                        italic: new ToggleSwitch({ name: "italic", value: s.italic ?? false, selector }),
                        underline: new ToggleSwitch({ name: "underline", value: s.underline ?? false, selector })
                    }),
                    new ColorPicker({ name: "fontColor", displayName: "Header text color", value: color(s.fontColor), selector }),
                    new ColorPicker({ name: "backColor", displayName: "Header background", value: color(s.backColor), selector }),
                    new formattingSettings.NumUpDown({ name: "separatorWidth", displayName: "Separator after column: width (px)", value: s.separatorWidth ?? 0, selector }),
                    new ColorPicker({
                        name: "separatorColor",
                        displayName: "Separator after column: color",
                        value: { value: colorOf(s.separatorColor) || "#ffffff" },
                        selector
                    }),
                    ...sourceSlices
                ]
            });
        });
    }
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    behaviorCard = new class extends FormattingSettingsCard {
        showToolbar = new ToggleSwitch({
            name: "showToolbar",
            displayName: "Show toolbar",
            description: "Expand/collapse buttons, layout options and paging above the matrix",
            value: true
        });
        stickyHeaders = new ToggleSwitch({ name: "stickyHeaders", displayName: "Sticky headers", value: true });
        pageSize = new formattingSettings.NumUpDown({ name: "pageSize", displayName: "Page size", value: 100 });
        showDiagnostics = new ToggleSwitch({
            name: "showDiagnostics",
            displayName: "Show diagnostics",
            description: "Shows the data layout Power BI sends to the visual (for troubleshooting)",
            value: false
        });
        name: string = "state";
        displayName: string = "Behavior";
        // Note: Compact layout and Row header min width are managed in the toolbar
        // and persisted via capabilities, so they are intentionally omitted here.
        slices: Array<FormattingSettingsSlice> = [this.showToolbar, this.stickyHeaders, this.pageSize, this.showDiagnostics];
    }();

    labelsCard = new class extends FormattingSettingsCard {
        fontItems = [
            { value: "Segoe UI", displayName: "Segoe UI" },
            { value: "Arial", displayName: "Arial" },
            { value: "Verdana", displayName: "Verdana" },
            { value: "Tahoma", displayName: "Tahoma" },
            { value: "Calibri", displayName: "Calibri" },
            { value: "Times New Roman", displayName: "Times New Roman" }
        ];
        rowHeaderFontSize = new formattingSettings.NumUpDown({ name: "rowHeaderFontSize", displayName: "Row header font size", value: 12 });
        rowHeaderFontFamily = new ItemDropdown({ name: "rowHeaderFontFamily", displayName: "Row header font", value: { value: "Segoe UI", displayName: "Segoe UI" }, items: this.fontItems });
        rowHeaderBold = new ToggleSwitch({ name: "rowHeaderBold", displayName: "Row header bold", value: false });
        colHeaderFontSize = new formattingSettings.NumUpDown({ name: "colHeaderFontSize", displayName: "Column header font size", value: 11 });
        colHeaderFontFamily = new ItemDropdown({ name: "colHeaderFontFamily", displayName: "Column header font", value: { value: "Segoe UI", displayName: "Segoe UI" }, items: this.fontItems });
        colHeaderBold = new ToggleSwitch({ name: "colHeaderBold", displayName: "Column header bold", value: false });
        dataFontSize = new formattingSettings.NumUpDown({ name: "dataFontSize", displayName: "Data font size", value: 11 });
        dataFontFamily = new ItemDropdown({ name: "dataFontFamily", displayName: "Data font", value: { value: "Segoe UI", displayName: "Segoe UI" }, items: this.fontItems });
        dataBold = new ToggleSwitch({ name: "dataBold", displayName: "Data bold", value: false });
        name: string = "labels";
        displayName: string = "Labels";
        slices: Array<FormattingSettingsSlice> = [
            this.rowHeaderFontSize, this.rowHeaderFontFamily, this.rowHeaderBold,
            this.colHeaderFontSize, this.colHeaderFontFamily, this.colHeaderBold,
            this.dataFontSize, this.dataFontFamily, this.dataBold
        ];
    }();

    // Subtotals pane
    subtotalsCard = new class extends FormattingSettingsCard {
        showRowSubtotals = new ToggleSwitch({ name: "showRowSubtotals", displayName: "Row subtotals", value: true });
        showColumnSubtotals = new ToggleSwitch({ name: "showColumnSubtotals", displayName: "Column subtotals", value: true });
        rowSubtotalPositionItems = [
            { value: "Top", displayName: "Top" },
            { value: "Bottom", displayName: "Bottom" }
        ];
        rowSubtotalsType = new ItemDropdown({
            name: "rowSubtotalsType",
            displayName: "Row subtotal position",
            value: { value: "Top", displayName: "Top" },
            items: this.rowSubtotalPositionItems
        });
        name: string = "subtotal";
        displayName: string = "Subtotals";
        slices: Array<FormattingSettingsSlice> = [this.showRowSubtotals, this.showColumnSubtotals, this.rowSubtotalsType];
    }();

    grandTotalCard = new class extends FormattingSettingsCard {
        showRows = new ToggleSwitch({ name: "showRows", displayName: "Grand total row", value: true });
        showColumns = new ToggleSwitch({ name: "showColumns", displayName: "Grand total column", value: true });
        positionItems = [
            { value: "Top", displayName: "Top" },
            { value: "Bottom", displayName: "Bottom" }
        ];
        position = new ItemDropdown({
            name: "position",
            displayName: "Grand total position",
            value: { value: "Bottom", displayName: "Bottom" },
            items: this.positionItems
        });
        name: string = "grandTotal";
        displayName: string = "Grand Total";
        slices: Array<FormattingSettingsSlice> = [this.showRows, this.showColumns, this.position];
    }();

    // Colors card
    colorsCard = new class extends FormattingSettingsCard {
        rowHeaderColor = new ColorPicker({ name: "rowHeaderColor", displayName: "Row header color", value: { value: "" } });
        rowHeaderBg = new ColorPicker({ name: "rowHeaderBg", displayName: "Row header background", value: { value: "" } });
        colHeaderColor = new ColorPicker({ name: "colHeaderColor", displayName: "Column header color", value: { value: "" } });
        colHeaderBg = new ColorPicker({ name: "colHeaderBg", displayName: "Column header background", value: { value: "" } });
        cellColor = new ColorPicker({ name: "cellColor", displayName: "Cell text color", value: { value: "" } });
        cellBg = new ColorPicker({ name: "cellBg", displayName: "Cell background", value: { value: "" } });
        name: string = "colors";
        displayName: string = "Colors";
        slices: Array<FormattingSettingsSlice> = [this.rowHeaderColor, this.rowHeaderBg, this.colHeaderColor, this.colHeaderBg, this.cellColor, this.cellBg];
    }();

    gridCard = new class extends FormattingSettingsCard {
        showHorizontal = new ToggleSwitch({ name: "showHorizontal", displayName: "Horizontal grid", value: true });
        showVertical = new ToggleSwitch({ name: "showVertical", displayName: "Vertical grid", value: true });
        thickness = new formattingSettings.NumUpDown({ name: "thickness", displayName: "Grid thickness (px)", value: 1 });
        color = new ColorPicker({ name: "color", displayName: "Grid color", value: { value: "#d0d0d0" } });
        paddingVertical = new formattingSettings.NumUpDown({ name: "paddingVertical", displayName: "Cell padding: top/bottom (px)", value: 4 });
        paddingHorizontal = new formattingSettings.NumUpDown({ name: "paddingHorizontal", displayName: "Cell padding: left/right (px)", value: 6 });
        headerRowHeight = new formattingSettings.NumUpDown({ name: "headerRowHeight", displayName: "Header row height (px, 0 = auto)", value: 0 });
        rowHeight = new formattingSettings.NumUpDown({ name: "rowHeight", displayName: "Row height (px, 0 = auto)", value: 0 });
        name: string = "grid";
        displayName: string = "Grid";
        slices: Array<FormattingSettingsSlice> = [this.showHorizontal, this.showVertical, this.thickness, this.color,
            this.paddingVertical, this.paddingHorizontal, this.headerRowHeight, this.rowHeight];
    }();

    zebraCard = new class extends FormattingSettingsCard {
        enabled = new ToggleSwitch({ name: "enabled", displayName: "Zebra striping", value: false });
        oddColor = new ColorPicker({ name: "oddColor", displayName: "Odd rows", value: { value: "#f7f7f7" } });
        evenColor = new ColorPicker({ name: "evenColor", displayName: "Even rows", value: { value: "#ffffff" } });
        name: string = "zebra";
        displayName: string = "Zebra striping";
        slices: Array<FormattingSettingsSlice> = [this.enabled, this.oddColor, this.evenColor];
    }();

    zebraColsCard = new class extends FormattingSettingsCard {
        enabled = new ToggleSwitch({ name: "enabled", displayName: "Column striping", value: false });
        oddColor = new ColorPicker({ name: "oddColor", displayName: "Odd columns", value: { value: "#f7f7f7" } });
        evenColor = new ColorPicker({ name: "evenColor", displayName: "Even columns", value: { value: "#ffffff" } });
        name: string = "zebraColumns";
        displayName: string = "Column striping";
        slices: Array<FormattingSettingsSlice> = [this.enabled, this.oddColor, this.evenColor];
    }();

    themeCard = new class extends FormattingSettingsCard {
        presetItems = [
            { value: "Default", displayName: "Default" },
            { value: "Light Blue", displayName: "Light Blue" },
            { value: "Dark", displayName: "Dark" },
            { value: "Excel-like", displayName: "Excel-like" },
            { value: "Modern", displayName: "Modern" }
        ];
        preset = new ItemDropdown({
            name: "preset",
            displayName: "Preset",
            value: { value: "Default", displayName: "Default" },
            items: this.presetItems
        });
        name: string = "theme";
        displayName: string = "Theme";
        slices: Array<FormattingSettingsSlice> = [this.preset];
    }();

    alignmentCard = new class extends formattingSettings.CompositeCard {
        name: string = "alignment";
        displayName: string = "Alignment";
        groups: Array<formattingSettings.Group> = [
            alignmentGroup("rowHeader", "Row headers", "left"),
            alignmentGroup("colHeader", "Column headers", "center"),
            alignmentGroup("values", "Values", "right")
        ];
    }();

    measureHeaderCard = new MeasureHeaderCard();

    rowFormattingCard = new class extends FormattingSettingsCard {
        applyToRowHeaders = new ToggleSwitch({ name: "applyToRowHeaders", displayName: "Apply to row headers", value: true });
        precedenceItems = [
            { value: "cell", displayName: "Cell formatting wins" },
            { value: "row", displayName: "Row formatting wins" }
        ];
        precedence = new ItemDropdown({
            name: "precedence",
            displayName: "When row and cell formatting overlap",
            items: this.precedenceItems,
            value: this.precedenceItems[0]
        });
        name: string = "rowFormatting";
        displayName: string = "Row formatting";
        description: string = "Whole-row styles come from the Row Background Color, Row Font Color and Row Bold field wells";
        slices: Array<FormattingSettingsSlice> = [this.applyToRowHeaders, this.precedence];
    }();

    cards = [this.behaviorCard, this.labelsCard, this.alignmentCard, this.measureHeaderCard, this.rowFormattingCard, this.colorsCard, this.gridCard, this.zebraCard, this.zebraColsCard, this.subtotalsCard, this.grandTotalCard, this.themeCard];
}
