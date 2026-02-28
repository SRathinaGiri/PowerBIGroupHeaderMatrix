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

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import TextInput = formattingSettings.TextInput;
import ItemDropdown = formattingSettings.ItemDropdown;
import ColorPicker = formattingSettings.ColorPicker;

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    behaviorCard = new class extends FormattingSettingsCard {
        stickyHeaders = new ToggleSwitch({ name: "stickyHeaders", displayName: "Sticky headers", value: true });
        pageSize = new formattingSettings.NumUpDown({ name: "pageSize", displayName: "Page size", value: 100 });
        name: string = "state";
        displayName: string = "Behavior";
        // Note: Compact layout and Row header min width are managed in the toolbar
        // and persisted via capabilities, so they are intentionally omitted here.
        slices: Array<FormattingSettingsSlice> = [this.stickyHeaders, this.pageSize];
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
        rowSubtotals = new ToggleSwitch({ name: "rowSubtotals", displayName: "Row subtotals", value: true });
        columnSubtotals = new ToggleSwitch({ name: "columnSubtotals", displayName: "Column subtotals", value: true });
        rowSubtotalPositionItems = [
            { value: "Top", displayName: "Top" },
            { value: "Bottom", displayName: "Bottom" }
        ];
        rowSubtotalsType = new ItemDropdown({
            name: "rowSubtotalsType",
            displayName: "Row subtotal position",
            value: { value: "Bottom", displayName: "Bottom" },
            items: this.rowSubtotalPositionItems
        });
        rowSubtotalsPerLevel = new ToggleSwitch({ name: "rowSubtotalsPerLevel", displayName: "Per row level", value: true });
        columnSubtotalsPerLevel = new ToggleSwitch({ name: "columnSubtotalsPerLevel", displayName: "Per column level", value: true });
        name: string = "subtotal";
        displayName: string = "Subtotals";
        slices: Array<FormattingSettingsSlice> = [this.rowSubtotals, this.columnSubtotals, this.rowSubtotalsType, this.rowSubtotalsPerLevel, this.columnSubtotalsPerLevel];
    }();

    grandTotalCard = new class extends FormattingSettingsCard {
        show = new ToggleSwitch({ name: "show", displayName: "Show grand total", value: true });
        fallbackToRootValues = new ToggleSwitch({ name: "fallbackToRootValues", displayName: "Fallback to root values", value: false });
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
        slices: Array<FormattingSettingsSlice> = [this.show, this.fallbackToRootValues, this.position];
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
        name: string = "grid";
        displayName: string = "Grid";
        slices: Array<FormattingSettingsSlice> = [this.showHorizontal, this.showVertical, this.thickness, this.color];
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

    cards = [this.behaviorCard, this.labelsCard, this.colorsCard, this.gridCard, this.zebraCard, this.zebraColsCard, this.subtotalsCard, this.grandTotalCard];
}
