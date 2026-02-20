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
var FormattingSettingsCard = formattingSettings.SimpleCard;
var FormattingSettingsModel = formattingSettings.Model;
var ToggleSwitch = formattingSettings.ToggleSwitch;
var TextInput = formattingSettings.TextInput;
var ItemDropdown = formattingSettings.ItemDropdown;
var ColorPicker = formattingSettings.ColorPicker;
/**
 * Data Point Formatting Card
 */
class DataPointCardSettings extends FormattingSettingsCard {
    defaultColor = new formattingSettings.ColorPicker({
        name: "defaultColor",
        displayName: "Default color",
        value: { value: "" }
    });
    showAllDataPoints = new formattingSettings.ToggleSwitch({
        name: "showAllDataPoints",
        displayName: "Show all",
        value: true
    });
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        value: { value: "" }
    });
    fillRule = new formattingSettings.ColorPicker({
        name: "fillRule",
        displayName: "Color saturation",
        value: { value: "" }
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Text Size",
        value: 12
    });
    name = "dataPoint";
    displayName = "Data colors";
    slices = [this.defaultColor, this.showAllDataPoints, this.fill, this.fillRule, this.fontSize];
}
/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    dataPointCard = new DataPointCardSettings();
    behaviorCard = new class extends FormattingSettingsCard {
        stickyHeaders = new ToggleSwitch({ name: "stickyHeaders", displayName: "Sticky headers", value: true });
        name = "state";
        displayName = "Behavior";
        // Note: Compact layout and Row header min width are managed in the toolbar
        // and persisted via capabilities, so they are intentionally omitted here.
        slices = [this.stickyHeaders];
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
        name = "labels";
        displayName = "Labels";
        slices = [
            this.rowHeaderFontSize, this.rowHeaderFontFamily, this.rowHeaderBold,
            this.colHeaderFontSize, this.colHeaderFontFamily, this.colHeaderBold,
            this.dataFontSize, this.dataFontFamily, this.dataBold
        ];
    }();
    // Subtotals pane
    subtotalsCard = new class extends FormattingSettingsCard {
        rowSubtotals = new ToggleSwitch({ name: "rowSubtotals", displayName: "Row subtotals", value: true });
        columnSubtotals = new ToggleSwitch({ name: "columnSubtotals", displayName: "Column subtotals", value: true });
        rowSubtotalsType = new TextInput({ name: "rowSubtotalsType", displayName: "Row subtotal position (Top|Bottom)", value: "Top", placeholder: "Top|Bottom" });
        rowSubtotalsPerLevel = new ToggleSwitch({ name: "rowSubtotalsPerLevel", displayName: "Per row level", value: true });
        columnSubtotalsPerLevel = new ToggleSwitch({ name: "columnSubtotalsPerLevel", displayName: "Per column level", value: true });
        name = "subtotal";
        displayName = "Subtotals";
        slices = [this.rowSubtotals, this.columnSubtotals, this.rowSubtotalsType, this.rowSubtotalsPerLevel, this.columnSubtotalsPerLevel];
    }();
    grandTotalCard = new class extends FormattingSettingsCard {
        show = new ToggleSwitch({ name: "show", displayName: "Show grand total", value: true });
        fallbackToRootValues = new ToggleSwitch({ name: "fallbackToRootValues", displayName: "Fallback to root values", value: false });
        name = "grandTotal";
        displayName = "Grand Total";
        slices = [this.show, this.fallbackToRootValues];
    }();
    // Colors card
    colorsCard = new class extends FormattingSettingsCard {
        rowHeaderColor = new ColorPicker({ name: "rowHeaderColor", displayName: "Row header color", value: { value: "" } });
        rowHeaderBg = new ColorPicker({ name: "rowHeaderBg", displayName: "Row header background", value: { value: "" } });
        colHeaderColor = new ColorPicker({ name: "colHeaderColor", displayName: "Column header color", value: { value: "" } });
        colHeaderBg = new ColorPicker({ name: "colHeaderBg", displayName: "Column header background", value: { value: "" } });
        cellColor = new ColorPicker({ name: "cellColor", displayName: "Cell text color", value: { value: "" } });
        cellBg = new ColorPicker({ name: "cellBg", displayName: "Cell background", value: { value: "" } });
        name = "colors";
        displayName = "Colors";
        slices = [this.rowHeaderColor, this.rowHeaderBg, this.colHeaderColor, this.colHeaderBg, this.cellColor, this.cellBg];
    }();
    cards = [this.dataPointCard, this.behaviorCard, this.labelsCard, this.colorsCard, this.subtotalsCard, this.grandTotalCard];
}
//# sourceMappingURL=settings.js.map