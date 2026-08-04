"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;

// Defaults are real documented swatches from the ACE Visual Design Guide (no
// invented tints): Dark Blue (#193661, the primary brand dark) for the canvas,
// Light Grey Lighter 80% (#F6F6F6) for text.
class GanttConfigCard extends FormattingSettingsCard {
    bgColor = new formattingSettings.ColorPicker({
        name: "bgColor",
        displayName: "Background Color",
        value: { value: "#193661" }
    });

    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text Color",
        value: { value: "#F6F6F6" }
    });

    barHeight = new formattingSettings.NumUpDown({
        name: "barHeight",
        displayName: "Bar Height (px)",
        value: 11,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 4 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 30 }
        }
    });

    barFontSize = new formattingSettings.NumUpDown({
        name: "barFontSize",
        displayName: "Bar Label Font Size (px)",
        value: 9,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 6 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 20 }
        }
    });

    barFontColor = new formattingSettings.ColorPicker({
        name: "barFontColor",
        displayName: "Bar Label Font Color",
        value: { value: "#000000" }
    });

    locationColor = new formattingSettings.ColorPicker({
        name: "locationColor",
        displayName: "Location Row Color",
        value: { value: "#50A6D3" }
    });

    todayColor = new formattingSettings.ColorPicker({
        name: "todayColor",
        displayName: "Today Line Color",
        value: { value: "#EF5350" }
    });

    showYearLane = new formattingSettings.ToggleSwitch({
        name: "showYearLane",
        displayName: "Show Year Row",
        value: true
    });

    showQuarterLane = new formattingSettings.ToggleSwitch({
        name: "showQuarterLane",
        displayName: "Show Quarter Row",
        value: true
    });

    showMonthLane = new formattingSettings.ToggleSwitch({
        name: "showMonthLane",
        displayName: "Show Month Row",
        value: true
    });

    name: string = "ganttConfig";
    displayName: string = "Gantt Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.bgColor,
        this.textColor,
        this.barHeight,
        this.barFontSize,
        this.barFontColor,
        this.locationColor,
        this.todayColor,
        this.showYearLane,
        this.showQuarterLane,
        this.showMonthLane
    ];
}

class LabelTextCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font",
        value: "Arial, sans-serif"
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size (px)",
        value: 11.5,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 20 }
        }
    });
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Text Color",
        value: { value: "#F6F6F6" }
    });
    name: string = "labelText";
    displayName: string = "Location / Project Names";
    slices: Array<FormattingSettingsSlice> = [this.fontFamily, this.fontSize, this.color];
}

class LegendTextCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font",
        value: "Arial, sans-serif"
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size (px)",
        value: 10.5,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 7 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 18 }
        }
    });
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Text Color",
        value: { value: "#F6F6F6" }
    });
    name: string = "legendText";
    displayName: string = "Legend";
    slices: Array<FormattingSettingsSlice> = [this.fontFamily, this.fontSize, this.color];
}

class FilterTextCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font",
        value: "Arial, sans-serif"
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size (px)",
        value: 11,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 18 }
        }
    });
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Text Color",
        value: { value: "#F6F6F6" }
    });
    name: string = "filterText";
    displayName: string = "Date Filters";
    slices: Array<FormattingSettingsSlice> = [this.fontFamily, this.fontSize, this.color];
}

class HeaderTextCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font",
        value: "Arial, sans-serif"
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size (px)",
        value: 10.5,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 7 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 18 }
        }
    });
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Text Color",
        value: { value: "#F6F6F6" }
    });
    name: string = "headerText";
    displayName: string = "Date Rows (Year / Quarter / Month)";
    slices: Array<FormattingSettingsSlice> = [this.fontFamily, this.fontSize, this.color];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    public ganttConfig: GanttConfigCard = new GanttConfigCard();
    public labelText: LabelTextCard = new LabelTextCard();
    public legendText: LegendTextCard = new LegendTextCard();
    public filterText: FilterTextCard = new FilterTextCard();
    public headerText: HeaderTextCard = new HeaderTextCard();
    // Typed broadly because visual.ts appends a dynamically-built "Series Colors"
    // card (a plain formattingSettings.SimpleCard) whose slice count varies with
    // however many bar/milestone fields are currently bound.
    cards: Array<formattingSettings.SimpleCard> = [
        this.ganttConfig,
        this.labelText,
        this.legendText,
        this.filterText,
        this.headerText
    ];
}
