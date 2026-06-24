"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;

class GanttConfigCard extends FormattingSettingsCard {
    p6BarColor = new formattingSettings.ColorPicker({
        name: "p6BarColor",
        displayName: "P6 Bar Color",
        value: { value: "#4dabf7" }
    });

    maximoBarColor = new formattingSettings.ColorPicker({
        name: "maximoBarColor",
        displayName: "Maximo Bar Color",
        value: { value: "#26c6da" }
    });

    cappExecColor = new formattingSettings.ColorPicker({
        name: "cappExecColor",
        displayName: "CAPP Execution Milestone Color",
        value: { value: "#ab47bc" }
    });

    cappFundingColor = new formattingSettings.ColorPicker({
        name: "cappFundingColor",
        displayName: "CAPP Funding FY Milestone Color",
        value: { value: "#66bb6a" }
    });

    cappPlanningColor = new formattingSettings.ColorPicker({
        name: "cappPlanningColor",
        displayName: "CAPP Planning Milestone Color",
        value: { value: "#ffa726" }
    });

    bgColor = new formattingSettings.ColorPicker({
        name: "bgColor",
        displayName: "Background Color",
        value: { value: "#0d1b2a" }
    });

    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text Color",
        value: { value: "#d0e4f7" }
    });

    name: string = "ganttConfig";
    displayName: string = "Gantt Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.p6BarColor,
        this.maximoBarColor,
        this.cappExecColor,
        this.cappFundingColor,
        this.cappPlanningColor,
        this.bgColor,
        this.textColor
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    public ganttConfig: GanttConfigCard = new GanttConfigCard();
    cards = [this.ganttConfig];
}
