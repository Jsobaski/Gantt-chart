"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;

class GanttConfigCard extends FormattingSettingsCard {
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
        this.bgColor,
        this.textColor
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    public ganttConfig: GanttConfigCard = new GanttConfigCard();
    cards = [this.ganttConfig];
}
