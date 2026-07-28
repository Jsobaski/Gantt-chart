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
    // Typed broadly because visual.ts appends a dynamically-built "Series Colors"
    // card (a plain formattingSettings.SimpleCard) whose slice count varies with
    // however many bar/milestone fields are currently bound.
    cards: Array<formattingSettings.SimpleCard> = [this.ganttConfig];
}
