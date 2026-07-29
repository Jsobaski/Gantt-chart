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

    name: string = "ganttConfig";
    displayName: string = "Gantt Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.bgColor,
        this.textColor,
        this.barHeight,
        this.barFontSize,
        this.barFontColor
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    public ganttConfig: GanttConfigCard = new GanttConfigCard();
    // Typed broadly because visual.ts appends a dynamically-built "Series Colors"
    // card (a plain formattingSettings.SimpleCard) whose slice count varies with
    // however many bar/milestone fields are currently bound.
    cards: Array<formattingSettings.SimpleCard> = [this.ganttConfig];
}
