/*
 * Power BI D3.js Visual
 * Copyright (c) 2018 Jan Pieter Posthuma / DataScenarios
 * MIT License
 */

"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

// Power BI default theme palette used as initial color defaults
const DEFAULT_COLORS: string[] = [
    "#01B8AA", "#374649", "#FD625E", "#F2C80F",
    "#5F6B6D", "#8AD4EB", "#FE9666", "#A66999"
];

// ---------------------------------------------------------------------------
// Margin card
// ---------------------------------------------------------------------------
class MarginCard extends formattingSettings.SimpleCard {
    public top = new formattingSettings.NumUpDown({
        name: "top", displayName: "Top", value: 2
    });
    public bottom = new formattingSettings.NumUpDown({
        name: "bottom", displayName: "Bottom", value: 2
    });
    public left = new formattingSettings.NumUpDown({
        name: "left", displayName: "Left", value: 2
    });
    public right = new formattingSettings.NumUpDown({
        name: "right", displayName: "Right", value: 2
    });

    public name: string = "margin";
    public displayName: string = "Margin";
    public slices = [this.top, this.bottom, this.left, this.right];
}

// ---------------------------------------------------------------------------
// Colors card
// ---------------------------------------------------------------------------
class ColorsCard extends formattingSettings.SimpleCard {
    public color1 = new formattingSettings.ColorPicker({
        name: "color1", displayName: "Color 1", value: { value: DEFAULT_COLORS[0] }
    });
    public color2 = new formattingSettings.ColorPicker({
        name: "color2", displayName: "Color 2", value: { value: DEFAULT_COLORS[1] }
    });
    public color3 = new formattingSettings.ColorPicker({
        name: "color3", displayName: "Color 3", value: { value: DEFAULT_COLORS[2] }
    });
    public color4 = new formattingSettings.ColorPicker({
        name: "color4", displayName: "Color 4", value: { value: DEFAULT_COLORS[3] }
    });
    public color5 = new formattingSettings.ColorPicker({
        name: "color5", displayName: "Color 5", value: { value: DEFAULT_COLORS[4] }
    });
    public color6 = new formattingSettings.ColorPicker({
        name: "color6", displayName: "Color 6", value: { value: DEFAULT_COLORS[5] }
    });
    public color7 = new formattingSettings.ColorPicker({
        name: "color7", displayName: "Color 7", value: { value: DEFAULT_COLORS[6] }
    });
    public color8 = new formattingSettings.ColorPicker({
        name: "color8", displayName: "Color 8", value: { value: DEFAULT_COLORS[7] }
    });

    public name: string = "colors";
    public displayName: string = "Colors";
    public slices = [
        this.color1, this.color2, this.color3, this.color4,
        this.color5, this.color6, this.color7, this.color8
    ];
}

// ---------------------------------------------------------------------------
// Root model (exposed to FormattingSettingsService)
// ---------------------------------------------------------------------------
export class VisualFormattingSettingsModel extends formattingSettings.Model {
    public margin = new MarginCard();
    public colors = new ColorsCard();
    public cards = [this.margin, this.colors];
}
