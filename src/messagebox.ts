/*
 * Power BI D3.js Visual
 * Copyright (c) 2018 Jan Pieter Posthuma / DataScenarios
 * MIT License
 */

"use strict";

import * as d3 from "d3";

export enum MessageBoxType {
    None,
    Info,
    Warning,
    Error
}

export interface MessageBoxOptions {
    type: MessageBoxType;
    base: d3.Selection<HTMLDivElement, unknown, null, undefined>;
    text?: string;
    label1?: string;
    label2?: string;
    label3?: string;
    callback1?: () => void;
    callback2?: () => void;
    callback3?: () => void;
}

export class MessageBox {
    public static setMessageBox(options: MessageBoxOptions): void {
        // Clear any previously appended buttons before re-rendering
        options.base.selectAll(".inlineBtn").remove();

        options.base
            .style("display", options.type === MessageBoxType.None ? "none" : "inline-block")
            .classed("info",    options.type === MessageBoxType.Info)
            .classed("warning", options.type === MessageBoxType.Warning)
            .classed("error",   options.type === MessageBoxType.Error)
            .text(options.text ?? "");

        if (options.label1) {
            options.base
                .append("div")
                .classed("inlineBtn", true)
                .text(options.label1)
                .on("click", options.callback1 ?? (() => options.base.style("display", "none")));
        }
        if (options.label2) {
            options.base
                .append("div")
                .classed("inlineBtn", true)
                .text(options.label2)
                .on("click", options.callback2 ?? (() => options.base.style("display", "none")));
        }
        if (options.label3) {
            options.base
                .append("div")
                .classed("inlineBtn", true)
                .text(options.label3)
                .on("click", options.callback3 ?? (() => options.base.style("display", "none")));
        }
    }
}
