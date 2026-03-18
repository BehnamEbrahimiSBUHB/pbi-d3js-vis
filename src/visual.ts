/*
 * Power BI D3.js Visual
 * Copyright (c) 2018 Jan Pieter Posthuma / DataScenarios
 * MIT License
 *
 * Modernised to powerbi-visuals-tools v7 / API v5 / D3 v7
 */

/* eslint-disable powerbi-visuals/no-implied-inner-html */  // Visual injects SVG/CSS into the DOM by design
/* eslint-disable powerbi-visuals/no-banned-terms */         // eval() is the core execution mechanism of this visual
"use strict";

// Style (processed by webpack less-loader)
import "../style/visual.less";

// D3 v7 — also exposed globally so user eval'd scripts can reference `d3`
import * as d3 from "d3";

// CodeMirror v5 + required modes/addons
import * as CodeMirror from "codemirror";
import "codemirror/mode/javascript/javascript";
import "codemirror/mode/css/css";
import "codemirror/addon/dialog/dialog";
import "codemirror/addon/search/search";
import "codemirror/addon/search/searchcursor";

// UglifyJS — used for JS syntax validation before saving
import * as UglifyJS from "uglify-js";

// Power BI API
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions  = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions        = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                    = powerbi.extensibility.visual.IVisual;
import IVisualHost                = powerbi.extensibility.visual.IVisualHost;
import IViewport                  = powerbi.IViewport;
import DataView                   = powerbi.DataView;
import PrimitiveValue             = powerbi.PrimitiveValue;
import EditMode                   = powerbi.EditMode;
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
import VisualObjectInstance        = powerbi.VisualObjectInstance;
import VisualObjectInstancesToPersist = powerbi.VisualObjectInstancesToPersist;
import DataViewPropertyValue       = powerbi.DataViewPropertyValue;
import IVisualEventService         = powerbi.extensibility.IVisualEventService;
import ISelectionManager           = powerbi.extensibility.ISelectionManager;
import ILocalizationManager        = powerbi.extensibility.ILocalizationManager;

// Local modules
import { MessageBoxType, MessageBoxOptions, MessageBox } from "./messagebox";
import { VisualFormattingSettingsModel } from "./settings";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface ClassAndSelector {
    className: string;
    selectorName: string;
}

function createClassAndSelector(name: string): ClassAndSelector {
    return { className: name, selectorName: `.${name}` };
}

/** Convert a number to a CSS pixel string, e.g. 12 → "12px" */
function toPx(n: number): string {
    return `${n}px`;
}

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

export enum pbiD3jsVisType {
    Js,
    Css,
    Object
}

export interface D3JSDataObjects {
    dataObjects: D3JSDataObject[];
}

export interface D3JSDataObject {
    columnName: string;
    values: PrimitiveValue[];
}

// UglifyJS typings don't expose our custom error shape — use a standalone type
export interface CompileError {
    col: number;
    line: number;
    pos: number;
    filename: string;
    message: string;
    stack: string;
}

export interface CompileOutput {
    code?: string;
    error?: CompileError;
}

// ---------------------------------------------------------------------------
// Visual class
// ---------------------------------------------------------------------------

export class pbiD3jsVis implements IVisual {

    // ---- Infrastructure ----
    private target: HTMLElement;
    private host: IVisualHost;
    private viewport: IViewport;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private localizationManager: ILocalizationManager;

    // ---- Hidden persistent settings (written/read directly via persistProperties) ----
    private jsCode: string  = "";
    private cssCode: string = "";

    // ---- Rendering state ----
    private data: D3JSDataObjects;
    private D3jswidth: number  = 0;
    private D3jsheight: number = 0;
    private isHighContrast: boolean = false;
    private open: pbiD3jsVisType = pbiD3jsVisType.Js;
    private isSaved: boolean     = true;
    private reload: boolean      = false;
    private initialized: boolean = false;
    private lastError: string    = "";

    // ---- DOM selections ----
    private editContainer: d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private d3Container:   d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private d3jsFrame:     d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private messageBox:    d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private landingPage:   d3.Selection<HTMLDivElement, unknown, null, undefined>;

    // ---- CodeMirror ----
    private editor: CodeMirror.EditorFromTextArea;

    // ---- Message-box option presets ----
    private saveWarning:      MessageBoxOptions;
    private hideMessageBox:   MessageBoxOptions;
    private overwriteWarning: MessageBoxOptions;

    // ---- CSS class / selector map ----
    private static EditContainer  = createClassAndSelector("editContainer");
    private static D3Container    = createClassAndSelector("d3Container");
    private static EditorHeader   = createClassAndSelector("editorHeader");
    private static EditorTextArea = createClassAndSelector("editorTextArea");
    private static MessageBoxEl   = createClassAndSelector("messageBox");
    private static Icon           = createClassAndSelector("icon");
    private static New            = createClassAndSelector("new");
    private static Save           = createClassAndSelector("save");
    private static Reload         = createClassAndSelector("reload");
    private static Js             = createClassAndSelector("js");
    private static Css            = createClassAndSelector("css");
    private static Object         = createClassAndSelector("object");
    private static Space          = createClassAndSelector("space");
    private static Parse          = createClassAndSelector("parse");
    private static Help           = createClassAndSelector("help");
    private static D3jsLogo       = createClassAndSelector("d3jslogo");
    private static D3jsFrame      = createClassAndSelector("d3jsframe");
    private static LandingPage    = createClassAndSelector("landing-page");

    // ---- Persistent property identifiers ----
    private readonly visualProperties = {
        d3jsJs:  <DataViewObjectPropertyIdentifier>{ objectName: "general", propertyName: "js"  },
        d3jsCss: <DataViewObjectPropertyIdentifier>{ objectName: "general", propertyName: "css" },
    };

    private readonly helpUrl = "https://behnamebrahimisbuhb.github.io/pbi-d3js-vis/";

    // ---- Inline SVG icon set (MDL-style) ----
    private readonly IconSet = {
        new:    `<svg viewBox="0 0 16 16"><path d="M14 10.5v2h2v1h-2v2h-1v-2h-2v-1h2v-2h1zM10 11.5v2h-2v-13h3v11h-1zM4 13.5v-9h3v9h-3zM0 13.5v-5h3v5h-3zM12 9.5v-5h3v5h-3z"></path></svg>`,
        save:   `<svg viewBox="0 0 16 16"><path d="M1.992 1h12q0.406 0 0.711 0.289 0.289 0.305 0.289 0.711v13h-12.211l-1.789-1.797v-11.203q-0.008-0.406 0.289-0.703t0.711-0.297zM10.992 14h3v-12h-1v6h-10v-6h-1v10.789l1.203 1.211h0.797v-4h7v4zM11.992 2h-8v5h8v-5zM6.992 14h3v-3h-5v3h1v-2h1v2z"></path></svg>`,
        reload: `<svg viewBox="0 0 16 16"><path d="M16 7.875q0 2.281-1.078 4.133t-2.914 2.922-4.008 1.070-4.016-1.070-2.914-2.914-1.070-4.023 1.109-4.055 3.031-2.938h-2.141v-1h4v4h-1v-2.32q-1.844 0.891-2.922 2.602t-1.078 3.656q0 1.938 0.938 3.547t2.547 2.563 3.508 0.953 3.523-0.961 2.547-2.539q0.938-1.578 0.938-3.5 0-2.344-1.438-4.234t-3.695-2.508l0.266-0.961q1.273 0.344 2.359 1.086t1.867 1.773 1.211 2.273 0.43 2.445z"></path></svg>`,
        js:     `<svg viewBox="0 0 16 16"><path d="M16 4.422q0 1.953-1.328 3.266t-3.172 1.313q-0.336 0-0.727-0.063l-6.297 6.297q-0.766 0.766-1.859 0.766t-1.844-0.773q-0.773-0.758-0.773-1.852t0.766-1.852l6.297-6.297q-0.063-0.398-0.063-0.867 0-1.070 0.617-2.125 0.961-1.609 2.117-1.922t1.844-0.313 1.289 0.219 1.414 0.703l-3.078 3.078 0.797 0.797 3.078-3.078q0.484 0.813 0.703 1.414t0.219 1.289zM11.641 8q0.563 0 1.203-0.273t1.125-0.758q1.031-1.031 1.031-2.469 0-0.578-0.188-1.102l-2.813 2.805-2.203-2.203 2.805-2.813q-0.531-0.188-1.172-0.188-0.633 0-1.273 0.273t-1.125 0.758q-1.031 1.031-1.031 2.469 0 0.422 0.156 1.047l-6.68 6.688q-0.477 0.469-0.477 1.141t0.477 1.148 1.148 0.477 1.141-0.477l6.688-6.68q0.625 0.156 1.188 0.156z"></path></svg>`,
        css:    `<svg viewBox="0 0 16 16"><path d="M4.5 3q1.359 0 2.5 0.758v-2.758h9v9h-3.273l2.891 5h-10.969l2.086-3.617q-1.063 0.617-2.109 0.617t-1.859-0.344-1.445-0.977q-0.633-0.625-0.977-1.445-0.344-0.813-0.344-2.203t1.32-2.711 3.18-1.32zM15 9v-7h-7v2.711q0.984 1.211 1 2.758l1.133-1.969 2.016 3.5h2.852zM8 7.5q0-1.43-1.023-2.469-1.039-1.031-2.477-1.031t-2.469 1.031-1.031 2.477 1.023 2.469 2.477 1.023 2.477-1.031q1.023-1.039 1.023-2.469zM8.258 10.75q-1.055 1.836-1.875 3.25h7.5q-0.82-1.414-1.875-3.25t-1.875-3.25q-0.82 1.414-1.875 3.25z"></path></svg>`,
        help:   `<svg viewBox="0 0 16 16"><path d="M7.492 0q1.867 0 3.18 1.313t1.32 3.055q0.008 1.734-0.758 2.563t-1.242 1.273l-0.164 0.148q-0.875 0.828-1.195 1.266-0.641 0.883-0.641 1.883v1.5h-1v-1.5q0-1.617 0.758-2.43t1.242-1.266l0.117-0.109q0.922-0.859 1.242-1.305 0.641-0.898 0.641-1.75-0.008-0.852-0.273-1.492-0.266-0.648-1.016-1.398t-2.203-0.75-2.484 1.031-1.023 2.469h-1q0.008-1.844 1.32-3.172t3.18-1.328zM7.992 16h-1v-1h1v1z"></path></svg>`,
        space:  `<svg viewBox="0 0 4 16"><path d="M3.429 11.143v1.714q0 0.357-0.25 0.607t-0.607 0.25h-1.714q-0.357 0-0.607-0.25t-0.25-0.607v-1.714q0-0.357 0.25-0.607t0.607-0.25h1.714q0.357 0 0.607 0.25t0.25 0.607zM3.429 6.571v1.714q0 0.357-0.25 0.607t-0.607 0.25h-1.714q-0.357 0-0.607-0.25t-0.25-0.607v-1.714q0-0.357 0.25-0.607t0.607-0.25h1.714q0.357 0 0.607 0.25t0.25 0.607zM3.429 2v1.714q0 0.357-0.25 0.607t-0.607 0.25h-1.714q-0.357 0-0.607-0.25t-0.25-0.607v-1.714q0-0.357 0.25-0.607t0.607-0.25h1.714q0.357 0 0.607 0.25t0.25 0.607z"></path></svg>`,
        object: `<svg viewBox="0 0 16 16"><path d="M12 13.93l2.992-1.5v-4.375l-2.992 1.492v4.383zM8.008 8.055l-0.008 4.383 3 1.492v-4.383zM5 10.43l2.008-0.992v-2.5l0.992-0.492-0.008-1.891-2.992 1.492v4.383zM1.008 4.555l-0.008 4.383 3 1.492v-4.383zM7.375 3.742l-2.875-1.43-2.875 1.43 2.875 1.438zM14.375 7.242l-2.875-1.43-2.875 1.43 2.875 1.438zM11.5 4.688l4.492 2.25v6.109l-4.492 2.25-4.5-2.234v-2.508l-2.5 1.242-4.5-2.234 0.008-6.125 4.492-2.25 4.492 2.25 0.008 2.508z"></path></svg>`,
        parse:  `<svg viewBox="0 0 16 16"><path d="M5 2.922v10.156l7.258-5.078zM4 1l10 7-10 7v-14z"></path></svg>`,
        d3js:   `<svg width="#widthpx" viewBox="0 0 256 243" version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" preserveAspectRatio="xMidYMid meet">
            <defs>
                <linearGradient x1="-82.6367258%" y1="-92.819878%" x2="103.767353%" y2="106.041826%" id="lg1"><stop stop-color="#F9A03C" offset="0%"></stop><stop stop-color="#F7974E" offset="100%"></stop></linearGradient>
                <linearGradient x1="-258.923825%" y1="-248.970263%" x2="97.6202479%" y2="98.7684937%" id="lg2"><stop stop-color="#F9A03C" offset="0%"></stop><stop stop-color="#F7974E" offset="100%"></stop></linearGradient>
                <linearGradient x1="-223.162629%" y1="-261.967947%" x2="94.0283377%" y2="101.690818%" id="lg3"><stop stop-color="#F9A03C" offset="0%"></stop><stop stop-color="#F7974E" offset="100%"></stop></linearGradient>
                <linearGradient x1="11.3387123%" y1="-1.82169774%" x2="82.496193%" y2="92.1067478%" id="lg4"><stop stop-color="#F26D58" offset="0%"></stop><stop stop-color="#F9A03C" offset="100%"></stop></linearGradient>
                <linearGradient x1="15.8436473%" y1="3.85803114%" x2="120.126091%" y2="72.3802579%" id="lg5"><stop stop-color="#B84E51" offset="0%"></stop><stop stop-color="#F68E48" offset="100%"></stop></linearGradient>
                <linearGradient x1="46.9841705%" y1="23.4664325%" x2="51.881003%" y2="147.391179%" id="lg6"><stop stop-color="#F9A03C" offset="0%"></stop><stop stop-color="#F7974E" offset="100%"></stop></linearGradient>
            </defs>
            <g>
                <path d="M255.52,175.618667 C255.634667,174.504 255.717333,173.378667 255.781333,172.248 C255.858667,170.909333 175.218667,94.3973333 175.218667,94.3973333 L173.290667,94.3973333 C173.290667,94.3973333 255.026667,180.613333 255.52,175.618667 Z" fill="url(#lg1)"></path>
                <path d="M83.472,149.077333 C83.3653333,149.312 83.2586667,149.546667 83.1493333,149.781333 C83.0346667,150.026667 82.9173333,150.272 82.8,150.514667 C80.2293333,155.874667 118.786667,193.568 121.888,188.989333 C122.029333,188.786667 122.170667,188.573333 122.312,188.370667 C122.469333,188.130667 122.624,187.901333 122.778667,187.661333 C125.258667,183.896 84.5733333,146.629333 83.472,149.077333 Z" fill="url(#lg2)"></path>
                <path d="M137.957333,202.082667 C137.848,202.322667 137.072,203.634667 136.362667,204.328 C136.242667,204.568 174.002667,242.016 174.002667,242.016 L177.402667,242.016 C177.405333,242.016 141.957333,203.666667 137.957333,202.082667 Z" fill="url(#lg3)"></path>
                <path d="M255.834667,171.568 C254.069333,210.714667 221.682667,242.016 182.114667,242.016 L176.765333,242.016 L137.250667,203.088 C140.501333,198.504 143.522667,193.754667 146.213333,188.802667 L182.114667,188.802667 C193.469333,188.802667 202.709333,179.568 202.709333,168.208 C202.709333,156.853333 193.469333,147.613333 182.114667,147.613333 L160.869333,147.613333 C162.488,139.056 163.373333,130.232 163.373333,121.205333 C163.373333,112.04 162.472,103.090667 160.794667,94.3973333 L173.992,94.3973333 L255.602667,174.810667 C255.698667,173.733333 255.776,172.656 255.834667,171.568 Z M21.4666667,0 L0,0 L0,53.2133333 L21.4666667,53.2133333 C58.96,53.2133333 89.4666667,83.712 89.4666667,121.205333 C89.4666667,131.405333 87.192,141.088 83.1493333,149.781333 L122.312,188.370667 C135.170667,169.130667 142.688,146.032 142.688,121.205333 C142.688,54.3733333 88.3066667,0 21.4666667,0 Z" fill="url(#lg4)"></path>
                <path d="M182.114667,0 L95.1866667,0 C116.418667,12.9626667 134,31.344 145.978667,53.2133333 L182.114667,53.2133333 C193.469333,53.2133333 202.709333,62.448 202.709333,73.808 C202.709333,85.1653333 193.469333,94.4 182.114667,94.4 L173.994667,94.4 L255.605333,174.813333 C255.797333,172.632 255.917333,170.437333 255.917333,168.208 C255.917333,150.269333 249.48,133.813333 238.792,121.005333 C249.48,108.202667 255.917333,91.744 255.917333,73.808 C255.917333,33.112 222.813333,0 182.114667,0 Z" fill="url(#lg5)"></path>
                <path d="M176.765333,242.016 L95.808,242.016 C112.104,231.952 126.192,218.666667 137.250667,203.088 L176.765333,242.016 Z M122.312,188.370667 L83.152,149.781333 C72.3333333,173.032 48.7573333,189.202667 21.4666667,189.202667 L0,189.202667 L0,242.410667 L21.4666667,242.410667 C63.4773333,242.410667 100.557333,220.922667 122.312,188.370667 Z" fill="url(#lg6)"></path>
            </g>
        </svg>`
    };

    // ---- HTML templates ----
    private readonly tmplSVG = `<svg class="chart" id="chart" width="#width" height="#height"></svg>`;
    private readonly tmplCSS = `<style>#style</style>`;

    // ---------------------------------------------------------------------------
    // Constructor
    // ---------------------------------------------------------------------------

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host   = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.events            = options.host.eventService;
        this.selectionManager  = options.host.createSelectionManager();

        try {
            this.localizationManager = options.host.createLocalizationManager();
        } catch (_) {
            // Localization manager may not be available in all hosts
            this.localizationManager = null;
        }

        // Expose d3 and PBI services globally so user-authored eval'd scripts can reference them
        (window as any).d3 = d3;
        (window as any).__pbiD3Visual = {
            selectionManager:  this.selectionManager,
            tooltipService:    this.host.tooltipService,
            colorPalette:      this.host.colorPalette,
            host:              this.host,
        };

        // Required so child elements with position:absolute are contained correctly
        this.target.style.position = "relative";
        this.target.style.overflow = "hidden";
    }

    // ---------------------------------------------------------------------------
    // Initialise DOM (called once, on first update)
    // ---------------------------------------------------------------------------

    private init(options: VisualUpdateOptions): void {
        const editorIcons = [
            { title: "New",        class: pbiD3jsVis.New.className,    icon: this.IconSet.new,    selected: false },
            { title: "Save",       class: pbiD3jsVis.Save.className,   icon: this.IconSet.save,   selected: false },
            { title: "Reload",     class: pbiD3jsVis.Reload.className, icon: this.IconSet.reload, selected: false },
            { title: "",           class: pbiD3jsVis.Space.className,  icon: this.IconSet.space,  selected: false },
            { title: "JavaScript", class: pbiD3jsVis.Js.className,     icon: this.IconSet.js,     selected: true  },
            { title: "Style",      class: pbiD3jsVis.Css.className,    icon: this.IconSet.css,    selected: false },
            { title: "PBI object", class: pbiD3jsVis.Object.className, icon: this.IconSet.object, selected: false },
            { title: "",           class: pbiD3jsVis.Space.className,  icon: this.IconSet.space,  selected: false },
            { title: "Parse",      class: pbiD3jsVis.Parse.className,  icon: this.IconSet.parse,  selected: false },
            { title: "Help",       class: pbiD3jsVis.Help.className,   icon: this.IconSet.help,   selected: false },
        ];

        this.viewport = options.viewport;

        // Edit-mode container
        this.editContainer = d3.select<HTMLElement, unknown>(this.target)
            .append("div")
            .classed(pbiD3jsVis.EditContainer.className, true) as d3.Selection<HTMLDivElement, unknown, null, undefined>;

        const editorHeader = this.editContainer
            .append("div")
            .classed(pbiD3jsVis.EditorHeader.className, true);

        editorHeader.selectAll(pbiD3jsVis.Icon.selectorName)
            .data(editorIcons)
            .enter()
            .append("div")
            .classed(pbiD3jsVis.Icon.className, true)
            .classed("selected", d => d.selected)
            .attr("tooltip", d => d.title)
            .attr("tabindex", d => d.title ? "0" : null)  // keyboard nav: skip spacers
            .attr("role", d => d.title ? "button" : null)
            .attr("aria-label", d => d.title || null)
            .each(function(d) { this.classList.add(d.class); })
            .html(d => d.icon)
            .on("keydown", function(_event: KeyboardEvent) {
                const ev = _event as KeyboardEvent;
                if (ev.key === "Enter" || ev.key === " ") {
                    ev.preventDefault();
                    (this as HTMLElement).click();
                }
            });

        this.messageBox = editorHeader
            .append("div")
            .classed(pbiD3jsVis.MessageBoxEl.className, true)
            .style("display", "none") as d3.Selection<HTMLDivElement, unknown, null, undefined>;

        // Version label (visible in editor mode for debugging)
        editorHeader
            .append("div")
            .style("float", "right")
            .style("font-size", "10px")
            .style("color", "#999")
            .style("line-height", "24px")
            .style("margin-right", "8px")
            .style("user-select", "text")
            .text("v3.1.0.0");

        // Debug copy button — copies last error to clipboard
        this.lastError = "";
        editorHeader
            .append("div")
            .style("float", "right")
            .style("font-size", "10px")
            .style("color", "#c00")
            .style("line-height", "24px")
            .style("margin-right", "6px")
            .style("cursor", "pointer")
            .text("📋 Copy Debug")
            .on("click", () => {
                const info = [
                    "Visual: pbi-d3js-vis v3.1.0.0",
                    "API: 5.11.0",
                    "Last error: " + (this.lastError || "(none)"),
                    "Module parse: " + (() => { try { const r = UglifyJS.minify("const x=1;", {compress:false,mangle:false,module:true} as any) as any; return r.error ? "FAIL: "+r.error.message : "OK"; } catch(e) { return "EXCEPTION: "+e; } })(),
                    "UglifyJS version: " + ((UglifyJS as any).version || "unknown"),
                ].join("\n");
                navigator.clipboard.writeText(info).then(() => {
                    window.alert("Debug info copied to clipboard!");
                }).catch(() => {
                    window.prompt("Copy this:", info);
                });
            });

        // CodeMirror textarea placeholder
        this.editContainer
            .append("textarea")
            .classed(pbiD3jsVis.EditorTextArea.className, true);

        // Render-mode container
        this.d3Container = d3.select<HTMLElement, unknown>(this.target)
            .append("div")
            .classed(pbiD3jsVis.D3Container.className, true) as d3.Selection<HTMLDivElement, unknown, null, undefined>;

        this.d3jsFrame = this.d3Container
            .append("div")
            .classed(pbiD3jsVis.D3jsFrame.className, true) as d3.Selection<HTMLDivElement, unknown, null, undefined>;

        // Context menu (right-click) on the render area
        this.d3Container.on("contextmenu", (event: MouseEvent) => {
            this.selectionManager.showContextMenu({}, { x: event.clientX, y: event.clientY });
            event.preventDefault();
        });

        // Landing page — shown when no data or no JS code is present
        this.landingPage = d3.select<HTMLElement, unknown>(this.target)
            .append("div")
            .classed(pbiD3jsVis.LandingPage.className, true)
            .style("display", "none") as d3.Selection<HTMLDivElement, unknown, null, undefined>;

        this.buildLandingPage();

        this.open = pbiD3jsVisType.Js;

        this.hideMessageBox = {
            type: MessageBoxType.None,
            base: this.messageBox
        };
        this.saveWarning = {
            type: MessageBoxType.Warning,
            base: this.messageBox,
            text:   this.loc("Editor_SaveWarning"),
            label1: this.loc("Button_Yes"),
            label2: this.loc("Button_No"),
            label3: this.loc("Button_Cancel")
        };
        this.overwriteWarning = {
            type: MessageBoxType.Warning,
            base: this.messageBox,
            text:   this.loc("Editor_OverwriteWarning"),
            label1: this.loc("Button_Yes"),
            label2: this.loc("Button_No")
        };
    }

    // ---------------------------------------------------------------------------
    // update() — called by Power BI on every data/viewport/format change
    // ---------------------------------------------------------------------------

    public update(options: VisualUpdateOptions): void {
        try {
            if (this.events) { this.events.renderingStarted(options); }

            // One-time DOM initialisation
            if (!this.initialized) {
                this.initialized = true;
                this.init(options);
            }
            this.viewport = options.viewport;

            // ---- Landing page decision (done first, before anything can throw) ----
            const hasData    = (options.dataViews?.length ?? 0) > 0
                && (options.dataViews[0]?.table?.rows?.length ?? 0) > 0;
            const inEditMode = options.editMode === EditMode.Advanced;
            const objects    = options.dataViews?.[0]?.metadata?.objects;
            this.jsCode      = (objects?.general?.js  as string) ?? "";
            this.cssCode     = (objects?.general?.css as string) ?? "";

            const showLanding = !inEditMode && (!hasData || this.jsCode === "");

            this.landingPage
                .style("display", showLanding ? "flex" : "none")
                .style("width",   toPx(this.viewport.width))
                .style("height",  toPx(this.viewport.height));
            this.editContainer.style("display", inEditMode                     ? "inline" : "none");
            this.d3Container.style("display",   !inEditMode && !showLanding    ? "inline" : "none");

            // Early exit — nothing more to do while the landing page is visible
            if (showLanding) {
                if (this.events) { this.events.renderingFinished(options); }
                return;
            }

            // ---- Normal rendering path ----
            this.isHighContrast = this.host.colorPalette?.isHighContrast ?? false;
            (window as any).__pbiD3Visual.isHighContrast = this.isHighContrast;
            (window as any).__pbiD3Visual.colorPalette   = this.host.colorPalette;

            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                options.dataViews?.[0]
            );

            if (inEditMode) {
                this.renderEdit();
            } else {
                this.renderVisual(options);
            }

            if (this.events) { this.events.renderingFinished(options); }
        } catch (e) {
            console.error("D3.js visual update error:", e);
            this.lastError = String(e) + "\n" + ((e as any)?.stack || "");
            if (this.events) { this.events.renderingFailed(options, String(e)); }
            // Show error visually on screen for debugging
            this.target.innerHTML = `<div style="padding:12px;font-family:Consolas,monospace;font-size:12px;color:#c00;white-space:pre-wrap;word-break:break-all;background:#fff8f8;border:2px solid #c00;margin:4px;overflow:auto;max-height:100%"><b>pbi-d3js-vis v3.1.0.0 ERROR</b>\n\n${String(e)}\n\n${(e as any)?.stack || ""}</div>`;
        }
    }

    // ---------------------------------------------------------------------------
    // getFormattingModel — new format-pane API (replaces enumerateObjectInstances)
    // ---------------------------------------------------------------------------

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ---------------------------------------------------------------------------
    // Landing page
    // ---------------------------------------------------------------------------

    private loc(key: string): string {
        try {
            return this.localizationManager?.getDisplayName(key) || key;
        } catch (_) {
            return key;
        }
    }

    private buildLandingPage(): void {
        const logo = this.IconSet.d3js.replace(/#width/g, "72");
        const steps = [
            { title: this.loc("LandingPage_Step1_Title"), desc: this.loc("LandingPage_Step1_Desc") },
            { title: this.loc("LandingPage_Step2_Title"), desc: this.loc("LandingPage_Step2_Desc") },
            { title: this.loc("LandingPage_Step3_Title"), desc: this.loc("LandingPage_Step3_Desc") },
        ];

        const inner = this.landingPage
            .append("div")
            .classed("landing-page-inner", true);

        // D3.js logo
        inner.append("div").classed("landing-logo", true).html(logo);

        // Title + description
        inner.append("h1").classed("landing-title", true).text(this.loc("LandingPage_Title"));
        inner.append("p").classed("landing-desc", true).text(this.loc("LandingPage_Description"));

        // Step-by-step list
        const list = inner.append("ul").classed("landing-steps", true);
        steps.forEach((step, i) => {
            const li = list.append("li").classed("landing-step", true);
            li.append("div").classed("landing-step-num", true).text(String(i + 1));
            const body = li.append("div").classed("landing-step-body", true);
            body.append("strong").text(step.title);
            body.append("span").text(` — ${step.desc}`);
        });

        // Documentation link
        inner.append("a")
            .classed("landing-link", true)
            .attr("href", this.helpUrl)
            .attr("target", "_blank")
            .attr("rel", "noopener noreferrer")
            .text(this.loc("LandingPage_DocLink"));
    }

    // ---------------------------------------------------------------------------
    // Edit mode
    // ---------------------------------------------------------------------------

    private renderEdit(): void {
        const viewport   = this.viewport;
        const textarea   = this.editContainer.select<HTMLTextAreaElement>("textarea");
        this.open        = this.getSelectedType();

        // Destroy previous CodeMirror instance
        d3.selectAll(".CodeMirror").remove();

        const textareaEl = document.querySelector(pbiD3jsVis.EditorTextArea.selectorName) as HTMLTextAreaElement;
        this.editor = CodeMirror.fromTextArea(textareaEl, { lineNumbers: true });

        this.switchContext(textarea, this.open);

        this.editor.setSize(viewport.width, viewport.height - 24);
        this.editor.on("change", () => {
            MessageBox.setMessageBox(this.hideMessageBox);
            if (this.reload) {
                this.isSaved = true;
                this.reload  = false;
            } else {
                this.isSaved = false;
            }
            this.editor.save();
        });

        this.registerEvents(textarea);
    }

    // ---------------------------------------------------------------------------
    // Render (view) mode
    // ---------------------------------------------------------------------------

    private renderVisual(options: VisualUpdateOptions): void {
        const settings = this.formattingSettings;
        this.D3jswidth  = this.viewport.width  - settings.margin.left.value - settings.margin.right.value;
        this.D3jsheight = this.viewport.height - settings.margin.top.value  - settings.margin.bottom.value;

        const logoWidth = Math.min(this.D3jswidth, 100);

        const d3IconData = [{ title: "D3.js logo: (c) Mike Bostock", class: pbiD3jsVis.D3jsLogo.className, icon: this.IconSet.d3js }];

        this.d3Container.selectAll(pbiD3jsVis.D3jsLogo.selectorName).remove();

        const d3logo = this.d3Container
            .selectAll(pbiD3jsVis.D3jsLogo.selectorName)
            .data(d3IconData)
            .enter()
            .append("div")
            .attr("tooltip", d => d.title)
            .each(function(d) { this.classList.add(d.class); })
            .style("top",  toPx((this.D3jsheight - logoWidth / 2) / 2))
            .style("left", toPx((this.D3jswidth  - logoWidth / 2) / 2))
            .html(d => d.icon.replace(/#width/g, logoWidth.toString()));

        if (this.jsCode !== "") {
            d3logo.classed("fading", true);
            this.renderD3js(options, this.D3jsheight, this.D3jswidth);
        }
    }

    private renderD3js(options: VisualUpdateOptions, height: number, width: number): void {
        this.data = this.convert(options.dataViews[0]);

        const settings  = this.formattingSettings;
        const d3jsCode  = this.buildHeader(this.data, height, width) + this.jsCode;
        const cssBlock  = this.tmplCSS.replace("#style", this.cssCode);
        const svgBlock  = this.tmplSVG
            .replace(/#height/g, toPx(height))
            .replace(/#width/g,  toPx(width));

        this.d3jsFrame
            .style("height",       toPx(height))
            .style("width",        toPx(width))
            .style("padding-top",  toPx(settings.margin.top.value))
            .style("padding-left", toPx(settings.margin.left.value))
            .html(cssBlock + svgBlock);

        try {
            eval(d3jsCode);
        } catch (ex) {
            console.error("D3.js visual: error during user script evaluation:", ex);
        }

        // Remove the logo overlay once user code has rendered
        d3.selectAll(pbiD3jsVis.D3jsLogo.selectorName).remove();
    }

    // ---------------------------------------------------------------------------
    // Data conversion
    // ---------------------------------------------------------------------------

    private convert(dataView: DataView): D3JSDataObjects {
        if (!dataView?.table?.columns) {
            return { dataObjects: [] };
        }

        const { columns, rows } = dataView.table;
        const dataObjects = columns.map((col, c) => ({
            columnName: col.displayName.replace(/\s+/g, "").toLowerCase(),
            values:     (rows ?? []).map(row => row[c])
        }));

        return { dataObjects };
    }

    // ---------------------------------------------------------------------------
    // Header generator — creates the `pbi` context object for user scripts
    // ---------------------------------------------------------------------------

    private buildHeader(data: D3JSDataObjects, height: number, width: number): string {
        return this.buildHeaderBase(data, height, width, true);
    }

    private buildHeaderView(data: D3JSDataObjects, height: number, width: number): string {
        return this.buildHeaderBase(data, height, width, false);
    }

    private buildHeaderBase(data: D3JSDataObjects, height: number, width: number, minify: boolean): string {
        const nl  = minify ? "" : "\n";
        const t   = minify ? "" : "\t";
        const tt  = minify ? "" : "\t\t";
        const ttt = minify ? "" : "\t\t\t";

        // Resolve colors — in high-contrast mode substitute system colors
        const colors = this.formattingSettings.colors;
        let colorArray: string[];
        if (this.isHighContrast) {
            const pal = this.host.colorPalette;
            const fg  = (pal as any).foreground?.value  ?? "#ffffff";
            const bg  = (pal as any).background?.value  ?? "#000000";
            const fgS = (pal as any).foregroundSelected?.value ?? fg;
            colorArray = [fg, bg, fgS, fg, bg, fgS, fg, bg];
        } else {
            colorArray = [
                colors.color1.value.value,
                colors.color2.value.value,
                colors.color3.value.value,
                colors.color4.value.value,
                colors.color5.value.value,
                colors.color6.value.value,
                colors.color7.value.value,
                colors.color8.value.value,
            ];
        }

        let code = `var pbi = {${nl}`;
        code += `${t}width:${width},${nl}`;
        code += `${t}height:${height},${nl}`;
        code += `${t}isHighContrast:${this.isHighContrast},${nl}`;
        code += `${t}colors:[${nl}${tt}"${colorArray.join(`",${nl}${tt}"`)}${nl}${t}],${nl}`;
        // Expose PBI services via the global __pbiD3Visual reference
        code += `${t}get selectionManager(){return window.__pbiD3Visual.selectionManager;},${nl}`;
        code += `${t}get tooltipService(){return window.__pbiD3Visual.tooltipService;},${nl}`;
        code += `${t}get colorPalette(){return window.__pbiD3Visual.colorPalette;},${nl}`;
        code += `${t}dsv:function(accessor,callback){${nl}${tt}var data=[`;

        if (data?.dataObjects?.length > 0) {
            for (let r = 0; r < data.dataObjects[0].values.length; r++) {
                code += `${nl}${ttt}{`;
                for (const col of data.dataObjects) {
                    code += `${col.columnName}:'${col.values[r]}',`;
                }
                code += `},`;
            }
        }

        code += `${nl}${tt}];${nl}`;
        code += `${tt}if(arguments.length<2){${nl}${ttt}callback=accessor;accessor=null;${nl}${tt}}else{${nl}${ttt}data=data.map(function(d){return accessor(d);});${nl}${tt}}${nl}`;
        code += `${tt}callback(data);${nl}`;
        code += `${t}}${nl}};${nl}`;
        return code;
    }

    // ---------------------------------------------------------------------------
    // Persistence
    // ---------------------------------------------------------------------------

    private persist(code: string, type: pbiD3jsVisType): void {
        if (type === pbiD3jsVisType.Object) { return; }

        const propId   = type === pbiD3jsVisType.Css ? this.visualProperties.d3jsCss : this.visualProperties.d3jsJs;
        const props: { [key: string]: DataViewPropertyValue } = {};
        props[propId.propertyName] = code;

        const objects: VisualObjectInstancesToPersist = {
            merge: [<VisualObjectInstance>{
                objectName: propId.objectName,
                selector:   null,
                properties: props
            }]
        };
        this.host.persistProperties(objects);
    }

    // ---------------------------------------------------------------------------
    // Editor event registration
    // ---------------------------------------------------------------------------

    private registerEvents(textarea: d3.Selection<HTMLTextAreaElement, unknown, HTMLDivElement, unknown>): void {

        // New
        this.editContainer.select(pbiD3jsVis.New.selectorName).on("click", () => {
            this.overwriteWarning.callback1 = () => {
                this.reload = true;
                this.persist("", this.open);
                this.editor.setValue("");
                this.editor.refresh();
            };
            MessageBox.setMessageBox(this.overwriteWarning);
        });

        // Save
        this.editContainer.select(pbiD3jsVis.Save.selectorName).on("click", () => {
            MessageBox.setMessageBox(this.hideMessageBox);
            if (this.parseCode(this.editor, this.open)) {
                this.isSaved = true;
                this.persist(this.editor.getValue(), this.open);
            }
        });

        // JS tab
        this.editContainer.select(pbiD3jsVis.Js.selectorName).on("click", () => {
            this.switchTab(textarea, pbiD3jsVisType.Js);
        });

        // CSS tab
        this.editContainer.select(pbiD3jsVis.Css.selectorName).on("click", () => {
            this.switchTab(textarea, pbiD3jsVisType.Css);
        });

        // Object (read-only PBI header) tab
        this.editContainer.select(pbiD3jsVis.Object.selectorName).on("click", () => {
            this.switchTab(textarea, pbiD3jsVisType.Object);
        });

        // Parse / validate
        this.editContainer.select(pbiD3jsVis.Parse.selectorName).on("click", () => {
            MessageBox.setMessageBox(this.hideMessageBox);
            this.parseCode(this.editor, this.open);
        });

        // Reload from saved
        this.editContainer.select(pbiD3jsVis.Reload.selectorName).on("click", () => {
            const code = this.open === pbiD3jsVisType.Css ? this.cssCode : this.jsCode;
            this.reload = true;
            textarea.text(code);
            this.editor.setValue(code);
            this.editor.refresh();
        });

        // Help
        this.editContainer.select(pbiD3jsVis.Help.selectorName).on("click", () => {
            this.host.launchUrl(this.helpUrl);
        });
    }

    /** Handle tab switch with optional unsaved-changes prompt */
    private switchTab(
        textarea: d3.Selection<HTMLTextAreaElement, unknown, HTMLDivElement, unknown>,
        targetType: pbiD3jsVisType
    ): void {
        MessageBox.setMessageBox(this.hideMessageBox);
        if (!this.parseCode(this.editor, this.open)) { return; }

        if (this.isSaved) {
            this.switchContext(textarea, targetType);
        } else {
            this.saveWarning.callback1 = () => {
                this.isSaved = true;
                this.persist(this.editor.getValue(), this.open);
                this.switchContext(textarea, targetType);
            };
            this.saveWarning.callback2 = () => {
                this.switchContext(textarea, targetType);
            };
            MessageBox.setMessageBox(this.saveWarning);
        }
    }

    private switchContext(
        textarea: d3.Selection<HTMLTextAreaElement, unknown, HTMLDivElement, unknown>,
        type: pbiD3jsVisType
    ): void {
        let code: string;
        let mode: string;
        let readOnly: boolean | "nocursor";
        let openType = type;

        switch (type) {
            case pbiD3jsVisType.Css:
                code     = this.cssCode;
                mode     = "css";
                readOnly = false;
                break;
            case pbiD3jsVisType.Object:
                code     = this.buildHeaderView(this.data, this.D3jsheight, this.D3jswidth);
                mode     = "javascript";
                readOnly = "nocursor";
                openType = pbiD3jsVisType.Js;  // don't allow persist of read-only view
                break;
            case pbiD3jsVisType.Js:
            default:
                code     = this.jsCode;
                mode     = "javascript";
                readOnly = false;
                break;
        }

        this.open   = openType;
        this.reload = true;
        this.switchIcons(type);

        textarea.text(code);
        this.editor.setValue(code);
        this.editor.setOption("mode",     mode);
        this.editor.setOption("readOnly", readOnly);
        this.editor.refresh();
    }

    // ---------------------------------------------------------------------------
    // JS parse / validation (UglifyJS)
    // ---------------------------------------------------------------------------

    private parseCode(editor: CodeMirror.EditorFromTextArea, type: pbiD3jsVisType): boolean {
        if (type !== pbiD3jsVisType.Js) { return true; }

        const result = UglifyJS.minify(editor.getValue(), { compress: false, mangle: false, module: true } as any) as unknown as CompileOutput;

        // Warn if user references the old d3.select("svg") selector
        const cursor = editor.getDoc().getSearchCursor('d3.select("svg")');
        if (cursor.findNext()) {
            result.error = {
                message:  `Replace 'd3.select("svg")' with 'd3.select("#chart")'`,
                line:     cursor.from().line + 1,
                col:      cursor.from().ch,
                pos:      0,
                filename: "",
                stack:    ""
            };
        }

        if (result.error !== undefined) {
            const { message, line, col } = result.error;
            const selectLen = result.error.message.startsWith("Replace") ? ('d3.select("svg")').length : 1;
            const errText = `Parse error: ${message} at (${line}:${col})`;
            this.lastError = errText + " | options: " + JSON.stringify({compress:false,mangle:false,module:true});
            MessageBox.setMessageBox({
                type: MessageBoxType.Error,
                base: this.messageBox,
                text: errText
            });
            editor.getDoc().setSelection(
                { line: line - 1, ch: col },
                { line: line - 1, ch: col + selectLen }
            );
            editor.focus();
            return false;
        }
        return true;
    }

    // ---------------------------------------------------------------------------
    // Icon highlight helpers
    // ---------------------------------------------------------------------------

    private switchIcons(type: pbiD3jsVisType): void {
        this.editContainer.select(pbiD3jsVis.Js.selectorName).classed("selected",     type === pbiD3jsVisType.Js);
        this.editContainer.select(pbiD3jsVis.Css.selectorName).classed("selected",    type === pbiD3jsVisType.Css);
        this.editContainer.select(pbiD3jsVis.Object.selectorName).classed("selected", type === pbiD3jsVisType.Object);
    }

    private getSelectedType(): pbiD3jsVisType {
        if (this.editContainer.select(pbiD3jsVis.Js.selectorName).classed("selected"))     { return pbiD3jsVisType.Js;     }
        if (this.editContainer.select(pbiD3jsVis.Css.selectorName).classed("selected"))    { return pbiD3jsVisType.Css;    }
        if (this.editContainer.select(pbiD3jsVis.Object.selectorName).classed("selected")) { return pbiD3jsVisType.Object; }
        return pbiD3jsVisType.Js; // default
    }
}
