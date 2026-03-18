/*
 * Power BI D3.js Visual — pbi-d3js-vis
 * MIT License
 * Modernised to powerbi-visuals-tools v7 / API v5.11 / D3 v7
 */

/* eslint-disable powerbi-visuals/no-implied-inner-html */
/* eslint-disable powerbi-visuals/no-banned-terms */
"use strict";

import "../style/visual.less";
import * as d3 from "d3";
import * as CodeMirror from "codemirror";
import "codemirror/mode/javascript/javascript";
import "codemirror/mode/css/css";
import "codemirror/addon/dialog/dialog";
import "codemirror/addon/search/search";
import "codemirror/addon/search/searchcursor";
import * as UglifyJS from "uglify-js";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions   = powerbi.extensibility.visual.VisualConstructorOptions;
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

import { MessageBoxType, MessageBoxOptions, MessageBox } from "./messagebox";
import { VisualFormattingSettingsModel } from "./settings";

// ---------------------------------------------------------------------------
const VERSION = "3.2.0.0";

interface CAS { className: string; selectorName: string; }
function cas(n: string): CAS { return { className: n, selectorName: `.${n}` }; }
function px(n: number): string { return `${n}px`; }

export enum VisualEditorTab { Js, Css, Object }

export interface D3JSDataObjects { dataObjects: D3JSDataObject[]; }
export interface D3JSDataObject { columnName: string; values: PrimitiveValue[]; }
interface CompileOutput { code?: string; error?: { col: number; line: number; message: string; pos?: number; filename?: string; stack?: string }; }

// ---------------------------------------------------------------------------
export class pbiD3jsVis implements IVisual {

    // Infrastructure
    private target: HTMLElement;
    private host: IVisualHost;
    private viewport: IViewport;
    private fmtService: FormattingSettingsService;
    private fmtSettings: VisualFormattingSettingsModel;
    private events: any;           // IVisualEventService — may be undefined
    private selMgr: any;           // ISelectionManager — may be undefined
    private locMgr: any;           // ILocalizationManager — may be undefined
    private constructorOk = false; // did constructor succeed?
    private initialized   = false; // did init() succeed?
    private lastError     = "";

    // State
    private jsCode  = "";
    private cssCode = "";
    private data: D3JSDataObjects = { dataObjects: [] };
    private w = 0;
    private h = 0;
    private isHC = false;
    private tab: VisualEditorTab = VisualEditorTab.Js;
    private saved  = true;
    private reload = false;

    // DOM
    private editBox:    d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private viewBox:    d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private frame:      d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private msgBox:     d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private landing:    d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private editor: CodeMirror.EditorFromTextArea;

    // Presets
    private hideMsg:  MessageBoxOptions;
    private saveMsg:  MessageBoxOptions;
    private overMsg:  MessageBoxOptions;

    // Static selectors
    private static S = {
        edit:   cas("editContainer"), view:   cas("d3Container"), header: cas("editorHeader"),
        ta:     cas("editorTextArea"), mb:    cas("messageBox"),  icon:   cas("icon"),
        newI:   cas("new"),   saveI:  cas("save"),   reload: cas("reload"),
        js:     cas("js"),    css:    cas("css"),     obj:    cas("object"),
        space:  cas("space"), parse:  cas("parse"),   help:   cas("help"),
        logo:   cas("d3jslogo"), frame: cas("d3jsframe"), land: cas("landing-page"),
    };

    private readonly props = {
        js:  { objectName: "general", propertyName: "js"  } as DataViewObjectPropertyIdentifier,
        css: { objectName: "general", propertyName: "css" } as DataViewObjectPropertyIdentifier,
    };

    private readonly helpUrl = "https://behnamebrahimisbuhb.github.io/pbi-d3js-vis/";

    // SVG icons (unchanged from before, inlined)
    private readonly I = {
        new:    `<svg viewBox="0 0 16 16"><path d="M14 10.5v2h2v1h-2v2h-1v-2h-2v-1h2v-2h1zM10 11.5v2h-2v-13h3v11h-1zM4 13.5v-9h3v9h-3zM0 13.5v-5h3v5h-3zM12 9.5v-5h3v5h-3z"/></svg>`,
        save:   `<svg viewBox="0 0 16 16"><path d="M1.992 1h12q.406 0 .711.289.289.305.289.711v13h-12.211l-1.789-1.797v-11.203q-.008-.406.289-.703t.711-.297zM10.992 14h3v-12h-1v6h-10v-6h-1v10.789l1.203 1.211h.797v-4h7v4zM11.992 2h-8v5h8v-5zM6.992 14h3v-3h-5v3h1v-2h1v2z"/></svg>`,
        reload: `<svg viewBox="0 0 16 16"><path d="M16 7.875q0 2.281-1.078 4.133t-2.914 2.922-4.008 1.07-4.016-1.07-2.914-2.914-1.07-4.023 1.109-4.055 3.031-2.938h-2.141v-1h4v4h-1v-2.32q-1.844.891-2.922 2.602t-1.078 3.656q0 1.938.938 3.547t2.547 2.563 3.508.953 3.523-.961 2.547-2.539q.938-1.578.938-3.5 0-2.344-1.438-4.234t-3.695-2.508l.266-.961q1.273.344 2.359 1.086t1.867 1.773 1.211 2.273.43 2.445z"/></svg>`,
        js:     `<svg viewBox="0 0 16 16"><path d="M16 4.422q0 1.953-1.328 3.266t-3.172 1.313q-.336 0-.727-.063l-6.297 6.297q-.766.766-1.859.766t-1.844-.773q-.773-.758-.773-1.852t.766-1.852l6.297-6.297q-.063-.398-.063-.867 0-1.07.617-2.125.961-1.609 2.117-1.922t1.844-.313 1.289.219 1.414.703l-3.078 3.078.797.797 3.078-3.078q.484.813.703 1.414t.219 1.289z"/></svg>`,
        css:    `<svg viewBox="0 0 16 16"><path d="M4.5 3q1.359 0 2.5.758v-2.758h9v9h-3.273l2.891 5h-10.969l2.086-3.617q-1.063.617-2.109.617t-1.859-.344-1.445-.977q-.633-.625-.977-1.445-.344-.813-.344-2.203t1.32-2.711 3.18-1.32z"/></svg>`,
        help:   `<svg viewBox="0 0 16 16"><path d="M7.492 0q1.867 0 3.18 1.313t1.32 3.055q.008 1.734-.758 2.563t-1.242 1.273l-.164.148q-.875.828-1.195 1.266-.641.883-.641 1.883v1.5h-1v-1.5q0-1.617.758-2.43t1.242-1.266l.117-.109q.922-.859 1.242-1.305.641-.898.641-1.75-.008-.852-.273-1.492-.266-.648-1.016-1.398t-2.203-.75-2.484 1.031-1.023 2.469h-1q.008-1.844 1.32-3.172t3.18-1.328zM7.992 16h-1v-1h1v1z"/></svg>`,
        space:  `<svg viewBox="0 0 4 16"><path d="M3.429 11.143v1.714q0 .357-.25.607t-.607.25h-1.714q-.357 0-.607-.25t-.25-.607v-1.714q0-.357.25-.607t.607-.25h1.714q.357 0 .607.25t.25.607z"/></svg>`,
        object: `<svg viewBox="0 0 16 16"><path d="M12 13.93l2.992-1.5v-4.375l-2.992 1.492v4.383zM8.008 8.055l-.008 4.383 3 1.492v-4.383z"/></svg>`,
        parse:  `<svg viewBox="0 0 16 16"><path d="M5 2.922v10.156l7.258-5.078zM4 1l10 7-10 7v-14z"/></svg>`,
        d3js:   `<svg width="#widthpx" viewBox="0 0 256 243" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="lg4" x1="11%" y1="-1%" x2="82%" y2="92%"><stop stop-color="#F26D58" offset="0%"/><stop stop-color="#F9A03C" offset="100%"/></linearGradient><linearGradient id="lg5" x1="15%" y1="3%" x2="120%" y2="72%"><stop stop-color="#B84E51" offset="0%"/><stop stop-color="#F68E48" offset="100%"/></linearGradient><linearGradient id="lg6" x1="46%" y1="23%" x2="51%" y2="147%"><stop stop-color="#F9A03C" offset="0%"/><stop stop-color="#F7974E" offset="100%"/></linearGradient></defs><path d="M255.8,171.6C254.1,210.7 221.7,242 182.1,242L176.8,242L137.3,203.1C140.5,198.5 143.5,193.8 146.2,188.8L182.1,188.8C193.5,188.8 202.7,179.6 202.7,168.2C202.7,156.9 193.5,147.6 182.1,147.6L160.9,147.6C162.5,139.1 163.4,130.2 163.4,121.2C163.4,112 162.5,103.1 160.8,94.4L174,94.4L255.6,174.8C255.7,173.7 255.8,172.7 255.8,171.6ZM21.5,0L0,0L0,53.2L21.5,53.2C59,53.2 89.5,83.7 89.5,121.2C89.5,131.4 87.2,141.1 83.1,149.8L122.3,188.4C135.2,169.1 142.7,146 142.7,121.2C142.7,54.4 88.3,0 21.5,0Z" fill="url(#lg4)"/><path d="M182.1,0L95.2,0C116.4,13 134,31.3 146,53.2L182.1,53.2C193.5,53.2 202.7,62.4 202.7,73.8C202.7,85.2 193.5,94.4 182.1,94.4L174,94.4L255.6,174.8C255.8,172.6 255.9,170.4 255.9,168.2C255.9,150.3 249.5,133.8 238.8,121C249.5,108.2 255.9,91.7 255.9,73.8C255.9,33.1 222.8,0 182.1,0Z" fill="url(#lg5)"/><path d="M176.8,242L95.8,242C112.1,232 126.2,218.7 137.3,203.1L176.8,242ZM122.3,188.4L83.2,149.8C72.3,173 48.8,189.2 21.5,189.2L0,189.2L0,242.4L21.5,242.4C63.5,242.4 100.6,220.9 122.3,188.4Z" fill="url(#lg6)"/></svg>`,
    };

    private readonly tmplSVG = `<svg class="chart" id="chart" width="#width" height="#height"></svg>`;
    private readonly tmplCSS = `<style>#style</style>`;

    // ===== CONSTRUCTOR =====
    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        try {
            this.host = options.host;
            this.fmtService = new FormattingSettingsService();
            this.fmtSettings = new VisualFormattingSettingsModel();
            try { this.events = options.host.eventService; } catch (_) { /* noop */ }
            try { this.selMgr = options.host.createSelectionManager(); } catch (_) { /* noop */ }
            try { this.locMgr = options.host.createLocalizationManager(); } catch (_) { /* noop */ }

            (window as any).d3 = d3;
            (window as any).__pbiD3Visual = { host: this.host };

            this.target.style.position = "relative";
            this.target.style.overflow = "hidden";
            this.constructorOk = true;
        } catch (e) {
            this.showFatalError("CONSTRUCTOR", e);
        }
    }

    // ===== UPDATE =====
    public update(options: VisualUpdateOptions): void {
        if (!this.constructorOk) return;
        try {
            if (this.events) { try { this.events.renderingStarted(options); } catch (_) { /* noop */ } }

            if (!this.initialized) {
                this.initialized = true;
                this.initDOM(options);
            }
            this.viewport = options.viewport;

            const hasRows    = !!(options.dataViews?.[0]?.table?.rows?.length);
            const inEdit     = options.editMode === EditMode.Advanced;
            const objs       = options.dataViews?.[0]?.metadata?.objects;
            this.jsCode      = (objs?.general?.js  as string) ?? "";
            this.cssCode     = (objs?.general?.css as string) ?? "";
            const showLand   = !inEdit && (!hasRows || this.jsCode === "");

            // Visibility
            this.landing.style("display", showLand ? "flex" : "none")
                        .style("width", px(this.viewport.width))
                        .style("height", px(this.viewport.height));
            this.editBox.style("display", inEdit && !showLand ? "inline" : "none");
            this.viewBox.style("display", !inEdit && !showLand ? "inline" : "none");

            if (showLand) {
                if (this.events) { try { this.events.renderingFinished(options); } catch (_) { /* noop */ } }
                return;
            }

            // Populate formatting settings
            try {
                this.fmtSettings = this.fmtService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel, options.dataViews?.[0]);
            } catch (_) { /* keep defaults */ }

            this.isHC = !!(this.host.colorPalette as any)?.isHighContrast;

            if (inEdit) {
                this.renderEdit();
            } else {
                this.renderVisual(options);
            }

            if (this.events) { try { this.events.renderingFinished(options); } catch (_) { /* noop */ } }
        } catch (e) {
            this.lastError = String(e) + "\n" + ((e as any)?.stack || "");
            console.error("pbi-d3js-vis update:", e);
            if (this.events) { try { this.events.renderingFailed(options, String(e)); } catch (_) { /* noop */ } }
            this.showFatalError("UPDATE", e);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        try {
            if (this.fmtSettings) {
                return this.fmtService.buildFormattingModel(this.fmtSettings);
            }
        } catch (_) { /* noop */ }
        return { cards: [] } as any;
    }

    // ===== FATAL ERROR DISPLAY =====
    private showFatalError(phase: string, e: any): void {
        try {
            this.target.innerHTML =
                `<div style="padding:16px;font:13px/1.5 Consolas,monospace;color:#c00;background:#fff8f8;border:2px solid #c00;margin:8px;overflow:auto;max-height:calc(100% - 16px);white-space:pre-wrap;word-break:break-all">` +
                `<b>pbi-d3js-vis v${VERSION} — ${phase} ERROR</b>\n\n` +
                `${String(e)}\n\n${(e as any)?.stack || "(no stack)"}\n\n` +
                `Tip: copy this text and report it.` +
                `</div>`;
        } catch (_) { /* truly fatal, nothing we can do */ }
    }

    // ===== LOCALISATION HELPER =====
    private loc(key: string): string {
        try { return this.locMgr?.getDisplayName(key) || key; } catch (_) { return key; }
    }

    // ===== INIT DOM =====
    private initDOM(options: VisualUpdateOptions): void {
        const S = pbiD3jsVis.S;
        const icons = [
            { t: "New",        c: S.newI.className,  i: this.I.new,    s: false },
            { t: "Save",       c: S.saveI.className, i: this.I.save,   s: false },
            { t: "Reload",     c: S.reload.className, i: this.I.reload, s: false },
            { t: "",           c: S.space.className,  i: this.I.space,  s: false },
            { t: "JavaScript", c: S.js.className,     i: this.I.js,     s: true  },
            { t: "Style",      c: S.css.className,    i: this.I.css,    s: false },
            { t: "PBI object", c: S.obj.className,    i: this.I.object, s: false },
            { t: "",           c: S.space.className,  i: this.I.space,  s: false },
            { t: "Parse",      c: S.parse.className,  i: this.I.parse,  s: false },
            { t: "Help",       c: S.help.className,   i: this.I.help,   s: false },
        ];

        if (options.viewport) this.viewport = options.viewport;

        // Edit container
        this.editBox = d3.select(this.target).append("div")
            .classed(S.edit.className, true) as any;

        const hdr = this.editBox.append("div").classed(S.header.className, true);

        hdr.selectAll(S.icon.selectorName).data(icons).enter()
            .append("div").classed(S.icon.className, true)
            .classed("selected", d => d.s)
            .attr("tooltip", d => d.t).attr("tabindex", d => d.t ? "0" : null)
            .attr("role", d => d.t ? "button" : null)
            .each(function(d) { this.classList.add(d.c); })
            .html(d => d.i)
            .on("keydown", function(_ev: KeyboardEvent) {
                if (_ev.key === "Enter" || _ev.key === " ") { _ev.preventDefault(); (this as HTMLElement).click(); }
            });

        this.msgBox = hdr.append("div").classed(S.mb.className, true)
            .style("display", "none") as any;

        // Version + debug
        hdr.append("div").style("float","right").style("font","10px Consolas").style("color","#999")
            .style("line-height","24px").style("margin-right","8px").text("v" + VERSION);

        hdr.append("div").style("float","right").style("font","10px Consolas").style("color","#c00")
            .style("line-height","24px").style("margin-right","6px").style("cursor","pointer")
            .text("📋Debug").on("click", () => {
                const info = `pbi-d3js-vis v${VERSION}\nError: ${this.lastError || "(none)"}\nUglify module: ${(() => { try { const r = (UglifyJS as any).minify("const x=1;",{compress:false,mangle:false,module:true}); return r.error ? "FAIL:"+r.error.message : "OK"; } catch(e2) { return "EX:"+e2; } })()}`;
                try { navigator.clipboard.writeText(info); } catch (_) { /* noop */ }
                try { window.prompt("Debug info:", info); } catch (_) { /* noop */ }
            });

        this.editBox.append("textarea").classed(S.ta.className, true);

        // View container
        this.viewBox = d3.select(this.target).append("div")
            .classed(S.view.className, true) as any;
        this.frame = this.viewBox.append("div").classed(S.frame.className, true) as any;

        // Context menu
        if (this.selMgr) {
            this.viewBox.on("contextmenu", (ev: MouseEvent) => {
                try { this.selMgr.showContextMenu({}, { x: ev.clientX, y: ev.clientY }); } catch (_) { /* noop */ }
                ev.preventDefault();
            });
        }

        // Landing page
        this.landing = d3.select(this.target).append("div")
            .classed(S.land.className, true).style("display","none") as any;
        this.buildLanding();

        // Message box presets
        this.hideMsg = { type: MessageBoxType.None, base: this.msgBox };
        this.saveMsg = { type: MessageBoxType.Warning, base: this.msgBox,
            text: this.loc("Editor_SaveWarning"),
            label1: this.loc("Button_Yes"), label2: this.loc("Button_No"), label3: this.loc("Button_Cancel") };
        this.overMsg = { type: MessageBoxType.Warning, base: this.msgBox,
            text: this.loc("Editor_OverwriteWarning"),
            label1: this.loc("Button_Yes"), label2: this.loc("Button_No") };

        this.tab = VisualEditorTab.Js;
    }

    // ===== LANDING PAGE =====
    private buildLanding(): void {
        const logo = this.I.d3js.replace(/#width/g, "72");
        const steps = [
            { t: this.loc("LandingPage_Step1_Title"), d: this.loc("LandingPage_Step1_Desc") },
            { t: this.loc("LandingPage_Step2_Title"), d: this.loc("LandingPage_Step2_Desc") },
            { t: this.loc("LandingPage_Step3_Title"), d: this.loc("LandingPage_Step3_Desc") },
        ];
        const inner = this.landing.append("div").classed("landing-page-inner", true);
        inner.append("div").classed("landing-logo", true).html(logo);
        inner.append("h1").classed("landing-title", true).text(this.loc("LandingPage_Title"));
        inner.append("p").classed("landing-desc", true).text(this.loc("LandingPage_Description"));
        const ul = inner.append("ul").classed("landing-steps", true);
        steps.forEach((s, i) => {
            const li = ul.append("li").classed("landing-step", true);
            li.append("div").classed("landing-step-num", true).text(String(i + 1));
            const b = li.append("div").classed("landing-step-body", true);
            b.append("strong").text(s.t);
            b.append("span").text(` — ${s.d}`);
        });
        inner.append("a").classed("landing-link", true)
            .attr("href", this.helpUrl).attr("target", "_blank").attr("rel", "noopener")
            .text(this.loc("LandingPage_DocLink"));
    }

    // ===== EDIT MODE =====
    private renderEdit(): void {
        const S = pbiD3jsVis.S;
        const ta = this.editBox.select<HTMLTextAreaElement>("textarea");
        this.tab = this.getTab();

        d3.selectAll(".CodeMirror").remove();
        const el = this.target.querySelector(S.ta.selectorName) as HTMLTextAreaElement;
        if (!el) { this.showFatalError("EDIT", new Error("textarea not found in DOM")); return; }

        this.editor = CodeMirror.fromTextArea(el, { lineNumbers: true });
        this.switchCtx(ta, this.tab);
        this.editor.setSize(this.viewport.width, this.viewport.height - 24);

        this.editor.on("change", () => {
            MessageBox.setMessageBox(this.hideMsg);
            this.saved = this.reload; this.reload = false;
            this.editor.save();
        });
        this.regEvents(ta);
    }

    // ===== VIEW MODE =====
    private renderVisual(options: VisualUpdateOptions): void {
        const S = pbiD3jsVis.S;
        const m = this.fmtSettings.margin;
        this.w = this.viewport.width  - m.left.value - m.right.value;
        this.h = this.viewport.height - m.top.value  - m.bottom.value;
        const lw = Math.min(this.w, 100);

        this.viewBox.selectAll(S.logo.selectorName).remove();
        const logo = this.viewBox.selectAll(S.logo.selectorName)
            .data([1]).enter().append("div")
            .classed(S.logo.className, true)
            .style("top", px((this.h - lw/2)/2)).style("left", px((this.w - lw/2)/2))
            .html(this.I.d3js.replace(/#width/g, lw.toString()));

        if (this.jsCode) {
            logo.classed("fading", true);
            this.runD3(options);
        }
    }

    private runD3(options: VisualUpdateOptions): void {
        const S = pbiD3jsVis.S;
        this.data = this.convert(options.dataViews?.[0]);
        const code = this.buildHdr(this.data, this.h, this.w) + this.jsCode;
        const css  = this.tmplCSS.replace("#style", this.cssCode);
        const svg  = this.tmplSVG.replace(/#height/g, px(this.h)).replace(/#width/g, px(this.w));
        const m    = this.fmtSettings.margin;

        this.frame.style("height",px(this.h)).style("width",px(this.w))
            .style("padding-top",px(m.top.value)).style("padding-left",px(m.left.value))
            .html(css + svg);

        try { eval(code); } catch (ex) { console.error("D3 eval:", ex); }
        this.viewBox.selectAll(S.logo.selectorName).remove();
    }

    // ===== DATA =====
    private convert(dv: DataView | undefined): D3JSDataObjects {
        if (!dv?.table?.columns) return { dataObjects: [] };
        return {
            dataObjects: dv.table.columns.map((col, c) => ({
                columnName: col.displayName.replace(/\s+/g, "").toLowerCase(),
                values: (dv.table.rows ?? []).map(r => r[c])
            }))
        };
    }

    // ===== HEADER BUILDER =====
    private buildHdr(data: D3JSDataObjects, h: number, w: number): string { return this.hdrBase(data,h,w,true); }
    private buildHdrView(data: D3JSDataObjects, h: number, w: number): string { return this.hdrBase(data,h,w,false); }

    private hdrBase(data: D3JSDataObjects, h: number, w: number, min: boolean): string {
        const n = min?"":"\n", t = min?"":"\t", tt = min?"":"\t\t", ttt = min?"":"\t\t\t";
        const cc = this.fmtSettings.colors;
        const ca = [cc.color1,cc.color2,cc.color3,cc.color4,cc.color5,cc.color6,cc.color7,cc.color8]
            .map(c => c.value.value);

        let s = `var pbi={${n}${t}width:${w},${n}${t}height:${h},${n}${t}isHighContrast:${this.isHC},${n}`;
        s += `${t}colors:[${ca.map(c=>`"${c}"`).join(",")}],${n}`;
        s += `${t}dsv:function(accessor,callback){${n}${tt}var data=[`;
        if (data?.dataObjects?.length) {
            for (let r = 0; r < data.dataObjects[0].values.length; r++) {
                s += `${n}${ttt}{`;
                for (const col of data.dataObjects) s += `${col.columnName}:'${col.values[r]}',`;
                s += `},`;
            }
        }
        s += `${n}${tt}];${n}${tt}if(arguments.length<2){callback=accessor;accessor=null;}else{data=data.map(function(d){return accessor(d);});}${n}${tt}callback(data);${n}${t}}${n}};${n}`;
        return s;
    }

    // ===== PERSIST =====
    private persist(code: string, tab: VisualEditorTab): void {
        if (tab === VisualEditorTab.Object) return;
        const p = tab === VisualEditorTab.Css ? this.props.css : this.props.js;
        const props: Record<string, DataViewPropertyValue> = {};
        props[p.propertyName] = code;
        this.host.persistProperties({ merge: [{ objectName: p.objectName, selector: null, properties: props } as VisualObjectInstance] });
    }

    // ===== EDITOR EVENTS =====
    private regEvents(ta: d3.Selection<HTMLTextAreaElement, unknown, any, any>): void {
        const S = pbiD3jsVis.S;
        this.editBox.select(S.newI.selectorName).on("click", () => {
            this.overMsg.callback1 = () => { this.reload=true; this.persist("",this.tab); this.editor.setValue(""); this.editor.refresh(); };
            MessageBox.setMessageBox(this.overMsg);
        });
        this.editBox.select(S.saveI.selectorName).on("click", () => {
            MessageBox.setMessageBox(this.hideMsg);
            if (this.parseJS(this.editor, this.tab)) { this.saved=true; this.persist(this.editor.getValue(), this.tab); }
        });
        this.editBox.select(S.js.selectorName).on("click",  () => this.swTab(ta, VisualEditorTab.Js));
        this.editBox.select(S.css.selectorName).on("click", () => this.swTab(ta, VisualEditorTab.Css));
        this.editBox.select(S.obj.selectorName).on("click", () => this.swTab(ta, VisualEditorTab.Object));
        this.editBox.select(S.parse.selectorName).on("click", () => { MessageBox.setMessageBox(this.hideMsg); this.parseJS(this.editor, this.tab); });
        this.editBox.select(S.reload.selectorName).on("click", () => {
            const c = this.tab === VisualEditorTab.Css ? this.cssCode : this.jsCode;
            this.reload=true; ta.text(c); this.editor.setValue(c); this.editor.refresh();
        });
        this.editBox.select(S.help.selectorName).on("click", () => { try { this.host.launchUrl(this.helpUrl); } catch(_){/* noop */} });
    }

    private swTab(ta: d3.Selection<HTMLTextAreaElement, unknown, any, any>, to: VisualEditorTab): void {
        MessageBox.setMessageBox(this.hideMsg);
        if (!this.parseJS(this.editor, this.tab)) return;
        if (this.saved) { this.switchCtx(ta, to); return; }
        this.saveMsg.callback1 = () => { this.saved=true; this.persist(this.editor.getValue(),this.tab); this.switchCtx(ta,to); };
        this.saveMsg.callback2 = () => this.switchCtx(ta, to);
        MessageBox.setMessageBox(this.saveMsg);
    }

    private switchCtx(ta: d3.Selection<HTMLTextAreaElement, unknown, any, any>, to: VisualEditorTab): void {
        let code: string, mode: string, ro: boolean|"nocursor", open = to;
        switch (to) {
            case VisualEditorTab.Css:    code=this.cssCode; mode="css"; ro=false; break;
            case VisualEditorTab.Object:
                code = this.data?.dataObjects?.length ? this.buildHdrView(this.data, this.h, this.w) : "// No data loaded yet";
                mode="javascript"; ro="nocursor"; open=VisualEditorTab.Js; break;
            default: code=this.jsCode; mode="javascript"; ro=false; break;
        }
        this.tab=open; this.reload=true; this.switchIcons(to);
        ta.text(code); this.editor.setValue(code); this.editor.setOption("mode",mode);
        this.editor.setOption("readOnly",ro); this.editor.refresh();
    }

    // ===== JS PARSE =====
    private parseJS(ed: CodeMirror.EditorFromTextArea, tab: VisualEditorTab): boolean {
        if (tab !== VisualEditorTab.Js) return true;
        const result = (UglifyJS as any).minify(ed.getValue(), {compress:false,mangle:false,module:true}) as CompileOutput;
        const cur = ed.getDoc().getSearchCursor('d3.select("svg")');
        if (cur.findNext()) {
            result.error = { message: `Replace 'd3.select("svg")' with 'd3.select("#chart")'`, line: cur.from().line+1, col: cur.from().ch };
        }
        if (result.error) {
            const {message,line,col} = result.error;
            const len = message.startsWith("Replace") ? 16 : 1;
            this.lastError = `Parse: ${message} @${line}:${col}`;
            MessageBox.setMessageBox({ type:MessageBoxType.Error, base:this.msgBox, text:`Parse error: ${message} at (${line}:${col})` });
            ed.getDoc().setSelection({line:line-1,ch:col},{line:line-1,ch:col+len}); ed.focus();
            return false;
        }
        return true;
    }

    // ===== ICON HELPERS =====
    private switchIcons(t: VisualEditorTab): void {
        const S = pbiD3jsVis.S;
        this.editBox.select(S.js.selectorName).classed("selected",  t===VisualEditorTab.Js);
        this.editBox.select(S.css.selectorName).classed("selected", t===VisualEditorTab.Css);
        this.editBox.select(S.obj.selectorName).classed("selected", t===VisualEditorTab.Object);
    }
    private getTab(): VisualEditorTab {
        const S = pbiD3jsVis.S;
        if (this.editBox.select(S.css.selectorName).classed("selected")) return VisualEditorTab.Css;
        if (this.editBox.select(S.obj.selectorName).classed("selected")) return VisualEditorTab.Object;
        return VisualEditorTab.Js;
    }
}
