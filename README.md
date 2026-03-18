# D3.js Visual for Power BI

> Write custom **D3.js v7** visualizations directly inside Power BI — no external tools, no build step, just open the editor and start coding.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![API](https://img.shields.io/badge/Power%20BI%20API-5.11.0-0078d4)](https://github.com/microsoft/powerbi-visuals-api)
[![D3](https://img.shields.io/badge/D3.js-v7-F9A03C)](https://d3js.org)
[![pbiviz tools](https://img.shields.io/badge/pbiviz--tools-7.0.2-blueviolet)](https://github.com/microsoft/PowerBI-visuals-tools)
[![GitHub Pages](https://img.shields.io/badge/docs-GitHub%20Pages-222)](https://behnamebrahimisbuhb.github.io/pbi-d3js-vis/)

![D3.js Visual screenshot](assets/Screenshot.png)

---

## 📖 Documentation

Full documentation, API reference, and live code examples are on the project's **GitHub Pages** site:

**👉 <https://behnamebrahimisbuhb.github.io/pbi-d3js-vis/>**

---

## ✨ Features

| Feature | Details |
|---|---|
| **In-report code editor** | Full [CodeMirror 5](https://codemirror.net/5/) editor with syntax highlighting, line numbers, and search |
| **D3.js v7 built-in** | D3 v7 bundled and exposed globally — no CDN or version conflicts |
| **Live data binding** | Power BI table data injected via `pbi.dsv()` (D3-compatible callback pattern) |
| **JS validation** | [UglifyJS](https://github.com/mishoo/UglifyJS) parses your script before saving; errors reported with line & column |
| **8 theme-aware colours** | Configurable in the Format pane; auto-overridden in high-contrast mode |
| **Rendering events** | Signals `renderingStarted` / `renderingFinished` / `renderingFailed` to Power BI |
| **Landing page** | Guided welcome screen shown when no data or code is present |
| **High contrast** | Windows high-contrast / forced-colours fully supported |
| **Context menu** | Right-click context menu via Power BI `ISelectionManager` |
| **Keyboard navigation** | Full tab-order and Enter/Space activation on all editor icons |
| **Multi-visual selection** | `supportsMultiVisualSelection` enabled |
| **Tooltips** | Power BI tooltip service exposed via `pbi.tooltipService` |
| **Localisation ready** | All UI strings go through `ILocalizationManager`; add a `.resjson` for your language |

---

## 🚀 Getting started

### 1. Import the visual

Download the latest `.pbiviz` from [**Releases**](https://github.com/BehnamEbrahimiSBUHB/pbi-d3js-vis/releases/latest), then import it in Power BI Desktop or Service:

> **Insert → Get more visuals → Import a visual from a file**

### 2. Connect your data

Drag any fields (dimensions, measures, or both) into the **Dataset** bucket in the Fields pane.  
Up to **30,000 rows** are supported.

### 3. Open the editor

Click the **pencil / Focus mode** icon in the visual's top-right corner to enter advanced edit mode.

### 4. Write and run

Write D3.js v7 code in the editor. Press **Save** to execute it. The `pbi` object is always available:

```js
// Minimal bar chart
const svg = d3.select("#chart");
svg.attr("width", pbi.width).attr("height", pbi.height);

pbi.dsv(function(data) {
    const x = d3.scaleBand()
        .domain(data.map(d => d.category))
        .range([0, pbi.width]).padding(0.2);

    const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => +d.value)])
        .range([pbi.height, 0]);

    svg.selectAll("rect")
        .data(data).join("rect")
        .attr("x",      d => x(d.category))
        .attr("y",      d => y(+d.value))
        .attr("width",  x.bandwidth())
        .attr("height", d => pbi.height - y(+d.value))
        .attr("fill",   pbi.colors[0]);
});
```

---

## 📐 The `pbi` API

Every user script runs with a `pbi` context object pre-defined:

| Property / Method | Type | Description |
|---|---|---|
| `pbi.width` | `number` | Usable canvas width in pixels (after margins) |
| `pbi.height` | `number` | Usable canvas height in pixels (after margins) |
| `pbi.colors` | `string[]` | 8 hex colours from the Format pane (auto-overridden in high-contrast) |
| `pbi.isHighContrast` | `boolean` | `true` when high-contrast mode is active |
| `pbi.dsv(accessor?, cb)` | `void` | Power BI data via D3-style callback; optional accessor transforms rows |
| `pbi.selectionManager` | `ISelectionManager` | Cross-filtering and context menus |
| `pbi.tooltipService` | `ITooltipService` | Native Power BI tooltips |
| `pbi.colorPalette` | `IColorPalette` | Access current report theme colours |

> **Global:** `d3` (v7) is available everywhere in your script.  
> PBI services are also accessible via `window.__pbiD3Visual`.

---

## 🛠️ Building from source

Requirements: **Node.js ≥ 18**, **npm ≥ 9**

```bash
# Install dependencies
npm install

# Start dev server (live reload in Power BI)
npm run start

# Package .pbiviz for distribution
npm run package

# Lint
npm run lint
```

The packaged file is written to `dist/`.

### Tech stack

| Package | Version | Role |
|---|---|---|
| `powerbi-visuals-tools` | 7.0.2 | Build toolchain (webpack) |
| `powerbi-visuals-api` | 5.11.0 | Type definitions |
| `d3` | 7.x | Visualization library |
| `codemirror` | 5.x | In-visual code editor |
| `uglify-js` | 3.x | JS syntax validation |
| `powerbi-visuals-utils-formattingmodel` | 6.x | Format pane integration |
| `typescript` | 5.5 | Language |
| `eslint` + `eslint-plugin-powerbi-visuals` | 9.x | Linting |

---

## 📁 Repository structure

```
pbi-d3js-vis/
├── src/
│   ├── visual.ts          # Main visual class
│   ├── settings.ts        # FormattingSettingsModel (format pane)
│   └── messagebox.ts      # Reusable message/dialog component
├── style/
│   └── visual.less        # All styles (editor + landing page)
├── strings/
│   └── en-US.resjson      # English localisation strings
├── assets/
│   ├── icon.png
│   └── Screenshot.png
├── capabilities.json       # Power BI data roles & feature flags
├── pbiviz.json             # Visual metadata & build config
├── tsconfig.json
├── eslint.config.mjs
└── dist/                   # Packaged .pbiviz output (git-ignored)
```

---

## 🌿 Branches

| Branch | Purpose |
|---|---|
| `main` | Source code |
| `gh-pages` | Documentation site (served at the GitHub Pages URL above) |

---

## 📜 Changelog

### v2.0.0 (2025)
- **Modernised** to `powerbi-visuals-tools` v7, API 5.11.0, D3 v7
- Replaced `externalJS` array with webpack-bundled npm packages
- Replaced `DataViewObjectsParser` with `FormattingSettingsModel` (new Format Pane API)
- Replaced TSLint with ESLint (`eslint-plugin-powerbi-visuals`)
- Removed jQuery dependency
- Added rendering events, landing page, high-contrast support, context menu, keyboard navigation, multi-visual selection, tooltips, and localisation
- Fixed MessageBox button-accumulation bug
- Fixed implicit `undefined` return in `getSelectedType()`

### v1.2.0 (2018 — original by Jan Pieter Posthuma)
- Initial release with D3 v3, API 1.9.0, and CodeMirror editor

---

## 🤝 Contributing

Issues and pull requests are welcome.  
Please open an [issue](https://github.com/BehnamEbrahimiSBUHB/pbi-d3js-vis/issues) first for significant changes.

---

## 📄 License

[MIT](LICENSE) © 2025 [Behnam Ebrahimi](https://github.com/BehnamEbrahimiSBUHB)  
Originally forked from [liprec/powerbi-d3jsvisual](https://github.com/liprec/powerbi-d3jsvisual) by Jan Pieter Posthuma.
