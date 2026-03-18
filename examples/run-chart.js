// ── Run Chart ─────────────────────────────────────────────────────────────
// Fields required: 1 Date column + 1 Measure in the Dataset bucket
// Copy this entire file into the editor, then press Save to render.
// --------------------------------------------------------------------------

const margin = { top: 20, right: 60, bottom: 60, left: 60 };
const w = pbi.width  - margin.left - margin.right;
const h = pbi.height - margin.top  - margin.bottom;

const svg = d3.select("#chart")
    .attr("width",  pbi.width)
    .attr("height", pbi.height)
    .append("g")
    .attr("transform", `translate(${margin.left},${margin.top})`);

pbi.dsv(function(raw) {

    // ── Auto-detect date column and value column ───────────────────────────
    const keys     = Object.keys(raw[0]);
    const dateKey  = keys.find(k => !isNaN(Date.parse(raw[0][k]))) || keys[0];
    const valueKey = keys.find(k => k !== dateKey) || keys[1];

    const data = raw
        .map(d => ({ date: new Date(d[dateKey]), value: +d[valueKey] }))
        .filter(d => !isNaN(d.date) && !isNaN(d.value))
        .sort((a, b) => a.date - b.date);

    if (data.length === 0) return;

    // ── Scales ─────────────────────────────────────────────────────────────
    const x = d3.scaleTime()
        .domain(d3.extent(data, d => d.date))
        .range([0, w]);

    const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.value)]).nice()
        .range([h, 0]);

    // ── Gridlines ──────────────────────────────────────────────────────────
    svg.append("g")
        .call(d3.axisLeft(y).tickSize(-w).tickFormat(() => ""))
        .call(g => g.select(".domain").remove())
        .call(g => g.selectAll("line")
            .attr("stroke", "#e8e8e8")
            .attr("stroke-dasharray", "2,2"));

    // ── X Axis ─────────────────────────────────────────────────────────────
    svg.append("g")
        .attr("transform", `translate(0,${h})`)
        .call(d3.axisBottom(x).ticks(6).tickFormat(d3.timeFormat("%d %b %Y")))
        .selectAll("text")
            .attr("transform", "rotate(-35)")
            .style("text-anchor", "end")
            .attr("font-size", 11);

    // ── Y Axis ─────────────────────────────────────────────────────────────
    svg.append("g")
        .call(d3.axisLeft(y))
        .selectAll("text")
            .attr("font-size", 11);

    // ── Statistics ─────────────────────────────────────────────────────────
    const mean   = d3.mean(data, d => d.value);
    const median = d3.median(data, d => d.value);

    // Mean line
    svg.append("line")
        .attr("x1", 0).attr("x2", w)
        .attr("y1", y(mean)).attr("y2", y(mean))
        .attr("stroke", pbi.colors[1])
        .attr("stroke-width", 1.5)
        .attr("stroke-dasharray", "6,3");

    svg.append("text")
        .attr("x", w + 4).attr("y", y(mean) + 4)
        .attr("font-size", 10).attr("fill", pbi.colors[1])
        .text("Mean");

    // Median line
    svg.append("line")
        .attr("x1", 0).attr("x2", w)
        .attr("y1", y(median)).attr("y2", y(median))
        .attr("stroke", pbi.colors[3])
        .attr("stroke-width", 1.5)
        .attr("stroke-dasharray", "3,3");

    svg.append("text")
        .attr("x", w + 4).attr("y", y(median) + 4)
        .attr("font-size", 10).attr("fill", pbi.colors[3])
        .text("Median");

    // ── Run Line ───────────────────────────────────────────────────────────
    svg.append("path")
        .datum(data)
        .attr("fill", "none")
        .attr("stroke", pbi.colors[0])
        .attr("stroke-width", 2)
        .attr("stroke-linejoin", "round")
        .attr("stroke-linecap", "round")
        .attr("d", d3.line()
            .x(d => x(d.date))
            .y(d => y(d.value)));

    // ── Data Points ────────────────────────────────────────────────────────
    // Points above mean → accent colour, below → base colour
    svg.selectAll("circle")
        .data(data)
        .join("circle")
        .attr("cx",     d => x(d.date))
        .attr("cy",     d => y(d.value))
        .attr("r",      4.5)
        .attr("fill",   d => d.value > mean ? pbi.colors[2] : pbi.colors[0])
        .attr("stroke", "#fff")
        .attr("stroke-width", 1.5);

    // ── Axis Labels ────────────────────────────────────────────────────────
    svg.append("text")
        .attr("x", w / 2)
        .attr("y", h + 55)
        .attr("text-anchor", "middle")
        .attr("font-size", 12)
        .attr("fill", "#555")
        .text(dateKey);

    svg.append("text")
        .attr("transform", "rotate(-90)")
        .attr("x", -(h / 2))
        .attr("y", -48)
        .attr("text-anchor", "middle")
        .attr("font-size", 12)
        .attr("fill", "#555")
        .text(valueKey);
});
