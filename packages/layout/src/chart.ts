/**
 * Charts (DrawingML c:chartSpace parts): what one shows — its kind, title, categories and series, read
 * from the values Word caches in the chart (so a chart from Word shows without its spreadsheet) — and
 * how it is drawn: a list of simple shapes (rectangles, lines, paths, text) in a w×h box, drawn as SVG
 * on screen and as vector graphics in the PDF.
 */

export type ChartKind = "column" | "bar" | "line" | "pie" | "area" | "doughnut" | "scatter";

/** Where the legend goes (none: no legend). */
export type LegendPos = "bottom" | "top" | "right" | "none";

export interface ChartSeries { name: string; values: number[]; color?: string }

export interface ChartSpec {
  kind: ChartKind;
  title?: string;
  /** Column, bar and area: series stacked on each other (Word's "stacked" grouping). */
  stacked?: boolean;
  /** The legend's place (default: bottom). */
  legend?: LegendPos;
  /** Each value written on its bar, point or slice. */
  labels?: boolean;
  /** The value gridlines (default: shown). */
  gridlines?: boolean;
  /** The colour set: series (and pie slices) in this order, unless a series has its own. Scatter: categories are the x values. */
  colors?: string[];
  /** The chart's border colour (its outline), if it has one. */
  border?: string;
  categories: string[];
  series: ChartSeries[];
}

/** Office's chart colours (accent 1–6 of Word's current default theme, used when a document has none). */
export const CHART_COLORS = ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"];

/** The theme's colours by name (a:clrScheme in the theme part), for colours given as accent 1–6 etc. */
export function chartTheme(themeXml: string | undefined): Record<string, string> {
  const out: Record<string, string> = { dk1: "000000", lt1: "FFFFFF", dk2: "0E2841", lt2: "E8E8E8", ...Object.fromEntries(CHART_COLORS.map((c, i) => [`accent${i + 1}`, c])) };
  const scheme = themeXml ? /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(themeXml)?.[0] ?? "" : "";
  for (const m of scheme.matchAll(/<a:(dk1|lt1|dk2|lt2|accent[1-6])>([\s\S]*?)<\/a:\1>/g)) {
    const v = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[2])?.[1] ?? /lastClr="([0-9A-Fa-f]{6})"/.exec(m[2])?.[1];
    if (v) out[m[1]] = v.toUpperCase();
  }
  return out;
}

/** The first colour in a piece of DrawingML (srgbClr, or schemeClr through the theme, with lumMod / lumOff / shade / tint). */
function firstColor(xml: string, theme: Record<string, string>): string | undefined {
  const m = /<a:(srgbClr|schemeClr) val="([^"]+)"(\/>|>([\s\S]*?)<\/a:\1>)/.exec(xml);
  if (!m) return undefined;
  const alias: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };
  const base = m[1] === "srgbClr" ? m[2] : theme[alias[m[2]] ?? m[2]];
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return undefined;
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16));
  for (const t of (m[4] ?? "").matchAll(/<a:(lumMod|lumOff|shade|tint) val="(\d+)"/g)) {
    const v = Number(t[2]) / 100000;
    if (t[1] === "lumMod" || t[1] === "shade") [r, g, b] = [r * v, g * v, b * v];
    else if (t[1] === "lumOff") [r, g, b] = [r + 255 * v, g + 255 * v, b + 255 * v];
    else [r, g, b] = [r + (255 - r) * (1 - v), g + (255 - g) * (1 - v), b + (255 - b) * (1 - v)];
  }
  return [r, g, b].map((c) => Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, "0")).join("").toUpperCase();
}

const unescape = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** The cached points of a c:strCache / c:numCache (or literal) inside `xml`, by index. */
function points(xml: string | undefined): string[] {
  if (!xml) return [];
  const out: string[] = [];
  for (const m of xml.matchAll(/<c:pt idx="(\d+)"[^>]*>\s*<c:v>([\s\S]*?)<\/c:v>/g)) out[Number(m[1])] = unescape(m[2]);
  return out;
}

/** What a chart part shows, or undefined for a kind not drawn yet (3-D, bubble, stock…: shown as a box). */
export function parseChart(xml: string, theme: Record<string, string> = chartTheme(undefined)): ChartSpec | undefined {
  const plot = /<c:plotArea>([\s\S]*?)<\/c:plotArea>/.exec(xml)?.[1] ?? "";
  const group = /<c:(barChart|lineChart|pieChart|doughnutChart|areaChart|scatterChart)>([\s\S]*?)<\/c:\1>/.exec(plot);
  if (!group) return undefined;
  const kind: ChartKind = group[1] === "lineChart" ? "line" : group[1] === "barChart" ? (/<c:barDir val="bar"/.test(group[2]) ? "bar" : "column")
    : group[1] === "areaChart" ? "area" : group[1] === "scatterChart" ? "scatter" : group[1] === "doughnutChart" ? "doughnut" : "pie";
  // Stacked (and 100% stacked, drawn stacked) columns, bars and areas.
  const stacked = /<c:grouping val="(stacked|percentStacked)"/.test(group[2]);
  const series: ChartSeries[] = [];
  let categories: string[] = [];
  for (const s of group[2].matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const body = s[1];
    const name = points(/<c:tx>([\s\S]*?)<\/c:tx>/.exec(body)?.[1])[0] ?? unescape(/<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>/.exec(body)?.[1] ?? `Series ${series.length + 1}`);
    // (A scatter chart's x values stand in the categories' place.)
    const cat = points(/<c:(?:cat|xVal)>([\s\S]*?)<\/c:(?:cat|xVal)>/.exec(body)?.[1]);
    const val = points(/<c:(?:val|yVal)>([\s\S]*?)<\/c:(?:val|yVal)>/.exec(body)?.[1]).map((v) => Number(v));
    if (cat.length > categories.length) categories = cat;
    // The series' own colour: its fill (bars, slices) or line.
    const spPr = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(body.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, ""))?.[1] ?? "";
    // (A scatter series without a line: its markers' fill.)
    const marker = /<c:marker>[\s\S]*?<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(body)?.[1] ?? "";
    const fill = /<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(spPr)?.[1]
      ?? (/<a:ln\b[^>]*>\s*<a:noFill\/>/.test(spPr) ? /<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(marker)?.[1] : undefined)
      ?? /<a:ln\b[\s\S]*?<\/a:ln>/.exec(spPr)?.[0] ?? "";
    const color = firstColor(fill, theme);
    series.push({ name, values: val, ...(color ? { color } : {}) });
  }
  if (!series.length) return undefined;
  const n = Math.max(categories.length, ...series.map((s) => s.values.length));
  for (let i = 0; i < n; i++) categories[i] ??= String(i + 1);
  const titleXml = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml)?.[1];
  const title = titleXml ? [...titleXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => unescape(m[1])).join("") || undefined : undefined;
  // The chart's own outline (the chart space's spPr, after the chart).
  const outer = /<\/c:chart>([\s\S]*)$/.exec(xml)?.[1] ?? "";
  const ln = /<c:spPr>[\s\S]*?(<a:ln\b[\s\S]*?<\/a:ln>)/.exec(outer)?.[1];
  const border = ln && !/<a:noFill\/>/.test(ln) ? firstColor(ln, theme) : undefined;
  // The legend's place (none: no c:legend), values written on the chart, the value gridlines.
  const legendXml = /<c:legend>([\s\S]*?)<\/c:legend>/.exec(xml)?.[1];
  const pos = legendXml ? /<c:legendPos val="(\w+)"/.exec(legendXml)?.[1] ?? "r" : undefined;
  const legend: LegendPos = !legendXml ? "none" : pos === "b" ? "bottom" : pos === "t" ? "top" : "right";
  const labels = /<c:showVal val="1"\/>/.test(group[2]);
  const gridlines = /<c:valAx>[\s\S]*?<c:majorGridlines\b[\s\S]*?<\/c:valAx>/.test(plot);
  // A pie's or doughnut's slice colours (each data point's own fill).
  let colors: string[] | undefined;
  if (kind === "pie" || kind === "doughnut") {
    const first = /<c:ser>([\s\S]*?)<\/c:ser>/.exec(group[2])?.[1] ?? "";
    const slices: string[] = [];
    for (const d of first.matchAll(/<c:dPt>([\s\S]*?)<\/c:dPt>/g)) {
      const i = Number(/<c:idx val="(\d+)"/.exec(d[1])?.[1]);
      const c = firstColor(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(d[1])?.[1] ?? "", theme);
      if (Number.isFinite(i) && c) slices[i] = c;
    }
    if (slices.length) colors = Array.from({ length: slices.length }, (_, i) => slices[i] ?? CHART_COLORS[i % CHART_COLORS.length]);
  }
  return {
    kind, ...(stacked && kind !== "line" ? { stacked } : {}), ...(title ? { title } : {}), ...(border ? { border } : {}),
    ...(legend !== "bottom" ? { legend } : {}), ...(labels ? { labels } : {}), ...(!gridlines && kind !== "pie" && kind !== "doughnut" ? { gridlines: false } : {}),
    ...(colors ? { colors } : {}), categories, series,
  };
}

export type ChartShape =
  | { kind: "rect"; x: number; y: number; w: number; h: number; fill: string }
  | { kind: "line"; x1: number; y1: number; x2: number; y2: number; color: string; width: number }
  | { kind: "path"; d: string; fill?: string; stroke?: string; width?: number }
  | { kind: "text"; x: number; y: number; text: string; size: number; color: string; anchor: "start" | "middle" | "end"; bold?: boolean };

/** A round number step for an axis from 0 to `max` (1, 2, 5 × 10ⁿ). */
function niceStep(range: number, ticks = 5): number {
  if (range <= 0) return 1;
  const raw = range / ticks;
  const p = 10 ** Math.floor(Math.log10(raw));
  const f = raw / p;
  return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p;
}

const fmt = (v: number) => (Math.abs(v) >= 1000 ? v.toLocaleString("en-US") : String(Math.round(v * 100) / 100));
/** A rough text width (pt) at `size`, for laying out labels (Calibri-like proportions). */
const textWidth = (t: string, size: number) => t.length * size * 0.5;

/** Colour sets for charts (Insert → Chart → Colours): series, or a pie's slices, in this order. */
export const CHART_PALETTES: { id: string; name: string; colors: string[] }[] = [
  { id: "office", name: "Office", colors: CHART_COLORS },
  { id: "colorful", name: "Colorful", colors: ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"] },
  { id: "blue", name: "Blue", colors: ["1F3864", "2F5597", "4472C4", "8FAADC", "B4C7E7", "5B9BD5"] },
  { id: "green", name: "Green", colors: ["375623", "548235", "70AD47", "A9D18E", "C5E0B4", "8FBC5A"] },
  { id: "warm", name: "Warm", colors: ["843C0C", "C55A11", "ED7D31", "F4B183", "FFC000", "BF9000"] },
  { id: "gray", name: "Grey", colors: ["262626", "595959", "7F7F7F", "A6A6A6", "BFBFBF", "404040"] },
  { id: "pastel", name: "Pastel", colors: ["8EB4E3", "F4B183", "A9D18E", "FFD966", "C9A0DC", "9DC3E6"] },
];

/** Dark text on a light colour, white on a dark one (a label inside a bar or slice). */
const onColor = (hex: string) => {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "404040" : "FFFFFF";
};

/** The chart drawn in a w×h box: shapes from back to front. */
export function chartShapes(spec: ChartSpec, w: number, h: number): ChartShape[] {
  const out: ChartShape[] = [{ kind: "rect", x: 0, y: 0, w, h, fill: "FFFFFF" }];
  if (spec.border) out.push({ kind: "path", d: `M0.5,0.5 L${w - 0.5},0.5 L${w - 0.5},${h - 0.5} L0.5,${h - 0.5} Z`, stroke: spec.border, width: 0.75 });
  const palette = spec.colors?.length ? spec.colors : CHART_COLORS;
  const color = (i: number, s?: ChartSeries) => s?.color ?? palette[i % palette.length];
  const text = "595959";
  const labelColor = "404040";
  const round = spec.kind === "pie" || spec.kind === "doughnut";
  const stacked = !!spec.stacked && (spec.kind === "column" || spec.kind === "bar" || spec.kind === "area");
  const grid = spec.gridlines !== false;
  const label = (x: number, y: number, v: number, anchor: "start" | "middle" | "end" = "middle", fill = labelColor) =>
    out.push({ kind: "text", x, y, text: fmt(v), size: 8, color: fill, anchor });
  let top = 8;
  let right = w - 12;
  if (spec.title) {
    out.push({ kind: "text", x: w / 2, y: top + 14, text: spec.title, size: 14, color: text, anchor: "middle" });
    top += 26;
  }
  // The legend: series (or, for a pie or doughnut, categories); below, above, to the right, or none.
  const order = spec.series.map((s, i) => ({ s, i }));
  if (spec.kind === "bar") order.reverse();
  const names = round ? spec.categories : order.map((o) => o.s.name);
  const place = spec.legend ?? "bottom";
  const shown = place !== "none" && (names.length > 1 || round);
  const key = (x: number, y: number, n: number) => {
    const i = round ? n : order[n].i;
    if (spec.kind === "line") out.push({ kind: "line", x1: x, y1: y - 3, x2: x + 12, y2: y - 3, color: color(i, spec.series[i]), width: 2.25 });
    else if (spec.kind === "scatter") out.push({ kind: "path", d: circle(x + 4, y - 4, 3.5), fill: color(i, spec.series[i]) });
    else out.push({ kind: "rect", x, y: y - 8, w: 8, h: 8, fill: round ? color(i) : color(i, spec.series[i]) });
    out.push({ kind: "text", x: x + 14, y, text: names[n], size: 9, color: text, anchor: "start" });
  };
  let bottom = h - 8;
  if (shown && place === "right") {
    const lw = Math.min(w * 0.35, Math.max(...names.map((nm) => textWidth(nm, 9))) + 26);
    const y0 = (top + bottom) / 2 - (names.length * 16) / 2 + 12;
    names.forEach((_, n) => key(w - lw, y0 + n * 16, n));
    right = w - lw - 12;
  } else if (shown) {
    const items = names.map((name) => 10 + 4 + textWidth(name, 9) + 12);
    let x = (w - items.reduce((a, it) => a + it, 0)) / 2;
    const y = place === "top" ? top + 10 : h - 14;
    names.forEach((_, n) => { key(x, y, n); x += items[n]; });
    if (place === "top") top += 20;
    else bottom -= 20;
  }
  if (round) {
    const values = spec.series[0]?.values.map((v) => Math.max(0, v || 0)) ?? [];
    const total = values.reduce((a, v) => a + v, 0) || 1;
    const r = Math.max(10, Math.min(right - 4, bottom - top - 8) / 2);
    const cx = (right + 12) / 2;
    const cy = top + (bottom - top) / 2;
    let a = -Math.PI / 2;
    const marks: { x: number; y: number; v: number; fill: string }[] = [];
    values.forEach((v, i) => {
      const da = (v / total) * Math.PI * 2;
      if (da <= 0) return;
      const x1 = cx + r * Math.cos(a), y1 = cy + r * Math.sin(a);
      const x2 = cx + r * Math.cos(a + da), y2 = cy + r * Math.sin(a + da);
      const d = da >= Math.PI * 2 - 1e-6
        ? circle(cx, cy, r)
        : `M${cx},${cy} L${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${da > Math.PI ? 1 : 0},1 ${x2.toFixed(2)},${y2.toFixed(2)} Z`;
      out.push({ kind: "path", d, fill: color(i), stroke: "FFFFFF", width: 1 });
      // Its value inside the slice (in the ring, for a doughnut).
      const mid = a + da / 2, at = spec.kind === "doughnut" ? 0.75 : 0.62;
      marks.push({ x: cx + r * at * Math.cos(mid), y: cy + r * at * Math.sin(mid) + 3, v: spec.series[0].values[i], fill: onColor(color(i)) });
      a += da;
    });
    // A doughnut: the middle cut out (half the radius, Word's hole size).
    if (spec.kind === "doughnut") out.push({ kind: "path", d: circle(cx, cy, r * 0.5), fill: "FFFFFF", stroke: "FFFFFF", width: 1 });
    if (spec.labels) for (const m of marks) label(m.x, m.y, m.v, "middle", m.fill);
    return out;
  }
  // The value range: each series' values, or (stacked) each category's total above and below zero.
  const finite = (v: number | undefined) => (v !== undefined && Number.isFinite(v) ? v : 0);
  const n = spec.categories.length;
  let all: number[];
  if (stacked) {
    all = [];
    for (let i = 0; i < n; i++) {
      all.push(spec.series.reduce((a, s) => a + Math.max(0, finite(s.values[i])), 0));
      all.push(spec.series.reduce((a, s) => a + Math.min(0, finite(s.values[i])), 0));
    }
  } else all = spec.series.flatMap((s) => s.values.filter((v) => Number.isFinite(v)));
  const range = (vals: number[]) => {
    const max = Math.max(0, ...vals);
    const min = Math.min(0, ...vals);
    const step = niceStep(max - min);
    const hi = Math.ceil((max + (max - min) * 0.05) / step) * step || step;
    const lo = min < 0 ? Math.floor((min - (max - min) * 0.05) / step) * step : 0;
    const ticks: number[] = [];
    for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Math.round(v * 1e6) / 1e6);
    return { lo, hi, ticks };
  };
  const { lo, hi, ticks } = range(all);
  const labelW = Math.max(...ticks.map((t) => textWidth(fmt(t), 9))) + 6;
  // Where each bar segment starts and ends (stacked: on top of the series before it, by sign).
  const segments = (i: number) => {
    let up = 0, down = 0;
    return spec.series.map((s) => {
      const v = s.values[i];
      if (v === undefined || !Number.isFinite(v)) return undefined;
      if (!stacked) return { from: 0, to: v };
      if (v >= 0) { const from = up; up += v; return { from, to: up }; }
      const from = down; down += v; return { from, to: down };
    });
  };
  const plotTop = top + 4;
  const plotBottom = bottom - 16;
  if (spec.kind === "bar") {
    // Horizontal bars: categories down the side, values along the bottom.
    const catW = Math.min(w * 0.3, Math.max(...spec.categories.map((c) => textWidth(c, 9))) + 8);
    const left = 8 + catW;
    const X = (v: number) => left + ((v - lo) / (hi - lo)) * (right - left);
    for (const t of ticks) {
      if (grid) out.push({ kind: "line", x1: X(t), y1: plotTop, x2: X(t), y2: plotBottom, color: "D9D9D9", width: 0.75 });
      out.push({ kind: "text", x: X(t), y: plotBottom + 12, text: fmt(t), size: 9, color: text, anchor: "middle" });
    }
    const band = (plotBottom - plotTop) / Math.max(1, n);
    const k = spec.series.length;
    const barH = stacked ? band * 0.6 : band / (k + 1.2);
    const marks: (() => void)[] = [];
    spec.categories.forEach((c, i) => {
      const y0 = plotBottom - band * (i + 1) + band * 0.1;
      out.push({ kind: "text", x: left - 4, y: y0 + band * 0.4 + 3, text: c, size: 9, color: text, anchor: "end" });
      segments(i).forEach((seg, j) => {
        if (!seg) return;
        const xa = X(Math.min(seg.from, seg.to)), xb = X(Math.max(seg.from, seg.to));
        const y = stacked ? y0 + band * 0.1 : y0 + barH * (k - 1 - j + 0.1);
        const fill = color(j, spec.series[j]);
        out.push({ kind: "rect", x: xa, y, w: Math.max(0.5, xb - xa), h: barH * 0.95, fill });
        const v = seg.to - seg.from, ty = y + barH * 0.475 + 3;
        if (spec.labels) marks.push(stacked ? () => label((xa + xb) / 2, ty, v, "middle", onColor(fill)) : () => label(v >= 0 ? xb + 3 : xa - 3, ty, v, v >= 0 ? "start" : "end"));
      });
    });
    out.push({ kind: "line", x1: X(0), y1: plotTop, x2: X(0), y2: plotBottom, color: "BFBFBF", width: 0.75 });
    marks.forEach((m) => m());
    return out;
  }
  const left = 8 + labelW;
  const Y = (v: number) => plotBottom - ((v - lo) / (hi - lo)) * (plotBottom - plotTop);
  for (const t of ticks) {
    if (grid) out.push({ kind: "line", x1: left, y1: Y(t), x2: right, y2: Y(t), color: "D9D9D9", width: 0.75 });
    out.push({ kind: "text", x: left - 4, y: Y(t) + 3, text: fmt(t), size: 9, color: text, anchor: "end" });
  }
  if (spec.kind === "scatter") {
    // Scatter: the categories are x values (numbers), on an axis of their own; a marker per point.
    const xs = spec.categories.map((c) => Number(c));
    const xr = range(xs.filter((v) => Number.isFinite(v)));
    const X = (v: number) => left + ((v - xr.lo) / (xr.hi - xr.lo)) * (right - left);
    for (const t of xr.ticks) {
      if (grid) out.push({ kind: "line", x1: X(t), y1: plotTop, x2: X(t), y2: plotBottom, color: "D9D9D9", width: 0.75 });
      out.push({ kind: "text", x: X(t), y: plotBottom + 12, text: fmt(t), size: 9, color: text, anchor: "middle" });
    }
    out.push({ kind: "line", x1: left, y1: Y(0), x2: right, y2: Y(0), color: "BFBFBF", width: 0.75 });
    spec.series.forEach((s, j) => {
      s.values.forEach((v, i) => {
        if (!Number.isFinite(v) || !Number.isFinite(xs[i])) return;
        out.push({ kind: "path", d: circle(X(xs[i]), Y(v), 3.5), fill: color(j, s) });
        if (spec.labels) label(X(xs[i]) + 5, Y(v) - 4, v, "start");
      });
    });
    return out;
  }
  const band = (right - left) / Math.max(1, n);
  spec.categories.forEach((c, i) => {
    out.push({ kind: "text", x: left + band * (i + 0.5), y: plotBottom + 12, text: c, size: 9, color: text, anchor: "middle" });
  });
  const marks: (() => void)[] = [];
  if (spec.kind === "column") {
    const k = spec.series.length;
    const barW = stacked ? band * 0.6 : band / (k + 1.2);
    for (let i = 0; i < n; i++) {
      segments(i).forEach((seg, j) => {
        if (!seg) return;
        const ya = Y(Math.max(seg.from, seg.to)), yb = Y(Math.min(seg.from, seg.to));
        const x = stacked ? left + band * i + band * 0.2 : left + band * i + band * 0.1 + barW * (j + 0.1);
        const fill = color(j, spec.series[j]);
        out.push({ kind: "rect", x, y: ya, w: barW * 0.95, h: Math.max(0.5, yb - ya), fill });
        const v = seg.to - seg.from, cx = x + barW * 0.475;
        if (spec.labels) marks.push(stacked ? () => label(cx, (ya + yb) / 2 + 3, v, "middle", onColor(fill)) : () => label(cx, v >= 0 ? ya - 3 : yb + 9, v));
      });
    }
  } else if (spec.kind === "area") {
    // Areas: each series filled down to zero (or, stacked, down to the series below it); later ones in front.
    const X = (i: number) => left + band * (i + 0.5);
    const base = new Array<number>(n).fill(0);
    spec.series.forEach((s, j) => {
      const lower = stacked ? [...base] : new Array<number>(n).fill(0);
      const upper = lower.map((b, i) => b + finite(s.values[i]));
      if (stacked) upper.forEach((v, i) => { base[i] = v; });
      if (!n) return;
      const d = `M${upper.map((v, i) => `${X(i).toFixed(2)},${Y(v).toFixed(2)}`).join(" L")} L${lower.map((v, i) => `${X(i).toFixed(2)},${Y(v).toFixed(2)}`).reverse().join(" L")} Z`;
      out.push({ kind: "path", d, fill: color(j, s) });
      if (spec.labels) upper.forEach((v, i) => { if (Number.isFinite(s.values[i])) marks.push(() => label(X(i), Y(v) - 3, s.values[i])); });
    });
  } else {
    spec.series.forEach((s, j) => {
      const pts = s.values.map((v, i) => (Number.isFinite(v) ? `${(left + band * (i + 0.5)).toFixed(2)},${Y(v).toFixed(2)}` : "")).filter(Boolean);
      if (pts.length) out.push({ kind: "path", d: `M${pts.join(" L")}`, stroke: color(j, s), width: 2.25 });
      if (spec.labels) s.values.forEach((v, i) => { if (Number.isFinite(v)) marks.push(() => label(left + band * (i + 0.5), Y(v) - 5, v)); });
    });
  }
  out.push({ kind: "line", x1: left, y1: Y(0), x2: right, y2: Y(0), color: "BFBFBF", width: 0.75 });
  marks.forEach((m) => m());
  return out;
}

/** A circle as a path (two arcs). */
function circle(cx: number, cy: number, r: number): string {
  return `M${(cx - r).toFixed(2)},${cy.toFixed(2)} A${r},${r} 0 1,1 ${(cx + r).toFixed(2)},${cy.toFixed(2)} A${r},${r} 0 1,1 ${(cx - r).toFixed(2)},${cy.toFixed(2)} Z`;
}

const escXml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/** The chart as an SVG document (for the screen). */
export function chartSvg(spec: ChartSpec, w: number, h: number): string {
  const body = chartShapes(spec, w, h).map((s) => {
    switch (s.kind) {
      case "rect": return `<rect x="${s.x.toFixed(2)}" y="${s.y.toFixed(2)}" width="${s.w.toFixed(2)}" height="${s.h.toFixed(2)}" fill="#${s.fill}"/>`;
      case "line": return `<line x1="${s.x1.toFixed(2)}" y1="${s.y1.toFixed(2)}" x2="${s.x2.toFixed(2)}" y2="${s.y2.toFixed(2)}" stroke="#${s.color}" stroke-width="${s.width}"/>`;
      case "path": return `<path d="${s.d}" fill="${s.fill ? `#${s.fill}` : "none"}"${s.stroke ? ` stroke="#${s.stroke}" stroke-width="${s.width ?? 1}" stroke-linejoin="round"` : ""}/>`;
      case "text": return `<text x="${s.x.toFixed(2)}" y="${s.y.toFixed(2)}" font-family="Calibri, Carlito, Arial, sans-serif" font-size="${s.size}" fill="#${s.color}" text-anchor="${s.anchor}"${s.bold ? ' font-weight="700"' : ""}>${escXml(s.text)}</text>`;
    }
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}pt" height="${h}pt">${body}</svg>`;
}
