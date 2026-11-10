/**
 * Charts (DrawingML c:chartSpace parts): what one shows — its kind, title, categories and series, read
 * from the values Word caches in the chart (so a chart from Word shows without its spreadsheet) — and
 * how it is drawn: a list of simple shapes (rectangles, lines, paths, text) in a w×h box, drawn as SVG
 * on screen and as vector graphics in the PDF.
 */

export type ChartKind = "column" | "bar" | "line" | "pie";

export interface ChartSeries { name: string; values: number[]; color?: string }

export interface ChartSpec {
  kind: ChartKind;
  title?: string;
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

/** What a chart part shows, or undefined for a kind not drawn yet (3-D, area, scatter…: shown as a box). */
export function parseChart(xml: string, theme: Record<string, string> = chartTheme(undefined)): ChartSpec | undefined {
  const plot = /<c:plotArea>([\s\S]*?)<\/c:plotArea>/.exec(xml)?.[1] ?? "";
  const group = /<c:(barChart|lineChart|pieChart|doughnutChart)>([\s\S]*?)<\/c:\1>/.exec(plot);
  if (!group) return undefined;
  const kind: ChartKind = group[1] === "lineChart" ? "line" : group[1] === "barChart" ? (/<c:barDir val="bar"/.test(group[2]) ? "bar" : "column") : "pie";
  const series: ChartSeries[] = [];
  let categories: string[] = [];
  for (const s of group[2].matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const body = s[1];
    const name = points(/<c:tx>([\s\S]*?)<\/c:tx>/.exec(body)?.[1])[0] ?? unescape(/<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>/.exec(body)?.[1] ?? `Series ${series.length + 1}`);
    const cat = points(/<c:cat>([\s\S]*?)<\/c:cat>/.exec(body)?.[1]);
    const val = points(/<c:val>([\s\S]*?)<\/c:val>/.exec(body)?.[1]).map((v) => Number(v));
    if (cat.length > categories.length) categories = cat;
    // The series' own colour: its fill (bars, slices) or line.
    const spPr = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(body.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, ""))?.[1] ?? "";
    const fill = /<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(spPr)?.[1] ?? /<a:ln\b[\s\S]*?<\/a:ln>/.exec(spPr)?.[0] ?? "";
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
  return { kind, ...(title ? { title } : {}), ...(border ? { border } : {}), categories, series };
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

/** The chart drawn in a w×h box: shapes from back to front. */
export function chartShapes(spec: ChartSpec, w: number, h: number): ChartShape[] {
  const out: ChartShape[] = [{ kind: "rect", x: 0, y: 0, w, h, fill: "FFFFFF" }];
  if (spec.border) out.push({ kind: "path", d: `M0.5,0.5 L${w - 0.5},0.5 L${w - 0.5},${h - 0.5} L0.5,${h - 0.5} Z`, stroke: spec.border, width: 0.75 });
  const color = (i: number, s?: ChartSeries) => s?.color ?? CHART_COLORS[i % CHART_COLORS.length];
  const text = "595959";
  let top = 8;
  if (spec.title) {
    out.push({ kind: "text", x: w / 2, y: top + 14, text: spec.title, size: 14, color: text, anchor: "middle" });
    top += 26;
  }
  // The legend below: series (or, for a pie, categories), centred.
  const order = spec.series.map((s, i) => ({ s, i }));
  if (spec.kind === "bar") order.reverse();
  const legend = spec.kind === "pie" ? spec.categories : order.map((o) => o.s.name);
  const legendH = legend.length > 1 || spec.kind === "pie" ? 20 : 0;
  if (legendH) {
    const items = legend.map((name) => ({ name, width: 10 + 4 + textWidth(name, 9) + 12 }));
    let x = (w - items.reduce((a, it) => a + it.width, 0)) / 2;
    const y = h - 14;
    items.forEach((it, n) => {
      const i = spec.kind === "pie" ? n : order[n].i;
      if (spec.kind === "line") out.push({ kind: "line", x1: x, y1: y - 3, x2: x + 12, y2: y - 3, color: color(i, spec.series[i]), width: 2.25 });
      else out.push({ kind: "rect", x, y: y - 8, w: 8, h: 8, fill: spec.kind === "pie" ? color(i) : color(i, spec.series[i]) });
      out.push({ kind: "text", x: x + 14, y, text: it.name, size: 9, color: text, anchor: "start" });
      x += it.width;
    });
  }
  const bottom = h - 8 - legendH;
  if (spec.kind === "pie") {
    const values = spec.series[0]?.values.map((v) => Math.max(0, v || 0)) ?? [];
    const total = values.reduce((a, v) => a + v, 0) || 1;
    const r = Math.max(10, Math.min(w - 16, bottom - top - 8) / 2);
    const cx = w / 2;
    const cy = top + (bottom - top) / 2;
    let a = -Math.PI / 2;
    values.forEach((v, i) => {
      const da = (v / total) * Math.PI * 2;
      if (da <= 0) return;
      const x1 = cx + r * Math.cos(a), y1 = cy + r * Math.sin(a);
      const x2 = cx + r * Math.cos(a + da), y2 = cy + r * Math.sin(a + da);
      const d = da >= Math.PI * 2 - 1e-6
        ? `M${cx - r},${cy} A${r},${r} 0 1,1 ${cx + r},${cy} A${r},${r} 0 1,1 ${cx - r},${cy} Z`
        : `M${cx},${cy} L${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${da > Math.PI ? 1 : 0},1 ${x2.toFixed(2)},${y2.toFixed(2)} Z`;
      out.push({ kind: "path", d, fill: color(i), stroke: "FFFFFF", width: 1 });
      a += da;
    });
    return out;
  }
  // Axes: values from 0 (or below, for negatives) to a round maximum.
  const all = spec.series.flatMap((s) => s.values.filter((v) => Number.isFinite(v)));
  const max = Math.max(0, ...all);
  const min = Math.min(0, ...all);
  const step = niceStep(max - min);
  const hi = Math.ceil((max + (max - min) * 0.05) / step) * step || step;
  const lo = min < 0 ? Math.floor((min - (max - min) * 0.05) / step) * step : 0;
  const ticks: number[] = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Math.round(v * 1e6) / 1e6);
  const labelW = Math.max(...ticks.map((t) => textWidth(fmt(t), 9))) + 6;
  const n = spec.categories.length;
  if (spec.kind === "bar") {
    // Horizontal bars: categories down the side, values along the bottom.
    const catW = Math.min(w * 0.3, Math.max(...spec.categories.map((c) => textWidth(c, 9))) + 8);
    const left = 8 + catW;
    const right = w - 12;
    const plotTop = top + 4;
    const plotBottom = bottom - 16;
    const X = (v: number) => left + ((v - lo) / (hi - lo)) * (right - left);
    for (const t of ticks) {
      out.push({ kind: "line", x1: X(t), y1: plotTop, x2: X(t), y2: plotBottom, color: "D9D9D9", width: 0.75 });
      out.push({ kind: "text", x: X(t), y: plotBottom + 12, text: fmt(t), size: 9, color: text, anchor: "middle" });
    }
    const band = (plotBottom - plotTop) / Math.max(1, n);
    const k = spec.series.length;
    const barH = (band / (k + 1.2)) ;
    spec.categories.forEach((c, i) => {
      const y0 = plotBottom - band * (i + 1) + band * 0.1;
      out.push({ kind: "text", x: left - 4, y: y0 + band * 0.4 + 3, text: c, size: 9, color: text, anchor: "end" });
      spec.series.forEach((s, j) => {
        const v = s.values[i] ?? 0;
        const xa = X(Math.min(0, v)), xb = X(Math.max(0, v));
        out.push({ kind: "rect", x: xa, y: y0 + barH * (k - 1 - j + 0.1), w: Math.max(0.5, xb - xa), h: barH * 0.95, fill: color(j, s) });
      });
    });
    out.push({ kind: "line", x1: X(0), y1: plotTop, x2: X(0), y2: plotBottom, color: "BFBFBF", width: 0.75 });
    return out;
  }
  const left = 8 + labelW;
  const right = w - 12;
  const plotTop = top + 4;
  const plotBottom = bottom - 16;
  const Y = (v: number) => plotBottom - ((v - lo) / (hi - lo)) * (plotBottom - plotTop);
  for (const t of ticks) {
    out.push({ kind: "line", x1: left, y1: Y(t), x2: right, y2: Y(t), color: "D9D9D9", width: 0.75 });
    out.push({ kind: "text", x: left - 4, y: Y(t) + 3, text: fmt(t), size: 9, color: text, anchor: "end" });
  }
  const band = (right - left) / Math.max(1, n);
  spec.categories.forEach((c, i) => {
    out.push({ kind: "text", x: left + band * (i + 0.5), y: plotBottom + 12, text: c, size: 9, color: text, anchor: "middle" });
  });
  if (spec.kind === "column") {
    const k = spec.series.length;
    const barW = band / (k + 1.2);
    spec.series.forEach((s, j) => {
      s.values.forEach((v, i) => {
        if (!Number.isFinite(v)) return;
        const ya = Y(Math.max(0, v)), yb = Y(Math.min(0, v));
        out.push({ kind: "rect", x: left + band * i + band * 0.1 + barW * (j + 0.1), y: ya, w: barW * 0.95, h: Math.max(0.5, yb - ya), fill: color(j, s) });
      });
    });
  } else {
    spec.series.forEach((s, j) => {
      const pts = s.values.map((v, i) => (Number.isFinite(v) ? `${(left + band * (i + 0.5)).toFixed(2)},${Y(v).toFixed(2)}` : "")).filter(Boolean);
      if (pts.length) out.push({ kind: "path", d: `M${pts.join(" L")}`, stroke: color(j, s), width: 2.25 });
    });
  }
  out.push({ kind: "line", x1: left, y1: Y(0), x2: right, y2: Y(0), color: "BFBFBF", width: 0.75 });
  return out;
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
