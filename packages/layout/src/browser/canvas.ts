/**
 * Browser measurer: text widths from the canvas, line metrics from the browser's font metrics.
 * The editor uses it; the Node measurer (node/fonts.ts) remains the reference for PDF and tests.
 */
import type { FontKey, FontMetrics, Measurer } from "../measure.js";

const PROBE = 100; // px; widths are returned per 1pt-equivalent unit (divided back out)

export class CanvasMeasurer implements Measurer {
  private ctx: CanvasRenderingContext2D;
  private widths = new Map<string, number>();
  private metricsCache = new Map<string, FontMetrics>();

  constructor() {
    const canvas = document.createElement("canvas");
    this.ctx = canvas.getContext("2d")!;
  }

  private font(f: FontKey) {
    return `${f.italic ? "italic " : ""}${f.bold ? "bold " : ""}${PROBE}px "${f.family}"`;
  }

  width(text: string, f: FontKey): number {
    const key = `${this.font(f)}|${text}`;
    let w = this.widths.get(key);
    if (w === undefined) {
      this.ctx.font = this.font(f);
      w = this.ctx.measureText(text).width / PROBE;
      if (this.widths.size > 20000) this.widths.clear();
      this.widths.set(key, w);
    }
    return w;
  }

  metrics(f: FontKey): FontMetrics {
    const key = this.font(f);
    let m = this.metricsCache.get(key);
    if (!m) {
      this.ctx.font = key;
      const t = this.ctx.measureText("Hg");
      const ascent = (t.fontBoundingBoxAscent ?? PROBE * 0.9) / PROBE;
      const descent = (t.fontBoundingBoxDescent ?? PROBE * 0.25) / PROBE;
      m = { lineHeight: ascent + descent, ascent };
      this.metricsCache.set(key, m);
    }
    return m;
  }
}
