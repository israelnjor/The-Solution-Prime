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

  private normalLineHeight(font: string): number {
    if (typeof document === "undefined" || !document.body) return 0;
    const el = document.createElement("span");
    el.style.cssText = "position:absolute;visibility:hidden;white-space:nowrap;line-height:normal;display:inline-block";
    el.style.font = font;
    el.style.lineHeight = "normal";
    el.textContent = "Hg";
    document.body.appendChild(el);
    const h = el.getBoundingClientRect().height / PROBE;
    el.remove();
    return h;
  }

  metrics(f: FontKey): FontMetrics {
    const key = this.font(f);
    let m = this.metricsCache.get(key);
    if (!m) {
      this.ctx.font = key;
      const t = this.ctx.measureText("Hg");
      const ascent = (t.fontBoundingBoxAscent ?? PROBE * 0.9) / PROBE;
      const descent = (t.fontBoundingBoxDescent ?? PROBE * 0.25) / PROBE;
      // The font's external leading (Word adds it: Times New Roman 12pt lines are 13.8pt): the
      // browser's normal line height beyond ascent + descent, measured once per font.
      const extra = Math.max(0, this.normalLineHeight(key) - ascent - descent);
      m = { lineHeight: ascent + descent + extra, ascent: ascent + extra };
      this.metricsCache.set(key, m);
    }
    return m;
  }
}
