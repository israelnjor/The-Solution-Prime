/**
 * PDF export (ADR-008): draws the layout engine's pages with pdf-lib. One rendering path for export,
 * print and preview — the PDF shows exactly the pages the layout engine computed.
 */
import fontkit from "@pdf-lib/fontkit";
import { LineCapStyle, PDFDocument, type PDFFont, type PDFImage, type PDFPage, StandardFonts, rgb } from "pdf-lib";
import type { FontKey, ImageSource, LayoutResult, Page, Segment } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";

/** Supplies font files for embedding, and splits text into runs per drawing font (glyph fallback). */
export interface FontSource {
  fontData(font: FontKey): Uint8Array | undefined;
  split?(text: string, font: FontKey): { text: string; font: FontKey }[];
}

/**
 * Symbol fonts (Symbol, Wingdings…): Word writes their characters at U+F020–F0FF, while the fonts
 * themselves often list them at 0x20–0xFF. Use the font's own code where it has only that one.
 */
export function symbolText(text: string, has: (cp: number) => boolean): string {
  return text.replace(/[-]/g, (c) => {
    const cp = c.charCodeAt(0);
    return !has(cp) && has(cp - 0xf000) ? String.fromCharCode(cp - 0xf000) : c;
  });
}

const HIGHLIGHT: Record<string, string> = {
  yellow: "FFFF00", green: "00FF00", cyan: "00FFFF", magenta: "FF00FF", blue: "0000FF", red: "FF0000",
  darkBlue: "000080", darkCyan: "008080", darkGreen: "008000", darkMagenta: "800080", darkRed: "800000",
  darkYellow: "808000", darkGray: "808080", lightGray: "C0C0C0", black: "000000", white: "FFFFFF",
};

function color(hex: string | undefined) {
  const h = hex && /^[0-9A-F]{6}$/i.test(hex) ? hex : "000000";
  return rgb(parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255);
}

const fromBase64 = (b64: string) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));

class Renderer {
  private fonts = new Map<string, Promise<PDFFont>>();
  private images = new Map<string, Promise<PDFImage | undefined>>();

  constructor(private pdf: PDFDocument, private doc: PrimeDocument, private source: FontSource, private total: number) {}

  private font(key: FontKey): Promise<PDFFont> {
    const id = `${key.family}|${key.bold}|${key.italic}`;
    let font = this.fonts.get(id);
    if (!font) {
      const data = this.source.fontData(key);
      font = data
        ? this.pdf.embedFont(data, { subset: true }).catch(() => this.pdf.embedFont(StandardFonts.Helvetica))
        : this.pdf.embedFont(StandardFonts.Helvetica);
      this.fonts.set(id, font);
    }
    return font;
  }

  private image(src: ImageSource): Promise<PDFImage | undefined> {
    const id = src.imageId ?? src.part ?? "";
    let image = this.images.get(id);
    if (!image) {
      const res = src.imageId ? this.doc.images[src.imageId] : undefined;
      const b64 = res?.data ?? this.doc.package.parts.find((p) => p.name === src.part)?.data;
      const ext = (res?.extension ?? src.part?.split(".").pop() ?? "").toLowerCase();
      image = !b64 ? Promise.resolve(undefined)
        : ext === "png" ? this.pdf.embedPng(fromBase64(b64))
        : ext === "jpg" || ext === "jpeg" ? this.pdf.embedJpg(fromBase64(b64))
        : Promise.resolve(undefined); // EMF/WMF/GIF etc.: drawn as a placeholder box for now
      this.images.set(id, image.catch(() => undefined));
    }
    return this.images.get(id)!;
  }

  async page(p: Page) {
    const out = this.pdf.addPage([p.width, p.height]);
    const Y = (y: number) => p.height - y;
    for (const r of p.rects) {
      if (r.fill) out.drawRectangle({ x: r.x, y: Y(r.y + r.height), width: r.width, height: r.height, color: color(r.fill) });
      if (r.sides) {
        const stroke = { thickness: r.stroke?.width ?? 0.5, color: color(r.stroke?.color) };
        const line = (x1: number, y1: number, x2: number, y2: number) =>
          out.drawLine({ start: { x: x1, y: Y(y1) }, end: { x: x2, y: Y(y2) }, ...stroke });
        if (r.sides.top) line(r.x, r.y, r.x + r.width, r.y);
        if (r.sides.bottom) line(r.x, r.y + r.height, r.x + r.width, r.y + r.height);
        if (r.sides.left) line(r.x, r.y, r.x, r.y + r.height);
        if (r.sides.right) line(r.x + r.width, r.y, r.x + r.width, r.y + r.height);
      }
    }
    for (const line of p.lines) {
      for (const img of line.images) {
        const embedded = await this.image(img);
        if (embedded) out.drawImage(embedded, { x: img.x, y: Y(img.y + img.height), width: img.width, height: img.height });
        else out.drawRectangle({ x: img.x, y: Y(img.y + img.height), width: img.width, height: img.height, borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 0.5 });
      }
      for (const seg of line.segments) await this.segment(out, p, line.baseline, seg);
      // Ink: each stroke as a path (SVG coordinates: from the page's top-left, y downwards).
      for (const st of line.strokes ?? []) {
        if (st.points.length < 2) continue;
        const pts = st.points.length === 2 ? [...st.points, st.points[0] + 0.01, st.points[1]] : st.points;
        let path = `M ${pts[0]} ${pts[1]}`;
        for (let i = 2; i < pts.length; i += 2) path += ` L ${pts[i]} ${pts[i + 1]}`;
        out.drawSvgPath(path, { x: 0, y: p.height, borderColor: color(st.color), borderWidth: st.width, borderLineCap: LineCapStyle.Round });
      }
      for (const r of line.rules) {
        out.drawLine({ start: { x: r.x1, y: Y(r.y1) }, end: { x: r.x2, y: Y(r.y2) }, thickness: r.width, color: rgb(0, 0, 0) });
      }
    }
  }

  private async segment(out: PDFPage, p: Page, baseline: number, seg: Segment) {
    const text = seg.field === "PAGE" ? String(p.number) : seg.field === "NUMPAGES" ? String(this.total) : seg.text;
    const size = seg.vertAlign ? seg.size * 0.65 : seg.size;
    const shift = seg.vertAlign === "superscript" ? -seg.size * 0.33 : seg.vertAlign === "subscript" ? seg.size * 0.14 : 0;
    const y = p.height - (baseline + shift + (seg.dy ?? 0));
    const pieces = this.source.split ? this.source.split(text, seg.font) : [{ text, font: seg.font }];
    let x = seg.x;
    for (const raw of pieces) {
      const font = await this.font(raw.font);
      // pdf-lib keeps the fontkit font it embeds; it tells which characters the font has.
      const face = (font as unknown as { embedder?: { font?: { hasGlyphForCodePoint?(cp: number): boolean } } }).embedder?.font;
      const piece = face?.hasGlyphForCodePoint ? { ...raw, text: symbolText(raw.text, (cp) => face.hasGlyphForCodePoint!(cp)) } : raw;
      let width: number;
      try {
        width = font.widthOfTextAtSize(piece.text, size);
      } catch {
        width = 0;
      }
      if (seg.highlight) {
        out.drawRectangle({ x, y: y - size * 0.22, width, height: size * 1.1, color: color(HIGHLIGHT[seg.highlight]) });
      }
      if (piece.text.trim()) {
        try {
          out.drawText(piece.text, { x, y, size, font, color: color(seg.color) });
        } catch {
          // A glyph the embedded font cannot encode: skip the piece rather than fail the export.
        }
      }
      if (seg.underline) out.drawLine({ start: { x, y: y - size * 0.12 }, end: { x: x + width, y: y - size * 0.12 }, thickness: size * 0.05, color: color(seg.color) });
      if (seg.strike) out.drawLine({ start: { x, y: y + size * 0.3 }, end: { x: x + width, y: y + size * 0.3 }, thickness: size * 0.05, color: color(seg.color) });
      x += width;
    }
  }
}

export async function renderPdf(doc: PrimeDocument, layout: LayoutResult, fonts: FontSource): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  pdf.setProducer("Prime Word");
  pdf.setCreator("Prime Word");
  const renderer = new Renderer(pdf, doc, fonts, layout.pages.length);
  for (const page of layout.pages) await renderer.page(page);
  return pdf.save();
}
