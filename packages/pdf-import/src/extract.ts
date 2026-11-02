/**
 * Reading a PDF (ADR-009): pdf.js parses it; this turns its output into plain page content — text
 * pieces with position, size, font and colour; ruled lines and filled rectangles (table borders and
 * shading); pictures as PNG. Coordinates are points from the page's top-left corner, y downwards.
 */
import { encodePng } from "./png.js";

/** The pdf.js module (the Node "legacy" build or the browser build, loaded by the caller). */
export type PdfJs = typeof import("pdfjs-dist");

export interface TextPiece {
  text: string;
  /** Left edge and baseline, pt. */
  x: number;
  y: number;
  width: number;
  /** Font size in pt. */
  size: number;
  font: { family: string; bold: boolean; italic: boolean };
  /** RRGGBB. */
  color: string;
  /** A rule drawn under it (underlining). */
  underline?: boolean;
  /** The web address of a link area it lies in. */
  link?: string;
}

export interface Rule {
  /** A horizontal or vertical stroke (or a thin filled rectangle): from (x1, y1) to (x2, y2). */
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
  color: string;
}

export interface Fill {
  x: number;
  y: number;
  w: number;
  h: number;
  color: string;
}

export interface Picture {
  x: number;
  y: number;
  w: number;
  h: number;
  png: Uint8Array;
}

export interface PageContent {
  number: number;
  width: number;
  height: number;
  texts: TextPiece[];
  rules: Rule[];
  fills: Fill[];
  pictures: Picture[];
}

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
/** m1 then m2 (PDF: new CTM = m × CTM). */
const mul = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
];
/** A matrix argument: six numbers, or (newer pdf.js) one array of six. */
const matrix = (args: unknown[]): Matrix =>
  (args.length === 1 && typeof args[0] === "object" && args[0] ? Array.from(args[0] as ArrayLike<number>) : args) as Matrix;
const apply = (m: Matrix, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

/** "#rrggbb", [r, g, b] (0–255) or a gray level, as RRGGBB. */
function hex(args: unknown[]): string {
  const a = args[0];
  if (typeof a === "string") return a.replace("#", "").toUpperCase().padStart(6, "0").slice(-6);
  const [r, g, b] = (args.length >= 3 ? args : [a, a, a]) as number[];
  const h = (v: number) => Math.max(0, Math.min(255, Math.round(v > 1 ? v : v * 255))).toString(16).padStart(2, "0");
  return (h(r) + h(g) + h(b)).toUpperCase();
}

/** "BCDEEE+TimesNewRomanPS-BoldItalicMT" → Times New Roman, bold, italic. */
export function fontOf(name: string | undefined, flags: { bold?: boolean; italic?: boolean } = {}): TextPiece["font"] {
  const raw = (name ?? "").replace(/^[A-Z]{6}\+/, "");
  const [base, style = ""] = raw.split(/[-,]/, 2);
  // Weights and widths Word lists as fonts of their own: "Calibri Light", "Segoe UI Semibold".
  const variant = /Semilight|SemiLight|Semibold|SemiBold|Light|Condensed|Narrow|Black/.exec(style)?.[0];
  const family = ((base || "Calibri")
    .replace(/(PS)?MT$|PSMT$|PS$/, "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\s+/g, " ")
    .trim() + (variant ? ` ${variant.replace(/^Semi(.)/, (_, c: string) => `Semi${c.toLowerCase()}`)}` : ""));
  const s = `${style} ${raw}`.toLowerCase();
  return {
    family,
    bold: !!flags.bold || /bold|black|heavy/.test(style.toLowerCase()) || /,bold/.test(s),
    italic: !!flags.italic || /italic|oblique/.test(style.toLowerCase()),
  };
}

/** Where pdf.js finds its support files (shipped with the app, for offline use). */
export interface PdfJsFiles {
  cMapUrl?: string;
  standardFontDataUrl?: string;
  wasmUrl?: string;
  iccUrl?: string;
}

/** Page content of every page. */
export async function extractPages(data: Uint8Array, pdfjs: PdfJs, files: PdfJsFiles = {},
  onPage?: (done: number, total: number) => void): Promise<PageContent[]> {
  // fontExtraProperties: keeps font names (bold and italic come from them).
  const task = pdfjs.getDocument({
    data, verbosity: 0, fontExtraProperties: true, disableFontFace: true,
    ...(files.cMapUrl ? { cMapUrl: files.cMapUrl, cMapPacked: true } : {}),
    ...(files.standardFontDataUrl ? { standardFontDataUrl: files.standardFontDataUrl } : {}),
    ...(files.wasmUrl ? { wasmUrl: files.wasmUrl } : {}),
    ...(files.iccUrl ? { iccUrl: files.iccUrl } : {}),
  } as Parameters<PdfJs["getDocument"]>[0]);
  const doc = await task.promise;
  const pages: PageContent[] = [];
  try {
    for (let n = 1; n <= doc.numPages; n++) {
      pages.push(await extractPage(await doc.getPage(n), n, pdfjs));
      onPage?.(n, doc.numPages);
    }
  } finally {
    await task.destroy();
  }
  return pages;
}

async function extractPage(page: import("pdfjs-dist").PDFPageProxy, number: number, pdfjs: PdfJs): Promise<PageContent> {
  const vp = page.getViewport({ scale: 1 });
  const H = vp.height;
  // The viewport maps PDF space (y up, page box origin) to ours (y down, top-left).
  const toPage = vp.transform as Matrix;
  const ops = await page.getOperatorList();
  const OPS = pdfjs.OPS;
  const rules: Rule[] = [];
  const fills: Fill[] = [];
  const pictures: Picture[] = [];
  /** Fill colour at each text position (x, y rounded): textContent has no colours. */
  const textColor = new Map<string, string>();

  type State = { ctm: Matrix; fill: string; stroke: string; lineWidth: number };
  let st: State = { ctm: IDENTITY, fill: "000000", stroke: "000000", lineWidth: 1 };
  const stack: State[] = [];
  let tm: Matrix = IDENTITY;
  let lineStart: Matrix = IDENTITY;
  const P = (x: number, y: number) => apply(toPage, ...apply(st.ctm, x, y));

  const paintPath = (paint: number, data: ArrayLike<number>) => {
    const filled = [OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke].includes(paint);
    const stroked = [OPS.stroke, OPS.closeStroke, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke].includes(paint);
    if (!filled && !stroked) return;
    // Split into subpaths of straight segments (curves are not table borders: skipped).
    const subpaths: [number, number][][] = [];
    let cur: [number, number][] = [];
    let curved = false;
    for (let i = 0; i < data.length;) {
      const cmd = data[i++];
      if (cmd === 0) { if (cur.length) subpaths.push(cur); cur = [P(data[i], data[i + 1])]; i += 2; }
      else if (cmd === 1) { cur.push(P(data[i], data[i + 1])); i += 2; }
      else if (cmd === 2) { curved = true; cur.push(P(data[i + 4], data[i + 5])); i += 6; }
      else if (cmd === 3) { curved = true; cur.push(P(data[i + 2], data[i + 3])); i += 4; }
      else if (cmd === 4) { if (cur.length) cur.push(cur[0]); }
      else break;
    }
    if (cur.length) subpaths.push(cur);
    if (curved) return;
    for (const sp of subpaths) {
      const xs = sp.map((p) => p[0]);
      const ys = sp.map((p) => p[1]);
      const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
      const axisAligned = sp.every((p, k) => k === 0 || Math.abs(p[0] - sp[k - 1][0]) < 0.01 || Math.abs(p[1] - sp[k - 1][1]) < 0.01);
      if (!axisAligned) continue;
      const scale = Math.hypot(st.ctm[0], st.ctm[1]) || 1;
      if (filled && sp.length >= 4) {
        const w = x1 - x0;
        const h = y1 - y0;
        // Thin filled rectangles are how Word draws borders; larger ones are shading.
        if (h <= 2.5 && w > h) rules.push({ x1: x0, y1: (y0 + y1) / 2, x2: x1, y2: (y0 + y1) / 2, width: Math.max(h, 0.25), color: st.fill });
        else if (w <= 2.5 && h > w) rules.push({ x1: (x0 + x1) / 2, y1: y0, x2: (x0 + x1) / 2, y2: y1, width: Math.max(w, 0.25), color: st.fill });
        else if (st.fill !== "FFFFFF" && w * h < vp.width * vp.height * 0.9) fills.push({ x: x0, y: y0, w, h, color: st.fill });
      }
      if (stroked) {
        for (let k = 1; k < sp.length; k++) {
          const [a, b] = [sp[k - 1], sp[k]];
          if (Math.hypot(b[0] - a[0], b[1] - a[1]) < 0.5) continue;
          rules.push({ x1: Math.min(a[0], b[0]), y1: Math.min(a[1], b[1]), x2: Math.max(a[0], b[0]), y2: Math.max(a[1], b[1]), width: st.lineWidth * scale, color: st.stroke });
        }
      }
    }
  };

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i];
    const args = ops.argsArray[i] as unknown[];
    switch (fn) {
      case OPS.save: stack.push({ ...st }); break;
      case OPS.restore: st = stack.pop() ?? st; break;
      case OPS.transform: st.ctm = mul(matrix(args), st.ctm); break;
      case OPS.setFillRGBColor: case OPS.setFillGray: case OPS.setFillColor: st.fill = hex(args); break;
      case OPS.setStrokeRGBColor: case OPS.setStrokeGray: case OPS.setStrokeColor: st.stroke = hex(args); break;
      case OPS.setLineWidth: st.lineWidth = Number(args[0]) || 0.75; break;
      case OPS.beginText: tm = IDENTITY; lineStart = IDENTITY; break;
      case OPS.setTextMatrix: tm = matrix(args); lineStart = tm; break;
      case OPS.moveText: case OPS.setLeadingMoveText:
        lineStart = mul([1, 0, 0, 1, Number(args[0]), Number(args[1])], lineStart); tm = lineStart; break;
      case OPS.showText: case OPS.showSpacedText: {
        const [x, y] = apply(toPage, ...apply(mul(tm, st.ctm), 0, 0));
        textColor.set(`${Math.round(x)},${Math.round(y)}`, st.fill);
        break;
      }
      case OPS.constructPath: {
        const [paint, path] = args as [number, ArrayLike<number>[] | ArrayLike<number>];
        const data = Array.isArray(path) && path.length && typeof path[0] !== "number" ? (path as ArrayLike<number>[])[0] : (path as ArrayLike<number>);
        if (data) paintPath(paint, data);
        break;
      }
      case OPS.paintImageXObject: case OPS.paintInlineImageXObject: {
        const [x0, y0] = P(0, 0);
        const [x1, y1] = P(1, 1);
        const box = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) };
        if (box.w < 2 || box.h < 2) break;
        const img = fn === OPS.paintInlineImageXObject ? args[0] : await new Promise<unknown>((resolve) => {
          const id = args[0] as string;
          const objs = id.startsWith("g_") ? page.commonObjs : page.objs;
          objs.get(id, resolve);
        });
        const png = imagePng(img as ImageLike);
        if (png) pictures.push({ ...box, png });
        break;
      }
      default:
    }
  }

  const tc = await page.getTextContent();
  const texts: TextPiece[] = [];
  for (const item of tc.items) {
    if (!("str" in item) || !item.str) continue;
    const m = mul(item.transform as Matrix, toPage);
    const size = Math.hypot(item.transform[2], item.transform[3]);
    if (!size) continue;
    const obj = page.commonObjs.has(item.fontName) ? page.commonObjs.get(item.fontName) as { name?: string; bold?: boolean; italic?: boolean } : undefined;
    const x = m[4];
    const y = m[5];
    texts.push({
      text: item.str, x, y, width: item.width, size: Math.round(size * 100) / 100,
      font: fontOf(obj?.name, obj ?? {}),
      color: textColor.get(`${Math.round(x)},${Math.round(y)}`) ?? "000000",
    });
  }
  // Underlining: a thin horizontal rule just below a piece's baseline, under its text. Such rules are
  // part of the text, not table borders.
  // Lines drawn in pieces (Word draws an underline word by word): touching pieces on one line become one.
  const merged: Rule[] = [];
  for (const r of [...rules].sort((a, b) => a.y1 - b.y1 || a.x1 - b.x1)) {
    const last = merged[merged.length - 1];
    // Real horizontal lines only (not the tiny ticks Word draws at table corners).
    const horizontal = (x: Rule) => Math.abs(x.y2 - x.y1) < 0.5 && x.x2 - x.x1 > 1;
    if (last && horizontal(r) && horizontal(last) && Math.abs(last.y1 - r.y1) < 0.3 &&
        Math.abs(last.width - r.width) < 0.2 && last.color === r.color && r.x1 <= last.x2 + 1.5) {
      last.x2 = Math.max(last.x2, r.x2);
    } else merged.push({ ...r });
  }
  rules.length = 0;
  rules.push(...merged);
  const underlines = new Set<Rule>();
  const verticals = rules.filter((v) => Math.abs(v.x2 - v.x1) < 1);
  const joined = (u: Rule) => verticals.some((v) => (Math.abs(v.x1 - u.x1) < 2 || Math.abs(v.x1 - u.x2) < 2) && v.y1 <= u.y1 + 2 && v.y2 >= u.y1 - 2);
  for (const t of texts) {
    if (!t.text.trim()) continue;
    // Under the text (it may run on under the next words), and not part of a ruled grid: no vertical
    // rule meets its ends, as one would at a table border.
    const r = rules.find((u) => Math.abs(u.y1 - u.y2) < 0.5 && u.width <= Math.max(1.5, t.size * 0.12) &&
      u.y1 > t.y && u.y1 - t.y < t.size * 0.3 &&
      Math.min(u.x2, t.x + t.width) - Math.max(u.x1, t.x) >= t.width * 0.6 && !joined(u));
    if (r) { t.underline = true; underlines.add(r); }
  }
  // Links: the page's link areas (web addresses only; links within the document are not kept).
  for (const a of await page.getAnnotations({ intent: "display" })) {
    const url = (a as { subtype?: string; url?: string; unsafeUrl?: string }).url ?? (a as { unsafeUrl?: string }).unsafeUrl;
    if ((a as { subtype?: string }).subtype !== "Link" || !url) continue;
    const [x1, y1, x2, y2] = (a as { rect: number[] }).rect;
    const [ax, ay] = apply(toPage, x1, y1);
    const [bx, by] = apply(toPage, x2, y2);
    const box = { x0: Math.min(ax, bx), x1: Math.max(ax, bx), y0: Math.min(ay, by), y1: Math.max(ay, by) };
    for (const t of texts) {
      const cx = t.x + t.width / 2;
      if (cx >= box.x0 && cx <= box.x1 && t.y >= box.y0 - 1 && t.y - t.size * 0.3 <= box.y1 + 1) t.link = url;
    }
  }
  return { number, width: vp.width, height: H, texts, rules: rules.filter((r) => !underlines.has(r)), fills, pictures };
}

interface ImageLike {
  width: number;
  height: number;
  kind?: number;
  data?: Uint8ClampedArray | Uint8Array;
}

/** pdf.js decoded image data (1-bit gray, RGB or RGBA) as PNG. */
function imagePng(img: ImageLike | undefined): Uint8Array | undefined {
  if (!img?.data || !img.width || !img.height) return undefined;
  const { width, height, kind, data } = img;
  const rgba = new Uint8Array(width * height * 4);
  if (kind === 1) {
    // 1 bit per pixel, rows padded to whole bytes; 1 = white.
    const row = (width + 7) >> 3;
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      const v = (data[y * row + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0;
      rgba.set([v, v, v, 255], (y * width + x) * 4);
    }
  } else if (kind === 2 || data.length === width * height * 3) {
    for (let p = 0, q = 0; p < width * height; p++, q += 3) rgba.set([data[q], data[q + 1], data[q + 2], 255], p * 4);
  } else if (data.length === width * height * 4) {
    rgba.set(data);
  } else return undefined;
  return encodePng(rgba, width, height);
}
