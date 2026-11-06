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
  /** The encoded picture (PNG or JPEG). */
  data: Uint8Array;
  contentType: "image/png" | "image/jpeg";
}

/**
 * Encodes a picture taken out of a PDF (RGBA pixels, and the size it is shown at, in pt). The app
 * passes one that writes photographs as JPEG and scales large pictures down; without one, pictures
 * are written as PNG.
 */
export type PictureEncoder = (rgba: Uint8Array, width: number, height: number, shown: { w: number; h: number }) =>
  Promise<{ data: Uint8Array; contentType: "image/png" | "image/jpeg" } | undefined>;

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
/**
 * TeX's fonts (Computer Modern, Latin Modern, AMS): their names say the design, not "Bold" or "Italic"
 * (CMSSI10 = sans italic, CMBX12 = bold): mapped to the nearest Windows fonts.
 */
function texFont(raw: string): TextPiece["font"] | undefined {
  const n = raw.toUpperCase();
  const lm = /^LM(ROMAN|SANS|MONO)/.exec(n);
  if (lm) {
    const style = n.slice(lm[0].length);
    return { family: lm[1] === "SANS" ? "Arial" : lm[1] === "MONO" ? "Consolas" : "Cambria", bold: /BOLD|DEMI/.test(style), italic: /ITALIC|OBLIQUE|SLANT/.test(style) };
  }
  if (!/^(CM|MSAM|MSBM|EUF|EUR)/.test(n)) return undefined;
  if (/^(CMSY|CMEX|CMBSY|MSAM|MSBM|EU)/.test(n)) return { family: "Cambria Math", bold: false, italic: false };
  if (/^CMSS/.test(n)) return { family: "Arial", bold: /^CMSS(BX|DC)/.test(n), italic: /^CMSS(I|SL|BXI)/.test(n) || /^CMSSBXO/.test(n) };
  if (/^CMTT|^CMITT|^CMSLTT/.test(n)) return { family: "Consolas", bold: false, italic: /^CM(I|SL)TT/.test(n) };
  return {
    family: "Cambria",
    bold: /^CM(BX|B\d|MIB|BSY)/.test(n),
    italic: /^CM(TI|SL|MI|BXTI|BXSL|U)/.test(n),
  };
}

/**
 * Text without control characters: fonts without a Unicode map (TeX's CMEX: large brackets) come out as
 * codes below 32. CMEX's brackets become ordinary brackets; anything else unreadable is dropped.
 */
function cleanText(text: string, font: string | undefined): string {
  const cmex: Record<number, string> = {
    0x00: "(", 0x01: ")", 0x02: "[", 0x03: "]", 0x08: "{", 0x09: "}", 0x0c: "|", 0x0d: "|",
    0x10: "(", 0x11: ")", 0x12: "(", 0x13: ")", 0x14: "[", 0x15: "]", 0x1a: "{", 0x1b: "}",
  };
  const isCmex = /CMEX/.test(font ?? "");
  return text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, (c) => (isCmex ? cmex[c.charCodeAt(0)] ?? "" : ""));
}

/** Letters of a blackboard-bold font (TeX's MSBM: ℝ, ℂ, ℕ…) as their Unicode characters. */
function doubleStruck(text: string): string {
  const special: Record<string, string> = { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" };
  return text.replace(/[A-Z]/g, (c) => special[c] ?? String.fromCodePoint(0x1d538 + c.charCodeAt(0) - 65));
}

export function fontOf(name: string | undefined, flags: { bold?: boolean; italic?: boolean } = {}): TextPiece["font"] {
  const raw = (name ?? "").replace(/^[A-Z]{6}\+/, "");
  const tex = texFont(raw);
  if (tex) return tex;
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
  /** How pictures are encoded (default: PNG). */
  encodePicture?: PictureEncoder;
}

/** Page content of every page. */
export async function extractPages(data: Uint8Array, pdfjs: PdfJs, files: PdfJsFiles = {},
  onPage?: (done: number, total: number) => void): Promise<PageContent[]> {
  // fontExtraProperties: keeps font names (bold and italic come from them).
  const task = pdfjs.getDocument({
    data, verbosity: 0, fontExtraProperties: true, disableFontFace: true,
    // Pictures as pixel data everywhere (in a browser pdf.js would otherwise hand over bitmaps it may
    // release before we read them): converting needs their pixels.
    isOffscreenCanvasSupported: false, isImageDecoderSupported: false,
    ...(files.cMapUrl ? { cMapUrl: files.cMapUrl, cMapPacked: true } : {}),
    ...(files.standardFontDataUrl ? { standardFontDataUrl: files.standardFontDataUrl } : {}),
    ...(files.wasmUrl ? { wasmUrl: files.wasmUrl } : {}),
    ...(files.iccUrl ? { iccUrl: files.iccUrl } : {}),
  } as Parameters<PdfJs["getDocument"]>[0]);
  const doc = await task.promise;
  const pages: PageContent[] = [];
  try {
    for (let n = 1; n <= doc.numPages; n++) {
      pages.push(await extractPage(await doc.getPage(n), n, pdfjs, files.encodePicture));
      onPage?.(n, doc.numPages);
    }
  } finally {
    await task.destroy();
  }
  return pages;
}

async function extractPage(page: import("pdfjs-dist").PDFPageProxy, number: number, pdfjs: PdfJs, encode?: PictureEncoder): Promise<PageContent> {
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
  /** Each text run as drawn: where, its colour, its characters (pdf.js's text pieces merge runs). */
  const textColor: { x: number; y: number; color: string; text: string }[] = [];
  /** Inside a soft-masked (partly transparent) group: shapes there are see-through, e.g. shadows. */
  let masked = 0;
  const maskedGroups: boolean[] = [];
  /** The colour of the text drawn last at or just before (x, y) on its baseline. */
  const colorAt = (x: number, y: number) => {
    let best: { x: number; color: string } | undefined;
    for (const t of textColor) if (Math.abs(t.y - y) < 1 && t.x <= x + 0.5 && (!best || t.x >= best.x)) best = t;
    return best?.color;
  };

  /** fillAlpha: fill opacity (0–1); mostly see-through shapes (shadows) are not backgrounds. */
  type State = { ctm: Matrix; fill: string; stroke: string; lineWidth: number; fillAlpha: number };
  let st: State = { ctm: IDENTITY, fill: "000000", stroke: "000000", lineWidth: 1, fillAlpha: 1 };
  const stack: State[] = [];
  let tm: Matrix = IDENTITY;
  // Text state, to advance the text position after each run (as pdf.js draws it).
  const text = { size: 0, fontMatrix: 0.001, charSpacing: 0, wordSpacing: 0, hScale: 1 };
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
    if (curved) {
      // A filled shape with curves (a rounded box, a banner): its bounding box, as a background area.
      if (!filled || st.fill === "FFFFFF" || st.fillAlpha < 0.5 || masked) return;
      for (const sp of subpaths) {
        const xs = sp.map((p) => p[0]);
        const ys = sp.map((p) => p[1]);
        const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
        if ((x1 - x0) * (y1 - y0) > 200 && (x1 - x0) * (y1 - y0) < vp.width * vp.height * 0.9) fills.push({ x: x0, y: y0, w: x1 - x0, h: y1 - y0, color: st.fill });
      }
      return;
    }
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
        else if (st.fill !== "FFFFFF" && st.fillAlpha >= 0.5 && !masked && w * h < vp.width * vp.height * 0.9) fills.push({ x: x0, y: y0, w, h, color: st.fill });
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
      case OPS.setFont: {
        text.size = Number(args[1]) || 0;
        const font = page.commonObjs.has(args[0] as string) ? page.commonObjs.get(args[0] as string) as { fontMatrix?: number[] } : undefined;
        text.fontMatrix = font?.fontMatrix?.[0] ?? 0.001;
        break;
      }
      case OPS.setCharSpacing: text.charSpacing = Number(args[0]) || 0; break;
      case OPS.setWordSpacing: text.wordSpacing = Number(args[0]) || 0; break;
      case OPS.setHScale: text.hScale = (Number(args[0]) || 100) / 100; break;
      case OPS.setGState:
        for (const [key, value] of (args[0] as [string, unknown][]) ?? []) if (key === "ca") st.fillAlpha = Number(value);
        break;
      case OPS.beginText: tm = IDENTITY; lineStart = IDENTITY; break;
      case OPS.setTextMatrix: tm = matrix(args); lineStart = tm; break;
      case OPS.moveText: case OPS.setLeadingMoveText:
        lineStart = mul([1, 0, 0, 1, Number(args[0]), Number(args[1])], lineStart); tm = lineStart; break;
      case OPS.showText: case OPS.showSpacedText: {
        const [x, y] = apply(toPage, ...apply(mul(tm, st.ctm), 0, 0));
        const glyphs = (args[0] ?? []) as ({ unicode?: string; width?: number; isSpace?: boolean } | number | null)[];
        textColor.push({ x, y, color: st.fill, text: glyphs.map((g) => (g && typeof g === "object" ? g.unicode ?? "" : "")).join("") });
        // Advance past the run, so the next run on the line starts where this one ended.
        let advance = 0;
        for (const g of glyphs) {
          if (typeof g === "number") advance -= (g * text.size) / 1000;
          else if (g) advance += (g.width ?? 0) * text.size * text.fontMatrix + text.charSpacing + (g.isSpace ? text.wordSpacing : 0);
        }
        tm = mul([1, 0, 0, 1, advance * text.hScale, 0], tm);
        break;
      }
      case OPS.beginGroup: {
        const soft = !!(args[0] as { smask?: unknown } | undefined)?.smask;
        maskedGroups.push(soft);
        if (soft) masked++;
        break;
      }
      case OPS.endGroup: {
        if (maskedGroups.pop()) masked--;
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
        const pixels = imageRgba(img as ImageLike);
        if (pixels) {
          const encoded = (encode && await encode(pixels.rgba, pixels.width, pixels.height, { w: box.w, h: box.h }))
            ?? { data: encodePng(pixels.rgba, pixels.width, pixels.height), contentType: "image/png" as const };
          pictures.push({ ...box, ...encoded });
        }
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
    const piece: TextPiece = {
      text: cleanText(/MSBM/.test(obj?.name ?? "") ? doubleStruck(item.str) : item.str, obj?.name), x, y, width: item.width, size: Math.round(size * 100) / 100,
      font: fontOf(obj?.name, obj ?? {}),
      color: colorAt(x, y) ?? "000000",
    };
    texts.push(...splitByColour(piece));
  }
  /**
   * A piece that pdf.js merged from runs of different colours ("is a" black, "vector space" red): split
   * where the drawn runs change colour, using the runs' own characters and positions.
   */
  function splitByColour(p: TextPiece): TextPiece[] {
    // The runs on its baseline from the one in progress where the piece starts to its end.
    const line = textColor.filter((t) => Math.abs(t.y - p.y) < 1 && t.text).sort((a, b) => a.x - b.x);
    let from = 0;
    line.forEach((t, k) => { if (t.x <= p.x + 0.5) from = k; });
    const runs = line.slice(from).filter((t, k) => k === 0 || t.x < p.x + p.width - 0.5);
    if (new Set(runs.map((r) => r.color)).size < 2) return [p];
    const out: TextPiece[] = [];
    let at = 0;
    runs.forEach((r, k) => {
      // The run's characters, found in the piece's text from where the previous run ended.
      const want = r.text.replace(/\s+/g, "");
      let end = at;
      let got = 0;
      while (end < p.text.length && got < want.length) { if (!/\s/.test(p.text[end])) got++; end++; }
      if (k === runs.length - 1) end = p.text.length;
      const text = p.text.slice(at, end);
      const x0 = Math.max(r.x, p.x);
      const x1 = k + 1 < runs.length ? runs[k + 1].x : p.x + p.width;
      if (text) {
        const last = out[out.length - 1];
        if (last && last.color === r.color) { last.text += text; last.width = x1 - last.x; }
        else out.push({ ...p, text, x: x0, width: x1 - x0, color: r.color });
      }
      at = end;
    });
    return out.length ? out : [p];
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
  /** In a browser, pdf.js hands over decoded pictures as bitmaps instead of pixel data. */
  bitmap?: CanvasImageSource & { width: number; height: number };
}

/** pdf.js decoded image data (1-bit gray, RGB or RGBA) as RGBA pixels. */
function imageRgba(img: ImageLike | undefined): { rgba: Uint8Array; width: number; height: number } | undefined {
  if (img?.bitmap && !img.data && typeof OffscreenCanvas !== "undefined") {
    // A bitmap (pdf.js in a browser): its pixels, read back through a canvas.
    const { width, height } = img.bitmap;
    if (!width || !height) return undefined;
    const canvas = new OffscreenCanvas(width, height);
    const ctx = canvas.getContext("2d");
    if (!ctx) return undefined;
    ctx.drawImage(img.bitmap, 0, 0);
    const pixels = ctx.getImageData(0, 0, width, height).data;
    return { rgba: new Uint8Array(pixels.buffer, pixels.byteOffset, pixels.byteLength), width, height };
  }
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
  return { rgba, width, height };
}
