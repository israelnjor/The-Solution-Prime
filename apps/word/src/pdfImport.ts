/**
 * PDF → Word in the app: pdf.js (with its support files shipped in public/pdfjs, for offline use) and
 * our reconstruction (@prime/pdf-import). Loaded only when the converter is used.
 */
import { type LayoutResult, layoutDocument, type Measurer, type Page } from "@prime/layout";
import type { ConversionResult, RecognizedWord } from "@prime/pdf-import";
import type { PrimeDocument } from "@prime/pdm";
import { invoke } from "@tauri-apps/api/core";
import { isDesktop } from "./platform";

type PdfJs = typeof import("pdfjs-dist");
let loaded: Promise<PdfJs> | undefined;

/** pdf.js, with its worker started from the app's own files. */
export function pdfjs(): Promise<PdfJs> {
  loaded ??= Promise.all([import("pdfjs-dist"), import("pdfjs-dist/build/pdf.worker.min.mjs?url")]).then(([lib, worker]) => {
    lib.GlobalWorkerOptions.workerSrc = worker.default;
    return lib;
  });
  return loaded;
}

const files = () => {
  const base = new URL("./pdfjs/", location.href).href;
  return { cMapUrl: `${base}cmaps/`, standardFontDataUrl: `${base}standard_fonts/`, wasmUrl: `${base}wasm/`, iccUrl: `${base}iccs/` };
};

/** Pictures are kept at 200 dpi at most, at the size they are shown on the page. */
const PICTURE_DPI = 200;

/**
 * Pictures from a PDF, as small as they can be without losing what is in them: photographs (many
 * colours, nothing see-through) as JPEG, drawings and screenshots as PNG; anything sharper than
 * 200 dpi at its printed size scaled down. A 480-page book's pictures went from 68 MB to a fraction.
 */
async function encodePicture(rgba: Uint8Array, width: number, height: number, shown: { w: number; h: number }) {
  if (typeof OffscreenCanvas === "undefined") return undefined;
  // Many distinct colours (in a sample of the pixels) and no transparency: a photograph.
  let opaque = true;
  const colours = new Set<number>();
  const step = Math.max(1, Math.floor((width * height) / 4096));
  for (let p = 0; p < width * height; p += step) {
    const i = p * 4;
    if (rgba[i + 3] !== 255) opaque = false;
    colours.add(((rgba[i] >> 3) << 10) | ((rgba[i + 1] >> 3) << 5) | (rgba[i + 2] >> 3));
  }
  const photo = opaque && colours.size > 600;
  const scale = Math.min(1, ((shown.w / 72) * PICTURE_DPI) / width, ((shown.h / 72) * PICTURE_DPI) / height);
  const w = Math.max(1, Math.round(width * scale));
  const h = Math.max(1, Math.round(height * scale));
  const source = new OffscreenCanvas(width, height);
  source.getContext("2d")!.putImageData(new ImageData(new Uint8ClampedArray(rgba), width, height), 0, 0);
  let canvas = source;
  if (w !== width || h !== height) {
    canvas = new OffscreenCanvas(w, h);
    const ctx = canvas.getContext("2d")!;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(source, 0, 0, w, h);
  }
  const type = photo ? "image/jpeg" : "image/png";
  const blob = await canvas.convertToBlob(photo ? { type, quality: 0.85 } : { type });
  return { data: new Uint8Array(await blob.arrayBuffer()), contentType: type as "image/png" | "image/jpeg" };
}

/** What the system's recogniser returns: lines of words, boxes in the image's pixels. */
interface Recognized { lines: { words: { text: string; x: number; y: number; w: number; h: number }[] }[] }
/** Recognises an 8-bit grey image (one byte a pixel). */
type RecognizeImage = (gray: Uint8Array, width: number, height: number) => Promise<Recognized>;

let recognizer: Promise<{ run: RecognizeImage; limit: number } | undefined> | undefined;

/**
 * The operating system's text recognition (offline): Windows' own on Windows; none elsewhere yet.
 * Self-tests put a stand-in at globalThis.primeRecognizeText.
 */
function systemRecognizer() {
  recognizer ??= (async () => {
    const stand = (globalThis as { primeRecognizeText?: RecognizeImage }).primeRecognizeText;
    if (stand) return { run: stand, limit: 10_000 };
    if (!isDesktop) return undefined;
    const limit = await invoke<number>("recognition_limit").catch(() => 0);
    if (!limit) return undefined;
    const run: RecognizeImage = (gray, width, height) =>
      invoke<Recognized>("recognize_text", gray, { headers: { "x-width": String(width), "x-height": String(height) } });
    return { run, limit };
  })();
  return recognizer;
}

/** Text recognition resolution: 300 dpi, as scanners use (within the recogniser's largest image). */
const RECOGNITION_DPI = 300;

/**
 * Reads a page that has no text: drawn at 300 dpi and recognised (Windows reads light text on dark
 * bands as well as dark on light), each word's colour and weight taken from its pixels.
 */
export async function recognizePage(page: import("pdfjs-dist").PDFPageProxy): Promise<RecognizedWord[] | undefined> {
  const r = await systemRecognizer();
  if (!r) return undefined;
  const base = page.getViewport({ scale: 1 });
  const scale = Math.min(RECOGNITION_DPI / 72, (r.limit - 1) / Math.max(base.width, base.height));
  const vp = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(vp.width);
  canvas.height = Math.round(vp.height);
  const ctx = canvas.getContext("2d", { willReadFrequently: true })!;
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvas, canvasContext: ctx, viewport: vp } as Parameters<typeof page.render>[0]).promise;
  const { width, height } = canvas;
  const rgba = ctx.getImageData(0, 0, width, height).data;
  canvas.width = canvas.height = 0;
  const gray = new Uint8Array(width * height);
  for (let p = 0, i = 0; p < gray.length; p++, i += 4) gray[p] = (rgba[i] * 299 + rgba[i + 1] * 587 + rgba[i + 2] * 114) / 1000;
  const { lines } = await r.run(gray, width, height);
  const words: RecognizedWord[] = [];
  lines.forEach((l, line) => {
    for (const w of l.words) {
      const { color, ink } = inkOf(rgba, width, height, w);
      words.push({ text: w.text, x: w.x / scale, y: w.y / scale, w: w.w / scale, h: w.h / scale, line, color, ink });
    }
  });
  return words;
}

/**
 * A word's colour (its darkest-against-the-background pixels; the background from just around its
 * box) and its ink (the share of its box the letters cover: bold covers more).
 */
function inkOf(rgba: Uint8ClampedArray, width: number, height: number, b: { x: number; y: number; w: number; h: number }) {
  const x0 = Math.max(0, Math.floor(b.x)), y0 = Math.max(0, Math.floor(b.y));
  const x1 = Math.min(width - 1, Math.ceil(b.x + b.w)), y1 = Math.min(height - 1, Math.ceil(b.y + b.h));
  const ring: number[][] = [[], [], []];
  const add = (x: number, y: number) => {
    if (x < 0 || y < 0 || x >= width || y >= height) return;
    const i = (y * width + x) * 4;
    for (let c = 0; c < 3; c++) ring[c].push(rgba[i + c]);
  };
  for (let x = x0 - 2; x <= x1 + 2; x++) { add(x, y0 - 2); add(x, y1 + 2); }
  for (let y = y0 - 2; y <= y1 + 2; y++) { add(x0 - 2, y); add(x1 + 2, y); }
  const bg = ring.map((v) => v.sort((p, q) => p - q)[v.length >> 1] ?? 255);
  let max = 0;
  const dist = (i: number) => Math.abs(rgba[i] - bg[0]) + Math.abs(rgba[i + 1] - bg[1]) + Math.abs(rgba[i + 2] - bg[2]);
  for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) max = Math.max(max, dist((y * width + x) * 4));
  if (max < 40) return { color: "000000", ink: 0 };
  let n = 0, inked = 0;
  const sum = [0, 0, 0];
  for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) {
    const i = (y * width + x) * 4;
    const d = dist(i);
    if (d > max * 0.5) inked++;
    if (d >= max * 0.75) { n++; for (let c = 0; c < 3; c++) sum[c] += rgba[i + c]; }
  }
  const rgb = sum.map((v) => Math.round(v / Math.max(n, 1)));
  // Near-black is black (anti-aliasing lightens the darkest pixels a little).
  const black = Math.max(...rgb) < 60 && Math.max(...rgb) - Math.min(...rgb) < 24;
  const color = black ? "000000" : rgb.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
  return { color, ink: inked / ((x1 - x0 + 1) * (y1 - y0 + 1)) };
}

export async function convertPdfFile(bytes: Uint8Array, onPage?: (done: number, total: number) => void): Promise<ConversionResult> {
  const [{ convertPdf }, lib] = await Promise.all([import("@prime/pdf-import"), pdfjs()]);
  // pdf.js takes ownership of the buffer it is given: pass a copy, the caller keeps the original.
  return convertPdf(bytes.slice(), lib, { ...files(), encodePicture, recognize: recognizePage }, onPage);
}

/** The original PDF, for drawing its pages. */
export async function openPdfForView(bytes: Uint8Array) {
  const lib = await pdfjs();
  return lib.getDocument({ data: bytes.slice(), verbosity: 0, ...files() } as Parameters<PdfJs["getDocument"]>[0]).promise;
}

/** Drawings queued per canvas: pdf.js draws one page at a time on a canvas, and only the latest counts. */
const queue = new WeakMap<HTMLCanvasElement, Promise<void>>();
const latest = new WeakMap<HTMLCanvasElement, number>();

/** Draws page `n` (1-based) of a PDF at `scale` (1 = 72 dpi) on a canvas. */
export function drawPdfPage(doc: Awaited<ReturnType<typeof openPdfForView>>, n: number, canvas: HTMLCanvasElement, scale: number): Promise<void> {
  const ticket = (latest.get(canvas) ?? 0) + 1;
  latest.set(canvas, ticket);
  const run = (queue.get(canvas) ?? Promise.resolve()).catch(() => undefined).then(async () => {
    if (latest.get(canvas) !== ticket) return; // a newer drawing replaces this one
    const page = await doc.getPage(n);
    const vp = page.getViewport({ scale });
    canvas.width = Math.round(vp.width);
    canvas.height = Math.round(vp.height);
    await page.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport: vp } as Parameters<typeof page.render>[0]).promise;
  });
  queue.set(canvas, run);
  return run;
}

export function layoutOf(doc: PrimeDocument, measurer: Measurer): LayoutResult {
  return layoutDocument(doc, measurer);
}

const IMAGE_TYPES: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };
const images = new Map<string, HTMLImageElement>();
function image(doc: PrimeDocument, id?: string, part?: string): HTMLImageElement | undefined {
  const key = id ?? part ?? "";
  let img = images.get(key);
  if (img) return img.complete ? img : undefined;
  const res = id ? doc.images[id] : undefined;
  const p = !res && part ? doc.package.parts.find((x) => x.name === part) : undefined;
  const type = res?.contentType ?? IMAGE_TYPES[p?.name.split(".").pop()?.toLowerCase() ?? ""];
  const data = res?.data ?? p?.data;
  if (!type || !data) return undefined;
  img = new Image();
  img.src = `data:${type};base64,${data}`;
  images.set(key, img);
  return undefined;
}

/**
 * Draws a laid-out page of the converted document on a canvas: what Print Layout and the PDF export
 * show, at `scale` (1 = 72 dpi). Pictures still decoding are drawn on the next call.
 */
export function drawLayoutPage(doc: PrimeDocument, page: Page, canvas: HTMLCanvasElement, scale: number) {
  canvas.width = Math.round(page.width * scale);
  canvas.height = Math.round(page.height * scale);
  const ctx = canvas.getContext("2d")!;
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, page.width, page.height);
  for (const r of page.rects) {
    if (r.fill) { ctx.fillStyle = `#${r.fill}`; ctx.fillRect(r.x, r.y, r.width, r.height); }
    if (r.sides) {
      ctx.strokeStyle = `#${r.stroke?.color && r.stroke.color !== "auto" ? r.stroke.color : "000000"}`;
      ctx.lineWidth = r.stroke?.width ?? 0.5;
      ctx.beginPath();
      if (r.sides.top) { ctx.moveTo(r.x, r.y); ctx.lineTo(r.x + r.width, r.y); }
      if (r.sides.bottom) { ctx.moveTo(r.x, r.y + r.height); ctx.lineTo(r.x + r.width, r.y + r.height); }
      if (r.sides.left) { ctx.moveTo(r.x, r.y); ctx.lineTo(r.x, r.y + r.height); }
      if (r.sides.right) { ctx.moveTo(r.x + r.width, r.y); ctx.lineTo(r.x + r.width, r.y + r.height); }
      ctx.stroke();
    }
  }
  for (const line of page.lines) {
    for (const img of line.images) {
      const el = image(doc, img.imageId, img.part);
      if (el) ctx.drawImage(el, img.x, img.y, img.width, img.height);
    }
    for (const s of line.segments) {
      const size = s.vertAlign ? s.size * 0.65 : s.size;
      const lift = s.vertAlign === "superscript" ? -s.size * 0.33 : s.vertAlign === "subscript" ? s.size * 0.14 : 0;
      ctx.font = `${s.font.italic ? "italic " : ""}${s.font.bold ? "bold " : ""}${size}px "${s.font.family}"`;
      ctx.fillStyle = `#${s.color && s.color !== "auto" ? s.color : "000000"}`;
      const text = s.field === "PAGE" ? String(page.number) : s.text;
      const y = line.baseline + lift + (s.dy ?? 0);
      ctx.fillText(text, s.x, y);
      if (s.underline) ctx.fillRect(s.x, y + size * 0.12, ctx.measureText(text).width, Math.max(0.5, size * 0.05));
    }
    for (const st of line.strokes ?? []) {
      ctx.strokeStyle = `#${st.color}`;
      ctx.lineWidth = st.width;
      ctx.beginPath();
      for (let i = 0; i < st.points.length; i += 2) (i ? ctx.lineTo : ctx.moveTo).call(ctx, st.points[i], st.points[i + 1]);
      ctx.stroke();
    }
  }
}
