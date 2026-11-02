/**
 * PDF → Word in the app: pdf.js (with its support files shipped in public/pdfjs, for offline use) and
 * our reconstruction (@prime/pdf-import). Loaded only when the converter is used.
 */
import { type LayoutResult, layoutDocument, type Measurer, type Page } from "@prime/layout";
import type { ConversionResult } from "@prime/pdf-import";
import type { PrimeDocument } from "@prime/pdm";

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

export async function convertPdfFile(bytes: Uint8Array, onPage?: (done: number, total: number) => void): Promise<ConversionResult> {
  const [{ convertPdf }, lib] = await Promise.all([import("@prime/pdf-import"), pdfjs()]);
  // pdf.js takes ownership of the buffer it is given: pass a copy, the caller keeps the original.
  return convertPdf(bytes.slice(), lib, files(), onPage);
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
