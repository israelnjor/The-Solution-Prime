/**
 * Turning a photo or scan of a handwritten signature into a clean e-signature: a transparent PNG of the
 * ink alone. All on this computer, with plain canvas pixel work:
 *  1. even out the lighting (a phone photo is darker at one side, with shadows): each pixel is compared
 *     with a heavily blurred copy of the page around it, so the paper becomes uniformly white;
 *  2. find the ink: a threshold between paper and ink (Otsu's method), adjustable;
 *  3. keep the ink, with soft edges, as the only opaque pixels; optionally recolour it;
 *  4. crop to the ink, with a small margin.
 */

export type InkColour = "original" | "black" | "blue";

export interface Cleaned {
  canvas: HTMLCanvasElement;
  /** Share of the kept area that is ink (a sanity check: a blank photo has almost none). */
  coverage: number;
}

const MAX_SIDE = 1600;

/** A picture (photo, scan) as a canvas, scaled down to at most MAX_SIDE pixels and turned `turns` quarter turns. */
export function toCanvas(img: CanvasImageSource & { width: number; height: number }, turns = 0): HTMLCanvasElement {
  const k = Math.min(1, MAX_SIDE / Math.max(img.width, img.height));
  const w = Math.round(img.width * k);
  const h = Math.round(img.height * k);
  const c = document.createElement("canvas");
  const side = turns % 2 === 1;
  c.width = side ? h : w;
  c.height = side ? w : h;
  const ctx = c.getContext("2d")!;
  ctx.translate(c.width / 2, c.height / 2);
  ctx.rotate((turns * Math.PI) / 2);
  ctx.drawImage(img, -w / 2, -h / 2, w, h);
  return c;
}

/** Luminance (0–255) of each pixel. */
function luminance(d: Uint8ClampedArray, n: number): Float32Array {
  const out = new Float32Array(n);
  for (let i = 0; i < n; i++) out[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2];
  return out;
}

/** The page's lighting: a large blur (box blur, three passes ≈ Gaussian), from a small copy for speed. */
function background(src: HTMLCanvasElement): Float32Array {
  const scale = 8;
  const sw = Math.max(1, Math.round(src.width / scale));
  const sh = Math.max(1, Math.round(src.height / scale));
  const small = document.createElement("canvas");
  small.width = sw;
  small.height = sh;
  const sctx = small.getContext("2d")!;
  sctx.drawImage(src, 0, 0, sw, sh);
  let lum = luminance(sctx.getImageData(0, 0, sw, sh).data, sw * sh);
  // Ink is darker than paper: lift dark spots first (a max filter), so strokes do not darken the background.
  const maxed = new Float32Array(lum.length);
  for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
    let m = 0;
    for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) {
      const yy = Math.min(sh - 1, Math.max(0, y + dy)), xx = Math.min(sw - 1, Math.max(0, x + dx));
      m = Math.max(m, lum[yy * sw + xx]);
    }
    maxed[y * sw + x] = m;
  }
  lum = maxed;
  const r = Math.max(2, Math.round(Math.min(sw, sh) / 12));
  for (let pass = 0; pass < 3; pass++) {
    const tmp = new Float32Array(lum.length);
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      let s = 0, n = 0;
      for (let k = -r; k <= r; k++) { const xx = x + k; if (xx >= 0 && xx < sw) { s += lum[y * sw + xx]; n++; } }
      tmp[y * sw + x] = s / n;
    }
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      let s = 0, n = 0;
      for (let k = -r; k <= r; k++) { const yy = y + k; if (yy >= 0 && yy < sh) { s += tmp[yy * sw + x]; n++; } }
      lum[y * sw + x] = s / n;
    }
  }
  // Back to full size (nearest small pixel; the lighting changes slowly).
  const out = new Float32Array(src.width * src.height);
  for (let y = 0; y < src.height; y++) {
    const sy = Math.min(sh - 1, Math.floor(y / scale));
    for (let x = 0; x < src.width; x++) out[y * src.width + x] = lum[sy * sw + Math.min(sw - 1, Math.floor(x / scale))];
  }
  return out;
}

/** Otsu's threshold on values 0–255: the split that best separates two groups (paper and ink). */
function otsu(values: Float32Array): number {
  const hist = new Array(256).fill(0);
  for (const v of values) hist[Math.max(0, Math.min(255, Math.round(v)))]++;
  const total = values.length;
  let sum = 0;
  for (let i = 0; i < 256; i++) sum += i * hist[i];
  let sumB = 0, wB = 0, best = 0, threshold = 128;
  for (let t = 0; t < 256; t++) {
    wB += hist[t];
    if (!wB) continue;
    const wF = total - wB;
    if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sum - sumB) / wF;
    const between = wB * wF * (mB - mF) ** 2;
    if (between > best) { best = between; threshold = t; }
  }
  return threshold;
}

/**
 * The ink of a signature photo on a transparent background. `sensitivity` (−1…1) moves the threshold:
 * higher keeps fainter strokes (and more noise).
 */
export function cleanSignature(src: HTMLCanvasElement, opts: { sensitivity: number; colour: InkColour }): Cleaned {
  const w = src.width, h = src.height;
  const ctx = src.getContext("2d")!;
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  const lum = luminance(d, w * h);
  const bg = background(src);
  // 1. Lighting evened out: paper ≈ 255 everywhere, ink stays dark.
  const flat = new Float32Array(w * h);
  for (let i = 0; i < flat.length; i++) flat[i] = Math.min(255, (lum[i] / Math.max(1, bg[i])) * 255);
  // 2. Threshold between ink and paper (Otsu on the evened-out image), moved by the sensitivity.
  const t = Math.max(40, Math.min(250, otsu(flat) + opts.sensitivity * 40));
  const soft = 18; // a soft edge either side of the threshold keeps strokes smooth
  // 3. Ink only, with its colour; 4. its bounding box.
  const out = new ImageData(w, h);
  const o = out.data;
  let x0 = w, y0 = h, x1 = -1, y1 = -1, ink = 0;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = y * w + x;
    const a = Math.max(0, Math.min(1, (t + soft - flat[i]) / (2 * soft)));
    if (a <= 0.02) continue;
    const [r, g, b] = opts.colour === "black" ? [20, 20, 25] : opts.colour === "blue" ? [16, 45, 140]
      : [d[i * 4], d[i * 4 + 1], d[i * 4 + 2]].map((c) => Math.min(c, 160)); // original ink, a little darker
    o.set([r, g, b, Math.round(a * 255)], i * 4);
    if (a > 0.5) { ink++; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
  }
  const full = document.createElement("canvas");
  full.width = w;
  full.height = h;
  full.getContext("2d")!.putImageData(out, 0, 0);
  if (x1 < 0) return { canvas: full, coverage: 0 };
  const pad = Math.round(Math.max(x1 - x0, y1 - y0) * 0.04) + 4;
  const cx = Math.max(0, x0 - pad), cy = Math.max(0, y0 - pad);
  const cw = Math.min(w, x1 + pad) - cx, ch = Math.min(h, y1 + pad) - cy;
  const cropped = document.createElement("canvas");
  cropped.width = cw;
  cropped.height = ch;
  cropped.getContext("2d")!.drawImage(full, cx, cy, cw, ch, 0, 0, cw, ch);
  return { canvas: cropped, coverage: ink / (cw * ch) };
}

/** A canvas cropped to its non-transparent pixels (for drawn and typed signatures). */
export function cropToInk(src: HTMLCanvasElement): HTMLCanvasElement {
  const { width: w, height: h } = src;
  const d = src.getContext("2d")!.getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = -1, y1 = -1;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    if (d[(y * w + x) * 4 + 3] > 10) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
  }
  if (x1 < 0) return src;
  const pad = 6;
  const out = document.createElement("canvas");
  out.width = Math.min(w, x1 + pad) - Math.max(0, x0 - pad);
  out.height = Math.min(h, y1 + pad) - Math.max(0, y0 - pad);
  out.getContext("2d")!.drawImage(src, Math.max(0, x0 - pad), Math.max(0, y0 - pad), out.width, out.height, 0, 0, out.width, out.height);
  return out;
}

/** A canvas as PNG data (base64) with its pixel size. */
export function pngOf(c: HTMLCanvasElement): { data: string; width: number; height: number } {
  return { data: c.toDataURL("image/png").split(",")[1], width: c.width, height: c.height };
}
