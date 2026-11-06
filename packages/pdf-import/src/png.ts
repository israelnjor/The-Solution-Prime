/**
 * A small PNG encoder for pictures taken out of PDFs: greyscale, RGB or RGBA (the smallest that holds
 * the picture), each row with the PNG filter that suits it best (the usual "minimum sum" choice).
 */
import { zlibSync } from "fflate";

const CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

const paeth = (a: number, b: number, c: number) => {
  const p = a + b - c;
  const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
};

export function encodePng(rgba: Uint8Array, width: number, height: number): Uint8Array {
  // The fewest channels that hold the picture: grey, RGB, or RGBA when anything is see-through.
  let opaque = true;
  let grey = true;
  for (let i = 0; i < rgba.length && (opaque || grey); i += 4) {
    if (rgba[i + 3] !== 255) opaque = false;
    if (rgba[i] !== rgba[i + 1] || rgba[i] !== rgba[i + 2]) grey = false;
  }
  const channels = !opaque ? 4 : grey ? 1 : 3;
  const colourType = channels === 4 ? 6 : channels === 3 ? 2 : 0;
  const stride = width * channels;
  const pixels = new Uint8Array(stride * height);
  for (let p = 0, q = 0; p < width * height; p++) {
    const i = p * 4;
    if (channels === 1) pixels[q++] = rgba[i];
    else { pixels[q++] = rgba[i]; pixels[q++] = rgba[i + 1]; pixels[q++] = rgba[i + 2]; if (channels === 4) pixels[q++] = rgba[i + 3]; }
  }
  // Each row filtered (None, Sub, Up, Average, Paeth): the one with the smallest sum of values.
  const raw = new Uint8Array((stride + 1) * height);
  const candidate = new Uint8Array(stride);
  const best = new Uint8Array(stride);
  for (let y = 0; y < height; y++) {
    const row = pixels.subarray(y * stride, (y + 1) * stride);
    const up = y > 0 ? pixels.subarray((y - 1) * stride, y * stride) : undefined;
    let bestSum = Infinity;
    let bestType = 0;
    for (let type = 0; type < 5; type++) {
      let sum = 0;
      for (let x = 0; x < stride; x++) {
        const a = x >= channels ? row[x - channels] : 0;
        const b = up ? up[x] : 0;
        const c = up && x >= channels ? up[x - channels] : 0;
        const v = type === 0 ? row[x] : type === 1 ? row[x] - a : type === 2 ? row[x] - b : type === 3 ? row[x] - ((a + b) >> 1) : row[x] - paeth(a, b, c);
        candidate[x] = v & 255;
        sum += candidate[x] < 128 ? candidate[x] : 256 - candidate[x];
      }
      if (sum < bestSum) { bestSum = sum; bestType = type; best.set(candidate); }
    }
    raw[y * (stride + 1)] = bestType;
    raw.set(best, y * (stride + 1) + 1);
  }
  const header = new Uint8Array(13);
  const v = new DataView(header.buffer);
  v.setUint32(0, width);
  v.setUint32(4, height);
  header.set([8, colourType, 0, 0, 0], 8); // 8 bits per channel, deflate, adaptive filtering, no interlace
  const parts = [Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", header), chunk("IDAT", zlibSync(raw, { level: 6 })), chunk("IEND", new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}
