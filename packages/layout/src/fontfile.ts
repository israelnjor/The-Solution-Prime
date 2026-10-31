/** Font file helpers shared by Node and the browser. */

/** Copies one font out of a TrueType collection (.ttc) into a standalone font file. */
export function extractFont(buf: Uint8Array, index: number): Uint8Array {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  if (dv.getUint32(0) !== 0x74746366 /* "ttcf" */) return buf;
  const offset = dv.getUint32(12 + 4 * index);
  const count = dv.getUint16(offset + 4);
  const tables = Array.from({ length: count }, (_, i) => {
    const r = offset + 12 + 16 * i;
    return { tag: dv.getUint32(r), sum: dv.getUint32(r + 4), at: dv.getUint32(r + 8), length: dv.getUint32(r + 12) };
  });
  const head = 12 + 16 * count;
  const out = new Uint8Array(head + tables.reduce((n, t) => n + ((t.length + 3) & ~3), 0));
  const od = new DataView(out.buffer);
  out.set(buf.subarray(offset, offset + 12), 0);
  let pos = head;
  tables.forEach((t, i) => {
    od.setUint32(12 + 16 * i, t.tag);
    od.setUint32(16 + 16 * i, t.sum);
    od.setUint32(20 + 16 * i, pos);
    od.setUint32(24 + 16 * i, t.length);
    out.set(buf.subarray(t.at, t.at + t.length), pos);
    pos += (t.length + 3) & ~3;
  });
  return out;
}
