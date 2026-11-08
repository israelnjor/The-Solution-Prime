/**
 * Compound File Binary (MS-CFB): the container of Word 97–2003 documents (.doc) and of password-
 * protected .docx files (an encrypted package inside one). Read: every stream, by path. Written:
 * streams (in storages) as a version 3 file (512-byte sectors), small streams in the mini stream.
 */

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;

export function isCfb(bytes: Uint8Array): boolean {
  return bytes.length >= 512 && SIGNATURE.every((b, i) => bytes[i] === b);
}

/** All streams of a compound file, by path ("WordDocument", "1Table", "\u0006DataSpaces/Version"…). */
export function readCfb(bytes: Uint8Array): Map<string, Uint8Array> {
  if (!isCfb(bytes)) throw new Error("Not a compound file.");
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (o: number) => dv.getUint32(o, true);
  const sectorSize = 1 << dv.getUint16(0x1e, true);
  const miniSize = 1 << dv.getUint16(0x20, true);
  const firstDir = u32(0x30);
  const cutoff = u32(0x38);
  const firstMiniFat = u32(0x3c);
  const firstDifat = u32(0x44);
  const sectorAt = (n: number) => (n + 1) * sectorSize;
  // The FAT, from the DIFAT (109 entries in the header, then a chain of DIFAT sectors).
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) { const s = u32(0x4c + i * 4); if (s !== FREE) fatSectors.push(s); }
  for (let d = firstDifat, guard = 0; d !== END && d !== FREE && guard < 100000; guard++) {
    const o = sectorAt(d);
    const per = sectorSize / 4 - 1;
    for (let i = 0; i < per; i++) { const s = u32(o + i * 4); if (s !== FREE) fatSectors.push(s); }
    d = u32(o + per * 4);
  }
  const fat: number[] = [];
  for (const s of fatSectors) { const o = sectorAt(s); for (let i = 0; i < sectorSize / 4 && o + i * 4 + 4 <= bytes.length; i++) fat.push(u32(o + i * 4)); }
  const chain = (start: number, table: number[]): number[] => {
    const out: number[] = [];
    for (let s = start; s !== END && s !== FREE && s < table.length && out.length <= table.length; s = table[s]) out.push(s);
    return out;
  };
  const readChain = (start: number, size?: number): Uint8Array => {
    const sectors = chain(start, fat);
    const out = new Uint8Array(sectors.length * sectorSize);
    sectors.forEach((s, i) => out.set(bytes.subarray(sectorAt(s), Math.min(bytes.length, sectorAt(s) + sectorSize)), i * sectorSize));
    return size === undefined ? out : out.subarray(0, size);
  };
  const dir = readChain(firstDir);
  const ddv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  interface Entry { name: string; type: number; left: number; right: number; child: number; start: number; size: number }
  const entries: Entry[] = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const len = ddv.getUint16(o + 64, true);
    let name = "";
    for (let i = 0; i < Math.max(0, len / 2 - 1); i++) name += String.fromCharCode(ddv.getUint16(o + i * 2, true));
    entries.push({ name, type: dir[o + 66], left: ddv.getUint32(o + 68, true), right: ddv.getUint32(o + 72, true), child: ddv.getUint32(o + 76, true), start: ddv.getUint32(o + 116, true), size: ddv.getUint32(o + 120, true) });
  }
  const root = entries[0];
  const miniStream = readChain(root.start, root.size);
  const miniFat: number[] = [];
  if (firstMiniFat !== END) { const mf = readChain(firstMiniFat); const mdv = new DataView(mf.buffer, mf.byteOffset, mf.byteLength); for (let i = 0; i + 4 <= mf.length; i += 4) miniFat.push(mdv.getUint32(i, true)); }
  const readMini = (start: number, size: number) => {
    const sectors = chain(start, miniFat);
    const out = new Uint8Array(sectors.length * miniSize);
    sectors.forEach((s, i) => out.set(miniStream.subarray(s * miniSize, s * miniSize + miniSize), i * miniSize));
    return out.subarray(0, size);
  };
  const streams = new Map<string, Uint8Array>();
  const visit = (id: number, prefix: string, seen: Set<number>) => {
    if (id === NOSTREAM || id >= entries.length || seen.has(id)) return;
    seen.add(id);
    const e = entries[id];
    visit(e.left, prefix, seen);
    visit(e.right, prefix, seen);
    const path = prefix + e.name;
    if (e.type === 2) streams.set(path, e.size < cutoff ? readMini(e.start, e.size) : readChain(e.start, e.size));
    else if (e.type === 1) visit(e.child, `${path}/`, seen);
  };
  visit(root.child, "", new Set());
  return streams;
}

/** CFB's order of names in a storage: shorter first, then by upper-case code units. */
function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

/** A compound file holding these streams (paths with "/" between storages). */
export function writeCfb(streams: Map<string, Uint8Array>): Uint8Array {
  const SECTOR = 512;
  const MINI = 64;
  const CUTOFF = 4096;
  interface Node { name: string; type: 1 | 2 | 5; data?: Uint8Array; children: Node[]; id?: number; start?: number; left?: number; right?: number; child?: number }
  const root: Node = { name: "Root Entry", type: 5, children: [] };
  for (const [path, data] of streams) {
    const parts = path.split("/");
    let parent = root;
    for (const part of parts.slice(0, -1)) {
      let s = parent.children.find((c) => c.name === part && c.type === 1);
      if (!s) { s = { name: part, type: 1, children: [] }; parent.children.push(s); }
      parent = s;
    }
    parent.children.push({ name: parts[parts.length - 1], type: 2, data, children: [] });
  }
  // Directory ids in order; each storage's children as a balanced tree (all black: readers accept it).
  const all: Node[] = [];
  const number = (n: Node) => { n.id = all.length; all.push(n); for (const c of n.children) number(c); };
  number(root);
  const tree = (list: Node[]): number => {
    if (!list.length) return NOSTREAM;
    const sorted = [...list].sort((a, b) => compareNames(a.name, b.name));
    const build = (lo: number, hi: number): number => {
      if (lo > hi) return NOSTREAM;
      const mid = (lo + hi) >> 1;
      sorted[mid].left = build(lo, mid - 1);
      sorted[mid].right = build(mid + 1, hi);
      return sorted[mid].id!;
    };
    return build(0, sorted.length - 1);
  };
  for (const n of all) n.child = tree(n.children);

  // Small streams go in the mini stream (64-byte sectors).
  const miniFat: number[] = [];
  const miniParts: Uint8Array[] = [];
  let miniSectors = 0;
  const big: Node[] = [];
  for (const n of all) {
    if (n.type !== 2 || !n.data) continue;
    if (n.data.length < CUTOFF) {
      if (!n.data.length) { n.start = END; continue; }
      const count = Math.ceil(n.data.length / MINI);
      n.start = miniSectors;
      for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? END : miniSectors + i + 1);
      const padded = new Uint8Array(count * MINI);
      padded.set(n.data);
      miniParts.push(padded);
      miniSectors += count;
    } else big.push(n);
  }
  const miniStream = new Uint8Array(miniSectors * MINI);
  { let o = 0; for (const p of miniParts) { miniStream.set(p, o); o += p.length; } }
  const sectorsOf = (len: number) => Math.ceil(len / SECTOR);
  const dirSectors = sectorsOf(all.length * 128);
  const miniFatSectors = sectorsOf(miniFat.length * 4);
  const miniStreamSectors = sectorsOf(miniStream.length);
  const bigSectors = big.reduce((a, n) => a + sectorsOf(n.data!.length), 0);
  // FAT and DIFAT sectors cover every sector, themselves included.
  const content = dirSectors + miniFatSectors + miniStreamSectors + bigSectors;
  let fatSectors = 1;
  let difatSectors = 0;
  for (;;) {
    difatSectors = fatSectors > 109 ? Math.ceil((fatSectors - 109) / (SECTOR / 4 - 1)) : 0;
    const total = content + fatSectors + difatSectors;
    if (fatSectors * (SECTOR / 4) >= total) break;
    fatSectors++;
  }
  const total = content + fatSectors + difatSectors;
  const fat = new Array<number>(fatSectors * (SECTOR / 4)).fill(FREE);
  let next = 0;
  const allocate = (count: number, mark?: number): number => {
    if (!count) return END;
    const start = next;
    for (let i = 0; i < count; i++) fat[next + i] = mark ?? (i === count - 1 ? END : next + i + 1);
    next += count;
    return start;
  };
  const fatStart = allocate(fatSectors, FATSECT);
  const difatStart = allocate(difatSectors, DIFSECT);
  const dirStart = allocate(dirSectors);
  const miniFatStart = allocate(miniFatSectors);
  root.start = miniStream.length ? allocate(miniStreamSectors) : END;
  for (const n of big) n.start = allocate(sectorsOf(n.data!.length));

  const out = new Uint8Array((total + 1) * SECTOR);
  const dv = new DataView(out.buffer);
  const at = (s: number) => (s + 1) * SECTOR;
  // Header.
  out.set(SIGNATURE, 0);
  dv.setUint16(0x18, 0x3e, true); // minor version
  dv.setUint16(0x1a, 3, true); // major version 3
  dv.setUint16(0x1c, 0xfffe, true); // little-endian
  dv.setUint16(0x1e, 9, true); // 512-byte sectors
  dv.setUint16(0x20, 6, true); // 64-byte mini sectors
  dv.setUint32(0x2c, fatSectors, true);
  dv.setUint32(0x30, dirStart, true);
  dv.setUint32(0x38, CUTOFF, true);
  dv.setUint32(0x3c, miniFatSectors ? miniFatStart : END, true);
  dv.setUint32(0x40, miniFatSectors, true);
  dv.setUint32(0x44, difatSectors ? difatStart : END, true);
  dv.setUint32(0x48, difatSectors, true);
  for (let i = 0; i < 109; i++) dv.setUint32(0x4c + i * 4, i < fatSectors ? fatStart + i : FREE, true);
  // DIFAT sectors: the FAT sectors after the first 109, then the next DIFAT sector.
  for (let d = 0; d < difatSectors; d++) {
    const per = SECTOR / 4 - 1;
    for (let i = 0; i < per; i++) { const k = 109 + d * per + i; dv.setUint32(at(difatStart + d) + i * 4, k < fatSectors ? fatStart + k : FREE, true); }
    dv.setUint32(at(difatStart + d) + per * 4, d === difatSectors - 1 ? END : difatStart + d + 1, true);
  }
  // FAT.
  fat.forEach((v, i) => dv.setUint32(at(fatStart) + i * 4, v, true));
  // Directory.
  for (let i = 0; i < dirSectors * 4; i++) {
    const o = at(dirStart) + i * 128;
    const n = all[i];
    if (!n) { dv.setUint32(o + 68, NOSTREAM, true); dv.setUint32(o + 72, NOSTREAM, true); dv.setUint32(o + 76, NOSTREAM, true); continue; }
    for (let c = 0; c < Math.min(31, n.name.length); c++) dv.setUint16(o + c * 2, n.name.charCodeAt(c), true);
    dv.setUint16(o + 64, (Math.min(31, n.name.length) + 1) * 2, true);
    out[o + 66] = n.type;
    out[o + 67] = 1; // black
    dv.setUint32(o + 68, n.left ?? NOSTREAM, true);
    dv.setUint32(o + 72, n.right ?? NOSTREAM, true);
    dv.setUint32(o + 76, n.type === 2 ? NOSTREAM : n.child ?? NOSTREAM, true);
    dv.setUint32(o + 116, n.type === 1 ? 0 : n.start ?? END, true);
    dv.setUint32(o + 120, n.type === 5 ? miniStream.length : n.type === 2 ? n.data!.length : 0, true);
  }
  // Mini FAT, mini stream, big streams.
  miniFat.forEach((v, i) => dv.setUint32(at(miniFatStart) + i * 4, v, true));
  for (let i = miniFat.length; i < miniFatSectors * (SECTOR / 4); i++) dv.setUint32(at(miniFatStart) + i * 4, FREE, true);
  if (miniStream.length) out.set(miniStream, at(root.start));
  for (const n of big) out.set(n.data!, at(n.start!));
  return out;
}
