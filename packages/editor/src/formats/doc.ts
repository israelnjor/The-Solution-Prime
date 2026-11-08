/**
 * Word 97–2003 documents (.doc), read (MS-DOC): the main text from the piece table, character
 * formatting (font, size, bold, italic, underline, strike, caps, colour, highlight, super/subscript,
 * hidden), paragraph formatting (alignment, indents, spacing, line spacing, keep with next, page break
 * before), styles (their formatting; headings by built-in identifier), lists (as bullets or numbers),
 * tables (cell widths), fields (results shown; links kept), inline pictures (PNG, JPEG), breaks.
 * Headers, footers, footnotes, comments and floating objects are left out (with a note).
 */
import { isCfb, readCfb } from "@prime/docx";
import type { Block, Inline, ParagraphProps, PrimeDocument, RunProps, Table, TableCell } from "@prime/pdm";
import { addImage, headingStyle, listId, newDocument } from "./common.js";

const ICO = ["", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0"];
const HIGHLIGHT = ["", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray"];
const UNDERLINE: Record<number, string> = { 1: "single", 2: "words", 3: "double", 4: "dotted", 6: "thick", 7: "dash", 9: "dotDash", 10: "dotDotDash", 11: "wave" };

interface Sprm { op: number; arg: DataView; at: number; len: number }

/** The sprms of a grpprl. */
function sprms(bytes: Uint8Array, start: number, end: number): Sprm[] {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out: Sprm[] = [];
  let o = start;
  while (o + 2 <= end) {
    const op = dv.getUint16(o, true);
    o += 2;
    const spra = op >> 13;
    let len = [1, 1, 2, 4, 2, 2, 0, 3][spra];
    if (spra === 6) {
      if (op === 0xd608 || op === 0xd606) { len = dv.getUint16(o, true) - 1; o += 2; } // sprmTDefTable: a 2-byte length (counting itself… less one)
      else if (op === 0xc615) { len = bytes[o]; o += 1; if (len === 255) len = 0; }
      else { len = bytes[o]; o += 1; }
    }
    if (o + len > end + 1) break;
    out.push({ op, arg: new DataView(bytes.buffer, bytes.byteOffset + o, Math.max(0, Math.min(len, bytes.length - o))), at: o, len });
    o += len;
  }
  return out;
}

interface CharFmt { b?: boolean; i?: boolean; strike?: boolean; caps?: boolean; scaps?: boolean; hidden?: boolean; ul?: string; size?: number; ftc?: number; color?: string; hl?: string; va?: "superscript" | "subscript"; special?: boolean; picLocation?: number }
interface ParaFmt { jc?: number; left?: number; right?: number; first?: number; before?: number; after?: number; line?: number; lineMult?: boolean; keepNext?: boolean; pageBreakBefore?: boolean; inTable?: boolean; ttp?: boolean; ilfo?: number; ilvl?: number; outline?: number; cellEdges?: number[] }

const toggle = (v: number, base?: boolean) => (v === 0 ? false : v === 1 ? true : v === 0x81 ? !base : base ?? false);

function applyChar(f: CharFmt, list: Sprm[]): CharFmt {
  const c = { ...f };
  for (const s of list) {
    const b = s.len >= 1 ? s.arg.getUint8(0) : 0;
    switch (s.op) {
      case 0x0835: c.b = toggle(b, f.b); break;
      case 0x0836: c.i = toggle(b, f.i); break;
      case 0x0837: c.strike = toggle(b, f.strike); break;
      case 0x083a: c.scaps = toggle(b, f.scaps); break;
      case 0x083b: c.caps = toggle(b, f.caps); break;
      case 0x083c: c.hidden = toggle(b, f.hidden); break;
      case 0x2a3e: c.ul = b ? UNDERLINE[b] ?? "single" : undefined; break;
      case 0x4a43: c.size = s.arg.getUint16(0, true) / 2; break;
      case 0x4a4f: c.ftc = s.arg.getUint16(0, true); break;
      case 0x2a42: c.color = ICO[b] || undefined; break;
      case 0x6870: if (s.len >= 4 && !(s.arg.getUint8(3) & 0xff)) c.color = [0, 1, 2].map((k) => s.arg.getUint8(k).toString(16).padStart(2, "0")).join("").toUpperCase(); break;
      case 0x2a0c: c.hl = HIGHLIGHT[b] || undefined; break;
      case 0x2a48: c.va = b === 1 ? "superscript" : b === 2 ? "subscript" : undefined; break;
      case 0x0855: c.special = b !== 0; break;
      case 0x6a03: c.picLocation = s.arg.getUint32(0, true); break;
      default: break;
    }
  }
  return c;
}

function applyPara(f: ParaFmt, list: Sprm[]): ParaFmt {
  const p = { ...f };
  for (const s of list) {
    const i16 = () => s.arg.getInt16(0, true);
    switch (s.op) {
      case 0x2403: case 0x2461: p.jc = s.arg.getUint8(0); break;
      case 0x840f: case 0x845e: p.left = i16(); break;
      case 0x840e: case 0x845d: p.right = i16(); break;
      case 0x8411: case 0x8460: p.first = i16(); break;
      case 0xa413: p.before = s.arg.getUint16(0, true); break;
      case 0xa414: p.after = s.arg.getUint16(0, true); break;
      case 0x6412: p.line = i16(); p.lineMult = s.arg.getInt16(2, true) === 1; break;
      case 0x2406: p.keepNext = s.arg.getUint8(0) !== 0; break;
      case 0x2407: p.pageBreakBefore = s.arg.getUint8(0) !== 0; break;
      case 0x2416: p.inTable = s.arg.getUint8(0) !== 0; break;
      case 0x2417: p.ttp = s.arg.getUint8(0) !== 0; break;
      case 0x260a: p.ilvl = s.arg.getUint8(0); break;
      case 0x460b: p.ilfo = s.arg.getInt16(0, true); break;
      case 0x2640: p.outline = s.arg.getUint8(0); break;
      case 0xd608: {
        const n = s.arg.getUint8(0);
        const edges: number[] = [];
        for (let k = 0; k <= n && 1 + k * 2 + 2 <= s.len; k++) edges.push(s.arg.getInt16(1 + k * 2, true));
        p.cellEdges = edges;
        break;
      }
      default: break;
    }
  }
  return p;
}

export function readDoc(bytes: Uint8Array): { doc: PrimeDocument; warnings: string[] } {
  if (!isCfb(bytes)) throw new Error("This is not a Word 97–2003 document.");
  const streams = readCfb(bytes);
  const wd = streams.get("WordDocument");
  if (!wd) throw new Error("This compound file holds no Word document.");
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  if (dv.getUint16(0, true) !== 0xa5ec) throw new Error("This Word document's header is not recognised.");
  const flags = dv.getUint16(0x0a, true);
  if (flags & 0x0100) throw new Error("This Word 97–2003 document is protected with a password; open it in Word and save it as .docx first.");
  const table = streams.get(flags & 0x0200 ? "1Table" : "0Table");
  if (!table) throw new Error("This Word document's table stream is missing.");
  const data = streams.get("Data");
  const tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const warnings = new Set<string>();

  // The FIB: text lengths and where the structures are in the table stream.
  const csw = dv.getUint16(32, true);
  const lwAt = 34 + csw * 2;
  const cslw = dv.getUint16(lwAt, true);
  const rgLw = lwAt + 2;
  const ccpText = dv.getInt32(rgLw + 3 * 4, true);
  const ccpFtn = dv.getInt32(rgLw + 4 * 4, true);
  const ccpHdd = dv.getInt32(rgLw + 5 * 4, true);
  const ccpAtn = dv.getInt32(rgLw + 7 * 4, true);
  if (ccpFtn > 0) warnings.add("Footnotes were left out.");
  if (ccpHdd > 0) warnings.add("Headers and footers were left out.");
  if (ccpAtn > 0) warnings.add("Comments were left out.");
  const fcAt = rgLw + cslw * 4 + 2;
  const fcLcb = (i: number) => ({ fc: dv.getUint32(fcAt + i * 8, true), lcb: dv.getUint32(fcAt + i * 8 + 4, true) });

  // Piece table (the CLX): where each stretch of text is, and whether it is 8-bit (cp1252) or UTF-16.
  const clx = fcLcb(33);
  let o = clx.fc;
  const clxEnd = clx.fc + clx.lcb;
  while (o < clxEnd && table[o] === 0x01) o += 3 + tdv.getUint16(o + 1, true);
  if (table[o] !== 0x02) throw new Error("This Word document's text could not be found.");
  const lcb = tdv.getUint32(o + 1, true);
  const plc = o + 5;
  const nPieces = (lcb - 4) / 12;
  const pieces: { cp: number; cpEnd: number; fc: number; compressed: boolean }[] = [];
  for (let k = 0; k < nPieces; k++) {
    const cp = tdv.getUint32(plc + k * 4, true);
    const cpEnd = tdv.getUint32(plc + (k + 1) * 4, true);
    const raw = tdv.getUint32(plc + (nPieces + 1) * 4 + k * 8 + 2, true);
    const compressed = (raw & 0x40000000) !== 0;
    pieces.push({ cp, cpEnd, fc: compressed ? (raw & 0x3fffffff) / 2 : raw, compressed });
  }
  const cp1252 = new TextDecoder("windows-1252");
  let text = "";
  const fcs: number[] = []; // each character's file offset (for its formatting)
  for (const p of pieces) {
    if (p.cp >= ccpText) break;
    const end = Math.min(p.cpEnd, ccpText);
    const n = end - p.cp;
    if (p.compressed) {
      text += cp1252.decode(wd.subarray(p.fc, p.fc + n));
      for (let k = 0; k < n; k++) fcs.push(p.fc + k);
    } else {
      for (let k = 0; k < n; k++) { text += String.fromCharCode(dv.getUint16(p.fc + k * 2, true)); fcs.push(p.fc + k * 2); }
    }
  }

  // Formatted runs (FKPs): character and paragraph, by file offset.
  const runs = (bte: { fc: number; lcb: number }, kind: "chp" | "pap") => {
    const out: { start: number; end: number; grpprl: Uint8Array; istd?: number }[] = [];
    const n = (bte.lcb - 4) / 8;
    for (let k = 0; k < n; k++) {
      const pn = tdv.getUint32(bte.fc + (n + 1) * 4 + k * 4, true) & 0x3fffff;
      const page = wd.subarray(pn * 512, pn * 512 + 512);
      const pdv = new DataView(page.buffer, page.byteOffset, 512);
      const count = page[511];
      for (let r = 0; r < count; r++) {
        const start = pdv.getUint32(r * 4, true);
        const end = pdv.getUint32((r + 1) * 4, true);
        if (kind === "chp") {
          const off = page[(count + 1) * 4 + r] * 2;
          out.push({ start, end, grpprl: off ? page.subarray(off + 1, off + 1 + page[off]) : new Uint8Array(0) });
        } else {
          const off = page[(count + 1) * 4 + r * 13] * 2;
          if (!off) { out.push({ start, end, grpprl: new Uint8Array(0), istd: 0 }); continue; }
          let cb = page[off];
          let at2 = off + 1;
          let size: number;
          if (cb === 0) { cb = page[at2]; at2++; size = cb * 2; } else size = cb * 2 - 1;
          const istd = pdv.getUint16(at2, true);
          out.push({ start, end, grpprl: page.subarray(at2 + 2, at2 + size), istd });
        }
      }
    }
    return out.sort((a, b) => a.start - b.start);
  };
  const chpRuns = runs(fcLcb(12), "chp");
  const papRuns = runs(fcLcb(13), "pap");
  const find = <T extends { start: number; end: number }>(list: T[], fc: number): T | undefined => {
    let lo = 0;
    let hi = list.length - 1;
    while (lo <= hi) { const m = (lo + hi) >> 1; if (list[m].end <= fc) lo = m + 1; else if (list[m].start > fc) hi = m - 1; else return list[m]; }
    return undefined;
  };

  // Fonts.
  const fonts: string[] = [];
  {
    const f = fcLcb(15);
    if (f.lcb) {
      let p = f.fc;
      const count = tdv.getUint16(p, true);
      p += 4;
      for (let k = 0; k < count && p < f.fc + f.lcb; k++) {
        const len = table[p] + 1;
        let name = "";
        for (let q = p + 40; q + 1 < p + len; q += 2) { const ch = tdv.getUint16(q, true); if (!ch) break; name += String.fromCharCode(ch); }
        fonts.push(name);
        p += len;
      }
    }
  }

  // Styles: built-in identifier, base style, formatting.
  interface DocStyle { sti: number; base: number; name: string; chp: Sprm[]; pap: Sprm[] }
  const styles: (DocStyle | undefined)[] = [];
  {
    const st = fcLcb(1);
    if (st.lcb) {
      const cbStshi = tdv.getUint16(st.fc, true);
      const shi = st.fc + 2;
      const cstd = tdv.getUint16(shi, true);
      const cbBase = tdv.getUint16(shi + 2, true);
      let p = shi + cbStshi;
      for (let k = 0; k < cstd; k++) {
        const cb = tdv.getUint16(p, true);
        const s = p + 2;
        if (cb) {
          const sti = tdv.getUint16(s, true) & 0x0fff;
          const stk = tdv.getUint16(s + 2, true) & 0x000f;
          const base = tdv.getUint16(s + 2, true) >> 4;
          const cupx = tdv.getUint16(s + 4, true) & 0x000f;
          let q = s + cbBase;
          const cch = tdv.getUint16(q, true);
          let name = "";
          for (let c = 0; c < cch; c++) name += String.fromCharCode(tdv.getUint16(q + 2 + c * 2, true));
          q += 2 + cch * 2 + 2;
          const upx: Uint8Array[] = [];
          for (let u = 0; u < cupx && q < s + cb; u++) {
            if ((q - s) % 2) q++;
            const len = tdv.getUint16(q, true);
            upx.push(table.subarray(q + 2, q + 2 + len));
            q += 2 + len;
          }
          const isPara = stk === 1;
          const papU = isPara ? upx[0] : undefined;
          const chpU = isPara ? upx[1] : upx[0];
          styles[k] = {
            sti, base: base === 0x0fff ? -1 : base, name,
            pap: papU && papU.length > 2 ? sprms(papU, 2, papU.length) : [],
            chp: chpU ? sprms(chpU, 0, chpU.length) : [],
          };
        }
        p = s + cb;
      }
    }
  }
  const styleFmt = new Map<number, { c: CharFmt; p: ParaFmt }>();
  const resolveStyle = (istd: number, depth = 0): { c: CharFmt; p: ParaFmt } => {
    const known = styleFmt.get(istd);
    if (known) return known;
    const s = styles[istd];
    const base = s && s.base >= 0 && depth < 10 ? resolveStyle(s.base, depth + 1) : { c: { size: 10 }, p: {} };
    const r = s ? { c: applyChar(base.c, s.chp), p: applyPara(base.p, s.pap) } : base;
    styleFmt.set(istd, r);
    return r;
  };

  // Lists: each list override (ilfo, 1-based) → its list → each level's number format (23: a bullet).
  const listFormats = new Map<number, number[]>(); // lsid → nfc per level
  const lfoLists: number[] = []; // ilfo-1 → lsid
  {
    const lst = fcLcb(73);
    if (lst.lcb) {
      const count = tdv.getUint16(lst.fc, true);
      const lstfs: { lsid: number; simple: boolean }[] = [];
      for (let k = 0; k < count; k++) { const at2 = lst.fc + 2 + k * 28; lstfs.push({ lsid: tdv.getInt32(at2, true), simple: (table[at2 + 26] & 1) === 1 }); }
      let q = lst.fc + 2 + count * 28;
      for (const l of lstfs) {
        const nfcs: number[] = [];
        for (let lv = 0; lv < (l.simple ? 1 : 9) && q + 28 <= table.length; lv++) {
          nfcs.push(table[q + 4]);
          const cbChpx = table[q + 24];
          const cbPapx = table[q + 25];
          q += 28 + cbPapx + cbChpx;
          const cch = tdv.getUint16(q, true);
          q += 2 + cch * 2;
        }
        listFormats.set(l.lsid, nfcs);
      }
    }
    const lfo = fcLcb(74);
    if (lfo.lcb) {
      const n = tdv.getUint32(lfo.fc, true);
      for (let k = 0; k < n; k++) lfoLists.push(tdv.getInt32(lfo.fc + 4 + k * 16, true));
    }
  }
  const listKindOf = (ilfo: number, ilvl: number): "bullet" | "number" | undefined => {
    const nfcs = listFormats.get(lfoLists[ilfo - 1]);
    const nfc = nfcs?.[Math.min(ilvl, nfcs.length - 1)];
    return nfc === undefined ? "bullet" : nfc === 0xff ? undefined : nfc === 23 ? "bullet" : "number";
  };

  // The document: Normal's font and size as its defaults, no spacing of its own.
  const doc = newDocument();
  const normal = resolveStyle(0).c;
  const normalFont = normal.ftc !== undefined ? fonts[normal.ftc] : undefined;
  doc.styles.defaults.run = { ...doc.styles.defaults.run, size: normal.size ?? 12, ...(normalFont ? { font: { ascii: normalFont, hAnsi: normalFont, cs: normalFont, eastAsia: normalFont } } : {}) };
  doc.styles.defaults.paragraph = { spacing: { after: 0, line: 240, lineRule: "auto" } };

  const runProps = (c: CharFmt): RunProps => {
    const p: RunProps = {};
    if (c.b) p.bold = true;
    if (c.i) p.italic = true;
    if (c.strike) p.strike = true;
    if (c.caps) p.caps = true;
    if (c.scaps) p.smallCaps = true;
    if (c.hidden) p.hidden = true;
    if (c.ul) p.underline = { val: c.ul };
    if (c.size && c.size !== doc.styles.defaults.run?.size) p.size = c.size;
    const font = c.ftc !== undefined ? fonts[c.ftc] : undefined;
    if (font && font !== normalFont) p.font = { ascii: font, hAnsi: font, cs: font };
    if (c.color) p.color = c.color;
    if (c.hl) p.highlight = c.hl;
    if (c.va) p.vertAlign = c.va;
    return p;
  };
  const paraProps = (f: ParaFmt, sti: number, styleName = ""): ParagraphProps => {
    const p: ParagraphProps = {};
    if (f.jc) p.align = ["", "center", "right", "both"][f.jc] || undefined;
    if (!p.align) delete p.align;
    if (f.left || f.right || f.first) {
      p.indent = {};
      if (f.left) p.indent.left = f.left / 20;
      if (f.right) p.indent.right = f.right / 20;
      if (f.first && f.first > 0) p.indent.firstLine = f.first / 20;
      if (f.first && f.first < 0) p.indent.hanging = -f.first / 20;
    }
    if (f.before || f.after || (f.line && !(f.lineMult && f.line === 240))) {
      p.spacing = {};
      if (f.before) p.spacing.before = f.before / 20;
      if (f.after) p.spacing.after = f.after / 20;
      if (f.line) { if (f.lineMult) { p.spacing.line = f.line; p.spacing.lineRule = "auto"; } else { p.spacing.line = Math.abs(f.line); p.spacing.lineRule = f.line < 0 ? "exact" : "atLeast"; } }
    }
    if (f.keepNext) p.keepNext = true;
    if (f.pageBreakBefore) p.pageBreakBefore = true;
    const level = sti >= 1 && sti <= 9 ? sti : f.outline !== undefined && f.outline < 9 ? f.outline + 1 : undefined;
    if (level) p.style = headingStyle(doc, level);
    else if (sti === 62) p.style = "Title";
    const kind = f.ilfo && f.ilfo > 0 ? listKindOf(f.ilfo, f.ilvl ?? 0) : undefined;
    // ("List Bullet 2", "List Number 3": Word's nested list styles, each a list of its own one level deeper.)
    const nested = /^List (?:Bullet|Number)\s*([2-9])$/i.exec(styleName);
    if (kind) { p.list = { numId: listId(doc, kind), level: Math.min(8, Math.max(f.ilvl ?? 0, nested ? Number(nested[1]) - 1 : 0)) }; delete p.indent; }
    return p;
  };

  // Pictures: the PICF in the Data stream holds an Office drawing whose BLIP is the PNG or JPEG.
  const picture = (loc: number): { data: Uint8Array; type: string; w: number; h: number } | undefined => {
    if (!data || loc + 68 > data.length) return undefined;
    const ddv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const lcbPic = ddv.getUint32(loc, true);
    const cbHeader = ddv.getUint16(loc + 4, true);
    const dxaGoal = ddv.getInt16(loc + 28, true);
    const dyaGoal = ddv.getInt16(loc + 30, true);
    const mx = ddv.getUint16(loc + 32, true) || 1000;
    const my = ddv.getUint16(loc + 34, true) || 1000;
    const end = Math.min(data.length, loc + lcbPic);
    for (let q = loc + cbHeader; q + 8 <= end; q++) {
      const type = ddv.getUint16(q + 2, true);
      if (type !== 0xf01e && type !== 0xf01d) continue;
      const len = ddv.getUint32(q + 4, true);
      const body = data.subarray(q + 8, Math.min(end, q + 8 + len));
      const sig = type === 0xf01e ? [0x89, 0x50, 0x4e, 0x47] : [0xff, 0xd8, 0xff];
      for (let s = 0; s < Math.min(64, body.length); s++) {
        if (sig.every((v, k) => body[s + k] === v)) {
          return { data: body.slice(s), type: type === 0xf01e ? "image/png" : "image/jpeg", w: (dxaGoal * mx) / 1000 / 20, h: (dyaGoal * my) / 1000 / 20 };
        }
      }
    }
    return undefined;
  };

  // Paragraphs: text up to each paragraph mark (\r) or cell mark (\x07).
  const blocks: Block[] = [];
  let tbl: Table | undefined;
  let cell: Block[] = [];
  let row: TableCell[] = [];
  let fieldStack: { instr: string; start: number; inResult: boolean; inlines: Inline[] }[] = [];
  let start = 0;
  let pictures = 0;
  for (let k = 0; k < text.length; k++) {
    const ch = text[k];
    if (ch !== "\r" && ch !== "\x07") continue;
    const pap = find(papRuns, fcs[k]);
    const style = resolveStyle(pap?.istd ?? 0);
    const pf = applyPara(style.p, pap ? sprms(pap.grpprl, 0, pap.grpprl.length) : []);
    const sti = styles[pap?.istd ?? 0]?.sti ?? 0;
    // The paragraph's inlines.
    const inlines: Inline[] = [];
    let target = inlines;
    fieldStack = [];
    const push = (i: Inline) => { if (fieldStack.length && !fieldStack[fieldStack.length - 1].inResult) return; (fieldStack.length ? fieldStack[fieldStack.length - 1].inlines : target).push(i); };
    for (let q = start; q < k; q++) {
      const c = text[q];
      const chp = find(chpRuns, fcs[q]);
      const cf = applyChar(style.c, chp ? sprms(chp.grpprl, 0, chp.grpprl.length) : []);
      if (c === "\x13") { fieldStack.push({ instr: "", start: q, inResult: false, inlines: [] }); continue; }
      if (c === "\x14" && fieldStack.length) { fieldStack[fieldStack.length - 1].inResult = true; continue; }
      if (c === "\x15" && fieldStack.length) {
        const f = fieldStack.pop()!;
        const link = /^\s*HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(f.instr);
        const into = fieldStack.length ? fieldStack[fieldStack.length - 1].inlines : target;
        if (fieldStack.length && !fieldStack[fieldStack.length - 1].inResult) continue;
        if (link) into.push({ type: "link", ...(/\\l/.test(f.instr) ? { anchor: link[1] } : { href: link[1] }), content: f.inlines });
        else into.push(...f.inlines);
        continue;
      }
      if (fieldStack.length && !fieldStack[fieldStack.length - 1].inResult) { fieldStack[fieldStack.length - 1].instr += c; continue; }
      const props = runProps(cf);
      const withProps = Object.keys(props).length ? { props } : {};
      if (cf.special && c === "\x01") {
        const pic = cf.picLocation !== undefined ? picture(cf.picLocation) : undefined;
        if (pic) { pictures++; push({ type: "image", imageId: addImage(doc, pic.data, pic.type), width: pic.w || 100, height: pic.h || 100 }); }
        else warnings.add("Some pictures (other than PNG or JPEG) were left out.");
        continue;
      }
      if (cf.special && c === "\x08") { warnings.add("Floating pictures and shapes were left out."); continue; }
      if (cf.special) continue; // footnote references, annotations and other special characters
      if (c === "\t") { push({ type: "tab", ...withProps }); continue; }
      if (c === "\x0b") { push({ type: "break", kind: "line" }); continue; }
      if (c === "\x0c") { push({ type: "break", kind: "page" }); continue; }
      if (c === "\x0e") { push({ type: "break", kind: "column" }); continue; }
      const t = c === "\x1e" ? "‑" : c === "\x1f" ? "­" : c;
      if (t < " " && t !== "­") continue;
      const into = fieldStack.length ? fieldStack[fieldStack.length - 1].inlines : target;
      const last = into[into.length - 1];
      if (last?.type === "text" && JSON.stringify(last.props ?? {}) === JSON.stringify(props)) last.text += t;
      else into.push({ type: "text", text: t, ...withProps });
    }
    start = k + 1;
    const para = { type: "paragraph" as const, props: paraProps(pf, sti, styles[pap?.istd ?? 0]?.name), content: inlines };
    if (pf.inTable) {
      if (pf.ttp) {
        // The row's end: its cells; their widths from the row's definition.
        if (cell.length) { row.push({ props: {}, blocks: cell }); cell = []; }
        const edges = pf.cellEdges ?? [];
        const widths = edges.slice(1).map((e, n) => (e - edges[n]) / 20);
        row.forEach((c, n) => { if (widths[n]) c.props.width = { w: String(Math.round(widths[n] * 20)), type: "dxa" }; });
        if (!tbl) tbl = { type: "table", props: { style: "TableGrid" }, grid: widths.length ? widths : row.map(() => 100), rows: [] };
        if (row.length) tbl.rows.push({ props: {}, cells: row });
        if (tbl.grid.length < row.length) tbl.grid = row.map((_, n) => widths[n] ?? 100);
        row = [];
      } else {
        cell.push(para);
        if (ch === "\x07") { row.push({ props: {}, blocks: cell }); cell = []; }
      }
    } else {
      if (tbl) { blocks.push(tbl); tbl = undefined; }
      blocks.push(para);
    }
  }
  if (tbl) blocks.push(tbl);
  doc.sections[0].blocks = blocks.length ? blocks : [{ type: "paragraph", props: {}, content: [] }];
  void pictures;
  return { doc, warnings: [...warnings] };
}
