/**
 * Rich Text Format (.rtf), read and written. Read: text with its character formatting (font, size,
 * bold, italic, underline, strike, colour, highlight, super/subscript, caps, hidden), paragraph
 * formatting (alignment, indents, spacing, keep with next), headings (by style name or outline level),
 * lists (as bullets or numbers), tables, pictures (PNG, JPEG), links, special characters and Unicode,
 * page size and margins, and the title and author. Written: the same, from what the document shows
 * (styles resolved), list labels as text.
 */
import { withProperties } from "@prime/docx";
import type { Block, Inline, ParagraphProps, PrimeDocument, RunProps, Table, TableCell } from "@prime/pdm";
import { addImage, headingStyle, imageBytes, listId, newDocument, type Piece, type Shown, type ShownParagraph, shownBody } from "./common.js";

// --- reading -----------------------------------------------------------------------------------------

interface CharState {
  b?: boolean; i?: boolean; ul?: string; strike?: boolean; fs?: number; f?: number; cf?: number; hl?: number;
  va?: "superscript" | "subscript"; caps?: boolean; scaps?: boolean; v?: boolean;
}
interface Group { c: CharState; skip: boolean; dest?: string; uc: number }
interface ParaState { align?: string; li?: number; ri?: number; fi?: number; sb?: number; sa?: number; sl?: number; slmult?: boolean; keepn?: boolean; style?: number; intbl?: boolean; ls?: number; ilvl?: number; outline?: number }

const HIGHLIGHTS: Record<string, string> = { "FFFF00": "yellow", "00FF00": "green", "00FFFF": "cyan", "FF00FF": "magenta", "0000FF": "blue", "FF0000": "red", "000080": "darkBlue", "008080": "darkCyan", "008000": "darkGreen", "800080": "darkMagenta", "800000": "darkRed", "808000": "darkYellow", "808080": "darkGray", "C0C0C0": "lightGray", "000000": "black" };

const SPECIAL: Record<string, string> = {
  emdash: "\u2014", endash: "\u2013", bullet: "\u2022", lquote: "\u2018", rquote: "\u2019", ldblquote: "\u201C", rdblquote: "\u201D",
  emspace: "\u2003", enspace: "\u2002", qmspace: "\u2005", zwj: "\u200D", zwnj: "\u200C",
};
const PARA_WORDS = new Set(["pard", "ql", "qc", "qr", "qj", "li", "lin", "ri", "rin", "fi", "sb", "sa", "sl", "slmult", "keepn", "s", "outlinelevel", "intbl", "ls", "ilvl", "trowd", "cellx"]);

/** Destinations whose text is not part of the body. */
const SKIP = new Set(["fonttbl", "colortbl", "stylesheet", "info", "pict", "fldinst", "listtable", "listoverridetable", "rsidtbl", "generator", "themedata",
  "colorschememapping", "datastore", "latentstyles", "xmlnstbl", "mmathPr", "header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf",
  "footnote", "annotation", "atnid", "atnauthor", "comment", "bkmkstart", "bkmkend", "pgdsctbl", "listtext", "pntext", "pntxta", "pntxtb", "shpinst", "shp",
  "nonshppict", "objdata", "object", "result", "template", "filetbl", "revtbl", "protusertbl", "wgrffmtfilter", "panose", "falt", "fontemb", "fontfile", "title",
  "author", "subject", "keywords", "operator", "category", "comment", "doccomm", "company", "manager", "userprops", "xe", "tc", "picprop", "blipuid", "sp", "sn", "sv"]);

export function readRtf(bytes: Uint8Array): { doc: PrimeDocument; warnings: string[] } {
  const src = new TextDecoder("latin1").decode(bytes);
  const doc = newDocument();
  const warnings = new Set<string>();
  // RTF's own defaults: 12 pt, no space between paragraphs, single lines.
  doc.styles.defaults.run = { ...(doc.styles.defaults.run ?? {}), size: 12 };
  doc.styles.defaults.paragraph = { spacing: { after: 0, line: 240, lineRule: "auto" } };

  let codepage = 1252;
  let deff = 0;
  const fonts = new Map<number, string>();
  const colors: (string | undefined)[] = [];
  const styleNames = new Map<number, string>();
  const info: Record<string, string> = {};
  let decoder = new TextDecoder("windows-1252");

  const blocks: Block[] = [];
  let inlines: Inline[] = [];
  let para: ParaState = {};
  // Tables: the current row's cells (each its paragraphs), the rows so far.
  let cellBlocks: Block[] = [];
  let rowCells: TableCell[] = [];
  let rowWidths: number[] = [];
  let table: Table | undefined;
  // A field: its instruction (HYPERLINK "…"), and the inlines its result adds.
  const fields: { instr: string; start: number }[] = [];
  // A picture being read.
  let pict: { hex: string[]; type?: string; w?: number; h?: number; picw?: number; pich?: number; sx?: number; sy?: number } | undefined;
  // A list item's label (\listtext): bullet or number.
  let listText = "";

  const stack: Group[] = [];
  let g: Group = { c: {}, skip: false, uc: 1 };
  let pendingSkip = 0; // characters to skip after \uN

  const runProps = (c: CharState): RunProps => {
    const p: RunProps = {};
    if (c.b) p.bold = true;
    if (c.i) p.italic = true;
    if (c.ul) p.underline = { val: c.ul };
    if (c.strike) p.strike = true;
    if (c.fs !== undefined && c.fs !== 24) p.size = c.fs / 2;
    else if (c.fs === undefined) { /* the 12 pt default */ }
    const f = fonts.get(c.f ?? deff);
    if (f && c.f !== undefined && c.f !== deff) p.font = { ascii: f, hAnsi: f, cs: f };
    const col = c.cf !== undefined ? colors[c.cf] : undefined;
    if (col) p.color = col;
    const hl = c.hl !== undefined ? colors[c.hl] : undefined;
    if (hl && HIGHLIGHTS[hl]) p.highlight = HIGHLIGHTS[hl];
    if (c.va) p.vertAlign = c.va;
    if (c.caps) p.caps = true;
    if (c.scaps) p.smallCaps = true;
    if (c.v) p.hidden = true;
    return p;
  };
  const sameProps = (a?: RunProps, b?: RunProps) => JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
  const addText = (text: string) => {
    if (!text) return;
    const props = runProps(g.c);
    const last = inlines[inlines.length - 1];
    if (last?.type === "text" && sameProps(last.props, props)) last.text += text;
    else inlines.push({ type: "text", text, ...(Object.keys(props).length ? { props } : {}) });
  };
  const addInline = (i: Inline) => inlines.push(i);

  const paraProps = (): ParagraphProps => {
    const p: ParagraphProps = {};
    if (para.align) p.align = para.align;
    if (para.li || para.ri || para.fi) {
      p.indent = {};
      if (para.li) p.indent.left = para.li / 20;
      if (para.ri) p.indent.right = para.ri / 20;
      if (para.fi && para.fi > 0) p.indent.firstLine = para.fi / 20;
      if (para.fi && para.fi < 0) p.indent.hanging = -para.fi / 20;
    }
    if (para.sb || para.sa || para.sl) {
      p.spacing = {};
      if (para.sb) p.spacing.before = para.sb / 20;
      if (para.sa) p.spacing.after = para.sa / 20;
      if (para.sl) {
        if (para.slmult && para.sl > 0) { p.spacing.line = para.sl; p.spacing.lineRule = "auto"; }
        else { p.spacing.line = Math.abs(para.sl); p.spacing.lineRule = para.sl < 0 ? "exact" : "atLeast"; }
      }
    }
    if (para.keepn) p.keepNext = true;
    const name = para.style !== undefined ? styleNames.get(para.style) : undefined;
    const h = /^heading ([1-9])$/i.exec(name ?? "");
    const level = h ? Number(h[1]) : para.outline !== undefined && para.outline < 9 ? para.outline + 1 : undefined;
    if (level) p.style = headingStyle(doc, level);
    else if (/^title$/i.test(name ?? "")) p.style = "Title";
    if (para.ls !== undefined && listText) {
      const kind = /\d|^[a-z]\W|^[ivxlc]+\W/i.test(listText.trim()) ? "number" : "bullet";
      p.list = { numId: listId(doc, kind), level: para.ilvl ?? 0 };
      delete p.indent; // the list's own indent
    }
    return p;
  };
  const endParagraph = () => {
    const p = { type: "paragraph" as const, props: paraProps(), content: inlines };
    inlines = [];
    listText = "";
    if (para.intbl) cellBlocks.push(p);
    else {
      if (table) { blocks.push(table); table = undefined; }
      blocks.push(p);
    }
  };
  const endCell = () => {
    if (inlines.length || !cellBlocks.length) {
      cellBlocks.push({ type: "paragraph", props: paraProps(), content: inlines });
      inlines = [];
    }
    rowCells.push({ props: {}, blocks: cellBlocks });
    cellBlocks = [];
  };
  const endRow = () => {
    if (inlines.length || cellBlocks.length) endCell();
    if (!rowCells.length) return;
    const widths = rowWidths.map((x, i) => (x - (rowWidths[i - 1] ?? 0)) / 20);
    if (!table) table = { type: "table", props: { style: "TableGrid" }, grid: widths.length ? widths : rowCells.map(() => 100), rows: [] };
    rowCells.forEach((c, i) => { if (widths[i]) c.props.width = { w: String(Math.round(widths[i] * 20)), type: "dxa" }; });
    table.rows.push({ props: {}, cells: rowCells });
    if (table.grid.length < rowCells.length) table.grid = rowCells.map((_, i) => widths[i] ?? table!.grid[i] ?? 100);
    rowCells = [];
  };

  // Destinations collecting text: the font, colour and style tables, the document info, a field's instruction.
  let destText = "";
  const finishGroup = (closing: Group) => {
    switch (closing.dest) {
      case "fontentry": { const name = destText.replace(/;\s*$/, "").trim(); if (closing.c.f !== undefined && name) fonts.set(closing.c.f, name); break; }
      case "styleentry": { const name = destText.replace(/;\s*$/, "").trim(); styleNames.set((closing as Group & { s?: number }).s ?? 0, name); break; }
      case "title": case "author": case "subject": case "keywords": info[closing.dest] = destText.trim(); break;
      case "fldinst": { const f = fields[fields.length - 1]; if (f) f.instr = destText.trim(); break; }
      case "listtext": listText = destText; break;
      case "pict": {
        if (pict?.type && pict.hex.length) {
          const hex = pict.hex.join("").replace(/[^0-9a-f]/gi, "");
          const data = new Uint8Array(hex.length >> 1);
          for (let i = 0; i < data.length; i++) data[i] = parseInt(hex.substr(i * 2, 2), 16);
          const id = addImage(doc, data, pict.type);
          const w = (pict.w ?? pict.picw ?? 1440) / 20 * ((pict.sx ?? 100) / 100);
          const h = (pict.h ?? pict.pich ?? 1440) / 20 * ((pict.sy ?? 100) / 100);
          addInline({ type: "image", imageId: id, width: w, height: h });
        } else if (pict) warnings.add("Some pictures (other than PNG or JPEG) were left out.");
        pict = undefined;
        break;
      }
      case "fldrslt": {
        // A link's result becomes the link.
        const f = fields[fields.length - 1];
        const m = f && /^HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(f.instr);
        if (f && m) {
          const content = inlines.splice(f.start);
          const anchor = /\\l/.test(f.instr);
          inlines.push({ type: "link", ...(anchor ? { anchor: m[1] } : { href: m[1] }), content });
        }
        break;
      }
      case "field": fields.pop(); break;
      default: break;
    }
    if (["fontentry", "styleentry", "title", "author", "subject", "keywords", "fldinst", "listtext"].includes(closing.dest ?? "")) destText = "";
  };

  const n = src.length;
  let i = 0;
  let text = "";
  const flush = () => {
    if (!text) return;
    if (g.dest && ["fontentry", "styleentry", "title", "author", "subject", "keywords", "fldinst", "listtext"].includes(g.dest)) destText += text;
    else if (g.dest === "pict") pict?.hex.push(text);
    else if (!g.skip) addText(text);
    text = "";
  };
  const emit = (s: string) => {
    if (pendingSkip > 0) { pendingSkip--; return; }
    text += s;
  };

  while (i < n) {
    const ch = src[i];
    if (ch === "{") {
      flush();
      stack.push(g);
      g = { c: { ...g.c }, skip: g.skip, dest: g.dest === "fonttbl" ? "fontentry" : g.dest === "stylesheet" ? "styleentry" : g.dest, uc: g.uc };
      i++;
      continue;
    }
    if (ch === "}") {
      flush();
      const closing = g;
      g = stack.pop() ?? { c: {}, skip: false, uc: 1 };
      if (closing.dest !== g.dest || closing.dest === "fontentry" || closing.dest === "styleentry") finishGroup(closing);
      i++;
      continue;
    }
    if (ch === "\\") {
      const next = src[i + 1];
      if (next === "\\" || next === "{" || next === "}") { emit(next); i += 2; continue; }
      if (next === "'") {
        const byte = parseInt(src.substr(i + 2, 2), 16);
        if (pendingSkip > 0) pendingSkip--;
        else if (!Number.isNaN(byte)) text += decoder.decode(new Uint8Array([byte]));
        i += 4;
        continue;
      }
      if (next === "~") { emit("\u00A0"); i += 2; continue; }
      if (next === "-") { emit("\u00AD"); i += 2; continue; }
      if (next === "_") { emit("\u2011"); i += 2; continue; }
      if (next === "*") { flush(); g.skip = true; g.dest = g.dest ?? "unknown"; i += 2; continue; }
      if (next === "\r" || next === "\n") { flush(); if (!g.skip && !g.dest) endParagraph(); i += 2; continue; }
      const m = /^([a-zA-Z]+)(-?\d+)? ?/.exec(src.slice(i + 1, i + 40));
      if (!m) { i++; continue; }
      i += 1 + m[0].length;
      const word = m[1];
      const arg = m[2] !== undefined ? Number(m[2]) : undefined;
      if (word === "bin" && arg) { i += arg; continue; }
      flush();
      // Destinations.
      if (SKIP.has(word) || word === "fonttbl" || word === "colortbl" || word === "stylesheet") {
        if (word === "fonttbl" || word === "colortbl" || word === "stylesheet") { g.dest = word; g.skip = true; continue; }
        // (Not the copy for older readers Word writes beside each picture.)
        if (word === "pict" && (g.dest === "nonshppict" || stack.some((x) => x.dest === "nonshppict"))) { g.skip = true; continue; }
        if (word === "pict") { g.dest = "pict"; g.skip = true; pict = { hex: [] }; continue; }
        if (["title", "author", "subject", "keywords"].includes(word) && stack.some((x) => x.dest === "info")) { g.dest = word; g.skip = true; continue; }
        if (word === "info") { g.dest = "info"; g.skip = true; continue; }
        if (word === "fldinst") { g.dest = "fldinst"; g.skip = true; continue; }
        if (word === "listtext" || word === "pntext") { g.dest = "listtext"; g.skip = true; continue; }
        if (["header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf"].includes(word)) warnings.add("Headers and footers were left out.");
        if (word === "footnote") warnings.add("Footnotes were left out.");
        g.dest = word;
        g.skip = true;
        continue;
      }
      if (word === "field") { g.dest = "field"; fields.push({ instr: "", start: inlines.length }); continue; }
      if (word === "fldrslt") { g.dest = "fldrslt"; g.skip = false; const f = fields[fields.length - 1]; if (f) f.start = inlines.length; continue; }
      if (g.dest === "colortbl") {
        const cur = (colors as (string | undefined)[] & { cur?: number[] });
        cur.cur ??= [0, 0, 0];
        if (word === "red") cur.cur[0] = arg ?? 0;
        if (word === "green") cur.cur[1] = arg ?? 0;
        if (word === "blue") cur.cur[2] = arg ?? 0;
        continue;
      }
      if (g.dest === "pict") {
        if (word === "pngblip") pict!.type = "image/png";
        else if (word === "jpegblip") pict!.type = "image/jpeg";
        else if (word === "picwgoal") pict!.w = arg;
        else if (word === "pichgoal") pict!.h = arg;
        else if (word === "picw") pict!.picw = arg;
        else if (word === "pich") pict!.pich = arg;
        else if (word === "picscalex") pict!.sx = arg;
        else if (word === "picscaley") pict!.sy = arg;
        continue;
      }
      if (g.dest === "styleentry" && word === "s") { (g as Group & { s?: number }).s = arg ?? 0; continue; }
      if (g.dest === "styleentry" && (word === "cs" || word === "ds" || word === "ts")) { (g as Group & { s?: number }).s = -1; continue; }
      // Paragraph and table settings count only in the body (not inside a list label, the style table…).
      if (g.skip && g.dest !== "fldrslt" && PARA_WORDS.has(word)) continue;
      switch (word) {
        // Document.
        case "ansicpg": codepage = arg ?? 1252; try { decoder = new TextDecoder(`windows-${codepage}`); } catch { decoder = new TextDecoder("windows-1252"); } break;
        case "deff": deff = arg ?? 0; break;
        case "paperw": doc.sections[0].props.pageWidth = (arg ?? 12240) / 20; break;
        case "paperh": doc.sections[0].props.pageHeight = (arg ?? 15840) / 20; break;
        case "margl": doc.sections[0].props.margins = { ...doc.sections[0].props.margins, left: (arg ?? 1440) / 20 }; break;
        case "margr": doc.sections[0].props.margins = { ...doc.sections[0].props.margins, right: (arg ?? 1440) / 20 }; break;
        case "margt": doc.sections[0].props.margins = { ...doc.sections[0].props.margins, top: (arg ?? 1440) / 20 }; break;
        case "margb": doc.sections[0].props.margins = { ...doc.sections[0].props.margins, bottom: (arg ?? 1440) / 20 }; break;
        case "landscape": doc.sections[0].props.orientation = "landscape"; break;
        case "uc": g.uc = arg ?? 1; break;
        case "u": { const code = arg! < 0 ? arg! + 65536 : arg!; emit(String.fromCharCode(code)); pendingSkip = g.uc; break; }
        // Characters.
        case "par": case "sect": if (!g.skip || g.dest === "fldrslt") endParagraph(); break;
        case "line": if (!g.skip) addInline({ type: "break", kind: "line" }); break;
        case "page": if (!g.skip) addInline({ type: "break", kind: "page" }); break;
        case "column": if (!g.skip) addInline({ type: "break", kind: "column" }); break;
        case "tab": if (!g.skip) addInline({ type: "tab" }); else if (g.dest === "listtext") destText += "\t"; break;
        case "cell": case "nestcell": if (!g.skip) endCell(); break;
        case "row": case "nestrow": if (!g.skip) endRow(); break;
        // Character formatting.
        case "plain": g.c = {}; break;
        case "b": g.c.b = arg !== 0; break;
        case "i": g.c.i = arg !== 0; break;
        case "ul": g.c.ul = arg === 0 ? undefined : "single"; break;
        case "uld": g.c.ul = "dotted"; break;
        case "uldb": g.c.ul = "double"; break;
        case "uldash": g.c.ul = "dash"; break;
        case "ulw": g.c.ul = "words"; break;
        case "ulwave": g.c.ul = "wave"; break;
        case "ulth": g.c.ul = "thick"; break;
        case "ulnone": g.c.ul = undefined; break;
        case "strike": g.c.strike = arg !== 0; break;
        case "fs": g.c.fs = arg; break;
        case "f": g.c.f = arg; break;
        case "cf": g.c.cf = arg; break;
        case "highlight": case "cb": case "chcbpat": g.c.hl = arg; break;
        case "super": g.c.va = "superscript"; break;
        case "sub": g.c.va = "subscript"; break;
        case "nosupersub": g.c.va = undefined; break;
        case "caps": g.c.caps = arg !== 0; break;
        case "scaps": g.c.scaps = arg !== 0; break;
        case "v": g.c.v = arg !== 0; break;
        // Paragraph formatting.
        case "pard": para = { intbl: false }; break;
        case "ql": para.align = undefined; break;
        case "qc": para.align = "center"; break;
        case "qr": para.align = "right"; break;
        case "qj": para.align = "both"; break;
        case "li": case "lin": para.li = arg; break;
        case "ri": case "rin": para.ri = arg; break;
        case "fi": para.fi = arg; break;
        case "sb": para.sb = arg; break;
        case "sa": para.sa = arg; break;
        case "sl": para.sl = arg; break;
        case "slmult": para.slmult = arg === 1; break;
        case "keepn": para.keepn = true; break;
        case "s": para.style = arg; break;
        case "outlinelevel": para.outline = arg; break;
        case "intbl": para.intbl = true; break;
        case "ls": para.ls = arg; break;
        case "ilvl": para.ilvl = arg; break;
        // Tables.
        case "trowd": rowWidths = []; break;
        case "cellx": rowWidths.push(arg ?? 0); break;
        default: if (SPECIAL[word] && !g.skip) emit(SPECIAL[word]); break;
      }
      continue;
    }
    if (ch === "\r" || ch === "\n") { i++; continue; }
    if (g.dest === "colortbl" && ch === ";") {
      const cur = (colors as (string | undefined)[] & { cur?: number[] }).cur;
      colors.push(cur ? cur.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase() : undefined);
      (colors as (string | undefined)[] & { cur?: number[] }).cur = undefined;
      i++;
      continue;
    }
    emit(ch);
    i++;
  }
  flush();
  if (inlines.length) endParagraph();
  if (rowCells.length || cellBlocks.length) endRow();
  if (table) blocks.push(table);
  doc.sections[0].blocks = blocks.length ? blocks : [{ type: "paragraph", props: {}, content: [] }];
  const deffName = fonts.get(deff);
  if (deffName) doc.styles.defaults.run = { ...doc.styles.defaults.run, font: { ascii: deffName, hAnsi: deffName, cs: deffName, eastAsia: deffName } };
  const titled = info.title || info.author || info.subject || info.keywords
    ? withProperties(doc, { ...(info.title ? { title: info.title } : {}), ...(info.author ? { author: info.author } : {}), ...(info.subject ? { subject: info.subject } : {}), ...(info.keywords ? { keywords: info.keywords } : {}) })
    : doc;
  return { doc: titled, warnings: [...warnings] };
}

// --- writing -----------------------------------------------------------------------------------------

/** Text for RTF: control characters escaped, anything beyond ASCII as \uN?. */
function esc(s: string): string {
  let out = "";
  for (const ch of s) {
    const c = ch.codePointAt(0)!;
    if (ch === "\\" || ch === "{" || ch === "}") out += `\\${ch}`;
    else if (c === 0xa0) out += "\\~";
    else if (c === 0xad) out += "\\-";
    else if (c < 0x80) out += ch;
    else if (c > 0xffff) {
      const hi = Math.floor((c - 0x10000) / 0x400) + 0xd800;
      const lo = ((c - 0x10000) % 0x400) + 0xdc00;
      out += `\\u${hi - 65536}?\\u${lo - 65536}?`;
    } else out += `\\u${c > 32767 ? c - 65536 : c}?`;
  }
  return out;
}

const tw = (pt: number) => Math.round(pt * 20);
const HL_COLORS: Record<string, string> = Object.fromEntries(Object.entries(HIGHLIGHTS).map(([k, v]) => [v, k]));

export function writeRtf(doc: PrimeDocument, props: { title?: string; author?: string } = {}): Uint8Array {
  const body = shownBody(doc);
  const fonts: string[] = [];
  const colors: string[] = [];
  const font = (name: string) => { let i = fonts.indexOf(name); if (i < 0) { fonts.push(name); i = fonts.length - 1; } return i; };
  const color = (hex: string) => { const h = hex.toUpperCase(); let i = colors.indexOf(h); if (i < 0) { colors.push(h); i = colors.length - 1; } return i + 1; };
  font("Calibri");

  const runCodes = (r: Piece["run"]) => {
    let s = `\\f${font(r.font)}\\fs${Math.round(r.size * 2)}`;
    if (r.bold) s += "\\b";
    if (r.italic) s += "\\i";
    if (r.underline) s += r.underlineStyle === "double" ? "\\uldb" : r.underlineStyle === "dotted" ? "\\uld" : r.underlineStyle === "wave" ? "\\ulwave" : "\\ul";
    if (r.strike) s += "\\strike";
    if (r.caps) s += "\\caps";
    if (r.hidden) s += "\\v";
    if (r.vertAlign === "superscript") s += "\\super";
    if (r.vertAlign === "subscript") s += "\\sub";
    if (r.color && r.color !== "auto" && r.color !== "000000") s += `\\cf${color(r.color)}`;
    if (r.highlight && HL_COLORS[r.highlight]) s += `\\highlight${color(HL_COLORS[r.highlight])}`;
    return s;
  };
  const pieces = (sp: ShownParagraph): string => {
    let out = "";
    if (sp.p.listLabel) {
      const r = sp.p.labelRun ?? sp.p.mark;
      out += `{${runCodes(r)} ${esc(sp.p.listLabel)}\\tab}`;
    }
    for (const piece of sp.pieces) {
      const codes = runCodes(piece.run);
      if (piece.kind === "text") {
        const t = `{${codes} ${esc(piece.text)}}`;
        out += piece.link ? `{\\field{\\*\\fldinst{HYPERLINK ${piece.link.startsWith("#") ? `\\\\l "${esc(piece.link.slice(1))}"` : `"${esc(piece.link)}"`}}}{\\fldrslt{${t}}}}` : t;
      } else if (piece.kind === "tab") out += `{${codes}\\tab}`;
      else if (piece.kind === "line") out += "\\line ";
      else if (piece.kind === "page") out += "\\page ";
      else if (piece.kind === "column") out += "\\column ";
      else if (piece.kind === "image") {
        const img = imageBytes(doc, piece.imageId);
        if (!img || !/png|jpe?g/.test(img.contentType)) continue;
        let hex = "";
        for (let i = 0; i < img.data.length; i++) { hex += img.data[i].toString(16).padStart(2, "0"); if (i % 64 === 63) hex += "\n"; }
        out += `{\\pict${img.contentType.includes("png") ? "\\pngblip" : "\\jpegblip"}\\picwgoal${tw(piece.width)}\\pichgoal${tw(piece.height)}\n${hex}}`;
      }
    }
    return out;
  };
  const paraCodes = (sp: ShownParagraph, inTable: boolean) => {
    const p = sp.p;
    let s = "\\pard\\plain";
    if (inTable) s += "\\intbl";
    s += { center: "\\qc", right: "\\qr", both: "\\qj", justify: "\\qj" }[p.align] ?? "\\ql";
    if (p.left) s += `\\li${tw(p.left)}`;
    if (p.right) s += `\\ri${tw(p.right)}`;
    if (p.firstLine) s += `\\fi${tw(p.firstLine)}`;
    if (p.spaceBefore) s += `\\sb${tw(p.spaceBefore)}`;
    if (p.spaceAfter) s += `\\sa${tw(p.spaceAfter)}`;
    if (p.lineRule === "auto" && p.line && p.line !== 240) s += `\\sl${Math.round(p.line)}\\slmult1`;
    else if (p.lineRule === "exact" && p.line) s += `\\sl-${Math.round(p.line)}\\slmult0`;
    else if (p.lineRule === "atLeast" && p.line) s += `\\sl${Math.round(p.line)}\\slmult0`;
    if (p.keepNext) s += "\\keepn";
    if (p.pageBreakBefore) s += "\\pagebb";
    if (sp.heading) s += `\\outlinelevel${sp.heading - 1}`;
    return `${s} `;
  };
  const out: string[] = [];
  const blocks = (list: Shown[], inTable: boolean) => {
    list.forEach((b, i) => {
      if (b.kind === "paragraph") {
        const last = inTable && i === list.length - 1;
        out.push(`${paraCodes(b, inTable)}${pieces(b)}${last ? "\\cell" : "\\par"}\n`);
      } else {
        // A table in a table cell: its text as paragraphs of the cell.
        if (inTable) { for (const row of b.rows) for (const c of row.cells) blocks(c.blocks.map((x) => x), false); return; }
        for (const row of b.rows) {
          let x = 0;
          let col = 0;
          let defs = "\\trowd\\trgaph108\\trleft-108";
          for (const c of row.cells) {
            const w = b.grid.slice(col, col + c.span).reduce((a, v) => a + v, 0) || 100;
            col += c.span;
            x += w;
            defs += `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10${c.shading ? `\\clcbpat${color(c.shading)}` : ""}\\cellx${tw(x)}`;
          }
          out.push(`${defs}\n`);
          for (const c of row.cells) {
            const paras = c.blocks.filter((x): x is ShownParagraph => x.kind === "paragraph");
            if (!paras.length) out.push("\\pard\\intbl\\cell\n");
            else blocks(paras, true);
          }
          out.push(`${defs}\\row\n`);
        }
      }
    });
  };
  blocks(body, false);
  const s = doc.sections[0].props;
  const m = s.margins ?? {};
  const page = `\\paperw${tw(s.pageWidth ?? 612)}\\paperh${tw(s.pageHeight ?? 792)}\\margl${tw(m.left ?? 72)}\\margr${tw(m.right ?? 72)}\\margt${tw(m.top ?? 72)}\\margb${tw(m.bottom ?? 72)}${s.orientation === "landscape" ? "\\landscape" : ""}`;
  const infoGroup = props.title || props.author ? `{\\info${props.title ? `{\\title ${esc(props.title)}}` : ""}${props.author ? `{\\author ${esc(props.author)}}` : ""}}` : "";
  const head = `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1{\\fonttbl${fonts.map((f, i) => `{\\f${i}\\fnil ${esc(f)};}`).join("")}}{\\colortbl;${colors.map((c) => `\\red${parseInt(c.slice(0, 2), 16)}\\green${parseInt(c.slice(2, 4), 16)}\\blue${parseInt(c.slice(4, 6), 16)};`).join("")}}${infoGroup}${page}\n`;
  return new TextEncoder().encode(`${head}${out.join("")}}`);
}
