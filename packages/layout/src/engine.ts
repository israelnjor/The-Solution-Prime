/**
 * Layout engine v0 — Prime Document Model -> pages of positioned lines (ADR-007).
 *
 * Scope of v0: paragraphs (line breaking, indents, spacing, line rules, borders), lists, manual
 * page breaks, keep-with-next, keep-lines-together, widow/orphan control, inline images, tables
 * (rows, spans, repeating header rows), sections with their own page size and margins.
 * Also: horizontal alignment and justification, headers/footers on every page, PAGE/NUMPAGES
 * fields, picture placement, table cell borders and shading — everything a renderer needs.
 * Not yet: custom tab stops, floating objects, footnotes, columns, hyphenation.
 */
import {
  type Block, type FloatInfo, type Inline, type Paragraph, type PrimeDocument, type Relationship, type Section, type Story, type Table,
  paragraphText,
} from "@prime/pdm";
import { MATH_FONT, type MathBox, MathTypesetter } from "./math.js";
import type { FontKey, Measurer } from "./measure.js";
import { type ResolvedParagraph, type ResolvedRun, Resolver } from "./resolve.js";

const MARKER = /\[[A-Z]-\d{3}\]/g;
const DEFAULT_TAB = 36;
const CELL_MARGIN = 5.4;
const ROW_BORDER = 0.5;

export interface Segment {
  x: number;
  text: string;
  font: FontKey;
  size: number;
  color?: string;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  vertAlign?: string;
  /** Text computed at render time (its cached text is only a placeholder). */
  field?: "PAGE" | "NUMPAGES";
  /** Vertical offset from the line's baseline (equations: scripts, fractions). */
  dy?: number;
}

/** A straight line drawn with the text (fraction bars, overlines, radicals). */
export interface Rule {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
}

/** Where a picture's bytes live: a modeled image resource, or a package part (preserved drawings). */
export interface ImageSource {
  imageId?: string;
  part?: string;
}

export interface PlacedImage extends ImageSource {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
  fill?: string;
  /** Border sides to draw (table cell borders, paragraph borders). */
  sides?: { top: boolean; bottom: boolean; left: boolean; right: boolean };
  /** Border colour and width (defaults: black, 0.5pt). */
  stroke?: { color?: string; width: number };
}

export interface PositionedLine {
  /** Absolute position on the page, in pt from the top-left corner. */
  x: number;
  y: number;
  height: number;
  baseline: number;
  segments: Segment[];
  images: PlacedImage[];
  rules: Rule[];
  /** Set for header and footer lines, and for floating pictures and text boxes. */
  story?: "header" | "footer" | "float";
  /** Pen strokes (ink): x, y pairs on the page, pen width in pt, colour. */
  strokes?: { points: number[]; width: number; color: string }[];
}

export interface Page {
  number: number;
  section: number;
  width: number;
  height: number;
  lines: PositionedLine[];
  /** Cell shading and borders, drawn beneath the text. */
  rects: Rect[];
  /** Top of the body text area (below the header). */
  bodyTop: number;
  /** Bottom of the body text area (above the footer). */
  bodyBottom: number;
  /** Left and right edges of the body text area. */
  bodyLeft: number;
  bodyRight: number;
}

export interface LayoutResult {
  pages: Page[];
  /** Page and top-of-first-line position of every paragraph containing a content marker. */
  markers: Record<string, { page: number; y: number }>;
  /** Every paragraph in reading order (body and table cells): text, page, top of first line. */
  paragraphs: { text: string; page: number; y: number }[];
  /** Where pages 2, 3, ... begin, in document terms — for drawing page breaks in the editor. */
  pageStarts: PageStart[];
  /** Where text wraps around floating objects, in document terms — so the editor can do the same. */
  wraps: WrapPoint[];
  /** Floating objects relative to their anchor paragraph — so the editor can draw them with its text. */
  floats: PlacedFloat[];
}

/**
 * A floating object as placed: its box relative to the anchor paragraph (x from the column's left
 * edge, y from the top of the paragraph's first line), and what it shows — a picture, or text lines
 * positioned relative to the box.
 */
export interface PlacedFloat {
  page: number;
  path: number[];
  /** Which of the paragraph's floating objects (in document order). */
  index: number;
  x: number;
  y: number;
  width: number;
  height: number;
  image?: ImageSource;
  lines: PositionedLine[];
  /** Ink: strokes relative to the box (0–1 each way). */
  ink?: FloatInfo["ink"];
}

/**
 * A line where wrapping around floating objects starts (or changes): the line moves down by `skip`,
 * then runs beside space taken on the left and/or right (sizes in pt, heights from the line's top
 * after the skip). Body paragraphs only: path and offset as in PageStart.
 */
export interface WrapPoint {
  page: number;
  path: number[];
  offset: number;
  skip: number;
  left?: { width: number; height: number };
  right?: { width: number; height: number };
}

/**
 * The first content of a page: a line of a paragraph (path + offset) or a table row (path + row).
 * Paths index the editor's document: top-level blocks with one section break between sections,
 * then children of content controls.
 */
export interface PageStart {
  page: number;
  path: number[];
  offset?: number;
  row?: number;
}

// --- measured structures -------------------------------------------------------------------

interface MLine {
  width: number;
  /** Width available to the line (for alignment). */
  avail: number;
  /** Ends with a manual line or page break (such lines are not justified). */
  endsWithBreak: boolean;
  images: (ImageSource & { x: number; width: number; height: number })[];
  /** Rules relative to the line's left edge and baseline. */
  rules: Rule[];
  height: number;
  ascent: number;
  segments: Segment[];
  startX: number;
  pageBreakAfter: boolean;
  /** Height that must fit on the page: for a picture line, just the picture (the rest may spill). */
  fitHeight: number;
  /** Offset in the paragraph (editor position units) where the line starts. */
  startOffset: number;
  /** Where wrapping around floating objects starts or changes on this line (see WrapPoint). */
  wrap?: Omit<WrapPoint, "page" | "path" | "offset">;
}

interface ParaBox {
  kind: "para";
  /** Index path of the block: top-level index (sections and their breaks counted), then children. */
  path: number[];
  rp: ResolvedParagraph;
  lines: MLine[];
  /** Plain text of the paragraph, used to identify it in results. */
  anchor: string;
  /** The paragraph itself (re-broken around floating objects), and the floating objects anchored in it. */
  para: Paragraph;
  floats: FloatInfo[];
}

/**
 * Where a line may go, for wrapping around floating objects: given the line's top (relative to the
 * paragraph's first line) and estimated height, the free span of the column, or where to move down to.
 */
type Slot = (dy: number, height: number) =>
  | { left: number; right: number; leftUntil: number; rightUntil: number }
  | { skipTo: number };

/** Narrowest gap beside a floating object that text flows into (Word leaves narrower gaps empty). */
const MIN_WRAP = 36;

/** Floating objects in a paragraph's content (inside links, fields and other containers too). */
function floatsIn(items: Inline[]): FloatInfo[] {
  return items.flatMap((i) => i.type === "runPreserved" ? (i.float ? [i.float] : [])
    : "content" in i && Array.isArray(i.content) ? floatsIn(i.content as Inline[]) : []);
}

interface CellBox {
  dx: number;
  /** Cell width including its margins. */
  width: number;
  fill?: string;
  borders: boolean;
  /** Vertical merge role, to omit the borders inside a merged cell. */
  vMerge?: "restart" | "continue";
  /** Content height of the cell (for merged cells: across all the rows it spans). */
  height: number;
  lines: { line: MLine; dy: number }[];
  /** Paragraph anchors (text) with their offset from the cell top. */
  markers: [string, number][];
  /** Floating objects anchored in the cell's paragraphs: their paragraph's top, path in the cell, index. */
  floats?: { float: FloatInfo; dy: number; path: number[]; index: number }[];
}

interface RowBox {
  index: number;
  height: number;
  header: boolean;
  cantSplit: boolean;
  cells: CellBox[];
}

interface TableBox {
  kind: "table";
  path: number[];
  rows: RowBox[];
}

type Box = ParaBox | TableBox;

/** Stacked content of a cell or a header/footer. */
type StoryBox = Pick<CellBox, "height" | "lines" | "markers" | "floats">;

type Atom = { offset: number } & (
  // `breakAfter`: the line may break after this piece although no space follows (after a hyphen).
  | { kind: "word"; text: string; run: ResolvedRun; field?: Segment["field"]; breakAfter?: boolean }
  | { kind: "space"; text: string; run: ResolvedRun }
  | { kind: "tab"; run: ResolvedRun }
  | { kind: "br" }
  | { kind: "page" }
  | { kind: "image"; width: number; height: number; src: ImageSource }
  | { kind: "math"; box: MathBox; display: boolean }
);

interface Face {
  font: FontKey;
  size: number;
  /** Drawing characters the run's own font lacks (Word spaces such lines more tightly). */
  fallback?: boolean;
  /** A list label's font: counts above the baseline only (Word measures its descent from the text). */
  label?: boolean;
}

const key = (r: ResolvedRun): FontKey => ({ family: r.font, bold: r.bold, italic: r.italic });

/**
 * Characters of the symbol fonts, as Word writes them (U+F0xx, or 0xxx), with their Unicode
 * equivalents: the bullets Word offers, mostly. Symbol fonts are drawn badly by PDF libraries (their
 * character tables are not Unicode), so these are drawn as the Unicode character in a regular font.
 */
const SYMBOL_CHARS: Record<string, Record<number, string>> = {
  symbol: { 0xb7: "•", 0x2d: "−", 0xae: "→", 0xac: "←", 0xad: "↑", 0xaf: "↓", 0xd7: "·", 0xa8: "♦", 0xa9: "♥", 0xaa: "♠", 0xa7: "♣", 0xb0: "°", 0xb1: "±", 0xb4: "×", 0xb8: "÷", 0xb9: "≠", 0xa3: "≤", 0xb3: "≥", 0xa5: "∞", 0xd6: "√", 0xe0: "◊" },
  wingdings: { 0x6c: "●", 0x6e: "■", 0x6f: "□", 0x71: "❑", 0x75: "◆", 0x76: "❖", 0x77: "⬥", 0xa7: "▪", 0xa8: "◻", 0x9f: "•", 0xd8: "➢", 0xe0: "➔", 0xe8: "➔", 0xfc: "✓", 0xfb: "✗", 0xfe: "☑", 0xfd: "☒", 0x4a: "☺", 0x4c: "☹", 0xab: "★" },
};
const SYMBOL_FALLBACK = "Segoe UI Symbol";

/** A symbol-font run's text as Unicode in a regular font, when every character has an equivalent. */
function unsymbol(text: string, run: ResolvedRun): { text: string; run: ResolvedRun; dy?: number } | undefined {
  const table = SYMBOL_CHARS[run.font.toLowerCase()];
  if (!table) return undefined;
  let out = "";
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    const u = table[cp >= 0xf000 ? cp - 0xf000 : cp];
    if (!u) return undefined;
    out += u;
  }
  // Word's Symbol bullet is about 0.42em across, Arial's 0.25em (measured in Word's PDF output):
  // the bullet is drawn larger to match. Other characters come from Segoe UI Symbol.
  // Enlarged, it would sit higher than Word's (its centre is 0.35em up): lowered back by `dy`.
  if (out === "•") return { text: out, run: { ...run, font: "Arial", size: run.size * 1.7 }, dy: run.size * 0.35 * 0.7 };
  return { text: out, run: { ...run, font: SYMBOL_FALLBACK } };
}

/** Segment styling taken from a run. */
const look = (r: ResolvedRun): Partial<Segment> => {
  const out: Partial<Segment> = {};
  if (r.color) out.color = r.color;
  if (r.underline) out.underline = true;
  if (r.strike) out.strike = true;
  if (r.highlight) out.highlight = r.highlight;
  if (r.vertAlign) out.vertAlign = r.vertAlign;
  return out;
};

const RENDER_FIELDS = new Set(["PAGE", "NUMPAGES"]);

export class LayoutEngine {
  private resolver: Resolver;
  private math: MathTypesetter;
  /** Relationships of the story being measured (main document or a header/footer). */
  private rels: Relationship[];
  /** Word compatibility mode of the document (12 = 2007, 14 = 2010, 15 = 2013 and later). */
  private compat: number;

  constructor(private doc: PrimeDocument, private m: Measurer) {
    this.resolver = new Resolver(doc);
    this.math = new MathTypesetter(m);
    this.rels = doc.package.preservedRels;
    const settings = doc.package.parts.find((p) => /word\/settings\.xml$/.test(p.name));
    const xml = settings ? new TextDecoder().decode(Uint8Array.from(atob(settings.data), (c) => c.charCodeAt(0))) : "";
    this.compat = Number(/compatibilityMode" w:uri="[^"]*" w:val="(\d+)"/.exec(xml)?.[1] ?? 12);
  }

  // --- paragraph measurement -----------------------------------------------------------------

  /**
   * The paragraph as a sequence of atoms. Each atom carries its offset in the paragraph, counted the
   * way the editor counts positions: one per character of text, one per inline object, nothing for
   * containers (links, fields, content controls, tracked changes).
   */
  private atoms(p: Paragraph, rp: ResolvedParagraph): Atom[] {
    const out: Atom[] = [];
    let pos = 0;
    const size = (i: Inline): number =>
      i.type === "text" ? i.text.length
        : i.type === "link" || i.type === "field" || i.type === "contentControl" || i.type === "revision"
          ? i.content.reduce((n, c) => n + size(c), 0) : 1;
    // Open fields: their instruction, and whether we are in the displayed result yet.
    const fields: { instr: string; result: boolean }[] = [];
    const fieldKind = (): Segment["field"] => {
      const f = fields[fields.length - 1];
      const kind = f?.instr.trim().split(/\s+/)[0]?.toUpperCase();
      return kind && RENDER_FIELDS.has(kind) ? (kind as Segment["field"]) : undefined;
    };
    const visit = (items: Inline[]) => {
      for (const i of items) {
        const offset = pos;
        if (i.type === "fieldChar") {
          if (i.kind === "begin") fields.push({ instr: "", result: false });
          else if (i.kind === "separate" && fields.length) fields[fields.length - 1].result = true;
          else if (i.kind === "end") fields.pop();
          pos += 1;
          continue;
        }
        if (i.type === "instr") {
          if (fields.length && !fields[fields.length - 1].result) fields[fields.length - 1].instr += i.text;
          pos += 1;
          continue;
        }
        if (fields.some((f) => !f.result)) { pos += size(i); continue; }
        if (i.type === "field") {
          fields.push({ instr: i.instr, result: true });
          visit(i.content);
          fields.pop();
          continue;
        }
        // Tracked deletions are laid out too, as Word's markup view shows them (struck through).
        if (i.type === "link" || i.type === "contentControl" || i.type === "revision") {
          visit(i.content);
          continue;
        }
        pos += size(i);
        if (i.type === "image") {
          out.push({ kind: "image", width: i.width, height: i.height, src: { imageId: i.imageId }, offset });
          continue;
        }
        if (i.type === "break") { out.push({ kind: i.kind === "page" ? "page" : "br", offset }); continue; }
        if (i.type === "preserved" && (i.preserved.name === "m:oMath" || i.preserved.name === "m:oMathPara")) {
          const { box, display } = this.math.layout(i.preserved.xml, rp.mark.size);
          out.push({ kind: "math", box, display, offset });
          continue;
        }
        if (i.type === "runPreserved") {
          // Inline pictures we keep verbatim (borders, crops...) still occupy their size on the line.
          // Borders and effects extend the picture's footprint by its effect extent.
          const extent = /<wp:inline[ >][\s\S]*?<wp:extent cx="(\d+)" cy="(\d+)"/.exec(i.preserved.xml);
          const effect = /<wp:effectExtent l="(\d+)" t="(\d+)" r="(\d+)" b="(\d+)"/.exec(i.preserved.xml)?.slice(1).map(Number) ?? [0, 0, 0, 0];
          if (extent) {
            const relId = /r:embed="([^"]+)"/.exec(i.preserved.xml)?.[1];
            const rel = this.rels.find((r) => r.id === relId && !r.external);
            out.push({
              kind: "image",
              width: (Number(extent[1]) + effect[0] + effect[2]) / 12700,
              height: (Number(extent[2]) + effect[1] + effect[3]) / 12700,
              src: rel ? { part: rel.target } : {},
              offset,
            });
          }
          continue;
        }
        if (i.type !== "text" && i.type !== "tab") continue;
        const run = this.resolver.run(rp.styleId, i.props, rp.tableStyle);
        if (run.hidden) continue;
        if (i.type === "tab") { out.push({ kind: "tab", run, offset }); continue; }
        let at = offset;
        for (const piece of i.text.split(/( +)/)) {
          if (!piece) continue;
          const sym = unsymbol(piece, run);
          if (sym) {
            out.push({ kind: "word", text: sym.text, run: sym.run, offset: at });
            at += piece.length;
            continue;
          }
          const text = run.caps ? piece.toUpperCase() : piece;
          const field = fieldKind();
          if (piece.startsWith(" ")) {
            out.push({ kind: "space", text, run, offset: at });
            at += piece.length;
            continue;
          }
          // Word breaks lines after a hyphen inside a word (e.g. in long addresses and compounds).
          let from = 0;
          for (const part of text.split(/(?<=[^\s-]-)(?=[^\s-])/)) {
            const last = from + part.length === text.length;
            out.push({ kind: "word", text: part, run, offset: at + from, ...(field ? { field } : {}), ...(last ? {} : { breakAfter: true }) });
            from += part.length;
          }
          at += piece.length;
        }
      }
    };
    visit(p.content);
    return out;
  }

  private textWidth(text: string, run: ResolvedRun): number {
    const size = run.vertAlign ? run.size * 0.65 : run.size;
    return this.m.width(text, key(run)) * size;
  }

  /** Fonts that actually draw `text` in this run: the run's font plus any fallback fonts. */
  private faces(run: ResolvedRun, text?: string): Face[] {
    const fonts = text && this.m.faces ? this.m.faces(text, key(run)) : [key(run)];
    return fonts.map((font, i) => ({ font, size: run.size, ...(i > 0 ? { fallback: true } : {}) }));
  }

  private lineMetrics(faces: Face[]): { height: number; ascent: number; labelExtra: number } {
    let height = 0;
    let ascent = 0;
    let descent = 0;
    let labelAscent = 0;
    for (const f of faces) {
      const fm = this.m.metrics(f.font, f.fallback);
      if (f.label) { labelAscent = Math.max(labelAscent, fm.ascent * f.size); continue; }
      height = Math.max(height, fm.lineHeight * f.size);
      ascent = Math.max(ascent, fm.ascent * f.size);
      descent = Math.max(descent, (fm.lineHeight - fm.ascent) * f.size);
    }
    // A taller list label (e.g. a Symbol bullet) raises the line by its ascent only: with Times New
    // Roman 12pt text, Word's line is 12.06 (Symbol's ascent) + 2.60 (Times' descent) = 14.66pt.
    let labelExtra = 0;
    if (labelAscent > ascent) {
      const raised = Math.max(height, labelAscent + (descent || height - ascent));
      labelExtra = raised - height;
      height = raised;
      ascent = labelAscent;
    }
    return { height, ascent, labelExtra };
  }

  private applyLineRule(natural: number, rp: ResolvedParagraph): number {
    if (rp.lineRule === "exact") return rp.line / 20;
    if (rp.lineRule === "atLeast") return Math.max(natural, rp.line / 20);
    return natural * (rp.line / 240);
  }

  /** Greedy line breaking at spaces, using real glyph advances. */
  breakLines(p: Paragraph, rp: ResolvedParagraph, width: number, slot?: Slot): MLine[] {
    const atoms = this.atoms(p, rp);
    const lines: MLine[] = [];
    let label = rp.listLabel;
    // Wrapping around floating objects: the line's top so far, and space skipped to get below one.
    let dy = 0;
    let skip = 0;
    let free: { left: number; right: number; leftUntil: number; rightUntil: number } | undefined;
    let lastFree = "";
    const lineGuess = slot ? this.applyLineRule(this.lineMetrics(this.faces(rp.mark)).height, rp) : 0;

    let segs: Segment[] = [];
    let images: MLine["images"] = [];
    let rules: Rule[] = [];
    let mathAscent = 0;
    let mathDescent = 0;
    let runs: Face[] = [];
    let imageHeight = 0;
    let x = 0;
    let pendingSpace: { text: string; run: ResolvedRun; w: number } | undefined;
    let lineStart = 0;
    let startX = 0;
    let avail = 0;

    const begin = () => {
      const first = lines.length === 0;
      startX = rp.left + (first ? rp.firstLine : 0);
      if (first && label !== undefined) {
        let lr = rp.labelRun ?? rp.mark;
        // Line height still comes from the label's own font (as in Word), whatever font draws it.
        runs.push(...this.faces(lr, label).map((f) => ({ ...f, label: true })));
        const sym = unsymbol(label, lr);
        if (sym) ({ text: label, run: lr } = sym);
        const lw = this.textWidth(label, lr);
        segs.push({ x: startX, text: label, font: key(lr), size: lr.size, ...look(lr), ...(sym?.dy ? { dy: sym.dy } : {}) });
        // The label sits in the hanging indent; text starts at the next tab stop (usually the left indent).
        const after = startX + lw;
        startX = rp.firstLine < 0 && after <= rp.left ? rp.left : (Math.floor(after / DEFAULT_TAB) + 1) * DEFAULT_TAB;
        label = undefined;
      }
      avail = width - rp.right - startX;
      x = 0;
      free = undefined;
      if (slot) {
        for (let guard = 0; guard < 20; guard++) {
          const span = slot(dy + skip, lineGuess);
          if ("skipTo" in span) { skip = Math.max(skip, span.skipTo - dy); continue; }
          // Text starts at the indent or the object's edge, whichever is further in.
          startX = Math.max(startX, span.left);
          avail = Math.min(width - rp.right, span.right) - startX;
          free = span;
          break;
        }
      }
    };
    const finish = (pageBreakAfter = false, endsWithBreak = pageBreakAfter, nextStart = lineStart) => {
      const textMetrics = this.lineMetrics(runs.length ? runs : this.faces(rp.mark));
      // Equations extend the line only where they reach beyond the text's own ascent and descent.
      const ascentAll = Math.max(textMetrics.ascent, mathAscent);
      const descentAll = Math.max(textMetrics.height - textMetrics.ascent, mathDescent);
      const metrics = { ascent: ascentAll, height: ascentAll + descentAll };
      // Line spacing does not scale pictures: Word sizes a picture line to the picture, plus the text
      // descent, plus the extra space the line rule adds to text (an "exact" rule still clips it).
      // Line spacing scales the text's height; a taller list label adds its extra unscaled (Word: a
      // Symbol bullet on 1.16-spaced Times New Roman lines adds 0.86pt, not 0.99pt).
      const extra = textMetrics.labelExtra;
      const text = rp.lineRule === "exact" ? rp.line / 20
        : rp.lineRule === "atLeast" ? Math.max(metrics.height, rp.line / 20)
        : this.applyLineRule(metrics.height - extra, rp) + extra;
      const picture = imageHeight && rp.lineRule !== "exact"
        ? imageHeight + (metrics.height - metrics.ascent) + Math.max(text - metrics.height, 0) : 0;
      // Space skipped to get below a floating object belongs to the line (its text sits lower).
      lines.push({
        width: x, height: Math.max(text, picture) + skip, ascent: Math.max(metrics.ascent, imageHeight) + skip,
        fitHeight: (picture > text ? Math.max(imageHeight, text) : Math.max(text, picture)) + skip,
        segments: segs, startX, pageBreakAfter, avail, endsWithBreak, images, rules, startOffset: lineStart,
      });
      // Record where wrapping starts or changes, for the editor (heights from the line's top after the skip).
      const narrowed = free && (free.left > 0.01 || free.right < width - 0.01);
      const signature = narrowed ? `${free!.left}:${free!.right}` : "";
      if (skip > 0 || signature !== lastFree) {
        const top = dy + skip;
        const line = lines[lines.length - 1];
        line.wrap = {
          skip,
          ...(narrowed && free!.left > rp.left + 0.01 ? { left: { width: free!.left - rp.left, height: free!.leftUntil - top } } : {}),
          ...(narrowed && free!.right < width - rp.right - 0.01 ? { right: { width: width - rp.right - free!.right, height: free!.rightUntil - top } } : {}),
        };
      }
      lastFree = signature;
      dy += lines[lines.length - 1].height;
      skip = 0;
      lineStart = nextStart;
      segs = [];
      images = [];
      rules = [];
      mathAscent = 0;
      mathDescent = 0;
      runs = [];
      imageHeight = 0;
      pendingSpace = undefined;
      begin();
    };
    begin();

    for (let i = 0; i < atoms.length; i++) {
      const a = atoms[i];
      if (a.kind === "space") {
        const w = this.textWidth(a.text, a.run);
        pendingSpace = pendingSpace ? { ...pendingSpace, text: pendingSpace.text + a.text, w: pendingSpace.w + w } : { text: a.text, run: a.run, w };
        continue;
      }
      if (a.kind === "br") { finish(false, true, a.offset + 1); continue; }
      if (a.kind === "page") { finish(true, true, a.offset + 1); continue; }
      if (a.kind === "tab") {
        const pos = startX + x + (pendingSpace?.w ?? 0);
        const next = (Math.floor(pos / DEFAULT_TAB + 1e-6) + 1) * DEFAULT_TAB;
        x = next - startX;
        pendingSpace = undefined;
        runs.push(...this.faces(a.run));
        continue;
      }
      // A word may span several runs (e.g. "H₂O"): collect atoms until the next space or break.
      const word: Atom[] = [a];
      const ends = (w: Atom) => w.kind === "word" && w.breakAfter;
      while (!ends(word[word.length - 1]) && ["word", "image", "math"].includes(atoms[i + 1]?.kind)) word.push(atoms[++i]);
      // A word wider than a whole line is split where the line ends (Word does the same).
      if (word.length === 1 && a.kind === "word" && a.text.length > 1) {
        // Measured against the line's full width: such a word starts a new line, then breaks.
        if (this.textWidth(a.text, a.run) > avail + 0.01) {
          let n = 1;
          while (n < a.text.length && this.textWidth(a.text.slice(0, n + 1), a.run) <= avail) n++;
          atoms.splice(i, 1,
            { ...a, text: a.text.slice(0, n), breakAfter: true },
            { ...a, text: a.text.slice(n), offset: a.offset + n });
          i--;
          continue;
        }
      }
      const parts = word.map((w) => (w.kind === "word" ? this.textWidth(w.text, w.run)
        : w.kind === "image" ? w.width : w.kind === "math" ? w.box.width : 0));
      const ww = parts.reduce((s, v) => s + v, 0);
      const spaceW = pendingSpace?.w ?? 0;
      if (x > 0 && x + spaceW + ww > avail) finish(false, false, word[0].offset);
      else if (pendingSpace && x > 0) {
        segs.push({ x: startX + x, text: pendingSpace.text, font: key(pendingSpace.run), size: pendingSpace.run.size, ...look(pendingSpace.run) });
        x += spaceW;
      }
      pendingSpace = undefined;
      word.forEach((w, k) => {
        if (w.kind === "word") {
          segs.push({ x: startX + x, text: w.text, font: key(w.run), size: w.run.size, ...look(w.run), ...(w.field ? { field: w.field } : {}) });
          runs.push(...this.faces(w.run, w.text));
        } else if (w.kind === "image") {
          imageHeight = Math.max(imageHeight, w.height);
          images.push({ x: startX + x, width: w.width, height: w.height, ...w.src });
        } else if (w.kind === "math") {
          const at = startX + x;
          for (const g of w.box.glyphs) segs.push({ x: at + g.x, text: g.text, font: MATH_FONT, size: g.size, dy: g.dy });
          for (const r of w.box.rules) rules.push({ ...r, x1: at + r.x1, x2: at + r.x2 });
          mathAscent = Math.max(mathAscent, w.box.ascent);
          mathDescent = Math.max(mathDescent, w.box.descent);
        }
        x += parts[k];
      });
    }
    // The paragraph mark: a trailing empty line after a page break does not exist in Word.
    const last = lines[lines.length - 1];
    if (!(segs.length === 0 && x === 0 && imageHeight === 0 && last?.pageBreakAfter)) finish();
    // A display equation (m:oMathPara) is centred unless the paragraph says otherwise.
    const display = atoms.some((a) => a.kind === "math" && a.display);
    const align = display && rp.align === "left" ? "center" : rp.align;
    lines.forEach((line, n) => this.align(line, align, n === lines.length - 1));
    return lines;
  }

  /** Horizontal alignment; justified lines spread the slack over their spaces (not the last line). */
  private align(line: MLine, align: string, lastLine: boolean) {
    const slack = line.avail - line.width;
    if (slack <= 0.01) return;
    let shift: (x: number) => number;
    if (align === "center") shift = () => slack / 2;
    else if (align === "right" || align === "end") shift = () => slack;
    else if ((align === "both" || align === "distribute") && !lastLine && !line.endsWithBreak) {
      const spaces = line.segments.filter((s) => /^ +$/.test(s.text));
      const count = spaces.reduce((n, s) => n + s.text.length, 0);
      if (!count) return;
      const per = slack / count;
      // Every segment moves right by the extra space given to the spaces before it.
      shift = (x) => spaces.filter((s) => s.x < x).reduce((n, s) => n + s.text.length, 0) * per;
    } else return;
    line.segments = line.segments.map((s) => ({ ...s, x: s.x + shift(s.x) }));
    line.images = line.images.map((i) => ({ ...i, x: i.x + shift(i.x) }));
    line.rules = line.rules.map((r) => ({ ...r, x1: r.x1 + shift(r.x1), x2: r.x2 + shift(r.x1) }));
  }

  private measureParagraph(p: Paragraph, width: number, tableStyle?: string, path: number[] = []): ParaBox {
    const rp = this.resolver.paragraph(p, tableStyle);
    return { kind: "para", path, rp, lines: this.breakLines(p, rp, width), anchor: paragraphText(p), para: p, floats: floatsIn(p.content) };
  }

  private stories = new Map<Story, StoryBox>();

  /** A header or footer's measured content (empty when there is none). */
  private story(story: Story | undefined, width: number): StoryBox {
    if (!story) return { height: 0, lines: [], markers: [] };
    let box = this.stories.get(story);
    if (!box) {
      const saved = this.rels;
      this.rels = story.preservedRels ?? [];
      box = this.measureCell(story.blocks, width);
      this.rels = saved;
      this.stories.set(story, box);
    }
    return box;
  }

  private paraHeight(b: ParaBox): number {
    return b.rp.spaceBefore + b.rp.borderTop + b.lines.reduce((s, l) => s + l.height, 0) + b.rp.spaceAfter + b.rp.borderBottom;
  }

  // --- tables --------------------------------------------------------------------------------

  private measureTable(t: Table, available: number, path: number[] = []): TableBox {
    const tableStyle = this.resolver.tableStyle(t.props.style);
    const grid = (from: number, n: number) => t.grid.slice(from, from + n).reduce((s, v) => s + v, 0);
    // Table position: centred/right tables use the slack; otherwise Word before 2013 outdents the table
    // so that cell text lines up with the margin.
    const total = grid(0, t.grid.length);
    const offset = t.props.align === "center" ? (available - total) / 2
      : t.props.align === "right" || t.props.align === "end" ? available - total
      : this.compat < 15 ? -CELL_MARGIN : 0;
    const borders = this.hasBorders(t, tableStyle);

    // Lay out every cell, remembering the grid column it starts in (to match vertical merges).
    const measured = t.rows.map((row) => {
      let col = 0;
      return row.cells.map((cell) => {
        const n = cell.props.span ?? 1;
        const start = col;
        col += n;
        const fill = cell.props.shading?.fill?.toUpperCase();
        const frame = {
          dx: offset + grid(0, start) + CELL_MARGIN, width: grid(start, n), borders, vMerge: cell.props.vMerge,
          ...(fill && fill !== "AUTO" ? { fill } : {}),
        };
        // w:hideMark: an empty cell's paragraph mark does not count towards the row's height.
        const hidden = cell.props.extra?.some((x) => x.name === "w:hideMark") &&
          cell.blocks.every((b) => b.type === "paragraph" && b.content.length === 0);
        const box: CellBox = cell.props.vMerge === "continue" || hidden
          ? { ...frame, height: 0, lines: [], markers: [] }
          : { ...frame, ...this.measureCell(cell.blocks, grid(start, n) - 2 * CELL_MARGIN, tableStyle) };
        return { start, vMerge: cell.props.vMerge, box };
      });
    });

    // Row heights come from ordinary cells; a vertically merged cell then stretches the last row it spans.
    const heights = t.rows.map((row, r) => {
      let h = Math.max(0, ...measured[r].filter((c) => !c.vMerge).map((c) => c.box.height));
      const rule = row.props.height;
      if (rule?.val) {
        const v = Number(rule.val) / 20;
        h = rule.hRule === "exact" ? v : Math.max(h, v);
      }
      return h + ROW_BORDER;
    });
    measured.forEach((cells, r) => cells.forEach((c) => {
      if (c.vMerge !== "restart") return;
      let last = r;
      while (measured[last + 1]?.some((x) => x.start === c.start && x.vMerge === "continue")) last++;
      const have = heights.slice(r, last + 1).reduce((s, v) => s + v, 0) - ROW_BORDER;
      if (c.box.height > have) heights[last] += c.box.height - have;
    }));

    const rows = t.rows.map((row, r): RowBox => ({
      index: r,
      height: heights[r],
      header: !!row.props.header,
      cantSplit: !!row.props.cantSplit,
      cells: measured[r].map((c) => c.box),
    }));
    return { kind: "table", path, rows };
  }

  /** True when the table (or its style) defines visible borders. */
  private hasBorders(t: Table, tableStyle?: string): boolean {
    const xml = [
      ...(t.props.extra ?? []),
      ...this.doc.styles.styles.filter((s) => s.id === tableStyle).flatMap((s) => s.table?.extra ?? []),
    ].find((x) => x.name === "w:tblBorders")?.xml;
    return !!xml && /w:val="(?!none|nil)[^"]+"/.test(xml);
  }

  /** Stacks a cell's blocks; adjacent paragraph spacing collapses as on the page. */
  private measureCell(blocks: Block[], width: number, tableStyle?: string): StoryBox {
    const lines: CellBox["lines"] = [];
    const markers: CellBox["markers"] = [];
    const floats: NonNullable<CellBox["floats"]> = [];
    let y = 0;
    let prevAfter = 0;
    for (const box of this.measureBlocks(blocks, width, tableStyle)) {
      if (box.kind === "table") {
        y += box.rows.reduce((s, r) => s + r.height, 0);
        prevAfter = 0;
        continue;
      }
      y += Math.max(box.rp.spaceBefore - prevAfter, 0) + box.rp.borderTop;
      prevAfter = box.rp.spaceAfter;
      markers.push([box.anchor, y]);
      box.floats.forEach((float, index) => floats.push({ float, dy: y, path: box.path, index }));
      for (const line of box.lines) {
        lines.push({ line, dy: y });
        y += line.height;
      }
      y += box.rp.spaceAfter + box.rp.borderBottom;
    }
    return { height: y, lines, markers, ...(floats.length ? { floats } : {}) };
  }

  /**
   * Splits a row at line boundaries so its first part fits in `limit` (Word lets rows break across
   * pages unless "cantSplit" is set). Returns undefined when not even one line fits.
   */
  private splitRow(row: RowBox, limit: number): [RowBox, RowBox] | undefined {
    const first: CellBox[] = [];
    const rest: CellBox[] = [];
    let firstHeight = 0;
    let restHeight = 0;
    for (const c of row.cells) {
      const keep = c.lines.filter((l) => l.dy + l.line.height <= limit);
      const move = c.lines.filter((l) => l.dy + l.line.height > limit);
      const cut = move[0]?.dy ?? Infinity;
      const shift = move.length ? cut : c.height;
      const bottom = keep.length ? keep[keep.length - 1].dy + keep[keep.length - 1].line.height : 0;
      first.push({ ...c, height: bottom, lines: keep, markers: c.markers.filter(([, dy]) => dy < cut) });
      rest.push({
        ...c,
        height: Math.max(c.height - shift, 0),
        lines: move.map((l) => ({ ...l, dy: l.dy - shift })),
        markers: c.markers.filter(([, dy]) => dy >= cut).map(([m, dy]) => [m, dy - shift] as [string, number]),
      });
      firstHeight = Math.max(firstHeight, bottom);
      restHeight = Math.max(restHeight, c.height - shift);
    }
    if (firstHeight === 0) return undefined;
    return [{ ...row, height: firstHeight, cells: first }, { ...row, height: restHeight + ROW_BORDER, cells: rest }];
  }

  private measureBlocks(blocks: Block[], width: number, tableStyle?: string, base: number[] = [], first = 0): Box[] {
    const out: Box[] = [];
    blocks.forEach((b, i) => {
      const path = [...base, first + i];
      if (b.type === "paragraph") out.push(this.measureParagraph(b, width, tableStyle, path));
      else if (b.type === "contentControl") out.push(...this.measureBlocks(b.blocks, width, tableStyle, path));
      else if (b.type === "table") out.push(this.measureTable(b, width, path));
    });
    return out;
  }

  // --- pagination ----------------------------------------------------------------------------

  layout(): LayoutResult {
    const pages: Page[] = [];
    const markers: LayoutResult["markers"] = {};
    const paragraphs: LayoutResult["paragraphs"] = [];
    const pageStarts: PageStart[] = [];
    const wraps: WrapPoint[] = [];
    /** Room already reported for wrapping on this page, per side, and how far down it reaches. */
    let reported: { left?: { width: number; until: number }; right?: { width: number; until: number } } = {};
    const floats: PlacedFloat[] = [];
    // A new page is "open" until its first body content is placed (recorded as the page start).
    let pageOpen = false;
    let page!: Page;
    let y = 0;
    let top = 0;
    let bottom = 0;
    let left = 0;
    let prev: ParaBox | undefined;
    // At the top of a page after a page break Word drops space-before; at the start of a new section
    // it keeps it, collapsed against the previous paragraph's space-after.
    let sectionTop = false;
    /** Areas of the current page that text wraps around (floating objects), in page coordinates. */
    type Exclusion = { x1: number; x2: number; y1: number; y2: number };
    let exclusions: Exclusion[] = [];

    // Headers and footers are inherited from the previous section unless redefined.
    let headers: Section["headers"] = {};
    let footers: Section["footers"] = {};
    let sectionPages = 0;

    const newPage = (section: Section, index: number) => {
      const w = section.props.pageWidth ?? 612;
      const h = section.props.pageHeight ?? 792;
      const mg = section.props.margins ?? {};
      const width = w - (mg.left ?? 72) - (mg.right ?? 72);
      const kind = sectionPages === 0 && section.props.titlePage ? "first" : "default";
      sectionPages++;
      const header = this.story(headers[kind], width);
      const footer = this.story(footers[kind], width);
      // A header or footer taller than its margin pushes the body text away from it.
      top = Math.max(mg.top ?? 72, (mg.header ?? 36) + header.height);
      bottom = Math.min(h - (mg.bottom ?? 72), h - (mg.footer ?? 36) - footer.height);
      left = mg.left ?? 72;
      page = { number: pages.length + 1, section: index, width: w, height: h, lines: [], rects: [], bodyTop: top, bodyBottom: bottom, bodyLeft: left, bodyRight: w - (mg.right ?? 72) };
      pages.push(page);
      pageOpen = pages.length > 1;
      for (const l of header.lines) put(l.line, left, (mg.header ?? 36) + l.dy, { story: "header" });
      for (const l of footer.lines) put(l.line, left, h - (mg.footer ?? 36) - footer.height + l.dy, { story: "footer" });
      y = top;
      prev = undefined;
      exclusions = [];
      reported = {};
    };
    const room = () => bottom - y;
    const atTop = () => y === top;
    const mark = (anchor: string, at: number) => {
      const pos = { page: page.number, y: Math.round(at * 10) / 10 };
      paragraphs.push({ text: anchor.slice(0, 80), ...pos });
      for (const n of anchor.match(MARKER) ?? []) markers[n] ??= pos;
    };
    const put = (line: MLine, x: number, at: number, extra: { story?: PositionedLine["story"]; path?: number[] } = {}) => {
      if (extra.path && pageOpen) {
        pageStarts.push({ page: page.number, path: extra.path, offset: line.startOffset });
        pageOpen = false;
      }
      if (extra.path && line.wrap) {
        // Report room on a side only where earlier reported room does not already cover it (the
        // editor's boxes stay in effect below their line, like the objects they stand for).
        const top = at + line.wrap.skip;
        const side = (s: "left" | "right") => {
          const w = line.wrap![s];
          const r = reported[s];
          if (!w || (r && Math.abs(r.width - w.width) < 0.5 && r.until >= top + w.height - 0.5)) return undefined;
          reported[s] = { width: w.width, until: top + w.height };
          return w;
        };
        const left = side("left");
        const right = side("right");
        if (line.wrap.skip > 0 || left || right) {
          wraps.push({ page: page.number, path: extra.path, offset: line.startOffset, skip: line.wrap.skip,
            ...(left ? { left } : {}), ...(right ? { right } : {}) });
        }
      }
      const baseline = at + line.ascent;
      page.lines.push({
        ...(extra.story ? { story: extra.story } : {}),
        x, y: at, height: line.height, baseline,
        segments: line.segments.map((s) => ({ ...s, x: x + s.x })),
        // Inline pictures stand on the baseline.
        images: line.images.map((i) => ({ ...i, x: x + i.x, y: baseline - i.height })),
        rules: line.rules.map((r) => ({ ...r, x1: x + r.x1, x2: x + r.x2, y1: baseline + r.y1, y2: baseline + r.y2 })),
      });
    };

    /**
     * Places a floating object anchored in a paragraph whose first line starts at `paraTop`: draws it
     * (as "float" lines) and returns the area text must wrap around, if any.
     */
    const placeFloat = (f: FloatInfo, paraTop: number, path: number[], index: number,
      column: [number, number] = [page.bodyLeft, page.bodyRight]): Exclusion | undefined => {
      const h = f.x.from;
      // In a table cell (Word's "layout in cell"), the column is the cell.
      const hArea = h === "page" ? [0, page.width]
        : h === "leftMargin" || h === "insideMargin" ? [0, page.bodyLeft]
        : h === "rightMargin" || h === "outsideMargin" ? [page.bodyRight, page.width]
        : column; // margin, column, character (approximated by the column)
      const v = f.y.from;
      const vArea = v === "page" ? [0, page.height] : v === "margin" ? [page.bodyTop, page.bodyBottom]
        : v === "topMargin" ? [0, page.bodyTop] : v === "bottomMargin" ? [page.bodyBottom, page.height]
        : [paraTop, paraTop]; // paragraph, line
      const insets = f.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
      const content = f.kind === "textbox" && f.blocks ? this.measureCell(f.blocks, f.width - insets.left - insets.right) : undefined;
      const height = content && (f.autoHeight || f.height < 1) ? content.height + insets.top + insets.bottom : f.height;
      const at = (pos: FloatInfo["x"], area: number[], size: number) =>
        pos.offset !== undefined ? area[0] + pos.offset
          : pos.align === "center" ? (area[0] + area[1] - size) / 2
          : pos.align === "right" || pos.align === "bottom" || pos.align === "outside" ? area[1] - size : area[0];
      const x = at(f.x, hArea, f.width);
      const fy = at(f.y, vArea, height);
      const from = page.lines.length;
      if (f.kind === "picture" && f.imagePart) {
        page.lines.push({ story: "float", x, y: fy, height, baseline: fy + height, segments: [], rules: [],
          images: [{ part: f.imagePart, x, y: fy, width: f.width, height }] });
      }
      if (content) for (const l of content.lines) put(l.line, x + insets.left, fy + insets.top + l.dy, { story: "float" });
      if (f.kind === "ink" && f.ink) {
        page.lines.push({ story: "float", x, y: fy, height, baseline: fy + height, segments: [], rules: [], images: [],
          strokes: f.ink.strokes.map((s) => ({ points: s.map((v, i) => (i % 2 ? fy + v * height : x + v * f.width)), width: f.ink!.width, color: f.ink!.color })) });
      }
      // The same, relative to the anchor paragraph and the object's box, for the editor.
      const shift = (l: PositionedLine): PositionedLine => ({
        ...l, x: l.x - x, y: l.y - fy, baseline: l.baseline - fy,
        segments: l.segments.map((g) => ({ ...g, x: g.x - x })),
        images: l.images.map((i) => ({ ...i, x: i.x - x, y: i.y - fy })),
        rules: l.rules.map((r) => ({ ...r, x1: r.x1 - x, x2: r.x2 - x, y1: r.y1 - fy, y2: r.y2 - fy })),
      });
      floats.push({
        page: page.number, path, index, x: x - column[0], y: fy - paraTop, width: f.width, height,
        ...(f.kind === "picture" && f.imagePart ? { image: { part: f.imagePart } } : {}),
        ...(f.kind === "ink" && f.ink ? { ink: f.ink } : {}),
        lines: f.kind === "textbox" ? page.lines.slice(from).map(shift) : [],
      });
      if (f.wrap === "none") return undefined;
      const span = f.wrap === "topAndBottom" ? { x1: -Infinity, x2: Infinity } : { x1: x - f.dist.left, x2: x + f.width + f.dist.right };
      return { ...span, y1: fy - f.dist.top, y2: fy + height + f.dist.bottom };
    };

    let topIndex = 0;
    this.doc.sections.forEach((section, si) => {
      const carried = prev;
      headers = { ...headers, ...section.headers };
      footers = { ...footers, ...section.footers };
      sectionPages = 0;
      if (si === 0 || section.props.breakType !== "continuous") newPage(section, si);
      prev = carried;
      // The first page of the document behaves like a section start: space-before is kept.
      sectionTop = true;
      const mg = section.props.margins ?? {};
      const width = (section.props.pageWidth ?? 612) - (mg.left ?? 72) - (mg.right ?? 72);
      const boxes = this.measureBlocks(section.blocks, width, undefined, [], topIndex);
      topIndex += section.blocks.length + 1; // the section's blocks, then its section break

      // Height a box needs at minimum on its first page (for keep-with-next decisions).
      const minimum = (b: Box | undefined): number => {
        if (!b) return 0;
        if (b.kind === "table") return (b.rows.find((r) => r.header)?.height ?? 0) + (b.rows.find((r) => !r.header)?.height ?? 0);
        const n = b.rp.widowControl ? Math.min(2, b.lines.length) : 1;
        return b.rp.spaceBefore + b.lines.slice(0, n).reduce((s, l) => s + l.height, 0);
      };

      boxes.forEach((box, bi) => {
        if (box.kind === "table") {
          const headers = box.rows.filter((r) => r.header);
          box.rows.forEach((row) => {
            let current = row;
            while (current.height > room() + 0.01 && !atTop()) {
              const parts = current.cantSplit || current.header ? undefined : this.splitRow(current, room());
              if (parts) {
                placeRow(parts[0], box.path);
                current = parts[1];
              }
              newPage(section, si);
              if (!current.header) for (const h of headers) placeRow(h, box.path, true);
            }
            placeRow(current, box.path);
          });
          prev = undefined;
          return;
        }
        const rp = box.rp;
        if (rp.pageBreakBefore && !atTop()) newPage(section, si);
        // Adjacent spacing collapses to the larger of space-after and space-before (Word behaviour,
        // confirmed against Word's positions in Experiment 2).
        let before = atTop() && !sectionTop ? 0 : Math.max(rp.spaceBefore - (prev?.rp.spaceAfter ?? 0), 0);
        sectionTop = false;
        if (prev && rp.contextualSpacing && prev.rp.contextualSpacing && prev.rp.styleId === rp.styleId) {
          y -= prev.rp.spaceAfter;
          before = 0;
        }
        const full = before + this.paraHeight(box) - rp.spaceBefore;
        if (!atTop()) {
          // Keep with next: this paragraph (and any keep-with-next chain) plus the start of what follows.
          let need = full;
          if (rp.keepNext) {
            let j = bi + 1;
            while (boxes[j]?.kind === "para" && (boxes[j] as ParaBox).rp.keepNext) need += this.paraHeight(boxes[j++] as ParaBox);
            need += minimum(boxes[j]);
          }
          const pageHeight = bottom - top;
          if ((rp.keepNext && need > room() && need <= pageHeight) || (rp.keepLines && full > room() && full <= pageHeight)) {
            newPage(section, si);
            before = 0;
          }
        }
        y += before;
        // Paragraph borders: the top border above the first line; side borders drawn per page.
        const bd = rp.borders;
        const contentLeft = left + rp.left;
        const contentRight = left + width - rp.right;
        const stroke = (b: { width: number; color?: string }) => ({ width: b.width, color: b.color });
        let sideTop = y;
        const drawSides = (to: number) => {
          if (bd.left) page.rects.push({ x: contentLeft - bd.left.space - bd.left.width / 2, y: sideTop, width: 0, height: to - sideTop, sides: { top: false, bottom: false, left: true, right: false }, stroke: stroke(bd.left) });
          if (bd.right) page.rects.push({ x: contentRight + bd.right.space + bd.right.width / 2, y: sideTop, width: 0, height: to - sideTop, sides: { top: false, bottom: false, left: true, right: false }, stroke: stroke(bd.right) });
        };
        const borderSpan = () => ({
          x: contentLeft - (bd.left ? bd.left.space + bd.left.width / 2 : 0),
          width: contentRight - contentLeft + (bd.left ? bd.left.space + bd.left.width / 2 : 0) + (bd.right ? bd.right.space + bd.right.width / 2 : 0),
        });
        if (bd.top) {
          page.rects.push({ ...borderSpan(), y: y + bd.top.width / 2, height: 0, sides: { top: true, bottom: false, left: false, right: false }, stroke: stroke(bd.top) });
          y += rp.borderTop;
        }
        let remaining = box.lines;
        let first = true;
        let endedOnBreak = false;
        // Floating objects anchored here, as placed on a page (moved if the paragraph moves on).
        let anchored: { page: Page; from: number; to: number; exclusions: Exclusion[]; floatsFrom: number } | undefined;
        while (remaining.length) {
          if (first && box.floats.length) {
            if (anchored) {
              anchored.page.lines.splice(anchored.from, anchored.to - anchored.from);
              exclusions = exclusions.filter((e) => !anchored!.exclusions.includes(e));
              floats.splice(anchored.floatsFrom);
            }
            const from = page.lines.length;
            const floatsFrom = floats.length;
            const placed = box.floats.map((f, i) => placeFloat(f, y, box.path, i)).filter((e): e is Exclusion => !!e);
            exclusions.push(...placed);
            anchored = { page, from, to: page.lines.length, exclusions: placed, floatsFrom };
          }
          if (first) {
            // Wrap around floating objects on this page: lines beside one get the free part of the
            // column; where too little is free, they move below it.
            const paraTop = y;
            const slot: Slot = (dy, height) => {
              const at = paraTop + dy;
              const over = exclusions.filter((e) => e.y1 < at + height && e.y2 > at);
              if (!over.length) return { left: 0, right: width, leftUntil: 0, rightUntil: 0 };
              let free = [[0, width]];
              for (const e of over) {
                const [a, b] = [e.x1 - left, e.x2 - left];
                free = free.flatMap(([l, r]) => (b <= l || a >= r ? [[l, r]] : [[l, Math.max(l, a)], [Math.min(r, b), r]]).filter(([p, q]) => q - p > 0.01));
              }
              const best = free.sort((p, q) => (q[1] - q[0]) - (p[1] - p[0]))[0];
              if (best && best[1] - best[0] >= MIN_WRAP) {
                // How far down the objects on each side of the free span reach (relative to the paragraph).
                const until = (side: Exclusion[]) => (side.length ? Math.max(...side.map((e) => e.y2)) - paraTop : 0);
                return {
                  left: best[0], right: best[1],
                  leftUntil: until(over.filter((e) => e.x2 - left <= best[0] + 0.01)),
                  rightUntil: until(over.filter((e) => e.x1 - left >= best[1] - 0.01)),
                };
              }
              // No room beside: the line moves below all the objects beside it (as Word does: with a
              // picture on each side, below the lower one). A paragraph pushed down keeps its space before.
              return { skipTo: Math.max(...over.map((e) => e.y2)) - paraTop + (dy === 0 ? rp.spaceBefore : 0) };
            };
            remaining = exclusions.some((e) => e.y2 > y) ? this.breakLines(box.para, rp, width, slot) : box.lines;
          }
          let fit = 0;
          let h = 0;
          while (fit < remaining.length && h + remaining[fit].fitHeight <= room() + 0.01) h += remaining[fit++].height;
          // Widow/orphan control works within the lines between manual page breaks, not across them.
          const breakAt = remaining.findIndex((l) => l.pageBreakAfter);
          const segment = breakAt < 0 ? remaining.length : breakAt + 1;
          if (fit < segment && rp.widowControl && segment >= 2) {
            if (fit === 1 && first) fit = 0;                                  // orphan: lone first line
            else if (segment - fit === 1 && fit >= 2) fit -= 1;               // widow: lone last line
          }
          if (fit === 0 && atTop()) fit = 1;                                  // taller than a page: place anyway
          if (fit === 0) { drawSides(y); newPage(section, si); sideTop = y; continue; }
          for (let k = 0; k < fit; k++) {
            const line = remaining[k];
            if (first && k === 0) mark(box.anchor, y);
            put(line, left, y, { path: box.path });
            y += line.height;
            if (line.pageBreakAfter) {
              remaining = remaining.slice(k + 1);
              first = false;
              endedOnBreak = remaining.length === 0;
              drawSides(y);
              newPage(section, si);
              sideTop = y;
              fit = -1;
              break;
            }
          }
          if (fit === -1) continue;
          remaining = remaining.slice(fit);
          first = false;
          if (remaining.length) { drawSides(y); newPage(section, si); sideTop = y; }
        }
        if (!endedOnBreak) drawSides(y + (bd.bottom ? bd.bottom.space + bd.bottom.width : 0));
        if (rp.borderLine && !endedOnBreak) {
          const b = rp.borderLine;
          page.rects.push({
            ...(bd.left || bd.right ? borderSpan() : { x: contentLeft, width: contentRight - contentLeft }),
            y: y + b.space + b.width / 2, height: 0,
            sides: { top: true, bottom: false, left: false, right: false }, stroke: { color: b.color, width: b.width },
          });
        }
        // A paragraph ending in a page break leaves its spacing behind on the previous page.
        if (!endedOnBreak) y += rp.spaceAfter + rp.borderBottom;
        prev = box;
      });

      function placeRow(row: RowBox, tablePath: number[], repeated = false) {
        if (!repeated && pageOpen) {
          pageStarts.push({ page: page.number, path: tablePath, row: row.index });
          pageOpen = false;
        }
        for (const [ci, c] of row.cells.entries()) {
          if (c.fill || c.borders) {
            page.rects.push({
              x: left + c.dx - CELL_MARGIN, y, width: c.width, height: row.height,
              ...(c.fill ? { fill: c.fill } : {}),
              ...(c.borders ? { sides: { top: c.vMerge !== "continue", bottom: c.vMerge !== "restart", left: true, right: true } } : {}),
            });
          }
          for (const [anchor, dy] of c.markers) mark(anchor, y + dy);
          for (const l of c.lines) put(l.line, left + c.dx, y + l.dy);
          // Floating objects in the cell (signatures, logos): placed in the cell; text does not wrap.
          for (const cf of c.floats ?? []) {
            placeFloat(cf.float, y + cf.dy, [...tablePath, row.index, ci, ...cf.path], cf.index, [left + c.dx, left + c.dx + c.width - 2 * CELL_MARGIN]);
          }
        }
        y += row.height;
      }
    });
    return { pages, markers, paragraphs, pageStarts, wraps, floats };
  }
}

export function layoutDocument(doc: PrimeDocument, measurer: Measurer): LayoutResult {
  return new LayoutEngine(doc, measurer).layout();
}
