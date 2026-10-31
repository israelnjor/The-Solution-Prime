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
  type Block, type Inline, type Paragraph, type PrimeDocument, type Relationship, type Section, type Story, type Table,
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
  /** Set for header and footer lines. */
  story?: "header" | "footer";
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
}

interface ParaBox {
  kind: "para";
  /** Index path of the block: top-level index (sections and their breaks counted), then children. */
  path: number[];
  rp: ResolvedParagraph;
  lines: MLine[];
  /** Plain text of the paragraph, used to identify it in results. */
  anchor: string;
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
type StoryBox = Pick<CellBox, "height" | "lines" | "markers">;

type Atom = { offset: number } & (
  | { kind: "word"; text: string; run: ResolvedRun; field?: Segment["field"] }
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
}

const key = (r: ResolvedRun): FontKey => ({ family: r.font, bold: r.bold, italic: r.italic });

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
          const text = run.caps ? piece.toUpperCase() : piece;
          const field = fieldKind();
          out.push(piece.startsWith(" ") ? { kind: "space", text, run, offset: at }
            : { kind: "word", text, run, offset: at, ...(field ? { field } : {}) });
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
    return fonts.map((font) => ({ font, size: run.size }));
  }

  private lineMetrics(faces: Face[]): { height: number; ascent: number } {
    let height = 0;
    let ascent = 0;
    for (const f of faces) {
      const fm = this.m.metrics(f.font);
      height = Math.max(height, fm.lineHeight * f.size);
      ascent = Math.max(ascent, fm.ascent * f.size);
    }
    return { height, ascent };
  }

  private applyLineRule(natural: number, rp: ResolvedParagraph): number {
    if (rp.lineRule === "exact") return rp.line / 20;
    if (rp.lineRule === "atLeast") return Math.max(natural, rp.line / 20);
    return natural * (rp.line / 240);
  }

  /** Greedy line breaking at spaces, using real glyph advances. */
  breakLines(p: Paragraph, rp: ResolvedParagraph, width: number): MLine[] {
    const atoms = this.atoms(p, rp);
    const lines: MLine[] = [];
    let label = rp.listLabel;

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
        const lr = rp.labelRun ?? rp.mark;
        const lw = this.textWidth(label, lr);
        segs.push({ x: startX, text: label, font: key(lr), size: lr.size, ...look(lr) });
        runs.push(...this.faces(lr, label));
        // The label sits in the hanging indent; text starts at the next tab stop (usually the left indent).
        const after = startX + lw;
        startX = rp.firstLine < 0 && after <= rp.left ? rp.left : (Math.floor(after / DEFAULT_TAB) + 1) * DEFAULT_TAB;
        label = undefined;
      }
      avail = width - rp.right - startX;
      x = 0;
    };
    const finish = (pageBreakAfter = false, endsWithBreak = pageBreakAfter, nextStart = lineStart) => {
      const textMetrics = this.lineMetrics(runs.length ? runs : this.faces(rp.mark));
      // Equations extend the line only where they reach beyond the text's own ascent and descent.
      const ascentAll = Math.max(textMetrics.ascent, mathAscent);
      const descentAll = Math.max(textMetrics.height - textMetrics.ascent, mathDescent);
      const metrics = { ascent: ascentAll, height: ascentAll + descentAll };
      // Line spacing does not scale pictures: Word sizes a picture line to the picture, plus the text
      // descent, plus the extra space the line rule adds to text (an "exact" rule still clips it).
      const text = this.applyLineRule(metrics.height, rp);
      const picture = imageHeight && rp.lineRule !== "exact"
        ? imageHeight + (metrics.height - metrics.ascent) + Math.max(text - metrics.height, 0) : 0;
      lines.push({
        width: x, height: Math.max(text, picture), ascent: Math.max(metrics.ascent, imageHeight),
        fitHeight: picture > text ? Math.max(imageHeight, text) : Math.max(text, picture),
        segments: segs, startX, pageBreakAfter, avail, endsWithBreak, images, rules, startOffset: lineStart,
      });
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
      while (["word", "image", "math"].includes(atoms[i + 1]?.kind)) word.push(atoms[++i]);
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
    return { kind: "para", path, rp, lines: this.breakLines(p, rp, width), anchor: paragraphText(p) };
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
        const box: CellBox = cell.props.vMerge === "continue"
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
      for (const line of box.lines) {
        lines.push({ line, dy: y });
        y += line.height;
      }
      y += box.rp.spaceAfter + box.rp.borderBottom;
    }
    return { height: y, lines, markers };
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
    };
    const room = () => bottom - y;
    const atTop = () => y === top;
    const mark = (anchor: string, at: number) => {
      const pos = { page: page.number, y: Math.round(at * 10) / 10 };
      paragraphs.push({ text: anchor.slice(0, 80), ...pos });
      for (const n of anchor.match(MARKER) ?? []) markers[n] ??= pos;
    };
    const put = (line: MLine, x: number, at: number, extra: { story?: "header" | "footer"; path?: number[] } = {}) => {
      if (extra.path && pageOpen) {
        pageStarts.push({ page: page.number, path: extra.path, offset: line.startOffset });
        pageOpen = false;
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
        while (remaining.length) {
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
        for (const c of row.cells) {
          if (c.fill || c.borders) {
            page.rects.push({
              x: left + c.dx - CELL_MARGIN, y, width: c.width, height: row.height,
              ...(c.fill ? { fill: c.fill } : {}),
              ...(c.borders ? { sides: { top: c.vMerge !== "continue", bottom: c.vMerge !== "restart", left: true, right: true } } : {}),
            });
          }
          for (const [anchor, dy] of c.markers) mark(anchor, y + dy);
          for (const l of c.lines) put(l.line, left + c.dx, y + l.dy);
        }
        y += row.height;
      }
    });
    return { pages, markers, paragraphs, pageStarts };
  }
}

export function layoutDocument(doc: PrimeDocument, measurer: Measurer): LayoutResult {
  return new LayoutEngine(doc, measurer).layout();
}
