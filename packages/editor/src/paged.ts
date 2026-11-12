/**
 * Paginated editing (Phase 4, ADR-007). The layout engine decides where pages start; the editor keeps
 * one flowing document and receives page-gap spacers at those points. After each redraw the spacers
 * are sized from the real screen positions, so every page's content starts exactly at the top of
 * its sheet's body area — any small difference between engine and browser corrects itself.
 */
import { isLineShape, type LayoutResult, layoutDocument, type Measurer, type PlacedFloat, Resolver, sectionColumns, shapePath } from "@prime/layout";
import type { PrimeDocument, SectionProps, ShapeStyle } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { fromProseMirror } from "./convert.js";
import { inkSvg } from "./ink.js";
import { type ResolvedFloat, type ResolvedStart, resolveFloats, resolveHyphens, resolvePageStarts, resolveWraps } from "./pages.js";

/** Whether `timed` records (self-tests turn it on; entries would otherwise pile up over a session). */
const timing = () => (globalThis as { primeTiming?: boolean }).primeTiming === true;

/** Records an event on the browser's performance timeline ("pw:" entries; see the perf self-test). */
export function mark(name: string) {
  if (timing()) performance.measure(`pw:${name}`, { start: performance.now(), end: performance.now() });
}

/** Times a step for the browser's performance timeline (when timing is on). */
export function timed<T>(name: string, fn: () => T): T {
  if (!timing()) return fn();
  const t0 = performance.now();
  try { return fn(); } finally { performance.measure(`pw:${name}`, { start: t0, end: performance.now() }); }
}

/** Gap between page sheets on screen, in pt. */
export const PAGE_GAP = 16;

/**
 * Where each page starts on screen (pt from the first page's top): pages are stacked at their own
 * heights with PAGE_GAP between (a section can have its own page size, e.g. landscape pages).
 */
export function pageTops(pages: readonly { height: number }[], gap = PAGE_GAP): number[] {
  const tops: number[] = [];
  let y = 0;
  for (const p of pages) { tops.push(y); y += p.height + gap; }
  return tops;
}

/**
 * How far the text column of each section is from the first section's on screen (pt): pages are
 * centred on the widest one, so a wider (landscape) section's text starts further left (`left`
 * negative) and ends further right (`right` negative). Indexed by section.
 */
export function sectionOffsets(sections: readonly SectionProps[]): { left: number; right: number }[] {
  const width = (s: SectionProps) => s.pageWidth ?? 612;
  const widest = Math.max(...sections.map(width), 0);
  const column = (s: SectionProps) => {
    const x = (widest - width(s)) / 2;
    return [x + (s.margins?.left ?? 72), x + width(s) - (s.margins?.right ?? 72)];
  };
  const [l0, r0] = sections.length ? column(sections[0]) : [0, 0];
  return sections.map((s) => { const [l, r] = column(s); return { left: l - l0, right: r0 - r }; });
}
const PX = 96 / 72;

export interface PagedOptions {
  /** The document the editor was opened from (styles, lists, package parts). */
  base: () => PrimeDocument;
  measurer: Measurer;
  /** Called with every new layout (to draw page sheets, headers and footers). */
  onLayout?: (layout: LayoutResult) => void;
  /**
   * Called after spacers are sized, with the pages whose preceding content overflowed into the
   * margin (the screen ran longer than the layout). Margin masks must not hide those.
   */
  onFit?: (overflow: Set<number>) => void;
  /** Quiet time after typing before re-paginating, in ms. */
  delay?: number;
}

interface PagedState {
  decorations: DecorationSet;
  /** The text columns the editor shows (the editor is then a multi-column container), if any. */
  columns?: ScreenColumns;
}

/** Equal text columns, as the browser draws them: how many, the space between (pt), a line between. */
interface ScreenColumns { count: number; gap: number; rule: boolean }

/**
 * The text columns to show on screen. The editor is one multi-column container: blocks of sections
 * with these columns flow through them (the browser splits paragraphs between columns, as Word
 * does); every other block spans the page (column-span: all), as do page gaps, which end a page's
 * set of columns. The browser's columns are equally wide and as wide as the first section's text: a
 * section with other columns shows one column on screen (its printed pages are right).
 */
function screenColumns(props: SectionProps[]): { columns?: ScreenColumns; inColumns: boolean[] } {
  const body = (q: SectionProps) => (q.pageWidth ?? 612) - (q.margins?.left ?? 72) - (q.margins?.right ?? 72);
  const offsets = sectionOffsets(props);
  const specs = props.map((q, i) => {
    const c = sectionColumns(q, body(q));
    const equal = c && new Set(c.widths.map((w) => Math.round(w * 10))).size === 1;
    return c && equal && !offsets[i].left && !offsets[i].right ? { count: c.widths.length, gap: c.gaps[0], rule: c.sep } : undefined;
  });
  const columns = specs.find(Boolean);
  const same = (c: ScreenColumns | undefined) => !!c && !!columns && c.count === columns.count && Math.abs(c.gap - columns.gap) < 0.1;
  return { ...(columns ? { columns } : {}), inColumns: specs.map(same) };
}

/** The text columns the editor shows for a document (see screenColumns). */
function columnsIn(doc: Node): ScreenColumns | undefined {
  const sections: SectionProps[] = [];
  doc.forEach((node) => { if (node.type.name === "section_break") sections.push(node.attrs.section?.props ?? {}); });
  sections.push(doc.attrs.section?.props ?? {});
  return screenColumns(sections).columns;
}

export const paginationKey = new PluginKey<PagedState>("pagination");

/** The page a position is on, as the pages are drawn (1 before the first page break). */
export function pageAt(state: EditorState, pos = state.selection.head): number {
  let page = 1;
  for (const d of paginationKey.getState(state)?.decorations.find(0, pos) ?? []) {
    const p = (d.spec as { page?: number }).page;
    if (p && p > page && d.from <= pos) page = p;
  }
  return page;
}

function gap(page: number, inline: boolean): () => HTMLElement {
  return () => {
    const el = document.createElement(inline ? "span" : "div");
    el.className = "pw-page-gap";
    el.dataset.page = String(page);
    el.contentEditable = "false";
    return el;
  };
}

/**
 * Room for floating objects where the layout wraps text around them, at the start of the affected
 * line: a block of the skip's height moves the line down (a float would not move an empty line), then
 * CSS floats of the engine's sizes narrow it on either side.
 */
function wrapSpace(w: ReturnType<typeof resolveWraps>[number]): () => HTMLElement {
  return () => {
    const el = document.createElement("span");
    el.className = "pw-wrap";
    el.contentEditable = "false";
    // Each box clears earlier ones on its side: CSS would otherwise set it beside them.
    const box = (side: "left" | "right", width: string, height: number) => {
      const b = document.createElement("span");
      b.style.cssText = `float:${side};clear:${side};width:${width};height:${Math.max(0, height)}pt`;
      el.appendChild(b);
    };
    if (w.skip > 0) {
      const b = document.createElement("span");
      b.style.cssText = `display:block;height:${w.skip}pt`;
      el.appendChild(b);
    }
    if (w.left) box("left", `${w.left.width}pt`, w.left.height);
    if (w.right) box("right", `${w.right.width}pt`, w.right.height);
    return el;
  };
}

/** Object URLs for image parts, made once per part (pictures are large; data URLs would be copied). */
const imageUrls = new WeakMap<PrimeDocument, Map<string, string>>();
function imageUrl(base: PrimeDocument, part: string): string | undefined {
  let urls = imageUrls.get(base);
  if (!urls) imageUrls.set(base, (urls = new Map()));
  let url = urls.get(part);
  if (url === undefined) {
    const p = base.package.parts.find((x) => x.name === part);
    const ext = part.split(".").pop()?.toLowerCase() ?? "";
    const type = ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" } as Record<string, string>)[ext];
    if (!p || !type) return undefined;
    url = URL.createObjectURL(new Blob([Uint8Array.from(atob(p.data), (c) => c.charCodeAt(0))], { type }));
    urls.set(part, url);
  }
  return url;
}

/** Object URLs for image resources (pictures inserted in Prime Word or read as ordinary pictures). */
function resourceUrl(base: PrimeDocument, id: string): string | undefined {
  let urls = imageUrls.get(base);
  if (!urls) imageUrls.set(base, (urls = new Map()));
  const key = `resource:${id}`;
  let url = urls.get(key);
  const res = base.images[id];
  if (url === undefined && res) {
    url = URL.createObjectURL(new Blob([Uint8Array.from(atob(res.data), (c) => c.charCodeAt(0))], { type: res.contentType }));
    urls.set(key, url);
  }
  return url;
}

/** What a float layer shows (redrawn when it changes: its place, size, picture, text, fill or outline). */
function floatKey(f: PlacedFloat, left: number): string {
  const text = f.lines.map((l) => l.segments.map((g) => `${g.text}${g.font.bold ? "b" : ""}${g.font.italic ? "i" : ""}${g.color ?? ""}`).join("")).join("|");
  const style = f.shape ? `${f.shape.geometry}:${f.shape.fill ?? ""}:${f.shape.line ? `${f.shape.line.color}@${f.shape.line.width}` : ""}` : "";
  return `float-${f.x}-${f.y}-${f.width}-${f.height}-${f.image?.part ?? f.image?.imageId ?? ""}-${left}-${style}-${text}`;
}

/**
 * A floating picture or text box, drawn inside its anchor paragraph where the layout put it, so it
 * moves with its text on screen. `left` is the paragraph's left indent (the layer's origin).
 */
function floatLayer(f: PlacedFloat, base: PrimeDocument, left: number): () => HTMLElement {
  return () => {
    const el = document.createElement("span");
    el.className = "pw-float";
    el.contentEditable = "false";
    el.style.cssText = `left:${f.x - left}pt;top:${f.y}pt;width:${f.width}pt;height:${f.height}pt`;
    // For moving and resizing (floats.ts): which float, and its box in layout terms.
    // Pictures and ink keep their proportions when resized; text boxes do not.
    Object.assign(el.dataset, { index: String(f.index), x: String(f.x), y: String(f.y), w: String(f.width), h: String(f.height), kind: f.image || f.ink ? "picture" : f.textbox ? "textbox" : "shape" });
    el.title = "Drag to move; drag a corner to resize; Delete to remove";
    for (const corner of ["nw", "ne", "sw", "se"]) {
      const h = document.createElement("span");
      h.className = `pw-float-handle pw-${corner}`;
      h.dataset.handle = corner;
      el.appendChild(h);
    }
    // A shape, or a text box's fill and border (beneath its text).
    if (f.shape) el.appendChild(shapeSvg(f.shape, f.width, f.height));
    const src = f.image?.part ? imageUrl(base, f.image.part) : f.image?.imageId ? resourceUrl(base, f.image.imageId) : undefined;
    if (src) {
      const img = document.createElement("img");
      img.src = src;
      img.alt = "";
      img.style.cssText = "width:100%;height:100%";
      el.appendChild(img);
    }
    if (f.ink) el.appendChild(inkSvg(f.ink, f.width, f.height));
    for (const line of f.lines) {
      for (const s of line.segments) {
        const t = document.createElement("span");
        t.className = "pw-float-text";
        t.textContent = s.text;
        t.style.cssText = `left:${s.x}pt;top:${line.y}pt;line-height:${line.height}pt;font-family:"${s.font.family}";font-size:${s.size}pt;` +
          `font-weight:${s.font.bold ? 700 : 400};font-style:${s.font.italic ? "italic" : "normal"}` +
          (s.color ? `;color:#${s.color}` : "") + (s.underline ? ";text-decoration:underline" : "");
        el.appendChild(t);
      }
    }
    return el;
  };
}

/** A shape drawn as SVG in its box (the same outline the PDF draws). */
function shapeSvg(style: ShapeStyle, w: number, h: number): SVGSVGElement {
  const NS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("class", "pw-shape");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.setAttribute("overflow", "visible");
  const path = document.createElementNS(NS, "path");
  path.setAttribute("d", shapePath(style.geometry, w, h));
  path.setAttribute("fill", style.fill && !isLineShape(style.geometry) ? `#${style.fill}` : "none");
  if (style.line) {
    path.setAttribute("stroke", `#${style.line.color}`);
    path.setAttribute("stroke-width", String(style.line.width));
  }
  svg.appendChild(path);
  return svg;
}

/** A soft hyphen: where the layout hyphenated a word, the screen may break it too (and nowhere else). */
function softHyphen(): HTMLElement {
  const el = document.createElement("span");
  el.className = "pw-shy";
  el.textContent = String.fromCharCode(0xad); // U+00AD SOFT HYPHEN
  return el;
}

/**
 * Classes blocks get from the page layout: outside text columns, a block spans the page (pw-span); an
 * empty paragraph that only holds a continuous section break takes no room (pw-break-holder, as in
 * the layout). For the top-level blocks between `from` and `to` (all by default).
 */
function layoutClasses(doc: Node, from = 0, to = doc.content.size): Decoration[] {
  const out: Decoration[] = [];
  const sections: SectionProps[] = [];
  doc.forEach((node) => { if (node.type.name === "section_break") sections.push(node.attrs.section?.props ?? {}); });
  sections.push(doc.attrs.section?.props ?? {});
  const { columns, inColumns } = screenColumns(sections);
  // Sections with their own page width or margins: their blocks' text column (see sectionOffsets;
  // paragraph and table margins add these variables).
  const offsets = sectionOffsets(sections);
  let si = 0;
  doc.forEach((node, pos, index) => {
    if (node.type.name === "section_break") { si++; return; }
    if (pos + node.nodeSize < from || pos > to) return;
    const o = offsets[si];
    if (o.left || o.right) out.push(Decoration.node(pos, pos + node.nodeSize, { style: `--pw-sx:${o.left}pt;--pw-sr:${o.right}pt` }, { layoutClass: true }));
    if (columns && !inColumns[si]) out.push(Decoration.node(pos, pos + node.nodeSize, { class: "pw-span" }, { layoutClass: true }));
    const next = doc.maybeChild(index + 1);
    const following = next?.type.name === "section_break" ? sections[si + 1] : undefined;
    if (node.type.name === "paragraph" && node.content.size === 0 && following?.breakType === "continuous" && index > 0
      // (a section on other paper starts a new page: its break holder is an ordinary empty paragraph)
      && (following.pageWidth ?? 612) === (sections[si].pageWidth ?? 612) && (following.pageHeight ?? 792) === (sections[si].pageHeight ?? 792)) {
      out.push(Decoration.node(pos, pos + node.nodeSize, { class: "pw-break-holder" }, { layoutClass: true }));
    }
  });
  return out;
}

function decorations(doc: Node, starts: ResolvedStart[], wraps: ReturnType<typeof resolveWraps>,
  floats: ResolvedFloat[] = [], base?: PrimeDocument, hyphens: number[] = []): DecorationSet {
  const out: Decoration[] = [];
  // Text columns: blocks outside them span the page. The layout's column starts are not forced on
  // screen: where the screen's text ran a little longer or shorter than the layout's, forced breaks gave
  // a page more columns than it has, off the side of the page. The browser balances each page's
  // columns (a page gap ends them); a column break typed in the text still breaks (styles.css).
  out.push(...layoutClasses(doc));
  for (const pos of hyphens) out.push(Decoration.widget(pos, softHyphen, { side: -1, key: "shy", marks: [] }));
  if (base) {
    const resolver = new Resolver(base);
    for (const f of floats) {
      const para = doc.nodeAt(f.pos);
      if (!para || para.type.name !== "paragraph") continue;
      const left = resolver.paragraph({ type: "paragraph", props: para.attrs.props, content: [] }).left;
      out.push(Decoration.node(f.pos, f.pos + para.nodeSize, { style: "position:relative" }));
      out.push(Decoration.widget(f.pos + 1, floatLayer(f, base, left), { side: -1, key: floatKey(f, left) }));
    }
  }
  for (const w of wraps) {
    out.push(Decoration.widget(w.pos, wrapSpace(w), { side: -1, key: `wrap-${w.page}-${w.skip}-${w.left?.width}-${w.left?.height}-${w.right?.width}-${w.right?.height}` }));
  }
  for (const s of starts) {
    if (s.kind === "row") {
      const row = doc.nodeAt(s.pos);
      if (row) out.push(Decoration.node(s.pos, s.pos + row.nodeSize, { class: "pw-page-row", "data-page": String(s.page) }, { page: s.page }));
    } else {
      out.push(Decoration.widget(s.pos, gap(s.page, s.kind === "inline"), { side: -1, key: `gap-${s.page}-${s.kind}`, page: s.page }));
    }
  }
  return DecorationSet.create(doc, out);
}

/**
 * How much space the paragraph at `pos` has before it (in CSS px) beyond the space after the block
 * above (the two collapse): what a page break before it takes away (`.pw-page-gap + p`).
 */
function spaceBefore(view: EditorView, pos: number): number {
  const el = view.nodeDOM(pos) as HTMLElement | null;
  if (!el || el.tagName !== "P") return 0;
  const css = (v: string) => (v.endsWith("pt") ? (parseFloat(v) * 96) / 72 : parseFloat(v) || 0);
  let prev = el.previousElementSibling;
  while (prev?.classList.contains("pw-page-gap")) prev = prev.previousElementSibling;
  const after = prev ? parseFloat(getComputedStyle(prev).marginBottom) || 0 : 0;
  return Math.max(0, css(el.style.marginTop) - after);
}

/**
 * The start of the first line of paragraph `el` that crosses the bottom of the body (`limit`), when
 * it is not the paragraph's first line and starts after the page's own start (`after`): a break there
 * moves only the lines that do not fit, as the layout splits paragraphs. Moving the whole paragraph
 * left much of the page empty and pushed that much more onto every page after it, so a long book
 * corrected nearly every page, round after round. `planned` turns a screen y into the planned
 * position fit works in.
 */
function lineBreak(view: EditorView, el: Element, limit: number, after: number, planned: (clientY: number) => number):
  { kind: "inline"; pos: number; top: number } | undefined {
  if (el.tagName !== "P") return undefined;
  // The paragraph's lines, from its text (spacers and pictures' boxes aside).
  const lines: { top: number; bottom: number; left: number; y: number }[] = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!n.textContent || n.parentElement?.closest(".pw-page-gap, .pw-wrap, .pw-float-anchor")) continue;
    range.selectNodeContents(n);
    for (const r of Array.from(range.getClientRects())) {
      if (!r.height) continue;
      const top = planned(r.top);
      const bottom = planned(r.bottom);
      const line = lines.find((l) => top < l.bottom - 1 && bottom > l.top + 1);
      if (line) {
        line.top = Math.min(line.top, top);
        line.bottom = Math.max(line.bottom, bottom);
        if (r.left < line.left) { line.left = r.left; line.y = (r.top + r.bottom) / 2; }
      } else lines.push({ top, bottom, left: r.left, y: (r.top + r.bottom) / 2 });
    }
  }
  lines.sort((a, b) => a.top - b.top);
  const i = lines.findIndex((l) => l.bottom > limit);
  if (i < 1 || lines[i].top <= after + 1) return undefined;
  // The line's first character: in the first text node with text on that line or later, found by
  // binary search over its characters (from the DOM: hit-testing works only on screen).
  const from = lines[i].top - 1;
  const charTop = (n: globalThis.Node, k: number) => { range.setStart(n, k); range.setEnd(n, k + 1); const r = range.getClientRects()[0]; return r ? planned(r.top + r.height / 2) : -Infinity; };
  walker.currentNode = el;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.textContent?.length ?? 0;
    if (!len || n.parentElement?.closest(".pw-page-gap, .pw-wrap, .pw-float-anchor")) continue;
    if (charTop(n, len - 1) < from) continue; // all of it is on earlier lines
    let lo = 0;
    let hi = len - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (charTop(n, mid) >= from) hi = mid;
      else lo = mid + 1;
    }
    let pos: number;
    try { pos = view.posAtDOM(n, lo); } catch { return undefined; }
    const $pos = view.state.doc.resolve(pos);
    // The break must fall inside this paragraph, after its start.
    if (!$pos.parent.isTextblock || $pos.parentOffset === 0) return undefined;
    // The line box starts halfway between the text of the line before and its own.
    return { kind: "inline", pos, top: (lines[i - 1].bottom + lines[i].top) / 2 };
  }
  return undefined;
}

/**
 * Where the break for a page should move when the screen's content overflows that page: before the
 * first row, top-level block or (in a paragraph) line, after the page's own start, that crosses the
 * bottom of the body.
 */
function earlierBreak(view: EditorView, blocks: Element[], limit: number, after: number, before: number,
  top: (el: Element) => number, bottom: (el: Element) => number, planned: (clientY: number) => number):
  { kind: "row" | "block" | "inline"; pos: number; top: number } | undefined {
  // Only the blocks on this page (found by binary search: blocks are in screen order), and the rows of
  // its tables — scanning the whole document for every page made long documents slow.
  let lo = 0;
  let hi = blocks.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (top(blocks[mid]) <= after + 1) lo = mid + 1;
    else hi = mid;
  }
  const near: Element[] = [];
  // From the block that starts before the page (a table continuing onto it) to the page's end.
  for (let i = Math.max(0, lo - 1); i < blocks.length && top(blocks[i]) < before; i++) near.push(blocks[i]);
  const fits = (el: Element) => top(el) > after + 1 && top(el) < before && bottom(el) > limit;
  const rows = near.flatMap((el) => Array.from(el.querySelectorAll("tr"))).filter(fits);
  // A paragraph that crosses the bottom, even one continued from the page before.
  const split = near.find((el) => el.tagName === "P" && top(el) < limit && bottom(el) > limit);
  // A table that crosses the bottom breaks before its crossing row, not before the whole table
  // (moving a whole table for a row that ran 2px over left most of a page empty, and every page
  // after it then overflowed in turn); only a table whose first row crosses moves as a whole.
  const candidates = [...rows, ...near.filter((el) => fits(el) && !rows.some((r) => el.contains(r)))]
    .sort((a, b) => top(a) - top(b));
  for (const el of candidates) {
    if (split && top(split) < top(el)) break; // the paragraph comes first
    const line = el === split ? lineBreak(view, el, limit, after, planned) : undefined;
    if (line) return line;
    const $pos = view.state.doc.resolve(view.posAtDOM(el, 0));
    for (let d = $pos.depth; d > 0; d--) {
      const name = $pos.node(d).type.name;
      if (el.tagName === "TR" && name === "table_row") return { kind: "row", pos: $pos.before(d), top: top(el) };
      if (el.tagName !== "TR" && d === 1) return { kind: "block", pos: $pos.before(1), top: top(el) };
    }
  }
  return split ? lineBreak(view, split, limit, after, planned) : undefined;
}

export function pagination(opts: PagedOptions): Plugin {
  let layout: LayoutResult | undefined;
  /** The page starts currently drawn (from the layout, possibly moved earlier by `fit`). */
  let starts: ResolvedStart[] = [];
  /** The page starts as the last layout gave them (before corrections), kept on their text like `starts`. */
  let raw: ResolvedStart[] = [];
  /** Where text wraps around floating objects, and the objects themselves (from the layout). */
  let wraps: ReturnType<typeof resolveWraps> = [];
  /** Whether the drawn decorations include layout classes (text columns, break holders): kept up after edits. */
  let layoutClassed = false;
  /** Where the layout hyphenated words (editor positions). */
  let hyphens: number[] = [];
  let floats: ResolvedFloat[] = [];
  /** Rounds of corrections (each moves every break that needs it) since the last layout: a loop guard. */
  let rounds = 0;
  /** Pages whose overflow has no earlier break (e.g. one block taller than the page), until the next layout. */
  let unfixable = new Set<number>();

  /** Where each page starts (pt from the first page's top), for the current layout. */
  let tops: number[] = [];
  /** Where page `page`'s body starts, in screen px from the editor's top (the first page's body top). */
  const bodyTopOf = (page: number) => {
    const pages = layout!.pages;
    return ((tops[page - 1] ?? 0) + (pages[page - 1]?.bodyTop ?? pages[0].bodyTop) - pages[0].bodyTop) * PX;
  };
  /** How far down content before page `page` may reach: the previous page's body bottom (+1 px). */
  const limitOf = (page: number) => {
    const pages = layout!.pages;
    const previous = pages[page - 2];
    return previous ? (tops[page - 2] + previous.bodyBottom - pages[0].bodyTop) * PX + 1 : Infinity;
  };

  /** Spacer sizes currently in the stylesheet, by spacer ("gap-3", "row-3"). */
  let applied = new Map<string, number>();

  /**
   * Sizes each spacer so the next content lands at the top of its page's body area. Sizes go into a
   * stylesheet outside the editor: changing the editor's own DOM would make ProseMirror redraw it.
   * If the screen's content overflows a page (it ran longer than the layout predicted), that page's
   * break moves earlier, so a page never overflows; the screen has the last word.
   *
   * Positions are read in one pass (each read after a style change would re-measure the whole
   * document — in a long document, that is what made typing slow): growing a spacer by d moves
   * everything after it down by d, so the later positions follow from the earlier sizes. One
   * check afterwards catches the rare case where that does not hold exactly (text reflowing
   * around an inline spacer), with a second pass. Breaks that must move earlier are found with the
   * same arithmetic, all in one pass and one redraw (one redraw per page made a long document's
   * pages jump around for seconds after each pause in typing).
   */
  const fit = (view: EditorView, sheet: HTMLStyleElement, pass = 0) => {
    if (!layout?.pages.length) return;
    const box = timed("fit.reflow", () => view.dom.getBoundingClientRect());
    // Screen measurements include any zoom; sizes we set are in unzoomed CSS pixels.
    // (Unrounded CSS width: offsetWidth is a whole number, and its rounding error grows down the document.)
    const cssWidth = parseFloat(getComputedStyle(view.dom).width);
    const zoom = cssWidth ? box.width / cssWidth : 1;
    // Nothing moves during this pass (sizes are written at its end): each element is measured once.
    // (Hundreds of corrections in a long book measured every block again for each one: seconds.)
    const rects = new Map<Element, DOMRect>();
    const rect = (el: Element) => { let r = rects.get(el); if (!r) rects.set(el, (r = el.getBoundingClientRect())); return r; };
    const top = (el: Element) => (rect(el).top - box.top) / zoom;
    const bottom = (el: Element) => (rect(el).bottom - box.top) / zoom;
    let blocks: Element[] | undefined;
    const topLevel = () => (blocks ??= Array.from(view.dom.children).filter((el) => !el.classList.contains("pw-page-gap")));
    const bodyTopAt = (page: number) => bodyTopOf(page);
    const overflow = new Set<number>();
    const items = Array.from(view.dom.querySelectorAll<HTMLElement>(".pw-page-gap, .pw-page-row"));
    const tops = items.map(top); // the one measuring pass
    const sizes = new Map<string, number>();
    let shift = 0; // how far the planned sizes move content, so far
    let off = false;
    let moves = 0;
    let previousEnd = 0; // where the previous page's content starts, after the planned sizes
    // When the previous page's break moved earlier, content between the new break and the old spacer
    // (measured above `at`) never had the old spacer's size, which `shift` has already taken away.
    let lastMove: { at: number; size: number } | undefined;
    const extra = (measured: number) => (lastMove && measured < lastMove.at ? lastMove.size : 0);
    // Positions as they will be once the planned sizes apply (valid up to the next spacer).
    const plannedTop = (el: Element) => { const t = top(el); return t + shift + extra(t); };
    const plannedBottom = (el: Element) => bottom(el) + shift + extra(top(el));
    const plannedY = (clientY: number) => { const t = (clientY - box.top) / zoom; return t + shift + extra(t); };
    for (let i = 0; i < items.length; i++) {
      const el = items[i];
      const page = Number(el.dataset.page);
      const key = `${el.classList.contains("pw-page-row") ? "row" : "gap"}-${page}`;
      const current = applied.get(key) ?? 0;
      // Where the spacer would sit with no size of its own, after the sizes planned before it.
      const natural = tops[i] + shift;
      const size = Math.max(0, bodyTopAt(page) - natural);
      // Content before this spacer should end inside the previous page's body area.
      const limit = limitOf(page);
      if (natural > limit) {
        if (rounds < 10 && !unfixable.has(page)) {
          // Only content after the previous page's start may move to this page.
          const found = timed("fit.earlierBreak", () => earlierBreak(view, topLevel(), limit, Math.max(previousEnd, bodyTopAt(page - 1)), natural, plannedTop, plannedBottom, plannedY));
          // A correction moves the break earlier: finding the break it already has (a line that ends a
          // fraction of a pixel low) changes nothing, and counted as a correction round after round.
          const now = starts.find((s) => s.page === page);
          const moved = found && (!now || found.pos < now.pos) ? found : undefined;
          if (!moved && natural - limit < 3) { sizes.set(key, size); shift += size - current; previousEnd = natural + size; lastMove = undefined; if (Math.abs(size - current) > 0.5) off = true; continue; }
          if (moved) {
            moves++;
            mark(`correction p${page} +${Math.round(natural - limit)}px`);
            starts = starts.map((s) => (s.page === page ? { page, kind: moved.kind, pos: moved.pos } : s));
            // The spacer moves up to the new break: content from there moves by its new size, and
            // content after the old break no longer has the old spacer's size.
            const movedSize = Math.max(0, bodyTopAt(page) - moved.top);
            // A paragraph that starts a page has no space before it (styles.css): the one that started
            // it gets its space back, pushing what follows down (a long book's headings made the next
            // pages overflow again, round after round), and a new one loses its space, so its spacer
            // grows by that much for the paragraph to land at the top.
            const regained = now?.kind === "block" ? spaceBefore(view, now.pos) : 0;
            const lost = moved.kind === "block" ? spaceBefore(view, moved.pos) : 0;
            sizes.delete(key);
            sizes.set(`${moved.kind === "row" ? "row" : "gap"}-${page}`, movedSize + lost);
            shift += movedSize - current + regained;
            previousEnd = moved.top + movedSize;
            lastMove = { at: tops[i], size: current - regained };
            continue;
          }
          unfixable.add(page);
        }
        overflow.add(page);
      }
      sizes.set(key, size);
      shift += size - current;
      previousEnd = natural + size;
      lastMove = undefined;
      if (Math.abs(size - current) > 0.5) off = true;
    }
    if (moves) {
      rounds++;
      mark(`round ${moves}`);
      // Sizes first, then the breaks: the redraw that follows refits (and checks) everything.
      write(sheet, sizes);
      view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false));
      return;
    }
    if (off || sizes.size !== applied.size) timed("fit.write", () => write(sheet, sizes));
    rects.clear(); // the page now has the new sizes: measure again from here
    // Verify: every spacer should now end at its page's body top.
    if (off && pass < 2) {
      const wrong = items.some((el) => Math.abs(top(el) + (sizes.get(`${el.classList.contains("pw-page-row") ? "row" : "gap"}-${el.dataset.page}`) ?? 0) - bodyTopAt(Number(el.dataset.page))) > 1);
      if (wrong) return fit(view, sheet, pass + 1);
    }
    opts.onFit?.(overflow);
  };

  /**
   * Fitting after an edit at `from`: pages before it are untouched, and everything after the first
   * spacer that follows it keeps its place once that one spacer takes up the difference, so only that
   * spacer is measured and resized (measuring all of them after every keystroke cost ~30 ms in a
   * 944-page book). Returns false when that is not enough — the edited page now overflows — and a
   * full fit must run.
   */
  const fastFit = (view: EditorView, sheet: HTMLStyleElement, from: number): boolean => {
    if (!layout?.pages.length) return true;
    const items = view.dom.querySelectorAll<HTMLElement>(".pw-page-gap, .pw-page-row");
    // The first spacer after the edit (spacers are in document order).
    let lo = 0;
    let hi = items.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      let pos: number;
      try { pos = view.posAtDOM(items[mid], 0); } catch { return false; }
      if (pos <= from) lo = mid + 1;
      else hi = mid;
    }
    if (lo >= items.length) return true; // the last page: nothing below it to place
    const el = items[lo];
    const page = Number(el.dataset.page);
    const key = `${el.classList.contains("pw-page-row") ? "row" : "gap"}-${page}`;
    const box = view.dom.getBoundingClientRect();
    const cssWidth = parseFloat(getComputedStyle(view.dom).width);
    const zoom = cssWidth ? box.width / cssWidth : 1;
    const natural = (el.getBoundingClientRect().top - box.top) / zoom;
    if (natural > limitOf(page)) return false;
    const size = Math.max(0, bodyTopOf(page) - natural);
    if (Math.abs(size - (applied.get(key) ?? 0)) > 0.5) write(sheet, new Map(applied).set(key, size));
    return true;
  };

  const write = (sheet: HTMLStyleElement, sizes: Map<string, number>) => {
    const rules: string[] = [];
    for (const [key, size] of sizes) {
      const [kind, page] = key.split("-");
      rules.push(kind === "row"
        ? `.pw-editor tr.pw-page-row[data-page="${page}"] > td { padding-top: ${size}px; }`
        : `.pw-editor .pw-page-gap[data-page="${page}"] { height: ${size}px; }`);
    }
    sheet.textContent = rules.join("\n");
    applied = new Map(sizes);
  };

  return new Plugin<PagedState>({
    key: paginationKey,
    state: {
      init: () => ({ decorations: DecorationSet.empty }),
      apply(tr, old) {
        const drawn = tr.getMeta(paginationKey) as ResolvedStart[] | undefined;
        if (drawn) {
          const columns = columnsIn(tr.doc);
          const set = decorations(tr.doc, drawn, wraps, floats, opts.base(), hyphens);
          layoutClassed = set.find(undefined, undefined, (spec) => spec.layoutClass).length > 0;
          return { decorations: set, ...(columns ? { columns } : {}) };
        }
        if (!tr.docChanged) return old;
        // Keep page starts and wrap points on their text until the next layout replaces them (a
        // correction made before then must not put a break where the text was before an edit).
        starts = starts.map((p) => ({ ...p, pos: tr.mapping.map(p.pos) }));
        raw = raw.map((p) => ({ ...p, pos: tr.mapping.map(p.pos) }));
        wraps = wraps.map((w) => ({ ...w, pos: tr.mapping.map(w.pos) }));
        hyphens = hyphens.map((h) => tr.mapping.map(h));
        floats = floats.map((f) => ({ ...f, pos: tr.mapping.map(f.pos) }));
        let mapped = old.decorations.map(tr.mapping, tr.doc);
        // Blocks the edit replaced (a table rebuilt when a column is resized, a restyled paragraph) lose
        // their node decorations: their layout classes are worked out again (else a table outside text
        // columns fell into them, its rows over the text).
        if (layoutClassed) {
          const from = tr.before.content.findDiffStart(tr.doc.content);
          if (from !== null) {
            const to = tr.before.content.findDiffEnd(tr.doc.content)?.b ?? tr.doc.content.size;
            const $from = tr.doc.resolve(Math.min(from, tr.doc.content.size));
            const $to = tr.doc.resolve(Math.min(Math.max(to, from), tr.doc.content.size));
            const a = $from.depth ? $from.before(1) : from;
            const b = $to.depth ? $to.after(1) : Math.max(to, from);
            mapped = mapped.remove(mapped.find(a, b, (spec) => spec.layoutClass)).add(tr.doc, layoutClasses(tr.doc, a, b));
          }
        }
        return { decorations: mapped, ...(old.columns ? { columns: old.columns } : {}) };
      },
    },
    props: {
      decorations: (state) => paginationKey.getState(state)?.decorations,
      // Text columns: the editor is a multi-column container (see screenColumns).
      attributes: (state): Record<string, string> => {
        const c = paginationKey.getState(state)?.columns;
        return c ? { class: "pw-columns", style: `column-count:${c.count};column-gap:${c.gap}pt;column-fill:balance${c.rule ? ";column-rule:0.5pt solid #000" : ""}` } : {};
      },
    },
    view(view) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      let fitTimer: ReturnType<typeof setTimeout> | undefined;
      const sheet = document.createElement("style");
      sheet.dataset.primePages = "";
      document.head.appendChild(sheet);
      // A timeout, not requestAnimationFrame: frames do not run in hidden or headless windows, and
      // reading positions forces layout anyway.
      // The next layout waits for a pause in typing: `delay` ms with no edit and none of our own work.
      // (A long fit blocks the page; when it ended, the layout's timer had already run out and fired
      // before the next keystroke could put it off, so a long document laid itself out again and
      // again mid-sentence.) The pause also grows with what the last layout cost, up to 2 s.
      let lastEdit = 0;
      let busyUntil = 0;
      let cost = 0;
      const pause = () => Math.max(opts.delay ?? 250, Math.min(2000, cost));
      const schedule = () => {
        clearTimeout(timer);
        const due = Math.max(lastEdit, busyUntil) + pause() - performance.now();
        timer = setTimeout(() => (Math.max(lastEdit, busyUntil) + pause() - performance.now() > 1 ? schedule() : relayout()), Math.max(0, due));
      };
      // What the next fit must cover: everything, or (after an edit) only what follows a position.
      let want: number | "full" | undefined;
      const refit = (from?: number) => {
        clearTimeout(fitTimer);
        want = from === undefined || want === "full" ? "full" : Math.min(want ?? from, from);
        fitTimer = setTimeout(() => {
          const w = want;
          want = undefined;
          timed("fit", () => { if (w === "full" || w === undefined || !fastFit(view, sheet, w)) fit(view, sheet); });
          busyUntil = performance.now();
        }, 0);
      };
      const relayout = () => {
        const began = performance.now();
        queueMicrotask(() => { cost = performance.now() - began; busyUntil = performance.now(); });
        const doc = timed("toModel", () => fromProseMirror(view.state.doc, opts.base()));
        const result = timed("layout", () => layoutDocument(doc, opts.measurer));
        layout = result;
        tops = pageTops(result.pages);
        timed("onLayout", () => opts.onLayout?.(result));
        const next = timed("pageStarts", () => resolvePageStarts(view.state.doc, result.pageStarts));
        const nextWraps = resolveWraps(view.state.doc, result.wraps);
        const nextHyphens = resolveHyphens(view.state.doc, result.hyphens ?? []);
        const nextFloats = resolveFloats(view.state.doc, result.floats);
        // Corrections stand where the layout did not change: a page whose start and whose previous
        // page's start are where the last layout put them keeps its correction (an edit in chapter 10
        // leaves chapter 1's pages alone; starting all corrections again after every pause made a long
        // book correct hundreds of breaks each time). If nothing changed, nothing is drawn again.
        const id = (p: ResolvedStart | undefined) => (p ? `${p.kind}:${p.pos}` : "");
        const was = new Map(raw.map((p) => [p.page, id(p)]));
        const nextAt = new Map(next.map((p) => [p.page, id(p)]));
        const fixed = new Map(starts.map((p) => [p.page, p]));
        const kept = (page: number) => (was.get(page) ?? "") === (nextAt.get(page) ?? "");
        const merged = next.map((p) => (kept(p.page) && kept(p.page - 1) && fixed.get(p.page)) || p);
        const same = merged.length === starts.length && merged.every((p, i) => id(p) === id(starts[i]) && p.page === starts[i].page)
          && JSON.stringify(nextWraps) === JSON.stringify(wraps) && JSON.stringify(nextFloats) === JSON.stringify(floats)
          && nextHyphens.length === hyphens.length && nextHyphens.every((h, i) => h === hyphens[i])
          && JSON.stringify(columnsIn(view.state.doc) ?? null) === JSON.stringify(paginationKey.getState(view.state)?.columns ?? null);
        raw = next;
        wraps = nextWraps;
        hyphens = nextHyphens;
        floats = nextFloats;
        rounds = 0; // the loop guard counts per layout
        unfixable = new Set();
        if (!same) {
          starts = merged;
          timed("drawBreaks", () => view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false)));
        }
        refit();
      };
      timer = setTimeout(relayout, 0);
      const onResize = () => refit();
      window.addEventListener("resize", onResize);
      document.fonts?.addEventListener?.("loadingdone", onResize);
      return {
        update(v, prev) {
          if (v.state.doc !== prev.doc) {
            lastEdit = performance.now();
            schedule();
          }
          if (v.state.doc !== prev.doc) {
            // An edit moves only what follows it (see fastFit).
            const from = prev.doc.content.findDiffStart(v.state.doc.content);
            if (from !== null) refit(from);
          } else if (paginationKey.getState(v.state) !== paginationKey.getState(prev)) refit();
        },
        destroy() {
          clearTimeout(timer);
          clearTimeout(fitTimer);
          sheet.remove();
          window.removeEventListener("resize", onResize);
          document.fonts?.removeEventListener?.("loadingdone", onResize);
        },
      };
    },
  });
}
