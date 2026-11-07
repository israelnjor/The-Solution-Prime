/**
 * References, as Word does them, with Word's own fields so Word can update them too:
 *  - a table of contents: a TOC field whose entries (styles TOC 1–9) link to bookmarks on the
 *    headings and show their pages with PAGEREF fields;
 *  - captions ("Figure 1: …") numbered by SEQ fields, and a table of figures (TOC \c "Figure");
 *  - cross-references: REF (the heading's text, or "Figure 2") and PAGEREF (its page) to a bookmark;
 *  - Update Field (F9): contents rebuilt, captions renumbered, references refreshed.
 * Pages come from the layout (`pages`: the page of each bookmark).
 */
import { Resolver } from "@prime/layout";
import type { Block, Inline, Paragraph, ParagraphProps, PrimeDocument, Style } from "@prime/pdm";
import type { Node, Schema } from "prosemirror-model";
import { type Command, type EditorState, TextSelection, type Transaction } from "prosemirror-state";
import { bibliographyBlocks, citationText, citedTags, readSources } from "./citations.js";
import { storyToProseMirror } from "./convert.js";

export type Pages = Record<string, number>;

// --- bookmarks -------------------------------------------------------------------------------

const BOOKMARK = /<w:bookmarkStart\b[^>]*w:name="([^"]+)"/;
const bookmarkName = (n: Node) => (n.type.name === "inline_preserved" && n.attrs.preserved?.name === "w:bookmarkStart" ? BOOKMARK.exec(n.attrs.preserved.xml)?.[1] : undefined);

/** The next free bookmark id (Word needs them unique and numeric). */
function nextBookmarkId(doc: Node): number {
  let max = 0;
  doc.descendants((n) => {
    if (n.type.name === "inline_preserved" && /^w:bookmark(Start|End)$/.test(n.attrs.preserved?.name ?? "")) {
      max = Math.max(max, Number(/w:id="(\d+)"/.exec(n.attrs.preserved.xml)?.[1] ?? 0));
    }
  });
  return max + 1;
}

const randomName = (prefix: string) => `${prefix}${String(Math.floor(100000000 + Math.random() * 899999999))}`;

/**
 * A bookmark (name starting with `prefix`) on a paragraph, from its start to `end` (default: its
 * end); an existing one with that prefix is reused. Returns the name.
 */
function ensureBookmark(tr: Transaction, paraPos: number, prefix: string, endOffset?: number): string {
  const para = tr.doc.nodeAt(paraPos)!;
  let found: string | undefined;
  para.forEach((n) => { const name = bookmarkName(n); if (name?.startsWith(prefix) && !found) found = name; });
  if (found) return found;
  const id = nextBookmarkId(tr.doc);
  const name = randomName(prefix);
  const s = tr.doc.type.schema;
  const node = (pname: string, xml: string) => s.nodes.inline_preserved.create({ preserved: { name: pname, xml } });
  tr.insert(paraPos + 1 + (endOffset ?? para.content.size), node("w:bookmarkEnd", `<w:bookmarkEnd w:id="${id}"/>`));
  tr.insert(paraPos + 1, node("w:bookmarkStart", `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`));
  return name;
}

/** Where a bookmark is: its paragraph (position and node) and the text it covers. */
function findBookmark(doc: Node, name: string): { pos: number; node: Node; text: string } | undefined {
  let out: { pos: number; node: Node; text: string } | undefined;
  doc.descendants((n, pos) => {
    if (out) return false;
    if (!n.isTextblock) return true;
    let start = -1;
    let end = -1;
    n.forEach((c, offset) => {
      if (bookmarkName(c) === name) start = offset + c.nodeSize;
      const id = start >= 0 && c.type.name === "inline_preserved" && c.attrs.preserved?.name === "w:bookmarkEnd" ? offset : -1;
      if (id >= 0 && end < 0) end = id;
    });
    if (start >= 0) out = { pos, node: n, text: n.textBetween(start, end >= 0 ? end : n.content.size, undefined, "") };
    return false;
  });
  return out;
}

// --- fields ----------------------------------------------------------------------------------

/** A complex field in the document: positions of its begin, separate and end markers. */
export interface FieldAt { instr: string; begin: number; separate?: number; end: number }

/** Every complex field (w:fldChar … w:instrText … w:fldChar), outermost and nested, in order. */
export function fields(doc: Node): FieldAt[] {
  const out: FieldAt[] = [];
  const open: { instr: string; begin: number; separate?: number }[] = [];
  doc.descendants((n, pos) => {
    if (n.type.name === "field_char") {
      if (n.attrs.kind === "begin") open.push({ instr: "", begin: pos });
      else if (n.attrs.kind === "separate" && open.length) open[open.length - 1].separate = pos;
      else if (n.attrs.kind === "end" && open.length) out.push({ ...open.pop()!, end: pos });
    } else if (n.type.name === "instr" && open.length && open[open.length - 1].separate === undefined) {
      open[open.length - 1].instr += n.attrs.text;
    }
    return true;
  });
  return out.sort((a, b) => a.begin - b.begin);
}

const kindOf = (instr: string) => instr.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
const argOf = (instr: string) => instr.trim().split(/\s+/)[1]?.replace(/^"|"$/g, "");

/** The inlines of a complex field with its cached result. */
const fieldInlines = (instr: string, result: Inline[]): Inline[] => [
  { type: "fieldChar", kind: "begin" }, { type: "instr", text: ` ${instr} ` }, { type: "fieldChar", kind: "separate" },
  ...result, { type: "fieldChar", kind: "end" },
];

/** A field's nodes (with `marks`), for inserting at the cursor. */
function fieldNodes(schema: Schema, instr: string, result: string, marks: readonly import("prosemirror-model").Mark[]): Node[] {
  return [
    schema.nodes.field_char.create({ kind: "begin", props: {} }, null, marks),
    schema.nodes.instr.create({ text: ` ${instr} `, props: {} }, null, marks),
    schema.nodes.field_char.create({ kind: "separate", props: {} }, null, marks),
    ...(result ? [schema.text(result, marks)] : []),
    schema.nodes.field_char.create({ kind: "end", props: {} }, null, marks),
  ];
}

/** Replaces a field's cached result with `text` (formatting of its first character kept); positions in tr.doc. */
function setResult(tr: Transaction, f: FieldAt, text: string) {
  if (f.separate === undefined) return;
  const from = f.separate + 1;
  const to = f.end;
  if (tr.doc.textBetween(from, to, undefined, "") === text) return;
  const marks = tr.doc.resolve(Math.min(from + 1, to)).marks();
  if (text) tr.replaceWith(from, to, tr.doc.type.schema.text(text, marks));
  else tr.delete(from, to);
}

// --- styles ----------------------------------------------------------------------------------

function ensureStyle(doc: PrimeDocument, style: Style) {
  if (!doc.styles.styles.some((s) => s.id === style.id)) doc.styles.styles.push(style);
}

const tocStyle = (level: number): Style => ({
  id: `TOC${level}`, type: "paragraph", name: `toc ${level}`, basedOn: "Normal", next: "Normal",
  paragraph: { spacing: { after: 5 }, ...(level > 1 ? { indent: { left: 11 * (level - 1) } } : {}) },
});

/** Word's Caption style: small, italic, in the theme's dark blue-grey. */
const CAPTION_STYLE: Style = {
  id: "Caption", type: "paragraph", name: "caption", basedOn: "Normal", next: "Normal",
  paragraph: { spacing: { after: 10, line: 240, lineRule: "auto" } }, run: { italic: true, size: 9, color: "44546A" },
};
const TABLE_OF_FIGURES_STYLE: Style = { id: "TableofFigures", type: "paragraph", name: "table of figures", basedOn: "Normal", next: "Normal", paragraph: { spacing: { after: 0 } } };
const TOC_HEADING_STYLE: Style = { id: "TOCHeading", type: "paragraph", name: "TOC Heading", basedOn: "Heading1", next: "Normal", paragraph: { outlineLevel: 9 } };

// --- headings and captions -------------------------------------------------------------------

/** A paragraph's outline level (0 = Heading 1), from its own formatting or its style chain. */
/** A paragraph's outline level (0 = Heading 1 … 8), from its own setting or its style's; undefined: body text. */
export function outlineLevelOf(doc: PrimeDocument, props: ParagraphProps): number | undefined {
  if (props.outlineLevel !== undefined) return props.outlineLevel;
  const seen = new Set<string>();
  for (let id = props.style; id && !seen.has(id);) {
    seen.add(id);
    const st = doc.styles.styles.find((s) => s.id === id);
    if (!st) break;
    if (st.paragraph?.outlineLevel !== undefined) return st.paragraph.outlineLevel;
    const m = /^heading (\d)$/i.exec(st.name ?? "") ?? /^Heading(\d)$/.exec(st.id);
    if (m) return Number(m[1]) - 1;
    id = st.basedOn;
  }
  return undefined;
}

export interface Target { pos: number; node: Node; text: string; level: number; label?: string; number?: string }

/** Headings (outline levels 1…`max`), in order, with their list numbers ("1.2") if numbered. */
export function headings(state: EditorState | Node, doc: PrimeDocument, max = 9): Target[] {
  const out: Target[] = [];
  const resolver = new Resolver(doc);
  const root = "doc" in state ? state.doc : state;
  root.descendants((n, pos) => {
    if (n.type.name !== "paragraph") return true;
    const props = n.attrs.props as ParagraphProps;
    // Every paragraph through the resolver, in order, so list numbers count as in the document.
    const rp = resolver.paragraph({ type: "paragraph", props, content: [] } as Paragraph);
    const level = outlineLevelOf(doc, props);
    const text = n.textContent.trim();
    if (level !== undefined && level < max && text && !/^TOC/i.test(props.style ?? "")) {
      out.push({ pos, node: n, level: level + 1, text: rp.listLabel ? `${rp.listLabel.trim()} ${text}` : text });
    }
    return false;
  });
  return out;
}

/** Caption paragraphs (with a SEQ field), optionally of one label ("Figure", "Table"). */
export function captions(state: EditorState | Node, label?: string): Target[] {
  const out: Target[] = [];
  const root = "doc" in state ? state.doc : state;
  const seqs = fields(root).filter((f) => kindOf(f.instr) === "SEQ" && (!label || argOf(f.instr) === label));
  for (const f of seqs) {
    const $p = root.resolve(f.begin);
    const pos = $p.before();
    if (out.some((c) => c.pos === pos)) continue;
    const node = $p.parent;
    out.push({
      pos, node, level: 1, text: node.textContent.trim(), label: argOf(f.instr),
      number: f.separate !== undefined ? root.textBetween(f.separate + 1, f.end, undefined, "") : "",
    });
  }
  return out;
}

// --- tables of contents ----------------------------------------------------------------------

export interface TocOptions {
  /** Headings from level 1 to this one (TOC \o "1-n"). */
  levels?: number;
  /** A table of figures instead: captions with this label (TOC \c "Figure"). */
  caption?: string;
}

const tocInstr = (o: TocOptions) => (o.caption ? `TOC \\h \\z \\c "${o.caption}"` : `TOC \\o "1-${o.levels ?? 3}" \\h \\z \\u`);
const parseToc = (instr: string): TocOptions => {
  const c = /\\c\s+"([^"]+)"/.exec(instr)?.[1];
  if (c) return { caption: c };
  const o = /\\o\s+"\d+-(\d+)"/.exec(instr)?.[1];
  return { levels: o ? Number(o) : 3 };
};

/**
 * The entries of a contents table (or table of figures) as paragraphs, bookmarking each heading
 * (or caption) on the way, inside one TOC field.
 */
function tocBlocks(tr: Transaction, doc: PrimeDocument, o: TocOptions, pages: Pages, width: number): Block[] {
  const targets = o.caption ? captions(tr.doc, o.caption) : headings(tr.doc, doc, o.levels ?? 3);
  const since = tr.steps.length;
  ensureStyle(doc, o.caption ? TABLE_OF_FIGURES_STYLE : tocStyle(1));
  const tabs = { name: "w:tabs", xml: `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${Math.round(width * 20)}"/></w:tabs>` };
  const blocks: Block[] = targets.map((t, i) => {
    const name = ensureBookmark(tr, tr.mapping.slice(since).map(t.pos), "_Toc");
    const level = o.caption ? 0 : t.level;
    if (level) ensureStyle(doc, tocStyle(level));
    const page = String(pages[name] ?? "");
    const content: Inline[] = [
      ...(i === 0 ? [{ type: "fieldChar", kind: "begin" }, { type: "instr", text: ` ${tocInstr(o)} ` }, { type: "fieldChar", kind: "separate" }] as Inline[] : []),
      { type: "link", anchor: name, content: [{ type: "text", text: t.text }, { type: "tab" }, ...fieldInlines(`PAGEREF ${name} \\h`, page ? [{ type: "text", text: page }] : [])] },
      ...(i === targets.length - 1 ? [{ type: "fieldChar", kind: "end" }] as Inline[] : []),
    ];
    return { type: "paragraph", props: { style: o.caption ? "TableofFigures" : `TOC${level}`, extra: [tabs] }, content };
  });
  if (!blocks.length) {
    const what = o.caption ? `${o.caption.toLowerCase()} captions` : "headings";
    blocks.push({ type: "paragraph", props: {}, content: [...fieldInlines(tocInstr(o), [{ type: "text", text: `No ${what} found: use the Heading styles${o.caption ? " or Insert Caption" : ""}, then update this table.` }])] });
  }
  return blocks;
}

/** The text column's width (pt) for the right-aligned page numbers. */
function columnWidth(state: EditorState): number {
  const p = (state.doc.attrs.section?.props ?? {}) as { pageWidth?: number; margins?: { left?: number; right?: number } };
  return (p.pageWidth ?? 612) - (p.margins?.left ?? 72) - (p.margins?.right ?? 72);
}

/** Inserts a table of contents (or of figures) at the cursor, with a "Contents" heading. */
export function insertToc(doc: () => PrimeDocument, o: TocOptions, pages: Pages = {}, heading = !o.caption): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const d = doc();
    const tr = state.tr;
    const blocks = tocBlocks(tr, d, o, pages, columnWidth(state));
    if (heading) {
      ensureStyle(d, TOC_HEADING_STYLE);
      blocks.unshift({ type: "paragraph", props: { style: "TOCHeading" }, content: [{ type: "text", text: "Contents" }] });
    }
    const nodes = storyToProseMirror(blocks, state.schema).content;
    // As whole paragraphs, before the paragraph at the cursor (in its place if it is empty).
    const $c = tr.doc.resolve(tr.mapping.map(state.selection.from));
    const depth = $c.depth;
    if (depth === 0) tr.insert($c.pos, nodes); // between blocks
    else if ($c.parent.content.size === 0) tr.replaceWith($c.before(depth), $c.after(depth), nodes);
    else tr.insert($c.before(depth), nodes);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// --- captions --------------------------------------------------------------------------------

/**
 * Inserts a caption ("Figure 3: text", numbered by a SEQ field) as a paragraph below (or above)
 * the block at the cursor, then renumbers the captions after it.
 */
export function insertCaption(doc: () => PrimeDocument, label: string, text: string, where: "below" | "above" = "below"): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    ensureStyle(doc(), CAPTION_STYLE);
    const $c = state.selection.$from;
    // The top-level block (a picture's paragraph, a table) the caption belongs to.
    const depth = state.selection instanceof TextSelection || $c.depth > 0 ? 1 : 0;
    const at = depth ? (where === "below" ? $c.after(depth) : $c.before(depth)) : state.selection.to;
    const blocks: Block[] = [{
      type: "paragraph", props: { style: "Caption" },
      content: [{ type: "text", text: `${label} ` }, ...fieldInlines(`SEQ ${label} \\* ARABIC`, [{ type: "text", text: "1" }]), ...(text ? [{ type: "text", text: `: ${text}` } as Inline] : [])],
    }];
    const tr = state.tr.insert(at, storyToProseMirror(blocks, state.schema).content);
    renumber(tr, label);
    const caption = tr.doc.nodeAt(at);
    if (caption) tr.setSelection(TextSelection.near(tr.doc.resolve(at + caption.nodeSize - 1)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** SEQ fields numbered in document order (per label, or every label). */
function renumber(tr: Transaction, label?: string) {
  const count = new Map<string, number>();
  const todo: [FieldAt, string][] = [];
  for (const f of fields(tr.doc)) {
    if (kindOf(f.instr) !== "SEQ") continue;
    const l = argOf(f.instr) ?? "";
    const n = (count.get(l) ?? 0) + 1;
    count.set(l, n);
    if (!label || l === label) todo.push([f, String(n)]);
  }
  for (const [f, n] of todo.reverse()) setResult(tr, f, n);
}

// --- cross-references ------------------------------------------------------------------------

export type RefShow = "text" | "page" | "labelNumber" | "aboveBelow";

/**
 * A cross-reference at the cursor to a heading or caption: its text (or for a caption just
 * "Figure 2"), its page, or "above"/"below". Word's REF and PAGEREF fields, to a _Ref bookmark.
 */
export function insertCrossReference(target: Target, show: RefShow, pages: Pages = {}): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    let endOffset: number | undefined;
    if (show === "labelNumber") {
      // Only "Figure 2": the bookmark ends after the caption's SEQ field.
      const seq = fields(state.doc).find((f) => kindOf(f.instr) === "SEQ" && f.begin > target.pos && f.end < target.pos + target.node.nodeSize);
      if (seq) endOffset = seq.end + 1 - (target.pos + 1);
    }
    const name = ensureBookmark(tr, target.pos, "_Ref", endOffset);
    const at = tr.mapping.map(state.selection.from);
    const marks = state.selection.$from.marks().filter((m) => m.type.name !== "field");
    const below = target.pos < state.selection.from;
    const found = findBookmark(tr.doc, name);
    const text = show === "page" ? String(pages[name] ?? "")
      : show === "aboveBelow" ? (below ? "above" : "below")
      : found?.text.trim() ?? target.text;
    const instr = show === "page" ? `PAGEREF ${name} \\h` : show === "aboveBelow" ? `PAGEREF ${name} \\p \\h` : `REF ${name} \\h`;
    tr.insert(at, fieldNodes(state.schema, instr, text, marks));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// --- updating ----------------------------------------------------------------------------------

/**
 * Update Field (F9) over the whole document: contents tables and tables of figures rebuilt from
 * the headings and captions, captions renumbered, cross-references and page numbers refreshed.
 */
export function updateFields(doc: () => PrimeDocument, pages: Pages = {}): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const d = doc();
    const tr = state.tr;
    renumber(tr);
    // Tables of contents, last first; the document is scanned again for each (bookmarks get added).
    const outerTocs = () => {
      const tocs = fields(tr.doc).filter((f) => kindOf(f.instr) === "TOC");
      return tocs.filter((f) => !tocs.some((g) => g !== f && g.begin < f.begin && g.end > f.end));
    };
    for (let k = outerTocs().length - 1; k >= 0; k--) {
      const f = outerTocs()[k];
      const $b = tr.doc.resolve(f.begin);
      const $e = tr.doc.resolve(f.end);
      if ($b.depth !== $e.depth) continue;
      const from = $b.before();
      const to = $e.after();
      const since = tr.steps.length;
      const blocks = tocBlocks(tr, d, parseToc(f.instr), pages, columnWidth(state));
      const map = tr.mapping.slice(since);
      tr.replaceWith(map.map(from, -1), map.map(to, 1), storyToProseMirror(blocks, state.schema).content);
    }
    // Reference lists, rebuilt the same way (from the citations as they now are).
    const bibs = () => fields(tr.doc).filter((f) => kindOf(f.instr) === "BIBLIOGRAPHY");
    for (let k = bibs().length - 1; k >= 0; k--) {
      const f = bibs()[k];
      const $b = tr.doc.resolve(f.begin);
      const $e = tr.doc.resolve(f.end);
      if ($b.depth !== $e.depth) continue;
      tr.replaceWith($b.before(), $e.after(), storyToProseMirror(bibliographyBlocks(state.apply(tr), d), state.schema).content);
    }
    // Citations, in the current style (IEEE numbers by order of first citation).
    const { style, sources } = readSources(d);
    const cited = citedTags(state.apply(tr));
    for (const f of fields(tr.doc).reverse()) {
      if (kindOf(f.instr) !== "CITATION") continue;
      const tag = argOf(f.instr) ?? "";
      const page = /\\p\s+(\S+)/.exec(f.instr)?.[1];
      setResult(tr, f, citationText(sources.find((s) => s.tag === tag), style, cited.indexOf(tag) + 1, page));
    }
    // REF and PAGEREF results (last first: results change length).
    for (const f of fields(tr.doc).reverse()) {
      const kind = kindOf(f.instr);
      const name = argOf(f.instr);
      if (!name || (kind !== "REF" && kind !== "PAGEREF")) continue;
      if (kind === "REF") {
        const b = findBookmark(tr.doc, name);
        if (b) setResult(tr, f, b.text.trim());
      } else if (/\\p\b/.test(f.instr)) {
        const b = findBookmark(tr.doc, name);
        if (b) setResult(tr, f, b.pos < f.begin ? "above" : "below");
      } else if (pages[name] !== undefined) setResult(tr, f, String(pages[name]));
    }
    dispatch(tr);
    return true;
  };
}

/**
 * Page numbers in PAGEREF results set to where the layout put their bookmarks — for saving, so a
 * file opened elsewhere shows the right pages without updating. Works on the document model.
 */
export function refreshPageRefs(doc: PrimeDocument, pages: Pages): void {
  // The model's blocks are shared with the editor's conversion cache: change a copy of them.
  doc.sections = structuredClone(doc.sections);
  const visit = (items: Inline[]) => {
    let open: { instr: string; result: boolean; ref?: string }[] = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it.type === "fieldChar") {
        if (it.kind === "begin") open.push({ instr: "", result: false });
        else if (it.kind === "separate" && open.length) {
          const f = open[open.length - 1];
          f.result = true;
          const m = /^\s*PAGEREF\s+(\S+)/.exec(f.instr);
          const n = m && !/\\p\b/.test(f.instr) ? pages[m[1]] : undefined;
          // The result is the text right after the separator, up to the end marker.
          if (n !== undefined && items[i + 1]?.type === "text" && items[i + 2]?.type === "fieldChar") (items[i + 1] as { text: string }).text = String(n);
        } else if (it.kind === "end") open = open.slice(0, -1);
      } else if (it.type === "instr" && open.length && !open[open.length - 1].result) open[open.length - 1].instr += it.text;
      else if ("content" in it && Array.isArray(it.content)) visit(it.content as Inline[]);
    }
  };
  const blocks = (list: Block[]) => {
    for (const b of list) {
      if (b.type === "paragraph") visit(b.content);
      else if (b.type === "table") for (const r of b.rows) for (const c of r.cells) blocks(c.blocks);
      else if (b.type === "contentControl") blocks(b.blocks);
    }
  };
  for (const s of doc.sections) blocks(s.blocks);
}

/** The bookmark a click at `pos` leads to: a contents entry's link, or a REF/PAGEREF field's target. */
export function referenceAt(state: EditorState, pos: number): string | undefined {
  const link = state.doc.resolve(pos).marks().find((m) => m.type.name === "link" && m.attrs.anchor);
  if (link) return link.attrs.anchor as string;
  const f = fields(state.doc).filter((x) => x.begin < pos && x.end >= pos && /^(REF|PAGEREF)$/.test(kindOf(x.instr))).pop();
  return f ? argOf(f.instr) : undefined;
}

/** Puts the cursor at a bookmark (Ctrl+Click on a reference). False if it is not in the document. */
export function gotoBookmark(name: string): Command {
  return (state, dispatch) => {
    const b = findBookmark(state.doc, name);
    if (!b) return false;
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(b.pos + 1))).scrollIntoView());
    return true;
  };
}
