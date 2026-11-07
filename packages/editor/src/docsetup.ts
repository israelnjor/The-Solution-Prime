/**
 * Page setup (margins, orientation, paper size) and headers/footers. Section properties travel in
 * the editor document: the last section's in the doc's attributes, earlier ones in section breaks.
 */
import type { HeaderFooterKind, SectionProps, Story } from "@prime/pdm";
import { type Command, type EditorState, TextSelection, type Transaction } from "prosemirror-state";

export interface SectionAttrs {
  props: SectionProps;
  headers: Partial<Record<HeaderFooterKind, Story>>;
  footers: Partial<Record<HeaderFooterKind, Story>>;
}

export const PAPER: Record<string, { width: number; height: number }> = {
  Letter: { width: 612, height: 792 },
  A4: { width: 595.3, height: 841.9 },
  Legal: { width: 612, height: 1008 },
  A5: { width: 419.5, height: 595.3 },
  Executive: { width: 522, height: 756 },
};

const EMPTY: SectionAttrs = { props: {}, headers: {}, footers: {} };

/** Every section: where its attributes live (-1 = the document's own) and the attributes. */
function sections(state: EditorState): { pos: number; attrs: SectionAttrs }[] {
  const out: { pos: number; attrs: SectionAttrs }[] = [];
  state.doc.forEach((node, offset) => {
    if (node.type.name === "section_break") out.push({ pos: offset, attrs: node.attrs.section as SectionAttrs });
  });
  out.push({ pos: -1, attrs: (state.doc.attrs.section as SectionAttrs | null) ?? EMPTY });
  return out;
}

/** The section containing the cursor. */
function current(state: EditorState): { pos: number; attrs: SectionAttrs; index: number } {
  const all = sections(state);
  const cursor = state.selection.from;
  const index = all.findIndex((s) => s.pos === -1 || s.pos >= cursor);
  return { ...all[index], index };
}

/** The sections the selection touches, as Word's ribbon applies its settings (all of a one-section document). */
function touched(state: EditorState): { pos: number; attrs: SectionAttrs }[] {
  const all = sections(state);
  const { from, to } = state.selection;
  return all.filter((s, i) => {
    const start = i === 0 ? 0 : all[i - 1].pos;
    const end = s.pos === -1 ? state.doc.content.size : s.pos;
    return end >= from && start <= to;
  });
}

function setSection(tr: Transaction, pos: number, attrs: SectionAttrs): Transaction {
  if (pos === -1) return tr.setDocAttribute("section", attrs);
  return tr.setNodeMarkup(pos, undefined, { section: attrs });
}

/** Page setup of the section at the cursor (with Word's defaults filled in). */
export function pageSetup(state: EditorState): Required<Pick<SectionProps, "pageWidth" | "pageHeight">> & SectionProps {
  const props = current(state).attrs.props;
  return { pageWidth: 612, pageHeight: 792, ...props, margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0, ...props.margins } };
}

/**
 * Applies page setup to the sections the selection touches (the whole of a one-section document), as
 * Word's ribbon does. Orientation swaps width and height when needed.
 */
export function setPageSetup(change: { paper?: { width: number; height: number }; orientation?: "portrait" | "landscape"; margins?: SectionProps["margins"] }): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    let tr = state.tr;
    for (const s of touched(state)) {
      const p = { ...s.attrs.props };
      let w = change.paper?.width ?? p.pageWidth ?? 612;
      let h = change.paper?.height ?? p.pageHeight ?? 792;
      const orientation = change.orientation ?? (p.orientation === "landscape" || w > h ? "landscape" : "portrait");
      if ((orientation === "landscape") !== (w > h)) [w, h] = [h, w];
      p.pageWidth = w;
      p.pageHeight = h;
      if (orientation === "landscape") p.orientation = "landscape";
      else delete p.orientation;
      if (change.paper) delete p.paperCode; // the printer code belonged to the old size
      if (change.margins) p.margins = { ...p.margins, ...change.margins };
      tr = setSection(tr, s.pos, { ...s.attrs, props: p });
    }
    dispatch(tr);
    return true;
  };
}

/** How a section starts: Word's section break kinds (w:sectPr/w:type; "nextPage" is the default). */
export type SectionStart = "nextPage" | "continuous" | "evenPage" | "oddPage";

/**
 * Inserts a section break at the cursor, splitting the paragraph there (as Word does). The text before
 * it keeps the section's settings, headers and footers; the text after starts as `start` says, with
 * the same page setup, and its headers and footers continue the ones before ("Link to Previous").
 */
export function insertSectionBreak(start: SectionStart): Command {
  return (state, dispatch) => {
    const $cursor = state.selection.$from;
    if (!dispatch) return true;
    const cur = current(state);
    const tr = state.tr.deleteSelection();
    const $from = tr.doc.resolve(tr.mapping.map($cursor.pos));
    // Where the break goes, between two top-level blocks.
    let at: number;
    if ($from.depth === 1 && $from.parent.isTextblock && $from.parentOffset > 0 && $from.parentOffset < $from.parent.content.size) {
      tr.split($from.pos);
      at = $from.pos + 1;
    } else if ($from.depth >= 1 && $from.parentOffset === 0 && $from.depth === 1) {
      at = $from.before(1);
    } else {
      at = $from.depth ? $from.after(1) : $from.pos;
    }
    const before: SectionAttrs = { ...cur.attrs };
    tr.insert(at, state.schema.nodes.section_break.create({ section: before }));
    // The section after the break: same page setup, starting as asked, headers and footers inherited.
    const props = { ...cur.attrs.props };
    if (start === "nextPage") delete props.breakType;
    else props.breakType = start;
    setSection(tr, cur.pos === -1 ? -1 : tr.mapping.map(cur.pos), { props, headers: {}, footers: {} });
    // The cursor at the start of what follows the break (a new empty paragraph at the end of the document).
    let after = at + 1;
    if (after >= tr.doc.content.size || !tr.doc.resolve(after).nodeAfter?.isTextblock) {
      tr.insert(after, state.schema.nodes.paragraph.create());
    }
    tr.setSelection(TextSelection.near(tr.doc.resolve(after + 1)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Inserts a column break: the text after it starts the next column (the next page, without columns). */
export const insertColumnBreak: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const tr = state.tr.replaceSelectionWith(state.schema.nodes.hard_break.create({ kind: "column" }), false);
  tr.split(tr.selection.from);
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * The header or footer shown for the section at the cursor: its own, or the one it inherits from
 * an earlier section (Word's "Link to Previous").
 */
export function storyAt(state: EditorState, which: "header" | "footer", kind: HeaderFooterKind = "default"): Story | undefined {
  const all = sections(state);
  for (let i = current(state).index; i >= 0; i--) {
    const story = (which === "header" ? all[i].attrs.headers : all[i].attrs.footers)?.[kind];
    if (story) return story;
  }
  return undefined;
}

/** Sets the header or footer of the section at the cursor (an empty story removes it). */
export function setStory(which: "header" | "footer", story: Story | undefined, kind: HeaderFooterKind = "default"): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { pos, attrs } = current(state);
    const key = which === "header" ? "headers" : "footers";
    const stories = { ...(attrs[key] ?? {}) };
    if (story) stories[kind] = story;
    else delete stories[kind];
    dispatch(setSection(state.tr, pos, { ...attrs, [key]: stories }));
    return true;
  };
}

/** Inserts a page-number field (PAGE or NUMPAGES) at the cursor; its text is filled in when laid out. */
export function insertField(instr: "PAGE" | "NUMPAGES"): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { from } = state.selection;
    const tr = state.tr.insertText("1", from);
    tr.addMark(from, from + 1, state.schema.marks.field.create({ instr: ` ${instr} ` }));
    tr.removeStoredMark(state.schema.marks.field);
    dispatch(tr);
    return true;
  };
}

/** A field with its current result as text, e.g. DATE \@ "d MMMM yyyy" showing today's date. */
export function insertFieldResult(instr: string, result: string): Command {
  return (state, dispatch) => {
    if (!result) return false;
    if (!dispatch) return true;
    const { from, to } = state.selection;
    const tr = state.tr.insertText(result, from, to);
    tr.addMark(from, from + result.length, state.schema.marks.field.create({ instr: ` ${instr} ` }));
    tr.removeStoredMark(state.schema.marks.field);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** The section at the cursor: its 1-based number and how many sections there are. */
export function sectionInfo(state: EditorState): { number: number; count: number; titlePage: boolean } {
  const all = sections(state);
  const c = current(state);
  return { number: c.index + 1, count: all.length, titlePage: !!c.attrs.props.titlePage };
}

/** The section's own header/footer (not one inherited from an earlier section). */
export function ownStory(state: EditorState, which: "header" | "footer", kind: HeaderFooterKind = "default"): Story | undefined {
  const { attrs } = current(state);
  return (which === "header" ? attrs.headers : attrs.footers)?.[kind];
}

/** "Different first page": the section's first page uses its first-page header and footer. */
export function setTitlePage(on: boolean): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { pos, attrs } = current(state);
    const props = { ...attrs.props };
    if (on) props.titlePage = true;
    else delete props.titlePage;
    dispatch(setSection(state.tr, pos, { ...attrs, props }));
    return true;
  };
}

// --- line numbers ----------------------------------------------------------------------------------

/** Line numbering: off, or numbers restarting on each page, each section, or running on. */
export type LineNumbering = "none" | "newPage" | "newSection" | "continuous";

/** The line numbering of the section at the cursor. */
export function lineNumbersOf(state: EditorState): LineNumbering {
  const xml = current(state).attrs.props.extra?.find((x) => x.name === "w:lnNumType")?.xml;
  if (!xml || !/w:countBy="[1-9]/.test(xml)) return "none";
  // Without w:restart: a new start on each page.
  return (/w:restart="(newSection|continuous)"/.exec(xml)?.[1] as LineNumbering | undefined) ?? "newPage";
}

/** Sets line numbering for the sections the selection touches (Word's Layout → Line Numbers). */
export function setLineNumbers(mode: LineNumbering): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    let tr = state.tr;
    for (const s of touched(state)) {
      const extra = (s.attrs.props.extra ?? []).filter((x) => x.name !== "w:lnNumType");
      if (mode !== "none") extra.push({ name: "w:lnNumType", xml: `<w:lnNumType w:countBy="1" w:restart="${mode}"/>` });
      tr = setSection(tr, s.pos, { ...s.attrs, props: { ...s.attrs.props, extra } });
    }
    dispatch(tr);
    return true;
  };
}

/** Whether the paragraph at the cursor is left out of line numbering. */
export function lineNumbersSuppressed(state: EditorState): boolean {
  const props = state.selection.$from.parent.attrs.props as { extra?: { name: string }[] } | undefined;
  return !!props?.extra?.some((x) => x.name === "w:suppressLineNumbers");
}

/** Leaves the selected paragraphs out of line numbering, or puts them back (Word: "Suppress for Current Paragraph"). */
export const toggleSuppressLineNumbers: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const on = !lineNumbersSuppressed(state);
  const tr = state.tr;
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const props = { ...(node.attrs.props ?? {}) } as { extra?: { name: string; xml: string }[] };
    const extra = (props.extra ?? []).filter((x) => x.name !== "w:suppressLineNumbers");
    if (on) extra.push({ name: "w:suppressLineNumbers", xml: "<w:suppressLineNumbers/>" });
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: { ...props, extra } });
    return false;
  });
  dispatch(tr);
  return true;
};

// --- hyphenation -------------------------------------------------------------------------------------

/** Whether words are hyphenated automatically (Word's Layout → Hyphenation → Automatic). */
export function hyphenationOn(state: EditorState): boolean {
  return !!(state.doc.attrs.hyphenation as { auto?: boolean } | null)?.auto;
}

/** Turns automatic hyphenation on (Word's defaults: 0.25" zone, no limit) or off. */
export function setHyphenation(on: boolean): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const old = state.doc.attrs.hyphenation as Record<string, unknown> | null;
      dispatch(state.tr.setDocAttribute("hyphenation", on ? { ...(old ?? {}), auto: true } : null));
    }
    return true;
  };
}

// --- text columns ------------------------------------------------------------------------------------

/** Text columns as the Columns dialog sets them: their number, the space between, a line between, and
 * (when not equal) each column's width; widths and spaces in pt. */
export interface ColumnsSpec {
  count: number;
  space: number;
  line?: boolean;
  /** Each column's own width (unequal columns); omitted: equal columns. */
  widths?: number[];
}

/** The text columns of the section at the cursor (one column when it has none). */
export function columnsOf(state: EditorState): ColumnsSpec {
  const xml = current(state).attrs.props.extra?.find((x) => x.name === "w:cols")?.xml;
  if (!xml) return { count: 1, space: 36 };
  const open = /<w:cols\b[^>]*>/.exec(xml)?.[0] ?? xml;
  const line = /w:sep="(1|true|on)"/.test(open);
  const space = Number(/w:space="(\d+)"/.exec(open)?.[1] ?? 720) / 20;
  if (/w:equalWidth="(0|false|off)"/.test(open)) {
    const widths = [...xml.matchAll(/<w:col\b([^>]*)\/?>/g)].map((m) => Number(/w:w="(\d+)"/.exec(m[1])?.[1] ?? 0) / 20);
    if (widths.length > 1) return { count: widths.length, space, line, widths };
  }
  return { count: Math.max(1, Number(/w:num="(\d+)"/.exec(open)?.[1] ?? 1)), space, line };
}

/** Word's column presets for a text width: One, Two, Three, Left (a narrow first column), Right. */
export function columnPreset(name: "one" | "two" | "three" | "left" | "right", textWidth: number): ColumnsSpec {
  const space = 36;
  if (name === "one") return { count: 1, space };
  if (name === "two") return { count: 2, space };
  if (name === "three") return { count: 3, space };
  const narrow = Math.round(((textWidth - space) / 3) * 10) / 10;
  const wide = Math.round((textWidth - space - narrow) * 10) / 10;
  return { count: 2, space, widths: name === "left" ? [narrow, wide] : [wide, narrow] };
}

/**
 * Sets text columns, as Word's Layout → Columns does: with text selected, for the paragraphs the
 * selection covers only — continuous section breaks go before and after them (where they do not
 * already start or end their section), so the text around keeps its own columns; with nothing
 * selected (or a selection across sections), for the sections the selection touches.
 */
export function setColumns(spec: ColumnsSpec): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tw = (pt: number) => Math.round(pt * 20);
    const count = Math.max(1, Math.min(12, Math.round(spec.count)));
    const xml = spec.widths && spec.widths.length > 1
      ? `<w:cols w:num="${spec.widths.length}"${spec.line ? ' w:sep="1"' : ""} w:space="${tw(spec.space)}" w:equalWidth="0">${spec.widths
        .map((w, i) => `<w:col w:w="${tw(w)}"${i < spec.widths!.length - 1 ? ` w:space="${tw(spec.space)}"` : ""}/>`).join("")}</w:cols>`
      : `<w:cols${count > 1 ? ` w:num="${count}"` : ""}${spec.line && count > 1 ? ' w:sep="1"' : ""} w:space="${tw(spec.space)}"/>`;
    const withColumns = (props: SectionProps): SectionProps => ({ ...props, extra: [...(props.extra ?? []).filter((x) => x.name !== "w:cols"), { name: "w:cols", xml }] });
    const all = sections(state);
    const { from, to, empty } = state.selection;
    const section = all.findIndex((s) => s.pos === -1 || s.pos >= from);
    const own = all[section];
    const start = section === 0 ? 0 : all[section - 1].pos + state.doc.child(state.doc.resolve(all[section - 1].pos).index(0)).nodeSize;
    const end = own.pos === -1 ? state.doc.content.size : own.pos;
    if (!empty && to <= end) {
      // The whole paragraphs (or tables) the selection covers.
      const $from = state.doc.resolve(from);
      const $to = state.doc.resolve(to);
      const blockFrom = $from.depth ? $from.before(1) : from;
      const blockTo = $to.depth ? $to.after(1) : to;
      const hasBefore = blockFrom > start;
      const hasAfter = blockTo < end;
      const s = own.attrs;
      // The selected part: the new columns; it starts on the same page (continuous) unless it starts the section.
      const middle: SectionProps = { ...withColumns(s.props), ...(hasBefore ? { breakType: "continuous" } : {}) };
      if (!hasBefore && s.props.breakType === undefined) delete middle.breakType;
      const middleAttrs: SectionAttrs = hasBefore ? { props: middle, headers: {}, footers: {} } : { ...s, props: middle };
      const tr = state.tr;
      if (hasAfter) {
        // The break after the selection ends the selected part; what follows keeps the section's own settings.
        tr.insert(blockTo, state.schema.nodes.section_break.create({ section: middleAttrs }));
        setSection(tr, own.pos === -1 ? -1 : tr.mapping.map(own.pos), { props: { ...s.props, breakType: "continuous" }, headers: {}, footers: {} });
      } else {
        setSection(tr, own.pos === -1 ? -1 : tr.mapping.map(own.pos), middleAttrs);
      }
      // The break before the selection ends the part before it, as the section was.
      if (hasBefore) tr.insert(blockFrom, state.schema.nodes.section_break.create({ section: { ...s } }));
      dispatch(tr.scrollIntoView());
      return true;
    }
    let tr = state.tr;
    for (const s of touched(state)) tr = setSection(tr, s.pos, { ...s.attrs, props: withColumns(s.attrs.props) });
    dispatch(tr);
    return true;
  };
}
