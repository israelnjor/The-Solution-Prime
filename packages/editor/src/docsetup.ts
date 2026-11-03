/**
 * Page setup (margins, orientation, paper size) and headers/footers. Section properties travel in
 * the editor document: the last section's in the doc's attributes, earlier ones in section breaks.
 */
import type { HeaderFooterKind, SectionProps, Story } from "@prime/pdm";
import type { Command, EditorState, Transaction } from "prosemirror-state";

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
 * Applies page setup to the whole document (Word's default "Apply to: Whole document").
 * Orientation swaps width and height when needed.
 */
export function setPageSetup(change: { paper?: { width: number; height: number }; orientation?: "portrait" | "landscape"; margins?: SectionProps["margins"] }): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    let tr = state.tr;
    for (const s of sections(state)) {
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
