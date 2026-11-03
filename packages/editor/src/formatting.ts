/**
 * Quick formatting, as on Word's Home tab: clear formatting, change case, grow/shrink font,
 * superscript/subscript, and the format painter (copy one place's formatting, apply it elsewhere).
 */
import { Resolver } from "@prime/layout";
import type { ParagraphProps, PrimeDocument, RunProps } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { updateRun } from "./commands.js";

/** Text nodes in the selection, clipped to it, with the paragraph style around each. */
function texts(state: EditorState): { node: Node; from: number; to: number; styleId?: string }[] {
  const out: { node: Node; from: number; to: number; styleId?: string }[] = [];
  for (const r of state.selection.ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos, parent) => {
      if (node.isText) out.push({ node, from: Math.max(pos, r.$from.pos), to: Math.min(pos + node.nodeSize, r.$to.pos), styleId: parent?.attrs.props?.style });
    });
  }
  return out;
}

function paragraphs(state: EditorState): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ node, pos });
  });
  return out;
}

const runOf = (node: Node): RunProps => node.marks.find((m) => m.type.name === "run")?.attrs.props ?? {};

/** Character formatting off: the text keeps only its character style (Ctrl+Space). */
export const clearCharFormatting: Command = updateRun((p) => (p.style ? { style: p.style } : {}));

/** Paragraph formatting back to the style's (Ctrl+Q): indents, spacing, alignment… (the style and list stay). */
export const clearParagraphFormatting: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const tr = state.tr;
  for (const { node, pos } of paragraphs(state)) {
    const p = node.attrs.props as ParagraphProps;
    const keep: ParagraphProps = {};
    if (p.style) keep.style = p.style;
    if (p.list) keep.list = p.list;
    const extra = p.extra?.filter((e) => !["w:tabs", "w:pBdr", "w:shd"].includes(e.name));
    if (extra?.length) keep.extra = extra;
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: keep });
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * Word's Clear All Formatting: the text back to plain Normal paragraphs — character formatting,
 * paragraph formatting, the paragraph style and lists all go.
 */
export const clearFormatting: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const { from, to, empty } = state.selection;
  const tr = state.tr;
  if (empty) tr.setStoredMarks([]);
  else tr.removeMark(from, to, state.schema.marks.run);
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const extra = (node.attrs.props as ParagraphProps).extra?.filter((e) => !["w:tabs", "w:pBdr", "w:shd"].includes(e.name));
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: extra?.length ? { extra } : {} });
    return false;
  });
  dispatch(tr.scrollIntoView());
  return true;
};

export type CaseMode = "sentence" | "lower" | "upper" | "title" | "toggle" | "cycle";

const sentence = (s: string, startsSentence: boolean) => {
  let start = startsSentence;
  return s.toLowerCase().replace(/[\p{L}\p{N}]|[.!?]/gu, (c) => {
    if (/[.!?]/.test(c)) { start = true; return c; }
    if (start) { start = false; return c.toUpperCase(); }
    return c;
  });
};
const title = (s: string, atWordStart: boolean) =>
  s.toLowerCase().replace(/[\p{L}\p{N}']+/gu, (w, i: number) => (i === 0 && !atWordStart ? w : w[0].toUpperCase() + w.slice(1)));

/**
 * Change Case: Sentence case, lowercase, UPPERCASE, Capitalize Each Word, tOGGLE cASE — or "cycle"
 * (Shift+F3): lowercase → UPPERCASE → Capitalize Each Word. Formatting is kept.
 */
export function changeCase(mode: CaseMode): Command {
  return (state, dispatch) => {
    const parts = texts(state);
    if (!parts.length || state.selection.empty) return false;
    if (!dispatch) return true;
    let m = mode;
    if (m === "cycle") {
      const sel = state.doc.textBetween(state.selection.from, state.selection.to, " ");
      m = sel === sel.toLowerCase() ? "upper" : sel === sel.toUpperCase() ? "title" : "lower";
    }
    const tr = state.tr;
    for (const p of parts) {
      const text = state.doc.textBetween(p.from, p.to);
      const $p = state.doc.resolve(p.from);
      const before = $p.parent.textBetween(0, $p.parentOffset, undefined, " ");
      const start = p.from;
      const out =
        m === "lower" ? text.toLowerCase()
        : m === "upper" ? text.toUpperCase()
        : m === "toggle" ? [...text].map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join("")
        : m === "title" ? title(text, !/[\p{L}\p{N}']$/u.test(before))
        : sentence(text, !before.trim() || /[.!?]\s*$/.test(before));
      if (out !== text) tr.replaceWith(tr.mapping.map(start), tr.mapping.map(p.to), state.schema.text(out, p.node.marks));
    }
    // Keep the same text selected.
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Word's font-size steps (Grow Font / Shrink Font go to the next one). */
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

function nextSize(size: number, dir: 1 | -1, byPoint: boolean): number {
  if (byPoint) return Math.max(1, Math.min(1638, size + dir));
  if (dir > 0) return FONT_SIZES.find((s) => s > size + 0.01) ?? Math.min(1638, Math.floor(size / 10) * 10 + 10);
  const lower = [...FONT_SIZES].reverse().find((s) => s < size - 0.01);
  return lower ?? Math.max(1, size - 1);
}

/**
 * Grow / Shrink Font (Ctrl+Shift+> / <: to the next size on Word's list; Ctrl+] / [: by one point),
 * each run from its own size — mixed sizes stay mixed.
 */
export function growFont(doc: () => PrimeDocument, dir: 1 | -1, byPoint = false): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const resolver = new Resolver(doc());
    const size = (props: RunProps, styleId?: string) => resolver.run(styleId ?? "", props).size;
    if (state.selection.empty) {
      const styleId = state.selection.$from.parent.attrs.props?.style;
      return updateRun((p) => ({ ...p, size: nextSize(size(p, styleId), dir, byPoint) }))(state, dispatch);
    }
    const run = state.schema.marks.run;
    const tr = state.tr;
    for (const t of texts(state)) {
      const p = runOf(t.node);
      tr.addMark(t.from, t.to, run.create({ props: { ...p, size: nextSize(size(p, t.styleId), dir, byPoint) } }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Superscript or subscript on, or back to normal if the selection already has it (Ctrl+Shift+= / Ctrl+=). */
export function toggleVertAlign(kind: "superscript" | "subscript"): Command {
  return (state, dispatch, view) => {
    const parts = texts(state);
    const on = state.selection.empty
      ? (state.storedMarks ?? state.selection.$from.marks()).find((m) => m.type.name === "run")?.attrs.props?.vertAlign === kind
      : parts.length > 0 && parts.every((t) => runOf(t.node).vertAlign === kind);
    return updateRun((p) => {
      const { vertAlign: _v, ...rest } = p;
      return on ? rest : { ...rest, vertAlign: kind };
    })(state, dispatch, view);
  };
}

/** Formatting copied by the format painter: the character formatting, and the paragraph's. */
export interface CopiedFormat { run: RunProps; paragraph: ParagraphProps }

/** The formatting at the start of the selection (Format Painter, Ctrl+Shift+C). */
export function copyFormat(state: EditorState): CopiedFormat {
  const $from = state.selection.$from;
  const marks = state.selection.empty ? (state.storedMarks ?? $from.marks()) : ($from.nodeAfter?.marks ?? $from.marks());
  const run = { ...(marks.find((m) => m.type.name === "run")?.attrs.props ?? {}) } as RunProps;
  const { list: _l, extra, ...paragraph } = ($from.parent.attrs.props ?? {}) as ParagraphProps;
  const tabs = extra?.filter((e) => ["w:tabs", "w:pBdr", "w:shd"].includes(e.name));
  return { run, paragraph: tabs?.length ? { ...paragraph, extra: tabs } : paragraph };
}

/**
 * Applies copied formatting (Ctrl+Shift+V, or the painter's brush): the character formatting to the
 * selected text; and, when whole paragraphs are selected (or the cursor is in one with nothing
 * selected), the paragraph formatting too — as Word does.
 */
export function pasteFormat(f: CopiedFormat): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    const run = state.schema.marks.run;
    const { from, to, empty } = state.selection;
    if (empty) tr.setStoredMarks([run.create({ props: f.run })]);
    else for (const t of texts(state)) tr.addMark(t.from, t.to, run.create({ props: f.run }));
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const whole = empty || (from <= pos + 1 && to >= pos + node.nodeSize - 1);
      if (whole) {
        const p = node.attrs.props as ParagraphProps;
        const kept = p.extra?.filter((e) => !["w:tabs", "w:pBdr", "w:shd"].includes(e.name)) ?? [];
        const extra = [...kept, ...(f.paragraph.extra ?? [])];
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: { ...f.paragraph, ...(p.list ? { list: p.list } : {}), ...(extra.length ? { extra } : {}) } });
      }
      return false;
    });
    dispatch(tr.scrollIntoView());
    return true;
  };
}
