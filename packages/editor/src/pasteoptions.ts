/**
 * Paste Options, as in Word: straight after a paste, the pasted text can be switched between Keep
 * Source Formatting (as pasted), Merge Formatting (the destination's look, keeping bold, italic,
 * underline and lists) and Keep Text Only. Typing anything else ends the choice.
 */
import type { ParagraphProps, RunProps } from "@prime/pdm";
import { Fragment, type Node, Slice } from "prosemirror-model";
import { type Command, type EditorState, Plugin, PluginKey, TextSelection } from "prosemirror-state";

export type PasteMode = "keep" | "merge" | "text";

export interface PasteState {
  from: number;
  to: number;
  /** What was pasted, as pasted (to switch back to "keep"). */
  original: Slice;
  mode: PasteMode;
}

export const pasteOptionsKey = new PluginKey<PasteState | null>("pasteOptions");

/** Remembers the last paste while nothing else changes. */
export function pasteOptions(): Plugin<PasteState | null> {
  return new Plugin<PasteState | null>({
    key: pasteOptionsKey,
    state: {
      init: () => null,
      apply(tr, value, old) {
        const set = tr.getMeta(pasteOptionsKey) as PasteState | null | undefined;
        if (set !== undefined) return set;
        if (tr.getMeta("paste") || tr.getMeta("uiEvent") === "paste") {
          const from = tr.mapping.map(old.selection.from, -1);
          const to = tr.selection.to;
          return to > from ? { from, to, original: tr.doc.slice(from, to), mode: "keep" } : null;
        }
        return tr.docChanged ? null : value;
      },
    },
  });
}

const KEPT: (keyof RunProps)[] = ["bold", "italic", "underline", "strike", "doubleStrike", "vertAlign"];

/** Merge: the destination's character formatting with the pasted bold/italic/underline; plain paragraphs (lists kept). */
function merged(slice: Slice, dest: RunProps, state: EditorState): Slice {
  const run = state.schema.marks.run;
  const map = (frag: Fragment): Fragment => {
    const out: Node[] = [];
    frag.forEach((n) => {
      if (n.isText) {
        const src = (n.marks.find((m) => m.type === run)?.attrs.props ?? {}) as RunProps;
        const props: RunProps = { ...dest };
        for (const k of KEPT) if (src[k] !== undefined) (props as Record<string, unknown>)[k] = src[k];
        out.push(n.mark([...n.marks.filter((m) => m.type !== run && m.type.name !== "link"), ...n.marks.filter((m) => m.type.name === "link"), run.create({ props })]));
      } else if (n.type.name === "paragraph") {
        const p = n.attrs.props as ParagraphProps;
        out.push(n.type.create({ ...n.attrs, props: p.list ? { list: p.list } : {} }, map(n.content), n.marks));
      } else out.push(n.copy(map(n.content)));
    });
    return Fragment.fromArray(out);
  };
  return new Slice(map(slice.content), slice.openStart, slice.openEnd);
}

/** Switches the last paste to another mode. */
export function pasteAs(mode: PasteMode): Command {
  return (state, dispatch) => {
    const p = pasteOptionsKey.getState(state);
    if (!p) return false;
    if (!dispatch) return true;
    const tr = state.tr.delete(p.from, p.to);
    const at = p.from;
    // The destination's formatting: what the text just before the paste has.
    const dest = (tr.doc.resolve(at).marks().find((m) => m.type.name === "run")?.attrs.props ?? {}) as RunProps;
    if (mode === "keep") tr.replace(at, at, p.original);
    else if (mode === "merge") tr.replace(at, at, merged(p.original, dest, state));
    else {
      const lines = p.original.content.textBetween(0, p.original.content.size, "\n", " ").split("\n");
      const marks = tr.doc.resolve(at).marks();
      let pos = at;
      lines.forEach((line, i) => {
        if (i > 0) { tr.split(pos); pos += 2; }
        if (line) { tr.insert(pos, state.schema.text(line, marks)); pos += line.length; }
      });
    }
    // The new content's end: the document grew by exactly its size (after the old paste was taken out).
    const end = at + tr.doc.content.size - (state.doc.content.size - (p.to - p.from));
    tr.setSelection(TextSelection.near(tr.doc.resolve(end), -1));
    dispatch(tr.setMeta(pasteOptionsKey, { ...p, from: at, to: end, mode }).scrollIntoView());
    return true;
  };
}
