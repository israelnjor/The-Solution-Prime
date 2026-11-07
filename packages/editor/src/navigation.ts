/**
 * The navigation pane's outline: the document's headings, each with the stretch of the document it
 * heads (its section: up to the next heading of the same or a higher level), and moving a section,
 * as dragging a heading in Word's navigation pane does.
 */
import type { PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, TextSelection } from "prosemirror-state";
import { headings } from "./references.js";

export interface OutlineItem {
  /** The heading paragraph's position. */
  pos: number;
  /** 1 = Heading 1. */
  level: number;
  text: string;
  /** Where its section ends (the next heading of the same or a higher level, or the document's end). */
  end: number;
}

/** The headings of the document body (not those inside tables or text boxes), with their sections. */
export function outline(state: EditorState | Node, doc: PrimeDocument): OutlineItem[] {
  const root = "doc" in state ? state.doc : state;
  const list = headings(root, doc).filter((h) => root.resolve(h.pos).depth === 0);
  return list.map((h, i) => {
    const next = list.slice(i + 1).find((x) => x.level <= h.level);
    return { pos: h.pos, level: h.level, text: h.text, end: next ? next.pos : root.content.size };
  });
}

/** The heading whose section holds `pos` (the last heading at or before it), or -1. */
export function outlineIndexAt(items: OutlineItem[], pos: number): number {
  let lo = 0;
  let hi = items.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (items[mid].pos <= pos) { found = mid; lo = mid + 1; } else hi = mid - 1;
  }
  return found;
}

/**
 * Moves the section of heading `index` (the heading and everything under it) to `target`, a position
 * between top-level blocks: another heading's position (before it) or a section's end (after it).
 * Refused when the target is inside the section itself.
 */
export function moveSection(items: OutlineItem[], index: number, target: number): Command {
  return (state, dispatch) => {
    const item = items[index];
    if (!item || target >= item.pos && target <= item.end) return false;
    const { pos: from, end: to } = item;
    if (to > state.doc.content.size || target > state.doc.content.size) return false;
    if (dispatch) {
      const slice = state.doc.slice(from, to);
      const tr = state.tr.delete(from, to);
      const at = tr.mapping.map(target, target > from ? -1 : 1);
      tr.insert(at, slice.content);
      tr.setSelection(TextSelection.create(tr.doc, at + 1));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}
