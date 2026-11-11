/**
 * Deleting pages (Print Layout): everything on a page, from where it starts to where the next one starts,
 * as laid out now. Word has no single command for it — people select a page's text and press Delete, and
 * a blank page left by a page break at the end is the usual trouble; this does both.
 */
import { type Command, type EditorState, TextSelection } from "prosemirror-state";
import { pageAt, paginationKey } from "./paged.js";

/** Where each page starts in the document (index 0 = page 1, at 0), from the pagination as laid out now. */
export function pageStarts(state: EditorState): number[] | undefined {
  const set = paginationKey.getState(state)?.decorations;
  if (!set) return undefined;
  const starts = [0];
  for (const d of set.find()) {
    const p = (d.spec as { page?: number }).page;
    if (p && p > 1 && (starts[p - 1] === undefined || d.from < starts[p - 1])) starts[p - 1] = d.from;
  }
  // A page whose start was not marked (should not happen) is skipped rather than guessed.
  return starts.map((s, i) => s ?? starts[i - 1] ?? 0);
}

/** The pages the selection is on (first, last); undefined outside Print Layout. */
export function selectedPages(state: EditorState): { first: number; last: number; total: number } | undefined {
  const starts = pageStarts(state);
  if (!starts) return undefined;
  const { from, to } = state.selection;
  return { first: pageAt(state, from), last: pageAt(state, to), total: starts.length };
}

/**
 * Deletes pages `first`…`last` (1-based). Deleting the last pages also removes a page break just before
 * them (or the blank page would stay); deleting every page leaves one empty paragraph.
 */
export function deletePages(first: number, last: number): Command {
  return (state, dispatch) => {
    const starts = pageStarts(state);
    if (!starts || first < 1 || last < first || first > starts.length) return false;
    if (!dispatch) return true;
    const doc = state.doc;
    const end = doc.content.size;
    let from = starts[first - 1];
    const to = last >= starts.length ? end : starts[last];
    const tr = state.tr;
    if (from <= 0 && to >= end) {
      tr.replaceWith(0, end, state.schema.nodes.paragraph.create());
      tr.setSelection(TextSelection.create(tr.doc, 1));
    } else {
      if (to >= end && from > 0) {
        // The page before ends with a page break (Ctrl+Enter puts it at the end of a paragraph): it goes too.
        const before = doc.resolve(from).nodeBefore;
        const lastChild = before?.lastChild;
        if (before?.type.name === "paragraph" && lastChild?.type.name === "hard_break" && lastChild.attrs.kind === "page") from -= 2;
        else {
          // …or the break is inline, just before where the page starts.
          const inline = doc.resolve(from).nodeBefore;
          if (inline?.type.name === "hard_break" && inline.attrs.kind === "page") from -= inline.nodeSize;
        }
      }
      tr.delete(from, to);
      const at = Math.max(0, Math.min(from, tr.doc.content.size));
      tr.setSelection(TextSelection.near(tr.doc.resolve(at), to >= end ? -1 : 1));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Deletes the page the cursor is on (or every page the selection covers). */
export const deleteCurrentPage: Command = (state, dispatch) => {
  const pages = selectedPages(state);
  return pages ? deletePages(pages.first, pages.last)(state, dispatch) : false;
};
