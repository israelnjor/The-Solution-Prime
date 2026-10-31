/**
 * Find and Replace. Matches are found within each paragraph's text (inline objects such as pictures
 * and equations break a match); they are highlighted with decorations.
 */
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, Plugin, PluginKey, TextSelection } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

export interface FindQuery {
  text: string;
  matchCase?: boolean;
  wholeWord?: boolean;
}

export interface Match {
  from: number;
  to: number;
}

const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** All matches in document order. */
export function findAll(doc: Node, q: FindQuery): Match[] {
  if (!q.text) return [];
  const pattern = new RegExp(q.wholeWord ? `\\b${escape(q.text)}\\b` : escape(q.text), q.matchCase ? "g" : "gi");
  const out: Match[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Map the paragraph's text to positions; non-text inline nodes become a separator character.
    let text = "";
    const at: number[] = [];
    node.forEach((child, offset) => {
      const s = child.isText ? child.text! : "￼";
      for (let i = 0; i < s.length; i++) at.push(pos + 1 + offset + i);
      text += s;
    });
    for (const m of text.matchAll(pattern)) {
      if (!m[0].length || m[0].includes("￼")) continue;
      out.push({ from: at[m.index!], to: at[m.index! + m[0].length - 1] + 1 });
    }
    return false;
  });
  return out;
}

interface FindState {
  query: FindQuery;
  matches: Match[];
  current: number;
  decorations: DecorationSet;
}

export const findKey = new PluginKey<FindState>("find");

function build(doc: Node, query: FindQuery, current: number): FindState {
  const matches = findAll(doc, query);
  const decorations = DecorationSet.create(doc, matches.map((m, i) =>
    Decoration.inline(m.from, m.to, { class: i === current ? "pw-find pw-find-current" : "pw-find" })));
  return { query, matches, current: Math.min(current, Math.max(0, matches.length - 1)), decorations };
}

/** Highlights matches of the query set with `setFindQuery`. */
export function findPlugin(): Plugin<FindState> {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: (_, state) => build(state.doc, { text: "" }, 0),
      apply(tr, old) {
        const meta = tr.getMeta(findKey) as { query?: FindQuery; current?: number } | undefined;
        if (meta) return build(tr.doc, meta.query ?? old.query, meta.current ?? old.current);
        return tr.docChanged ? build(tr.doc, old.query, old.current) : old;
      },
    },
    props: { decorations: (state) => findKey.getState(state)?.decorations },
  });
}

export const findState = (state: EditorState) => findKey.getState(state);

export function setFindQuery(query: FindQuery): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.setMeta(findKey, { query, current: 0 }));
    return true;
  };
}

/** Selects the next (+1) or previous (-1) match after the cursor. */
export function findNext(dir: 1 | -1 = 1): Command {
  return (state, dispatch) => {
    const fs = findKey.getState(state);
    if (!fs?.matches.length) return false;
    const pos = dir > 0 ? state.selection.to : state.selection.from;
    let index = dir > 0 ? fs.matches.findIndex((m) => m.from >= pos) : -1;
    if (dir < 0) for (let i = fs.matches.length - 1; i >= 0; i--) if (fs.matches[i].to <= pos) { index = i; break; }
    if (index < 0) index = dir > 0 ? 0 : fs.matches.length - 1; // wrap around
    const m = fs.matches[index];
    if (dispatch) {
      dispatch(state.tr.setSelection(TextSelection.create(state.doc, m.from, m.to)).setMeta(findKey, { current: index }).scrollIntoView());
    }
    return true;
  };
}

/** Replaces the selected match (if the selection is one) and moves to the next. */
export function replaceCurrent(replacement: string): Command {
  return (state, dispatch, view) => {
    const fs = findKey.getState(state);
    const { from, to } = state.selection;
    const hit = fs?.matches.find((m) => m.from === from && m.to === to);
    if (!hit) return findNext(1)(state, dispatch, view);
    if (dispatch) {
      const tr = state.tr.insertText(replacement, from, to);
      dispatch(tr.setSelection(TextSelection.create(tr.doc, from + replacement.length)));
    }
    return true;
  };
}

/** Replaces every match in one undoable step; returns the count through `onCount`. */
export function replaceAll(replacement: string, onCount?: (n: number) => void): Command {
  return (state, dispatch) => {
    const fs = findKey.getState(state);
    if (!fs?.matches.length) { onCount?.(0); return false; }
    if (dispatch) {
      const tr = state.tr;
      for (const m of [...fs.matches].reverse()) tr.insertText(replacement, m.from, m.to);
      dispatch(tr.scrollIntoView());
    }
    onCount?.(fs.matches.length);
    return true;
  };
}
