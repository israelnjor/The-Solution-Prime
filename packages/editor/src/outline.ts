/**
 * Outline view (View → Outline): the document as an outline. Each paragraph is indented by its outline
 * level (body text under the heading it follows) with Word's symbol before it: ⊕ a heading with text
 * under it, ⊖ an empty heading, ○ body text. Levels below "Show Level" are hidden; a heading's ⊕
 * clicked hides or shows what is under it. Promote / Demote change a paragraph's heading level (its
 * Heading style); Move Up / Down move paragraphs (blocks.ts).
 */
import type { ParagraphProps, PrimeDocument, Style } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { outlineLevelOf } from "./references.js";

/** 1–9 a heading level; 10 body text. */
export const BODY = 10;

interface OutlineState {
  /** Levels shown: 1–9, or BODY for all. */
  show: number;
  /** Collapsed headings, by position. */
  collapsed: number[];
  decorations: DecorationSet;
}

export const outlineKey = new PluginKey<OutlineState>("outline");

/** A top-level paragraph's outline level (1–9) or BODY. Tables and other blocks count as body text. */
function levelOf(node: Node, doc: PrimeDocument): number {
  if (node.type.name !== "paragraph") return BODY;
  const l = outlineLevelOf(doc, node.attrs.props as ParagraphProps);
  return l !== undefined && l < 9 ? l + 1 : BODY;
}

function build(state: { doc: Node }, base: PrimeDocument, show: number, collapsed: number[]): DecorationSet {
  const decos: Decoration[] = [];
  const doc = state.doc;
  const levels: { pos: number; node: Node; level: number }[] = [];
  doc.forEach((node, pos) => levels.push({ pos, node, level: levelOf(node, base) }));
  // Indent: a heading by its level; body text one step under the heading before it.
  let heading = 0;
  let hideUnder = Infinity; // inside a collapsed heading: hide everything deeper than it
  levels.forEach(({ pos, node, level }, i) => {
    if (level <= hideUnder) hideUnder = Infinity;
    const isHeading = level < BODY;
    if (isHeading) heading = level;
    const indent = isHeading ? level - 1 : heading;
    const next = levels[i + 1];
    const hasChildren = isHeading && !!next && next.level > level;
    const symbol = isHeading ? (hasChildren ? "plus" : "minus") : "body";
    const folded = isHeading && collapsed.includes(pos) && hasChildren;
    const hidden = hideUnder !== Infinity || level > show;
    const classes = ["pw-ol", `pw-ol-${symbol}`, hidden ? "pw-ol-hidden" : "", folded ? "pw-ol-folded" : ""].filter(Boolean).join(" ");
    decos.push(Decoration.node(pos, pos + node.nodeSize, { class: classes, style: `--pw-ol-indent: ${indent * 26}px` }));
    if (folded && hideUnder === Infinity) hideUnder = level;
  });
  return DecorationSet.create(doc, decos);
}

/** The outline's indents, symbols and hidden levels (only in Outline view). */
export function outlinePlugin(base: () => PrimeDocument): Plugin<OutlineState> {
  return new Plugin<OutlineState>({
    key: outlineKey,
    state: {
      init: (_, state) => ({ show: BODY, collapsed: [], decorations: build(state, base(), BODY, []) }),
      apply(tr, old, _o, state) {
        const meta = tr.getMeta(outlineKey) as Partial<Pick<OutlineState, "show" | "collapsed">> | undefined;
        if (!meta && !tr.docChanged) return old;
        const collapsed = (meta?.collapsed ?? old.collapsed).map((p) => (meta?.collapsed ? p : tr.mapping.map(p)));
        const show = meta?.show ?? old.show;
        return { show, collapsed, decorations: build(state, base(), show, collapsed) };
      },
    },
    props: {
      decorations: (state) => outlineKey.getState(state)?.decorations,
      handleDOMEvents: {
        // A click on a heading's ⊕ (drawn to the paragraph's left) folds or unfolds it.
        mousedown(view, event) {
          const target = event.target as HTMLElement;
          const p = target.closest?.(".pw-ol-plus");
          if (!p || event.clientX >= p.getBoundingClientRect().left) return false;
          const pos = view.posAtDOM(p, 0) - 1;
          const st = outlineKey.getState(view.state);
          if (!st || pos < 0) return false;
          const collapsed = st.collapsed.includes(pos) ? st.collapsed.filter((x) => x !== pos) : [...st.collapsed, pos];
          view.dispatch(view.state.tr.setMeta(outlineKey, { collapsed }));
          event.preventDefault();
          return true;
        },
      },
    },
  });
}

/** Outlining → Show Level: headings down to `level` (BODY: everything). */
export function setOutlineLevel(level: number): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.setMeta(outlineKey, { show: level, collapsed: [] }));
    return true;
  };
}

export const outlineShown = (state: EditorState) => outlineKey.getState(state)?.show ?? BODY;

const headingStyle = (level: number): Style => ({
  id: `Heading${level}`, type: "paragraph", name: `heading ${level}`, basedOn: "Normal", next: "Normal",
  paragraph: { keepNext: true, keepLines: true, spacing: { before: level === 1 ? 12 : 2, after: 0 }, outlineLevel: level - 1 },
  run: { font: { ascii: "Calibri Light", hAnsi: "Calibri Light" }, size: Math.max(11, 17 - level), color: level <= 2 ? "2F5496" : "1F3763" },
});

/** The document's style for Heading `level` (Word's own, added when the document has none). */
function headingStyleId(doc: PrimeDocument, level: number): string {
  const found = doc.styles.styles.find((s) => s.type === "paragraph" && (s.id === `Heading${level}` || new RegExp(`^heading ${level}$`, "i").test(s.name ?? "")));
  if (found) return found.id;
  const st = headingStyle(level);
  doc.styles.styles.push(st);
  return st.id;
}

/**
 * Promote (-1) or demote (+1) the paragraphs in the selection: a heading one level up or down (Heading 9
 * demoted stays); body text promoted becomes a heading at the level of the heading before it. `to`
 * sets the level outright (1 = Heading 1, BODY = body text, Normal style).
 */
export function changeOutlineLevel(base: () => PrimeDocument, by: -1 | 1 | 0, to?: number): Command {
  return (state, dispatch) => {
    const doc = base();
    const { from, to: end } = state.selection;
    const targets: { pos: number; node: Node; level: number }[] = [];
    let before = 1;
    let seenStart = false;
    state.doc.forEach((node, pos) => {
      const level = levelOf(node, doc);
      if (pos + node.nodeSize > from && pos <= end && node.type.name === "paragraph") { targets.push({ pos, node, level }); seenStart = true; }
      if (!seenStart && level < BODY) before = level;
    });
    if (!targets.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const t of targets) {
        let next = to ?? (t.level === BODY ? (by < 0 ? before : BODY) : Math.min(9, Math.max(1, t.level + by)));
        if (to === undefined && t.level === BODY && by > 0) continue;
        if (next === t.level) continue;
        const props = { ...(t.node.attrs.props as ParagraphProps) };
        delete props.outlineLevel;
        if (next >= BODY) { next = BODY; props.style = doc.styles.styles.find((s) => s.isDefault && s.type === "paragraph")?.id ?? "Normal"; }
        else props.style = headingStyleId(doc, next);
        tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, props });
      }
      if (tr.docChanged) dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** The outline level (1–9, BODY) of the paragraph at the cursor. */
export function outlineLevelAt(state: EditorState, base: PrimeDocument): number {
  const $from = state.selection.$from;
  const node = $from.depth ? $from.node(1) : state.doc.nodeAt($from.pos);
  return node ? levelOf(node, base) : BODY;
}
