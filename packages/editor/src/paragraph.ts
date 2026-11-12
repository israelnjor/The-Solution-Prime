/** Paragraph formatting: line spacing, spacing before/after, indents, borders. */
import type { Indent, ParagraphProps, Spacing } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

function paragraphs(state: EditorState): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ node, pos });
  });
  return out;
}

/** Rewrites the properties of every selected paragraph through `update`. */
export function updateParagraphs(update: (props: ParagraphProps) => ParagraphProps): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: update({ ...node.attrs.props }) });
      dispatch(tr);
    }
    return true;
  };
}

const clean = <T extends object>(o: T): T | undefined => {
  const entries = Object.entries(o).filter(([, v]) => v !== undefined);
  return entries.length ? (Object.fromEntries(entries) as T) : undefined;
};

/**
 * Line spacing. "multiple": 1 = single, 1.5, 2 = double (Word's auto rule, in 240ths of a line);
 * "atLeast" / "exact": a height in points.
 */
export function setLineSpacing(rule: "multiple" | "atLeast" | "exact", value: number): Command {
  return updateParagraphs((p) => {
    const line = rule === "multiple" ? Math.round(value * 240) : Math.round(value * 20);
    return { ...p, spacing: { ...p.spacing, line, lineRule: rule === "multiple" ? "auto" : rule } };
  });
}

/** Space before and/or after the paragraph, in points. */
export function setParagraphSpacing(change: { before?: number; after?: number }): Command {
  return updateParagraphs((p) => ({ ...p, spacing: clean<Spacing>({ ...p.spacing, ...change }) }));
}

/** Indents in points; firstLine > 0 is a first-line indent, < 0 a hanging indent. */
export function setIndent(change: { left?: number; right?: number; firstLine?: number }): Command {
  return updateParagraphs((p) => {
    const ind: Indent = { ...p.indent };
    if (change.left !== undefined) ind.left = change.left;
    if (change.right !== undefined) ind.right = change.right;
    if (change.firstLine !== undefined) {
      delete ind.firstLine;
      delete ind.hanging;
      if (change.firstLine > 0) ind.firstLine = change.firstLine;
      if (change.firstLine < 0) ind.hanging = -change.firstLine;
    }
    return { ...p, indent: clean<Indent>(ind) };
  });
}

export type BorderKind = "none" | "bottom" | "top" | "box";

const side = (name: string, on: boolean, color: string) =>
  on ? `<w:${name} w:val="single" w:sz="6" w:space="${name === "left" || name === "right" ? 4 : 1}" w:color="${color}"/>` : `<w:${name} w:val="nil"/>`;

/**
 * Paragraph borders. "none" writes explicit "nil" sides, so it also removes borders a style adds
 * (e.g. the line under a Title).
 */
export function setBorders(kind: BorderKind, color = "auto"): Command {
  return updateParagraphs((p) => {
    const extra = (p.extra ?? []).filter((x) => x.name !== "w:pBdr");
    const sides = { top: kind === "top" || kind === "box", left: kind === "box", bottom: kind === "bottom" || kind === "box", right: kind === "box" };
    const xml = `<w:pBdr>${side("top", sides.top, color)}${side("left", sides.left, color)}${side("bottom", sides.bottom, color)}${side("right", sides.right, color)}</w:pBdr>`;
    return { ...p, extra: [...extra, { name: "w:pBdr", xml }] };
  });
}

/** The border kind set directly on the first selected paragraph (for toolbar state). */
export function bordersAt(state: EditorState): BorderKind | undefined {
  const xml = paragraphs(state)[0]?.node.attrs.props.extra?.find((x: { name: string }) => x.name === "w:pBdr")?.xml as string | undefined;
  if (!xml) return undefined;
  const on = (n: string) => new RegExp(`<w:${n} w:val="(?!nil|none)`).test(xml);
  if (on("left") && on("right")) return "box";
  if (on("top")) return "top";
  if (on("bottom")) return "bottom";
  return "none";
}

/**
 * Increase / Decrease Indent: list items change level; other paragraphs move their left indent to the
 * next (or previous) half-inch step, as Word's buttons do (never below the margin).
 */
export function indentParagraphs(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      const STEP = 36;
      for (const { node, pos } of paras) {
        const props = { ...node.attrs.props };
        if (props.list) {
          const lvl = (props.list.level ?? 0) + delta;
          if (lvl < 0) delete props.list;
          else props.list = { ...props.list, level: Math.min(8, lvl) };
        } else {
          const left = props.indent?.left ?? 0;
          const next = delta > 0 ? (Math.floor(left / STEP + 1e-6) + 1) * STEP : Math.max(0, (Math.ceil(left / STEP - 1e-6) - 1) * STEP);
          props.indent = clean<Indent>({ ...props.indent, left: next || undefined });
          if (!props.indent) delete props.indent;
        }
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
      }
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Tab, as in Word: at the start of a list item it moves the item a level deeper; with several
 * paragraphs selected it indents them; otherwise it types a tab (to the next tab stop).
 */
export const tabKey: Command = (state, dispatch) => {
  const { $from, $to, empty } = state.selection;
  const para = $from.parent;
  if (para.type.name === "paragraph" && para.attrs.props?.list && empty && $from.parentOffset === 0) return indentParagraphs(1)(state, dispatch);
  if (!$from.sameParent($to)) return indentParagraphs(1)(state, dispatch);
  const tab = state.schema.nodes.tab;
  if (!$from.parent.inlineContent || !tab) return false;
  if (dispatch) dispatch(state.tr.replaceSelectionWith(tab.create(), true).scrollIntoView());
  return true;
};

/** Shift+Tab: a list item (or a selection of paragraphs) goes up a level / out by a step. */
export const shiftTabKey: Command = (state, dispatch) => indentParagraphs(-1)(state, dispatch);
