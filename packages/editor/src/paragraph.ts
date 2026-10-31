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
