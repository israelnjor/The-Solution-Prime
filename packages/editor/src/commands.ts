/**
 * Editing commands. Formatting edits the PDM properties carried by the editor (run marks and
 * paragraph attributes), so what the user changes is exactly what gets saved.
 */
import type { ParagraphProps, RunProps } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

type RunKey = "bold" | "italic" | "strike";

const runProps = (node: Node): RunProps => node.marks.find((m) => m.type.name === "run")?.attrs.props ?? {};

/** Text nodes in the selection with their clipped ranges. */
function selectedText(state: EditorState): { node: Node; from: number; to: number }[] {
  const out: { node: Node; from: number; to: number }[] = [];
  for (const range of state.selection.ranges) {
    state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
      if (node.isText) out.push({ node, from: Math.max(pos, range.$from.pos), to: Math.min(pos + node.nodeSize, range.$to.pos) });
    });
  }
  return out;
}

function hasProp(props: RunProps, key: RunKey | "underline"): boolean {
  if (key === "underline") return !!props.underline && props.underline.val !== "none";
  return !!props[key];
}

/** True when every selected character (or the cursor's formatting) has the property set directly. */
export function isActive(state: EditorState, key: RunKey | "underline"): boolean {
  if (state.selection.empty) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    return hasProp(marks.find((m) => m.type.name === "run")?.attrs.props ?? {}, key);
  }
  const texts = selectedText(state);
  return texts.length > 0 && texts.every((t) => hasProp(runProps(t.node), key));
}

function setProp(props: RunProps, key: RunKey | "underline", on: boolean): RunProps {
  const next: RunProps = { ...props };
  if (key === "underline") {
    if (on) next.underline = { val: "single" };
    else delete next.underline;
  } else if (on) next[key] = true;
  else delete next[key];
  return next;
}

/**
 * Rewrites the run formatting of the selection (or, with an empty selection, of the next typing)
 * through `update`, which receives each run's current properties.
 */
export function updateRun(update: (props: RunProps) => RunProps): Command {
  return (state, dispatch) => {
    const run = state.schema.marks.run;
    if (!dispatch) return true;
    const tr = state.tr;
    const clean = (p: RunProps) => Object.fromEntries(Object.entries(p).filter(([, v]) => v !== undefined)) as RunProps;
    if (state.selection.empty) {
      const marks = state.storedMarks ?? state.selection.$from.marks();
      const current = marks.find((m) => m.type === run)?.attrs.props ?? {};
      tr.setStoredMarks([...marks.filter((m) => m.type !== run), run.create({ props: clean(update(current)) })]);
    } else {
      for (const t of selectedText(state)) tr.addMark(t.from, t.to, run.create({ props: clean(update(runProps(t.node))) }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Toggles bold, italic, underline or strikethrough on the selection (or for the next typing). */
export function toggleRun(key: RunKey | "underline"): Command {
  return (state, dispatch, view) => {
    const on = !isActive(state, key);
    return updateRun((p) => setProp(p, key, on))(state, dispatch, view);
  };
}

/** Sets the font family (Latin and complex scripts), keeping any East Asian font as it was. */
export const setFont = (family: string): Command =>
  updateRun((p) => {
    // An explicit font replaces theme fonts for the same scripts; drop those keys rather than leave them undefined.
    const { asciiTheme: _a, hAnsiTheme: _h, csTheme: _c, ...rest } = p.font ?? {};
    return { ...p, font: { ...rest, ascii: family, hAnsi: family, cs: family } };
  });

/** Font size in points. */
export const setFontSize = (size: number): Command => updateRun((p) => ({ ...p, size }));

/** Text colour as hex (RRGGBB), or undefined for automatic. */
export const setColor = (hex: string | undefined): Command => updateRun((p) => ({ ...p, color: hex?.toUpperCase() }));

/** Highlight colour (Word's named highlight colours), or undefined for none. */
export const setHighlight = (name: string | undefined): Command => updateRun((p) => ({ ...p, highlight: name }));

/** Word's highlighter colours (w:highlight), by name. */
export const HIGHLIGHT_COLORS: [name: string, hex: string, label: string][] = [
  ["yellow", "FFFF00", "Yellow"], ["green", "00FF00", "Bright Green"], ["cyan", "00FFFF", "Turquoise"], ["magenta", "FF00FF", "Pink"],
  ["blue", "0000FF", "Blue"], ["red", "FF0000", "Red"], ["darkBlue", "000080", "Dark Blue"], ["darkCyan", "008080", "Teal"],
  ["darkGreen", "008000", "Green"], ["darkMagenta", "800080", "Violet"], ["darkRed", "800000", "Dark Red"], ["darkYellow", "808000", "Dark Yellow"],
  ["darkGray", "808080", "Gray 50%"], ["lightGray", "C0C0C0", "Gray 25%"], ["black", "000000", "Black"],
];

/**
 * Highlight in any colour (RRGGBB), or undefined for none. Word's highlighter colours are written as
 * highlights; any other colour as character shading (Word's way of giving text any background), so
 * the file opens in Word with the same colour.
 */
export const setHighlightColor = (hex: string | undefined): Command => updateRun((p) => {
  const out = { ...p };
  delete out.highlight;
  delete out.shading;
  if (!hex) return out;
  const named = HIGHLIGHT_COLORS.find(([, h]) => h === hex.toUpperCase());
  return named ? { ...out, highlight: named[0] } : { ...out, shading: { val: "clear", color: "auto", fill: hex.toUpperCase() } };
});

/** Direct run formatting at the start of the selection (for toolbar state). */
export function runAtSelection(state: EditorState): RunProps {
  const marks = state.storedMarks ?? state.selection.$from.marks();
  return marks.find((m) => m.type.name === "run")?.attrs.props ?? {};
}

/** Paragraph nodes touched by the selection, with their positions. */
function selectedParagraphs(state: EditorState): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ node, pos });
  });
  return out;
}

/** Sets (or clears, with undefined) a paragraph property on every selected paragraph. */
export function setParagraph<K extends keyof ParagraphProps>(key: K, value: ParagraphProps[K] | undefined): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of selectedParagraphs(state)) {
      const props: ParagraphProps = { ...node.attrs.props };
      if (value === undefined) delete props[key];
      else props[key] = value;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** The paragraph property shared by the selection's first paragraph (for toolbar state). */
export function paragraphProp<K extends keyof ParagraphProps>(state: EditorState, key: K): ParagraphProps[K] | undefined {
  return selectedParagraphs(state)[0]?.node.attrs.props[key];
}

export function wordCount(doc: Node): number {
  return (doc.textBetween(0, doc.content.size, " ", " ").match(/\S+/g) ?? []).length;
}
