/**
 * Bulleted and numbered lists. List formats live in the document's numbering definitions (outside
 * the editor's content), so creating a list adds a definition to the document model when needed.
 */
import type { ListDefinition, ListLevel, PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, TextSelection } from "prosemirror-state";

export type ListKind = "bullet" | "number";

const BULLETS = ["•", "o", "▪"];
const NUMBERS: [string, string][] = [["decimal", "%1."], ["lowerLetter", "%2."], ["lowerRoman", "%3."]];

function level(i: number, kind: ListKind): ListLevel {
  const indent = { left: 36 + 36 * i, hanging: 18 };
  if (kind === "bullet") {
    const symbol = BULLETS[i % 3];
    return {
      level: i, start: 1, format: "bullet", text: symbol, align: "left", paragraph: { indent },
      run: symbol === "o" ? { font: { ascii: "Courier New", hAnsi: "Courier New" } } : undefined,
    };
  }
  const [format] = NUMBERS[i % 3];
  return { level: i, start: 1, format, text: `%${i + 1}.`, align: "left", paragraph: { indent } };
}

/** The numbering instance (numId) for a list kind, creating its definition on first use. */
export function ensureList(doc: PrimeDocument, kind: ListKind): string {
  const tag = `prime-${kind}`;
  const existing = doc.lists.definitions.find((d) => d.extra?.some((x) => x.xml.includes(`w:val="${tag}"`)));
  if (existing) {
    const inst = doc.lists.instances.find((i) => i.definitionId === existing.id);
    if (inst) return inst.id;
  }
  const nextId = (ids: string[]) => String(Math.max(0, ...ids.map(Number).filter(Number.isFinite)) + 1);
  const def: ListDefinition = {
    id: nextId(doc.lists.definitions.map((d) => d.id)),
    levels: Array.from({ length: 9 }, (_, i) => level(i, kind)),
    // A name lets us find (and reuse) Prime Word's own list formats later.
    extra: [{ name: "w:name", xml: `<w:name w:val="${tag}"/>` }, { name: "w:multiLevelType", xml: '<w:multiLevelType w:val="hybridMultilevel"/>' }],
  };
  const numId = nextId(doc.lists.instances.map((i) => i.id));
  doc.lists.definitions.push(def);
  doc.lists.instances.push({ id: numId, definitionId: def.id });
  return numId;
}

/** The list kind of a numbering instance, from its first level's format. */
export function listKind(doc: PrimeDocument, numId: string | undefined): ListKind | undefined {
  if (!numId || numId === "0") return undefined;
  const inst = doc.lists.instances.find((i) => i.id === numId);
  const def = inst && doc.lists.definitions.find((d) => d.id === inst.definitionId);
  const format = def?.levels.find((l) => l.level === 0)?.format;
  return format === undefined ? undefined : format === "bullet" ? "bullet" : "number";
}

function paragraphs(state: EditorState): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ node, pos });
  });
  return out;
}

/** Turns the selected paragraphs into a list of `kind`, or back into plain paragraphs if they already are. */
export function toggleList(doc: () => PrimeDocument, kind: ListKind): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const all = paras.every((p) => listKind(doc(), p.node.attrs.props.list?.numId) === kind);
    const numId = all ? undefined : ensureList(doc(), kind);
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const props = { ...node.attrs.props };
      if (numId) props.list = { numId, level: props.list?.level ?? 0 };
      else delete props.list;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Moves list paragraphs one level deeper (+1) or shallower (-1); shallower than level 0 ends the list. */
export function changeLevel(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state).filter((p) => p.node.attrs.props.list);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const props = { ...node.attrs.props };
      const lvl = (props.list.level ?? 0) + delta;
      if (lvl < 0) delete props.list;
      else props.list = { ...props.list, level: Math.min(8, lvl) };
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    }
    dispatch(tr);
    return true;
  };
}

/** Enter on an empty list item ends the list, as in Word. */
export const exitEmptyListItem: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  const para = $from.parent;
  if (!empty || para.type.name !== "paragraph" || !para.attrs.props.list || para.content.size > 0) return false;
  if (dispatch) {
    const props = { ...para.attrs.props };
    delete props.list;
    const tr = state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, props });
    dispatch(tr.setSelection(TextSelection.create(tr.doc, $from.pos)));
  }
  return true;
};
