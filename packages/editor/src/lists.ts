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

/**
 * The list a paragraph at `pos` joins: the one right above it if that is the same kind (it continues),
 * else a new list numbered from 1. (A second instance of the same format, restarted, as Word writes it:
 * without the restart Word would carry on the numbers of the earlier list.)
 */
function listFor(state: EditorState, doc: PrimeDocument, kind: ListKind, pos: number): string {
  const $pos = state.doc.resolve(pos);
  const above = $pos.index() > 0 ? $pos.parent.child($pos.index() - 1) : undefined;
  const prevId = above?.type.name === "paragraph" ? above.attrs.props?.list?.numId : undefined;
  if (prevId && listKind(doc, prevId) === kind) return prevId;
  const numId = ensureList(doc, kind);
  if (kind === "bullet") return numId;
  let used = false;
  state.doc.descendants((n) => {
    if (used) return false;
    if (n.type.name === "paragraph") { used = n.attrs.props?.list?.numId === numId; return false; }
    return true;
  });
  if (!used) return numId;
  const definitionId = doc.lists.instances.find((i) => i.id === numId)!.definitionId;
  const id = String(Math.max(0, ...doc.lists.instances.map((i) => Number(i.id)).filter(Number.isFinite)) + 1);
  doc.lists.instances.push({
    id, definitionId,
    extra: [{ name: "w:lvlOverride", xml: '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' }],
  });
  return id;
}

/**
 * Enter in a list item: the new paragraph is the next item (ProseMirror's own split, at the end of a
 * paragraph, starts a plain one). An empty item ends the list instead (exitEmptyListItem, before this).
 */
export const splitListItem: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const para = $from.parent;
  if (para.type.name !== "paragraph" || !para.attrs.props?.list || $from.depth < 1) return false;
  if (dispatch) {
    const tr = state.tr;
    if (!state.selection.empty) tr.deleteSelection();
    const marks = state.storedMarks ?? (state.selection.$to.parentOffset ? $from.marks() : null);
    tr.split(tr.mapping.map(state.selection.from), 1, [{ type: para.type, attrs: para.attrs }]);
    if (marks) tr.ensureMarks(marks);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/** Turns the selected paragraphs into a list of `kind`, or back into plain paragraphs if they already are. */
export function toggleList(doc: () => PrimeDocument, kind: ListKind): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const all = paras.every((p) => listKind(doc(), p.node.attrs.props.list?.numId) === kind);
    const numId = all ? undefined : listFor(state, doc(), kind, paras[0].pos);
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
/**
 * Backspace at the start of a paragraph, as in Word: a numbered or bulleted item first loses its number;
 * a paragraph that a page (or column) break put on a new page first loses that break — the text moves
 * back, and the two paragraphs stay apart (joining them would leave the break at the end of the first,
 * with what is typed after it going to the page after). Otherwise Backspace joins as usual.
 */
export const backspaceAtParagraphStart: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  const para = $from.parent;
  if (!empty || $from.parentOffset !== 0 || para.type.name !== "paragraph") return false;
  if (para.attrs.props?.list) {
    if (dispatch) {
      const props = { ...para.attrs.props };
      delete props.list;
      const tr = state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, props });
      dispatch(tr.setSelection(TextSelection.create(tr.doc, $from.pos)));
    }
    return true;
  }
  const depth = $from.depth - 1;
  const index = $from.index(depth);
  const prev = index > 0 ? $from.node(depth).child(index - 1) : undefined;
  const last = prev?.type.name === "paragraph" ? prev.lastChild : undefined;
  if (last?.type.name !== "hard_break" || (last.attrs.kind !== "page" && last.attrs.kind !== "column")) return false;
  if (dispatch) {
    const end = $from.before() - 1; // the end of the previous paragraph's content
    dispatch(state.tr.delete(end - last.nodeSize, end).scrollIntoView());
  }
  return true;
};

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
