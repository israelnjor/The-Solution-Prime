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
    // Every level restarts (a numbered sub-list starts again at a.).
    extra: Array.from({ length: 9 }, (_, l) => ({ name: "w:lvlOverride", xml: `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="1"/></w:lvlOverride>` })),
  });
  return id;
}

/**
 * Enter, as in Word: the new paragraph keeps the paragraph's formatting — the next item of a list, the
 * same indent and alignment (ProseMirror's own split, at the end of a paragraph, starts a plain one).
 * At the end of a paragraph whose style names a different next style (a heading: Normal), the new
 * paragraph takes that style instead, without the heading's own formatting. An empty list item ends
 * the list instead (exitEmptyListItem, before this).
 */
export function splitParagraph(doc: () => PrimeDocument): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    const para = $from.parent;
    if (para.type.name !== "paragraph" || $from.depth < 1 || !$from.sameParent($to)) return false;
    if (dispatch) {
      const tr = state.tr;
      if (!state.selection.empty) tr.deleteSelection();
      let marks = state.storedMarks ?? (state.selection.$to.parentOffset ? $from.marks() : null);
      let attrs = para.attrs;
      const styleId: string | undefined = para.attrs.props?.style;
      if ($to.parentOffset === para.content.size && styleId) {
        const sheet = (state.doc.attrs.styles ?? doc().styles) as PrimeDocument["styles"];
        const next = sheet.styles.find((s) => s.id === styleId)?.next;
        if (next && next !== styleId) {
          const isDefault = sheet.styles.find((s) => s.id === next)?.isDefault;
          attrs = { ...para.attrs, props: isDefault ? {} : { style: next } };
          marks = null;
        }
      }
      tr.split(tr.mapping.map(state.selection.from), 1, [{ type: para.type, attrs }]);
      if (marks) tr.ensureMarks(marks);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * Bullets / Numbering. In a list item (the cursor in one paragraph), a sub-list: the item moves a level
 * deeper, as a bullet or a number as clicked (a bulleted sub-list under a numbered item is a bullet
 * list one level in, and the numbers above carry on after it). Otherwise the selected paragraphs become
 * a list of `kind`, or plain paragraphs again if they all are one already (select the items to end or
 * convert a whole list; Backspace, Shift+Tab or Enter on an empty item also end one).
 */
export function toggleList(doc: () => PrimeDocument, kind: ListKind): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const one = paras.length === 1 ? paras[0] : undefined;
    const cur = one?.node.attrs.props.list;
    if (one && cur && cur.numId !== "0" && listKind(doc(), cur.numId)) {
      const level = Math.min(8, (cur.level ?? 0) + 1);
      const numId = listKind(doc(), cur.numId) === kind ? cur.numId : listFor(state, doc(), kind, one.pos);
      const props = { ...one.node.attrs.props, list: { numId, level } };
      delete props.indent;
      dispatch(state.tr.setNodeMarkup(one.pos, undefined, { ...one.node.attrs, props }).scrollIntoView());
      return true;
    }
    const all = paras.every((p) => listKind(doc(), p.node.attrs.props.list?.numId) === kind);
    const numId = all ? undefined : listFor(state, doc(), kind, paras[0].pos);
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const props = { ...node.attrs.props };
      if (numId) {
        props.list = { numId, level: props.list?.level ?? 0 };
        delete props.indent; // the list level sets the indent (Word drops the paragraph's own)
      } else delete props.list;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/**
 * The list a paragraph at `pos` belongs to at `level`, from the items above it: the nearest item at that
 * level (stopping at a shallower one or a paragraph outside lists). Moving a sub-list item back up a
 * level rejoins that list (a bullet under "2." becomes "3."), as in Word.
 */
export function listAbove(doc: Node, pos: number, level: number): string | undefined {
  const $pos = doc.resolve(pos);
  for (let i = $pos.index() - 1; i >= 0; i--) {
    const n = $pos.parent.child(i);
    const l = n.type.name === "paragraph" ? n.attrs.props?.list : undefined;
    if (!l) return undefined;
    if ((l.level ?? 0) === level) return l.numId;
    if ((l.level ?? 0) < level) return undefined;
  }
  return undefined;
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
      else props.list = { ...props.list, numId: (delta < 0 && listAbove(tr.doc, pos, lvl)) || props.list.numId, level: Math.min(8, lvl) };
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
    // In a sub-list: back up a level, into the list there (Word); at the top level: out of the list.
    const level = (props.list.level ?? 0) - 1;
    if (level >= 0) props.list = { ...props.list, numId: listAbove(state.doc, $from.before(), level) ?? props.list.numId, level };
    else delete props.list;
    const tr = state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, props });
    dispatch(tr.setSelection(TextSelection.create(tr.doc, $from.pos)));
  }
  return true;
};

/** The bullet library (Home → Bullets ▾): the symbol a level shows. */
export const BULLET_STYLES: { id: string; text: string; font?: string }[] = [
  { id: "disc", text: "•" }, { id: "circle", text: "○", font: "Segoe UI Symbol" }, { id: "square", text: "■", font: "Segoe UI Symbol" },
  { id: "diamond", text: "◆", font: "Segoe UI Symbol" }, { id: "diamonds", text: "❖", font: "Segoe UI Symbol" }, { id: "arrow", text: "➢", font: "Segoe UI Symbol" },
  { id: "check", text: "✓", font: "Segoe UI Symbol" }, { id: "star", text: "★", font: "Segoe UI Symbol" }, { id: "dash", text: "–" },
];

/** The numbering library (Home → Numbering ▾): the number format and the text around it. */
export const NUMBER_STYLES: { id: string; format: string; before?: string; after: string; label: string }[] = [
  { id: "decimal", format: "decimal", after: ".", label: "1. 2. 3." },
  { id: "decimal-paren", format: "decimal", after: ")", label: "1) 2) 3)" },
  { id: "decimal-both", format: "decimal", before: "(", after: ")", label: "(1) (2) (3)" },
  { id: "upper-roman", format: "upperRoman", after: ".", label: "I. II. III." },
  { id: "upper-letter", format: "upperLetter", after: ".", label: "A. B. C." },
  { id: "lower-letter", format: "lowerLetter", after: ")", label: "a) b) c)" },
  { id: "lower-roman", format: "lowerRoman", after: ".", label: "i. ii. iii." },
];

/** A list format from the library: the usual levels, with the chosen symbol or number at `at`. */
function styledDefinition(doc: PrimeDocument, kind: ListKind, styleId: string, at: number): string {
  const tag = `prime-${kind}-${styleId}-${at}`;
  let def = doc.lists.definitions.find((d) => d.extra?.some((x) => x.xml.includes(`w:val="${tag}"`)));
  if (!def) {
    const levels = Array.from({ length: 9 }, (_, i) => level(i, kind));
    const base = levels[at];
    if (kind === "bullet") {
      const b = BULLET_STYLES.find((s) => s.id === styleId) ?? BULLET_STYLES[0];
      levels[at] = { ...base, text: b.text, run: b.font ? { font: { ascii: b.font, hAnsi: b.font } } : undefined };
    } else {
      const n = NUMBER_STYLES.find((s) => s.id === styleId) ?? NUMBER_STYLES[0];
      levels[at] = { ...base, format: n.format, text: `${n.before ?? ""}%${at + 1}${n.after}` };
    }
    const id = String(Math.max(0, ...doc.lists.definitions.map((d) => Number(d.id)).filter(Number.isFinite)) + 1);
    def = { id, levels, extra: [{ name: "w:name", xml: `<w:name w:val="${tag}"/>` }, { name: "w:multiLevelType", xml: '<w:multiLevelType w:val="hybridMultilevel"/>' }] };
    doc.lists.definitions.push(def);
  }
  return def.id;
}

/**
 * A bullet or number from the library. With the cursor in a list item: a sub-list in that style, a level
 * deeper. Otherwise the selected paragraphs, at the level each is on (select a list's items to change
 * its style). A list right above with the same format continues; otherwise a new list starts (numbers
 * from 1).
 */
export function applyListStyle(doc: () => PrimeDocument, kind: ListKind, styleId: string): Command {
  return (state, dispatch) => {
    const paras = paragraphs(state);
    // The cursor in a list item: a sub-list in the chosen style, a level deeper (as the Bullets and
    // Numbering buttons do there). Selected items keep their levels and change style.
    const single = paras.length === 1 && paras[0].node.attrs.props.list && state.selection.$from.sameParent(state.selection.$to);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const d = doc();
    const selected = new Set(paras.map((p) => p.pos));
    const chosen = new Map<number, string>(); // level → numId
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const lvl = Math.min(8, (node.attrs.props.list?.level ?? 0) + (single ? 1 : 0));
      let numId = chosen.get(lvl);
      if (!numId) {
        const defId = styledDefinition(d, kind, styleId, lvl);
        const instances = d.lists.instances.filter((i) => i.definitionId === defId);
        const above = listAbove(state.doc, pos, lvl);
        numId = above && instances.some((i) => i.id === above) ? above : undefined;
        if (!numId) {
          // An instance no paragraph outside the selection uses, else a new one (restarted).
          const used = new Set<string>();
          state.doc.descendants((n, p) => {
            if (n.type.name !== "paragraph") return true;
            if (!selected.has(p) && n.attrs.props?.list?.numId) used.add(n.attrs.props.list.numId);
            return false;
          });
          numId = instances.find((i) => !used.has(i.id))?.id;
          if (!numId) {
            numId = String(Math.max(0, ...d.lists.instances.map((i) => Number(i.id)).filter(Number.isFinite)) + 1);
            d.lists.instances.push({
              id: numId, definitionId: defId,
              ...(instances.length ? { extra: Array.from({ length: 9 }, (_, l) => ({ name: "w:lvlOverride", xml: `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="1"/></w:lvlOverride>` })) } : {}),
            });
          }
        }
        chosen.set(lvl, numId);
      }
      const props = { ...node.attrs.props, list: { numId, level: lvl } };
      delete props.indent;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}
