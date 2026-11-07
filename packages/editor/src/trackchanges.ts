/**
 * Track changes: recording edits as revisions, as Word's Review → Track Changes does. With it on, an
 * edit is rewritten before it is applied: text taken out stays, marked deleted (w:del); text put in is
 * marked inserted (w:ins), after any text it replaces; paragraph breaks put in or taken out are marked
 * on the paragraph's mark. Deleting your own insertion removes it outright, as in Word. Formatting
 * changes keep the formatting from before (formatchanges.ts).
 */
import type { Mark, MarkType, Node } from "prosemirror-model";
import { type Command, type EditorState, Selection, TextSelection, type Transaction } from "prosemirror-state";
import { Mapping, ReplaceStep, type Step } from "prosemirror-transform";
import { recordFormatting } from "./formatchanges.js";

/** Whether edits are recorded as tracked changes (the document's Word setting). */
export function trackingOn(state: EditorState): boolean {
  return !!state.doc.attrs.trackRevisions;
}

/** Turns track changes on or off (saved with the document, as Word does). */
export function setTracking(on: boolean): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.setDocAttribute("trackRevisions", on).setMeta(NO_TRACK, true));
    return true;
  };
}

/** Meta for transactions that are never recorded (accepting or rejecting changes, layout fix-ups). */
export const NO_TRACK = "pw-notrack";

/** Word's revision date: to the minute, so adjacent typing joins one revision. */
function revisionDate(now = new Date()): string {
  return `${now.toISOString().slice(0, 16)}:00Z`;
}

let paragraphRevisionId = 800000 + Math.floor(Math.random() * 50000);

interface Revision { kind: "ins" | "del"; author?: string }

function revisionsOf(marks: readonly Mark[]): Revision[] {
  return marks.filter((m) => m.type.name === "revision").map((m) => {
    const a = (m.attrs.attrs ?? {}) as Record<string, string>;
    return { kind: m.attrs.kind, author: a["w:author"] ?? a.author };
  });
}

type MarkRunExtra = { name: string; xml: string }[];
function paragraphRevision(node: Node): Revision | undefined {
  const r = ((node.attrs.props?.markRun?.extra ?? []) as MarkRunExtra).find((x) => x.name === "w:ins" || x.name === "w:del");
  if (!r) return undefined;
  return { kind: r.name === "w:ins" ? "ins" : "del", author: /w:author="([^"]*)"/.exec(r.xml)?.[1] };
}

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** Sets (or, with kind undefined, clears) the tracked change on a paragraph's mark. */
function setParagraphRevision(tr: Transaction, pos: number, kind: "ins" | "del" | undefined, author: string, date: string) {
  const node = tr.doc.nodeAt(pos);
  if (!node) return;
  const props = node.attrs.props ?? {};
  const markRun = props.markRun ?? {};
  const extra = ((markRun.extra ?? []) as MarkRunExtra).filter((x) => x.name !== "w:ins" && x.name !== "w:del");
  if (kind) extra.unshift({ name: `w:${kind}`, xml: `<w:${kind} w:id="${paragraphRevisionId++}" w:author="${escAttr(author)}" w:date="${date}"/>` });
  const nextRun = { ...markRun, extra };
  if (!extra.length) delete nextRun.extra;
  tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: { ...props, markRun: Object.keys(nextRun).length ? nextRun : undefined } });
}

/** The text a range holds, block boundaries as newlines (to tell real edits from rebuilt nodes). */
const textOf = (doc: Node, from: number, to: number) => doc.textBetween(from, to, "\n", "￼");

/**
 * The deletion of `from`–`to` in `tr.doc`, recorded: inline content is marked deleted (your own
 * insertions go for good) and paragraph breaks inside are marked on their paragraphs' marks.
 */
function recordDeletion(tr: Transaction, from: number, to: number, revision: MarkType, author: string, date: string) {
  const doc = tr.doc;
  const del = revision.create({ kind: "del", attrs: { "w:author": author, "w:date": date } });
  // Operations from the end back, so earlier positions stay valid.
  const ops: { at: number; run: () => void }[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock) {
      const end = pos + node.nodeSize - 1; // the paragraph mark
      if (end >= from && pos + node.nodeSize < to) {
        const r = paragraphRevision(node);
        if (r?.kind === "del") return true;
        const next = doc.resolve(pos + node.nodeSize).nodeAfter;
        if (r?.kind === "ins" && r.author === author && next?.isTextblock) ops.push({ at: end, run: () => tr.join(end + 1) });
        else ops.push({ at: end, run: () => setParagraphRevision(tr, pos, "del", author, date) });
      }
      return true;
    }
    if (!node.isInline) return true;
    const a = Math.max(from, pos);
    const b = Math.min(to, pos + node.nodeSize);
    if (a >= b) return false;
    const revs = revisionsOf(node.marks);
    if (revs.some((r) => r.kind === "del")) return false;
    if (revs.some((r) => r.kind === "ins" && r.author === author)) ops.push({ at: a, run: () => tr.delete(a, b) });
    else ops.push({ at: a, run: () => tr.addMark(a, b, del) });
    return false;
  });
  ops.sort((x, y) => y.at - x.at);
  for (const op of ops) op.run();
}

/** Marks what was put in between `from` and `to` as inserted, paragraph breaks included. */
function recordInsertion(tr: Transaction, from: number, to: number, revision: MarkType, author: string, date: string) {
  if (to <= from) return;
  tr.removeMark(from, to, revision);
  tr.addMark(from, to, revision.create({ kind: "ins", attrs: { "w:author": author, "w:date": date } }));
  const paragraphs: number[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock) {
      const end = pos + node.nodeSize - 1;
      if (end >= from && end < to) paragraphs.push(pos);
    }
    return !node.isInline;
  });
  for (const pos of paragraphs) setParagraphRevision(tr, pos, "ins", author, date);
}

/**
 * The transaction with its edits recorded as tracked changes, or `tr` itself when there is nothing to
 * record (track changes off, no change to the text, or a transaction marked NO_TRACK). Steps that only
 * rebuild nodes (same text before and after: a table whose columns were resized, a restyled paragraph)
 * apply as they are.
 */
export function trackTransaction(state: EditorState, tr: Transaction, author: string, date = revisionDate()): Transaction {
  if (!tr.docChanged || !trackingOn(state) || tr.getMeta(NO_TRACK) || tr.getMeta("addToHistory") === false) return tr;
  const revision = state.schema.marks.revision;
  const out = state.tr;
  // Positions in the transaction's step-by-step documents → positions in `out`.
  let map = new Mapping();
  // Ranges other steps changed (formatting), as out-step index and positions after that step.
  const changed: { at: number; from: number; to: number }[] = [];
  const onlyDelete = tr.steps.length === 1 && tr.steps[0] instanceof ReplaceStep && (tr.steps[0] as ReplaceStep).slice.size === 0;
  tr.steps.forEach((step, i) => {
    const before = tr.docs[i];
    const k = out.steps.length;
    if (step instanceof ReplaceStep && !rebuilds(step, before)) {
      const from = map.map(step.from, 1);
      const to = Math.max(from, map.map(step.to, -1));
      if (to > from) recordDeletion(out, from, to, revision, author, date);
      if (step.slice.size) {
        // The insertion goes after the text it replaces.
        const at = out.mapping.slice(k).map(to, 1);
        const n = out.steps.length;
        out.replace(at, at, step.slice);
        const m = out.mapping.slice(n);
        recordInsertion(out, m.map(at, -1), m.map(at, 1), revision, author, date);
      }
    } else {
      const mapped = step.map(map);
      if (mapped && !out.maybeStep(mapped).failed) {
        const r = mapped as unknown as { from?: number; to?: number; pos?: number };
        const m = mapped.getMap();
        if (r.from !== undefined && r.to !== undefined) changed.push({ at: out.steps.length, from: m.map(r.from, -1), to: m.map(r.to, 1) });
        else if (r.pos !== undefined) changed.push({ at: out.steps.length, from: r.pos, to: r.pos + 1 });
      }
    }
    const next = new Mapping();
    next.appendMap(step.getMap().invert());
    next.appendMapping(map);
    next.appendMapping(out.mapping.slice(k));
    map = next;
  });
  // Formatting changed: recorded on the runs and paragraphs it changed.
  if (changed.length) {
    const back = out.mapping.invert();
    for (const c of changed) {
      const m = out.mapping.slice(c.at);
      recordFormatting(out, state.doc, back, m.map(c.from, -1), m.map(c.to, 1), author, date);
    }
  }
  // The cursor: Backspace leaves it before the text it marked deleted; Delete and typing, after.
  const { selection } = tr;
  const backspace = onlyDelete && state.selection.empty && state.selection.head === (tr.steps[0] as ReplaceStep).to;
  const head = map.map(selection.head, backspace ? -1 : 1);
  const anchor = selection.empty ? head : map.map(selection.anchor, backspace ? -1 : 1);
  try {
    out.setSelection(TextSelection.create(out.doc, anchor, head));
  } catch {
    out.setSelection(Selection.near(out.doc.resolve(Math.min(head, out.doc.content.size))));
  }
  if (tr.storedMarksSet && tr.storedMarks) out.setStoredMarks(tr.storedMarks.filter((m) => m.type !== revision));
  for (const [key, value] of Object.entries((tr as unknown as { meta: Record<string, unknown> }).meta ?? {})) out.setMeta(key, value);
  if (tr.scrolledIntoView) out.scrollIntoView();
  return out;
}

/** A step that replaces nodes with others holding the same text (a rebuild, not an edit). */
function rebuilds(step: Step, doc: Node): boolean {
  if (!(step instanceof ReplaceStep) || !step.slice.size || step.to <= step.from) return false;
  return textOf(doc, step.from, step.to) === step.slice.content.textBetween(0, step.slice.content.size, "\n", "￼")
    && sameShape(doc.slice(step.from, step.to).content, step.slice.content);
}

/** Same blocks in the same order (types only): a rebuilt table, not text moved between paragraphs. */
function sameShape(a: Node["content"], b: Node["content"]): boolean {
  if (a.childCount !== b.childCount) return false;
  for (let i = 0; i < a.childCount; i++) {
    const x = a.child(i);
    const y = b.child(i);
    if (x.type !== y.type) return false;
    if (!x.isTextblock && !x.isLeaf && !sameShape(x.content, y.content)) return false;
  }
  return true;
}
