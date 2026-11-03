/**
 * Footnotes and endnotes: inserting one (a reference at the cursor, a note with Word's styles),
 * editing a note's text, and listing them in order with their numbers. The notes' text lives on the
 * document node (attribute `notes`), so every change to it is an ordinary, undoable step.
 */
import type { Block, Note, NoteKind, Notes, PrimeDocument, Style } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection } from "prosemirror-state";

type DocNotes = NonNullable<PrimeDocument["notes"]>;

const notesOf = (state: EditorState): DocNotes => (state.doc.attrs.notes as DocNotes | null) ?? {};

/** Word's styles for notes: the text 10pt with no space after; the reference mark superscript. */
const STYLES = (kind: NoteKind): Style[] => {
  const name = kind === "footnote" ? "Footnote" : "Endnote";
  return [
    { id: `${name}Text`, type: "paragraph", name: `${kind} text`, basedOn: "Normal", next: `${name}Text`, paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } }, run: { size: 10 } },
    { id: `${name}Reference`, type: "character", name: `${kind} reference`, run: { vertAlign: "superscript" } },
  ];
};

function ensureStyles(doc: PrimeDocument, kind: NoteKind) {
  for (const s of STYLES(kind)) if (!doc.styles.styles.some((x) => x.id === s.id)) doc.styles.styles.push(s);
}

/** A note as listed: its number (in reference order), where its reference is, and its text. */
export interface NoteEntry { kind: NoteKind; id: string; number: string; pos: number; note: Note }

const roman = (n: number) => {
  let out = "";
  for (const [v, r] of [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]] as const) while (n >= v) { out += r; n -= v; }
  return out;
};

/** Footnotes, then endnotes, in the order their references appear (numbered as they are shown). */
export function noteList(state: EditorState): NoteEntry[] {
  const notes = notesOf(state);
  const out: NoteEntry[] = [];
  const count = { footnote: 0, endnote: 0 };
  state.doc.descendants((n, pos) => {
    if (n.type.name !== "note_ref") return true;
    const kind = n.attrs.kind as NoteKind;
    const note = notes[kind]?.items.find((x) => x.id === n.attrs.id);
    if (note && !out.some((e) => e.kind === kind && e.id === note.id)) {
      const k = ++count[kind];
      out.push({ kind, id: note.id, number: kind === "footnote" ? String(k) : roman(k), pos, note });
    }
    return false;
  });
  return out.sort((a, b) => (a.kind === b.kind ? 0 : a.kind === "footnote" ? -1 : 1));
}

/**
 * Inserts a footnote (or endnote) reference at the cursor and an empty note for it; returns its id
 * through `onCreated` so the note can be opened for typing.
 */
export function insertNote(doc: () => PrimeDocument, kind: NoteKind, onCreated?: (id: string) => void): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const d = doc();
    ensureStyles(d, kind);
    const notes = notesOf(state);
    const current: Notes = notes[kind] ?? { items: [], special: [] };
    const id = String(Math.max(0, ...current.items.map((n) => Number(n.id)).filter(Number.isFinite)) + 1);
    const name = kind === "footnote" ? "Footnote" : "Endnote";
    const blocks: Block[] = [{
      type: "paragraph", props: { style: `${name}Text` },
      content: [{ type: "noteMark", kind, props: { style: `${name}Reference` } }, { type: "text", text: " " }],
    }];
    const next: DocNotes = { ...notes, [kind]: { ...current, items: [...current.items, { id, blocks }] } };
    const ref = state.schema.nodes.note_ref.create({ kind, id, props: { style: `${name}Reference` } });
    const { from, to } = state.selection;
    const tr = state.tr.setDocAttribute("notes", next).replaceWith(from, to, ref);
    dispatch(tr.scrollIntoView());
    onCreated?.(id);
    return true;
  };
}

/** Replaces a note's text. */
export function setNoteText(kind: NoteKind, id: string, blocks: Block[]): Command {
  return (state, dispatch) => {
    const notes = notesOf(state);
    const current = notes[kind];
    if (!current?.items.some((n) => n.id === id)) return false;
    if (dispatch) {
      const items = current.items.map((n) => (n.id === id ? { ...n, blocks } : n));
      dispatch(state.tr.setDocAttribute("notes", { ...notes, [kind]: { ...current, items } }));
    }
    return true;
  };
}

/** Deletes a note: its reference goes from the text (the saved file then leaves the note out). */
export function deleteNote(kind: NoteKind, id: string): Command {
  return (state, dispatch) => {
    let at = -1;
    state.doc.descendants((n, pos) => {
      if (at < 0 && n.type.name === "note_ref" && n.attrs.kind === kind && n.attrs.id === id) at = pos;
      return at < 0;
    });
    if (at < 0) return false;
    if (dispatch) dispatch(state.tr.delete(at, at + 1));
    return true;
  };
}

/** Selects a note's reference in the text (to go back from the note to where it is cited). */
export function gotoNoteRef(kind: NoteKind, id: string): Command {
  return (state, dispatch) => {
    const entry = noteList(state).find((e) => e.kind === kind && e.id === id);
    if (!entry) return false;
    if (dispatch) dispatch(state.tr.setSelection(NodeSelection.create(state.doc, entry.pos)).scrollIntoView());
    return true;
  };
}

/** The note whose reference is at `pos` (clicking a reference opens its note). */
export function noteAt(doc: Node, pos: number): { kind: NoteKind; id: string } | undefined {
  const n = doc.nodeAt(pos);
  return n?.type.name === "note_ref" ? { kind: n.attrs.kind, id: n.attrs.id } : undefined;
}
