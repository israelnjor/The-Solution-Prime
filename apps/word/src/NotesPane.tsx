import {
  createRender, createSchema, deleteNote, gotoNoteRef, noteList, type NoteEntry, paragraphStyles, setNoteText, storyFromProseMirror,
  storyToProseMirror, toggleRun,
} from "@prime/editor";
import type { Measurer } from "@prime/layout";
import type { Block, NoteKind, PrimeDocument } from "@prime/pdm";
import { baseKeymap } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { EditorState, Selection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { useEffect, useRef } from "react";

/** One note's text, in a small editor of its own; every change goes into the document (undoable there too). */
function NoteEditor({ entry, doc, measurer, focus, onChange }: {
  entry: NoteEntry; doc: PrimeDocument; measurer: Measurer; focus: boolean; onChange(blocks: Block[]): void;
}) {
  const mount = useRef<HTMLDivElement>(null);
  const changed = useRef(onChange);
  changed.current = onChange;
  const editor = useRef<EditorView>();
  /** The text this editor last sent to the document (anything else came from outside: undo, say). */
  const sent = useRef(JSON.stringify(entry.note.blocks));
  useEffect(() => {
    const v = editor.current;
    const now = JSON.stringify(entry.note.blocks);
    if (!v || now === sent.current) return;
    sent.current = now;
    v.updateState(EditorState.create({ doc: storyToProseMirror(entry.note.blocks, v.state.schema), plugins: v.state.plugins }));
  }, [entry.note.blocks]);
  useEffect(() => {
    if (!mount.current) return;
    const schema = createSchema(createRender(doc, measurer));
    const v = new EditorView(mount.current, {
      state: EditorState.create({
        doc: storyToProseMirror(entry.note.blocks, schema),
        plugins: [
          history(),
          keymap({ "Mod-z": undo, "Mod-y": redo, "Mod-b": toggleRun("bold"), "Mod-i": toggleRun("italic"), "Mod-u": toggleRun("underline") }),
          keymap(baseKeymap),
          paragraphStyles(doc, measurer),
        ],
      }),
      attributes: { "aria-label": `${entry.kind === "footnote" ? "Footnote" : "Endnote"} ${entry.number}` },
      dispatchTransaction(tr) {
        v.updateState(v.state.apply(tr));
        if (tr.docChanged) {
          const blocks = storyFromProseMirror(v.state.doc);
          sent.current = JSON.stringify(blocks);
          changed.current(blocks);
        }
      },
    });
    editor.current = v;
    if (focus) {
      v.focus();
      v.dispatch(v.state.tr.setSelection(Selection.atEnd(v.state.doc)));
      mount.current.scrollIntoView({ block: "nearest" });
    }
    return () => v.destroy();
    // One editor per note for as long as the pane shows it.
  }, [entry.kind, entry.id]);
  return <div ref={mount} className="note-editor" />;
}

/**
 * References → Show Notes: every footnote and endnote, numbered as in the document, each editable in
 * place. "Go to" selects the note's reference in the text; Delete removes the reference and its note.
 */
export function NotesPane({ view, base, measurer, focus, onClose }: {
  view: EditorView; base: () => PrimeDocument; measurer: Measurer; focus?: { kind: NoteKind; id: string }; onClose(): void;
}) {
  const entries = noteList(view.state);
  const section = (kind: NoteKind) => entries.filter((e) => e.kind === kind);
  const list = (kind: NoteKind) => (
    <>
      <h4>{kind === "footnote" ? "Footnotes" : "Endnotes"}</h4>
      {section(kind).length === 0 && <p className="hint">{kind === "footnote" ? "No footnotes yet: References → Insert Footnote (Alt+Ctrl+F)." : "No endnotes yet: References → Insert Endnote (Alt+Ctrl+D)."}</p>}
      {section(kind).map((e) => (
        <div key={`${e.kind}-${e.id}`} className="note">
          <div className="note-head">
            <strong>{e.number}</strong>
            <button className="link" onClick={() => { gotoNoteRef(e.kind, e.id)(view.state, view.dispatch); view.focus(); }}>Go to reference</button>
            <button className="link" onClick={() => { deleteNote(e.kind, e.id)(view.state, view.dispatch); }} aria-label={`Delete ${e.kind} ${e.number}`}>Delete</button>
          </div>
          <NoteEditor entry={e} doc={base()} measurer={measurer} focus={focus?.kind === e.kind && focus.id === e.id}
            onChange={(blocks) => setNoteText(e.kind, e.id, blocks)(view.state, view.dispatch)} />
        </div>
      ))}
    </>
  );
  return (
    <aside className="comments-pane notes-pane" aria-label="Notes">
      <header><h3>Notes</h3><button className="link" onClick={onClose} aria-label="Close notes">×</button></header>
      {list("footnote")}
      {list("endnote")}
    </aside>
  );
}
