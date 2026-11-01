import { addComment, anchoredComments, deleteComment, reviewKey } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useRef, useState } from "react";

const AUTHOR = "prime-word:author";

/** The name comments are signed with (asked for once, kept on this computer). */
export function authorName(): string {
  try { return localStorage.getItem(AUTHOR) ?? ""; } catch { return ""; }
}

const when = (date?: string) => {
  const d = date ? new Date(date) : undefined;
  return d && !Number.isNaN(d.getTime()) ? d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" }) : "";
};

/** Review → Comments: the document's comments in reading order; click one to go to its text. */
export function CommentsPane({ view, base, writing, onWritten, onClose }: {
  view: EditorView; base: () => PrimeDocument; writing: boolean; onWritten(): void; onClose(): void;
}) {
  const doc = view.state.doc;
  const comments = useMemo(() => anchoredComments(doc, base()), [doc, base]);
  const [author, setAuthor] = useState(authorName);
  const [draft, setDraft] = useState("");
  const [active, setActive] = useState<string>();
  const box = useRef<HTMLTextAreaElement>(null);
  useEffect(() => { if (writing) box.current?.focus(); }, [writing]);

  const saveAuthor = (name: string) => {
    setAuthor(name);
    try { localStorage.setItem(AUTHOR, name); } catch { /* storage unavailable */ }
  };
  const go = (id: string, from: number, to: number) => {
    setActive(id);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)).setMeta(reviewKey, { active: id }).scrollIntoView());
    view.focus();
  };
  const post = () => {
    if (addComment(base, author.trim() || "Prime Word user", draft)(view.state, view.dispatch)) {
      setDraft("");
      onWritten();
    }
  };

  return (
    <aside className="comments-pane" aria-label="Comments">
      <header>
        <strong>Comments</strong>
        <button className="link" onClick={onClose} aria-label="Close comments">×</button>
      </header>
      <label className="author">Your name
        <input value={author} onChange={(e) => saveAuthor(e.target.value)} placeholder="Shown on your comments" />
      </label>
      {writing && (
        <div className="comment new">
          <textarea ref={box} value={draft} onChange={(e) => setDraft(e.target.value)} rows={3} placeholder="Write a comment on the selected text"
            onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) post(); if (e.key === "Escape") onWritten(); }} />
          <div className="actions">
            <button onClick={onWritten}>Cancel</button>
            <button className="primary" onClick={post} disabled={!draft.trim()}>Post</button>
          </div>
        </div>
      )}
      {!comments.length && !writing && <p className="empty">No comments. Select some text and choose New Comment.</p>}
      {comments.map((c) => (
        <div key={c.id} className={`comment ${active === c.id ? "active" : ""}`} onClick={() => go(c.id, c.from, c.to)}>
          <div className="meta"><strong>{c.author || "Unknown"}</strong> <span>{when(c.date)}</span></div>
          <div className="quote">{doc.textBetween(c.from, c.to, " ").slice(0, 120)}</div>
          <div className="text">{c.text}</div>
          <button className="link delete" onClick={(e) => { e.stopPropagation(); deleteComment(c.id)(view.state, view.dispatch); }}>Delete</button>
        </div>
      ))}
    </aside>
  );
}
