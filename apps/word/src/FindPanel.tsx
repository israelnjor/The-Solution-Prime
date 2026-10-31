import { findNext, findState, replaceAll, replaceCurrent, setFindQuery } from "@prime/editor";
import type { EditorView } from "prosemirror-view";
import { useEffect, useRef, useState } from "react";

/** Find (Ctrl+F) and Replace (Ctrl+H). Matches are highlighted in the document while the panel is open. */
export function FindPanel({ view, replace, onClose }: { view: EditorView | null; replace: boolean; onClose(): void }) {
  const [text, setText] = useState("");
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [message, setMessage] = useState("");
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => { input.current?.focus(); input.current?.select(); }, [replace]);
  useEffect(() => {
    if (view) setFindQuery({ text, matchCase, wholeWord })(view.state, view.dispatch);
    setMessage("");
  }, [view, text, matchCase, wholeWord]);
  // Clear highlights when the panel closes.
  useEffect(() => () => { if (view) setFindQuery({ text: "" })(view.state, view.dispatch); }, [view]);

  const fs = view ? findState(view.state) : undefined;
  const count = fs?.matches.length ?? 0;
  const run = (cmd: (s: EditorView["state"], d: EditorView["dispatch"]) => boolean) => { if (view) cmd(view.state, view.dispatch); };

  return (
    <div className="find-panel" onKeyDown={(e) => { if (e.key === "Escape") { onClose(); view?.focus(); } }}>
      <div className="row">
        <input ref={input} placeholder="Find" value={text} onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") run(findNext(e.shiftKey ? -1 : 1)); }} />
        <span className="count">{text ? `${count} match${count === 1 ? "" : "es"}` : ""}</span>
        <button onClick={() => run(findNext(-1))} title="Previous (Shift+Enter)" disabled={!count}>↑</button>
        <button onClick={() => run(findNext(1))} title="Next (Enter)" disabled={!count}>↓</button>
        <button onClick={() => { onClose(); view?.focus(); }} title="Close (Esc)">✕</button>
      </div>
      {replace && (
        <div className="row">
          <input placeholder="Replace with" value={replacement} onChange={(e) => setReplacement(e.target.value)} />
          <button onClick={() => run(replaceCurrent(replacement))} disabled={!count}>Replace</button>
          <button onClick={() => run(replaceAll(replacement, (n) => setMessage(`Replaced ${n}`)))} disabled={!count}>Replace All</button>
        </div>
      )}
      <div className="row options">
        <label><input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> Match case</label>
        <label><input type="checkbox" checked={wholeWord} onChange={(e) => setWholeWord(e.target.checked)} /> Whole words</label>
        {message && <span className="message">{message}</span>}
      </div>
    </div>
  );
}
