import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo } from "react";
import { addToDictionary, type Misspelling, recheck, type Speller } from "./spelling";

/** Right-click menu on a misspelled word: suggestions, Ignore All, Add to Dictionary. */
export function SpellMenu({ view, speller, issue, x, y, ignored, onClose }: {
  view: EditorView; speller: Speller; issue: Misspelling; x: number; y: number; ignored: Set<string>; onClose(): void;
}) {
  const suggestions = useMemo(() => speller.suggest(issue.word).slice(0, 6), [speller, issue.word]);
  useEffect(() => {
    const close = () => onClose();
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", close);
    return () => { window.removeEventListener("mousedown", close); window.removeEventListener("keydown", close); };
  }, [onClose]);

  const replace = (s: string) => {
    view.dispatch(view.state.tr.insertText(s, issue.from, issue.to));
    onClose();
    view.focus();
  };
  return (
    <div className="spell-menu" style={{ left: x, top: y }} onMouseDown={(e) => e.stopPropagation()}>
      {suggestions.length ? suggestions.map((s) => (
        <button key={s} className="suggestion" onClick={() => replace(s)}>{s}</button>
      )) : <span className="none">No suggestions</span>}
      <hr />
      <button onClick={() => { ignored.add(issue.word); recheck(view); onClose(); }}>Ignore All</button>
      <button onClick={() => { addToDictionary(speller, issue.word); recheck(view); onClose(); }}>Add to Dictionary</button>
    </div>
  );
}
