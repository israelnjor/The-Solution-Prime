import { pasteAs, type PasteMode, pasteOptionsKey } from "@prime/editor";
import type { EditorView } from "prosemirror-view";
import { useEffect, useState } from "react";

const MODES: [PasteMode, string, string][] = [
  ["keep", "Keep Source Formatting", "K"], ["merge", "Merge Formatting", "M"], ["text", "Keep Text Only", "T"],
];

/**
 * The Paste Options button at the end of what was just pasted (as in Word). Ctrl on its own opens it;
 * K, M or T then chooses. It goes away once anything else is typed.
 */
export function PasteOptions({ view }: { view: EditorView }) {
  const paste = pasteOptionsKey.getState(view.state);
  const [open, setOpen] = useState(false);
  const choose = (m: PasteMode) => { pasteAs(m)(view.state, view.dispatch); setOpen(false); view.focus(); };

  useEffect(() => {
    if (!paste) { setOpen(false); return; }
    let ctrlAlone = false;
    const down = (e: KeyboardEvent) => {
      ctrlAlone = e.key === "Control";
      if (!open) return;
      const m = MODES.find(([, , k]) => k === e.key.toUpperCase());
      if (m) { e.preventDefault(); e.stopPropagation(); choose(m[0]); }
      if (e.key === "Escape") { e.preventDefault(); setOpen(false); view.focus(); }
    };
    const up = (e: KeyboardEvent) => { if (e.key === "Control" && ctrlAlone) setOpen((o) => !o); };
    window.addEventListener("keydown", down, true);
    window.addEventListener("keyup", up, true);
    return () => { window.removeEventListener("keydown", down, true); window.removeEventListener("keyup", up, true); };
  });

  if (!paste) return null;
  let at: { left: number; bottom: number };
  try { at = view.coordsAtPos(Math.min(paste.to, view.state.doc.content.size)); } catch { return null; }
  return (
    <div className="paste-options" style={{ left: at.left + 4, top: at.bottom + 2 }}>
      <button className="paste-button" title="Paste Options (Ctrl)" aria-label="Paste Options" aria-expanded={open} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>📋 (Ctrl) ▾</button>
      {open && (
        <div className="paste-menu" role="menu">
          {MODES.map(([m, label, key]) => (
            <button key={m} role="menuitemradio" aria-checked={paste.mode === m} className={paste.mode === m ? "on" : ""}
              onMouseDown={(e) => e.preventDefault()} onClick={() => choose(m)}>
              <span>{label}</span><kbd>{key}</kbd>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
