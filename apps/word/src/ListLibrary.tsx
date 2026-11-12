import { BULLET_STYLES, NUMBER_STYLES } from "@prime/editor";
import { useEffect, useRef, useState } from "react";

/** The ▾ beside Bullets / Numbering: the library of bullet symbols or number styles to pick from. */
export function ListLibrary({ kind, onPick }: { kind: "bullet" | "number"; onPick(styleId: string): void }) {
  const [at, setAt] = useState<{ top: number; left: number }>();
  const box = useRef<HTMLSpanElement>(null);
  useEffect(() => {
    if (!at) return;
    const close = (e: MouseEvent) => { if (!box.current?.contains(e.target as Node)) setAt(undefined); };
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") setAt(undefined); };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", esc);
    return () => { window.removeEventListener("mousedown", close); window.removeEventListener("keydown", esc); };
  }, [at]);
  const name = kind === "bullet" ? "Bullet library" : "Numbering library";
  return (
    <span className="list-library" ref={box}>
      <button className="list-more" title={`${name}: choose a ${kind === "bullet" ? "bullet symbol" : "number style"}`} aria-label={name}
        aria-haspopup="menu" aria-expanded={!!at}
        // Placed against the window (narrow ribbons flatten their groups), below the button, inside the window.
        onMouseDown={(e) => e.preventDefault()} // keep the editor's selection
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setAt(at ? undefined : { top: r.bottom + 4, left: Math.max(8, Math.min(r.left - 8, window.innerWidth - 236)) });
        }}>▾</button>
      {at && (
        <div className={`list-popup ${kind}`} role="menu" aria-label={name} style={{ top: at.top, left: at.left }}>
          {kind === "bullet"
            ? BULLET_STYLES.map((s) => (
              <button key={s.id} role="menuitem" title={s.id} onMouseDown={(e) => e.preventDefault()} onClick={() => { setAt(undefined); onPick(s.id); }}
                style={s.font ? { fontFamily: `"${s.font}"` } : undefined}>{s.text}</button>
            ))
            : NUMBER_STYLES.map((s) => (
              <button key={s.id} role="menuitem" onMouseDown={(e) => e.preventDefault()} onClick={() => { setAt(undefined); onPick(s.id); }}>
                {s.label.split(" ").map((l) => <span key={l}>{l}</span>)}
              </button>
            ))}
        </div>
      )}
    </span>
  );
}
