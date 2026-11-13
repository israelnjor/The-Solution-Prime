import { SHAPES } from "@prime/editor";
import { isLineShape, shapePath } from "@prime/layout";
import { type ReactNode, useEffect, useRef, useState } from "react";

/** A shape's outline, drawn as the page draws it (shapePath), in a w×h box. */
function ShapeIcon({ geometry, w = 30, h = 22 }: { geometry: string; w?: number; h?: number }) {
  const line = isLineShape(geometry);
  return (
    <svg viewBox={`-2 -2 ${w + 4} ${h + 4}`} width={w + 4} height={h + 4} aria-hidden="true">
      <path d={shapePath(geometry, w, h)} fill={line ? "none" : "#dce7f5"} stroke="#2b579a" strokeWidth={1.4} strokeLinejoin="round" />
    </svg>
  );
}

/**
 * Insert → Shapes: a button that opens the shapes as pictures to pick from (rather than a list of names);
 * each names itself on hover.
 */
export function ShapePicker({ icon, onPick }: { icon: ReactNode; onPick(geometry: string): void }) {
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
  return (
    <span className="shape-picker" ref={box}>
      <button title="Insert a shape" aria-label="Shapes" aria-haspopup="menu" aria-expanded={!!at} onMouseDown={(e) => e.preventDefault()}
        // Placed against the window (narrow ribbons flatten their groups), below the button, inside the window.
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setAt(at ? undefined : { top: r.bottom + 4, left: Math.max(8, Math.min(r.left, window.innerWidth - 268)) });
        }}>
        {icon}<span className="label keep">Shapes</span><span className="caret">▾</span>
      </button>
      {at && (
        <div className="shape-popup" role="menu" aria-label="Shapes" style={{ top: at.top, left: at.left }}>
          {SHAPES.map((s) => (
            <button key={s.geometry} role="menuitem" title={s.label} aria-label={s.label} onMouseDown={(e) => e.preventDefault()}
              onClick={() => { setAt(undefined); onPick(s.geometry); }}>
              <ShapeIcon geometry={s.geometry} />
            </button>
          ))}
        </div>
      )}
    </span>
  );
}
