import { SHAPE_GROUPS, type ShapeChoice, SHAPES } from "@prime/editor";
import { isLineShape, shadeFill, shapeDrawing } from "@prime/layout";
import type { ShapeStyle } from "@prime/pdm";
import { type ReactNode, useEffect, useRef, useState } from "react";

/** A shape drawn as the page draws it (its outline, details and a line's arrow heads), in a w×h box. */
export function ShapeIcon({ choice, w: width = 22, h: height = 18 }: { choice: ShapeChoice; w?: number; h?: number }) {
  const line = isLineShape(choice.geometry);
  // Upright shapes in an upright box (an up-down arrow needs its height for its shaft).
  const upright = /^(upArrow|downArrow|upDownArrow)$/.test(choice.geometry);
  const [w, h] = upright ? [height * 0.75, width] : [width, height];
  // Lines across the box, diagonally (as Word shows them); other shapes fill it.
  const style: ShapeStyle = { geometry: choice.geometry, ...(line ? {} : { fill: "DCE7F5" }), line: { color: "2B579A", width: 1.2, ...choice.ends } };
  return (
    <svg viewBox={`${-3 - (width - w) / 2} -3 ${width + 6} ${Math.max(h, height) + 6}`} width={width + 6} height={height + 6} aria-hidden="true">
      {shapeDrawing(style, w, h).map((p, i) => {
        const fill = p.fill === "line" ? "#2B579A" : p.fill === "none" || line ? "none" : p.fill === "shape" ? "#DCE7F5" : `#${shadeFill("DCE7F5", p.fill)}`;
        return <path key={i} d={p.d} fill={fill} stroke={p.stroke ? "#2B579A" : "none"} strokeWidth={1.2} strokeLinejoin="round" />;
      })}
    </svg>
  );
}

/**
 * Insert → Shapes: a button that opens Word's shapes, in its groups, as pictures to pick from; each
 * names itself on hover. Picking one starts drawing it on the page.
 */
export function ShapePicker({ icon, label = "Shapes", onPick }: { icon: ReactNode; label?: string; onPick(choice: ShapeChoice): void }) {
  const [at, setAt] = useState<{ top: number; left: number; maxHeight: number }>();
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
      <button title="Insert a shape: pick one, then drag on the page to draw it" aria-label={label} aria-haspopup="menu" aria-expanded={!!at} onMouseDown={(e) => e.preventDefault()}
        // Placed against the window (narrow ribbons flatten their groups), below the button, inside the window.
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setAt(at ? undefined : { top: r.bottom + 4, left: Math.max(8, Math.min(r.left, window.innerWidth - 348)), maxHeight: Math.max(200, window.innerHeight - r.bottom - 16) });
        }}>
        {icon}<span className="label keep">{label}</span><span className="caret">▾</span>
      </button>
      {at && (
        <div className="shape-popup" role="menu" aria-label="Shapes" style={{ top: at.top, left: at.left, maxHeight: at.maxHeight }}>
          {SHAPE_GROUPS.map((group) => (
            <section key={group} aria-label={group}>
              <h4>{group}</h4>
              <div className="shape-grid">
                {SHAPES.filter((s) => s.group === group).map((s) => (
                  <button key={s.id ?? s.geometry} role="menuitem" title={s.label} aria-label={s.label} onMouseDown={(e) => e.preventDefault()}
                    onClick={() => { setAt(undefined); onPick(s); }}>
                    <ShapeIcon choice={s} />
                  </button>
                ))}
              </div>
            </section>
          ))}
        </div>
      )}
    </span>
  );
}
