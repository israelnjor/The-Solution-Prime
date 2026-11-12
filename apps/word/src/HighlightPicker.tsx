import { HIGHLIGHT_COLORS } from "@prime/editor";
import { useEffect, useRef, useState } from "react";

/** Soft shades: light backgrounds that keep dark text readable (saved as character shading). */
const SHADES = ["FFF2CC", "FFE699", "E2F0D9", "C6E0B4", "DDEBF7", "BDD7EE", "FCE4D6", "F8CBAD", "EDE2F6", "FFD7E6", "D9F2F2", "EDEDED"];

/**
 * Home → Highlight: the button highlights in the last colour used (as Word's highlighter does); the ▾
 * opens every colour: Word's highlighter colours, soft shades, and any colour at all ("More colours").
 * `current` is the selection's highlight: a highlighter name or "#RRGGBB".
 */
export function HighlightPicker({ current, onPick }: { current?: string; onPick(hex: string | undefined): void }) {
  const [last, setLast] = useState("FFFF00");
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
  const pick = (hex: string | undefined) => {
    if (hex) setLast(hex.toUpperCase());
    setAt(undefined);
    onPick(hex?.toUpperCase());
  };
  const currentHex = current?.startsWith("#") ? current.slice(1).toUpperCase() : HIGHLIGHT_COLORS.find(([n]) => n === current)?.[1];
  const keep = (e: React.MouseEvent) => e.preventDefault(); // keep the editor's selection
  const swatch = (hex: string, title: string) => (
    <button key={hex} role="menuitem" className={`swatch${currentHex === hex ? " on" : ""}`} title={title} aria-label={title}
      style={{ background: `#${hex}` }} onMouseDown={keep} onClick={() => pick(hex)} />
  );
  return (
    <span className="highlight-picker" ref={box}>
      <button className={`highlight-main${currentHex ? " on" : ""}`} title="Text highlight colour" aria-label="Highlight" onMouseDown={keep} onClick={() => pick(last)}>
        <span className="ab">ab</span><span className="bar" style={{ background: `#${last}` }} />
      </button>
      <button className="list-more" title="Highlight colours" aria-label="Highlight colours" aria-haspopup="menu" aria-expanded={!!at} onMouseDown={keep}
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setAt(at ? undefined : { top: r.bottom + 4, left: Math.max(8, Math.min(r.left - 60, window.innerWidth - 228)) });
        }}>▾</button>
      {at && (
        <div className="highlight-popup" role="menu" aria-label="Highlight colours" style={{ top: at.top, left: at.left }}>
          <p>Highlighter</p>
          <div className="swatches">{HIGHLIGHT_COLORS.map(([, hex, label]) => swatch(hex, label))}</div>
          <p>Soft shades</p>
          <div className="swatches">{SHADES.map((hex) => swatch(hex, `#${hex}`))}</div>
          <button className="wide" role="menuitem" onMouseDown={keep} onClick={() => pick(undefined)}>No colour</button>
          <label className="wide more" onMouseDown={(e) => e.stopPropagation()}>
            More colours…
            <input type="color" value={`#${(currentHex ?? last).toLowerCase()}`} onChange={(e) => pick(e.target.value.slice(1))} />
          </label>
        </div>
      )}
    </span>
  );
}
