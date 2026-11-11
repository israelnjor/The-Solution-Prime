import { useEffect, useRef, useState } from "react";
import type { ShareFormat } from "./share";

/** The Share button: asks which copy to send (Word document or PDF), then opens the share panel. */
export function ShareMenu({ onShare, className, label = "Share" }: { onShare(format: ShareFormat): void; className?: string; label?: string }) {
  const [open, setOpen] = useState(false);
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => { if (!box.current?.contains(e.target as Node)) setOpen(false); };
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") setOpen(false); };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", esc);
    return () => { window.removeEventListener("mousedown", close); window.removeEventListener("keydown", esc); };
  }, [open]);
  const pick = (f: ShareFormat) => { setOpen(false); onShare(f); };
  // Placed against the window (the tab row scrolls sideways when narrow, and would clip it): below the
  // button, kept inside the window.
  const [at, setAt] = useState<{ top: number; left: number }>();
  const toggle = (e: React.MouseEvent<HTMLButtonElement>) => {
    const r = e.currentTarget.getBoundingClientRect();
    setAt({ top: r.bottom + 6, left: Math.max(8, Math.min(r.left, window.innerWidth - 258)) });
    setOpen((o) => !o);
  };
  return (
    <div className="share-menu" ref={box}>
      <button className={className} onClick={toggle} aria-expanded={open} aria-haspopup="menu"
        title="Send a copy of this document: by email, WhatsApp, Nearby Share…">{label}</button>
      {open && (
        <div className="share-popup" role="menu" style={at ? { position: "fixed", top: at.top, left: at.left, right: "auto" } : undefined}>
          <p>Send a copy as</p>
          <button role="menuitem" onClick={() => pick("docx")}>
            <strong>Word document</strong><span>They can edit it (.docx)</span>
          </button>
          <button role="menuitem" onClick={() => pick("pdf")}>
            <strong>PDF</strong><span>They can read and print it</span>
          </button>
        </div>
      )}
    </div>
  );
}
