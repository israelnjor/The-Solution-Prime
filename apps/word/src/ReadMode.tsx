import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { PdfPage, type ViewPdf } from "./PrintDialog";

/**
 * View → Read Mode: the document's pages to read, filling the window, one at a time or two side by
 * side (pages 1–2, 3–4…). Arrow keys, Page Up/Down, Space and a click on either half turn the pages;
 * Esc goes back to editing. Nothing can be changed here.
 */
export function ReadMode({ pdf, pages, start, name, onClose }: { pdf: ViewPdf; pages: number; start: number; name: string; onClose(): void }) {
  const [two, setTwo] = useState(() => { try { return localStorage.getItem("prime-word:read-two") !== "0"; } catch { return true; } });
  const per = two && pages > 1 ? 2 : 1;
  // The first page shown: in two-page view, always an odd page (1–2, 3–4…).
  const align = (p: number) => (per === 2 ? p - ((p - 1) % 2) : p);
  const [first, setFirst] = useState(() => align(Math.min(Math.max(1, start), pages)));
  useEffect(() => setFirst((p) => align(p)), [per]); // eslint-disable-line react-hooks/exhaustive-deps
  const area = useRef<HTMLDivElement>(null);
  const root = useRef<HTMLDivElement>(null);
  const [box, setBox] = useState({ w: 800, h: 600 });
  useLayoutEffect(() => {
    const el = area.current;
    if (!el) return;
    const measure = () => setBox({ w: el.clientWidth, h: el.clientHeight });
    measure();
    const watch = new ResizeObserver(measure);
    watch.observe(el);
    return () => watch.disconnect();
  }, []);
  useEffect(() => root.current?.focus(), []);
  const turn = (dir: 1 | -1) => setFirst((p) => Math.min(Math.max(1, p + dir * per), align(pages)));
  const setMode = (twoPages: boolean) => {
    setTwo(twoPages);
    try { localStorage.setItem("prime-word:read-two", twoPages ? "1" : "0"); } catch { /* storage unavailable */ }
  };
  const shown = [first, first + 1].slice(0, per).filter((p) => p <= pages);
  const gap = 24;
  const pageW = (box.w - gap * (per + 1)) / per;
  const pageH = box.h - gap * 2;

  return (
    <div ref={root} className="read-mode" role="dialog" aria-label="Read mode" tabIndex={-1}
      onKeyDown={(e) => {
        if (e.key === "Escape") { e.preventDefault(); onClose(); }
        else if (["ArrowRight", "ArrowDown", "PageDown", " "].includes(e.key)) { e.preventDefault(); turn(1); }
        else if (["ArrowLeft", "ArrowUp", "PageUp"].includes(e.key)) { e.preventDefault(); turn(-1); }
        else if (e.key === "Home") { e.preventDefault(); setFirst(1); }
        else if (e.key === "End") { e.preventDefault(); setFirst(align(pages)); }
      }}>
      <header>
        <strong title={name}>{name}</strong>
        <span className="read-pages">{shown.length === 2 ? `Pages ${shown[0]}–${shown[1]}` : `Page ${shown[0]}`} of {pages}</span>
        <div className="read-tools">
          <button className={per === 1 ? "on" : ""} aria-pressed={per === 1} onClick={() => setMode(false)} title="One page at a time">One Page</button>
          <button className={per === 2 ? "on" : ""} aria-pressed={per === 2} onClick={() => setMode(true)} disabled={pages < 2} title="Two pages side by side">Multiple Pages</button>
          <button onClick={onClose} title="Back to editing (Esc)">Edit Document</button>
        </div>
      </header>
      <div ref={area} className="read-area" onClick={(e) => {
        const r = e.currentTarget.getBoundingClientRect();
        turn(e.clientX > r.left + r.width / 2 ? 1 : -1);
      }}>
        {shown.map((n) => <PdfPage key={n} pdf={pdf} n={n} width={pageW} height={pageH} className="read-page" />)}
      </div>
      <footer>
        <button onClick={() => turn(-1)} disabled={first <= 1} aria-label="Previous page">◀</button>
        <input type="range" min={1} max={pages} step={per} value={first} aria-label="Page"
          onChange={(e) => setFirst(align(Number(e.target.value)))} />
        <button onClick={() => turn(1)} disabled={first + per > pages} aria-label="Next page">▶</button>
      </footer>
    </div>
  );
}
