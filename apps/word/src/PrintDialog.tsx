import { useEffect, useMemo, useRef, useState } from "react";
import { drawPdfPage, openPdfForView } from "./pdfImport";

export type ViewPdf = Awaited<ReturnType<typeof openPdfForView>>;

/** "1-3, 5, 8-" → page numbers (1-based, in order, no repeats), or an error message. */
export function parsePages(text: string, total: number): number[] | string {
  const out: number[] = [];
  const seen = new Set<number>();
  for (const part of text.split(/[,;]/).map((s) => s.trim()).filter(Boolean)) {
    const m = /^(\d*)\s*(?:[-–]\s*(\d*))?$/.exec(part);
    if (!m || (!m[1] && !m[2])) return `"${part}" is not a page or a range of pages (e.g. 1-3, 5).`;
    const range = part.includes("-") || part.includes("–");
    const a = m[1] ? Number(m[1]) : 1;
    const b = range ? (m[2] ? Number(m[2]) : total) : a;
    if (a < 1 || b < a) return `"${part}" is not a range of pages.`;
    if (a > total) return `The document has ${total} page${total === 1 ? "" : "s"}.`;
    for (let p = a; p <= Math.min(b, total); p++) if (!seen.has(p)) { seen.add(p); out.push(p); }
  }
  return out.length ? out : "Type the pages to print, e.g. 1-3, 5.";
}

/**
 * Prints pages of the PDF through the system's print dialog (printer, copies, two-sided…): each page
 * drawn as a 200 dpi picture on a sheet of its own size, in a hidden frame that is then printed.
 */
export async function printPdfPages(pdf: ViewPdf, pages: number[], onProgress?: (done: number, total: number) => void): Promise<void> {
  const DPI = 200;
  const urls: { url: string; w: number; h: number }[] = [];
  const canvas = document.createElement("canvas");
  try {
    for (const [i, n] of pages.entries()) {
      const page = await pdf.getPage(n);
      const vp = page.getViewport({ scale: 1 });
      await drawPdfPage(pdf, n, canvas, DPI / 72);
      const blob = await new Promise<Blob | null>((r) => canvas.toBlob(r, "image/png"));
      if (blob) urls.push({ url: URL.createObjectURL(blob), w: vp.width, h: vp.height });
      onProgress?.(i + 1, pages.length);
    }
    // One named page size per distinct sheet size (a landscape page prints on a landscape sheet).
    const sizes = [...new Set(urls.map((u) => `${u.w.toFixed(1)}x${u.h.toFixed(1)}`))];
    const css = sizes.map((s, i) => { const [w, h] = s.split("x"); return `@page s${i} { size: ${w}pt ${h}pt; margin: 0; } .s${i} { page: s${i}; width: ${w}pt; height: ${h}pt; }`; }).join("\n");
    const body = urls.map((u) => `<div class="s${sizes.indexOf(`${u.w.toFixed(1)}x${u.h.toFixed(1)}`)} sheet"><img src="${u.url}"></div>`).join("");
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden";
    document.body.appendChild(frame);
    const d = frame.contentDocument!;
    d.open();
    d.write(`<!doctype html><html><head><title>Print</title><style>html,body{margin:0;padding:0} .sheet{break-after:page;overflow:hidden} .sheet:last-child{break-after:auto} img{display:block;width:100%;height:100%} ${css}</style></head><body>${body}</body></html>`);
    d.close();
    await Promise.all([...d.images].map((img) => (img.complete ? Promise.resolve() : new Promise((r) => { img.onload = r; img.onerror = r; }))));
    // (Self-tests see the frame instead: the printer window cannot open in a test browser.)
    const hook = (globalThis as { primePrintHook?: (d: Document) => void }).primePrintHook;
    if (hook) hook(d);
    else {
      frame.contentWindow!.focus();
      frame.contentWindow!.print();
    }
    // The dialog has closed (print() returns when it does); the frame and pictures go after a moment.
    setTimeout(() => { frame.remove(); for (const u of urls) URL.revokeObjectURL(u.url); }, 60_000);
  } catch (e) {
    for (const u of urls) URL.revokeObjectURL(u.url);
    throw e;
  }
}

/** A page of the PDF, drawn to fit a box (sharp on high-density screens). */
export function PdfPage({ pdf, n, width, height, className }: { pdf: ViewPdf; n: number; width: number; height: number; className?: string }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const [size, setSize] = useState<{ w: number; h: number }>();
  useEffect(() => {
    let live = true;
    void pdf.getPage(n).then((page) => {
      if (!live || !canvas.current) return;
      const vp = page.getViewport({ scale: 1 });
      const fit = Math.min(width / vp.width, height / vp.height);
      setSize({ w: Math.floor(vp.width * fit), h: Math.floor(vp.height * fit) });
      return drawPdfPage(pdf, n, canvas.current, fit * (window.devicePixelRatio || 1));
    });
    return () => { live = false; };
  }, [pdf, n, width, height]);
  return <canvas ref={canvas} className={className} aria-label={`Page ${n}`} style={size ? { width: `${size.w}px`, height: `${size.h}px` } : { width: 0, height: 0 }} />;
}

/**
 * File → Print (Ctrl+P): which pages, a preview of them, then the system's print dialog (printer,
 * copies, two-sided). The pages are those of Print Layout and the PDF export.
 */
export function PrintDialog({ pdf, pages, current, onOpenPdf, onClose }: {
  pdf: ViewPdf; pages: number; current: number; onOpenPdf(): void; onClose(): void;
}) {
  const [which, setWhich] = useState<"all" | "current" | "custom">("all");
  const [custom, setCustom] = useState("");
  const [at, setAt] = useState(0);
  const [busy, setBusy] = useState<string>();
  const chosen = useMemo(() => (which === "all" ? Array.from({ length: pages }, (_, i) => i + 1) : which === "current" ? [current] : parsePages(custom, pages)), [which, custom, pages, current]);
  const list = typeof chosen === "string" ? [] : chosen;
  useEffect(() => setAt(0), [which, custom]);
  const shown = list[Math.min(at, list.length - 1)];

  const print = async () => {
    if (!list.length) return;
    try {
      setBusy("Preparing pages…");
      await printPdfPages(pdf, list, (done, total) => setBusy(`Preparing page ${done} of ${total}…`));
      onClose();
    } catch (e) {
      setBusy(undefined);
      alert(`Prime Word could not print:\n${(e as Error).message}`);
    }
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog print-dialog" role="dialog" aria-label="Print">
        <div className="print-options">
          <h3>Print</h3>
          <fieldset>
            <legend>Pages</legend>
            <label className="choice"><input type="radio" name="pages" checked={which === "all"} onChange={() => setWhich("all")} /> All pages ({pages})</label>
            <label className="choice"><input type="radio" name="pages" checked={which === "current"} onChange={() => setWhich("current")} /> Current page ({current})</label>
            <label className="choice"><input type="radio" name="pages" checked={which === "custom"} onChange={() => setWhich("custom")} /> Pages
              <input value={custom} placeholder="e.g. 1-3, 5" aria-label="Pages to print" onFocus={() => setWhich("custom")} onChange={(e) => { setCustom(e.target.value); setWhich("custom"); }}
                onKeyDown={(e) => { if (e.key === "Enter") void print(); }} />
            </label>
            {typeof chosen === "string" && custom && <p className="error">{chosen}</p>}
          </fieldset>
          <p className="hint">Printer, copies and two-sided printing are chosen next, in the printer window.</p>
          <div className="buttons">
            <button className="primary" disabled={!list.length || !!busy} onClick={() => void print()}>{busy ?? `Print ${list.length} page${list.length === 1 ? "" : "s"}…`}</button>
            <button onClick={onOpenPdf} title="Open the document as a PDF in your PDF viewer (to print from there)">Open as PDF</button>
            <button onClick={onClose}>Cancel</button>
          </div>
        </div>
        <div className="print-preview" aria-label="Preview">
          {shown ? <PdfPage pdf={pdf} n={shown} width={420} height={540} className="sheet" /> : <p className="hint">No pages chosen.</p>}
          <div className="pager">
            <button onClick={() => setAt((i) => Math.max(0, i - 1))} disabled={at <= 0} aria-label="Previous page">◀</button>
            <span>{list.length ? `Page ${shown} — ${Math.min(at, list.length - 1) + 1} of ${list.length}` : ""}</span>
            <button onClick={() => setAt((i) => Math.min(list.length - 1, i + 1))} disabled={at >= list.length - 1} aria-label="Next page">▶</button>
          </div>
        </div>
      </div>
    </div>
  );
}
