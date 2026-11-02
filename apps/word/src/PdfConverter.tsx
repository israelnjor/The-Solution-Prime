import type { LayoutResult, Measurer } from "@prime/layout";
import type { ConversionResult } from "@prime/pdf-import";
import { useCallback, useEffect, useRef, useState } from "react";
import { choosePdf } from "./platform";
import { convertPdfFile, drawLayoutPage, drawPdfPage, layoutOf, openPdfForView } from "./pdfImport";

type PdfView = Awaited<ReturnType<typeof openPdfForView>>;
type Stage =
  | { step: "choose" }
  | { step: "converting"; name: string; done: number; total: number }
  | { step: "review"; name: string; result: ConversionResult; original: PdfView; layout: LayoutResult }
  | { step: "failed"; name: string; message: string };

/**
 * PDF → Word, as its own workspace: choose a PDF; it is converted (page by page); review the original
 * and the converted document side by side; then open the result in the editor or save it as .docx.
 * The PDF itself is never changed.
 */
export function PdfConverter({ measurer, initial, onOpenInEditor, onSave, onClose }: {
  measurer: () => Measurer;
  /** A PDF to convert straight away (otherwise the person chooses one). */
  initial?: { bytes: Uint8Array; name: string };
  onOpenInEditor(result: ConversionResult, pdfName: string): void;
  onSave(result: ConversionResult, pdfName: string): void;
  onClose(): void;
}) {
  const [stage, setStage] = useState<Stage>({ step: "choose" });
  const [page, setPage] = useState(1);
  const [dragging, setDragging] = useState(false);
  const left = useRef<HTMLCanvasElement>(null);
  const right = useRef<HTMLCanvasElement>(null);

  const convert = useCallback(async (bytes: Uint8Array, name: string) => {
    setStage({ step: "converting", name, done: 0, total: 0 });
    try {
      const result = await convertPdfFile(bytes, (done, total) => setStage({ step: "converting", name, done, total }));
      const original = await openPdfForView(bytes);
      setPage(1);
      setStage({ step: "review", name, result, original, layout: layoutOf(result.doc, measurer()) });
    } catch (e) {
      setStage({ step: "failed", name, message: (e as Error).message ?? String(e) });
    }
  }, [measurer]);

  useEffect(() => { if (initial) void convert(initial.bytes, initial.name); }, [initial, convert]);

  const pick = async () => {
    const file = await choosePdf();
    if (file) await convert(file.bytes, file.name);
  };

  // Draw the current page of both documents (the original as pdf.js shows it; ours as the editor lays it out).
  useEffect(() => {
    if (stage.step !== "review") return;
    let cancelled = false;
    const draw = async () => {
      const scale = 1.25;
      if (left.current && page <= stage.original.numPages) await drawPdfPage(stage.original, page, left.current, scale);
      const ours = stage.layout.pages[page - 1];
      if (!cancelled && right.current && ours) drawLayoutPage(stage.result.doc, ours, right.current, scale);
      if (/[?&]selftest=/.test(location.search)) {
        console.info(`[selftest] align ${JSON.stringify({ review: stage.name, page, pages: stage.layout.pages.length, original: stage.original.numPages,
          left: left.current?.width, right: right.current?.width, warnings: stage.result.warnings })}`);
      }
    };
    void draw();
    // Pictures decode in the background: draw ours once more when they are ready.
    const again = setTimeout(() => {
      const ours = stage.layout.pages[page - 1];
      if (!cancelled && right.current && ours) drawLayoutPage(stage.result.doc, ours, right.current, 1.25);
    }, 600);
    return () => { cancelled = true; clearTimeout(again); };
  }, [stage, page]);

  const pages = stage.step === "review" ? Math.max(stage.original.numPages, stage.layout.pages.length) : 0;
  const stats = stage.step === "review" ? summary(stage.result) : undefined;

  return (
    <div className="pdfconv" role="dialog" aria-label="Convert PDF to Word"
      onKeyDown={(e) => {
        if (e.key === "Escape") onClose();
        if (stage.step === "review" && e.key === "PageDown") setPage((p) => Math.min(pages, p + 1));
        if (stage.step === "review" && e.key === "PageUp") setPage((p) => Math.max(1, p - 1));
      }}>
      <header className="pdfconv-bar">
        <strong>Convert PDF to Word</strong>
        {stage.step === "review" && (
          <>
            <span className="pdfconv-name">{stage.name}</span>
            <span className="spacer" />
            <button onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page <= 1} aria-label="Previous page">◀</button>
            <span>Page {page} of {pages}</span>
            <button onClick={() => setPage((p) => Math.min(pages, p + 1))} disabled={page >= pages} aria-label="Next page">▶</button>
            <span className="spacer" />
            <button onClick={() => void pick()}>Convert another…</button>
            <button onClick={() => onSave(stage.result, stage.name)}>Save as Word document…</button>
            <button className="primary" onClick={() => onOpenInEditor(stage.result, stage.name)}>Open in editor</button>
          </>
        )}
        {stage.step !== "review" && <span className="spacer" />}
        <button className="link" onClick={onClose} aria-label="Close the converter">×</button>
      </header>

      {stage.step === "choose" && (
        <div className={`pdfconv-drop ${dragging ? "over" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            const f = [...e.dataTransfer.files].find((x) => /\.pdf$/i.test(x.name));
            if (f) void f.arrayBuffer().then((b) => convert(new Uint8Array(b), f.name));
          }}>
          <h2>Turn a PDF into an editable Word document</h2>
          <p>The text, formatting, lists, tables, headers, footers and pictures are rebuilt. Everything happens on this computer; the PDF is not changed.</p>
          <button className="primary" onClick={() => void pick()} autoFocus>Choose a PDF…</button>
          <p className="hint">or drop a PDF here</p>
        </div>
      )}

      {stage.step === "converting" && (
        <div className="pdfconv-drop">
          <h2>Converting {stage.name}…</h2>
          <progress max={stage.total || 1} value={stage.done} />
          <p>{stage.total ? `Page ${stage.done} of ${stage.total}` : "Reading the PDF…"}</p>
        </div>
      )}

      {stage.step === "failed" && (
        <div className="pdfconv-drop">
          <h2>{stage.name} could not be converted</h2>
          <p>{stage.message}</p>
          <button className="primary" onClick={() => void pick()}>Choose another PDF…</button>
        </div>
      )}

      {stage.step === "review" && (
        <div className="pdfconv-review">
          <aside className="pdfconv-side">
            <h3>Result</h3>
            <ul className="pdfconv-stats">{stats!.map(([k, v]) => <li key={k}><span>{k}</span><strong>{v}</strong></li>)}</ul>
            <h3>Check</h3>
            {stage.result.warnings.length
              ? <ul className="pdfconv-warnings">{stage.result.warnings.map((w) => <li key={w}>{w}</li>)}</ul>
              : <p className="ok">Nothing flagged. Compare the pages before relying on the result.</p>}
          </aside>
          <figure>
            <figcaption>Original PDF</figcaption>
            <canvas ref={left} />
          </figure>
          <figure>
            <figcaption>Converted (as it will look in Prime Word)</figcaption>
            {page <= stage.layout.pages.length ? <canvas ref={right} /> : <p className="hint">The converted document has fewer pages.</p>}
          </figure>
        </div>
      )}
    </div>
  );
}

/** Counts for the review panel. */
function summary(r: ConversionResult): [string, string][] {
  const blocks = r.doc.sections.flatMap((s) => s.blocks);
  const paragraphs = blocks.filter((b) => b.type === "paragraph");
  const count = (pred: (b: (typeof blocks)[number]) => boolean) => blocks.filter(pred).length;
  return [
    ["Paragraphs", String(paragraphs.length)],
    ["Headings", String(paragraphs.filter((p) => p.type === "paragraph" && /^(Heading|Title)/.test(p.props.style ?? "")).length)],
    ["List items", String(paragraphs.filter((p) => p.type === "paragraph" && p.props.list).length)],
    ["Tables", String(count((b) => b.type === "table"))],
    ["Pictures", String(Object.keys(r.doc.images).length)],
    ["Header / footer", r.doc.sections[0].headers.default || r.doc.sections[0].footers.default ? "yes" : "no"],
  ];
}
