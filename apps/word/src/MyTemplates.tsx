import { readDocx, writeDocx } from "@prime/docx";
import type { Template } from "@prime/editor";
import { CanvasMeasurer } from "@prime/layout/dist/browser/canvas.js";
import type { PrimeDocument } from "@prime/pdm";
import { useEffect, useRef, useState } from "react";
import { deleteEntry, type LibraryEntry, listEntries, newId, saveEntry } from "./library";
import { drawLayoutPage, layoutOf } from "./pdfImport";

/** A saved template: the document, as a .docx (base64). */
export type SavedTemplate = LibraryEntry<{ docx: string; kind?: "letterhead" | "document" }>;

const toBase64 = (bytes: Uint8Array) => {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
};
const fromBase64 = (b64: string) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));

export async function saveTemplate(doc: PrimeDocument, name: string, kind: "letterhead" | "document" = "document") {
  await saveEntry("templates", { id: newId(name), name, created: Date.now(), data: { docx: toBase64(writeDocx(doc)), kind } } satisfies SavedTemplate);
}

/** A saved template as a template to start a document from. */
export function asTemplate(t: SavedTemplate): Template {
  return { id: t.id, name: t.name, description: t.data.kind === "letterhead" ? "Letterhead" : "My template", create: () => readDocx(fromBase64(t.data.docx)).doc };
}

/** The first page of a document, small, as the layout draws it. */
export function PageThumb({ doc, width = 120 }: { doc: PrimeDocument; width?: number }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const page = layoutOf(doc, new CanvasMeasurer()).pages[0];
    if (!page || !canvas.current) return;
    const scale = width / page.width;
    drawLayoutPage(doc, page, canvas.current, scale * 2);
    // Pictures decode in the background: draw once more when they are ready.
    const again = setTimeout(() => canvas.current && drawLayoutPage(doc, page, canvas.current, scale * 2), 400);
    return () => clearTimeout(again);
  }, [doc, width]);
  return <canvas ref={canvas} className="page-thumb" style={{ width }} />;
}

/** Saved templates and letterheads, to start from; with delete. */
export function MyTemplates({ onPick, refresh = 0 }: { onPick(t: Template): void; refresh?: number }) {
  const [items, setItems] = useState<{ entry: SavedTemplate; doc: PrimeDocument }[]>();
  const load = () => void listEntries<SavedTemplate["data"]>("templates").then((list) => {
    setItems(list.flatMap((entry) => {
      try { return [{ entry, doc: readDocx(fromBase64(entry.data.docx)).doc }]; } catch { return []; }
    }));
  }).catch(() => setItems([]));
  useEffect(load, [refresh]);
  if (!items) return null;
  if (!items.length) return <p className="recent-empty">Templates and letterheads you save appear here.</p>;
  return (
    <div className="template-grid">
      {items.map(({ entry, doc }) => (
        <div key={entry.id} className="template-card mine">
          <button className="mine-open" onClick={() => onPick({ ...asTemplate(entry), create: () => structuredClone(doc) })} title={`New document from ${entry.name}`}>
            <PageThumb doc={doc} />
            <strong>{entry.name}</strong>
            <span>{entry.data.kind === "letterhead" ? "Letterhead" : "My template"}</span>
          </button>
          <button className="mine-delete" aria-label={`Delete ${entry.name}`} title="Delete this template"
            onClick={() => { if (window.confirm(`Delete the template "${entry.name}"?`)) void deleteEntry("templates", entry.id).then(load); }}>×</button>
        </div>
      ))}
    </div>
  );
}
