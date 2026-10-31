/**
 * PDF export and printing from the app: the layout engine lays out the document (the same pages as
 * Print Layout) and @prime/pdf draws them. On the desktop, fonts are embedded from the installed
 * font files; the browser has no access to those and falls back to a standard font.
 */
import { extractFont, type FontKey, layoutDocument, type Measurer } from "@prime/layout";
import { renderPdf } from "@prime/pdf";
import type { PrimeDocument } from "@prime/pdm";
import { invoke } from "@tauri-apps/api/core";
import { save as saveDialog } from "@tauri-apps/plugin-dialog";
import { isDesktop } from "./platform";

const keyOf = (f: FontKey) => `${f.family}|${f.bold}|${f.italic}`;

/** Lays out and renders the document; returns the PDF bytes and the page count. */
export async function makePdf(doc: PrimeDocument, measurer: Measurer): Promise<{ bytes: Uint8Array; pages: number }> {
  const layout = layoutDocument(doc, measurer);
  const fonts = new Map<string, Uint8Array>();
  if (isDesktop) {
    const wanted = new Map<string, FontKey>();
    for (const page of layout.pages) for (const line of page.lines) for (const s of line.segments) wanted.set(keyOf(s.font), s.font);
    for (const [key, font] of wanted) {
      const found = await invoke<[string, number] | null>("font_lookup", { family: font.family, bold: font.bold, italic: font.italic });
      if (!found) continue;
      const file = new Uint8Array(await invoke<ArrayBuffer>("font_bytes", { path: found[0] }));
      fonts.set(key, extractFont(file, found[1]));
    }
  }
  const bytes = await renderPdf(doc, layout, { fontData: (f) => fonts.get(keyOf(f)) });
  return { bytes, pages: layout.pages.length };
}

const pdfName = (docName: string) => docName.replace(/\.docx$/i, "") + ".pdf";

/** Asks where to save, then writes the PDF. Returns false if the user cancelled. */
export async function exportPdf(doc: PrimeDocument, measurer: Measurer, docName: string): Promise<boolean> {
  const { bytes } = await makePdf(doc, measurer);
  if (isDesktop) {
    let path = await saveDialog({ defaultPath: pdfName(docName), filters: [{ name: "PDF", extensions: ["pdf"] }] });
    if (!path) return false;
    if (!/\.pdf$/i.test(path)) path += ".pdf";
    await invoke("save_document", bytes, { headers: { "x-path": encodeURIComponent(path) } });
    return true;
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type: "application/pdf" }));
  a.download = pdfName(docName);
  a.click();
  URL.revokeObjectURL(a.href);
  return true;
}

/** Prints via the PDF: opened in the default PDF viewer (desktop) or a new tab (browser). */
export async function printDocument(doc: PrimeDocument, measurer: Measurer, docName: string): Promise<void> {
  const { bytes } = await makePdf(doc, measurer);
  if (isDesktop) {
    await invoke("open_pdf", bytes, { headers: { "x-name": encodeURIComponent(docName) } });
  } else {
    window.open(URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type: "application/pdf" })), "_blank");
  }
}
