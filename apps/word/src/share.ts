/**
 * Share: sends a copy of the open document — as a Word document (.docx, for people to edit) or a PDF
 * (to read and print) — through the system's share panel: on Windows the panel with Mail, Outlook,
 * WhatsApp, Nearby Share, Phone Link and the like. What goes is a file, never a link. The document
 * itself is not saved or changed by sharing; the copy is as it is on screen now.
 *
 * In a browser, the browser's own share (where it can share files), otherwise a download.
 * Self-tests supply a stand-in at globalThis.primeShare.
 */
import { invoke } from "@tauri-apps/api/core";
import { isDesktop } from "./platform";

export type ShareFormat = "docx" | "pdf";

const TYPES: Record<ShareFormat, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
};

/** The shared copy's file name: the document's name with the format's extension. */
export const shareName = (docName: string, format: ShareFormat) =>
  `${docName.replace(/\.(docx?|dotx|rtf|odt|txt|pdf)$/i, "").trim() || "Document"}.${format}`;

type Stand = (name: string, bytes: Uint8Array) => Promise<void>;

/** Opens the share panel with `bytes` as the file `name`. */
export async function shareFile(name: string, bytes: Uint8Array, format: ShareFormat): Promise<void> {
  const stand = (globalThis as { primeShare?: Stand }).primeShare;
  if (stand) return stand(name, bytes);
  if (isDesktop) {
    await invoke("share_document", bytes, { headers: { "x-name": encodeURIComponent(name) } });
    return;
  }
  const file = new File([bytes as Uint8Array<ArrayBuffer>], name, { type: TYPES[format] });
  if (navigator.canShare?.({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: name });
    } catch (e) {
      if ((e as Error).name !== "AbortError") throw e;
    }
    return;
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}
