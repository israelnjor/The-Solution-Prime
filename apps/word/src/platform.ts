/**
 * Platform layer: one interface for opening and saving documents.
 * Desktop (Tauri): native dialogs and direct file access. Browser: File System Access API
 * (Edge/Chrome) or upload/download as a fallback.
 */
import { invoke } from "@tauri-apps/api/core";
import { ask, open as openDialog, save as saveDialog } from "@tauri-apps/plugin-dialog";

export const isDesktop = "__TAURI_INTERNALS__" in window;

/**
 * Where a document lives, so Save can write back to it. `modified` is when the file last changed as far
 * as we know (opened or saved by us); a later change means another program saved it.
 */
export type DocumentLocation =
  | { kind: "path"; path: string; modified?: number }
  | { kind: "handle"; handle: BrowserFileHandle; modified?: number };

/** Save refused: the file changed on disk after `modified` (saving would overwrite someone's work). */
export class ChangedOnDiskError extends Error {
  constructor() { super("The file was changed by another program."); }
}

export interface OpenedFile {
  bytes: Uint8Array;
  name: string;
  location?: DocumentLocation;
}

type BrowserFileHandle = {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(d: Blob): Promise<void>; close(): Promise<void> }>;
};
type PickerWindow = Window & {
  showOpenFilePicker?: (o: object) => Promise<BrowserFileHandle[]>;
  showSaveFilePicker?: (o: object) => Promise<BrowserFileHandle>;
};
const browser = window as PickerWindow;
const MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PICKER_TYPES = [{ description: "Word document", accept: { [MIME]: [".docx"] } }];
const FILTERS = [{ name: "Word document", extensions: ["docx"] }];
// PDFs are not opened: they are converted, in the PDF converter (choosePdf).
const PDF_FILTERS = [{ name: "PDF", extensions: ["pdf"] }];

const baseName = (path: string) => path.split(/[\\/]/).pop() ?? path;

/** Desktop: reads a document the user chose or opened before (see read_document in lib.rs). */
export async function openPath(path: string): Promise<OpenedFile> {
  const modified = await invoke<number | null>("document_modified", { path });
  const bytes = new Uint8Array(await invoke<ArrayBuffer>("read_document", { path }));
  return { bytes, name: baseName(path), location: { kind: "path", path, ...(modified ? { modified } : {}) } };
}

export interface RecentDocument {
  path: string;
  name: string;
  exists: boolean;
}

/** Recently opened or saved documents, most recent first (desktop only). */
export async function recentDocuments(): Promise<RecentDocument[]> {
  return isDesktop ? invoke<RecentDocument[]>("recent_documents") : [];
}

export async function forgetRecent(path: string): Promise<void> {
  if (isDesktop) await invoke("forget_recent", { path });
}

/** Asks for a PDF to convert. Undefined when the user cancels. */
export async function choosePdf(): Promise<{ bytes: Uint8Array; name: string } | undefined> {
  if (isDesktop) {
    const path = await openDialog({ multiple: false, directory: false, filters: PDF_FILTERS });
    if (typeof path !== "string") return undefined;
    return { bytes: new Uint8Array(await invoke<ArrayBuffer>("read_document", { path })), name: baseName(path) };
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".pdf,application/pdf";
    input.onchange = async () => {
      const f = input.files?.[0];
      resolve(f ? { bytes: new Uint8Array(await f.arrayBuffer()), name: f.name } : undefined);
    };
    input.click();
  });
}

/** Asks the user for a document. Undefined when they cancel. */
export async function openDocument(): Promise<OpenedFile | undefined> {
  if (isDesktop) {
    const path = await openDialog({ multiple: false, directory: false, filters: FILTERS });
    if (typeof path !== "string") return undefined;
    return openPath(path);
  }
  if (browser.showOpenFilePicker) {
    let handles: BrowserFileHandle[];
    try { handles = await browser.showOpenFilePicker({ types: PICKER_TYPES }); } catch { return undefined; }
    const f = await handles[0].getFile();
    return { bytes: new Uint8Array(await f.arrayBuffer()), name: f.name, location: { kind: "handle", handle: handles[0], modified: f.lastModified } };
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".docx";
    input.onchange = async () => {
      const f = input.files?.[0];
      resolve(f ? { bytes: new Uint8Array(await f.arrayBuffer()), name: f.name } : undefined);
    };
    input.click();
  });
}

/**
 * Saves to `location`, or asks where to save (Save As) when there is none.
 * Returns the new location and name, or undefined when the user cancels. Throws ChangedOnDiskError when
 * the file changed since `location.modified`, unless `overwrite` is set.
 */
export async function saveDocument(bytes: Uint8Array, name: string, location?: DocumentLocation, overwrite = false):
  Promise<{ location?: DocumentLocation; name: string } | undefined> {
  const known = overwrite ? undefined : location?.modified;
  if (isDesktop) {
    let path = location?.kind === "path" ? location.path : undefined;
    if (!path) {
      path = (await saveDialog({ defaultPath: name, filters: FILTERS })) ?? undefined;
      if (!path) return undefined;
      if (!/\.docx$/i.test(path)) path += ".docx";
    }
    // Atomic save (temporary file, flushed, renamed over the original) — see save_document in lib.rs.
    let modified: number;
    try {
      modified = await invoke<number>("save_document", bytes, {
        headers: { "x-path": encodeURIComponent(path), ...(known && location?.kind === "path" ? { "x-known-modified": String(known) } : {}) },
      });
    } catch (e) {
      throw String(e) === "CHANGED_ON_DISK" ? new ChangedOnDiskError() : new Error(String(e));
    }
    return { location: { kind: "path", path, modified }, name: baseName(path) };
  }
  const blob = new Blob([bytes as Uint8Array<ArrayBuffer>], { type: MIME });
  let handle = location?.kind === "handle" ? location.handle : undefined;
  if (!handle && browser.showSaveFilePicker) {
    try { handle = await browser.showSaveFilePicker({ suggestedName: name, types: PICKER_TYPES }); } catch { return undefined; }
  }
  if (handle) {
    if (known && handle === (location?.kind === "handle" ? location.handle : undefined) && (await handle.getFile()).lastModified > known) {
      throw new ChangedOnDiskError();
    }
    const w = await handle.createWritable();
    await w.write(blob);
    await w.close();
    return { location: { kind: "handle", handle, modified: (await handle.getFile()).lastModified }, name: handle.name };
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
  return { name };
}

/** A yes/no question: a native dialog on the desktop, the browser's confirm() otherwise. */
export async function confirmAction(message: string, title = "Prime Word"): Promise<boolean> {
  if (isDesktop) return ask(message, { title, kind: "warning" });
  return window.confirm(message);
}
