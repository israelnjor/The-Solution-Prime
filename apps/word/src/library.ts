/**
 * The personal library: templates and letterheads, signatures, saved blocks (addresses, signature
 * blocks, stamps). Kept on this computer only — desktop: the app's data folder (Rust); browser: IndexedDB.
 */
import { invoke } from "@tauri-apps/api/core";
import { isDesktop } from "./platform";

export type LibraryKind = "templates" | "signatures" | "blocks";

export interface LibraryEntry<T = unknown> {
  id: string;
  name: string;
  created: number;
  data: T;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function idb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("prime-word-library", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("entries");
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbDo<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await idb();
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction("entries", mode).objectStore("entries"));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

/** A new id for an entry named `name` (letters, digits and dashes; unique). */
export function newId(name: string): string {
  const slug = name.normalize("NFKD").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "item";
  return `${slug}-${Date.now().toString(36)}`;
}

export async function saveEntry<T>(kind: LibraryKind, entry: LibraryEntry<T>): Promise<void> {
  if (isDesktop) {
    await invoke("library_save", encoder.encode(JSON.stringify(entry)), { headers: { "x-kind": kind, "x-id": entry.id } });
  } else {
    await idbDo("readwrite", (s) => s.put(entry, `${kind}/${entry.id}`));
  }
}

/** All entries of a kind, newest first. */
export async function listEntries<T>(kind: LibraryKind): Promise<LibraryEntry<T>[]> {
  if (isDesktop) {
    const ids = await invoke<string[]>("library_list", { kind });
    const out: LibraryEntry<T>[] = [];
    for (const id of ids) {
      try {
        out.push(JSON.parse(decoder.decode(new Uint8Array(await invoke<ArrayBuffer>("library_load", { kind, id })))));
      } catch { /* an unreadable entry is skipped */ }
    }
    return out;
  }
  const all = await idbDo("readonly", (s) => s.getAll(IDBKeyRange.bound(`${kind}/`, `${kind}/￿`))) as LibraryEntry<T>[];
  return all.sort((a, b) => b.created - a.created);
}

export async function deleteEntry(kind: LibraryKind, id: string): Promise<void> {
  if (isDesktop) await invoke("library_delete", { kind, id });
  else await idbDo("readwrite", (s) => s.delete(`${kind}/${id}`));
}
