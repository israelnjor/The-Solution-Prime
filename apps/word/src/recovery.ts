/**
 * Autosave snapshots for crash recovery (ADR-010). While a document has unsaved changes, its
 * document model is snapshotted a few seconds after typing stops. A clean save, close or discard
 * deletes the snapshot; one that survives a crash is offered for restore at the next start.
 *
 * Desktop: atomic files in the app's local data folder (Rust). Browser: IndexedDB.
 */
import { invoke } from "@tauri-apps/api/core";
import type { PrimeDocument } from "@prime/pdm";
import { isDesktop } from "./platform";

export interface Snapshot {
  id: string;
  name: string;
  /** Where the document lives, if it has been saved before (desktop path). */
  path?: string;
  savedAt: number;
  warnings: string[];
  doc: PrimeDocument;
}

/** This window's editing session; its own snapshot is never offered for recovery. */
export const SESSION = (crypto.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`).replace(/[^A-Za-z0-9-]/g, "");

const encoder = new TextEncoder();

// --- browser storage (IndexedDB) -------------------------------------------------------------

function idb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("prime-word-recovery", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("snapshots", { keyPath: "id" });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbDo<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await idb();
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction("snapshots", mode).objectStore("snapshots"));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// --- API ---------------------------------------------------------------------------------

export async function saveSnapshot(s: Snapshot): Promise<void> {
  if (isDesktop) {
    await invoke("recovery_save", encoder.encode(JSON.stringify(s)), { headers: { "x-id": s.id } });
  } else {
    await idbDo("readwrite", (store) => store.put(s));
  }
}

export async function deleteSnapshot(id: string): Promise<void> {
  try {
    if (isDesktop) await invoke("recovery_delete", { id });
    else await idbDo("readwrite", (store) => store.delete(id));
  } catch {
    // Nothing to delete.
  }
}

/** Snapshots left behind by earlier sessions (newest first). */
export async function orphanedSnapshots(): Promise<Snapshot[]> {
  const out: Snapshot[] = [];
  try {
    if (isDesktop) {
      const list = await invoke<[string, number][]>("recovery_list");
      for (const [id] of list) {
        if (id === SESSION) continue;
        try {
          const bytes = new Uint8Array(await invoke<ArrayBuffer>("recovery_load", { id }));
          out.push(JSON.parse(new TextDecoder().decode(bytes)) as Snapshot);
        } catch {
          await deleteSnapshot(id); // unreadable leftovers are not useful
        }
      }
    } else {
      const all = await idbDo<Snapshot[]>("readonly", (store) => store.getAll() as IDBRequest<Snapshot[]>);
      out.push(...all.filter((s) => s.id !== SESSION));
    }
  } catch {
    // Recovery storage unavailable: nothing to offer.
  }
  return out.sort((a, b) => b.savedAt - a.savedAt);
}
