/**
 * Feedback: reports from the in-app form (FeedbackDialog), delivered to the Prime Productivity platform
 * (docs/feedback-api.md). Offline first: a report is written to the outbox on disk before anything is
 * sent, and stays there until the platform has it; sending is retried at start-up, when the connection
 * comes back, and every few minutes. Each report has its own id, so sending it again never makes a
 * duplicate. Nothing is sent but what the form shows (no document unless the tester attaches it).
 */
import { invoke } from "@tauri-apps/api/core";
import { isDesktop } from "./platform";
import { appVersion } from "./updates";

export type FeedbackCategory = "bug" | "difficulty" | "idea" | "general";
export const CATEGORIES: { id: FeedbackCategory; label: string; hint: string }[] = [
  { id: "bug", label: "Bug report", hint: "Something went wrong" },
  { id: "difficulty", label: "Hard to use", hint: "A feature was hard to find or use" },
  { id: "idea", label: "Feature request", hint: "Something you'd like Prime Word to do" },
  { id: "general", label: "General feedback", hint: "Anything else" },
];

/** At most this many bytes per attached file, and this many files. */
export const MAX_FILE = 10 * 1024 * 1024;
export const MAX_FILES = 5;

export interface Attachment { name: string; type: string; bytes: Uint8Array }

/** What Prime Word was doing when the form was opened (no document text). */
export interface FeedbackContext {
  screen: "home" | "document";
  homePage?: string;
  tab?: string;
  mode?: string;
  pages?: number;
  words?: number;
  openDocuments?: number;
  documentFormat?: string;
}

/** A report as stored in the outbox and sent (attachments as metadata; their bytes are files beside it). */
export interface Report {
  id: string;
  product: "prime-word";
  category: FeedbackCategory;
  description: string;
  email: string;
  followUp: boolean;
  createdAt: string;
  diagnostics: Record<string, unknown>;
  attachments: { file: string; name: string; type: string; size: number; sha256: string }[];
}

/** Where reports go: the platform's API (set at build time: VITE_FEEDBACK_API); empty: not set up yet. */
export const FEEDBACK_API = ((import.meta as { env?: Record<string, string> }).env?.VITE_FEEDBACK_API ?? "").replace(/\/+$/, "");

const enc = new TextEncoder();
const dec = new TextDecoder();
const newId = () => crypto.randomUUID?.() ?? `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;

/** This installation's random id (for the platform's rate limits; nothing personal). */
function installId(): string {
  try {
    let id = localStorage.getItem("prime-word:install-id");
    if (!id) { id = newId(); localStorage.setItem("prime-word:install-id", id); }
    return id;
  } catch {
    return "unknown";
  }
}

/** The tester's email and follow-up choice, remembered from the last report. */
export function rememberedSender(): { email: string; followUp: boolean } {
  try {
    return { email: localStorage.getItem("prime-word:feedback-email") ?? "", followUp: localStorage.getItem("prime-word:feedback-followup") === "1" };
  } catch {
    return { email: "", followUp: false };
  }
}
function rememberSender(email: string, followUp: boolean) {
  try { localStorage.setItem("prime-word:feedback-email", email); localStorage.setItem("prime-word:feedback-followup", followUp ? "1" : "0"); } catch { /* not kept */ }
}

/** Windows' version in words ("Windows 11 (10.0.22631)"), from the browser engine's own report. */
async function windowsVersion(): Promise<string> {
  const ua = navigator as Navigator & { userAgentData?: { platform: string; getHighEntropyValues(h: string[]): Promise<{ platformVersion?: string }> } };
  try {
    const data = ua.userAgentData;
    if (data?.platform === "Windows") {
      const { platformVersion = "" } = await data.getHighEntropyValues(["platformVersion"]);
      const major = Number(platformVersion.split(".")[0]);
      return `Windows ${major >= 13 ? "11" : "10"} (${platformVersion})`;
    }
    if (data?.platform) return data.platform;
  } catch { /* fall back */ }
  return /Windows NT [\d.]+/.exec(navigator.userAgent)?.[0] ?? navigator.platform;
}

/** The diagnostics sent with every report: versions, time, screen, language, what was open. */
export async function diagnostics(context: FeedbackContext): Promise<Record<string, unknown>> {
  return {
    appVersion: await appVersion(),
    os: await windowsVersion(),
    desktop: isDesktop,
    time: new Date().toISOString(),
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    language: navigator.language,
    screen: `${screen.width}×${screen.height} @${devicePixelRatio}`,
    context,
  };
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const hash = await crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>);
  return [...new Uint8Array(hash)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

// --- the outbox (desktop: files in the app's data folder; browser: IndexedDB) -----------------------

function idb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("prime-word-feedback", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("files");
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbDo<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await idb();
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction("files", mode).objectStore("files"));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

const outbox = {
  async save(id: string, file: string, bytes: Uint8Array): Promise<void> {
    if (isDesktop) await invoke("outbox_save", bytes, { headers: { "x-id": id, "x-file": file } });
    else await idbDo("readwrite", (s) => s.put(bytes, `${id}/${file}`));
  },
  async load(id: string, file: string): Promise<Uint8Array> {
    if (isDesktop) return new Uint8Array(await invoke<ArrayBuffer>("outbox_load", { id, file }));
    const v = await idbDo<Uint8Array | undefined>("readonly", (s) => s.get(`${id}/${file}`) as IDBRequest<Uint8Array | undefined>);
    if (!v) throw new Error("missing from the outbox");
    return v;
  },
  async list(): Promise<string[]> {
    if (isDesktop) return invoke<string[]>("outbox_list");
    const keys = await idbDo<IDBValidKey[]>("readonly", (s) => s.getAllKeys());
    return keys.map(String).filter((k) => k.endsWith("/report.json")).map((k) => k.slice(0, -"/report.json".length));
  },
  async remove(id: string): Promise<void> {
    if (isDesktop) { await invoke("outbox_delete", { id }); return; }
    const keys = await idbDo<IDBValidKey[]>("readonly", (s) => s.getAllKeys());
    for (const k of keys.map(String).filter((k) => k.startsWith(`${id}/`))) await idbDo("readwrite", (s) => s.delete(k));
  },
};

/** Reports waiting to be sent: how many, for the form's status line. */
let pending = 0;
const listeners = new Set<(n: number) => void>();
export function onPending(f: (n: number) => void): () => void {
  listeners.add(f);
  f(pending);
  return () => listeners.delete(f);
}
async function refreshPending() {
  pending = (await outbox.list().catch(() => [])).length;
  for (const f of listeners) f(pending);
}

/**
 * Keeps a new report: its attachments first, then the report itself (written last, so a report in the
 * outbox is always whole). From here on it is never lost.
 */
export async function keepReport(form: { category: FeedbackCategory; description: string; email: string; followUp: boolean; attachments: Attachment[]; context: FeedbackContext }): Promise<string> {
  const id = newId();
  const attachments: Report["attachments"] = [];
  for (const [i, a] of form.attachments.slice(0, MAX_FILES).entries()) {
    if (a.bytes.length > MAX_FILE) continue;
    const file = `a${i}`;
    await outbox.save(id, file, a.bytes);
    attachments.push({ file, name: a.name, type: a.type || "application/octet-stream", size: a.bytes.length, sha256: await sha256(a.bytes) });
  }
  const report: Report = {
    id, product: "prime-word", category: form.category, description: form.description.trim(), email: form.email.trim(),
    followUp: form.followUp, createdAt: new Date().toISOString(), diagnostics: await diagnostics(form.context), attachments,
  };
  await outbox.save(id, "report.json", enc.encode(JSON.stringify(report)));
  rememberSender(report.email, report.followUp);
  await refreshPending();
  return id;
}

/** Why a delivery stopped: no connection (try again later), the platform said no (kept, not retried), or not set up. */
export type Delivery = "sent" | "offline" | "refused" | "unconfigured";

class Refused extends Error {}
const sending = new Set<string>();

/**
 * Sends a report from the outbox (docs/feedback-api.md): announces it (the platform answers with upload
 * links for attachments it does not have yet), uploads those, then completes it; on success it leaves
 * the outbox. Each step can be repeated safely.
 */
export async function deliver(id: string): Promise<Delivery> {
  if (!FEEDBACK_API) return "unconfigured";
  if (sending.has(id)) return "offline";
  sending.add(id);
  try {
    const report = JSON.parse(dec.decode(await outbox.load(id, "report.json"))) as Report & { refused?: string };
    if (report.refused) return "refused";
    const headers = { "Content-Type": "application/json", "X-Prime-Install": installId(), "X-Prime-Product": report.product };
    const announce = await fetch(`${FEEDBACK_API}/v1/feedback`, { method: "POST", headers, body: JSON.stringify(report) });
    if (announce.status === 400 || announce.status === 413 || announce.status === 422) {
      // Refused for good (the report itself is wrong): kept, marked, not retried.
      await outbox.save(id, "report.json", enc.encode(JSON.stringify({ ...report, refused: `${announce.status} ${await announce.text().catch(() => "")}`.slice(0, 300) })));
      throw new Refused();
    }
    if (!announce.ok) return "offline";
    const { uploads = [] } = (await announce.json()) as { uploads?: { file: string; url: string; headers?: Record<string, string> }[] };
    for (const u of uploads) {
      const meta = report.attachments.find((a) => a.file === u.file);
      if (!meta) continue;
      const put = await fetch(u.url, { method: "PUT", headers: { "Content-Type": meta.type, ...u.headers }, body: (await outbox.load(id, u.file)) as Uint8Array<ArrayBuffer> });
      if (!put.ok) return "offline";
    }
    const done = await fetch(`${FEEDBACK_API}/v1/feedback/${encodeURIComponent(id)}/complete`, { method: "POST", headers });
    if (!done.ok) return "offline";
    await outbox.remove(id);
    await refreshPending();
    return "sent";
  } catch (e) {
    return e instanceof Refused ? "refused" : "offline";
  } finally {
    sending.delete(id);
  }
}

/** Sends every report waiting in the outbox, oldest first; stops at the first that cannot go now. */
export async function flushOutbox(): Promise<void> {
  if (!FEEDBACK_API) return;
  const ids = await outbox.list().catch(() => []);
  const reports: { id: string; at: string }[] = [];
  for (const id of ids) {
    try { reports.push({ id, at: (JSON.parse(dec.decode(await outbox.load(id, "report.json"))) as Report).createdAt }); } catch { /* unreadable: skipped */ }
  }
  for (const r of reports.sort((a, b) => a.at.localeCompare(b.at))) {
    if ((await deliver(r.id)) === "offline") break;
  }
  await refreshPending();
}

/**
 * Sends waiting reports a little after start-up, whenever the connection comes back, and every few
 * minutes while any are waiting. Returns a function that stops it.
 */
export function startOutbox(): () => void {
  const flush = () => { void flushOutbox(); };
  const first = setTimeout(flush, 15000);
  const every = setInterval(() => { if (pending) flush(); }, 5 * 60 * 1000);
  window.addEventListener("online", flush);
  void refreshPending();
  return () => { clearTimeout(first); clearInterval(every); window.removeEventListener("online", flush); };
}
