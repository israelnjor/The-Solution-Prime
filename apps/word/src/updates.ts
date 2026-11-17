/**
 * Updates: Prime Word checks its release page (tauri.conf.json → plugins.updater) for a newer version,
 * and installs one only when it is signed with Prime Word's own key. The check sends nothing about the
 * user or their documents: it fetches a small file listing the latest version. Desktop only.
 */
import { isDesktop } from "./platform";

/** A newer version, ready to download and install. */
export interface AvailableUpdate {
  version: string;
  /** What is new (the release notes), if given. */
  notes?: string;
  date?: string;
  /** Downloads it (progress 0–1, or undefined when the size is unknown) and installs it; the app then restarts. */
  install(onProgress?: (share: number | undefined) => void): Promise<void>;
}

/** This version of Prime Word. */
export async function appVersion(): Promise<string> {
  if (!isDesktop) return "0.1.0 (in the browser)";
  const { getVersion } = await import("@tauri-apps/api/app");
  return getVersion();
}

/** Looks for a newer version: undefined when this one is the latest. Throws when the check fails (offline). */
export async function checkForUpdate(): Promise<AvailableUpdate | undefined> {
  if (!isDesktop) return undefined;
  const { check } = await import("@tauri-apps/plugin-updater");
  const update = await check({ timeout: 20000 });
  if (!update) return undefined;
  return {
    version: update.version,
    ...(update.body ? { notes: update.body } : {}),
    ...(update.date ? { date: update.date } : {}),
    async install(onProgress) {
      let total = 0;
      let done = 0;
      await update.downloadAndInstall((event) => {
        if (event.event === "Started") total = event.data.contentLength ?? 0;
        if (event.event === "Progress") { done += event.data.chunkLength; onProgress?.(total ? Math.min(1, done / total) : undefined); }
        if (event.event === "Finished") onProgress?.(1);
      });
      // On Windows the installer closes the app itself; elsewhere it restarts here.
      const { relaunch } = await import("@tauri-apps/plugin-process");
      await relaunch();
    },
  };
}

const AUTO = "prime-word:check-updates";
const LAST = "prime-word:last-update-check";

/** Whether Prime Word looks for updates by itself (on, unless turned off in About). */
export function autoCheckOn(): boolean {
  try { return localStorage.getItem(AUTO) !== "off"; } catch { return true; }
}
export function setAutoCheck(on: boolean): void {
  try { localStorage.setItem(AUTO, on ? "on" : "off"); } catch { /* storage unavailable: stays on */ }
}

/**
 * The automatic check: at most once a day, a little after start-up, quietly (no message when offline
 * or up to date). Calls `found` with a newer version.
 */
export function scheduleUpdateCheck(found: (u: AvailableUpdate) => void): () => void {
  if (!isDesktop || !autoCheckOn()) return () => undefined;
  let last = 0;
  try { last = Number(localStorage.getItem(LAST) ?? 0); } catch { /* none */ }
  if (Date.now() - last < 20 * 60 * 60 * 1000) return () => undefined;
  const timer = setTimeout(() => {
    void checkForUpdate().then((u) => {
      try { localStorage.setItem(LAST, String(Date.now())); } catch { /* none */ }
      if (u) found(u);
    }).catch(() => undefined);
  }, 8000);
  return () => clearTimeout(timer);
}
