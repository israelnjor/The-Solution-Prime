import { useEffect, useState } from "react";
import { isDesktop } from "./platform";
import { appVersion, autoCheckOn, type AvailableUpdate, checkForUpdate, setAutoCheck } from "./updates";

/** The open-source software Prime Word is built with, and its licences. */
const CREDITS: [string, string][] = [
  ["Tauri", "MIT or Apache-2.0"],
  ["React", "MIT"],
  ["ProseMirror", "MIT"],
  ["pdf-lib and fontkit", "MIT"],
  ["PDF.js (Mozilla)", "Apache-2.0"],
  ["fflate", "MIT"],
  ["nspell", "MIT"],
  ["English dictionaries (dictionary-en, dictionary-en-gb)", "MIT and BSD"],
  ["WordNet 3.0 (Princeton University), for the thesaurus", "WordNet licence"],
];

type Status =
  | { kind: "idle" } | { kind: "checking" } | { kind: "latest" } | { kind: "offline"; message: string }
  | { kind: "found"; update: AvailableUpdate } | { kind: "installing"; update: AvailableUpdate; share?: number } | { kind: "failed"; message: string };

/**
 * Home → About (also File → About): which version this is, checking for and installing updates,
 * whether Prime Word checks by itself, and the open-source software it is built with. `found`: an
 * update the automatic check found.
 */
export function AboutPage({ found }: { found?: AvailableUpdate }) {
  const [version, setVersion] = useState("");
  const [status, setStatus] = useState<Status>(found ? { kind: "found", update: found } : { kind: "idle" });
  const [auto, setAuto] = useState(autoCheckOn());
  useEffect(() => { void appVersion().then(setVersion); }, []);

  const check = async () => {
    setStatus({ kind: "checking" });
    try {
      const update = await checkForUpdate();
      setStatus(update ? { kind: "found", update } : { kind: "latest" });
    } catch (e) {
      const message = String((e as Error)?.message ?? e);
      // Nothing published yet (the release page has no update file): as good as up to date.
      if (/valid release JSON|404|not found/i.test(message)) setStatus({ kind: "latest" });
      else setStatus({ kind: "offline", message });
    }
  };
  const install = async (update: AvailableUpdate) => {
    setStatus({ kind: "installing", update });
    try {
      await update.install((share) => setStatus({ kind: "installing", update, share }));
    } catch (e) {
      setStatus({ kind: "failed", message: String((e as Error)?.message ?? e) });
    }
  };

  return (
    <main className="home-main">
      <div className="home-content about-page">
        <div className="about-head">
          <img src="/app-icon.png" alt="" width={72} height={72} onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }} />
          <div>
            <h1>Prime Word</h1>
            <p className="about-version">Version {version || "…"}</p>
            <p className="home-hint">A word processor by Prime Productivity. © 2026 Prime Productivity.</p>
          </div>
        </div>

        <section className="about-updates" aria-live="polite">
          <h4>Updates</h4>
          {!isDesktop && <p className="hint">Updates are part of the installed app.</p>}
          {isDesktop && status.kind === "idle" && <p>Look for a newer version of Prime Word.</p>}
          {status.kind === "checking" && <p>Checking for updates…</p>}
          {status.kind === "latest" && <p>✓ You have the latest version.</p>}
          {status.kind === "offline" && <p>Prime Word could not check for updates. Check your internet connection and try again.<small className="hint about-detail">{status.message}</small></p>}
          {status.kind === "failed" && <p>The update could not be installed: {status.message}. Your version is unchanged.</p>}
          {(status.kind === "found" || status.kind === "installing") && (
            <div className="about-found">
              <p><strong>Prime Word {status.update.version} is available.</strong></p>
              {status.update.notes && <pre className="about-notes">{status.update.notes}</pre>}
              {status.kind === "found" && <p className="hint">Prime Word closes to install it and opens again. Save your documents first (unsaved changes are kept for recovery too).</p>}
              {status.kind === "installing" && (
                <>
                  <progress max={1} value={status.share} aria-label="Downloading the update" />
                  <p className="hint">{status.share === 1 ? "Installing… Prime Word will open again." : "Downloading the update…"}</p>
                </>
              )}
            </div>
          )}
          {isDesktop && (
            <div className="about-actions">
              {status.kind === "found"
                ? <button className="primary" onClick={() => void install(status.update)}>Install and restart</button>
                : <button onClick={() => void check()} disabled={status.kind === "checking" || status.kind === "installing"}>Check for updates</button>}
              <label className="choice"><input type="checkbox" checked={auto} onChange={(e) => { setAuto(e.target.checked); setAutoCheck(e.target.checked); }} />
                Check for updates automatically (once a day)</label>
            </div>
          )}
          {isDesktop && <p className="hint">Checking only asks the Prime Word release page for the latest version number; nothing about you or your documents is sent. Updates are installed only when signed by Prime Productivity.</p>}
        </section>

        <section className="about-credits" aria-label="Open-source software">
          <h4>Open-source software</h4>
          <p className="hint">Prime Word is built with these, with thanks to their authors:</p>
          <ul>{CREDITS.map(([name, licence]) => <li key={name}>{name} <span className="hint">— {licence}</span></li>)}</ul>
        </section>
      </div>
    </main>
  );
}
