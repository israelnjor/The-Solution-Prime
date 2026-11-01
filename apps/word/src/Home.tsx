import { type Template, TEMPLATES } from "@prime/editor";
import { useEffect, useRef } from "react";
import { isDesktop } from "./platform";
import { RecentList } from "./RecentList";
import type { Snapshot } from "./recovery";

const greeting = () => {
  const h = new Date().getHours();
  return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
};

/**
 * The home screen: shown when Prime Word starts without a document, and from the Prime Word button.
 * New (templates), Open, recent documents, and unsaved work to recover.
 */
export function Home({ recovery, onPick, onOpen, onOpenRecent, onRestore, onDiscard, onClose }: {
  recovery: Snapshot[];
  onPick(t: Template): void;
  onOpen(): void;
  onOpenRecent(path: string): void;
  onRestore(s: Snapshot): void;
  onDiscard(s: Snapshot): void;
  /** Back to the open document (absent when there is none to go back to). */
  onClose?(): void;
}) {
  const first = useRef<HTMLButtonElement>(null);
  useEffect(() => { first.current?.focus(); }, []);
  return (
    <div className="home" role="dialog" aria-label="Prime Word home" onKeyDown={(e) => {
      if (e.key === "Escape" && onClose) onClose();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") { e.preventDefault(); onOpen(); }
    }}>
      <nav className="home-rail" aria-label="Home">
        <div className="home-logo">Prime Word</div>
        {onClose && <button onClick={onClose}>← Back to document</button>}
        <button className="on">Home</button>
        <button onClick={() => onPick(TEMPLATES[0])}>New blank document</button>
        <button onClick={onOpen}>Open… <kbd>Ctrl+O</kbd></button>
        <div className="home-rail-foot">Prime Productivity</div>
      </nav>
      <main className="home-main">
        <h1>{greeting()}</h1>

        {recovery.length > 0 && (
          <section className="home-recovery" aria-label="Recovered work">
            <h2>Recovered unsaved work</h2>
            {recovery.map((s) => (
              <div key={s.id} className="recovery-item">
                <span><strong>{s.name}</strong> — last autosaved {new Date(s.savedAt).toLocaleString()}</span>
                <button className="primary" onClick={() => onRestore(s)}>Restore</button>
                <button onClick={() => onDiscard(s)}>Discard</button>
              </div>
            ))}
          </section>
        )}

        <section aria-label="New">
          <h2>New</h2>
          <div className="template-grid">
            {TEMPLATES.map((t, i) => (
              <button key={t.id} ref={i === 0 ? first : undefined} className="template-card" onClick={() => onPick(t)}>
                <span className={`template-thumb thumb-${t.id}`} aria-hidden="true"><i /><i /><i /><i /><i /></span>
                <strong>{t.name}</strong>
                <span>{t.description}</span>
              </button>
            ))}
          </div>
        </section>

        <section aria-label="Recent">
          <div className="home-row">
            <h2>Recent</h2>
            <button onClick={onOpen}>Open another document…</button>
          </div>
          {isDesktop ? <RecentList onOpen={onOpenRecent} />
            : <p className="recent-empty">Recent documents are listed in the desktop app. Use Open to choose a .docx file.</p>}
        </section>
      </main>
    </div>
  );
}
