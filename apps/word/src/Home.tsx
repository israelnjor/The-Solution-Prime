import { type Template, TEMPLATES } from "@prime/editor";
import { useEffect, useRef, useState } from "react";
import { isDesktop } from "./platform";
import { RecentList } from "./RecentList";
import { MyTemplates } from "./MyTemplates";
import type { Snapshot } from "./recovery";

const greeting = () => {
  const h = new Date().getHours();
  return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
};

/**
 * The home screen: shown when Prime Word starts without a document, and from the Prime Word button.
 * New (templates), Open, recent documents, and unsaved work to recover.
 */
export function Home({ recovery, onPick, onOpen, onOpenRecent, onRestore, onDiscard, onClose, onConvertPdf, onLetterhead, templatesVersion }: {
  recovery: Snapshot[];
  onPick(t: Template): void;
  onOpen(): void;
  onOpenRecent(path: string): void;
  onRestore(s: Snapshot): void;
  onDiscard(s: Snapshot): void;
  /** Back to the open document (absent when there is none to go back to). */
  onClose?(): void;
  onConvertPdf(): void;
  onLetterhead(): void;
  templatesVersion: number;
}) {
  const first = useRef<HTMLButtonElement>(null);
  /** The rail's pages: everything at a glance, or the person's own templates. */
  const [page, setPage] = useState<"home" | "templates">("home");
  useEffect(() => { first.current?.focus(); }, [page]);
  return (
    <div className="home" role="dialog" aria-label="Prime Word home" onKeyDown={(e) => {
      if (e.key === "Escape" && onClose) onClose();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") { e.preventDefault(); onOpen(); }
    }}>
      <nav className="home-rail" aria-label="Home">
        <div className="home-logo">Prime Word</div>
        {onClose && <button onClick={onClose}>← Back to document</button>}
        <button className={page === "home" ? "on" : ""} aria-current={page === "home" ? "page" : undefined} onClick={() => setPage("home")}>Home</button>
        <button className={page === "templates" ? "on" : ""} aria-current={page === "templates" ? "page" : undefined} onClick={() => setPage("templates")}>My templates</button>
        <button onClick={() => onPick(TEMPLATES[0])}>New blank document</button>
        <button onClick={onOpen}>Open… <kbd>Ctrl+O</kbd></button>
        <button onClick={onConvertPdf}>Convert PDF to Word…</button>
        <div className="home-rail-foot">Prime Productivity</div>
      </nav>
      {page === "templates" ? (
        <main className="home-main">
          <div className="home-row">
            <h1>My templates</h1>
            <button ref={first} className="primary" onClick={onLetterhead}>Design a letterhead…</button>
          </div>
          <p className="home-hint">Templates and letterheads you save are kept on this computer. To save one, open a document and choose
            File → Save as Template; every new document made from it starts as a copy.</p>
          <MyTemplates onPick={onPick} refresh={templatesVersion} />
        </main>
      ) : (
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

        <section aria-label="My templates">
          <div className="home-row">
            <h2>My templates</h2>
            <button onClick={onLetterhead}>Design a letterhead…</button>
          </div>
          <MyTemplates onPick={onPick} refresh={templatesVersion} />
        </section>

        <section aria-label="Convert PDF" className="home-convert">
          {/* A PDF page turning into a Word page. */}
          <svg className="home-convert-icon" viewBox="0 0 120 64" aria-hidden="true">
            <path d="M8 4h26l12 12v44H8z" fill="#fff" stroke="#d93025" strokeWidth="2.5" strokeLinejoin="round" />
            <path d="M34 4v12h12" fill="#fde3e1" stroke="#d93025" strokeWidth="2.5" strokeLinejoin="round" />
            <rect x="3" y="30" width="36" height="16" rx="3" fill="#d93025" />
            <text x="21" y="42.5" textAnchor="middle" fontFamily="Segoe UI, Arial, sans-serif" fontWeight="700" fontSize="12" fill="#fff">PDF</text>
            <path d="M54 32h14m-6-7 7 7-7 7" fill="none" stroke="#8a94a6" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M78 4h26l12 12v44H78z" fill="#fff" stroke="#2b579a" strokeWidth="2.5" strokeLinejoin="round" />
            <path d="M104 4v12h12" fill="#dce6f4" stroke="#2b579a" strokeWidth="2.5" strokeLinejoin="round" />
            <rect x="71" y="30" width="40" height="16" rx="3" fill="#2b579a" />
            <text x="91" y="42.5" textAnchor="middle" fontFamily="Segoe UI, Arial, sans-serif" fontWeight="700" fontSize="10" fill="#fff">DOCX</text>
          </svg>
          <div className="home-convert-text">
            <h2>Convert a PDF</h2>
            <p>Turn a PDF into an editable Word document: compare it page by page with the original, then edit or save it.</p>
          </div>
          <button className="home-convert-button" onClick={onConvertPdf}>Convert PDF to Word…</button>
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
      )}
    </div>
  );
}
