import { type Template, TEMPLATES } from "@prime/editor";
import { type ReactNode, type RefObject, useEffect, useRef, useState } from "react";
import { isDesktop } from "./platform";
import { RecentList } from "./RecentList";
import { MyTemplates } from "./MyTemplates";
import type { Snapshot } from "./recovery";
import { NewProjectDialog, ProjectCard, type ProjectActions, ProjectPage, ProjectsPage } from "./ProjectViews";
import { projectsAvailable, type ProjectSummary } from "./projects";

/** Which page of the home screen is showing (kept by the app, so Home comes back where it was left). */
export type HomeView = { page: "home" | "templates" | "projects" | "recovered"; project?: string };

const greeting = () => {
  const h = new Date().getHours();
  return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
};

/** How many cards of at least `min` px (with `gap` px between) fit across the element. */
function useColumns(ref: RefObject<HTMLElement | null>, min: number, gap: number, page: string): number {
  const [cols, setCols] = useState(4);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const measure = () => setCols(Math.max(1, Math.floor((el.clientWidth + gap) / (min + gap))));
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [ref, min, gap, page]);
  return cols;
}

/** A home section: its title, what can be done there on the right, then its content. */
function Section({ title, actions, children, className }: { title: string; actions?: ReactNode; children: ReactNode; className?: string }) {
  return (
    <section className={`home-section${className ? ` ${className}` : ""}`} aria-label={title}>
      <div className="home-section-head">
        <h2>{title}</h2>
        {actions && <div className="home-section-actions">{actions}</div>}
      </div>
      {children}
    </section>
  );
}

/** The built-in templates as cards. */
function TemplateCards({ onPick, first }: { onPick(t: Template): void; first?: RefObject<HTMLButtonElement> }) {
  return (
    <>
      {TEMPLATES.map((t, i) => (
        <button key={t.id} ref={i === 0 ? first : undefined} className="template-card" onClick={() => onPick(t)}>
          <span className={`template-thumb thumb-${t.id}`} aria-hidden="true"><i /><i /><i /><i /><i /></span>
          <strong>{t.name}</strong>
          <span>{t.description}</span>
        </button>
      ))}
    </>
  );
}

function RecoveryList({ recovery, onRestore, onDiscard }: { recovery: Snapshot[]; onRestore(s: Snapshot): void; onDiscard(s: Snapshot): void }) {
  return (
    <ul className="recovery-list">
      {recovery.map((s) => (
        <li key={s.id}>
          <span className="recovery-name"><strong>{s.name}</strong><span>Last autosaved {new Date(s.savedAt).toLocaleString()}</span></span>
          <button className="primary" onClick={() => onRestore(s)}>Restore</button>
          <button onClick={() => onDiscard(s)}>Discard</button>
        </li>
      ))}
    </ul>
  );
}

/**
 * The home screen: shown when Prime Word starts without a document, and from the Prime Word button.
 *
 * The rail holds the pages (Home, Projects, My templates, Recovered when there is any) and then the
 * actions (new, open, convert). The Home page is one column of sections, each with the same heading:
 * Start new (one row of templates, scrolling sideways), Projects (two rows at most), Recent, My templates
 * (one row), Convert a PDF, and — at the bottom, as one line — unsaved work kept from an earlier session.
 * Longer lists have pages of their own ("View all").
 */
export function Home({ recovery, onPick, onOpen, onOpenRecent, onRestore, onDiscard, onClose, onConvertPdf, onLetterhead, templatesVersion,
  view, onView, projects, onProjectsChanged, projectActions }: {
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
  view: HomeView;
  onView(v: HomeView): void;
  /** The person's projects (desktop), most recently used first. */
  projects: ProjectSummary[];
  onProjectsChanged(): void;
  projectActions: ProjectActions;
}) {
  const first = useRef<HTMLButtonElement>(null);
  const content = useRef<HTMLDivElement>(null);
  const page = view.page === "recovered" && !recovery.length ? "home" : view.page;
  const setPage = (p: HomeView["page"]) => onView({ page: p });
  const [creating, setCreating] = useState(false);
  const [templateCount, setTemplateCount] = useState(0);
  useEffect(() => { first.current?.focus(); }, [page]);
  const current = view.project ? projects.find((p) => p.path === view.project) : undefined;
  const openProject = (p: ProjectSummary) => onView({ page: "projects", project: p.path });
  const showProjects = projectsAvailable();
  // Two rows of project cards (200px at least), one row of my templates (170px at least).
  const projectCols = useColumns(content, 200, 12, page);
  const templateCols = useColumns(content, 170, 12, page);
  const projectLimit = projectCols * 2;
  const railPage = (p: HomeView["page"], label: ReactNode) => (
    <button className={page === p ? "on" : ""} aria-current={page === p ? "page" : undefined} onClick={() => setPage(p)}>{label}</button>
  );
  return (
    <div className="home" role="dialog" aria-label="Prime Word home" onKeyDown={(e) => {
      if (e.key === "Escape" && onClose) onClose();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") { e.preventDefault(); onOpen(); }
    }}>
      <nav className="home-rail" aria-label="Home">
        <div className="home-logo">Prime Word</div>
        {onClose && <button onClick={onClose}>← Back to document</button>}
        {railPage("home", "Home")}
        {showProjects && railPage("projects", "Projects")}
        {railPage("templates", "My templates")}
        {recovery.length > 0 && railPage("recovered", <>Recovered <span className="rail-count">{recovery.length}</span></>)}
        <hr className="rail-divider" />
        <button onClick={() => onPick(TEMPLATES[0])}>New blank document</button>
        <button onClick={onOpen}>Open… <kbd>Ctrl+O</kbd></button>
        <button onClick={onConvertPdf}>Convert PDF to Word…</button>
        <div className="home-rail-foot">Prime Productivity</div>
      </nav>

      {page === "projects" ? (current ? (
        <ProjectPage key={current.path} project={current} actions={projectActions} onBack={() => onView({ page: "projects" })}
          onChanged={onProjectsChanged} onForgotten={() => { onProjectsChanged(); onView({ page: "projects" }); }} />
      ) : (
        <ProjectsPage projects={projects} onOpen={openProject} onNew={() => setCreating(true)} />
      )) : page === "templates" ? (
        <main className="home-main">
          <div className="home-content">
            <div className="home-title">
              <h1>My templates</h1>
              <button ref={first} className="primary" onClick={onLetterhead}>Design a letterhead…</button>
            </div>
            <p className="home-hint">Templates and letterheads you save are kept on this computer. To save one, open a document and choose
              File → Save as Template; every new document made from it starts as a copy.</p>
            <MyTemplates onPick={onPick} refresh={templatesVersion} />
            <Section title="Built-in templates">
              <div className="template-grid"><TemplateCards onPick={onPick} /></div>
            </Section>
          </div>
        </main>
      ) : page === "recovered" ? (
        <main className="home-main">
          <div className="home-content">
            <div className="home-title"><h1>Recovered work</h1></div>
            <p className="home-hint">Prime Word keeps unsaved changes as you type. These documents were not saved before Prime Word
              last closed: restore one to carry on (it opens with its changes, still unsaved), or discard it.</p>
            <RecoveryList recovery={recovery} onRestore={onRestore} onDiscard={onDiscard} />
          </div>
        </main>
      ) : (
        <main className="home-main">
          <div className="home-content" ref={content}>
            <div className="home-title">
              <h1>{greeting()}</h1>
              <div className="home-title-actions">
                <button className="primary" onClick={() => onPick(TEMPLATES[0])}>New document</button>
                <button onClick={onOpen}>Open…</button>
              </div>
            </div>

            <Section title="Start new">
              <div className="template-strip" role="list"><TemplateCards onPick={onPick} first={first} /></div>
            </Section>

            {showProjects && (
              <Section title="Projects" actions={<>
                <button onClick={() => setCreating(true)}>New project…</button>
                {projects.length > projectLimit && <button className="link" onClick={() => setPage("projects")}>View all ({projects.length}) →</button>}
              </>}>
                {projects.length ? (
                  <div className="project-grid">
                    {projects.slice(0, projectLimit).map((p) => <ProjectCard key={p.path} p={p} onOpen={() => openProject(p)} />)}
                  </div>
                ) : (
                  <p className="home-hint">Keep everything for a piece of work together — a thesis, a client, a property sale. A project
                    is a folder on your computer: its documents, PDFs and pictures stay there, where you and your backups can find them.</p>
                )}
              </Section>
            )}

            <Section title="Recent" actions={<button onClick={onOpen}>Open another…</button>}>
              {isDesktop ? <RecentList onOpen={onOpenRecent} limit={6} columns />
                : <p className="recent-empty">Recent documents are listed in the desktop app. Use Open to choose a .docx file.</p>}
            </Section>

            <Section title="My templates" actions={<>
              <button onClick={onLetterhead}>Design a letterhead…</button>
              {templateCount > templateCols && <button className="link" onClick={() => setPage("templates")}>View all ({templateCount}) →</button>}
            </>}>
              <MyTemplates onPick={onPick} refresh={templatesVersion} limit={templateCols} onCount={setTemplateCount} />
            </Section>

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
                <p>Turn a PDF into an editable Word document, checked page by page against the original.</p>
              </div>
              <button className="primary" onClick={onConvertPdf}>Convert PDF to Word…</button>
            </section>

            {recovery.length > 0 && (
              <div className="home-recovered-note" role="status">
                <span>{recovery.length === 1 ? `${recovery[0].name} has` : `${recovery.length} documents have`} unsaved changes kept from an earlier session.</span>
                <button className="link" onClick={() => setPage("recovered")}>Review →</button>
              </div>
            )}
          </div>
        </main>
      )}
      {creating && <NewProjectDialog onClose={() => setCreating(false)} onCreated={(p) => { setCreating(false); onProjectsChanged(); onView({ page: "projects", project: p.path }); }} />}
    </div>
  );
}
