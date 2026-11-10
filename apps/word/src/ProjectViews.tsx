import { type Template, TEMPLATES } from "@prime/editor";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { confirmAction } from "./platform";
import {
  addProject, chooseFolder, defaultProjectsFolder, type FileKind, forgetProject, importFiles, joinPath, kindOf,
  openExternal, PROJECT_COLORS, type ProjectFile, projectFiles, type ProjectSummary, saveProject,
} from "./projects";

/** What a project page can ask the app to do with its files. */
export interface ProjectActions {
  /** Opens a document (in a tab). */
  openDocument(path: string): void;
  /** A new document from a template, saved in the project folder at once, then opened. */
  newDocument(folder: string, name: string, template: Template): Promise<void>;
  /** Converts a PDF (one in the project, or chosen) into a document saved in the project. */
  convertPdf(folder: string, pdfPath?: string): void;
  /** Opens several documents, each in a tab. */
  openAll(paths: string[]): void;
}

const when = (ms: number) => {
  if (!ms) return "";
  const d = new Date(ms);
  const days = Math.floor((Date.now() - ms) / 86_400_000);
  if (days < 1 && new Date().getDate() === d.getDate()) return `Today ${d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
  if (days < 2) return "Yesterday";
  if (days < 7) return d.toLocaleDateString([], { weekday: "long" });
  return d.toLocaleDateString([], { day: "numeric", month: "short", year: d.getFullYear() === new Date().getFullYear() ? undefined : "numeric" });
};

const size = (n: number) => (n < 1024 ? `${n} B` : n < 1024 ** 2 ? `${Math.round(n / 1024)} KB` : `${(n / 1024 ** 2).toFixed(1)} MB`);

/** A project's card: its colour, name, how many files, when it was last worked on. */
export function ProjectCard({ p, onOpen }: { p: ProjectSummary; onOpen(): void }) {
  return (
    <button className={`project-card${p.exists ? "" : " missing"}`} onClick={onOpen} style={{ ["--project" as string]: `#${p.meta.color ?? PROJECT_COLORS[0]}` }}>
      <span className="project-card-band" aria-hidden="true" />
      <strong>{p.meta.name}</strong>
      <span>{p.exists ? `${p.files} ${p.files === 1 ? "file" : "files"}${p.modified ? ` · ${when(p.modified)}` : ""}` : "Folder not found"}</span>
    </button>
  );
}

/** All projects (the rail's Projects page). */
export function ProjectsPage({ projects, onOpen, onNew }: { projects: ProjectSummary[]; onOpen(p: ProjectSummary): void; onNew(): void }) {
  const [query, setQuery] = useState("");
  const shown = projects.filter((p) => p.meta.name.toLowerCase().includes(query.trim().toLowerCase()));
  return (
    <main className="home-main">
      <div className="home-content">
      <div className="home-title">
        <h1>Projects</h1>
        <button className="primary" onClick={onNew}>New project…</button>
      </div>
      {projects.length > 6 && <input className="project-search" placeholder="Find a project" aria-label="Find a project" value={query} onChange={(e) => setQuery(e.target.value)} />}
      {shown.length ? (
        <div className="project-grid">{shown.map((p) => <ProjectCard key={p.path} p={p} onOpen={() => onOpen(p)} />)}</div>
      ) : (
        <p className="home-hint">{projects.length ? "No project has that name." : "No projects yet. A project is a folder that keeps a piece of work together: its documents, the PDFs that came with it, pictures and anything else."}</p>
      )}
      </div>
    </main>
  );
}

/**
 * New project: a name and colour, and its folder — a new one (in Documents unless the person picks
 * somewhere else), or a folder they already have. Nothing is created until Create project.
 */
export function NewProjectDialog({ onCreated, onClose }: { onCreated(p: ProjectSummary): void; onClose(): void }) {
  const [name, setName] = useState("");
  const [color, setColor] = useState(PROJECT_COLORS[0]);
  const [parent, setParent] = useState("");
  /** A folder the person already has, chosen to become the project (its files stay as they are). */
  const [existing, setExisting] = useState<string | undefined>();
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { input.current?.focus(); void defaultProjectsFolder().then(setParent); }, []);
  const clean = name.replace(/[\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().replace(/\.+$/, "");
  const lastPart = (folder: string) => folder.split(/[\/]/).filter(Boolean).pop() ?? "";
  const folder = existing ?? (clean && parent ? joinPath(parent, clean) : "");
  const create = async () => {
    if (!clean || !folder || busy) return;
    setBusy(true);
    setError("");
    try {
      onCreated(await addProject(folder, { name: clean, color }, !existing));
    } catch (e) {
      setError(String((e as Error).message ?? e));
      setBusy(false);
    }
  };
  const pickExisting = async () => {
    const f = await chooseFolder("Choose the folder to make a project", existing ?? (parent || undefined));
    if (!f) return;
    setExisting(f);
    setError("");
    if (!clean) setName(lastPart(f));
    input.current?.focus();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog project-dialog" role="dialog" aria-label="New project">
        <h3>New project</h3>
        <label>Name<input ref={input} value={name} onChange={(e) => { setName(e.target.value); setError(""); }} placeholder="e.g. Casa Viola — Lashibi sale"
          aria-label="Project name" onKeyDown={(e) => { if (e.key === "Enter") void create(); }} /></label>
        <div className="project-colors" role="radiogroup" aria-label="Colour">
          {PROJECT_COLORS.map((c) => (
            <button key={c} role="radio" aria-checked={c === color} aria-label={`Colour ${c}`} className={c === color ? "on" : ""}
              style={{ background: `#${c}` }} onClick={() => setColor(c)} />
          ))}
        </div>
        {existing ? (
          <div className="project-where">
            <span>Your folder: <strong>{existing}</strong><br />It becomes the project; the files in it stay as they are.</span>
            <button onClick={() => void pickExisting()}>Change…</button>
            <button onClick={() => setExisting(undefined)}>New folder instead</button>
          </div>
        ) : (
          <div className="project-where">
            <span>A new folder: <strong>{folder || (parent ? joinPath(parent, "…") : "…")}</strong></span>
            <button onClick={async () => { const f = await chooseFolder("Where to put the project's folder", parent || undefined); if (f) setParent(f); }}>Change…</button>
            <button onClick={() => void pickExisting()} title="Make a folder you already have into a project (its files stay as they are)">Use a folder I have…</button>
          </div>
        )}
        {error && <p className="error">{error}</p>}
        <div className="buttons">
          <button className="primary" disabled={!clean || !folder || busy} onClick={() => void create()}>{busy ? "Creating…" : "Create project"}</button>
          <button onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

const GROUPS: [string, FileKind[]][] = [["Documents", ["document"]], ["PDFs", ["pdf"]], ["Pictures", ["image"]], ["Other files", ["sheet", "slides", "other"]]];
const BADGE: Record<FileKind, string> = { document: "DOC", pdf: "PDF", image: "IMG", sheet: "XLS", slides: "PPT", other: "FILE" };

/** One project: its files (by kind, newest first) and what can be done in it. */
export function ProjectPage({ project, actions, onBack, onChanged, onForgotten }: {
  project: ProjectSummary; actions: ProjectActions; onBack(): void; onChanged(): void; onForgotten(): void;
}) {
  const [files, setFiles] = useState<ProjectFile[] | undefined>();
  const [error, setError] = useState("");
  const [query, setQuery] = useState("");
  const [renaming, setRenaming] = useState(false);
  const [newDoc, setNewDoc] = useState<{ name: string; template: string } | undefined>();
  const [busy, setBusy] = useState(false);
  /** Groups shown in full (long ones show their newest 200 until asked). */
  const [expanded, setExpanded] = useState<string[]>([]);
  const meta = project.meta;
  const refresh = useCallback(() => {
    projectFiles(project.path).then((f) => { setFiles(f); setError(""); }, (e) => setError(String((e as Error).message ?? e)));
  }, [project.path]);
  // Files change outside Prime Word too (Explorer, a download): looked at again when the window regains focus.
  useEffect(() => {
    refresh();
    window.addEventListener("focus", refresh);
    return () => window.removeEventListener("focus", refresh);
  }, [refresh]);
  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return (files ?? []).filter((f) => !q || f.name.toLowerCase().includes(q) || f.folder.toLowerCase().includes(q)).sort((a, b) => b.modified - a.modified);
  }, [files, query]);
  const documents = (files ?? []).filter((f) => kindOf(f.name) === "document");
  const update = async (m: Partial<typeof meta>) => { await saveProject(project.path, { ...meta, ...m }); onChanged(); };
  const createDoc = async () => {
    if (!newDoc?.name.trim()) return;
    setBusy(true);
    try {
      await actions.newDocument(project.path, newDoc.name.trim(), TEMPLATES.find((t) => t.id === newDoc.template) ?? TEMPLATES[0]);
      setNewDoc(undefined);
    } catch (e) { setError(String((e as Error).message ?? e)); }
    setBusy(false);
  };
  const open = (f: ProjectFile) => {
    const k = kindOf(f.name);
    if (k === "document") actions.openDocument(f.path);
    else void openExternal(f.path).catch((e) => setError(String(e)));
  };
  const convert = async (f: ProjectFile) => {
    if (await confirmAction(`Convert “${f.name}” into an editable Word document?

A new document is saved in this project; the PDF stays as it is.`, "Convert PDF")) {
      actions.convertPdf(project.path, f.path);
    }
  };
  return (
    <main className="home-main project-page" style={{ ["--project" as string]: `#${meta.color}` }}>
      <div className="home-content">
      <button className="project-back" onClick={onBack}>← All projects</button>
      <div className="project-head">
        <span className="project-dot" aria-hidden="true" />
        {renaming ? (
          <input className="project-rename" autoFocus defaultValue={meta.name} aria-label="Project name"
            onBlur={(e) => { const n = e.target.value.trim(); setRenaming(false); if (n && n !== meta.name) void update({ name: n }); }}
            onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") setRenaming(false); }} />
        ) : (
          <h1 title="Click to rename" onClick={() => setRenaming(true)}>{meta.name}</h1>
        )}
        <div className="project-colors small" role="radiogroup" aria-label="Colour">
          {PROJECT_COLORS.map((c) => (
            <button key={c} role="radio" aria-checked={c === meta.color} aria-label={`Colour ${c}`} className={c === meta.color ? "on" : ""}
              style={{ background: `#${c}` }} onClick={() => void update({ color: c })} />
          ))}
        </div>
      </div>
      <p className="project-path">
        <span title={project.path}>{project.path}</span>
        <button onClick={() => void openExternal(project.path).catch((e) => setError(String(e)))}>Show in folder</button>
        <button onClick={async () => { await forgetProject(project.path); onForgotten(); }}
          title="Take the project off Prime Word's list. Its folder and files stay as they are.">Remove from list</button>
      </p>
      {!project.exists && <p className="error">The project's folder is no longer at this place. If you moved it, use New project → Use an existing folder.</p>}

      <div className="project-actions">
        <button className="primary" onClick={() => setNewDoc({ name: "", template: "blank" })}>New document</button>
        <button onClick={() => actions.convertPdf(project.path)}>Convert a PDF into this project…</button>
        <button onClick={async () => { try { if ((await importFiles(project.path)).length) refresh(); } catch (e) { setError(String((e as Error).message ?? e)); } }}>Add files…</button>
        <button disabled={!documents.length} onClick={() => actions.openAll(documents.sort((a, b) => b.modified - a.modified).map((f) => f.path))}>
          Open all documents{documents.length ? ` (${documents.length})` : ""}
        </button>
      </div>
      {newDoc && (
        <div className="project-newdoc" role="group" aria-label="New document">
          <input autoFocus placeholder="Document name" aria-label="Document name" value={newDoc.name}
            onChange={(e) => setNewDoc({ ...newDoc, name: e.target.value })}
            onKeyDown={(e) => { if (e.key === "Enter") void createDoc(); if (e.key === "Escape") setNewDoc(undefined); }} />
          <select aria-label="Template" value={newDoc.template} onChange={(e) => setNewDoc({ ...newDoc, template: e.target.value })}>
            {TEMPLATES.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
          </select>
          <button className="primary" disabled={!newDoc.name.trim() || busy} onClick={() => void createDoc()}>Create</button>
          <button onClick={() => setNewDoc(undefined)}>Cancel</button>
        </div>
      )}
      {error && <p className="error">{error}</p>}

      {files === undefined ? <p className="home-hint">Looking in the folder…</p> : files.length === 0 ? (
        <p className="home-hint">No files yet. Start a new document, convert a PDF, or add files you already have — they are copied into the project's folder.</p>
      ) : (
        <>
          {files.length > 8 && <input className="project-search" placeholder="Find a file" aria-label="Find a file" value={query} onChange={(e) => setQuery(e.target.value)} />}
          {GROUPS.map(([title, kinds]) => {
            const list = shown.filter((f) => kinds.includes(kindOf(f.name)));
            if (!list.length) return null;
            return (
              <section key={title} aria-label={title}>
                <h2>{title} <span className="project-count">{list.length}</span></h2>
                <ul className="project-files">
                  {(expanded.includes(title) ? list : list.slice(0, 200)).map((f) => {
                    const k = kindOf(f.name);
                    return (
                      <li key={f.path}>
                        <button className="project-file" onClick={() => open(f)}
                          title={k === "document" ? `Open ${f.name}` : `Open ${f.name} in its program`}>
                          <span className={`file-badge ${k}`} aria-hidden="true">{BADGE[k]}</span>
                          <span className="project-file-name">
                            <strong>{f.name}</strong>
                            <span>{[f.folder, when(f.modified), size(f.size)].filter(Boolean).join(" · ")}</span>
                          </span>
                        </button>
                        {k === "pdf" && <button className="project-file-side" onClick={() => void convert(f)} title="Make an editable Word document from this PDF (asks first)">Convert…</button>}
                      </li>
                    );
                  })}
                </ul>
                {list.length > 200 && !expanded.includes(title) && (
                  <button className="project-more" onClick={() => setExpanded((x) => [...x, title])}>Show all {list.length} {title.toLowerCase()}</button>
                )}
              </section>
            );
          })}
        </>
      )}
      </div>
    </main>
  );
}
