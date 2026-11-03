import { citationText, insertCitation, newTag, type Person, readSources, type Source, type SourceType, writeSources } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useState } from "react";

const TYPES: [SourceType, string][] = [["Book", "Book"], ["JournalArticle", "Journal article"], ["InternetSite", "Website"], ["Report", "Report"]];
const EMPTY: Source = { tag: "", type: "Book", authors: [], title: "" };

/** "Mensah, Ama" per line → people; a line without a comma is an organisation. */
const parseAuthors = (text: string): Person[] => text.split("\n").map((l) => l.trim()).filter(Boolean).map((l) => {
  const i = l.indexOf(",");
  return i < 0 ? { last: l } : { last: l.slice(0, i).trim(), first: l.slice(i + 1).trim() || undefined };
});
const authorsText = (a: Person[]) => a.map((p) => (p.first ? `${p.last}, ${p.first}` : p.last)).join("\n");

/**
 * References → Manage Sources: the document's sources (kept in it, as Word keeps them), a form to add
 * or edit one, and Cite to put a citation at the cursor (with a page, if given).
 */
export function SourcesDialog({ view, doc, startNew, onChanged, onClose }: {
  view: EditorView; doc: () => PrimeDocument; startNew?: boolean; onChanged(): void; onClose(): void;
}) {
  const initial = readSources(doc());
  const [sources, setSources] = useState(initial.sources);
  const [editing, setEditing] = useState<Source | undefined>(startNew ? { ...EMPTY } : undefined);
  const [authors, setAuthors] = useState("");
  const [page, setPage] = useState("");
  const style = initial.style;

  const save = (list: Source[]) => { writeSources(doc(), style, list); setSources(list); onChanged(); };
  const edit = (s: Source) => { setEditing({ ...s }); setAuthors(authorsText(s.authors)); };
  const set = <K extends keyof Source>(k: K, v: Source[K]) => setEditing((e) => (e ? { ...e, [k]: v || undefined } as Source : e));
  const commit = () => {
    if (!editing || !editing.title.trim()) return;
    const s = { ...editing, authors: parseAuthors(authors), title: editing.title.trim() };
    const list = editing.tag ? sources.map((x) => (x.tag === editing.tag ? s : x)) : [...sources, { ...s, tag: newTag(s, sources) }];
    save(list);
    setEditing(undefined);
    setAuthors("");
  };
  const cite = (s: Source) => {
    insertCitation(doc, s.tag, page.trim() || undefined)(view.state, view.dispatch);
    onClose();
    view.focus();
  };

  const field = (label: string, k: keyof Source, placeholder = "") => (
    <label>{label}<input value={(editing?.[k] as string | undefined) ?? ""} placeholder={placeholder} onChange={(e) => set(k, e.target.value as never)} /></label>
  );

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog sources-dialog" role="dialog" aria-label="Sources">
        <h3>{editing ? (editing.tag ? "Edit source" : "New source") : "Sources"}</h3>
        {editing ? (
          <div className="source-form">
            <label>Type<select value={editing.type} onChange={(e) => set("type", e.target.value as SourceType)}>
              {TYPES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select></label>
            <label className="wide">Authors (one per line: Last, First — or an organisation)
              <textarea rows={3} value={authors} onChange={(e) => setAuthors(e.target.value)} placeholder={"Mensah, Ama\nOwusu, Kwame"} autoFocus /></label>
            <label className="wide">Title<input value={editing.title} onChange={(e) => set("title", e.target.value)} /></label>
            {field("Year", "year", "2024")}
            {editing.type === "JournalArticle" ? (
              <>{field("Journal", "journal")}{field("Volume", "volume")}{field("Issue", "issue")}{field("Pages", "pages", "12–19")}{field("DOI or URL", "url")}</>
            ) : editing.type === "InternetSite" ? (
              <>{field("Website name", "publisher")}{field("URL", "url", "https://")}
                <label>Accessed<input type="date" value={editing.accessed ?? ""} onChange={(e) => set("accessed", e.target.value)} /></label></>
            ) : (
              <>{field(editing.type === "Report" ? "Institution" : "Publisher", "publisher")}{field("City", "city")}</>
            )}
            <div className="hf-actions wide">
              <button onClick={() => { setEditing(undefined); setAuthors(""); }}>Back</button>
              <button className="primary" disabled={!editing.title.trim()} onClick={commit}>Save source</button>
            </div>
          </div>
        ) : (
          <>
            <p className="hint">Style: {style}. Sources are saved in this document, as Word saves them. Choose Cite to put a citation at the cursor.</p>
            <div className="source-list">
              {sources.length === 0 && <p className="hint">No sources yet.</p>}
              {sources.map((s) => (
                <div key={s.tag} className="source">
                  <div><strong>{s.title}</strong><span>{citationText(s, style === "IEEE" ? "APA" : style, 0)} · {TYPES.find(([v]) => v === s.type)?.[1]}</span></div>
                  <button className="primary" onClick={() => cite(s)}>Cite</button>
                  <button onClick={() => edit(s)}>Edit</button>
                  <button className="link" aria-label={`Delete ${s.title}`} onClick={() => { if (window.confirm(`Delete the source "${s.title}"? Citations of it will show "source missing".`)) save(sources.filter((x) => x !== s)); }}>Delete</button>
                </div>
              ))}
            </div>
            <div className="hf-actions">
              <label className="sig-name">Page (optional) <input value={page} onChange={(e) => setPage(e.target.value)} placeholder="12" style={{ width: 60 }} /></label>
              <button onClick={() => { setEditing({ ...EMPTY }); setAuthors(""); }}>New source…</button>
              <button className="primary" onClick={onClose}>Close</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
