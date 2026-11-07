import { type DocumentProperties, readProperties } from "@prime/docx";
import type { PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { useMemo, useState } from "react";

const when = (iso?: string) => {
  const d = iso ? new Date(iso) : undefined;
  return d && !Number.isNaN(d.getTime()) ? d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" }) : "—";
};

/** Counts for File → Info, as Word's Word Count gives them. */
function statistics(doc: Node) {
  let words = 0;
  let chars = 0;
  let charsNoSpaces = 0;
  let paragraphs = 0;
  doc.descendants((n) => {
    if (!n.isTextblock) return true;
    const text = n.textContent;
    if (text.trim()) paragraphs++;
    words += (text.match(/\S+/g) ?? []).length;
    chars += text.length;
    charsNoSpaces += text.replace(/\s/g, "").length;
    return false;
  });
  return { words, chars, charsNoSpaces, paragraphs };
}

/** File → Info: the document's properties (title, author, tags…) to edit, and its statistics. */
export function PropertiesDialog({ base, doc, pages, fileName, onSave, onClose }: {
  base: PrimeDocument; doc: Node; pages?: number; fileName: string; onSave(props: DocumentProperties): void; onClose(): void;
}) {
  const initial = useMemo(() => readProperties(base), [base]);
  const [p, setP] = useState<DocumentProperties>(initial);
  const stats = useMemo(() => statistics(doc), [doc]);
  const field = (key: "title" | "subject" | "author" | "keywords" | "category" | "status", label: string, hint?: string) => (
    <label>{label}<input value={p[key] ?? ""} placeholder={hint} aria-label={label} onChange={(e) => setP({ ...p, [key]: e.target.value })} /></label>
  );
  const save = () => {
    const changed: DocumentProperties = {};
    for (const k of ["title", "subject", "author", "keywords", "comments", "category", "status"] as const) {
      if ((p[k] ?? "") !== (initial[k] ?? "")) changed[k] = p[k] ?? "";
    }
    if (Object.keys(changed).length) onSave(changed);
    onClose();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog properties-dialog" role="dialog" aria-label="Document properties">
        <h3>Info — {fileName}</h3>
        <div className="properties-grid">
          <div className="properties-fields">
            {field("title", "Title", "Add a title")}
            {field("subject", "Subject")}
            {field("author", "Author")}
            {field("keywords", "Tags", "Add tags (keywords)")}
            {field("category", "Category")}
            {field("status", "Status", "e.g. Draft, Final")}
            <label>Comments<textarea rows={3} value={p.comments ?? ""} aria-label="Comments" onChange={(e) => setP({ ...p, comments: e.target.value })} /></label>
          </div>
          <dl className="properties-stats">
            <dt>Pages</dt><dd>{pages ?? "—"}</dd>
            <dt>Words</dt><dd>{stats.words.toLocaleString()}</dd>
            <dt>Characters (no spaces)</dt><dd>{stats.charsNoSpaces.toLocaleString()}</dd>
            <dt>Characters (with spaces)</dt><dd>{stats.chars.toLocaleString()}</dd>
            <dt>Paragraphs</dt><dd>{stats.paragraphs.toLocaleString()}</dd>
            <dt>Created</dt><dd>{when(initial.created)}</dd>
            <dt>Last modified</dt><dd>{when(initial.modified)}</dd>
            <dt>Last modified by</dt><dd>{initial.lastModifiedBy || "—"}</dd>
            <dt>Revision</dt><dd>{initial.revision ?? "—"}</dd>
          </dl>
        </div>
        <div className="buttons">
          <button className="primary" onClick={save}>OK</button>
          <button onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
