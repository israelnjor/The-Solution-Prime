import { captions, headings, insertCaption, insertCrossReference, type Pages, type RefShow } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useMemo, useState } from "react";

const LABELS = ["Figure", "Table", "Equation"];

/** References → Insert Caption: "Figure 3: …" below (or above) the picture or table at the cursor. */
export function CaptionDialog({ view, doc, onClose }: { view: EditorView; doc: () => PrimeDocument; onClose(): void }) {
  const [label, setLabel] = useState(() => (view.state.selection.$from.node(1)?.type.name === "table" ? "Table" : "Figure"));
  const [text, setText] = useState("");
  const [where, setWhere] = useState<"below" | "above">("below");
  const next = captions(view.state, label).length + 1;
  const insert = () => {
    insertCaption(doc, label, text.trim(), where)(view.state, view.dispatch, view);
    onClose();
    view.focus();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter") insert(); }}>
      <div className="hf-dialog caption-dialog" role="dialog" aria-label="Caption">
        <h3>Caption</h3>
        <div className="caption-preview">{label} {next}{text.trim() ? `: ${text.trim()}` : ""}</div>
        <label>Caption text <input value={text} onChange={(e) => setText(e.target.value)} placeholder="e.g. Monthly sales by region" autoFocus /></label>
        <div className="row">
          <label>Label <select value={label} onChange={(e) => setLabel(e.target.value)}>{LABELS.map((l) => <option key={l}>{l}</option>)}</select></label>
          <label>Position <select value={where} onChange={(e) => setWhere(e.target.value as "below" | "above")}>
            <option value="below">Below selected item</option><option value="above">Above selected item</option>
          </select></label>
        </div>
        <p className="hint">Captions are numbered automatically; a table of figures lists them with their pages.</p>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={insert}>OK</button>
        </div>
      </div>
    </div>
  );
}

type Kind = "Heading" | "Figure" | "Table" | "Equation";

/** References → Cross-reference: to a heading or caption, as its text, its number, its page, or above/below. */
export function CrossRefDialog({ view, doc, pages, onClose }: { view: EditorView; doc: () => PrimeDocument; pages: Pages; onClose(): void }) {
  const [kind, setKind] = useState<Kind>("Heading");
  const [show, setShow] = useState<RefShow>("text");
  const targets = useMemo(() => (kind === "Heading" ? headings(view.state, doc()) : captions(view.state, kind)), [kind, view, doc]);
  const [pick, setPick] = useState(0);
  const target = targets[pick];
  const shows: [RefShow, string][] = kind === "Heading"
    ? [["text", "Heading text"], ["page", "Page number"], ["aboveBelow", "Above/below"]]
    : [["labelNumber", `Only label and number ("${kind} 2")`], ["text", "Entire caption"], ["page", "Page number"], ["aboveBelow", "Above/below"]];
  const insert = () => {
    if (!target) return;
    insertCrossReference(target, shows.some(([s]) => s === show) ? show : shows[0][0], pages)(view.state, view.dispatch);
    onClose();
    view.focus();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog crossref-dialog" role="dialog" aria-label="Cross-reference">
        <h3>Cross-reference</h3>
        <div className="row">
          <label>Reference type <select value={kind} onChange={(e) => { setKind(e.target.value as Kind); setPick(0); setShow(e.target.value === "Heading" ? "text" : "labelNumber"); }}>
            {(["Heading", ...LABELS] as Kind[]).map((k) => <option key={k}>{k}</option>)}
          </select></label>
          <label>Insert reference to <select value={show} onChange={(e) => setShow(e.target.value as RefShow)}>
            {shows.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select></label>
        </div>
        <div className="crossref-list" role="listbox" aria-label={`${kind}s`}>
          {targets.length === 0 && <p className="hint">{kind === "Heading" ? "No headings yet: use the Heading styles." : `No ${kind.toLowerCase()} captions yet: use Insert Caption.`}</p>}
          {targets.map((t, i) => (
            <button key={t.pos} role="option" aria-selected={i === pick} className={i === pick ? "on" : ""} style={{ paddingLeft: 10 + (kind === "Heading" ? (t.level - 1) * 14 : 0) }}
              onClick={() => setPick(i)} onDoubleClick={() => { setPick(i); setTimeout(insert, 0); }}>{t.text}</button>
          ))}
        </div>
        <p className="hint">Ctrl+Click a cross-reference to go to what it refers to. References → Update Table (F9) refreshes them after edits.</p>
        <div className="hf-actions">
          <button onClick={onClose}>Close</button>
          <button className="primary" disabled={!target} onClick={insert}>Insert</button>
        </div>
      </div>
    </div>
  );
}
