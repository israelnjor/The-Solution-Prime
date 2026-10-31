import { setIndent, setLineSpacing, setParagraphSpacing } from "@prime/editor";
import { Resolver } from "@prime/layout";
import type { Paragraph, PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useState } from "react";

const CM = 72 / 2.54;
const round = (v: number, d = 2) => Math.round(v * 10 ** d) / 10 ** d;

type LineRule = "single" | "1.5" | "double" | "multiple" | "atLeast" | "exact";

/** Word's Paragraph dialog: indents (cm), spacing before/after (pt) and line spacing. */
export function ParagraphDialog({ view, doc, onClose }: { view: EditorView; doc: PrimeDocument; onClose(): void }) {
  // Start from the paragraph's effective formatting (style + direct), as Word shows it.
  const node = view.state.selection.$from.parent;
  const rp = new Resolver(doc).paragraph({ type: "paragraph", props: node.attrs.props ?? {}, content: [] } as Paragraph, undefined, false);
  const initialRule: LineRule = rp.lineRule === "exact" ? "exact" : rp.lineRule === "atLeast" ? "atLeast"
    : rp.line === 240 ? "single" : rp.line === 360 ? "1.5" : rp.line === 480 ? "double" : "multiple";
  const [f, setF] = useState({
    left: round(rp.left / CM), right: round(rp.right / CM),
    special: rp.firstLine > 0 ? "first" : rp.firstLine < 0 ? "hanging" : "none",
    by: round(Math.abs(rp.firstLine) / CM),
    before: round(rp.spaceBefore, 1), after: round(rp.spaceAfter, 1),
    rule: initialRule,
    value: rp.lineRule === "auto" ? round(rp.line / 240) : round(rp.line / 20, 1),
  });
  const set = (k: keyof typeof f) => (e: { target: { value: string } }) =>
    setF((s) => ({ ...s, [k]: k === "special" || k === "rule" ? e.target.value : Number(e.target.value) }));

  const apply = () => {
    const run = (cmd: (s: EditorView["state"], d: EditorView["dispatch"]) => boolean) => cmd(view.state, view.dispatch);
    run(setIndent({ left: f.left * CM, right: f.right * CM, firstLine: f.special === "none" ? 0 : (f.special === "first" ? 1 : -1) * f.by * CM }));
    run(setParagraphSpacing({ before: f.before, after: f.after }));
    if (f.rule === "single") run(setLineSpacing("multiple", 1));
    else if (f.rule === "1.5") run(setLineSpacing("multiple", 1.5));
    else if (f.rule === "double") run(setLineSpacing("multiple", 2));
    else if (f.rule === "multiple") run(setLineSpacing("multiple", f.value));
    else run(setLineSpacing(f.rule, f.value));
    onClose();
    view.focus();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter") apply(); }}>
      <div className="hf-dialog para-dialog" role="dialog" aria-label="Paragraph">
        <h3>Paragraph</h3>
        <fieldset>
          <legend>Indentation</legend>
          <label>Left <input type="number" step={0.1} value={f.left} onChange={set("left")} /> cm</label>
          <label>Right <input type="number" step={0.1} value={f.right} onChange={set("right")} /> cm</label>
          <label>Special <select value={f.special} onChange={set("special")}>
            <option value="none">(none)</option><option value="first">First line</option><option value="hanging">Hanging</option>
          </select></label>
          <label>By <input type="number" step={0.1} min={0} value={f.by} disabled={f.special === "none"} onChange={set("by")} /> cm</label>
        </fieldset>
        <fieldset>
          <legend>Spacing</legend>
          <label>Before <input type="number" step={1} min={0} value={f.before} onChange={set("before")} /> pt</label>
          <label>After <input type="number" step={1} min={0} value={f.after} onChange={set("after")} /> pt</label>
          <label>Line spacing <select value={f.rule} onChange={set("rule")}>
            <option value="single">Single</option><option value="1.5">1.5 lines</option><option value="double">Double</option>
            <option value="multiple">Multiple</option><option value="atLeast">At least</option><option value="exact">Exactly</option>
          </select></label>
          {(f.rule === "multiple" || f.rule === "atLeast" || f.rule === "exact") && (
            <label>{f.rule === "multiple" ? "At" : "At"} <input type="number" step={f.rule === "multiple" ? 0.05 : 1} min={0} value={f.value} onChange={set("value")} /> {f.rule === "multiple" ? "lines" : "pt"}</label>
          )}
        </fieldset>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={apply}>OK</button>
        </div>
      </div>
    </div>
  );
}
