import { setIndent, setLineSpacing, setParagraph, setParagraphSpacing, setTabStops } from "@prime/editor";
import { Resolver, type TabStop } from "@prime/layout";
import type { Paragraph, PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useState } from "react";

const CM = 72 / 2.54;
const round = (v: number, d = 2) => Math.round(v * 10 ** d) / 10 ** d;

type LineRule = "single" | "1.5" | "double" | "multiple" | "atLeast" | "exact";

const ALIGN: [string, string][] = [["left", "Left"], ["center", "Centred"], ["right", "Right"], ["both", "Justified"]];
const LEADER: [string, string][] = [["", "None"], ["dot", "....."], ["hyphen", "-----"], ["underscore", "_____"]];

/**
 * Word's Paragraph dialog: alignment, indents (cm), spacing before/after (pt), line spacing, line and
 * page breaks, and tab stops.
 */
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
    align: rp.align === "both" || rp.align === "distribute" ? "both" : rp.align === "end" ? "right" : rp.align === "start" ? "left" : rp.align,
  });
  const own = node.attrs.props ?? {};
  const [breaks, setBreaks] = useState({
    widowControl: own.widowControl ?? true, keepNext: !!own.keepNext, keepLines: !!own.keepLines, pageBreakBefore: !!own.pageBreakBefore,
  });
  const [tabs, setTabs] = useState<TabStop[]>(rp.tabs);
  const [tab, setTab] = useState({ pos: "", align: "left" as TabStop["align"], leader: "" });
  const addTab = () => {
    const pos = Number(tab.pos.replace(",", ".")) * CM;
    if (!(pos > 0)) return;
    setTabs((t) => [...t.filter((x) => Math.abs(x.pos - pos) > 0.5), { pos, align: tab.align, ...(tab.leader ? { leader: tab.leader as TabStop["leader"] } : {}) }].sort((a, b) => a.pos - b.pos));
    setTab((t) => ({ ...t, pos: "" }));
  };
  const set = (k: keyof typeof f) => (e: { target: { value: string } }) =>
    setF((s) => ({ ...s, [k]: k === "special" || k === "rule" || k === "align" ? e.target.value : Number(e.target.value) }));

  const apply = () => {
    const run = (cmd: (s: EditorView["state"], d: EditorView["dispatch"]) => boolean) => cmd(view.state, view.dispatch);
    run(setIndent({ left: f.left * CM, right: f.right * CM, firstLine: f.special === "none" ? 0 : (f.special === "first" ? 1 : -1) * f.by * CM }));
    run(setParagraphSpacing({ before: f.before, after: f.after }));
    if (f.rule === "single") run(setLineSpacing("multiple", 1));
    else if (f.rule === "1.5") run(setLineSpacing("multiple", 1.5));
    else if (f.rule === "double") run(setLineSpacing("multiple", 2));
    else if (f.rule === "multiple") run(setLineSpacing("multiple", f.value));
    else run(setLineSpacing(f.rule, f.value));
    if (f.align !== (rp.align === "both" || rp.align === "distribute" ? "both" : rp.align)) run(setParagraph("align", f.align === "left" ? undefined : f.align));
    for (const k of ["widowControl", "keepNext", "keepLines", "pageBreakBefore"] as const) {
      const was = k === "widowControl" ? own.widowControl ?? true : !!own[k];
      if (breaks[k] !== was) run(setParagraph(k, k === "widowControl" ? breaks[k] : breaks[k] || undefined));
    }
    const same = tabs.length === rp.tabs.length && tabs.every((t, i) => Math.abs(t.pos - rp.tabs[i].pos) < 0.5 && t.align === rp.tabs[i].align && t.leader === rp.tabs[i].leader);
    if (!same) run(setTabStops(() => doc, tabs));
    onClose();
    view.focus();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter") apply(); }}>
      <div className="hf-dialog para-dialog" role="dialog" aria-label="Paragraph">
        <h3>Paragraph</h3>
        <fieldset>
          <legend>General</legend>
          <label>Alignment <select value={f.align} onChange={set("align")}>
            {ALIGN.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select></label>
        </fieldset>
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
        <fieldset className="para-breaks">
          <legend>Line and page breaks</legend>
          {([["widowControl", "Widow/Orphan control"], ["keepNext", "Keep with next"], ["keepLines", "Keep lines together"], ["pageBreakBefore", "Page break before"]] as const).map(([k, l]) => (
            <label key={k} className="choice"><input type="checkbox" checked={breaks[k]} onChange={(e) => setBreaks((b) => ({ ...b, [k]: e.target.checked }))} /> {l}</label>
          ))}
        </fieldset>
        <fieldset className="para-tabs">
          <legend>Tab stops</legend>
          <div className="tab-add">
            <input aria-label="Tab stop position (cm)" placeholder="Position (cm)" value={tab.pos} onChange={(e) => setTab((t) => ({ ...t, pos: e.target.value }))}
              onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); addTab(); } }} />
            <select aria-label="Alignment" value={tab.align} onChange={(e) => setTab((t) => ({ ...t, align: e.target.value as TabStop["align"] }))}>
              {(["left", "center", "right", "decimal", "bar"] as const).map((a) => <option key={a} value={a}>{a === "center" ? "Centre" : a[0].toUpperCase() + a.slice(1)}</option>)}
            </select>
            <select aria-label="Leader" value={tab.leader} onChange={(e) => setTab((t) => ({ ...t, leader: e.target.value }))}>
              {LEADER.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
            <button type="button" onClick={addTab}>Set</button>
            <button type="button" disabled={!tabs.length} onClick={() => setTabs([])}>Clear all</button>
          </div>
          {tabs.length > 0 && (
            <ul className="tab-list">
              {tabs.map((t) => (
                <li key={t.pos}>{round(t.pos / CM)} cm · {t.align}{t.leader ? ` · ${t.leader}` : ""}
                  <button type="button" className="link" onClick={() => setTabs((all) => all.filter((x) => x !== t))}>Clear</button></li>
              ))}
            </ul>
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
