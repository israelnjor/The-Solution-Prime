import type { EditorView } from "prosemirror-view";
import { useState } from "react";

const RECENT = "prime-word:symbols";
const SETS: [string, string][] = [
  ["Common", "©®™§¶†‡•…–—‘’“”«»‹›°±×÷≠≈≤≥∞√∑∏∫∂µπΩ∆‰′″№℃℉"],
  ["Currency", "₵₦$€£¥¢₹₽₩₱₺₴₫₪﷼₿"],
  ["Fractions", "½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞¹²³⁴⁵⁶⁷⁸⁹⁰₀₁₂₃₄₅₆₇₈₉"],
  ["Arrows", "←↑→↓↔↕⇐⇑⇒⇓⇔↖↗↘↙↩↪⟵⟶➔➜"],
  ["Maths", "∀∃∄∈∉∋∅∩∪⊂⊃⊆⊇⊕⊗⊥∠∴∵∝∧∨¬≡≅∼∓⌈⌉⌊⌋"],
  ["Greek", "αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"],
  ["Letters", "àáâãäåæçèéêëìíîïñòóôõöøœùúûüýÿßÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØŒÙÚÛÜÝƐɛƆɔŊŋ"],
  ["Shapes", "✓✔✗✘★☆●○◆◇■□▲△▼▽►◄♠♣♥♦☐☑☒✉☎✂✎❖"],
];

function recent(): string[] {
  try { return JSON.parse(localStorage.getItem(RECENT) ?? "[]"); } catch { return []; }
}

/** Insert → Symbol: common symbols by category, recently used first; inserted with the text's formatting. */
export function SymbolDialog({ view, onClose }: { view: EditorView; onClose(): void }) {
  const [set, setSet] = useState(0);
  const [used, setUsed] = useState(recent);
  const insert = (ch: string) => {
    view.dispatch(view.state.tr.insertText(ch).scrollIntoView());
    const next = [ch, ...used.filter((c) => c !== ch)].slice(0, 20);
    setUsed(next);
    try { localStorage.setItem(RECENT, JSON.stringify(next)); } catch { /* storage unavailable */ }
  };
  const chars = [...SETS[set][1]];
  const code = (ch: string) => `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}`;
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") { onClose(); view.focus(); } }}>
      <div className="hf-dialog symbol-dialog" role="dialog" aria-label="Symbol">
        <h3>Symbol</h3>
        {used.length > 0 && (
          <>
            <div className="symbol-label">Recently used</div>
            <div className="symbol-grid recent">{used.map((c) => <button key={c} onClick={() => insert(c)} title={code(c)}>{c}</button>)}</div>
          </>
        )}
        <div className="hf-tabs">{SETS.map(([name], i) => <button key={name} className={set === i ? "on" : ""} onClick={() => setSet(i)}>{name}</button>)}</div>
        <div className="symbol-grid">{chars.map((c) => <button key={c} onClick={() => insert(c)} title={code(c)}>{c}</button>)}</div>
        <p className="hint">Click a symbol to insert it at the cursor; the dialog stays open for more.</p>
        <div className="hf-actions"><button className="primary" onClick={() => { onClose(); view.focus(); }}>Close</button></div>
      </div>
    </div>
  );
}
