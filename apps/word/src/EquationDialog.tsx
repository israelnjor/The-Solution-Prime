import { createRender, createSchema, insertEquation, latexToOmml, ommlToLatex } from "@prime/editor";
import type { Measurer } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { DOMSerializer } from "prosemirror-model";
import { NodeSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useRef, useState } from "react";

/** Pieces to click in, with where the cursor goes (the first "{}" or the end). */
const PALETTE: [string, string, string][] = [
  ["a/b", "\\frac{}{}", "Fraction"], ["x²", "^{}", "Superscript"], ["xᵢ", "_{}", "Subscript"], ["√", "\\sqrt{}", "Square root"],
  ["ⁿ√", "\\sqrt[]{}", "nth root"], ["∑", "\\sum_{}^{} ", "Sum"], ["∫", "\\int_{}^{} ", "Integral"], ["lim", "\\lim_{} ", "Limit"],
  ["( )", "\\left( \\right)", "Brackets that grow"], ["x̂", "\\hat{}", "Hat"], ["x̄", "\\bar{}", "Bar"], ["sin", "\\sin ", "Sine"],
  ["α", "\\alpha ", "alpha"], ["β", "\\beta ", "beta"], ["θ", "\\theta ", "theta"], ["π", "\\pi ", "pi"], ["Δ", "\\Delta ", "Delta"], ["Ω", "\\Omega ", "Omega"],
  ["×", "\\times ", "times"], ["÷", "\\div ", "divide"], ["±", "\\pm ", "plus or minus"], ["≤", "\\leq ", "less or equal"], ["≥", "\\geq ", "greater or equal"],
  ["≠", "\\neq ", "not equal"], ["≈", "\\approx ", "approximately"], ["∞", "\\infty ", "infinity"], ["→", "\\to ", "arrow"], ["∂", "\\partial ", "partial"],
];

/**
 * Insert → Equation (Alt+=): typed in LaTeX notation with a live preview, saved as Word's own math.
 * A selected equation (double-click one) opens here to be changed.
 */
export function EquationDialog({ view, doc, measurer, onClose }: { view: EditorView; doc: PrimeDocument; measurer: Measurer; onClose(): void }) {
  const sel = view.state.selection;
  const selected = sel instanceof NodeSelection && sel.node.type.name === "inline_preserved" && /^m:oMath(Para)?$/.test(sel.node.attrs.preserved?.name ?? "") ? sel.node : undefined;
  const initial = selected ? ommlToLatex(selected.attrs.preserved.xml) : undefined;
  const [src, setSrc] = useState(initial?.latex ?? "");
  const [display, setDisplay] = useState(initial?.display ?? false);
  const box = useRef<HTMLTextAreaElement>(null);
  const preview = useRef<HTMLDivElement>(null);
  const schema = useMemo(() => createSchema(createRender(doc, measurer)), [doc, measurer]);
  const result = useMemo(() => {
    if (!src.trim()) return { xml: "", error: "" };
    try { return { xml: latexToOmml(src, display), error: "" }; } catch (e) { return { xml: "", error: (e as Error).message }; }
  }, [src, display]);

  // The preview, drawn exactly as the document will draw it.
  useEffect(() => {
    const p = preview.current;
    if (!p) return;
    p.replaceChildren();
    if (!result.xml) return;
    const node = schema.nodes.inline_preserved.create({ preserved: { name: display ? "m:oMathPara" : "m:oMath", xml: result.xml } });
    p.append(DOMSerializer.renderSpec(document, schema.spec.nodes.get("inline_preserved")!.toDOM!(node)).dom);
  }, [result, schema, display]);

  const put = (snippet: string) => {
    const t = box.current;
    if (!t) return;
    const at = t.selectionStart;
    const next = src.slice(0, at) + snippet + src.slice(t.selectionEnd);
    setSrc(next);
    const hole = snippet.indexOf("{}");
    const caret = at + (hole >= 0 ? hole + 1 : snippet.length);
    requestAnimationFrame(() => { t.focus(); t.setSelectionRange(caret, caret); });
  };
  const insert = () => {
    if (!result.xml) return;
    insertEquation(result.xml)(view.state, view.dispatch);
    onClose();
    view.focus();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter" && e.ctrlKey) insert(); }}>
      <div className="hf-dialog equation-dialog" role="dialog" aria-label="Equation">
        <h3>{selected ? "Edit equation" : "Equation"}</h3>
        {selected && !initial && <p className="eq-error">This equation uses parts the editor cannot show as text; what you type here replaces it.</p>}
        <div className="eq-palette" aria-label="Symbols and structures">
          {PALETTE.map(([label, snippet, title]) => <button key={title} title={title} onClick={() => put(snippet)}>{label}</button>)}
        </div>
        <textarea ref={box} rows={3} value={src} onChange={(e) => setSrc(e.target.value)} autoFocus spellCheck={false}
          placeholder={"e.g.  x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}"} aria-label="Equation (LaTeX)" />
        <div className={`eq-preview ${display ? "display" : ""}`} ref={preview} aria-label="Preview" />
        {result.error && <p className="eq-error">{result.error}</p>}
        <label className="choice"><input type="checkbox" checked={display} onChange={(e) => setDisplay(e.target.checked)} /> Display: on its own line, centred</label>
        <p className="hint">Saved as a Word equation: Word opens and edits it as its own. Ctrl+Enter inserts.</p>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" disabled={!result.xml} onClick={insert}>{selected ? "Update" : "Insert"}</button>
        </div>
      </div>
    </div>
  );
}
