import { createRender, createSchema, insertEquation, latexToOmml, ommlToLatex } from "@prime/editor";
import type { Measurer } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { DOMSerializer } from "prosemirror-model";
import { NodeSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useRef, useState } from "react";

/** Pieces to click in, with where the cursor goes (the first "{}" or the end), in tabs. */
type Piece = [label: string, snippet: string, title: string];
const sym = (pairs: string): Piece[] => pairs.trim().split(/\s+/).map((p) => { const [c, name] = p.split("="); return [c, `\\${name} `, name]; });
const PALETTE: { tab: string; pieces: Piece[] }[] = [
  { tab: "Structures", pieces: [
    ["a/b", "\\frac{}{}", "Fraction"], ["x²", "^{}", "Power (superscript)"], ["xᵢ", "_{}", "Subscript"], ["xᵢ²", "_{}^{}", "Subscript and power"],
    ["√", "\\sqrt{}", "Square root"], ["ⁿ√", "\\sqrt[]{}", "nth root"], ["∑", "\\sum_{}^{} ", "Sum"], ["∏", "\\prod_{}^{} ", "Product"],
    ["∫", "\\int_{}^{} ", "Integral"], ["∬", "\\iint ", "Double integral"], ["∮", "\\oint ", "Contour integral"], ["lim", "\\lim_{} ", "Limit"],
    ["( )", "\\left( \\right)", "Brackets that grow"], ["[ ]", "\\left[ \\right]", "Square brackets that grow"], ["| |", "\\left| \\right|", "Absolute value"],
    ["(ⁿₖ)", "\\binom{}{}", "Binomial (n choose k)"], ["(⋮⋮)", "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}", "Matrix: & between cells, \\\\ between rows"],
    ["[⋮⋮]", "\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}", "Matrix in square brackets"], ["|⋮⋮|", "\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}", "Determinant"],
    ["{⋮", "\\begin{cases} x & x \\geq 0 \\\\ -x & x < 0 \\end{cases}", "Cases: a value & its condition on each row"],
    ["x̂", "\\hat{}", "Hat"], ["x̄", "\\bar{}", "Bar"], ["x⃗", "\\vec{}", "Vector"], ["ẋ", "\\dot{}", "Dot"], ["x̃", "\\tilde{}", "Tilde"],
    ["sin", "\\sin ", "Sine"], ["cos", "\\cos ", "Cosine"], ["tan", "\\tan ", "Tangent"], ["log", "\\log_{} ", "Logarithm to a base"], ["ln", "\\ln ", "Natural logarithm"],
    ["abc", "\\text{}", "Ordinary words"],
  ] },
  { tab: "Greek", pieces: sym(`α=alpha β=beta γ=gamma δ=delta ε=epsilon ζ=zeta η=eta θ=theta ι=iota κ=kappa λ=lambda μ=mu ν=nu ξ=xi π=pi ρ=rho σ=sigma
    τ=tau υ=upsilon φ=phi χ=chi ψ=psi ω=omega Γ=Gamma Δ=Delta Θ=Theta Λ=Lambda Ξ=Xi Π=Pi Σ=Sigma Φ=Phi Ψ=Psi Ω=Omega`) },
  { tab: "Operators", pieces: sym(`×=times ÷=div ⋅=cdot ±=pm ∓=mp ∘=circ ∂=partial ∇=nabla ∞=infty °=degree ′=prime ∠=angle ⋯=cdots …=ldots`) },
  { tab: "Relations", pieces: sym(`≤=leq ≥=geq ≠=neq ≈=approx ≡=equiv ∼=sim ∝=propto ⊥=perp ∥=parallel →=to ←=leftarrow ⇒=Rightarrow ⇐=Leftarrow ⇔=Leftrightarrow`) },
  { tab: "Sets & logic", pieces: sym(`∈=in ∉=notin ⊂=subset ⊆=subseteq ⊃=supset ∪=cup ∩=cap ∅=emptyset ∀=forall ∃=exists`) },
  { tab: "Examples", pieces: [
    ["Quadratic formula", "x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}", "Quadratic formula"],
    ["Pythagoras", "a^2 + b^2 = c^2", "Pythagoras' theorem"],
    ["Circle area", "A = \\pi r^2", "Area of a circle"],
    ["Derivative", "\\frac{d}{dx} x^n = n x^{n-1}", "Power rule"],
    ["Integral", "\\int_a^b f(x) dx = F(b) - F(a)", "Definite integral"],
    ["Sum", "\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}", "Sum of 1 to n"],
    ["Limit", "\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1", "A limit"],
    ["Mean", "\\bar{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_i", "Mean of a sample"],
    ["Compound interest", "A = P \\left(1 + \\frac{r}{n}\\right)^{nt}", "Compound interest"],
    ["Matrix", "A = \\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}", "A 2×2 matrix"],
    ["Piecewise", "|x| = \\begin{cases} x & x \\geq 0 \\\\ -x & x < 0 \\end{cases}", "Absolute value, piece by piece"],
    ["Binomial", "(a+b)^n = \\sum_{k=0}^{n} \\binom{n}{k} a^{n-k} b^k", "Binomial theorem"],
  ] },
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
  const [tab, setTab] = useState(0);
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
    // Several lines make a display equation even when "Display" is not ticked.
    const node = schema.nodes.inline_preserved.create({ preserved: { name: result.xml.startsWith("<m:oMathPara") ? "m:oMathPara" : "m:oMath", xml: result.xml } });
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
        <div className="eq-tabs" role="tablist" aria-label="Symbols and structures">
          {PALETTE.map((g, k) => <button key={g.tab} role="tab" aria-selected={k === tab} className={k === tab ? "on" : ""} onClick={() => setTab(k)}>{g.tab}</button>)}
        </div>
        <div className={`eq-palette ${PALETTE[tab].tab === "Examples" ? "examples" : ""}`} role="tabpanel">
          {PALETTE[tab].pieces.map(([label, snippet, title]) => <button key={title} title={`${title}\n${snippet.trim()}`} onClick={() => put(snippet)}>{label}</button>)}
        </div>
        <textarea ref={box} rows={4} value={src} onChange={(e) => setSrc(e.target.value)} autoFocus spellCheck={false}
          placeholder={"e.g.  x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\nPress Enter for another line."} aria-label="Equation (LaTeX)" />
        <div className={`eq-preview ${display ? "display" : ""}`} ref={preview} aria-label="Preview" />
        {result.error && <p className="eq-error">{result.error}</p>}
        <label className="choice"><input type="checkbox" checked={display} onChange={(e) => setDisplay(e.target.checked)} /> Display: on its own line, centred</label>
        <p className="hint">Each line you type is a line of the equation; words keep their spaces. Saved as a Word equation: Word opens and edits it as its own. Ctrl+Enter inserts.</p>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" disabled={!result.xml} onClick={insert}>{selected ? "Update" : "Insert"}</button>
        </div>
      </div>
    </div>
  );
}
