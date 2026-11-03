/**
 * Equations: typed in the LaTeX-style notation students know (\frac{a}{b}, x^2, \sqrt{x},
 * \sum_{i=1}^{n}, \alpha…) and stored as Word's own math (OMML), so Word edits them natively.
 * An equation from a document turns back into the same notation for editing, when it uses only
 * what this notation covers.
 */
import { DOMParser } from "@xmldom/xmldom";
import type { Command } from "prosemirror-state";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

const SYMBOLS: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", iota: "ι",
  kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", rho: "ρ", sigma: "σ", tau: "τ", upsilon: "υ", phi: "φ",
  varphi: "φ", chi: "χ", psi: "ψ", omega: "ω", Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
  Sigma: "Σ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  times: "×", cdot: "⋅", div: "÷", pm: "±", mp: "∓", leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠",
  approx: "≈", equiv: "≡", sim: "∼", propto: "∝", infty: "∞", to: "→", rightarrow: "→", leftarrow: "←",
  Rightarrow: "⇒", Leftarrow: "⇐", Leftrightarrow: "⇔", partial: "∂", nabla: "∇", in: "∈", notin: "∉",
  subset: "⊂", subseteq: "⊆", supset: "⊃", cup: "∪", cap: "∩", forall: "∀", exists: "∃", emptyset: "∅",
  angle: "∠", perp: "⊥", parallel: "∥", degree: "°", cdots: "⋯", ldots: "…", prime: "′", circ: "∘",
};
const NARY: Record<string, string> = { sum: "∑", prod: "∏", int: "∫", iint: "∬", oint: "∮", bigcup: "⋃", bigcap: "⋂" };
const FUNCTIONS = ["sin", "cos", "tan", "sec", "csc", "cot", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "log", "ln", "exp", "max", "min", "det", "gcd"];
const ACCENTS: Record<string, string> = { hat: "̂", tilde: "̃", dot: "̇", ddot: "̈", vec: "⃗" };
const DELIMS: Record<string, string> = { "(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}", "|": "|", ".": "" };

// --- LaTeX → OMML ---------------------------------------------------------------------------------

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const run = (text: string, upright = false) => `<m:r>${upright ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ""}<m:t>${esc(text)}</m:t></m:r>`;
const val = (tag: string, v: string) => `<m:${tag} m:val="${esc(v)}"/>`;

class Parser {
  private i = 0;
  constructor(private s: string) {}

  private peek() { return this.s[this.i]; }
  private skipSpace() { while (this.s[this.i] === " ") this.i++; }
  private command(): string {
    const m = /^\\([A-Za-z]+|.)/.exec(this.s.slice(this.i));
    if (!m) throw new Error("expected a command");
    this.i += m[0].length;
    return m[1];
  }
  private expect(ch: string) {
    this.skipSpace();
    if (this.s[this.i] !== ch) throw new Error(`expected “${ch}” at ${this.i + 1}`);
    this.i++;
  }

  /**
   * Items up to the end, or up to a closing brace / \right (not consumed). An operator's operand
   * (`toRelation`) also stops at a relation (=, <, ≤, →…), as Word's own linear format does.
   */
  expr(toRelation = false): string {
    let out = "";
    let text = "";
    const flush = () => { if (text) { out += run(text); text = ""; } };
    for (;;) {
      this.skipSpace();
      const c = this.peek();
      if (c === undefined || c === "}" || this.s.startsWith("\\right", this.i) || c === "&") break;
      if (toRelation && (/[=<>]/.test(c) || /^\\(leq|le|geq|ge|neq|ne|approx|equiv|to|rightarrow|Rightarrow|sim|propto)(?![A-Za-z])/.test(this.s.slice(this.i)))) break;
      // A plain character that no script follows can join the run before it.
      if (/[A-Za-z0-9.,+\-=<>!'|:;/*]/.test(c) && !/[\^_]/.test(this.s[this.i + 1] ?? "")) {
        if (c === "*") text += "∗";
        else text += c === "-" ? "−" : c;
        this.i++;
        continue;
      }
      flush();
      out += this.item();
    }
    flush();
    return out;
  }

  /** One item with any ^ and _ after it. */
  private item(): string {
    const base = this.atom();
    if (base.nary) return this.nary(base.nary);
    if (base.lim) return this.limit(base.lim);
    let sub: string | undefined;
    let sup: string | undefined;
    for (;;) {
      this.skipSpace();
      if (this.peek() === "^" && sup === undefined) { this.i++; sup = this.arg(); continue; }
      if (this.peek() === "_" && sub === undefined) { this.i++; sub = this.arg(); continue; }
      break;
    }
    if (sub !== undefined && sup !== undefined) return `<m:sSubSup><m:e>${base.xml}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
    if (sup !== undefined) return `<m:sSup><m:e>${base.xml}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
    if (sub !== undefined) return `<m:sSub><m:e>${base.xml}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
    return base.xml;
  }

  /** A group, a single character, or a command with its arguments. */
  private arg(): string {
    this.skipSpace();
    if (this.peek() === "{") { this.i++; const e = this.expr(); this.expect("}"); return e; }
    if (this.peek() === "\\") return this.atom().xml;
    const c = this.s[this.i++];
    if (c === undefined) throw new Error("something is missing at the end");
    return run(c === "-" ? "−" : c);
  }

  private atom(): { xml: string; nary?: string; lim?: string } {
    this.skipSpace();
    const c = this.peek();
    if (c === "{") { this.i++; const e = this.expr(); this.expect("}"); return { xml: e }; }
    if (c !== "\\") { this.i++; return { xml: run(c === "-" ? "−" : c) }; }
    const name = this.command();
    if (name === "frac") return { xml: `<m:f><m:num>${this.arg()}</m:num><m:den>${this.arg()}</m:den></m:f>` };
    if (name === "sqrt") {
      this.skipSpace();
      let deg = "";
      if (this.peek() === "[") { this.i++; const end = this.s.indexOf("]", this.i); if (end < 0) throw new Error("expected “]”"); deg = new Parser(this.s.slice(this.i, end)).expr(); this.i = end + 1; }
      return { xml: `<m:rad>${deg ? "" : `<m:radPr>${val("degHide", "1")}</m:radPr>`}<m:deg>${deg}</m:deg><m:e>${this.arg()}</m:e></m:rad>` };
    }
    if (name in NARY) return { xml: "", nary: NARY[name] };
    if (name === "lim") return { xml: "", lim: "lim" };
    if (name === "left") {
      const open = this.delim();
      const e = this.expr();
      if (!this.s.startsWith("\\right", this.i)) throw new Error("\\left without \\right");
      this.i += 6;
      const close = this.delim();
      const pr = (open === "(" ? "" : `<m:begChr m:val="${esc(open)}"/>`) + (close === ")" ? "" : `<m:endChr m:val="${esc(close)}"/>`);
      return { xml: `<m:d>${pr ? `<m:dPr>${pr}</m:dPr>` : ""}<m:e>${e}</m:e></m:d>` };
    }
    if (name === "text" || name === "mathrm") {
      this.expect("{");
      const end = this.s.indexOf("}", this.i);
      if (end < 0) throw new Error("expected “}”");
      const text = this.s.slice(this.i, end);
      this.i = end + 1;
      return { xml: run(text, true) };
    }
    if (name in ACCENTS) return { xml: `<m:acc><m:accPr>${val("chr", ACCENTS[name])}</m:accPr><m:e>${this.arg()}</m:e></m:acc>` };
    if (name === "bar" || name === "overline") return { xml: `<m:bar><m:barPr>${val("pos", "top")}</m:barPr><m:e>${this.arg()}</m:e></m:bar>` };
    if (FUNCTIONS.includes(name)) {
      return { xml: `<m:func><m:fName>${run(name, true)}</m:fName><m:e>${this.funcArg()}</m:e></m:func>` };
    }
    if (name in SYMBOLS) return { xml: run(SYMBOLS[name]) };
    if (name === "," || name === ";" || name === "quad") return { xml: run(name === "quad" ? " " : " ") };
    if (name === "{" || name === "}") return { xml: run(name) };
    throw new Error(`unknown command \\${name}`);
  }

  /** A function's argument: the next item (sin x, sin(x), sin^2 x is written \sin^2 x as in LaTeX). */
  private funcArg(): string {
    this.skipSpace();
    return this.peek() === undefined || this.peek() === "}" ? "" : this.item();
  }

  private delim(): string {
    this.skipSpace();
    for (const d of Object.keys(DELIMS)) if (this.s.startsWith(d, this.i)) { this.i += d.length; return DELIMS[d]; }
    throw new Error("expected a bracket after \\left or \\right");
  }

  /** \sum_{i=1}^{n} x: limits, then the next item as what is summed. */
  private nary(chr: string): string {
    let sub = "";
    let sup = "";
    for (;;) {
      this.skipSpace();
      if (this.peek() === "_") { this.i++; sub = this.arg(); continue; }
      if (this.peek() === "^") { this.i++; sup = this.arg(); continue; }
      break;
    }
    this.skipSpace();
    const e = this.expr(true);
    const hide = (sub ? "" : val("subHide", "1")) + (sup ? "" : val("supHide", "1"));
    return `<m:nary><m:naryPr>${val("chr", chr)}${chr === "∑" || chr === "∏" ? val("limLoc", "undOvr") : ""}${hide}</m:naryPr><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${e}</m:e></m:nary>`;
  }

  private limit(name: string): string {
    this.skipSpace();
    let lim = "";
    if (this.peek() === "_") { this.i++; lim = this.arg(); }
    return `<m:limLow><m:e>${run(name, true)}</m:e><m:lim>${lim}</m:lim></m:limLow>`;
  }

  done() { this.skipSpace(); if (this.i < this.s.length) throw new Error(`unexpected “${this.s[this.i]}” at ${this.i + 1}`); }
}

/** OMML for an equation typed as LaTeX; display equations go in an m:oMathPara (their own line, centred). Throws with a short message on a mistake. */
export function latexToOmml(src: string, display = false): string {
  const p = new Parser(src.trim());
  const body = p.expr();
  p.done();
  const math = `<m:oMath xmlns:m="${M}">${body}</m:oMath>`;
  return display ? `<m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${math.replace(` xmlns:m="${M}"`, "")}</m:oMathPara>` : math;
}

// --- OMML → LaTeX ---------------------------------------------------------------------------------

const BACK: Record<string, string> = Object.fromEntries(Object.entries(SYMBOLS).filter(([k]) => !["le", "ge", "ne", "varepsilon", "varphi", "rightarrow"].includes(k)).map(([k, v]) => [v, `\\${k} `]));
const NARY_BACK: Record<string, string> = Object.fromEntries(Object.entries(NARY).map(([k, v]) => [v, k]));
const ACCENT_BACK: Record<string, string> = Object.fromEntries(Object.entries(ACCENTS).map(([k, v]) => [v, k]));

class Unsupported extends Error {}

const kids = (e: Element) => Array.from(e.childNodes).filter((n): n is Element => n.nodeType === 1);
const kid = (e: Element, name: string) => kids(e).find((k) => k.localName === name);
const prop = (e: Element, pr: string, name: string) => {
  const p = kid(e, pr);
  const x = p && kid(p, name);
  return x?.getAttributeNS(M, "val") ?? x?.getAttribute("m:val") ?? undefined;
};
const group = (s: string) => (s.replace(/\s+$/, "").length === 1 ? s.trim() : `{${s.trim()}}`);
/** A command's argument: always braced (\hat{y}, not \haty). */
const brace = (s: string) => `{${s.trim()}}`;

function toLatex(e: Element | undefined): string {
  if (!e) return "";
  return kids(e).map(node).join("");
}

function node(e: Element): string {
  switch (e.localName) {
    case "r": {
      const text = kids(e).filter((k) => k.localName === "t").map((t) => t.textContent ?? "").join("");
      const upright = prop(e, "rPr", "sty") === "p";
      if (upright && /[A-Za-z]{2,}/.test(text) && !FUNCTIONS.includes(text)) return `\\text{${text}}`;
      if (upright && FUNCTIONS.includes(text)) return `\\${text} `;
      return [...text].map((c) => (c === "−" ? "-" : c === "∗" ? "*" : BACK[c] ?? c)).join("");
    }
    case "f": return `\\frac${brace(toLatex(kid(e, "num")))}${brace(toLatex(kid(e, "den")))}`;
    case "sSup": return `${toLatex(kid(e, "e"))}^${group(toLatex(kid(e, "sup")))}`;
    case "sSub": return `${toLatex(kid(e, "e"))}_${group(toLatex(kid(e, "sub")))}`;
    case "sSubSup": return `${toLatex(kid(e, "e"))}_${group(toLatex(kid(e, "sub")))}^${group(toLatex(kid(e, "sup")))}`;
    case "rad": {
      const deg = toLatex(kid(e, "deg"));
      return `\\sqrt${deg && prop(e, "radPr", "degHide") !== "1" ? `[${deg}]` : ""}${brace(toLatex(kid(e, "e")))}`;
    }
    case "nary": {
      const chr = prop(e, "naryPr", "chr") ?? "∫";
      const name = NARY_BACK[chr];
      if (!name) throw new Unsupported(chr);
      const sub = toLatex(kid(e, "sub"));
      const sup = toLatex(kid(e, "sup"));
      return `\\${name}${sub ? `_${group(sub)}` : ""}${sup ? `^${group(sup)}` : ""} ${toLatex(kid(e, "e"))}`;
    }
    case "d": {
      const open = prop(e, "dPr", "begChr") ?? "(";
      const close = prop(e, "dPr", "endChr") ?? ")";
      const tex = (c: string) => (c === "{" || c === "}" ? `\\${c}` : c || ".");
      // Several m:e (separated by a comma in Word's view) become comma-separated.
      const inner = kids(e).filter((k) => k.localName === "e").map(toLatex).join(prop(e, "dPr", "sepChr") ?? ",");
      return `\\left${tex(open)}${inner}\\right${tex(close)}`;
    }
    case "limLow": {
      const base = toLatex(kid(e, "e")).trim();
      if (base !== "\\text{lim}" && base !== "lim") throw new Unsupported("limLow");
      return `\\lim_${group(toLatex(kid(e, "lim")))} `;
    }
    case "func": return `${toLatex(kid(e, "fName")).trim()} ${toLatex(kid(e, "e"))}`;
    case "acc": {
      const name = ACCENT_BACK[prop(e, "accPr", "chr") ?? "̂"];
      if (!name) throw new Unsupported("accent");
      return `\\${name}${brace(toLatex(kid(e, "e")))}`;
    }
    case "bar": return `\\bar${brace(toLatex(kid(e, "e")))}`;
    case "rPr": case "fPr": case "sSupPr": case "sSubPr": case "sSubSupPr": case "radPr": case "naryPr": case "dPr": case "funcPr": case "accPr": case "barPr": case "limLowPr": case "ctrlPr": case "oMathParaPr":
      return "";
    case "oMath": return toLatex(e);
    default: throw new Unsupported(e.localName);
  }
}

/** The equation (an m:oMath or m:oMathPara) in LaTeX notation, or undefined if it uses more than the notation covers. */
export function ommlToLatex(xml: string): { latex: string; display: boolean } | undefined {
  try {
    const withNs = xml.includes("xmlns:m=") ? xml : xml.replace(/^<(m:oMath(?:Para)?)/, `<$1 xmlns:m="${M}"`);
    const root = new DOMParser().parseFromString(withNs, "application/xml").documentElement;
    if (!root) return undefined;
    const display = root.localName === "oMathPara";
    const maths = display ? kids(root).filter((k) => k.localName === "oMath") : [root];
    return { latex: maths.map(toLatex).join(" ").replace(/\s+/g, " ").trim(), display };
  } catch {
    return undefined;
  }
}

// --- inserting -------------------------------------------------------------------------------------

/**
 * An equation at the cursor (or in place of the selected one): inline, or display — in a paragraph
 * of its own, centred, as Word puts display equations.
 */
export function insertEquation(xml: string): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const display = xml.startsWith("<m:oMathPara");
    // Word's documents declare the namespace on the root; the element itself does not repeat it.
    const clean = xml.replace(` xmlns:m="${M}"`, "");
    const node = state.schema.nodes.inline_preserved.create({ preserved: { name: display ? "m:oMathPara" : "m:oMath", xml: clean } });
    const { $from, from, to } = state.selection;
    const tr = state.tr;
    if (display && $from.parent.content.size > 0 && !(state.selection.from === $from.start() && to === $from.end())) {
      // In the middle of text: the equation gets its own paragraph after this one.
      const para = state.schema.nodes.paragraph.create({ props: {} }, node);
      tr.insert($from.after(), para);
    } else tr.replaceWith(from, to, node);
    dispatch(tr.scrollIntoView());
    return true;
  };
}
