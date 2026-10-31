/**
 * Math typesetting: Office Math (OMML, m:oMath / m:oMathPara) -> a box of positioned glyphs and rules.
 *
 * A compact TeX-like box model: every element becomes a box with width, ascent and descent; glyph
 * and rule positions are relative to the box's left edge and baseline (dy > 0 is below the baseline).
 * Covers runs, accents, bars, scripts, fractions, delimiters, radicals, n-ary operators, functions,
 * limits, equation arrays and matrices. Unknown elements lay out their children in a row, so
 * nothing silently disappears.
 */
import { DOMParser } from "@xmldom/xmldom";
import type { FontKey, Measurer } from "./measure.js";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const MATH_FONT: FontKey = { family: "Cambria Math", bold: false, italic: false };

export interface MathGlyph {
  x: number;
  dy: number;
  text: string;
  size: number;
}

export interface MathRule {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
}

export interface MathBox {
  width: number;
  ascent: number;
  descent: number;
  /** Width of trailing spaces (accents and bars cover only the ink, not the spaces after it). */
  trail?: number;
  glyphs: MathGlyph[];
  rules: MathRule[];
}

// Approximate Cambria Math proportions (em): glyph extents, the math axis, script placement.
const ASCENT = 0.72;
const DESCENT = 0.22;
const AXIS = 0.25;
const SCRIPT = 0.7;
const RULE = 0.05;

const empty = (): MathBox => ({ width: 0, ascent: 0, descent: 0, glyphs: [], rules: [] });

function place(into: MathBox, box: MathBox, x: number, dy: number) {
  for (const g of box.glyphs) into.glyphs.push({ ...g, x: g.x + x, dy: g.dy + dy });
  for (const r of box.rules) into.rules.push({ ...r, x1: r.x1 + x, x2: r.x2 + x, y1: r.y1 + dy, y2: r.y2 + dy });
  into.ascent = Math.max(into.ascent, box.ascent - dy);
  into.descent = Math.max(into.descent, box.descent + dy);
}

function row(boxes: MathBox[], gap = 0): MathBox {
  const out = empty();
  let x = 0;
  boxes.forEach((b, i) => {
    place(out, b, x, 0);
    x += b.width + (i < boxes.length - 1 ? gap : 0);
  });
  out.width = x;
  // Trailing blank boxes (spaces), plus whatever blank tail the last inked box has.
  let trail = 0;
  for (let i = boxes.length - 1; i >= 0; i--) {
    const b = boxes[i];
    if (b.glyphs.length || b.rules.length) { trail += b.trail ?? 0; break; }
    trail += b.width;
  }
  out.trail = trail;
  return out;
}

/** Mathematical italic letters, as Word displays variables. */
function italic(ch: string): string {
  const c = ch.codePointAt(0)!;
  if (ch === "h") return "ℎ";
  if (c >= 65 && c <= 90) return String.fromCodePoint(0x1d434 + c - 65);
  if (c >= 97 && c <= 122) return String.fromCodePoint(0x1d44e + c - 97);
  return ch;
}

const RELATIONS = new Set([..."=<>≤≥≠≈≡∈∉⊂⊆→←⇒⇔"]);
const BINARY = new Set([..."+−-±×÷·∓∪∩∧∨⊕"]);
/** Combining accents drawn as a rule over the base (overline/bar). */
const OVERLINES = new Set(["̅", "̄", "¯"]);
const SPACING_ACCENT: Record<string, string> = {
  "̂": "^", "̃": "~", "̇": "˙", "̈": "¨", "́": "´", "̀": "`", "̌": "ˇ",
  "⃗": "→", "⃖": "←", "̆": "˘",
};

const kids = (e: Element) => {
  const out: Element[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
};
const child = (e: Element | undefined, name: string) => (e ? kids(e).find((c) => c.localName === name && c.namespaceURI === M) : undefined);
const mval = (e: Element | undefined) => e?.getAttributeNS(M, "val") || undefined;
const prop = (e: Element, pr: string, name: string) => mval(child(child(e, pr), name));

export class MathTypesetter {
  /** Size of the surrounding text; smaller sizes are script style (no operator spacing). */
  private base = 11;

  constructor(private m: Measurer) {}

  /** Lays out an m:oMath or m:oMathPara fragment. `size` is the surrounding text size in pt. */
  layout(xml: string, size: number): { box: MathBox; display: boolean } {
    this.base = size;
    const doc = new DOMParser({ errorHandler: { warning: () => undefined, error: () => undefined } })
      .parseFromString(`<root xmlns:m="${M}" xmlns:w="${W}">${xml}</root>`, "application/xml");
    const root = kids(doc.documentElement!)[0];
    const display = root?.localName === "oMathPara";
    if (!root) return { box: empty(), display };
    const maths = display ? kids(root).filter((c) => c.localName === "oMath") : [root];
    const lines = maths.map((e) => this.seq(kids(e), size, display));
    return { box: lines.length === 1 ? lines[0] : this.stack(lines, size, "center"), display };
  }

  private glyph(text: string, size: number): MathBox {
    const width = this.m.width(text, MATH_FONT) * size;
    return { width, ascent: ASCENT * size, descent: DESCENT * size, glyphs: [{ x: 0, dy: 0, text, size }], rules: [] };
  }

  private space(width: number): MathBox {
    return { ...empty(), width };
  }

  /** A sequence of math elements (the content of m:oMath, m:e, m:num, ...). */
  private seq(elements: Element[], size: number, display: boolean): MathBox {
    const boxes: MathBox[] = [];
    for (const e of elements) {
      if (e.namespaceURI === M && e.localName.endsWith("Pr")) continue;
      const b = this.element(e, size, display, boxes.length > 0);
      if (b) boxes.push(b);
    }
    return row(boxes);
  }

  private part(e: Element, name: string, size: number, display: boolean): MathBox {
    const c = child(e, name);
    return c ? this.seq(kids(c), size, display) : empty();
  }

  private element(e: Element, size: number, display: boolean, afterOperand = false): MathBox | undefined {
    if (e.namespaceURI === W) {
      // Ordinary runs inside math (w:r) render as upright text.
      if (e.localName === "r") return this.text(e.textContent ?? "", size, true);
      return undefined;
    }
    if (e.namespaceURI !== M) return undefined;
    switch (e.localName) {
      case "r": {
        const sz = /<w:sz w:val="(\d+)"/.exec(String(e))?.[1];
        const plain = prop(e, "rPr", "sty") === "p" || !!child(child(e, "rPr"), "nor");
        const text = kids(e).filter((c) => c.localName === "t").map((t) => t.textContent ?? "").join("");
        return this.text(text, sz ? Number(sz) / 2 : size, plain, afterOperand);
      }
      case "acc": return this.accent(e, size, display);
      case "bar": {
        const base = this.part(e, "e", size, display);
        const top = (prop(e, "barPr", "pos") ?? "bot") === "top";
        return this.overline(base, size, top);
      }
      case "sSup": case "sSub": case "sSubSup": case "sPre": return this.scripts(e, size, display);
      case "f": return this.fraction(e, size, display);
      case "d": return this.delimited(e, size, display);
      case "rad": return this.radical(e, size, display);
      case "nary": return this.nary(e, size, display);
      case "func": return row([this.part(e, "fName", size, display), this.space(size * 0.17), this.part(e, "e", size, display)]);
      case "limLow": case "limUpp": {
        const base = this.part(e, "e", size, display);
        const lim = this.part(e, "lim", size * SCRIPT, display);
        return this.limits(base, e.localName === "limUpp" ? lim : undefined, e.localName === "limLow" ? lim : undefined, size);
      }
      case "eqArr": return this.stack(kids(e).filter((c) => c.localName === "e").map((c) => this.seq(kids(c), size, display)), size, "center");
      case "m": return this.matrix(e, size, display);
      default: return this.seq(kids(e), size, display); // box, borderBox, groupChr, phant, sGroup...: their content
    }
  }

  private text(raw: string, size: number, plain: boolean, afterOperand = false): MathBox {
    const boxes: MathBox[] = [];
    const text = raw.replace(/ {2,}/g, " ");
    // Scripts and limits are set tight, as in Word (and TeX script style).
    const script = size < this.base * 0.85;
    let operand = afterOperand; // a binary operator needs something on its left; otherwise it is unary ("−b")
    for (const raw of text) {
      if (raw === " ") { boxes.push(this.space(size * 0.22)); continue; }
      const ch = raw === "-" ? "−" : raw;
      const g = this.glyph(plain ? ch : italic(ch), size);
      if (script) boxes.push(g);
      else if (RELATIONS.has(ch)) boxes.push(this.space(size * 0.28), g, this.space(size * 0.28));
      else if (BINARY.has(ch) && operand) boxes.push(this.space(size * 0.22), g, this.space(size * 0.22));
      else if (ch === ",") boxes.push(g, this.space(size * 0.17));
      else boxes.push(g);
      operand = !RELATIONS.has(ch) && !BINARY.has(ch) && !"([{,".includes(ch);
    }
    return row(boxes);
  }

  private overline(base: MathBox, size: number, top = true): MathBox {
    const out = row([base]);
    const gap = 0.08 * size;
    const y = top ? -(base.ascent + gap) : base.descent + gap;
    out.rules.push({ x1: 0, x2: base.width - (base.trail ?? 0), y1: y, y2: y, width: RULE * size });
    if (top) out.ascent = base.ascent + gap + RULE * size;
    else out.descent = base.descent + gap + RULE * size;
    return out;
  }

  private accent(e: Element, size: number, display: boolean): MathBox {
    const base = this.part(e, "e", size, display);
    const chr = prop(e, "accPr", "chr") ?? "̂";
    if (OVERLINES.has(chr)) return this.overline(base, size);
    const mark = this.glyph(SPACING_ACCENT[chr] ?? chr, size);
    const out = row([base]);
    place(out, mark, (base.width - (base.trail ?? 0) - mark.width) / 2, -(base.ascent - 0.45 * size));
    out.ascent = Math.max(out.ascent, base.ascent + 0.25 * size);
    return out;
  }

  private scripts(e: Element, size: number, display: boolean): MathBox {
    const base = this.part(e, "e", size, display);
    const sup = child(e, "sup") ? this.part(e, "sup", size * SCRIPT, display) : undefined;
    const sub = child(e, "sub") ? this.part(e, "sub", size * SCRIPT, display) : undefined;
    const up = Math.max(0.42 * size, base.ascent - 0.35 * size);
    // Word drops inline subscripts ~0.15em, keeping them within the text's own descent.
    const down = 0.15 * size;
    const scriptBox = empty();
    if (sup) place(scriptBox, sup, 0, -up);
    if (sub) place(scriptBox, sub, 0, down);
    scriptBox.width = Math.max(sup?.width ?? 0, sub?.width ?? 0);
    return e.localName === "sPre" ? row([scriptBox, base], size * 0.05) : row([base, scriptBox], size * 0.05);
  }

  private fraction(e: Element, size: number, display: boolean): MathBox {
    const type = prop(e, "fPr", "type") ?? "bar";
    const inner = display ? size : size * 0.8;
    const num = this.part(e, "num", inner, display);
    const den = this.part(e, "den", inner, display);
    if (type === "lin" || type === "skw") return row([num, this.glyph("/", size), den]);
    const width = Math.max(num.width, den.width) + 0.2 * size;
    const out = empty();
    const axis = -AXIS * size;
    const gap = 0.12 * size;
    place(out, num, (width - num.width) / 2, axis - gap - num.descent);
    place(out, den, (width - den.width) / 2, axis + gap + den.ascent);
    if (type !== "noBar") out.rules.push({ x1: 0.05 * size, x2: width - 0.05 * size, y1: axis, y2: axis, width: RULE * size });
    out.width = width;
    return out;
  }

  /** Draws a delimiter character tall enough for `content`, centred on the math axis. */
  private fence(chr: string, content: MathBox, size: number): MathBox {
    if (!chr) return empty();
    const need = content.ascent + content.descent;
    const scale = Math.max(1, need / ((ASCENT + DESCENT) * size));
    // A glyph's vertical centre sits on the axis at scale 1; keep it there when the glyph is enlarged.
    const out = empty();
    place(out, this.glyph(chr, size * scale), 0, AXIS * size * (scale - 1));
    out.width = this.m.width(chr, MATH_FONT) * size * scale;
    return out;
  }

  private delimited(e: Element, size: number, display: boolean): MathBox {
    const beg = prop(e, "dPr", "begChr") ?? "(";
    const end = prop(e, "dPr", "endChr") ?? ")";
    const sep = prop(e, "dPr", "sepChr") ?? "|";
    const items = kids(e).filter((c) => c.localName === "e").map((c) => this.seq(kids(c), size, display));
    const inner: MathBox[] = [];
    items.forEach((b, i) => {
      if (i) inner.push(this.glyph(sep, size));
      inner.push(b);
    });
    const content = row(inner);
    return row([this.fence(beg, content, size), content, this.fence(end, content, size)], size * 0.03);
  }

  private radical(e: Element, size: number, display: boolean): MathBox {
    const base = this.part(e, "e", size, display);
    const hideDeg = prop(e, "radPr", "degHide") === "1" || prop(e, "radPr", "degHide") === "on";
    const deg = hideDeg ? undefined : this.part(e, "deg", size * 0.5, display);
    const sign = this.fence("√", base, size);
    const out = empty();
    let x = 0;
    if (deg && deg.width) {
      place(out, deg, 0, -0.45 * size);
      x = Math.max(0, deg.width - sign.width * 0.4);
    }
    place(out, sign, x, 0);
    x += sign.width;
    place(out, base, x, 0);
    const top = -(base.ascent + 0.1 * size);
    out.rules.push({ x1: x, x2: x + base.width, y1: top, y2: top, width: RULE * size });
    out.ascent = Math.max(out.ascent, -top + RULE * size);
    out.width = x + base.width;
    return out;
  }

  private limits(base: MathBox, upper: MathBox | undefined, lower: MathBox | undefined, size: number): MathBox {
    const width = Math.max(base.width, upper?.width ?? 0, lower?.width ?? 0);
    const out = empty();
    place(out, base, (width - base.width) / 2, 0);
    if (upper) place(out, upper, (width - upper.width) / 2, -(base.ascent + 0.1 * size + upper.descent));
    if (lower) place(out, lower, (width - lower.width) / 2, base.descent + 0.1 * size + lower.ascent);
    out.width = width;
    return out;
  }

  private nary(e: Element, size: number, display: boolean): MathBox {
    const chr = prop(e, "naryPr", "chr") ?? "∫";
    const integral = /[∫∬∭∮]/.test(chr);
    const loc = prop(e, "naryPr", "limLoc") ?? (integral ? "subSup" : "undOvr");
    const hidden = (name: string) => ["1", "on", "true"].includes(prop(e, "naryPr", name) ?? "");
    // Display integrals are tall (Word uses the large variant); sums grow less.
    const op = this.glyph(chr, display ? size * (integral ? 2 : 1.4) : size * 1.1);
    const sub = hidden("subHide") ? undefined : this.part(e, "sub", size * SCRIPT, display);
    const sup = hidden("supHide") ? undefined : this.part(e, "sup", size * SCRIPT, display);
    let operator: MathBox;
    if (loc === "undOvr") operator = this.limits(op, sup?.width ? sup : undefined, sub?.width ? sub : undefined, size);
    else {
      const scripts = empty();
      if (sup?.width) place(scripts, sup, 0, -(op.ascent - (integral ? 0.45 : 0.2) * size));
      if (sub?.width) place(scripts, sub, 0, display ? op.descent + 0.1 * size : 0.15 * size);
      scripts.width = Math.max(sup?.width ?? 0, sub?.width ?? 0);
      operator = row([op, scripts], size * 0.05);
    }
    return row([operator, this.space(size * 0.17), this.part(e, "e", size, display)]);
  }

  /** Rows stacked vertically, the whole stack centred on the math axis. */
  private stack(rows: MathBox[], size: number, align: "center" | "left"): MathBox {
    const width = Math.max(0, ...rows.map((r) => r.width));
    const out = empty();
    let y = 0;
    rows.forEach((r, i) => {
      if (i) y += rows[i - 1].descent + size * 0.3 + r.ascent;
      place(out, r, align === "center" ? (width - r.width) / 2 : 0, y);
    });
    const middle = (out.descent - out.ascent) / 2;
    const centred = empty();
    place(centred, out, 0, -AXIS * size - middle);
    centred.width = width;
    return centred;
  }

  private matrix(e: Element, size: number, display: boolean): MathBox {
    const rows = kids(e).filter((c) => c.localName === "mr")
      .map((r) => kids(r).filter((c) => c.localName === "e").map((c) => this.seq(kids(c), size, display)));
    const cols = Math.max(0, ...rows.map((r) => r.length));
    const widths = Array.from({ length: cols }, (_, j) => Math.max(0, ...rows.map((r) => r[j]?.width ?? 0)));
    const gap = size * 0.6;
    const lines = rows.map((r) => {
      const out = empty();
      let x = 0;
      r.forEach((cell, j) => {
        place(out, cell, x + (widths[j] - cell.width) / 2, 0);
        x += widths[j] + gap;
      });
      out.width = Math.max(0, x - gap);
      return out;
    });
    return this.stack(lines, size, "center");
  }
}
