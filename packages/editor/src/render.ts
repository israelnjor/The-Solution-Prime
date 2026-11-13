/**
 * How the document looks in the editor (browser only).
 *
 * Paragraph styling is computed in one document-order pass by the layout engine's style resolver
 * (so list numbering, table styles and spacing match the layout engine), and applied as node
 * decorations. Runs render their own direct formatting; everything else is inherited via CSS.
 * Equations are drawn as SVG by the math typesetter.
 */
import { type CellLook, cellLooks, chartSvg, chartTheme, MathTypesetter, type Measurer, parseChart, type ResolvedRun, Resolver } from "@prime/layout";
import { type FloatInfo, type Paragraph, type PrimeDocument, paragraphText, type RunProps, type Table, type TextInline } from "@prime/pdm";
import { DOMSerializer, type Node } from "prosemirror-model";
import { Plugin, PluginKey, type Transaction } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from "prosemirror-transform";
import { Decoration, DecorationSet } from "prosemirror-view";
import { inkSvg } from "./ink.js";
import type { Render } from "./schema.js";
import { timed } from "./paged.js";

const pt = (v: number) => `${Math.round(v * 100) / 100}pt`;
const json = (v: unknown) => JSON.stringify(v);
const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
  darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
  darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};
const SVG = "http://www.w3.org/2000/svg";

/** CSS for run formatting; only the properties that are set. */
export function runCss(r: Partial<ResolvedRun>): string {
  const css: string[] = [];
  if (r.font) css.push(`font-family:"${r.font}"`);
  if (r.size !== undefined) css.push(`font-size:${pt(r.vertAlign ? r.size * 0.65 : r.size)}`);
  // Superscript or subscript with its size from a style: 65% of the text around it (as the layout draws it).
  else if (r.vertAlign) css.push("font-size:65%");
  if (r.bold !== undefined) css.push(`font-weight:${r.bold ? 700 : 400}`);
  if (r.italic !== undefined) css.push(`font-style:${r.italic ? "italic" : "normal"}`);
  if (r.color) css.push(`color:#${r.color}`);
  if (r.underline || r.strike) {
    css.push(`text-decoration:${[r.underline ? "underline" : "", r.strike ? "line-through" : ""].join(" ").trim()}`);
    const u = r.underline ? r.underlineStyle : undefined;
    if (u) css.push(`text-decoration-style:${u.startsWith("double") ? "double" : u.startsWith("dotted") ? "dotted" : u.startsWith("dash") ? "dashed" : u.startsWith("wav") ? "wavy" : "solid"}`);
    if (u && /thick|Heavy/.test(u)) css.push("text-decoration-thickness:2px");
  }
  if (r.highlight) css.push(`background:${HIGHLIGHT[r.highlight] ?? (r.highlight.startsWith("#") ? r.highlight : "transparent")}`);
  // Raised or lowered by the layout's amounts (a third of the full size up, 0.14 down: in the run's own
  // 65% size, 0.51em and 0.215em). Runs have no line height of their own (styles.css), so this does
  // not make the line taller; position:relative did the same but made each run a layer of its own,
  // which the browser composites on every frame (a thousand in a long book).
  if (r.vertAlign) css.push(`vertical-align:${r.vertAlign === "superscript" ? "0.51em" : "-0.215em"}`);
  if (r.caps) css.push("text-transform:uppercase");
  if (r.charSpacing) css.push(`letter-spacing:${pt(r.charSpacing)}`);
  if (r.hidden) css.push("display:none");
  return css.join(";");
}

const dataUrl = (contentType: string, b64: string) => `data:${contentType};base64,${b64}`;
const IMAGE_TYPES: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml" };

export function createRender(doc: PrimeDocument, measurer: Measurer): Render {
  const resolver = new Resolver(doc);
  const math = new MathTypesetter(measurer);
  const baseSize = resolver.run(doc.styles.styles.find((s) => s.type === "paragraph" && s.isDefault)?.id ?? "", undefined).size;

  const chip = (label: string, title: string, data: Record<string, string>) =>
    ["span", { class: "pw-chip", title, contenteditable: "false", ...data }, label] as const;

  const mathSvg = (xml: string, data: string): HTMLElement => {
    const { box, display } = math.layout(xml, baseSize);
    const height = box.ascent + box.descent;
    const wrap = document.createElement("span");
    wrap.className = display ? "pw-math pw-math-display" : "pw-math";
    wrap.contentEditable = "false";
    wrap.setAttribute("data-pw-inlinepreserved", data);
    const svg = document.createElementNS(SVG, "svg");
    svg.setAttribute("width", pt(box.width));
    svg.setAttribute("height", pt(height));
    svg.setAttribute("viewBox", `0 0 ${box.width} ${height}`);
    svg.style.verticalAlign = pt(-box.descent);
    svg.style.overflow = "visible";
    for (const g of box.glyphs) {
      const t = document.createElementNS(SVG, "text");
      t.setAttribute("x", String(g.x));
      t.setAttribute("y", String(box.ascent + g.dy));
      t.setAttribute("font-family", "Cambria Math");
      t.setAttribute("font-size", String(g.size));
      t.textContent = g.text;
      svg.appendChild(t);
    }
    for (const r of box.rules) {
      const l = document.createElementNS(SVG, "line");
      l.setAttribute("x1", String(r.x1));
      l.setAttribute("x2", String(r.x2));
      l.setAttribute("y1", String(box.ascent + r.y1));
      l.setAttribute("y2", String(box.ascent + r.y2));
      l.setAttribute("stroke", "currentColor");
      l.setAttribute("stroke-width", String(r.width));
      svg.appendChild(l);
    }
    wrap.appendChild(svg);
    return wrap;
  };

  return {
    paragraph: () => ["p", 0],
    run: (mark) => runCss(resolver.direct(mark.attrs.props)),
    table: (node) => {
      const grid: number[] = node.attrs.grid;
      const props = node.attrs.props;
      const width = pt(grid.reduce((a, b) => a + b, 0));
      // In its section's text column: --pw-sx / --pw-sr move it there in Print Layout (paged.ts).
      const sx = "var(--pw-sx, 0pt)";
      const column = `(100% - ${sx} - var(--pw-sr, 0pt))`;
      const left = props.align === "center" ? `calc(${sx} + (${column} - ${width}) / 2)` : props.align === "right" ? `calc(${sx} + ${column} - ${width})` : sx;
      return ["table", {
        class: "pw-table", style: `width:${width};margin-left:${left}`,
        "data-pw-props": json(props), "data-pw-grid": json(grid),
      },
        ["colgroup", ...grid.map((w) => ["col", { style: `width:${pt(w)}` }])], ["tbody", 0]];
    },
    cell: (node) => {
      const p = node.attrs.props;
      const attrs: Record<string, string> = { "data-pw-props": json(p) };
      if (p.span > 1) attrs.colspan = String(p.span);
      const css = [p.shading?.fill && p.shading.fill !== "auto" ? `background:#${p.shading.fill}` : "",
        p.vMerge === "continue" ? "border-top-color:transparent" : "",
        p.vAlign ? `vertical-align:${p.vAlign === "center" ? "middle" : p.vAlign}` : ""].filter(Boolean);
      if (css.length) attrs.style = css.join(";");
      // w:hideMark: an empty cell's paragraph mark adds no height to its row (as in the layout).
      if (p.extra?.some((x: { name: string }) => x.name === "w:hideMark") && node.childCount === 1 && node.firstChild!.childCount === 0) {
        attrs.class = "pw-hidemark";
      }
      return ["td", attrs, 0];
    },
    image: (node) => {
      // A floating picture is drawn by the layout (Print Layout) or below its paragraph (Web Layout).
      if (node.attrs.float) return ["span", { class: "pw-float-anchor", "data-pw-image": json(node.attrs) }];
      const res = doc.images[node.attrs.imageId];
      return ["img", {
        src: res ? dataUrl(res.contentType, res.data) : "",
        alt: node.attrs.description ?? "",
        style: `width:${pt(node.attrs.width)};height:${pt(node.attrs.height)}`,
        "data-pw-image": json(node.attrs),
      }];
    },
    runPreserved: (node) => {
      // Floating pictures and text boxes: drawn by the layout in Print Layout (see paged.ts), and below
      // their paragraph in Web Layout (see floatCanvas). The anchor itself shows nothing.
      if (node.attrs.float) return ["span", { class: "pw-float-anchor", "data-pw-runpreserved": json(node.attrs) }];
      const p = node.attrs.preserved as { name: string; xml: string };
      const extent = /<wp:inline[ >][\s\S]*?<wp:extent cx="(\d+)" cy="(\d+)"/.exec(p.xml);
      // A chart: drawn from its part (what it shows, as Word caches it).
      const chartId = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(p.xml)?.[1];
      const chartRel = chartId ? doc.package.preservedRels.find((r) => r.id === chartId && !r.external) : undefined;
      const chartPart = chartRel && doc.package.parts.find((x) => x.name === chartRel.target);
      if (extent && chartPart) {
        const w = Number(extent[1]) / 12700;
        const h = Number(extent[2]) / 12700;
        const text = (b64: string) => new TextDecoder().decode(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
        const themePart = doc.package.parts.find((x) => /^word\/theme\/theme\d*\.xml$/.test(x.name));
        const spec = parseChart(text(chartPart.data), chartTheme(themePart ? text(themePart.data) : undefined));
        const svg = spec ? chartSvg(spec, w, h) : `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}"><rect width="${w}" height="${h}" fill="#f4f6fa" stroke="#bbb"/><text x="${w / 2}" y="${h / 2}" text-anchor="middle" font-family="sans-serif" font-size="11" fill="#777">Chart</text></svg>`;
        return ["img", {
          src: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
          class: "pw-chart",
          alt: spec?.title ? `Chart: ${spec.title}` : "Chart",
          style: `width:${pt(w)};height:${pt(h)}`,
          "data-pw-runpreserved": json(node.attrs),
        }];
      }
      const relId = /r:embed="([^"]+)"/.exec(p.xml)?.[1];
      const rel = doc.package.preservedRels.find((r) => r.id === relId && !r.external);
      const part = rel && doc.package.parts.find((x) => x.name === rel.target);
      if (extent && part) {
        const ext = part.name.split(".").pop()!.toLowerCase();
        return ["img", {
          src: dataUrl(IMAGE_TYPES[ext] ?? "application/octet-stream", part.data),
          class: /<a:ln\b/.test(p.xml) ? "pw-bordered" : "",
          style: `width:${pt(Number(extent[1]) / 12700)};height:${pt(Number(extent[2]) / 12700)}`,
          "data-pw-runpreserved": json(node.attrs),
        }];
      }
      const data = { "data-pw-runpreserved": json(node.attrs) };
      if (p.name === "w:footnoteReference") return ["sup", { class: "pw-chip", contenteditable: "false", ...data }, "†"];
      if (p.name === "w:commentReference") return chip("💬", "Comment — see Review → Comments", data);
      return chip("◆", `${p.name} (kept unchanged; not editable yet)`, data);
    },
    inlinePreserved: (node) => {
      const p = node.attrs.preserved as { name: string; xml: string };
      const data = { "data-pw-inlinepreserved": json(node.attrs) };
      if (p.name === "m:oMath" || p.name === "m:oMathPara") return mathSvg(p.xml, data["data-pw-inlinepreserved"]);
      if (/^w:(bookmark|commentRange|perm|proofErr)/.test(p.name)) return ["span", { class: "pw-hidden", ...data }];
      if (p.name === "w:commentReference") return chip("💬", "Comment — see Review → Comments", data);
      return chip("◆", `${p.name} (kept unchanged; not editable yet)`, data);
    },
    blockPreserved: (node) => ["div", { class: "pw-chip-block", contenteditable: "false", "data-pw-preserved": json(node.attrs.preserved) },
      `Content kept unchanged (${node.attrs.preserved.name}) — not editable yet`],
    sectionBreak: (node) => ["div", { class: "pw-section-break", contenteditable: "false", "data-pw-section": json(node.attrs.section) }, "Section break"],
  };
}

/**
 * Paragraph styling and list labels, recomputed in document order whenever the document changes.
 */
/**
 * Whether a transaction only changed text (or character formatting) inside existing paragraphs: no
 * paragraph was added, removed, split, joined or given new properties. Then no paragraph's own
 * formatting or list number can have changed — the case for nearly every keystroke.
 */
function textOnly(tr: Transaction): boolean {
  return tr.steps.every((step, i) => {
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) return true;
    if (!(step instanceof ReplaceStep) || step.slice.openStart || step.slice.openEnd) return false;
    const $from = tr.docs[i].resolve(step.from);
    if (!$from.parent.isTextblock || !$from.sameParent(tr.docs[i].resolve(step.to))) return false;
    let inline = true;
    step.slice.content.forEach((n) => { inline &&= n.isInline; });
    return inline;
  });
}

/**
 * Whether an edit changed the kinds of text in a paragraph (font, size, bold, italic, raised/lowered):
 * the paragraph's line height follows its text (paragraphStyles), so a bigger font must not keep the
 * old, shorter lines (they overlapped). Typing more of the same text changes nothing here.
 */
function runsChanged(tr: Transaction): boolean {
  const kinds = (para: Node) => {
    const out = new Set<string>();
    para.forEach((n) => {
      if (!n.isText) return;
      const p = n.marks.find((m) => m.type.name === "run")?.attrs.props ?? {};
      out.add(JSON.stringify([p.font, p.size, p.bold, p.italic, p.vertAlign, p.style]));
    });
    return out;
  };
  return tr.steps.some((step, i) => {
    const before = tr.docs[i];
    const after = i + 1 < tr.docs.length ? tr.docs[i + 1] : tr.doc;
    const { from, to } = step as ReplaceStep | AddMarkStep | RemoveMarkStep;
    let changed = false;
    before.nodesBetween(from, Math.max(from, to), (node, pos) => {
      if (changed) return false;
      if (!node.isTextblock) return true;
      const now = after.nodeAt(step.getMap().map(pos, -1));
      if (!now) return false;
      const a = kinds(node), b = kinds(now);
      changed = a.size !== b.size || [...a].some((k) => !b.has(k));
      return false;
    });
    return changed;
  });
}

/**
 * Web Layout (no pages, no layout engine): a paragraph's floating objects in one block below it,
 * placed relative to each other as in the document (offsets from the column and the paragraph);
 * the text continues below them. Text box heights are estimated from their text.
 */
function floatCanvas(floats: FloatInfo[], doc: PrimeDocument, resolver: Resolver): HTMLElement {
  const section = doc.sections[0]?.props ?? {};
  const mg = section.margins ?? {};
  const W = (section.pageWidth ?? 612) - (mg.left ?? 72) - (mg.right ?? 72);
  const items = floats.flatMap((f) => {
    const x = f.x.offset !== undefined ? f.x.offset - (f.x.from === "page" || f.x.from === "leftMargin" ? mg.left ?? 72 : 0)
      : f.x.align === "right" || f.x.align === "outside" ? W - f.width : f.x.align === "center" ? (W - f.width) / 2 : 0;
    const y = ["paragraph", "line"].includes(f.y.from) ? f.y.offset ?? 0 : 0;
    const part = f.imagePart ? doc.package.parts.find((p) => p.name === f.imagePart) : undefined;
    const res = f.imageId ? doc.images[f.imageId] : undefined;
    const el = document.createElement("span");
    if (part || res) {
      const img = document.createElement("img");
      img.src = res ? dataUrl(res.contentType, res.data) : dataUrl(IMAGE_TYPES[part!.name.split(".").pop()!.toLowerCase()] ?? "application/octet-stream", part!.data);
      img.alt = "";
      img.style.cssText = "width:100%;height:100%";
      el.appendChild(img);
      return [{ el, x, y, width: f.width, height: f.height }];
    }
    if (f.kind === "ink" && f.ink) {
      el.appendChild(inkSvg(f.ink, f.width, f.height));
      return [{ el, x, y, width: f.width, height: f.height }];
    }
    if (f.kind !== "textbox" || !f.blocks) return [];
    let chars = 0;
    for (const b of f.blocks) {
      if (b.type !== "paragraph") continue;
      const rp = resolver.paragraph({ ...b, content: [] });
      const first = b.content.find((i): i is TextInline => i.type === "text");
      const line = document.createElement("span");
      line.style.cssText = `display:block;${runCss(resolver.run(rp.styleId, first?.props ?? b.props.markRun))};text-align:${rp.align === "center" ? "center" : rp.align === "right" ? "right" : "left"}`;
      line.textContent = paragraphText(b);
      chars += line.textContent.length;
      el.appendChild(line);
    }
    const height = f.autoHeight || f.height < 1 ? Math.max(1, Math.ceil((chars * 5) / Math.max(f.width, 1))) * 14 : f.height;
    return [{ el, x, y, width: f.width, height }];
  });
  const canvas = document.createElement("div");
  canvas.className = "pw-float-canvas";
  canvas.contentEditable = "false";
  if (!items.length) return canvas;
  const top = Math.min(...items.map((i) => i.y));
  const left = Math.min(0, ...items.map((i) => i.x));
  canvas.style.height = `${Math.max(...items.map((i) => i.y + i.height)) - top}pt`;
  for (const i of items) {
    i.el.style.cssText += `;position:absolute;left:${i.x - left}pt;top:${i.y - top}pt;width:${i.width}pt;height:${i.height}pt`;
    canvas.appendChild(i.el);
  }
  return canvas;
}

/** A table node's structure and formatting, as the layout's cell resolver reads it (no content needed). */
function tableModel(node: Node): Table {
  const rows: Table["rows"] = [];
  node.forEach((row) => {
    const cells: Table["rows"][number]["cells"] = [];
    row.forEach((cell) => cells.push({ props: cell.attrs.props, blocks: [] }));
    rows.push({ props: row.attrs.props, cells });
  });
  return { type: "table", props: node.attrs.props, grid: node.attrs.grid, rows };
}

/** A cell's resolved borders and shading as CSS (sides without a border keep the editor's light gridlines). */
function cellCss(look: CellLook): string {
  const css: string[] = [];
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const b = look.borders[side];
    if (b) css.push(`border-${side}:${pt(b.width)} ${b.style.startsWith("double") ? "double" : b.style.startsWith("dot") ? "dotted" : b.style.startsWith("dash") ? "dashed" : "solid"} #${b.color ?? "000000"}`);
  }
  if (look.fill) css.push(`background:#${look.fill}`);
  return css.join(";");
}

/**
 * ProseMirror's DecorationSet.forChild, made fast for long documents. For every top-level node on
 * every update it scans all top-level decorations (the page gaps: hundreds) for inline ones, and the
 * list of children that have decorations (paragraphs with spelling marks: thousands) for its own —
 * 7,500 paragraphs × both lists made each keystroke in a 480-page book cost ~15 ms. This finds the
 * child by binary search (the list is sorted by position) and skips the scan when the set has no
 * inline decorations; with inline ones it does exactly what ProseMirror does. (prosemirror-view is
 * pinned to an exact version; the fields used here are that version's.)
 */
{
  type Set = { local: Decoration[]; children: (number | DecorationSet)[] };
  const proto = DecorationSet.prototype as unknown as { forChild(offset: number, node: Node): DecorationSet };
  const original = proto.forChild;
  const hasInline = new WeakMap<object, boolean>();
  proto.forChild = function (this: DecorationSet & Set, offset: number, node: Node) {
    if (this === DecorationSet.empty || node.isLeaf || !Array.isArray(this.local) || !Array.isArray(this.children)) return original.call(this, offset, node);
    let inline = hasInline.get(this);
    // Decoration.inline: a getter in this version's code, not in its type definitions.
    if (inline === undefined) hasInline.set(this, (inline = this.local.some((d) => (d as unknown as { inline: boolean }).inline)));
    if (inline) return original.call(this, offset, node);
    const ch = this.children;
    let lo = 0;
    let hi = ch.length / 3;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if ((ch[mid * 3] as number) < offset) lo = mid + 1;
      else hi = mid;
    }
    return lo * 3 < ch.length && ch[lo * 3] === offset ? (ch[lo * 3 + 2] as DecorationSet) : DecorationSet.empty;
  };
}

/**
 * ProseMirror's DecorationSet.create, made linear at the top level. It gives each top-level node its
 * decorations by scanning every decoration once per node: 9,778 paragraphs × thousands of spelling
 * marks (or page gaps, or line heights) cost about a second each time a set was rebuilt in a 944-page
 * book. This sorts the decorations once and walks the nodes and decorations together, handing each
 * node its own (built by ProseMirror itself, positions relative to the node, as ProseMirror does);
 * the rest stay at the top, sorted and checked as ProseMirror does. Same sets as the original.
 */
{
  type Ctor = new (local: Decoration[], children: (number | DecorationSet)[]) => DecorationSet;
  type Span = Decoration & { copy(from: number, to: number): Decoration; type: { valid(node: Node, span: Decoration): boolean } };
  const original = DecorationSet.create;
  DecorationSet.create = function (doc: Node, decorations: Decoration[]): DecorationSet {
    if (decorations.length < 64 || doc.childCount < 64) return original.call(DecorationSet, doc, decorations);
    const spans = (decorations as Span[]).slice().sort((a, b) => a.from - b.from || a.to - b.to);
    const locals: Decoration[] = [];
    const children: (number | DecorationSet)[] = [];
    let i = 0;
    doc.forEach((child, start) => {
      const end = start + child.nodeSize;
      // Those that start before this node (or at its start) are not inside it.
      while (i < spans.length && spans[i].from <= start) locals.push(spans[i++]);
      if (child.isLeaf) return;
      const inside: Decoration[] = [];
      while (i < spans.length && spans[i].from < end) {
        const s = spans[i++];
        if (s.to < end) inside.push(s.copy(s.from - start - 1, s.to - start - 1));
        else locals.push(s); // crosses the node's end
      }
      if (inside.length) {
        const sub = original.call(DecorationSet, child, inside);
        if (sub !== DecorationSet.empty) children.push(start, end, sub);
      }
    });
    while (i < spans.length) locals.push(spans[i++]);
    locals.sort((a, b) => a.from - b.from || a.to - b.to);
    const valid = (locals as Span[]).filter((s) => s.type.valid(doc, s));
    return valid.length || children.length ? new (DecorationSet as unknown as Ctor)(valid, children) : DecorationSet.empty;
  } as typeof DecorationSet.create;
}

/** The property names in a CSS declaration list (as ProseMirror itself reads decoration styles). */
const CSS_PROPERTY = /\s*([\w\-\xa1-￿]+)\s*:(?:"(?:\\.|[^"])*"|'(?:\\.|[^'])*'|\(.*?\)|[^;])*/g;

/** A paragraph's look: its CSS, its style's id (data-style) and classes. */
interface Look { style: string; styleId: string; className: string }

/**
 * A paragraph that shows its look itself. (Decorations on every paragraph would make each keystroke
 * cost time in proportion to the square of the document's length: ProseMirror scans all top-level
 * decorations for every top-level node.) The plugin restyles it when its look changes.
 */
class ParagraphView {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  private look?: Look;
  constructor(node: Node, look: Look | undefined, private views: Set<ParagraphView>) {
    const spec = DOMSerializer.renderSpec(document, node.type.spec.toDOM!(node));
    this.dom = spec.dom as HTMLElement;
    this.contentDOM = (spec.contentDOM ?? spec.dom) as HTMLElement;
    this.node = node;
    this.show(look);
    views.add(this);
  }
  node: Node;
  show(look: Look | undefined) {
    if (look === this.look || (look && this.look && look.style === this.look.style && look.styleId === this.look.styleId && look.className === this.look.className)) {
      this.look = look;
      return;
    }
    // Only our own properties and classes change: decorations (e.g. position:relative) keep theirs.
    const old = this.look;
    this.look = look;
    if (old?.style) for (const [, prop] of old.style.matchAll(CSS_PROPERTY)) this.dom.style.removeProperty(prop);
    if (look?.style) this.dom.style.cssText += `;${look.style}`;
    if (look?.styleId) this.dom.setAttribute("data-style", look.styleId);
    else this.dom.removeAttribute("data-style");
    for (const c of old?.className.split(" ") ?? []) if (c) this.dom.classList.remove(c);
    for (const c of look?.className.split(" ") ?? []) if (c) this.dom.classList.add(c);
  }
  update(node: Node, looks: WeakMap<Node, Look>): boolean {
    if (node.type !== this.node.type || JSON.stringify(node.attrs) !== JSON.stringify(this.node.attrs)) return false;
    this.node = node;
    this.show(looks.get(node));
    return true;
  }
  ignoreMutation(m: { type: string; target: globalThis.Node }) {
    return m.type === "attributes" && m.target === this.dom; // our own restyling
  }
  destroy() { this.views.delete(this); }
}

export function paragraphStyles(doc: PrimeDocument, measurer: Measurer): Plugin {
  /**
   * The line heights of a paragraph's text runs: the most common (by characters), which sets the
   * paragraph's line height, and each run's own with its offset. Word makes a line as tall as the text on
   * it (the paragraph mark counts only when there is none); on screen runs add no height of their own
   * (styles.css), and runs taller than the common text carry theirs in a decoration (pw-tall), so only
   * the lines they are on grow — as in the layout.
   */
  const runHeights = (node: Node, rp: ReturnType<Resolver["paragraph"]>, resolver: Resolver, tableStyle?: string, cond?: RunProps) => {
    const runs: { offset: number; size: number; height: number }[] = [];
    const weight = new Map<number, number>();
    node.forEach((child, offset) => {
      if (!child.isText) return;
      const props = child.marks.find((m) => m.type.name === "run")?.attrs.props;
      const r = resolver.run(rp.styleId, props, tableStyle, cond);
      if (r.hidden) return;
      const height = measurer.metrics({ family: r.font, bold: r.bold, italic: r.italic }).lineHeight * (r.vertAlign ? r.size * 0.65 : r.size);
      runs.push({ offset, size: child.nodeSize, height });
      const k = Math.round(height * 100) / 100;
      weight.set(k, (weight.get(k) ?? 0) + child.nodeSize);
    });
    const common = [...weight].sort((x, y) => y[1] - x[1])[0]?.[0];
    return { common, runs };
  };

  const build = (pm: Node): StylesState => {
    const looks = new WeakMap<Node, Look>();
    // The editor's own stylesheet once a style was created or modified (styleedit.ts), else the file's.
    const sheet = pm.attrs.styles ? { ...doc, styles: pm.attrs.styles } : doc;
    const resolver = new Resolver(sheet);
    // A second resolver to look at the next paragraph (the first counts list numbers in order).
    const peek = new Resolver(sheet);
    const decorations: Decoration[] = [];
    const visit = (node: Node, pos: number, tableStyle?: string, cond?: RunProps) => {
      let prev: { styleId: string; contextual: boolean } | undefined;
      // A dropped capital just before: the room it takes beside this paragraph's first lines.
      let capRoom: string | undefined;
      node.forEach((child, offset, index) => {
        const at = pos + offset;
        if (child.type.name !== "paragraph") prev = undefined;
        if (child.type.name === "paragraph") {
          const p: Paragraph = { type: "paragraph", props: child.attrs.props, content: [] };
          const rp = resolver.paragraph(p, tableStyle, true, cond);
          // Contextual spacing: no space between paragraphs of the same style that both ask for it.
          const next = node.maybeChild(index + 1);
          const nrp = next?.type.name === "paragraph" ? peek.paragraph({ type: "paragraph", props: next.attrs.props, content: [] }, tableStyle) : undefined;
          const joinBefore = !!prev && rp.contextualSpacing && prev.contextual && prev.styleId === rp.styleId;
          const joinAfter = !!nrp && rp.contextualSpacing && nrp.contextualSpacing && nrp.styleId === rp.styleId;
          prev = { styleId: rp.styleId, contextual: rp.contextualSpacing };
          // The paragraph mark's formatting (what Word sizes an empty line by, and usually the text's):
          // the paragraph's base font and line height. The style's own font can differ (e.g. a large
          // Title style with 16pt text), and a mismatched base font would stretch every line.
          const run = rp.mark;
          const heights = runHeights(child, rp, resolver, tableStyle, cond);
          const single = heights.common ?? measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic }).lineHeight * run.size;
          const lineOf = (h: number) => (rp.lineRule === "exact" ? rp.line / 20 : rp.lineRule === "atLeast" ? Math.max(rp.line / 20, h) : h * (rp.line / 240));
          const lineHeight = lineOf(single);
          // Taller runs: their lines grow to their height (an exact line height never grows).
          if (rp.lineRule !== "exact") {
            for (const r of heights.runs) {
              if (r.height <= single + 0.05) continue;
              decorations.push(Decoration.inline(at + 1 + r.offset, at + 1 + r.offset + r.size, { class: "pw-tall", style: `line-height:${pt(lineOf(r.height))}` }));
            }
          }
          const css = [
            runCss(run),
            // Plus the section's own text column in Print Layout (--pw-sx / --pw-sr, see paged.ts).
            `margin:${pt(joinBefore ? 0 : rp.spaceBefore)} calc(${pt(rp.right)} + var(--pw-sr, 0pt)) ${pt(joinAfter ? 0 : rp.spaceAfter)} calc(${pt(rp.left)} + var(--pw-sx, 0pt))`,
            `text-indent:${pt(rp.firstLine)}`,
            // Kept with the next paragraph / its lines together (Word's keep options): honoured where the
            // browser breaks text itself (text columns on screen).
            rp.keepNext ? "break-after:avoid" : "",
            rp.keepLines ? "break-inside:avoid" : "",
            `text-align:${rp.align === "both" || rp.align === "distribute" ? "justify" : rp.align === "end" ? "right" : rp.align === "start" ? "left" : rp.align}`,
            `line-height:${pt(lineHeight)}`,
            ...(["top", "bottom", "left", "right"] as const).map((k) => {
              const b = rp.borders[k];
              return b ? `border-${k}:${pt(b.width)} solid #${b.color ?? "000000"};padding-${k}:${pt(b.space)}` : "";
            }),
          ].filter(Boolean).join(";");
          // A line with one tab to a right-aligned stop (a contents entry: "Heading ....... 4"): the text after
          // the tab sits at the right edge, the leader (dots) fills the gap. Browsers have no tab stops, so
          // the line becomes a flexible row in which the tab stretches.
          let tabs = 0;
          child.forEach((n) => { if (n.type.name === "tab") tabs++; });
          const rightStop = tabs === 1 ? rp.tabs.find((t) => t.align === "right") : undefined;
          const classes = rightStop ? `pw-righttab${rightStop.leader ? ` pw-leader-${rightStop.leader}` : ""}` : "";
          // A drop cap (its own paragraph, before the one it belongs to): floated left, so the next
          // paragraph's first lines wrap beside it, and moved down so its baseline sits on the baseline
          // of that paragraph's `lines`-th line (as the layout places it). "In margin": left of the text.
          let drop = "";
          const room = capRoom;
          capRoom = undefined;
          if (rp.dropCap && nrp) {
            const m = (r: typeof run) => measurer.metrics({ family: r.font, bold: r.bold, italic: r.italic });
            const nm = m(nrp.mark);
            const nSingle = nm.lineHeight * nrp.mark.size;
            const nLine = nrp.lineRule === "exact" ? nrp.line / 20 : nrp.lineRule === "atLeast" ? Math.max(nrp.line / 20, nSingle) : nSingle * (nrp.line / 240);
            // Baselines in CSS line boxes: half the leading above the font's ascent.
            const nBaseline = (nLine - nSingle) / 2 + nm.ascent * nrp.mark.size;
            const target = (rp.dropCap.lines - 1) * nLine + nBaseline;
            const lm = m(run);
            const baseline = (lineHeight - lm.lineHeight * run.size) / 2 + lm.ascent * run.size;
            const space = rp.dropCap.space;
            drop = rp.dropCap.kind === "margin"
              ? `;float:left;margin:${pt(nrp.spaceBefore)} 0 0 0;width:0;overflow:visible;direction:rtl;position:relative;top:${pt(target - baseline)};left:${pt(-space)}`
              // As tall as the lines beside it (the letter's own glyph box is taller, and would narrow one more).
              : `;float:left;margin:${pt(nrp.spaceBefore)} ${pt(space)} 0 0;text-indent:0;height:${pt(rp.dropCap.lines * nLine - 0.5)};overflow:visible;position:relative;top:${pt(target - baseline)}`;
            // Where the page spans text columns (each paragraph a block of its own, styles.css), the letter
            // cannot float into the next paragraph: that paragraph keeps the room for it instead.
            if (rp.dropCap.kind !== "margin") {
              const first = child.firstChild;
              const r = resolver.run(rp.styleId, first?.marks.find((mk) => mk.type.name === "run")?.attrs.props, tableStyle, cond);
              const width = measurer.width(child.textContent, { family: r.font, bold: r.bold, italic: r.italic }) * r.size;
              capRoom = `--pw-dc-w:${pt(width + space)};--pw-dc-h:${pt(rp.dropCap.lines * nLine - 0.5)}`;
            }
          }
          if (room) drop += `;${room}`;
          looks.set(child, { style: css + (rp.shading ? `;background:#${rp.shading}` : "") + drop, styleId: rp.styleId, className: rp.dropCap && nrp ? [classes, "pw-dropcap"].filter(Boolean).join(" ") : room ? [classes, "pw-dropcap-next"].filter(Boolean).join(" ") : classes });
          const floats: FloatInfo[] = [];
          child.forEach((n) => {
            if (n.type.name === "run_preserved" && n.attrs.float) floats.push(n.attrs.float);
            if (n.type.name === "image" && n.attrs.float) floats.push({ ...n.attrs.float, kind: "picture", width: n.attrs.width, height: n.attrs.height, imageId: n.attrs.imageId });
          });
          if (floats.length) {
            const key = `floats-${floats.map((f) => `${f.x.offset}:${f.y.offset}:${f.width}:${f.height}`).join("|")}`;
            decorations.push(Decoration.widget(at + child.nodeSize, () => floatCanvas(floats, doc, resolver), { side: -1, key }));
          }
          if (rp.listLabel !== undefined) {
            const label = rp.listLabel;
            const lr = rp.labelRun ?? rp.mark;
            const width = rp.firstLine < 0 ? -rp.firstLine : 18;
            decorations.push(Decoration.widget(at + 1, () => {
              const span = document.createElement("span");
              span.className = "pw-list-label";
              span.contentEditable = "false";
              // line-height 0: the label's font (e.g. Symbol) must not make the line taller than the
              // layout's, or the screen's pages would run longer than the document's.
              span.style.cssText = `${runCss(lr)};display:inline-block;min-width:${pt(width)};text-indent:0;line-height:0`;
              span.textContent = label;
              return span;
            }, { side: -1, key: `label-${label}-${width}-${runCss(lr)}` })); // no position: typing must not redraw every label after it
          }
        } else if (child.type.name === "table") {
          // Each cell's borders, shading and header/band text formatting, as the layout resolves them.
          const style = resolver.tableStyle(child.attrs.props.style);
          const looks = cellLooks(doc, tableModel(child), style);
          child.forEach((row, rowOffset, r) => {
            const rowAt = at + 1 + rowOffset;
            row.forEach((cell, cellOffset, k) => {
              const cellAt = rowAt + 1 + cellOffset;
              const look = looks[r]?.[k];
              const css = look ? cellCss(look) : "";
              if (css) decorations.push(Decoration.node(cellAt, cellAt + cell.nodeSize, { style: css }));
              visit(cell, cellAt + 1, style, look?.run);
            });
          });
        } else if (child.childCount) {
          visit(child, at + 1, tableStyle, cond);
        }
      });
    };
    visit(pm, 0);
    return { looks, decorations: DecorationSet.create(pm, decorations) };
  };
  /** Typing changes no looks: each edited paragraph keeps the look it had. */
  const carry = (tr: Transaction, old: StylesState): StylesState => {
    tr.steps.forEach((step, i) => {
      const before = tr.docs[i];
      const after = i + 1 < tr.docs.length ? tr.docs[i + 1] : tr.doc;
      const { from, to } = step as ReplaceStep | AddMarkStep | RemoveMarkStep;
      // Every paragraph the step touches (one for typing; several for bold across paragraphs).
      before.nodesBetween(from, to, (node, pos) => {
        if (!node.isTextblock) return true;
        const look = old.looks.get(node);
        const now = after.nodeAt(step.getMap().map(pos, -1));
        if (look && now && now !== node) old.looks.set(now, look);
        return false;
      });
    });
    return { looks: old.looks, decorations: old.decorations.map(tr.mapping, tr.doc) };
  };
  const views = new Set<ParagraphView>();
  const key = new PluginKey<StylesState>("paragraphStyles");
  return new Plugin<StylesState>({
    key,
    state: {
      init: (_, state) => build(state.doc),
      apply: (tr, old) => (!tr.docChanged ? old : textOnly(tr) && !runsChanged(tr) ? timed("styles.map", () => carry(tr, old)) : timed("styles.build", () => build(tr.doc))),
    },
    props: {
      decorations: (state) => key.getState(state)?.decorations,
      nodeViews: {
        paragraph: (node, view) => {
          const p = new ParagraphView(node, key.getState(view.state)?.looks.get(node), views);
          return { dom: p.dom, contentDOM: p.contentDOM, update: (n) => p.update(n, key.getState(view.state)!.looks), ignoreMutation: (m) => p.ignoreMutation(m), destroy: () => p.destroy() };
        },
      },
    },
    // After a rebuild, paragraphs that were not redrawn may have a new look (spacing next to a new
    // paragraph of the same style, a changed style): restyle them.
    view: () => ({
      update(view, prev) {
        const looks = key.getState(view.state)!.looks;
        if (looks === key.getState(prev)?.looks) return;
        timed("styles.restyle", () => {
          for (const p of views) {
            let look = looks.get(p.node);
            // ProseMirror keeps a paragraph's view without telling it when the new document has an equal
            // (but new) node there — the paragraph before a split by Enter. Its look is under the new
            // node: find it, or the paragraph would lose its formatting (a list item its indent).
            if (!look && p.dom.isConnected) {
              try {
                const now = view.state.doc.nodeAt(view.posAtDOM(p.contentDOM, 0) - 1);
                if (now && now !== p.node && now.type === p.node.type) { p.node = now; look = looks.get(now); }
              } catch { /* not in the document any more */ }
            }
            p.show(look);
          }
        });
      },
    }),
  });
}

interface StylesState { looks: WeakMap<Node, Look>; decorations: DecorationSet }
