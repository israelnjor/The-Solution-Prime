/**
 * How the document looks in the editor (browser only).
 *
 * Paragraph styling is computed in one document-order pass by the layout engine's style resolver
 * (so list numbering, table styles and spacing match the layout engine), and applied as node
 * decorations. Runs render their own direct formatting; everything else is inherited via CSS.
 * Equations are drawn as SVG by the math typesetter.
 */
import { MathTypesetter, type Measurer, type ResolvedRun, Resolver } from "@prime/layout";
import { type FloatInfo, type Paragraph, type PrimeDocument, paragraphText, type TextInline } from "@prime/pdm";
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
  if (r.bold !== undefined) css.push(`font-weight:${r.bold ? 700 : 400}`);
  if (r.italic !== undefined) css.push(`font-style:${r.italic ? "italic" : "normal"}`);
  if (r.color) css.push(`color:#${r.color}`);
  if (r.underline || r.strike) {
    css.push(`text-decoration:${[r.underline ? "underline" : "", r.strike ? "line-through" : ""].join(" ").trim()}`);
    const u = r.underline ? r.underlineStyle : undefined;
    if (u) css.push(`text-decoration-style:${u.startsWith("double") ? "double" : u.startsWith("dotted") ? "dotted" : u.startsWith("dash") ? "dashed" : u.startsWith("wav") ? "wavy" : "solid"}`);
    if (u && /thick|Heavy/.test(u)) css.push("text-decoration-thickness:2px");
  }
  if (r.highlight) css.push(`background:${HIGHLIGHT[r.highlight] ?? "transparent"}`);
  if (r.vertAlign) css.push(`vertical-align:${r.vertAlign === "superscript" ? "super" : "sub"}`);
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

  const tableHasBorders = (props: { style?: string; extra?: { name: string; xml: string }[] }) => {
    const style = resolver.tableStyle(props.style);
    const xml = [...(props.extra ?? []), ...doc.styles.styles.filter((s) => s.id === style).flatMap((s) => s.table?.extra ?? [])]
      .find((x) => x.name === "w:tblBorders")?.xml;
    return !!xml && /w:val="(?!none|nil)[^"]+"/.test(xml);
  };

  return {
    paragraph: () => ["p", 0],
    run: (mark) => runCss(resolver.direct(mark.attrs.props)),
    table: (node) => {
      const grid: number[] = node.attrs.grid;
      const props = node.attrs.props;
      const align = props.align === "center" ? "margin-left:auto;margin-right:auto" : props.align === "right" ? "margin-left:auto" : "";
      return ["table", {
        class: tableHasBorders(props) ? "pw-table pw-borders" : "pw-table", style: `width:${pt(grid.reduce((a, b) => a + b, 0))};${align}`,
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
  const build = (pm: Node): StylesState => {
    const looks = new WeakMap<Node, Look>();
    const resolver = new Resolver(doc);
    // A second resolver to look at the next paragraph (the first counts list numbers in order).
    const peek = new Resolver(doc);
    const decorations: Decoration[] = [];
    const visit = (node: Node, pos: number, tableStyle?: string) => {
      let prev: { styleId: string; contextual: boolean } | undefined;
      node.forEach((child, offset, index) => {
        const at = pos + offset;
        if (child.type.name !== "paragraph") prev = undefined;
        if (child.type.name === "paragraph") {
          const p: Paragraph = { type: "paragraph", props: child.attrs.props, content: [] };
          const rp = resolver.paragraph(p, tableStyle);
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
          const lineHeight = rp.lineRule === "exact" ? rp.line / 20
            : rp.lineRule === "atLeast" ? Math.max(rp.line / 20, measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic }).lineHeight * run.size)
            : measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic }).lineHeight * run.size * (rp.line / 240);
          const css = [
            runCss(run),
            `margin:${pt(joinBefore ? 0 : rp.spaceBefore)} ${pt(rp.right)} ${pt(joinAfter ? 0 : rp.spaceAfter)} ${pt(rp.left)}`,
            `text-indent:${pt(rp.firstLine)}`,
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
          looks.set(child, { style: css + (rp.shading ? `;background:#${rp.shading}` : ""), styleId: rp.styleId, className: classes });
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
          visit(child, at + 1, resolver.tableStyle(child.attrs.props.style));
        } else if (child.childCount) {
          visit(child, at + 1, tableStyle);
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
      apply: (tr, old) => (!tr.docChanged ? old : textOnly(tr) ? timed("styles.map", () => carry(tr, old)) : timed("styles.build", () => build(tr.doc))),
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
        timed("styles.restyle", () => { for (const p of views) p.show(looks.get(p.node)); });
      },
    }),
  });
}

interface StylesState { looks: WeakMap<Node, Look>; decorations: DecorationSet }
