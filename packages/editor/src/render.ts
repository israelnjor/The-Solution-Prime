/**
 * How the document looks in the editor (browser only).
 *
 * Paragraph styling is computed in one document-order pass by the layout engine's style resolver
 * (so list numbering, table styles and spacing match the layout engine), and applied as node
 * decorations. Runs render their own direct formatting; everything else is inherited via CSS.
 * Equations are drawn as SVG by the math typesetter.
 */
import { MathTypesetter, type Measurer, type ResolvedRun, Resolver } from "@prime/layout";
import type { Paragraph, PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { Plugin, type Transaction } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from "prosemirror-transform";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { Render } from "./schema.js";

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
  }
  if (r.highlight) css.push(`background:${HIGHLIGHT[r.highlight] ?? "transparent"}`);
  if (r.vertAlign) css.push(`vertical-align:${r.vertAlign === "superscript" ? "super" : "sub"}`);
  if (r.caps) css.push("text-transform:uppercase");
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
      return ["td", attrs, 0];
    },
    image: (node) => {
      const res = doc.images[node.attrs.imageId];
      return ["img", {
        src: res ? dataUrl(res.contentType, res.data) : "",
        alt: node.attrs.description ?? "",
        style: `width:${pt(node.attrs.width)};height:${pt(node.attrs.height)}`,
        "data-pw-image": json(node.attrs),
      }];
    },
    runPreserved: (node) => {
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
      return chip("◆", `${p.name} (kept unchanged; not editable yet)`, data);
    },
    inlinePreserved: (node) => {
      const p = node.attrs.preserved as { name: string; xml: string };
      const data = { "data-pw-inlinepreserved": json(node.attrs) };
      if (p.name === "m:oMath" || p.name === "m:oMathPara") return mathSvg(p.xml, data["data-pw-inlinepreserved"]);
      if (/^w:(bookmark|commentRange|perm|proofErr)/.test(p.name)) return ["span", { class: "pw-hidden", ...data }];
      if (p.name === "w:commentReference") return chip("💬", "Comment (kept unchanged; not editable yet)", data);
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

export function paragraphStyles(doc: PrimeDocument, measurer: Measurer): Plugin {
  const build = (pm: Node): DecorationSet => {
    const resolver = new Resolver(doc);
    const decorations: Decoration[] = [];
    const visit = (node: Node, pos: number, tableStyle?: string) => {
      node.forEach((child, offset) => {
        const at = pos + offset;
        if (child.type.name === "paragraph") {
          const p: Paragraph = { type: "paragraph", props: child.attrs.props, content: [] };
          const rp = resolver.paragraph(p, tableStyle);
          const run = resolver.run(rp.styleId, undefined, tableStyle);
          const lineHeight = rp.lineRule === "exact" ? rp.line / 20
            : rp.lineRule === "atLeast" ? Math.max(rp.line / 20, measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic }).lineHeight * run.size)
            : measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic }).lineHeight * run.size * (rp.line / 240);
          const css = [
            runCss(run),
            `margin:${pt(rp.spaceBefore)} ${pt(rp.right)} ${pt(rp.spaceAfter)} ${pt(rp.left)}`,
            `text-indent:${pt(rp.firstLine)}`,
            `text-align:${rp.align === "both" || rp.align === "distribute" ? "justify" : rp.align === "end" ? "right" : rp.align === "start" ? "left" : rp.align}`,
            `line-height:${pt(lineHeight)}`,
            ...(["top", "bottom", "left", "right"] as const).map((k) => {
              const b = rp.borders[k];
              return b ? `border-${k}:${pt(b.width)} solid #${b.color ?? "000000"};padding-${k}:${pt(b.space)}` : "";
            }),
          ].filter(Boolean).join(";");
          decorations.push(Decoration.node(at, at + child.nodeSize, { style: css, "data-style": rp.styleId }));
          if (rp.listLabel !== undefined) {
            const label = rp.listLabel;
            const lr = rp.labelRun ?? rp.mark;
            const width = rp.firstLine < 0 ? -rp.firstLine : 18;
            decorations.push(Decoration.widget(at + 1, () => {
              const span = document.createElement("span");
              span.className = "pw-list-label";
              span.contentEditable = "false";
              span.style.cssText = `${runCss(lr)};display:inline-block;min-width:${pt(width)};text-indent:0`;
              span.textContent = label;
              return span;
            }, { side: -1, key: `label-${at}-${label}` }));
          }
        } else if (child.type.name === "table") {
          visit(child, at + 1, resolver.tableStyle(child.attrs.props.style));
        } else if (child.childCount) {
          visit(child, at + 1, tableStyle);
        }
      });
    };
    visit(pm, 0);
    return DecorationSet.create(pm, decorations);
  };
  return new Plugin({
    state: {
      init: (_, state) => build(state.doc),
      apply: (tr, old) => (!tr.docChanged ? old : textOnly(tr) ? old.map(tr.mapping, tr.doc) : build(tr.doc)),
    },
    props: { decorations(state) { return this.getState(state); } },
  });
}
