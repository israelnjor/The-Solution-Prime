/**
 * Pasting from Word and the web. The schema's parse rules read HTML into paragraph properties,
 * tables and temporary "x_" formatting marks; after each paste, `pastePlugin` merges those marks
 * into run formatting, creates list definitions, registers pasted pictures and sizes tables.
 */
import type { ParagraphProps, PrimeDocument, RunProps } from "@prime/pdm";
import { Fragment, type Mark, type Node, type Schema, Slice } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import { ensureList } from "./lists.js";

// --- CSS ---------------------------------------------------------------------------------

/** A CSS length in points (pt, px, in, cm, mm, em), or undefined. */
export function toPt(value: string | undefined): number | undefined {
  const m = /^\s*(-?[\d.]+)\s*(pt|px|in|cm|mm|em|rem)?\s*$/i.exec(value ?? "");
  if (!m) return undefined;
  const n = Number(m[1]);
  switch ((m[2] ?? "px").toLowerCase()) {
    case "pt": return n;
    case "in": return n * 72;
    case "cm": return (n * 72) / 2.54;
    case "mm": return (n * 72) / 25.4;
    case "em": case "rem": return n * 12;
    default: return n * 0.75;
  }
}

const NAMED: Record<string, string> = {
  red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", orange: "FFA500", purple: "800080",
  gray: "808080", grey: "808080", white: "FFFFFF", navy: "000080", maroon: "800000", teal: "008080",
};

/** CSS colour as RRGGBB, or undefined for black/automatic/unknown. */
export function toHex(value: string | undefined): string | undefined {
  const v = (value ?? "").trim().toLowerCase();
  if (!v || v === "black" || v === "windowtext" || v === "inherit" || v === "initial" || v === "transparent") return undefined;
  if (NAMED[v]) return NAMED[v];
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return m[1].split("").map((c) => c + c).join("").toUpperCase();
  m = /^#([0-9a-f]{6})$/.exec(v);
  if (m) return m[1] === "000000" ? undefined : m[1].toUpperCase();
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (rgb) {
    const hex = rgb.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase();
    return hex === "000000" ? undefined : hex;
  }
  return undefined;
}

const HIGHLIGHTS: Record<string, string> = {
  FFFF00: "yellow", "00FF00": "green", "00FFFF": "cyan", FF00FF: "magenta", FF0000: "red", "0000FF": "blue", C0C0C0: "lightGray",
};

function declarations(css: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of css.split(";")) {
    const i = part.indexOf(":");
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

const GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "inherit", "initial"]);

/** Run formatting from inline CSS. */
export function cssToRunProps(css: string): RunProps {
  const d = declarations(css);
  const p: RunProps = {};
  if (d["font-weight"] && /^(bold|[6-9]00)$/.test(d["font-weight"])) p.bold = true;
  if (d["font-style"] === "italic") p.italic = true;
  const deco = d["text-decoration"] ?? d["text-decoration-line"] ?? "";
  if (deco.includes("underline")) p.underline = { val: "single" };
  if (deco.includes("line-through")) p.strike = true;
  const color = toHex(d.color);
  if (color) p.color = color;
  const size = toPt(d["font-size"]);
  if (size && size >= 4 && size <= 400) p.size = Math.round(size * 2) / 2;
  const family = (d["font-family"] ?? "").split(",").map((f) => f.trim().replace(/^["']|["']$/g, "")).find((f) => f && !GENERIC.has(f.toLowerCase()) && !f.startsWith("mso-"));
  if (family) p.font = { ascii: family, hAnsi: family, cs: family };
  const background = toHex(d["background-color"] ?? d.background ?? d["mso-highlight"]);
  if (background && HIGHLIGHTS[background]) p.highlight = HIGHLIGHTS[background];
  if (d["vertical-align"] === "super") p.vertAlign = "superscript";
  if (d["vertical-align"] === "sub") p.vertAlign = "subscript";
  return p;
}

// --- HTML elements ---------------------------------------------------------------------------

const ALIGN: Record<string, string> = { center: "center", right: "right", justify: "both" };

/**
 * Paragraph properties from an HTML paragraph or list item. Lists get a placeholder numbering
 * ("__bullet" / "__number") that pastePlugin replaces with a real list definition.
 */
export function paragraphFromHtml(el: HTMLElement): ParagraphProps {
  const d = declarations(el.getAttribute("style") ?? "");
  const p: ParagraphProps = {};
  const align = ALIGN[(d["text-align"] ?? el.getAttribute("align") ?? "").toLowerCase()];
  if (align) p.align = align;
  if (el.tagName === "LI") {
    let depth = 0;
    let kind: "bullet" | "number" = "bullet";
    for (let a = el.parentElement, first = true; a; a = a.parentElement) {
      if (a.tagName === "UL" || a.tagName === "OL") {
        if (first) kind = a.tagName === "OL" ? "number" : "bullet";
        first = false;
        depth++;
      }
    }
    p.list = { numId: `__${kind}`, level: Math.max(0, Math.min(8, depth - 1)) };
    return p;
  }
  // Word: <p style="mso-list:l0 level2 lfo1"> with the bullet or number in an mso-list:Ignore span.
  const msoList = d["mso-list"];
  if (msoList) {
    const level = Number(/level(\d)/.exec(msoList)?.[1] ?? 1) - 1;
    const marker = el.querySelector("span[style*='mso-list:Ignore']")?.textContent?.trim() ?? "";
    p.list = { numId: /^[\da-zA-Z]+[.)]?$/.test(marker) ? "__number" : "__bullet", level: Math.max(0, Math.min(8, level)) };
    return p;
  }
  const left = toPt(d["margin-left"]);
  if (left && left > 1) p.indent = { left };
  const first = toPt(d["text-indent"]);
  if (first) p.indent = { ...p.indent, ...(first > 0 ? { firstLine: first } : { hanging: -first }) };
  if (/MsoTitle/.test(el.className)) p.style = "Title";
  return p;
}

export function cellFromHtml(el: HTMLElement): Record<string, unknown> {
  const span = Number(el.getAttribute("colspan") ?? 1);
  const d = declarations(el.getAttribute("style") ?? "");
  const fill = toHex(d["background-color"] ?? d.background ?? el.getAttribute("bgcolor") ?? undefined);
  return {
    ...(span > 1 ? { span } : {}),
    ...(fill ? { shading: { val: "clear", color: "auto", fill } } : {}),
  };
}

const BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';

/** Table properties from HTML: bordered tables keep single borders; the grid is sized after pasting. */
export function tableFromHtml(el: HTMLElement): Record<string, unknown> {
  const style = el.getAttribute("style") ?? "";
  const bordered = Number(el.getAttribute("border") ?? 0) > 0 || /border[^:]*:\s*[^;]*solid/i.test(style) || !!el.querySelector("td[style*='solid'], th[style*='solid']");
  return { props: bordered ? { extra: [{ name: "w:tblBorders", xml: BORDERS }] } : {}, grid: [] };
}

// --- after the paste ---------------------------------------------------------------------------

const twips = (pt: number) => String(Math.round(pt * 20));

function textWidth(doc: PrimeDocument): number {
  const s = doc.sections[0]?.props ?? {};
  return (s.pageWidth ?? 612) - (s.margins?.left ?? 72) - (s.margins?.right ?? 72);
}

/** Merges temporary formatting marks into one run mark (keeping our own run properties). */
function textNode(node: Node, schema: Schema): Node {
  let props: RunProps = { ...(node.marks.find((m) => m.type.name === "run")?.attrs.props ?? {}) };
  for (const m of node.marks) {
    switch (m.type.name) {
      case "x_bold": props.bold = true; break;
      case "x_italic": props.italic = true; break;
      case "x_underline": props.underline = { val: "single" }; break;
      case "x_strike": props.strike = true; break;
      case "x_super": props.vertAlign = "superscript"; break;
      case "x_sub": props.vertAlign = "subscript"; break;
      case "x_style": props = { ...props, ...cssToRunProps(m.attrs.css) }; break;
    }
  }
  let marks: readonly Mark[] = node.marks.filter((m) => !m.type.name.startsWith("x_") && m.type.name !== "run");
  if (Object.keys(props).length) marks = schema.marks.run.create({ props }).addToSet(marks);
  return schema.text(node.text!, marks);
}

function dataUrl(src: string): { contentType: string; data: string; extension: string } | undefined {
  const m = /^data:(image\/(png|jpe?g|gif|bmp));base64,(.+)$/i.exec(src);
  if (!m) return undefined;
  const extension = m[2].toLowerCase() === "jpg" ? "jpeg" : m[2].toLowerCase();
  return { contentType: m[1].toLowerCase().replace("jpg", "jpeg"), data: m[3], extension };
}

function convert(node: Node, schema: Schema, doc: PrimeDocument): Node | undefined {
  if (node.isText) return textNode(node, schema);
  const marks = node.marks.filter((m) => !m.type.name.startsWith("x_"));
  let attrs = node.attrs;
  if (node.type.name === "paragraph" && typeof attrs.props.list?.numId === "string" && attrs.props.list.numId.startsWith("__")) {
    const numId = ensureList(doc, attrs.props.list.numId === "__number" ? "number" : "bullet");
    attrs = { ...attrs, props: { ...attrs.props, list: { ...attrs.props.list, numId } } };
  }
  if (node.type.name === "image" && String(attrs.imageId).startsWith("pending:")) {
    const res = dataUrl(String(attrs.imageId).slice("pending:".length));
    if (!res) return undefined; // not an embeddable picture
    let n = Object.keys(doc.images).length + 1;
    while (doc.images[`word/media/prime-image${n}.${res.extension}`]) n++;
    const imageId = `word/media/prime-image${n}.${res.extension}`;
    doc.images[imageId] = { contentType: res.contentType, extension: res.extension, data: res.data };
    const max = textWidth(doc);
    const scale = attrs.width > max ? max / attrs.width : 1;
    attrs = { ...attrs, imageId, width: attrs.width * scale, height: attrs.height * scale };
  }
  const children: Node[] = [];
  node.forEach((c) => {
    const converted = convert(c, schema, doc);
    if (converted) children.push(converted);
  });
  if (node.type.name === "table" && !(attrs.grid as number[]).length) {
    // A pasted HTML table: equal columns across the text width.
    const cols = Math.max(1, ...children.map((r) => { let n = 0; r.forEach((c) => { n += c.attrs.props.span ?? 1; }); return n; }));
    const col = textWidth(doc) / cols;
    attrs = { ...attrs, grid: Array(cols).fill(col) };
    const rows = children.map((r) => {
      const cells: Node[] = [];
      r.forEach((c) => cells.push(c.type.create({ props: { ...c.attrs.props, width: { w: twips(col * (c.attrs.props.span ?? 1)), type: "dxa" } } }, c.content)));
      return r.type.create(r.attrs, cells);
    });
    return node.type.create(attrs, rows, marks);
  }
  try {
    return node.type.create(attrs, Fragment.fromArray(children), marks);
  } catch {
    return node.type.createAndFill(attrs, Fragment.fromArray(children), marks) ?? undefined;
  }
}

/** Finishes every paste: formatting marks merged, lists created, pictures stored, tables sized. */
export function pastePlugin(doc: () => PrimeDocument): Plugin {
  return new Plugin({
    props: {
      transformPasted(slice, view) {
        const nodes: Node[] = [];
        slice.content.forEach((n) => {
          const c = convert(n, view.state.schema, doc());
          if (c) nodes.push(c);
        });
        return new Slice(Fragment.fromArray(nodes), slice.openStart, slice.openEnd);
      },
      // Pasting into an empty paragraph: like Word, it takes the first pasted paragraph's formatting
      // (otherwise a pasted heading would merge into the empty paragraph and lose its style).
      handlePaste(view, _event, slice) {
        const { $from, empty } = view.state.selection;
        const first = slice.content.firstChild;
        if (!empty || $from.parent.type.name !== "paragraph" || $from.parent.content.size > 0) return false;
        if (slice.openStart < 1 || first?.type.name !== "paragraph" || !Object.keys(first.attrs.props).length) return false;
        const at = $from.before();
        const tr = view.state.tr.replaceSelection(slice);
        tr.setNodeMarkup(tr.mapping.map(at, -1), undefined, { ...$from.parent.attrs, props: first.attrs.props });
        view.dispatch(tr.scrollIntoView().setMeta("paste", true).setMeta("uiEvent", "paste"));
        return true;
      },
    },
  });
}
