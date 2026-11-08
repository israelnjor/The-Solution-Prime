/**
 * OpenDocument Text (.odt: LibreOffice, OpenOffice, Google Docs' export), read and written. Read:
 * paragraphs and headings with their styles (named and automatic, inherited) as direct formatting:
 * fonts, sizes, bold, italic, underline, strike, colour, super/subscript, small caps, alignment,
 * indents, spacing, line height, keep with next, page breaks; lists (bullets or numbers, nested);
 * tables (spans, column widths, header rows); pictures; links; spaces, tabs, line breaks; page size and
 * margins; title and author. Written: the same, from what the document shows (styles resolved, list
 * labels as text).
 */
import { withProperties } from "@prime/docx";
import type { Block, Inline, Paragraph, ParagraphProps, PrimeDocument, RunProps, Table, TableCell } from "@prime/pdm";
import { DOMParser } from "@xmldom/xmldom";
import { strToU8, unzipSync, zipSync } from "fflate";
import { addImage, headingStyle, imageBytes, listId, newDocument, type Piece, type Shown, type ShownParagraph, shownBody } from "./common.js";

const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
  meta: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
};
type El = Element;
const kids = (e: El): El[] => { const out: El[] = []; for (let c = e.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as El); return out; };
const is = (e: El, ns: string, local: string) => e.namespaceURI === ns && e.localName === local;
const at = (e: El | undefined, ns: string, local: string) => (e?.getAttributeNS(ns, local) || undefined);

/** "2.54cm", "1in", "12pt", "5mm", "0.2inch" → pt. */
export function toPt(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|inch|pt|pc|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  return { cm: n * 72 / 2.54, mm: n * 72 / 25.4, in: n * 72, inch: n * 72, pt: n, pc: n * 12, px: n * 0.75 }[m[2] ?? "pt"] ?? n;
}

interface OdfStyle { family: string; parent?: string; text: Record<string, string>; para: Record<string, string>; displayName?: string; listStyle?: string; cell: Record<string, string> }

function readStyles(root: El | undefined, into: Map<string, OdfStyle>, defaults: Map<string, OdfStyle>, lists: Map<string, El>) {
  if (!root) return;
  for (const s of kids(root)) {
    if (is(s, NS.text, "list-style")) { const n = at(s, NS.style, "name"); if (n) lists.set(n, s); continue; }
    const isDefault = is(s, NS.style, "default-style");
    if (!is(s, NS.style, "style") && !isDefault) continue;
    const st: OdfStyle = { family: at(s, NS.style, "family") ?? "", text: {}, para: {}, cell: {} };
    st.parent = at(s, NS.style, "parent-style-name");
    st.displayName = at(s, NS.style, "display-name") ?? at(s, NS.style, "name");
    st.listStyle = at(s, NS.style, "list-style-name");
    for (const p of kids(s)) {
      const target = is(p, NS.style, "text-properties") ? st.text : is(p, NS.style, "paragraph-properties") ? st.para : is(p, NS.style, "table-cell-properties") || is(p, NS.style, "table-column-properties") ? st.cell : undefined;
      if (!target) continue;
      for (let i = 0; i < p.attributes.length; i++) { const a = p.attributes[i]; target[`${a.prefix}:${a.localName}`] = a.value; }
    }
    if (isDefault) defaults.set(st.family, st);
    else { const n = at(s, NS.style, "name"); if (n) into.set(n, st); }
  }
}

export function readOdt(bytes: Uint8Array): { doc: PrimeDocument; warnings: string[] } {
  const files = unzipSync(bytes);
  const xml = (name: string) => files[name] ? new DOMParser().parseFromString(new TextDecoder().decode(files[name]), "application/xml").documentElement as unknown as El : undefined;
  const content = xml("content.xml");
  if (!content) throw new Error("This is not an OpenDocument text file (content.xml is missing).");
  const stylesDoc = xml("styles.xml");
  const doc = newDocument();
  const warnings = new Set<string>();
  const styles = new Map<string, OdfStyle>();
  const defaults = new Map<string, OdfStyle>();
  const listStyles = new Map<string, El>();
  const fontFaces = new Map<string, string>();
  for (const root of [stylesDoc, content]) {
    for (const k of root ? kids(root) : []) {
      if (is(k, NS.office, "font-face-decls")) for (const f of kids(k)) { const n = at(f, NS.style, "name"); const fam = at(f, NS.svg, "font-family")?.replace(/^['"]|['"]$/g, ""); if (n) fontFaces.set(n, fam ?? n); }
      if (is(k, NS.office, "styles") || is(k, NS.office, "automatic-styles")) readStyles(k, styles, defaults, listStyles);
    }
  }
  // Page size and margins (the default master page's layout).
  const layouts = stylesDoc ? kids(stylesDoc).filter((k) => is(k, NS.office, "automatic-styles")).flatMap(kids).filter((k) => is(k, NS.style, "page-layout")) : [];
  const lp = layouts.flatMap(kids).find((k) => is(k, NS.style, "page-layout-properties"));
  if (lp) {
    const s = doc.sections[0].props;
    const w = toPt(at(lp, NS.fo, "page-width"));
    const h = toPt(at(lp, NS.fo, "page-height"));
    if (w) s.pageWidth = w;
    if (h) s.pageHeight = h;
    if (at(lp, NS.style, "print-orientation") === "landscape") s.orientation = "landscape";
    s.margins = { ...s.margins };
    for (const side of ["top", "bottom", "left", "right"] as const) { const v = toPt(at(lp, NS.fo, `margin-${side}`)); if (v !== undefined) s.margins[side] = v; }
  }
  // Defaults: the document's default paragraph style's font and size; no space between paragraphs unless styled.
  const dp = defaults.get("paragraph");
  const defaultFont = dp?.text["style:font-name"] ? fontFaces.get(dp.text["style:font-name"]) : undefined;
  doc.styles.defaults.run = { ...doc.styles.defaults.run, ...(defaultFont ? { font: { ascii: defaultFont, hAnsi: defaultFont, cs: defaultFont, eastAsia: defaultFont } } : {}), size: toPt(dp?.text["fo:font-size"]) ?? 12 };
  doc.styles.defaults.paragraph = { spacing: { after: 0, line: 240, lineRule: "auto" } };

  /** A style's properties of one kind, its parents' under its own. */
  const chain = (name: string | undefined, kind: "text" | "para" | "cell", family: string): Record<string, string> => {
    const list: Record<string, string>[] = [];
    const seen = new Set<string>();
    for (let n = name; n && !seen.has(n);) { seen.add(n); const s = styles.get(n); if (!s) break; list.unshift(s[kind]); n = s.parent; }
    const d = defaults.get(family)?.[kind] ?? {};
    return Object.assign({}, d, ...list);
  };
  const styleChainName = (name: string | undefined): string | undefined => {
    const seen = new Set<string>();
    for (let n = name; n && !seen.has(n);) { seen.add(n); const s = styles.get(n); if (!s) return undefined; if (s.displayName && !/^P\d+$/.test(s.displayName)) return s.displayName; n = s.parent; }
    return undefined;
  };
  const runOf = (t: Record<string, string>, base: RunProps = {}): RunProps => {
    const p: RunProps = { ...base };
    if (t["fo:font-weight"]) p.bold = t["fo:font-weight"] === "bold" || Number(t["fo:font-weight"]) >= 600;
    if (t["fo:font-style"]) p.italic = t["fo:font-style"] === "italic" || t["fo:font-style"] === "oblique";
    const ul = t["style:text-underline-style"];
    if (ul) { if (ul === "none") delete p.underline; else p.underline = { val: t["style:text-underline-type"] === "double" ? "double" : ul === "dotted" ? "dotted" : ul === "wave" ? "wave" : "single" }; }
    const lt = t["style:text-line-through-style"];
    if (lt) p.strike = lt !== "none";
    const size = t["fo:font-size"];
    if (size && !size.endsWith("%")) { const v = toPt(size); if (v) p.size = v; }
    const col = t["fo:color"];
    if (col && /^#[0-9a-f]{6}$/i.test(col)) p.color = col.slice(1).toUpperCase();
    const font = t["style:font-name"] ? fontFaces.get(t["style:font-name"]) ?? t["style:font-name"] : t["fo:font-family"];
    if (font) p.font = { ascii: font.replace(/'/g, ""), hAnsi: font.replace(/'/g, ""), cs: font.replace(/'/g, "") };
    const pos = t["style:text-position"];
    if (pos) { if (/^super|^\d/.test(pos) && !/^0/.test(pos)) p.vertAlign = "superscript"; else if (/^sub|^-/.test(pos)) p.vertAlign = "subscript"; }
    if (t["fo:font-variant"] === "small-caps") p.smallCaps = true;
    if (t["fo:text-transform"] === "uppercase") p.caps = true;
    if (t["text:display"] === "none") p.hidden = true;
    return p;
  };
  const paraOf = (pp: Record<string, string>): ParagraphProps => {
    const p: ParagraphProps = {};
    const align = pp["fo:text-align"];
    if (align) { const a = { center: "center", end: "right", right: "right", justify: "both" }[align]; if (a) p.align = a; }
    const left = toPt(pp["fo:margin-left"]);
    const right = toPt(pp["fo:margin-right"]);
    const indent = toPt(pp["fo:text-indent"]);
    if (left || right || indent) {
      p.indent = {};
      if (left) p.indent.left = left;
      if (right) p.indent.right = right;
      if (indent && indent > 0) p.indent.firstLine = indent;
      if (indent && indent < 0) p.indent.hanging = -indent;
    }
    const before = toPt(pp["fo:margin-top"]);
    const after = toPt(pp["fo:margin-bottom"]);
    const lh = pp["fo:line-height"];
    if (before || after || lh) {
      p.spacing = {};
      if (before) p.spacing.before = before;
      if (after) p.spacing.after = after;
      if (lh?.endsWith("%")) { p.spacing.line = Math.round((Number(lh.slice(0, -1)) / 100) * 240); p.spacing.lineRule = "auto"; }
      else if (lh && toPt(lh)) { p.spacing.line = Math.round(toPt(lh)! * 20); p.spacing.lineRule = "exact"; }
    }
    if (pp["fo:keep-with-next"] === "always") p.keepNext = true;
    if (pp["fo:break-before"] === "page") p.pageBreakBefore = true;
    return p;
  };

  const blocks: Block[] = [];
  const inlinesOf = (e: El, run: RunProps, out: Inline[], family = "paragraph") => {
    for (let c = e.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) {
        const text = (c.nodeValue ?? "").replace(/[\r\n\t ]+/g, " ");
        if (!text) continue;
        const last = out[out.length - 1];
        if (last?.type === "text" && JSON.stringify(last.props ?? {}) === JSON.stringify(run)) last.text += text;
        else out.push({ type: "text", text, ...(Object.keys(run).length ? { props: { ...run } } : {}) });
        continue;
      }
      if (c.nodeType !== 1) continue;
      const k = c as El;
      if (is(k, NS.text, "span")) inlinesOf(k, runOf(chain(at(k, NS.text, "style-name"), "text", "text"), run), out, family);
      else if (is(k, NS.text, "s")) out.push({ type: "text", text: " ".repeat(Number(at(k, NS.text, "c") ?? 1)), ...(Object.keys(run).length ? { props: { ...run } } : {}) });
      else if (is(k, NS.text, "tab")) out.push({ type: "tab" });
      else if (is(k, NS.text, "line-break")) out.push({ type: "break", kind: "line" });
      else if (is(k, NS.text, "a")) {
        const content: Inline[] = [];
        inlinesOf(k, runOf(chain(at(k, NS.text, "style-name"), "text", "text"), run), content, family);
        const href = at(k, NS.xlink, "href") ?? "";
        out.push({ type: "link", ...(href.startsWith("#") ? { anchor: href.slice(1) } : { href }), content });
      } else if (is(k, NS.draw, "frame")) {
        const img = kids(k).find((x) => is(x, NS.draw, "image"));
        const href = at(img, NS.xlink, "href");
        const data = href ? files[href.replace(/^\.\//, "")] : undefined;
        if (data) {
          const type = /\.png$/i.test(href!) ? "image/png" : /\.jpe?g$/i.test(href!) ? "image/jpeg" : /\.gif$/i.test(href!) ? "image/gif" : /\.bmp$/i.test(href!) ? "image/bmp" : undefined;
          if (type) {
            const id = addImage(doc, data, type);
            out.push({ type: "image", imageId: id, width: toPt(at(k, NS.svg, "width")) ?? 100, height: toPt(at(k, NS.svg, "height")) ?? 100 });
          } else warnings.add("Some pictures (not PNG, JPEG, GIF or BMP) were left out.");
        } else {
          const box = kids(k).find((x) => is(x, NS.draw, "text-box"));
          if (box) warnings.add("Text boxes were placed in the text.");
          if (box) for (const p of kids(box)) if (is(p, NS.text, "p")) inlinesOf(p, run, out, family);
        }
      } else if (is(k, NS.text, "note")) warnings.add("Footnotes and endnotes were left out.");
      else if (is(k, NS.text, "soft-page-break") || is(k, NS.text, "bookmark") || is(k, NS.text, "bookmark-start") || is(k, NS.text, "bookmark-end") || is(k, NS.office, "annotation") || is(k, NS.office, "annotation-end")) continue;
      else inlinesOf(k, run, out, family); // fields and the like: their text
    }
  };
  const paragraphOf = (e: El, list?: { kind: "bullet" | "number"; level: number }): Paragraph => {
    const name = at(e, NS.text, "style-name");
    const props = paraOf(chain(name, "para", "paragraph"));
    const run = runOf(chain(name, "text", "paragraph"));
    // The defaults' size and font are the document's own: not repeated on each run.
    if (run.size === doc.styles.defaults.run?.size) delete run.size;
    if (run.font && run.font.ascii === defaultFont) delete run.font;
    const displayName = styleChainName(name) ?? "";
    let level: number | undefined;
    if (is(e, NS.text, "h")) level = Number(at(e, NS.text, "outline-level") ?? 1);
    else { const m = /^heading[ _](?:20_)?(\d)$/i.exec(displayName); if (m) level = Number(m[1]); }
    if (level && level >= 1 && level <= 9) props.style = headingStyle(doc, level);
    else if (/^title$/i.test(displayName)) props.style = "Title";
    if (list) { props.list = { numId: listId(doc, list.kind), level: Math.min(8, list.level) }; delete props.indent; }
    const content: Inline[] = [];
    inlinesOf(e, run, content);
    return { type: "paragraph", props, content };
  };
  const listKind = (styleName: string | undefined, level: number): "bullet" | "number" => {
    const ls = styleName ? listStyles.get(styleName) : undefined;
    const lvl = ls ? kids(ls).find((x) => Number(at(x, NS.text, "level")) === level + 1) : undefined;
    return lvl && is(lvl, NS.text, "list-level-style-number") && at(lvl, NS.style, "num-format") ? "number" : "bullet";
  };
  const tableOf = (e: El): Table => {
    const grid: number[] = [];
    const rows: Table["rows"] = [];
    const walkRows = (parent: El, header: boolean) => {
      for (const k of kids(parent)) {
        if (is(k, NS.table, "table-column")) {
          const w = toPt(chain(at(k, NS.table, "style-name"), "cell", "table-column")["style:column-width"]) ?? 72;
          for (let i = 0; i < Number(at(k, NS.table, "number-columns-repeated") ?? 1); i++) grid.push(w);
        } else if (is(k, NS.table, "table-columns") || is(k, NS.table, "table-column-group")) walkRows(k, header);
        else if (is(k, NS.table, "table-header-rows")) walkRows(k, true);
        else if (is(k, NS.table, "table-rows") || is(k, NS.table, "table-row-group")) walkRows(k, header);
        else if (is(k, NS.table, "table-row")) {
          const cells: TableCell[] = [];
          for (const c of kids(k)) {
            if (!is(c, NS.table, "table-cell")) continue;
            const span = Number(at(c, NS.table, "number-columns-spanned") ?? 1);
            const cellBlocks: Block[] = [];
            body(c, cellBlocks);
            if (!cellBlocks.length) cellBlocks.push({ type: "paragraph", props: {}, content: [] });
            const bg = chain(at(c, NS.table, "style-name"), "cell", "table-cell")["fo:background-color"];
            cells.push({ props: { ...(span > 1 ? { span } : {}), ...(bg && /^#[0-9a-f]{6}$/i.test(bg) ? { shading: { val: "clear", color: "auto", fill: bg.slice(1).toUpperCase() } } : {}) }, blocks: cellBlocks });
          }
          if (cells.length) rows.push({ props: header ? { header: true } : {}, cells });
        }
      }
    };
    walkRows(e, false);
    const cols = Math.max(grid.length, ...rows.map((r) => r.cells.reduce((a, c) => a + (c.props.span ?? 1), 0)));
    while (grid.length < cols) grid.push(72);
    return { type: "table", props: { style: "TableGrid" }, grid: grid.slice(0, cols), rows };
  };
  const body = (parent: El, out: Block[], list?: { style?: string; level: number }) => {
    for (const k of kids(parent)) {
      if (is(k, NS.text, "p") || is(k, NS.text, "h")) out.push(paragraphOf(k, list ? { kind: listKind(list.style, list.level), level: list.level } : undefined));
      else if (is(k, NS.text, "list")) {
        const style = at(k, NS.text, "style-name") ?? list?.style;
        const level = list ? list.level + 1 : 0;
        for (const item of kids(k)) if (is(item, NS.text, "list-item") || is(item, NS.text, "list-header")) body(item, out, { style, level });
      } else if (is(k, NS.table, "table")) out.push(tableOf(k));
      else if (is(k, NS.text, "section") || is(k, NS.text, "index-body") || is(k, NS.text, "table-of-content") || is(k, NS.text, "alphabetical-index") || is(k, NS.text, "bibliography")) body(k, out, list);
      else if (is(k, NS.text, "sequence-decls") || is(k, NS.text, "variable-decls") || is(k, NS.text, "user-field-decls") || is(k, NS.office, "forms") || is(k, NS.text, "index-title-template")) continue;
    }
  };
  const officeBody = kids(content).find((k) => is(k, NS.office, "body"));
  const officeText = officeBody ? kids(officeBody).find((k) => is(k, NS.office, "text")) : undefined;
  if (!officeText) throw new Error("This OpenDocument file holds no text document (perhaps a spreadsheet or a presentation).");
  body(officeText, blocks);
  doc.sections[0].blocks = blocks.length ? blocks : [{ type: "paragraph", props: {}, content: [] }];
  // Title and author.
  const meta = xml("meta.xml");
  const metaEl = meta ? kids(meta).find((k) => is(k, NS.office, "meta")) : undefined;
  const metaText = (ns: string, local: string) => (metaEl ? kids(metaEl).find((k) => is(k, ns, local))?.textContent?.trim() : undefined);
  const title = metaText(NS.dc, "title");
  const author = metaText(NS.meta, "initial-creator") ?? metaText(NS.dc, "creator");
  return { doc: title || author ? withProperties(doc, { ...(title ? { title } : {}), ...(author ? { author } : {}) }) : doc, warnings: [...warnings] };
}

// --- writing -----------------------------------------------------------------------------------------

const x = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const cm = (pt: number) => `${(pt * 2.54 / 72).toFixed(3)}cm`;

/** Text as ODF: runs of spaces as text:s, tabs and breaks as their elements. */
function odfText(s: string): string {
  return x(s).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`).replace(/^ /, '<text:s/>').replace(/\t/g, "<text:tab/>").replace(/\n/g, "<text:line-break/>");
}

export function writeOdt(doc: PrimeDocument, props: { title?: string; author?: string } = {}): Uint8Array {
  const body = shownBody(doc);
  const autoStyles: string[] = [];
  const textStyles = new Map<string, string>();
  const paraStyles = new Map<string, string>();
  const fontNames = new Set<string>();
  const pictures: Record<string, Uint8Array> = {};
  const textStyle = (r: Piece["run"]): string => {
    fontNames.add(r.font);
    let a = `style:font-name="${x(r.font)}" fo:font-size="${r.size}pt"`;
    if (r.bold) a += ' fo:font-weight="bold"';
    if (r.italic) a += ' fo:font-style="italic"';
    if (r.underline) a += ` style:text-underline-style="${r.underlineStyle === "dotted" ? "dotted" : r.underlineStyle === "wave" ? "wave" : "solid"}" style:text-underline-width="auto" style:text-underline-color="font-color"${r.underlineStyle === "double" ? ' style:text-underline-type="double"' : ""}`;
    if (r.strike) a += ' style:text-line-through-style="solid"';
    if (r.color && r.color !== "auto") a += ` fo:color="#${r.color}"`;
    if (r.vertAlign === "superscript") a += ' style:text-position="super 58%"';
    if (r.vertAlign === "subscript") a += ' style:text-position="sub 58%"';
    if (r.caps) a += ' fo:text-transform="uppercase"';
    if (r.hidden) a += ' text:display="none"';
    let n = textStyles.get(a);
    if (!n) { n = `T${textStyles.size + 1}`; textStyles.set(a, n); autoStyles.push(`<style:style style:name="${n}" style:family="text"><style:text-properties ${a}/></style:style>`); }
    return n;
  };
  const paraStyle = (sp: ShownParagraph, breakBefore: boolean): string => {
    const p = sp.p;
    let a = "";
    const align = { center: "center", right: "end", both: "justify", justify: "justify" }[p.align];
    if (align) a += ` fo:text-align="${align}"`;
    if (p.left) a += ` fo:margin-left="${cm(p.left)}"`;
    if (p.right) a += ` fo:margin-right="${cm(p.right)}"`;
    if (p.firstLine) a += ` fo:text-indent="${cm(p.firstLine)}"`;
    a += ` fo:margin-top="${cm(p.spaceBefore)}" fo:margin-bottom="${cm(p.spaceAfter)}"`;
    if (p.lineRule === "auto" && p.line) a += ` fo:line-height="${Math.round((p.line / 240) * 100)}%"`;
    else if (p.lineRule === "exact" && p.line) a += ` fo:line-height="${cm(p.line / 20)}"`;
    else if (p.lineRule === "atLeast" && p.line) a += ` style:line-height-at-least="${cm(p.line / 20)}"`;
    if (p.keepNext) a += ' fo:keep-with-next="always"';
    if (breakBefore || p.pageBreakBefore) a += ' fo:break-before="page"';
    // The tab stops in use (a list label is followed by one at the indent).
    const tabs = p.tabs.length ? `<style:tab-stops>${p.tabs.map((t) => `<style:tab-stop style:position="${cm(t.pos - (p.left || 0))}"${t.align === "right" ? ' style:type="right"' : t.align === "center" ? ' style:type="center"' : ""}/>`).join("")}</style:tab-stops>` : "";
    const mark = p.mark;
    fontNames.add(mark.font);
    const key = `${a}|${tabs}|${mark.font}|${mark.size}|${sp.heading ?? ""}`;
    let n = paraStyles.get(key);
    if (!n) {
      n = `P${paraStyles.size + 1}`;
      paraStyles.set(key, n);
      autoStyles.push(`<style:style style:name="${n}" style:family="paragraph"${sp.heading ? ` style:parent-style-name="Heading_20_${sp.heading}"` : ""}><style:paragraph-properties${a}>${tabs}</style:paragraph-properties><style:text-properties style:font-name="${x(mark.font)}" fo:font-size="${mark.size}pt"/></style:style>`);
    }
    return n;
  };
  let picN = 0;
  const pieces = (list: Piece[]): string => {
    let out = "";
    for (const piece of list) {
      if (piece.kind === "page" || piece.kind === "column") continue; // handled by splitting the paragraph
      const t = textStyle(piece.run);
      if (piece.kind === "text") {
        const span = `<text:span text:style-name="${t}">${odfText(piece.text)}</text:span>`;
        out += piece.link ? `<text:a xlink:type="simple" xlink:href="${x(piece.link)}">${span}</text:a>` : span;
      } else if (piece.kind === "tab") out += `<text:span text:style-name="${t}"><text:tab/></text:span>`;
      else if (piece.kind === "line") out += "<text:line-break/>";
      else if (piece.kind === "image") {
        const img = imageBytes(doc, piece.imageId);
        if (!img) continue;
        const ext = img.contentType.split("/")[1]?.replace("jpeg", "jpg") ?? "png";
        const name = `Pictures/image${++picN}.${ext}`;
        pictures[name] = img.data;
        out += `<draw:frame draw:name="Picture ${picN}" text:anchor-type="as-char" svg:width="${cm(piece.width)}" svg:height="${cm(piece.height)}"><draw:image xlink:href="${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
      }
    }
    return out;
  };
  const paragraphXml = (sp: ShownParagraph): string => {
    // A page break inside a paragraph: the rest goes on a new page (as a paragraph that starts one).
    const parts: Piece[][] = [[]];
    for (const p of sp.pieces) { if (p.kind === "page") parts.push([]); else parts[parts.length - 1].push(p); }
    return parts.map((part, i) => {
      const label = i === 0 && sp.p.listLabel ? `<text:span text:style-name="${textStyle(sp.p.labelRun ?? sp.p.mark)}">${odfText(sp.p.listLabel)}<text:tab/></text:span>` : "";
      const tag = sp.heading ? "text:h" : "text:p";
      const level = sp.heading ? ` text:outline-level="${sp.heading}"` : "";
      return `<${tag} text:style-name="${paraStyle(sp, i > 0)}"${level}>${label}${pieces(part)}</${tag}>`;
    }).join("");
  };
  let tableN = 0;
  const blocks = (list: Shown[]): string => list.map((b) => {
    if (b.kind === "paragraph") return paragraphXml(b);
    const t = ++tableN;
    const total = b.grid.reduce((a, v) => a + v, 0);
    autoStyles.push(`<style:style style:name="Table${t}" style:family="table"><style:table-properties style:width="${cm(total)}" table:align="left"/></style:style>`);
    b.grid.forEach((w, i) => autoStyles.push(`<style:style style:name="Table${t}.C${i}" style:family="table-column"><style:table-column-properties style:column-width="${cm(w)}"/></style:style>`));
    const cellStyle = (shading?: string) => {
      const n = `Table${t}.Cell${shading ?? ""}`;
      if (!autoStyles.some((s) => s.includes(`style:name="${n}"`))) autoStyles.push(`<style:style style:name="${n}" style:family="table-cell"><style:table-cell-properties fo:padding="0.097cm" fo:border="0.5pt solid #000000"${shading ? ` fo:background-color="#${shading}"` : ""}/></style:style>`);
      return n;
    };
    const cols = b.grid.map((_, i) => `<table:table-column table:style-name="Table${t}.C${i}"/>`).join("");
    const rows = b.rows.map((r) => `<table:table-row>${r.cells.map((c) => {
      const inner = blocks(c.blocks) || "<text:p/>";
      const span = c.span > 1 ? ` table:number-columns-spanned="${c.span}"` : "";
      return `<table:table-cell table:style-name="${cellStyle(c.shading)}" office:value-type="string"${span}>${inner}</table:table-cell>${"<table:covered-table-cell/>".repeat(c.span - 1)}`;
    }).join("")}</table:table-row>`).join("");
    return `<table:table table:name="Table${t}" table:style-name="Table${t}">${cols}${rows}</table:table>`;
  }).join("");
  const text = blocks(body);

  const ns = `xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:text="${NS.text}" xmlns:table="${NS.table}" xmlns:draw="${NS.draw}" xmlns:fo="${NS.fo}" xmlns:xlink="${NS.xlink}" xmlns:dc="${NS.dc}" xmlns:meta="${NS.meta}" xmlns:svg="${NS.svg}"`;
  const fonts = `<office:font-face-decls>${[...fontNames].map((f) => `<style:font-face style:name="${x(f)}" svg:font-family="${x(/\s/.test(f) ? `'${f}'` : f)}"/>`).join("")}</office:font-face-decls>`;
  const contentXml = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ns} office:version="1.3">${fonts}<office:automatic-styles>${autoStyles.join("")}</office:automatic-styles><office:body><office:text>${text}</office:text></office:body></office:document-content>`;
  const s = doc.sections[0].props;
  const m = s.margins ?? {};
  const headingStyles = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((l) => `<style:style style:name="Heading_20_${l}" style:display-name="Heading ${l}" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="${l}"/>`).join("");
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${ns} office:version="1.3">${fonts}<office:styles><style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"/>${headingStyles}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="${cm(s.pageWidth ?? 612)}" fo:page-height="${cm(s.pageHeight ?? 792)}" style:print-orientation="${s.orientation === "landscape" ? "landscape" : "portrait"}" fo:margin-top="${cm(m.top ?? 72)}" fo:margin-bottom="${cm(m.bottom ?? 72)}" fo:margin-left="${cm(m.left ?? 72)}" fo:margin-right="${cm(m.right ?? 72)}"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`;
  const metaXml = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${ns} office:version="1.3"><office:meta><meta:generator>Prime Word</meta:generator>${props.title ? `<dc:title>${x(props.title)}</dc:title>` : ""}${props.author ? `<meta:initial-creator>${x(props.author)}</meta:initial-creator><dc:creator>${x(props.author)}</dc:creator>` : ""}</office:meta></office:document-meta>`;
  const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${Object.keys(pictures).map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${p.endsWith(".png") ? "image/png" : p.endsWith(".jpg") ? "image/jpeg" : "image/" + p.split(".").pop()}"/>`).join("")}</manifest:manifest>`;
  // The mimetype first and stored (not compressed), as the format requires.
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "META-INF/manifest.xml": strToU8(manifest),
    "content.xml": strToU8(contentXml),
    "styles.xml": strToU8(stylesXml),
    "meta.xml": strToU8(metaXml),
    ...Object.fromEntries(Object.entries(pictures).map(([k, v]) => [k, [v, { level: 0 }] as [Uint8Array, { level: 0 }]])),
  });
}
