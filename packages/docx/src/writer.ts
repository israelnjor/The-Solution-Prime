/** Prime Document Model -> DOCX. Emits clean OOXML in schema order; preserved content is written back verbatim. */
import { zipSync } from "fflate";
import type { Block, FloatInfo, FloatPosition, HeaderFooterKind, Inline, Paragraph, PrimeDocument, Relationship, Section, Story, Table } from "@prime/pdm";
import { ContentType, EMU_PER_PT, RelType, twips } from "./ooxml.js";
import { writeCellProps, writeParagraphProps, writeRowProps, writeRunProps, writeSectionProps, writeTableProps } from "./props.js";
import { writeNumbering, writeStyles } from "./styles.js";
import { NS, ROOT_NAMESPACES, attrString, el, escAttr, escText, fromBase64, type Item } from "./xml.js";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const encoder = new TextEncoder();
const dirname = (p: string) => (p.includes("/") ? p.slice(0, p.lastIndexOf("/")) : "");
const relsPath = (part: string) => `${dirname(part) ? dirname(part) + "/" : ""}_rels/${part.split("/").pop()}.rels`;

function relative(fromPart: string, to: string): string {
  const from = dirname(fromPart).split("/").filter(Boolean);
  const target = to.split("/");
  let i = 0;
  while (i < from.length && i < target.length - 1 && from[i] === target[i]) i++;
  return [...Array(from.length - i).fill(".."), ...target.slice(i)].join("/");
}

/** Relationships of one part. Preserved relationships keep their original ids; new ones get fresh ids. */
class Rels {
  private list: Relationship[] = [];
  private ids = new Set<string>();
  private cache = new Map<string, string>();
  private n = 0;

  constructor(readonly part: string, preserved: Relationship[] = []) {
    for (const r of preserved) {
      this.list.push(r);
      this.ids.add(r.id);
    }
  }

  add(type: string, target: string, external = false): string {
    const key = `${type}|${target}|${external}`;
    const cached = this.cache.get(key);
    if (cached) return cached;
    let id: string;
    do id = `rIdP${++this.n}`; while (this.ids.has(id));
    this.ids.add(id);
    this.list.push({ id, type, target, ...(external ? { external: true } : {}) });
    if (type !== RelType.hyperlink) this.cache.set(key, id);
    return id;
  }

  get size() {
    return this.list.length;
  }

  xml(): string {
    const rows = this.list.map((r) => {
      const target = r.external ? r.target : relative(this.part, r.target);
      return `<Relationship Id="${escAttr(r.id)}" Type="${escAttr(r.type)}" Target="${escAttr(target)}"${r.external ? ' TargetMode="External"' : ""}/>`;
    });
    return `${XML_DECL}<Relationships xmlns="${NS.pr}">${rows.join("")}</Relationships>`;
  }
}

class Writer {
  private files: Record<string, Uint8Array> = {};
  private overrides: Record<string, string> = {};
  private defaults: Record<string, string>;
  private used = new Set<string>();
  private drawingId = 0;

  constructor(private doc: PrimeDocument) {
    this.defaults = { ...doc.package.contentTypes.defaults };
  }

  write(): Uint8Array {
    const pkg = this.doc.package;
    for (const p of pkg.parts) this.put(p.name, fromBase64(p.data));
    const main = pkg.mainPart;
    const rels = new Rels(main, pkg.preservedRels);

    const styles = this.unique("word/styles", "xml");
    this.put(styles, encoder.encode(writeStyles(this.doc.styles)), ContentType.styles);
    rels.add(RelType.styles, styles);
    const lists = this.doc.lists;
    if (lists.definitions.length || lists.instances.length || lists.extra?.length) {
      const numbering = this.unique("word/numbering", "xml");
      this.put(numbering, encoder.encode(writeNumbering(lists)), ContentType.numbering);
      rels.add(RelType.numbering, numbering);
    }

    const sections = this.doc.sections;
    const body = sections.map((s, i) => this.section(s, i === sections.length - 1, rels)).join("");
    const rootAttrs = { ...ROOT_NAMESPACES, ...pkg.rootAttrs };
    this.put(main, encoder.encode(`${XML_DECL}<w:document${attrString(rootAttrs, "")}><w:body>${body}</w:body></w:document>`),
      pkg.contentTypes.overrides[main] ?? ContentType.document);
    this.put(relsPath(main), encoder.encode(rels.xml()));
    this.pruneComments(body);
    if (!this.files["_rels/.rels"]) {
      // A document created from scratch has no package relationships to carry through.
      const root = new Rels("");
      root.add(RelType.officeDocument, main);
      this.put("_rels/.rels", encoder.encode(root.xml()));
    }
    this.put("[Content_Types].xml", encoder.encode(this.contentTypes()));
    return zipSync(this.files, { level: 6 });
  }

  /**
   * Comments whose markers are no longer in the text (deleted, or undone after being added) are left out
   * of word/comments.xml, so Word does not show comments on nothing.
   */
  private pruneComments(body: string) {
    const file = this.files["word/comments.xml"];
    if (!file) return;
    const used = new Set([...body.matchAll(/<w:commentReference\b[^>]*w:id="([^"]*)"/g)].map((m) => m[1]));
    const xml = new TextDecoder().decode(file);
    const kept = xml.replace(/<w:comment\b[^>]*?w:id="([^"]*)"[\s\S]*?<\/w:comment>/g, (whole, id) => (used.has(id) ? whole : ""));
    if (kept !== xml) this.files["word/comments.xml"] = encoder.encode(kept);
  }

  private put(name: string, data: Uint8Array, contentType?: string) {
    this.files[name] = data;
    this.used.add(name);
    if (contentType) this.overrides[name] = contentType;
  }

  private unique(prefix: string, ext: string): string {
    let name = `${prefix}.${ext}`;
    for (let n = 1; this.used.has(name); n++) name = `${prefix}${n}.${ext}`;
    this.used.add(name);
    return name;
  }

  private contentTypes(): string {
    const d = { rels: ContentType.rels, xml: ContentType.xml, ...this.defaults };
    const original = this.doc.package.contentTypes.overrides;
    const rows = Object.entries(d).map(([ext, ct]) => `<Default Extension="${escAttr(ext)}" ContentType="${escAttr(ct)}"/>`);
    for (const name of Object.keys(this.files).sort()) {
      const ct = this.overrides[name] ?? original[name];
      if (ct) rows.push(`<Override PartName="/${escAttr(name)}" ContentType="${escAttr(ct)}"/>`);
    }
    return `${XML_DECL}<Types xmlns="${NS.ct}">${rows.join("")}</Types>`;
  }

  // --- sections and stories ----------------------------------------------------------------

  private section(s: Section, last: boolean, rels: Rels): string {
    const refs: Item[] = [];
    for (const [kind, story] of Object.entries(s.headers) as [HeaderFooterKind, Story][]) {
      refs.push(["w:headerReference", this.storyRef("header", kind, story, rels)]);
    }
    for (const [kind, story] of Object.entries(s.footers) as [HeaderFooterKind, Story][]) {
      refs.push(["w:footerReference", this.storyRef("footer", kind, story, rels)]);
    }
    const sectPr = writeSectionProps(s.props, refs);
    if (last) return this.blocks(s.blocks, rels) + sectPr;
    const tail = s.blocks[s.blocks.length - 1];
    if (tail?.type === "paragraph") return this.blocks(s.blocks.slice(0, -1), rels) + this.paragraph(tail, rels, sectPr);
    return this.blocks(s.blocks, rels) + `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
  }

  private storyRef(kind: "header" | "footer", type: HeaderFooterKind, story: Story, rels: Rels): string {
    const part = this.unique(`word/${kind}`, "xml");
    const storyRels = new Rels(part, story.preservedRels);
    const tag = kind === "header" ? "w:hdr" : "w:ftr";
    const blocks = this.blocks(story.blocks, storyRels) || "<w:p/>";
    this.put(part, encoder.encode(`${XML_DECL}<${tag}${attrString(ROOT_NAMESPACES, "")}>${blocks}</${tag}>`),
      kind === "header" ? ContentType.header : ContentType.footer);
    if (storyRels.size) this.put(relsPath(part), encoder.encode(storyRels.xml()));
    const id = rels.add(kind === "header" ? RelType.header : RelType.footer, part);
    return `<w:${kind}Reference w:type="${type}" r:id="${id}"/>`;
  }

  // --- blocks ------------------------------------------------------------------------------

  private blocks(blocks: Block[], rels: Rels): string {
    return blocks.map((b) => {
      if (b.type === "paragraph") return this.paragraph(b, rels);
      if (b.type === "table") return this.table(b, rels);
      if (b.type === "contentControl") return this.contentControl(b.settings, this.blocks(b.blocks, rels));
      return b.preserved.xml;
    }).join("");
  }

  private paragraph(p: Paragraph, rels: Rels, sectPr = ""): string {
    return `<w:p>${writeParagraphProps(p.props, sectPr)}${this.inlines(p.content, rels)}</w:p>`;
  }

  private table(t: Table, rels: Rels): string {
    const grid = t.grid.map((w) => el("w:gridCol", { w: twips(w) })).join("");
    const rows = t.rows.map((row) => {
      const extra = (row.extra ?? []).map((x) => x.xml).join("");
      const cells = row.cells.map((cell) => {
        let inner = this.blocks(cell.blocks, rels);
        const last = cell.blocks[cell.blocks.length - 1]?.type;
        if (last !== "paragraph" && last !== "contentControl") inner += "<w:p/>";
        return `<w:tc>${writeCellProps(cell.props)}${inner}</w:tc>`;
      }).join("");
      return `<w:tr>${extra}${writeRowProps(row.props)}${cells}</w:tr>`;
    }).join("");
    return `<w:tbl>${writeTableProps(t.props)}<w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
  }

  // --- inline content ----------------------------------------------------------------------

  private contentControl(settings: { xml: string }[], inner: string): string {
    return `<w:sdt>${settings.map((s) => s.xml).join("")}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  }

  /** `deleted`: inside a tracked deletion, where text is written as w:delText. */
  private inlines(items: Inline[], rels: Rels, deleted = false): string {
    return items.map((i) => this.inline(i, rels, deleted)).join("");
  }

  private inline(i: Inline, rels: Rels, deleted = false): string {
    const run = (inner: string) => `<w:r>${"props" in i ? writeRunProps(i.props) : ""}${inner}</w:r>`;
    switch (i.type) {
      case "text":
        return run(deleted
          ? `<w:delText xml:space="preserve">${escText(i.text)}</w:delText>`
          : `<w:t xml:space="preserve">${escText(i.text)}</w:t>`);
      case "tab":
        return run("<w:tab/>");
      case "break": {
        const a: Record<string, string> = {};
        if (i.kind !== "line") a.type = i.kind;
        if (i.clear) a.clear = i.clear;
        return run(el("w:br", a));
      }
      case "fieldChar":
        return run(el("w:fldChar", { fldCharType: i.kind }));
      case "instr":
        return run(deleted
          ? `<w:delInstrText xml:space="preserve">${escText(i.text)}</w:delInstrText>`
          : `<w:instrText xml:space="preserve">${escText(i.text)}</w:instrText>`);
      case "image":
        return run(this.drawing(i, rels));
      case "runPreserved":
        return run(i.preserved.xml);
      case "preserved":
        return i.preserved.xml;
      case "field":
        return `<w:fldSimple w:instr="${escAttr(i.instr)}">${this.inlines(i.content, rels, deleted)}</w:fldSimple>`;
      case "contentControl":
        return this.contentControl(i.settings, this.inlines(i.content, rels, deleted));
      case "revision":
        return `<w:${i.kind}${attrString(i.attrs, "")}>${this.inlines(i.content, rels, deleted || i.kind === "del")}</w:${i.kind}>`;
      case "link": {
        let a = "";
        if (i.href !== undefined) a += ` r:id="${rels.add(RelType.hyperlink, i.href, true)}"`;
        if (i.anchor !== undefined) a += ` w:anchor="${escAttr(i.anchor)}"`;
        if (i.tooltip !== undefined) a += ` w:tooltip="${escAttr(i.tooltip)}"`;
        return `<w:hyperlink${a} w:history="1">${this.inlines(i.content, rels, deleted)}</w:hyperlink>`;
      }
    }
  }

  private drawing(img: Extract<Inline, { type: "image" }>, rels: Rels): string {
    const res = this.doc.images[img.imageId];
    if (!res) throw new Error(`Image resource missing: ${img.imageId}`);
    const part = img.imageId.includes("/") ? img.imageId : `word/media/${img.imageId}.${res.extension}`;
    if (!this.files[part]) this.put(part, fromBase64(res.data));
    this.defaults[res.extension] ??= res.contentType;
    const rId = rels.add(RelType.image, part);
    const cx = Math.round(img.width * EMU_PER_PT);
    const cy = Math.round(img.height * EMU_PER_PT);
    const id = ++this.drawingId;
    const name = escAttr(img.name ?? `Picture ${id}`);
    const descr = img.description !== undefined ? ` descr="${escAttr(img.description)}"` : "";
    const body = `<wp:extent cx="${cx}" cy="${cy}"/>`
      + `<wp:effectExtent l="0" t="0" r="0" b="0"/>${img.float ? wrapXml(img.float) : ""}<wp:docPr id="${id}" name="${name}"${descr}/>`
      + `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
      + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>`
      + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
      + `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
      + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
      + `</pic:pic></a:graphicData></a:graphic>`;
    if (!img.float) return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${body}</wp:inline></w:drawing>`;
    // A floating picture: positioned from its column and paragraph, with its wrapping.
    const f = img.float;
    const emu = (pt: number) => Math.round(pt * EMU_PER_PT);
    const position = (axis: "H" | "V", p: FloatPosition) => `<wp:position${axis} relativeFrom="${p.from}">`
      + (p.offset !== undefined || !p.align ? `<wp:posOffset>${emu(p.offset ?? 0)}</wp:posOffset>` : `<wp:align>${p.align}</wp:align>`)
      + `</wp:position${axis}>`;
    return `<w:drawing><wp:anchor distT="${emu(f.dist.top)}" distB="${emu(f.dist.bottom)}" distL="${emu(f.dist.left)}" distR="${emu(f.dist.right)}"`
      + ` simplePos="0" relativeHeight="${251659264 + id}" behindDoc="${f.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
      + `<wp:simplePos x="0" y="0"/>${position("H", f.x)}${position("V", f.y)}${body}</wp:anchor></w:drawing>`;
  }
}

/** The wrapping element of a floating picture (between wp:effectExtent and wp:docPr). */
function wrapXml(f: FloatInfo): string {
  const side = f.side ?? "bothSides";
  const polygon = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/>'
    + '<wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
  switch (f.wrap) {
    case "square": return `<wp:wrapSquare wrapText="${side}"/>`;
    case "tight": return `<wp:wrapTight wrapText="${side}">${polygon}</wp:wrapTight>`;
    case "through": return `<wp:wrapThrough wrapText="${side}">${polygon}</wp:wrapThrough>`;
    case "topAndBottom": return "<wp:wrapTopAndBottom/>";
    default: return "<wp:wrapNone/>";
  }
}

export function writeDocx(doc: PrimeDocument): Uint8Array {
  return new Writer(doc).write();
}
