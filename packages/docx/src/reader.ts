/** DOCX -> Prime Document Model. */
import { unzipSync } from "fflate";
import {
  type Block, type FloatInfo, type FloatPosition, type HeaderFooterKind, type ImageInline, type Inline, type Notes, type Paragraph, type PrimeDocument,
  type Relationship, SCHEMA_VERSION, type Section, type Story, type Table, type TableRow,
} from "@prime/pdm";
import { readInk } from "./inkml.js";
import { IMAGE_TYPES, MODELED_REL_TYPES, RelType, EMU_PER_PT } from "./ooxml.js";
import { readCellProps, readParagraphProps, readRowProps, readRunProps, readSectionProps, readTableProps } from "./props.js";
import { readNumbering, readStyles } from "./styles.js";
import { NS, attr, attrs, child, elements, parseXml, preserve, qname, textContent, toBase64 } from "./xml.js";

export interface ReadResult {
  doc: PrimeDocument;
  /** Human-readable notes about content that is kept but not editable, or could not be kept. */
  warnings: string[];
}

/** Paragraph-level elements that carry no content and are dropped (proofing/render caches). */
const DROPPED = new Set(["w:proofErr", "w:lastRenderedPageBreak", "w:noProof"]);
/** Preserved elements that are harmless to keep silently (no visible content, or editable elsewhere: comments are in the Comments pane). */
const SILENT = new Set(["w:bookmarkStart", "w:bookmarkEnd", "w:permStart", "w:permEnd", "w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"]);
/** User-facing names for preserved content, used in conversion warnings. */
const FRIENDLY: Record<string, string> = {
  "w:drawing": "Floating image or shape",
  "w:drawing/inline": "Picture with border, effects or link",
  "w:pict": "Legacy drawing",
  "w:object": "Embedded object",
  "mc:AlternateContent": "Shape or text box",
  "w:footnoteReference": "Footnote",
  "w:endnoteReference": "Endnote",
  "w:commentRangeStart": "Comment",
  "w:commentRangeEnd": "Comment",
  "w:commentReference": "Comment",
  "w:ins": "Tracked insertion",
  "w:del": "Tracked deletion",
  "w:moveFrom": "Tracked move",
  "w:moveTo": "Tracked move",
  "w:sdt": "Content control",
  "m:oMath": "Equation (displayed; editing equations is not supported yet)",
  "m:oMathPara": "Equation (displayed; editing equations is not supported yet)",
  "w:hyperlink": "Hyperlink with unsupported options",
};
const REL_REF =/\br:(?:id|embed|link|pict|dm|lo|qs|cs|href)="([^"]+)"/g;

const dirname = (p: string) => (p.includes("/") ? p.slice(0, p.lastIndexOf("/")) : "");

function resolve(base: string, target: string): string {
  const parts = target.startsWith("/") ? [] : base ? base.split("/") : [];
  for (const seg of target.replace(/^\//, "").split("/")) {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

const relsPath = (part: string) => (part ? `${dirname(part) ? dirname(part) + "/" : ""}_rels/${part.split("/").pop()}.rels` : "_rels/.rels");

/** State for one story (main document, a header, a footer): its part, relationships, and what we used. */
class StoryContext {
  consumed = new Set<string>();
  referenced = new Set<string>();
  constructor(readonly part: string, readonly rels: Map<string, Relationship>) {}

  keep(e: Element) {
    const p = preserve(e);
    for (const m of p.xml.matchAll(REL_REF)) this.referenced.add(m[1]);
    return p;
  }

  /** Relationships to carry through: anything not modeled, plus anything referenced from preserved XML. */
  preservedRels(): Relationship[] {
    return [...this.rels.values()].filter((r) =>
      this.referenced.has(r.id) ||
      (!this.consumed.has(r.id) && !MODELED_REL_TYPES.has(r.type) && r.type !== RelType.image && r.type !== RelType.hyperlink));
  }
}

class Reader {
  private files: Record<string, Uint8Array>;
  private decoder = new TextDecoder();
  private warnCounts = new Map<string, number>();
  private images: PrimeDocument["images"] = {};
  private modeledParts = new Set<string>(["[Content_Types].xml"]);
  private mediaUsedByPreserved = new Set<string>();
  private contentTypes = { defaults: {} as Record<string, string>, overrides: {} as Record<string, string> };

  constructor(data: Uint8Array) {
    this.files = unzipSync(data);
  }

  warn(message: string) {
    this.warnCounts.set(message, (this.warnCounts.get(message) ?? 0) + 1);
  }

  xml(part: string): Element | undefined {
    const bytes = this.files[part];
    return bytes ? parseXml(this.decoder.decode(bytes)).documentElement! : undefined;
  }

  rels(part: string): Map<string, Relationship> {
    const out = new Map<string, Relationship>();
    const root = this.xml(relsPath(part));
    for (const r of root ? elements(root) : []) {
      const external = r.getAttribute("TargetMode") === "External";
      const target = r.getAttribute("Target") ?? "";
      out.set(r.getAttribute("Id") ?? "", {
        id: r.getAttribute("Id") ?? "",
        type: r.getAttribute("Type") ?? "",
        target: external ? target : resolve(dirname(part), target),
        ...(external ? { external: true } : {}),
      });
    }
    return out;
  }

  read(): ReadResult {
    const ct = this.xml("[Content_Types].xml");
    for (const e of ct ? elements(ct) : []) {
      if (e.localName === "Default") this.contentTypes.defaults[e.getAttribute("Extension")!.toLowerCase()] = e.getAttribute("ContentType")!;
      if (e.localName === "Override") this.contentTypes.overrides[e.getAttribute("PartName")!.replace(/^\//, "")] = e.getAttribute("ContentType")!;
    }
    const main = [...this.rels("").values()].find((r) => r.type === RelType.officeDocument)?.target;
    if (!main || !this.files[main]) throw new Error("Not a Word document: no main document part.");

    const ctx = new StoryContext(main, this.rels(main));
    this.modeledParts.add(main).add(relsPath(main));
    const relOf = (type: string) => [...ctx.rels.values()].find((r) => r.type === type)?.target;

    const stylesPart = relOf(RelType.styles);
    const numberingPart = relOf(RelType.numbering);
    if (stylesPart) this.modeledParts.add(stylesPart).add(relsPath(stylesPart));
    if (numberingPart) this.modeledParts.add(numberingPart).add(relsPath(numberingPart));

    const root = this.xml(main)!;
    const body = child(root, "w:body");
    if (!body) throw new Error("Main document has no body.");
    const rootAttrs: Record<string, string> = {};
    for (let i = 0; i < root.attributes.length; i++) rootAttrs[root.attributes[i].name] = root.attributes[i].value;

    const doc: PrimeDocument = {
      schemaVersion: SCHEMA_VERSION,
      styles: readStyles(stylesPart ? this.xml(stylesPart) : undefined),
      lists: readNumbering(numberingPart ? this.xml(numberingPart) : undefined, (m) => this.warn(m)),
      sections: this.readBody(body, ctx),
      images: this.images,
      package: { mainPart: main, parts: [], preservedRels: [], contentTypes: this.contentTypes, rootAttrs },
    };
    for (const e of elements(root)) {
      if (qname(e) === "w:background") doc.background = ctx.keep(e); // page colour
      else if (qname(e) !== "w:body") this.warn(`Document-level ${qname(e)} is not supported and was dropped.`);
    }
    const footnotes = relOf(RelType.footnotes);
    const endnotes = relOf(RelType.endnotes);
    if (footnotes || endnotes) {
      doc.notes = {};
      if (footnotes) doc.notes.footnote = this.readNotes(footnotes, "w:footnote");
      if (endnotes) doc.notes.endnote = this.readNotes(endnotes, "w:endnote");
    }

    doc.package.preservedRels = ctx.preservedRels();
    for (const r of doc.package.preservedRels) if (!r.external) this.mediaUsedByPreserved.add(r.target);
    for (const [name, data] of Object.entries(this.files)) {
      if (name.endsWith("/")) continue;
      const modeledMedia = Object.prototype.hasOwnProperty.call(this.images, name) && !this.mediaUsedByPreserved.has(name);
      if (!this.modeledParts.has(name) && !modeledMedia) doc.package.parts.push({ name, data: toBase64(data) });
    }
    const warnings = [...this.warnCounts].map(([m, n]) => (n > 1 ? `${m} (${n}×)` : m));
    return { doc, warnings };
  }

  // --- body and sections -------------------------------------------------------------------

  readBody(body: Element, ctx: StoryContext): Section[] {
    const sections: Section[] = [];
    let blocks: Block[] = [];
    const close = (sectPr: Element | undefined) => {
      sections.push({ ...this.readSection(sectPr, ctx), blocks });
      blocks = [];
    };
    for (const e of elements(body)) {
      const n = qname(e);
      if (n === "w:sectPr") close(e);
      else if (n === "w:p") {
        const { paragraph, sectPr } = this.readParagraph(e, ctx);
        blocks.push(paragraph);
        if (sectPr) close(sectPr);
      } else blocks.push(this.readBlock(e, ctx));
    }
    if (blocks.length || !sections.length) close(undefined);
    return sections;
  }

  readSection(sectPr: Element | undefined, ctx: StoryContext): Omit<Section, "blocks"> {
    const section: Omit<Section, "blocks"> = { props: sectPr ? readSectionProps(sectPr) : {}, headers: {}, footers: {} };
    for (const ref of sectPr ? elements(sectPr) : []) {
      const n = qname(ref);
      if (n !== "w:headerReference" && n !== "w:footerReference") continue;
      const rel = ctx.rels.get(attr(ref, "id", NS.r) ?? "");
      if (!rel || rel.external) continue;
      ctx.consumed.add(rel.id);
      const kind = (attr(ref, "type") ?? "default") as HeaderFooterKind;
      (n === "w:headerReference" ? section.headers : section.footers)[kind] = this.readStory(rel.target);
    }
    return section;
  }

  readStory(part: string): Story {
    const ctx = new StoryContext(part, this.rels(part));
    this.modeledParts.add(part).add(relsPath(part));
    const root = this.xml(part);
    const blocks = root ? this.readBlocks(elements(root), ctx) : [];
    const preservedRels = ctx.preservedRels();
    for (const r of preservedRels) if (!r.external) this.mediaUsedByPreserved.add(r.target);
    return preservedRels.length ? { blocks, preservedRels } : { blocks };
  }

  /** A footnotes or endnotes part: each note's text; separator notes kept as they are. */
  readNotes(part: string, tag: "w:footnote" | "w:endnote"): Notes | undefined {
    const root = this.xml(part);
    if (!root) return undefined;
    const ctx = new StoryContext(part, this.rels(part));
    this.modeledParts.add(part).add(relsPath(part));
    const out: Notes = { items: [], special: [], part };
    const rootAttrs: Record<string, string> = {};
    for (let i = 0; i < root.attributes.length; i++) rootAttrs[root.attributes[i].name] = root.attributes[i].value;
    out.rootAttrs = rootAttrs;
    for (const e of elements(root)) {
      if (qname(e) !== tag) continue;
      if (attr(e, "type") && attr(e, "type") !== "normal") out.special.push(ctx.keep(e));
      else out.items.push({ id: attr(e, "id") ?? "", blocks: this.readBlocks(elements(e), ctx) });
    }
    const preservedRels = ctx.preservedRels();
    for (const r of preservedRels) if (!r.external) this.mediaUsedByPreserved.add(r.target);
    if (preservedRels.length) out.preservedRels = preservedRels;
    return out;
  }

  readBlocks(nodes: Element[], ctx: StoryContext): Block[] {
    return nodes.map((e) => (qname(e) === "w:p" ? this.readParagraph(e, ctx).paragraph : this.readBlock(e, ctx)));
  }

  readBlock(e: Element, ctx: StoryContext): Block {
    if (qname(e) === "w:tbl") return this.readTable(e, ctx);
    const content = qname(e) === "w:sdt" ? child(e, "w:sdtContent") : undefined;
    if (content) {
      this.warn("Content control: its content is editable; its settings are kept unchanged.");
      return {
        type: "contentControl",
        settings: elements(e).filter((c) => c !== content).map((c) => ctx.keep(c)),
        blocks: this.readBlocks(elements(content), ctx),
      };
    }
    this.notePreserved(qname(e));
    return { type: "preserved", preserved: ctx.keep(e) };
  }

  notePreserved(name: string) {
    if (!SILENT.has(name)) this.warn(`${FRIENDLY[name] ?? `Unsupported element ${name}`}: kept unchanged, not yet editable.`);
  }

  // --- paragraphs and inline content ---------------------------------------------------------

  readParagraph(p: Element, ctx: StoryContext): { paragraph: Paragraph; sectPr?: Element } {
    const { props, sectPr } = readParagraphProps(child(p, "w:pPr"));
    const content = this.readInlines(elements(p).filter((c) => qname(c) !== "w:pPr"), ctx);
    return { paragraph: { type: "paragraph", props, content: mergeText(content) }, sectPr };
  }

  readInlines(nodes: Element[], ctx: StoryContext): Inline[] {
    const out: Inline[] = [];
    for (const e of nodes) {
      const n = qname(e);
      if (n === "w:r") out.push(...this.readRun(e, ctx));
      else if (n === "w:hyperlink") out.push(this.readHyperlink(e, ctx));
      else if (n === "w:fldSimple" && Object.keys(attrs(e)).every((k) => k === "instr")) {
        out.push({ type: "field", instr: attr(e, "instr") ?? "", content: this.readInlines(elements(e), ctx) });
      } else if (n === "w:sdt" && child(e, "w:sdtContent")) {
        const content = child(e, "w:sdtContent")!;
        this.warn("Content control: its content is editable; its settings are kept unchanged.");
        out.push({
          type: "contentControl",
          settings: elements(e).filter((c) => c !== content).map((c) => ctx.keep(c)),
          content: this.readInlines(elements(content), ctx),
        });
      } else if (n === "w:ins" || n === "w:del") {
        this.warn("Tracked changes are shown and can be accepted or rejected; new edits are not tracked yet.");
        out.push({ type: "revision", kind: n === "w:ins" ? "ins" : "del", attrs: rawAttrs(e), content: this.readInlines(elements(e), ctx) });
      } else if (DROPPED.has(n)) continue;
      else {
        this.notePreserved(n);
        out.push({ type: "preserved", preserved: ctx.keep(e) });
      }
    }
    return out;
  }

  readHyperlink(e: Element, ctx: StoryContext): Inline {
    const a = attrs(e);
    const known = Object.keys(a).every((k) => ["r:id", "anchor", "tooltip", "history"].includes(k));
    const rel = ctx.rels.get(a["r:id"] ?? "");
    if (!known || (a["r:id"] && (!rel || rel.type !== RelType.hyperlink))) {
      this.notePreserved("w:hyperlink");
      return { type: "preserved", preserved: ctx.keep(e) };
    }
    if (rel) ctx.consumed.add(rel.id);
    return {
      type: "link",
      ...(rel ? { href: rel.target } : {}),
      ...(a.anchor !== undefined ? { anchor: a.anchor } : {}),
      ...(a.tooltip !== undefined ? { tooltip: a.tooltip } : {}),
      content: this.readInlines(elements(e), ctx),
    };
  }

  readRun(r: Element, ctx: StoryContext): Inline[] {
    const props = readRunProps(child(r, "w:rPr"));
    const out: Inline[] = [];
    // Every inline produced from a run's children carries that run's formatting.
    const add = (i: Exclude<Inline, { type: "link" | "field" | "preserved" | "contentControl" | "revision" }>) => out.push(props ? { ...i, props } : i);
    for (const c of elements(r)) {
      const n = qname(c);
      if (n === "w:rPr" || DROPPED.has(n)) continue;
      if (n === "w:t" || n === "w:delText") add({ type: "text", text: textContent(c) });
      else if (n === "w:tab" && !c.attributes.length) add({ type: "tab" });
      else if ((n === "w:br" && Object.keys(attrs(c)).every((k) => k === "type" || k === "clear")) || n === "w:cr") {
        const type = attr(c, "type");
        const kind = type === "page" ? "page" : type === "column" ? "column" : "line";
        const clear = attr(c, "clear");
        add({ type: "break", kind, ...(clear ? { clear } : {}) });
      } else if (n === "w:fldChar" && !elements(c).length && Object.keys(attrs(c)).every((k) => k === "fldCharType")) {
        add({ type: "fieldChar", kind: attr(c, "fldCharType") as "begin" | "separate" | "end" });
      } else if (n === "w:instrText" || n === "w:delInstrText") add({ type: "instr", text: textContent(c) });
      else if ((n === "w:footnoteReference" || n === "w:endnoteReference") && Object.keys(attrs(c)).every((k) => k === "id")) {
        add({ type: "noteRef", kind: n === "w:footnoteReference" ? "footnote" : "endnote", id: attr(c, "id") ?? "" });
      } else if ((n === "w:footnoteRef" || n === "w:endnoteRef") && !c.attributes.length) {
        add({ type: "noteMark", kind: n === "w:footnoteRef" ? "footnote" : "endnote" });
      }
      else if (n === "w:drawing") {
        add(this.readDrawing(c, ctx) ?? this.floating(c, c, ctx) ?? this.runPreserved(c, ctx, child(c, "wp:inline") ? "w:drawing/inline" : n));
      } else if (n === "mc:AlternateContent") {
        // Word 2010+ shapes and text boxes: the modern form is the first choice (VML is the fallback).
        const drawing = elements(c).filter((x) => qname(x) === "mc:Choice").map((x) => child(x, "w:drawing")).find(Boolean);
        add(this.floating(c, drawing, ctx) ?? this.runPreserved(c, ctx));
      } else add(this.runPreserved(c, ctx));
    }
    return out;
  }

  runPreserved(e: Element, ctx: StoryContext, label = qname(e), shown = false): Extract<Inline, { type: "runPreserved" }> {
    const preserved = ctx.keep(e);
    // Word's watermark (a WordArt shape in a header) is shown and edited from the Design tab.
    const watermark = /PowerPlusWaterMarkObject/.test(preserved.xml);
    if (!shown && !watermark && !["w:sym", "w:noBreakHyphen", "w:softHyphen"].includes(label)) this.notePreserved(label);
    return { type: "runPreserved", preserved };
  }

  /**
   * A floating drawing (wp:anchor) Prime Word can show: kept verbatim, with its layout description.
   * Not reported as "kept unchanged" (it is shown, moved and resized); plain shapes still are.
   */
  floating(e: Element, drawing: Element | undefined, ctx: StoryContext): Extract<Inline, { type: "runPreserved" }> | undefined {
    const anchor = drawing && child(drawing, "wp:anchor");
    const float = anchor && this.readFloat(anchor, ctx);
    if (!float) return undefined;
    return { ...this.runPreserved(e, ctx, qname(e), float.kind !== "shape"), float };
  }

  readFloat(anchor: Element, ctx: StoryContext): FloatInfo | undefined {
    const pt = (v: string | null | undefined) => Number(v ?? 0) / EMU_PER_PT;
    const extent = child(anchor, "wp:extent");
    if (!extent) return undefined;
    const position = (name: string): FloatPosition => {
      const e = child(anchor, name);
      const offset = e && child(e, "wp:posOffset");
      const align = e && child(e, "wp:align");
      return {
        from: e?.getAttribute("relativeFrom") ?? "column",
        ...(offset ? { offset: pt(textContent(offset)) } : align ? { align: textContent(align).trim() } : { offset: 0 }),
      };
    };
    const wrapEl = elements(anchor).find((e) => /^wp:wrap/.test(qname(e)));
    const wrapName = wrapEl ? qname(wrapEl).slice("wp:wrap".length) : "None";
    const wrap = ({ Square: "square", Tight: "tight", Through: "through", TopAndBottom: "topAndBottom" } as const)[wrapName as "Square"] ?? "none";
    const side = wrapEl?.getAttribute("wrapText") as FloatInfo["side"] | null;
    const data = child(child(anchor, "a:graphic") ?? anchor, "a:graphicData");
    const uri = data?.getAttribute("uri") ?? "";
    const float: FloatInfo = {
      kind: "shape",
      width: pt(extent.getAttribute("cx")),
      height: pt(extent.getAttribute("cy")),
      x: position("wp:positionH"),
      y: position("wp:positionV"),
      wrap,
      ...(side ? { side } : {}),
      ...(anchor.getAttribute("behindDoc") === "1" ? { behind: true } : {}),
      dist: {
        top: pt(anchor.getAttribute("distT")), bottom: pt(anchor.getAttribute("distB")),
        left: pt(anchor.getAttribute("distL")), right: pt(anchor.getAttribute("distR")),
      },
    };
    if (data && uri.endsWith("/picture")) {
      const blip = data.getElementsByTagNameNS(NS.a, "blip")[0];
      const rel = blip && ctx.rels.get(attr(blip, "embed", NS.r) ?? "");
      if (rel && !rel.external) return { ...float, kind: "picture", imagePart: rel.target };
      return float;
    }
    if (data && uri.endsWith("/wordprocessingInk")) {
      // Ink: its strokes live in a part of their own (word/ink/inkN.xml).
      const part = elements(data).find((e) => e.localName === "contentPart");
      const rel = part && ctx.rels.get(attr(part, "id", NS.r) ?? "");
      const file = rel && !rel.external ? this.files[rel.target] : undefined;
      const ink = file ? readInk(this.decoder.decode(file)) : undefined;
      return ink ? { ...float, kind: "ink", ink } : float;
    }
    const shape = data && child(data, "wps:wsp");
    const content = shape && child(shape, "wps:txbx") && child(child(shape, "wps:txbx")!, "w:txbxContent");
    if (shape && content) {
      const body = child(shape, "wps:bodyPr");
      const inset = (name: string, fallback: number) => (body?.getAttribute(name) != null ? pt(body.getAttribute(name)) : fallback);
      return {
        ...float,
        kind: "textbox",
        blocks: this.readBlocks(elements(content), ctx),
        // Word's default inner margins: 0.1" left and right, 0.05" top and bottom.
        insets: { left: inset("lIns", 7.2), right: inset("rIns", 7.2), top: inset("tIns", 3.6), bottom: inset("bIns", 3.6) },
        ...(body && child(body, "a:spAutoFit") ? { autoHeight: true } : {}),
      };
    }
    return float;
  }

  /** Inline pictures become ImageInline; anything fancier (floating, cropped, shapes) is preserved. */
  readDrawing(drawing: Element, ctx: StoryContext): ImageInline | undefined {
    const inline = child(drawing, "wp:inline");
    if (!inline || elements(drawing).length !== 1) return undefined;
    const extent = child(inline, "wp:extent");
    const docPr = child(inline, "wp:docPr");
    const pic = child(child(child(inline, "a:graphic") ?? inline, "a:graphicData") ?? inline, "pic:pic");
    const blipFill = pic && child(pic, "pic:blipFill");
    const blip = blipFill && child(blipFill, "a:blip");
    if (!extent || !pic || !blipFill || !blip || !isPlainPicture(inline, pic, blipFill, blip)) return undefined;
    const rel = ctx.rels.get(attr(blip, "embed", NS.r) ?? "");
    if (!rel || rel.external || !this.files[rel.target]) return undefined;
    ctx.consumed.add(rel.id);
    const ext = rel.target.split(".").pop()!.toLowerCase();
    this.images[rel.target] ??= {
      contentType: this.contentTypes.overrides[rel.target] ?? this.contentTypes.defaults[ext] ?? IMAGE_TYPES[ext] ?? "application/octet-stream",
      extension: ext,
      data: toBase64(this.files[rel.target]),
    };
    const name = docPr?.getAttribute("name") ?? undefined;
    const description = docPr?.getAttribute("descr") ?? undefined;
    return {
      type: "image",
      imageId: rel.target,
      width: Number(extent.getAttribute("cx")) / EMU_PER_PT,
      height: Number(extent.getAttribute("cy")) / EMU_PER_PT,
      ...(name ? { name } : {}),
      ...(description ? { description } : {}),
    };
  }

  // --- tables ------------------------------------------------------------------------------

  readTable(tbl: Element, ctx: StoryContext): Table {
    const table: Table = { type: "table", props: readTableProps(child(tbl, "w:tblPr")), grid: [], rows: [] };
    for (const c of elements(tbl)) {
      const n = qname(c);
      if (n === "w:tblPr") continue;
      if (n === "w:tblGrid") {
        table.grid = elements(c).filter((g) => qname(g) === "w:gridCol").map((g) => Number(attr(g, "w") ?? 0) / 20);
      } else if (n === "w:tr") table.rows.push(this.readRow(c, ctx));
      else this.warn(`${n} inside a table is not supported and was dropped.`);
    }
    return table;
  }

  readRow(tr: Element, ctx: StoryContext): TableRow {
    const row: TableRow = { props: readRowProps(child(tr, "w:trPr")), cells: [] };
    for (const c of elements(tr)) {
      const n = qname(c);
      if (n === "w:trPr") continue;
      if (n === "w:tc") {
        const blocks = this.readBlocks(elements(c).filter((x) => qname(x) !== "w:tcPr"), ctx);
        row.cells.push({ props: readCellProps(child(c, "w:tcPr")), blocks });
      } else {
        if (n !== "w:tblPrEx") this.notePreserved(n);
        (row.extra ??= []).push(ctx.keep(c));
      }
    }
    return row;
  }
}

/** All attributes with their qualified names, including namespace declarations they rely on. */
function rawAttrs(e: Element): Record<string, string> {
  const out: Record<string, string> = {};
  for (let i = 0; i < e.attributes.length; i++) out[e.attributes[i].name] = e.attributes[i].value;
  return out;
}

const only = (e: Element, names: string[]) => elements(e).every((c) => names.includes(qname(c)));

/**
 * True when an inline picture carries nothing beyond what ImageInline models (image, size, name,
 * alt text). Borders, effects, recolouring, crops and picture hyperlinks are not modeled yet, so
 * such pictures are preserved verbatim instead of being regenerated without them.
 */
function isPlainPicture(inline: Element, pic: Element, blipFill: Element, blip: Element): boolean {
  const docPr = child(inline, "wp:docPr");
  const graphicData = child(child(inline, "a:graphic")!, "a:graphicData")!;
  const spPr = child(pic, "pic:spPr");
  const srcRect = child(blipFill, "a:srcRect");
  const geom = spPr && child(spPr, "a:prstGeom");
  const line = spPr && child(spPr, "a:ln");
  return only(inline, ["wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"])
    && (!docPr || only(docPr, ["a:extLst"]))
    && only(graphicData, ["pic:pic"])
    && only(pic, ["pic:nvPicPr", "pic:blipFill", "pic:spPr"])
    && only(blipFill, ["a:blip", "a:srcRect", "a:stretch"]) && !(srcRect && srcRect.attributes.length)
    && only(blip, ["a:extLst"]) && !attr(blip, "link", NS.r)
    && (!spPr || only(spPr, ["a:xfrm", "a:prstGeom", "a:noFill", "a:ln"]))
    && (!geom || geom.getAttribute("prst") === "rect")
    && (!line || only(line, ["a:noFill"]));
}

/** Merge adjacent text inlines whose formatting is identical (DOCX splits runs arbitrarily). */
function mergeText(items: Inline[]): Inline[] {
  const out: Inline[] = [];
  for (const i of items) {
    const prev = out[out.length - 1];
    if (i.type === "text" && prev?.type === "text" && JSON.stringify(prev.props) === JSON.stringify(i.props)) {
      out[out.length - 1] = { ...prev, text: prev.text + i.text };
    } else if (i.type === "link" || i.type === "field") {
      out.push({ ...i, content: mergeText(i.content) });
    } else out.push(i);
  }
  return out;
}

/** Why a file is not something we can open, in words for the person opening it. */
function notADocx(data: Uint8Array): string {
  const starts = (...bytes: number[]) => bytes.every((b, i) => data[i] === b);
  if (!data.length) return "The file is empty.";
  if (starts(0x25, 0x50, 0x44, 0x46)) return "This is a PDF file, not a Word document.";
  if (starts(0xd0, 0xcf, 0x11, 0xe0)) {
    return "This is an older Word document (.doc). Prime Word opens .docx files: open it in Word and use Save As → Word Document (.docx).";
  }
  if (starts(0x7b, 0x5c, 0x72, 0x74, 0x66)) return "This is a Rich Text (.rtf) file, not a Word document (.docx).";
  return "The file is damaged or is not a Word document (.docx). If it was downloaded or copied, try getting it again.";
}

export function readDocx(data: Uint8Array): ReadResult {
  let reader: Reader;
  try {
    reader = new Reader(data);
  } catch {
    throw new Error(notADocx(data));
  }
  try {
    return reader.read();
  } catch (e) {
    const message = (e as Error).message ?? String(e);
    if (/^XML parse error/.test(message)) throw new Error("Part of this document is damaged and cannot be read.");
    throw e;
  }
}
