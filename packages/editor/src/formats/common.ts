/**
 * Shared by the other file formats (plain text, RTF, OpenDocument, Word 97–2003): building a document
 * from what they hold, and reading back what a document shows (each paragraph's text with its
 * effective formatting, styles resolved), to write them.
 */
import { Resolver, type ResolvedParagraph, type ResolvedRun } from "@prime/layout";
import type { Block, ImageResource, Inline, Paragraph, ParagraphProps, PrimeDocument, RunProps, Style, Table } from "@prime/pdm";
import { blankDocument } from "../blank.js";
import { ensureList, type ListKind } from "../lists.js";

/** A new document (Prime Word's blank one: its styles, page and fonts) to fill. */
export function newDocument(): PrimeDocument {
  const doc = blankDocument();
  doc.sections[0].blocks = [];
  // Tables come in as Word's Table Grid: a single line around every cell.
  if (!doc.styles.styles.some((s) => s.id === "TableGrid")) {
    const line = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    doc.styles.styles.push({
      id: "TableGrid", type: "table", name: "Table Grid", basedOn: "TableNormal",
      paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } },
      table: { extra: [{ name: "w:tblBorders", xml: `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(line).join("")}</w:tblBorders>` }] },
    });
  }
  return doc;
}

/** The style for Heading `level` (1–9), added when the blank document lacks it. */
export function headingStyle(doc: PrimeDocument, level: number): string {
  const id = `Heading${level}`;
  if (!doc.styles.styles.some((s) => s.id === id)) {
    const st: Style = {
      id, type: "paragraph", name: `heading ${level}`, basedOn: "Normal", next: "Normal",
      paragraph: { keepNext: true, keepLines: true, spacing: { before: 2, after: 0 }, outlineLevel: level - 1 },
      run: { size: Math.max(11, 16 - level), color: "1F3763" },
    };
    doc.styles.styles.push(st);
  }
  return id;
}

/** A list (bullets or numbers) to put paragraphs in; one per kind and document. */
export function listId(doc: PrimeDocument, kind: ListKind): string {
  return ensureList(doc, kind);
}

export const paragraph = (content: Inline[], props: ParagraphProps = {}): Paragraph => ({ type: "paragraph", props, content });

/** Text as inlines: tabs and line breaks as their own. */
export function textInlines(text: string, props?: RunProps): Inline[] {
  const out: Inline[] = [];
  const parts = text.split(/(\t|\n)/);
  for (const p of parts) {
    if (!p) continue;
    if (p === "\t") out.push({ type: "tab", ...(props ? { props } : {}) });
    else if (p === "\n") out.push({ type: "break", kind: "line", ...(props ? { props } : {}) });
    else out.push({ type: "text", text: p, ...(props && Object.keys(props).length ? { props } : {}) });
  }
  return out;
}

/** Adds a picture to the document; returns its id. */
export function addImage(doc: PrimeDocument, data: Uint8Array, contentType: string): string {
  const extension = contentType.split("/")[1]?.replace("jpeg", "jpg") ?? "png";
  let n = Object.keys(doc.images).length + 1;
  while (doc.images[`image${n}.${extension}`]) n++;
  const id = `image${n}.${extension}`;
  let bin = "";
  for (let i = 0; i < data.length; i += 0x8000) bin += String.fromCharCode(...data.subarray(i, i + 0x8000));
  doc.images[id] = { contentType, extension, data: btoa(bin) } as ImageResource;
  return id;
}

export function imageBytes(doc: PrimeDocument, id: string): { data: Uint8Array; contentType: string } | undefined {
  const img = doc.images[id];
  if (!img) return undefined;
  const bin = atob(img.data);
  const data = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
  return { data, contentType: img.contentType };
}

// --- reading a document for writing it out ---------------------------------------------------------

/** One piece of a paragraph as shown: text, a tab, a line or page break, or a picture; with its formatting. */
export type Piece =
  | { kind: "text"; text: string; run: ResolvedRun; link?: string }
  | { kind: "tab" | "line" | "page" | "column"; run: ResolvedRun }
  | { kind: "image"; imageId: string; width: number; height: number; run: ResolvedRun };

export interface ShownParagraph { kind: "paragraph"; p: ResolvedParagraph; pieces: Piece[]; heading?: number; source: Paragraph }
export interface ShownTable { kind: "table"; grid: number[]; rows: { cells: { span: number; blocks: Shown[]; shading?: string }[] }[] }
export type Shown = ShownParagraph | ShownTable;

/**
 * The document's body as shown: paragraphs (styles resolved, list labels worked out, tracked
 * deletions left out, field results kept) and tables. For formats without Word's styles.
 */
export function shownBody(doc: PrimeDocument): Shown[] {
  const resolver = new Resolver(doc);
  const headingLevel = (style: string | undefined): number | undefined => {
    const seen = new Set<string>();
    for (let id = style; id && !seen.has(id);) {
      seen.add(id);
      const st = doc.styles.styles.find((s) => s.id === id);
      if (!st) break;
      if (st.paragraph?.outlineLevel !== undefined && st.paragraph.outlineLevel < 9) return st.paragraph.outlineLevel + 1;
      const m = /^heading (\d)$/i.exec(st.name ?? "");
      if (m) return Number(m[1]);
      id = st.basedOn;
    }
    return undefined;
  };
  const para = (p: Paragraph, tableStyle?: string): ShownParagraph => {
    const rp = resolver.paragraph(p, tableStyle);
    const pieces: Piece[] = [];
    const walk = (inlines: Inline[], link?: string, deleted = false) => {
      for (const i of inlines) {
        const run = () => resolver.run(rp.styleId, (i as { props?: RunProps }).props, tableStyle, rp.cond);
        switch (i.type) {
          case "text": if (!deleted) pieces.push({ kind: "text", text: i.text, run: run(), ...(link ? { link } : {}) }); break;
          case "tab": if (!deleted) pieces.push({ kind: "tab", run: run() }); break;
          case "break": if (!deleted) pieces.push({ kind: i.kind, run: run() }); break;
          case "image": if (!deleted && !i.float) pieces.push({ kind: "image", imageId: i.imageId, width: i.width, height: i.height, run: run() }); break;
          case "link": walk(i.content, i.href ?? (i.anchor ? `#${i.anchor}` : link), deleted); break;
          case "field": walk(i.content, link, deleted); break;
          case "contentControl": walk(i.content, link, deleted); break;
          case "revision": walk(i.content, link, deleted || i.kind === "del"); break;
          default: break;
        }
      }
    };
    walk(p.content);
    // Complex fields (begin / instructions / separate / result / end): the instructions are not shown.
    const heading = headingLevel(rp.styleId);
    return { kind: "paragraph", p: rp, pieces, ...(heading ? { heading } : {}), source: p };
  };
  const blocks = (list: Block[], tableStyle?: string): Shown[] => {
    const out: Shown[] = [];
    for (const b of list) {
      if (b.type === "paragraph") out.push(para(b, tableStyle));
      else if (b.type === "table") out.push(table(b));
      else if (b.type === "contentControl") out.push(...blocks(b.blocks, tableStyle));
    }
    return out;
  };
  const table = (t: Table): ShownTable => ({
    kind: "table",
    grid: t.grid,
    rows: t.rows.map((r) => ({
      cells: r.cells.filter((c) => c.props.vMerge !== "continue").map((c) => ({
        span: c.props.span ?? 1, blocks: blocks(c.blocks, t.props.style), ...(c.props.shading?.fill && c.props.shading.fill !== "auto" ? { shading: c.props.shading.fill } : {}),
      })),
    })),
  });
  return doc.sections.flatMap((s) => blocks(s.blocks));
}

/** A paragraph's text as shown (list label first), tabs as tabs, breaks as new lines. */
export function plainText(s: ShownParagraph): string {
  let t = s.p.listLabel ? `${s.p.listLabel}\t` : "";
  for (const piece of s.pieces) {
    if (piece.kind === "text") t += piece.text;
    else if (piece.kind === "tab") t += "\t";
    else if (piece.kind === "line" || piece.kind === "column") t += "\n";
  }
  return t;
}
