/**
 * Letterheads: an organisation's logo, name, tagline, address and contacts, laid out in the header
 * (and optionally the footer) of a document — centred, or the logo left or right of the details. The
 * result is an ordinary document; saved as a template, every new letter starts with it.
 */
import type { Block, Inline, ParagraphProps, PrimeDocument, RunProps, Story, TableCell } from "@prime/pdm";
import { blankDocument } from "./blank.js";

export interface LetterheadSpec {
  name: string;
  tagline?: string;
  /** Address lines. */
  address: string[];
  /** Phone, e-mail, website… */
  contacts: string[];
  layout: "centre" | "logoLeft" | "logoRight";
  /** Accent colour (name, rule), RRGGBB. */
  color: string;
  font?: string;
  /** A line under the header. */
  rule: boolean;
  /** Repeat the address and contacts in a footer (and keep only the name in the header). */
  footer: boolean;
  logo?: { data: string; contentType: string; width: number; height: number };
}

const text = (t: string, props?: RunProps): Inline => ({ type: "text", text: t, ...(props ? { props } : {}) });
const para = (content: Inline[], props: ParagraphProps = {}): Block => ({ type: "paragraph", props, content });
const NO_BORDERS = '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
const tight = { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } };
const rule = (color: string, side: "bottom" | "top") =>
  ({ name: "w:pBdr", xml: `<w:pBdr><w:${side} w:val="single" w:sz="12" w:space="4" w:color="${color}"/></w:pBdr>` });

/** The letterhead's header and footer stories (pictures are added to `doc.images`). */
export function letterheadStories(doc: PrimeDocument, spec: LetterheadSpec, width: number): { header: Story; footer?: Story } {
  const font = spec.font ? { font: { ascii: spec.font, hAnsi: spec.font, cs: spec.font } } : {};
  const nameRun: RunProps = { ...font, bold: true, size: 18, color: spec.color };
  const small: RunProps = { ...font, size: 9, color: "404040" };
  const details = [spec.address.filter(Boolean).join(", "), spec.contacts.filter(Boolean).join("  |  ")].filter(Boolean);

  let logo: Inline | undefined;
  if (spec.logo) {
    const ext = spec.logo.contentType === "image/jpeg" ? "jpeg" : "png";
    let n = 1;
    while (doc.images[`word/media/letterhead${n}.${ext}`]) n++;
    const imageId = `word/media/letterhead${n}.${ext}`;
    doc.images[imageId] = { contentType: spec.logo.contentType, extension: ext, data: spec.logo.data };
    // Logos are shown about 54pt tall, keeping their proportions.
    const h = 54;
    const w = Math.min((spec.logo.width / spec.logo.height) * h, width * 0.4);
    logo = { type: "image", imageId, width: w, height: (w / spec.logo.width) * spec.logo.height, name: "Logo" };
  }

  const align = spec.layout === "centre" ? "center" : spec.layout === "logoLeft" ? "right" : "left";
  const textBlocks = (a: string): Block[] => [
    para([text(spec.name, nameRun)], { ...tight, align: a }),
    ...(spec.tagline ? [para([text(spec.tagline, { ...font, italic: true, size: 10, color: "595959" })], { ...tight, align: a })] : []),
    ...(spec.footer ? [] : details.map((d) => para([text(d, small)], { ...tight, align: a }))),
  ];

  const header: Block[] = [];
  if (spec.layout === "centre") {
    if (logo) header.push(para([logo], { ...tight, align: "center" }));
    header.push(...textBlocks("center"));
  } else {
    // The logo beside the details: a borderless two-column table.
    const logoWidth = logo && logo.type === "image" ? logo.width + 12 : 0;
    const cell = (blocks: Block[], w: number, vAlign = "center"): TableCell => ({ props: { width: { w: String(Math.round(w * 20)), type: "dxa" }, vAlign }, blocks });
    const logoCell = cell([para(logo ? [logo] : [], { ...tight, align: spec.layout === "logoLeft" ? "left" : "right" })], logoWidth);
    const textCell = cell(textBlocks(align), width - logoWidth);
    header.push({
      type: "table",
      props: { width: { w: "0", type: "auto" }, extra: [{ name: "w:tblBorders", xml: NO_BORDERS }] },
      grid: spec.layout === "logoLeft" ? [logoWidth, width - logoWidth] : [width - logoWidth, logoWidth],
      rows: [{ props: {}, cells: spec.layout === "logoLeft" ? [logoCell, textCell] : [textCell, logoCell] }],
    });
  }
  // The line under the letterhead (an empty paragraph with a bottom border), or just some space.
  header.push(para([], { spacing: { before: 0, after: 6 }, ...(spec.rule ? { extra: [rule(spec.color, "bottom")] } : {}) }));

  const footer = spec.footer && details.length
    ? { blocks: details.map((d, i) => para([text(d, small)], { ...tight, align: "center", ...(i === 0 ? { extra: [rule(spec.color, "top")] } : {}) })) }
    : undefined;
  return { header: { blocks: header }, ...(footer ? { footer } : {}) };
}

/** A new document with the letterhead (a blank letter body). */
export function letterheadDocument(spec: LetterheadSpec): PrimeDocument {
  const doc = blankDocument();
  const section = doc.sections[0];
  const mg = section.props.margins ?? {};
  const width = (section.props.pageWidth ?? 612) - (mg.left ?? 72) - (mg.right ?? 72);
  const { header, footer } = letterheadStories(doc, spec, width);
  section.headers.default = header;
  if (footer) section.footers.default = footer;
  section.blocks = [para([], {})];
  return doc;
}
