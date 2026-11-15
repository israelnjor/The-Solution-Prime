/**
 * Blocks ("quick parts"): pieces of a document kept for reuse — a signature block, an address, a
 * stamp, a closing, or anything selected and saved. A part is a short story (blocks) with the pictures
 * it shows; inserting one copies it in at the cursor, or into the header or footer, aligned left,
 * centred or right.
 */
import type { Block, ImageResource, Inline, ParagraphProps, PrimeDocument, RunProps } from "@prime/pdm";
import { type Node, Slice } from "prosemirror-model";
import type { Command } from "prosemirror-state";
import { storyFromProseMirror, storyToProseMirror } from "./convert.js";
import { setStory, storyAt } from "./docsetup.js";
import { pinnedStart } from "./paged.js";

export interface Part {
  blocks: Block[];
  /** The pictures the blocks show, by image id. */
  images: Record<string, ImageResource>;
}

/**
 * Where a part goes: as the last paragraphs of the document, at the bottom of the last page (kept there
 * as the text grows), at the top of the letter (the sender's address), where the recipient's address goes
 * in a letter (above the greeting), at the cursor, or in the header or footer.
 */
export type PartPlace = "end" | "bottom" | "start" | "recipient" | "cursor" | "header" | "footer";
export type PartAlign = "left" | "center" | "right";

/** Every `imageId` in a value (pictures in paragraphs, links, tables, text boxes, floats). */
function eachImageId(value: unknown, f: (holder: { imageId: string }) => void) {
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) return value.forEach((v) => eachImageId(v, f));
  const o = value as Record<string, unknown>;
  if (typeof o.imageId === "string") f(o as { imageId: string });
  for (const k in o) if (k !== "imageId") eachImageId(o[k], f);
}

/** The selection as a part (its pictures included). */
export function partFromSelection(doc: () => PrimeDocument): (state: Parameters<Command>[0]) => Part | undefined {
  return (state) => {
    if (state.selection.empty) return undefined;
    const slice = state.selection.content(); // with the paragraphs (and table cells) around it
    let pm;
    try { pm = state.schema.nodes.doc.create({ section: null }, slice.content); } catch { return undefined; }
    const blocks = structuredClone(storyFromProseMirror(pm));
    const all = doc().images;
    const images: Record<string, ImageResource> = {};
    eachImageId(blocks, (h) => { if (all[h.imageId]) images[h.imageId] = all[h.imageId]; });
    return { blocks, images };
  };
}

/** The part's blocks, aligned, with its pictures added to the document under new ids. */
function prepared(part: Part, doc: PrimeDocument, align?: PartAlign): Block[] {
  const blocks = structuredClone(part.blocks);
  const renamed = new Map<string, string>();
  eachImageId(blocks, (h) => {
    const image = part.images[h.imageId];
    if (!image) return;
    let id = renamed.get(h.imageId);
    if (!id) {
      let n = Object.keys(doc.images).length + 1;
      while (doc.images[`word/media/prime-part${n}.${image.extension}`]) n++;
      id = `word/media/prime-part${n}.${image.extension}`;
      doc.images[id] = image;
      renamed.set(h.imageId, id);
    }
    h.imageId = id;
  });
  if (align) {
    for (const b of blocks) {
      if (b.type !== "paragraph") continue;
      b.props = { ...b.props, align };
      // A floating picture (a stamp) ignores the paragraph's alignment: align its position instead.
      for (const i of b.content) if (i.type === "image" && i.float) i.float = { ...i.float, x: { from: "column", align } };
    }
  }
  return blocks;
}

/** A greeting that opens a letter: "Dear Sir,", "Dear Ms Mensah", "Sir,", "To whom it may concern". */
const GREETING = /^\s*(dear\b|sirs?\b|madam\b|sir\s*\/\s*madam|to whom\b|hello\b|good (morning|afternoon|day)\b)/i;
const MONTH = "jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|june?|july?|aug(ust)?|sep(t(ember)?)?|oct(ober)?|nov(ember)?|dec(ember)?";
/** A date on a line of its own: "15 November 2026", "November 15, 2026", "15th Nov., 2026", "15/11/2026". */
const DATE = new RegExp(String.raw`^\s*(date\s*:?\s*)?((\w+day,?\s+)?\d{1,2}(st|nd|rd|th)?\s+(of\s+)?(${MONTH})\.?,?\s+\d{4}|(${MONTH})\.?\s+\d{1,2}(st|nd|rd|th)?,?\s+\d{4}|\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{4}-\d{2}-\d{2})\.?\s*$`, "i");

/** The top-level blocks of the first section, with their positions (the opening of a letter). */
function opening(doc: Node, max = 40): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  let pos = 0;
  for (let i = 0; i < doc.childCount && out.length < max; i++) {
    const node = doc.child(i);
    if (node.type.name === "section_break") break;
    out.push({ node, pos });
    pos += node.nodeSize;
  }
  return out;
}
const isEmptyPara = (n: Node | undefined) => !!n && n.type.name === "paragraph" && n.content.size === 0;
const alignOf = (n: Node) => (n.attrs.props as ParagraphProps | null)?.align;

/**
 * Where the recipient's address goes in a letter: above the greeting; with no greeting, after a date
 * near the top, or after the sender's address (short lines aligned right at the very top); otherwise at
 * the top. `gapBefore`: a blank line is needed between what is above and the address.
 */
export function recipientPlace(doc: Node): { pos: number; gapBefore: boolean } {
  const top = opening(doc);
  const greeting = top.findIndex(({ node }) => node.type.name === "paragraph" && node.textContent.length < 80 && GREETING.test(node.textContent));
  if (greeting >= 0) {
    // Above the greeting, and above the blank line(s) before it, so the address keeps them below.
    let i = greeting;
    while (i > 0 && isEmptyPara(top[i - 1].node)) i--;
    return { pos: top[i].pos, gapBefore: i > 0 };
  }
  const after = (i: number) => {
    // After block i and a blank line following it.
    const next = top[i + 1];
    return isEmptyPara(next?.node) ? { pos: next.pos + next.node.nodeSize, gapBefore: false } : { pos: top[i].pos + top[i].node.nodeSize, gapBefore: true };
  };
  const date = top.slice(0, 12).findIndex(({ node }) => node.type.name === "paragraph" && DATE.test(node.textContent));
  if (date >= 0) return after(date);
  let sender = 0;
  while (sender < top.length && top[sender].node.type.name === "paragraph" && top[sender].node.content.size > 0
    && top[sender].node.textContent.length < 60 && alignOf(top[sender].node) === "right") sender++;
  if (sender > 0 && sender < 8) return after(sender - 1);
  return { pos: 0, gapBefore: false };
}

/** Whether a part can be kept at the bottom of the last page (paragraphs only: a table cannot be framed). */
export function canPinToBottom(part: Part): boolean {
  return part.blocks.length > 0 && part.blocks.every((b) => b.type === "paragraph");
}

/** The frame that keeps a paragraph at the foot of the margin (Word shows it there too). */
const bottomFrame = (align: PartAlign = "left") =>
  ({ name: "w:framePr", xml: `<w:framePr w:wrap="notBeside" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="${align}" w:yAlign="bottom"/>` });

/**
 * Inserts a part as the last paragraphs of the document, at the bottom of the last page (after
 * everything, kept at the page's foot as the text above grows), at the cursor (as whole paragraphs), or
 * at the end of the header or footer.
 */
export function insertPart(part: Part, doc: () => PrimeDocument, place: PartPlace = "cursor", align?: PartAlign): Command {
  return (state, dispatch, view) => {
    if (!dispatch) return true;
    const blocks = prepared(part, doc(), align);
    if (place === "bottom" && canPinToBottom(part)) {
      // Trailing empty paragraphs would only take room at the foot of the page.
      while (blocks.length > 1) {
        const last = blocks[blocks.length - 1];
        if (last.type === "paragraph" && !last.content.length) blocks.pop();
        else break;
      }
      for (const b of blocks) {
        if (b.type !== "paragraph") continue;
        const extra = (b.props.extra ?? []).filter((x) => x.name !== "w:framePr");
        b.props = { ...b.props, align: align ?? b.props.align, keepLines: true, extra: [...extra, bottomFrame(align)] };
      }
      const pm = storyToProseMirror(blocks, state.schema);
      // After everything (after another block kept there: two signatures stack).
      const tr = state.tr.insert(state.doc.content.size, pm.content);
      dispatch(tr);
      return true;
    }
    if (place === "start" || place === "recipient") {
      const pm = storyToProseMirror(blocks, state.schema);
      const empty = state.doc.childCount === 1 && isEmptyPara(state.doc.firstChild ?? undefined);
      if (empty) {
        // An empty document: the part replaces its one empty paragraph.
        dispatch(state.tr.replaceWith(0, state.doc.content.size, pm.content).scrollIntoView());
        return true;
      }
      const at = place === "start" ? { pos: 0, gapBefore: false } : recipientPlace(state.doc);
      let content = pm.content;
      if (at.gapBefore) content = content.addToStart(state.schema.nodes.paragraph.create());
      // A blank line already following the place: the part's own trailing blank line is not needed too.
      const next = state.doc.resolve(at.pos).nodeAfter ?? undefined;
      if (isEmptyPara(next) && content.lastChild && isEmptyPara(content.lastChild) && content.childCount > 1) {
        content = content.cut(0, content.size - content.lastChild.nodeSize);
      }
      dispatch(state.tr.insert(at.pos, content).scrollIntoView());
      return true;
    }
    if (place === "cursor") {
      const pm = storyToProseMirror(blocks, state.schema);
      const { $from } = state.selection;
      // In an empty paragraph the part replaces it; elsewhere it goes in as paragraphs of its own.
      const empty = $from.parent.type.name === "paragraph" && $from.parent.content.size === 0;
      const tr = empty
        ? state.tr.replaceWith($from.before(), $from.after(), pm.content)
        : state.tr.replaceSelection(new Slice(pm.content, 0, 0));
      dispatch(tr.scrollIntoView());
      return true;
    }
    if (place === "end" || place === "bottom") {
      // The last paragraphs of the document (also a part that cannot be framed, such as a table, put at
      // the bottom). An empty last paragraph is taken over, so no blank line is left above the part.
      const pm = storyToProseMirror(blocks, state.schema);
      // Above a block kept at the bottom of the page, which must stay last.
      const pin = place === "end" ? pinnedStart(state.doc) : undefined;
      if (pin !== undefined) {
        dispatch(state.tr.insert(pin, pm.content).scrollIntoView());
        return true;
      }
      const last = state.doc.lastChild;
      const end = state.doc.content.size;
      const tr = last && last.type.name === "paragraph" && last.content.size === 0 && state.doc.childCount > 1
        ? state.tr.replaceWith(end - last.nodeSize, end, pm.content)
        : state.tr.insert(end, pm.content);
      dispatch(tr.scrollIntoView());
      return true;
    }
    const story = storyAt(state, place);
    return setStory(place, { ...story, blocks: [...(story?.blocks ?? []), ...blocks] })(state, dispatch, view);
  };
}

const text = (t: string, props?: RunProps): Inline => ({ type: "text", text: t, ...(props ? { props } : {}) });
const para = (content: Inline[], props: ParagraphProps = {}): Block => ({ type: "paragraph", props, content });
const tight = { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } };

export interface SignatureBlockSpec {
  /** "Yours faithfully," etc. */
  closing?: string;
  /** A saved signature picture (PNG). */
  signature?: { png: string; width: number; height: number };
  /** A line to sign on (under the signature). */
  line?: boolean;
  name: string;
  title?: string;
  organisation?: string;
  /** "Date: …" under the name. */
  date?: string;
}

/** A signature block: closing, signature picture (or room to sign), a line, name, title, organisation, date. */
export function signatureBlock(spec: SignatureBlockSpec): Part {
  const blocks: Block[] = [];
  const images: Record<string, ImageResource> = {};
  if (spec.closing) blocks.push(para([text(spec.closing)], { spacing: { before: 0, after: 0 } }));
  if (spec.signature) {
    const id = "word/media/signature.png";
    images[id] = { contentType: "image/png", extension: "png", data: spec.signature.png };
    // Shown about 120pt wide, never taller than 54pt.
    let w = 120;
    let h = (w / spec.signature.width) * spec.signature.height;
    if (h > 54) { w *= 54 / h; h = 54; }
    blocks.push(para([{ type: "image", imageId: id, width: w, height: h, name: "Signature", description: `Signature of ${spec.name}` }], tight));
  } else blocks.push(para([], { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } }), para([], tight));
  // The line to sign on: underscores keep it the same under any alignment.
  if (spec.line) blocks.push(para([text("_".repeat(30))], tight));
  blocks.push(para([text(spec.name, { bold: true })], tight));
  for (const line of [spec.title, spec.organisation, spec.date ? `Date: ${spec.date}` : undefined]) {
    if (line) blocks.push(para([text(line)], tight));
  }
  blocks.push(para([], {}));
  return { blocks, images };
}

/** An address: one tight paragraph per line, then a space. */
export function addressBlock(lines: string[], opts: { bold?: boolean } = {}): Part {
  const kept = lines.map((l) => l.trim()).filter(Boolean);
  return {
    blocks: [...kept.map((l, i) => para([text(l, i === 0 && opts.bold ? { bold: true } : undefined)], tight)), para([], {})],
    images: {},
  };
}

/** A picture as a part (a stamp, a logo): in line, or floating in front of the text. */
export function picturePart(png: { data: string; width: number; height: number }, opts: { width: number; name: string; float?: boolean }): Part {
  const id = "word/media/part.png";
  const height = (opts.width / png.width) * png.height;
  const image: Inline = {
    type: "image", imageId: id, width: opts.width, height, name: opts.name, description: opts.name,
    ...(opts.float ? {
      float: {
        kind: "picture", width: opts.width, height, imageId: id, x: { from: "column", offset: 0 }, y: { from: "paragraph", offset: 0 },
        wrap: "none", side: "bothSides", dist: { top: 0, bottom: 0, left: 9, right: 9 },
      },
    } : {}),
  } as Inline;
  return { blocks: [para([image], tight)], images: { [id]: { contentType: "image/png", extension: "png", data: png.data } } };
}
