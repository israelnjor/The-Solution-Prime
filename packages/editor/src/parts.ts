/**
 * Blocks ("quick parts"): pieces of a document kept for reuse — a signature block, an address, a
 * stamp, a closing, or anything selected and saved. A part is a short story (blocks) with the pictures
 * it shows; inserting one copies it in at the cursor, or into the header or footer, aligned left,
 * centred or right.
 */
import type { Block, ImageResource, Inline, ParagraphProps, PrimeDocument, RunProps } from "@prime/pdm";
import { Slice } from "prosemirror-model";
import type { Command } from "prosemirror-state";
import { storyFromProseMirror, storyToProseMirror } from "./convert.js";
import { setStory, storyAt } from "./docsetup.js";

export interface Part {
  blocks: Block[];
  /** The pictures the blocks show, by image id. */
  images: Record<string, ImageResource>;
}

export type PartPlace = "cursor" | "header" | "footer";
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

/** Inserts a part at the cursor (as whole paragraphs), or at the end of the header or footer. */
export function insertPart(part: Part, doc: () => PrimeDocument, place: PartPlace = "cursor", align?: PartAlign): Command {
  return (state, dispatch, view) => {
    if (!dispatch) return true;
    const blocks = prepared(part, doc(), align);
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
