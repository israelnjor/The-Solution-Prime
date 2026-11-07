/**
 * Review: tracked changes (accept / reject, move between them) and comments (list, add, delete,
 * highlight), on the model Word wrote — revisions are `revision` marks; comments are range markers
 * (w:commentRangeStart / End), a reference in a run, and an entry in word/comments.xml.
 */
import type { PrimeDocument } from "@prime/pdm";
import type { Mark, Node } from "prosemirror-model";
import { type Command, type EditorState, Plugin, PluginKey, TextSelection, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

// --- tracked changes -------------------------------------------------------------------------

export interface Change {
  from: number;
  to: number;
  kind: "ins" | "del";
  author?: string;
  date?: string;
  /** A paragraph mark (an inserted or deleted paragraph break): the position of its paragraph. */
  paragraph?: number;
}

/** A tracked change on a paragraph's mark (w:pPr/w:rPr/w:ins or w:del), if any. */
function markRevision(node: Node): { kind: "ins" | "del"; author?: string; date?: string } | undefined {
  const extra = (node.attrs.props?.markRun?.extra ?? []) as { name: string; xml: string }[];
  const r = extra.find((x) => x.name === "w:ins" || x.name === "w:del");
  if (!r) return undefined;
  const a = (n: string) => new RegExp(`w:${n}="([^"]*)"`).exec(r.xml)?.[1];
  return { kind: r.name === "w:ins" ? "ins" : "del", author: a("author"), date: a("date") };
}

const revisionOf = (node: Node): Mark | undefined => node.marks.find((m) => m.type.name === "revision");

/** All tracked changes, in document order (adjacent text with the same revision joined). */
export function changes(doc: Node): Change[] {
  const out: (Change & { mark: Mark })[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "paragraph") {
      const r = markRevision(node);
      // The mark sits at the paragraph's end.
      if (r) out.push({ from: pos + node.nodeSize - 1, to: pos + node.nodeSize - 1, ...r, paragraph: pos, mark: node.type.schema.marks.revision.create({ kind: r.kind }) });
    }
    if (!node.isInline) return true;
    const mark = revisionOf(node);
    if (!mark) return false;
    const last = out[out.length - 1];
    if (last && last.to === pos && last.mark.eq(mark)) last.to = pos + node.nodeSize;
    else {
      // Revision attributes keep their qualified names (w:author).
      const a = (mark.attrs.attrs ?? {}) as Record<string, string>;
      out.push({ from: pos, to: pos + node.nodeSize, kind: mark.attrs.kind, author: a["w:author"] ?? a.author, date: a["w:date"] ?? a.date, mark });
    }
    return false;
  });
  return out.map(({ mark: _mark, ...c }) => c).sort((a, b) => a.from - b.from);
}

/** The change at (or touching) the selection. */
export function changeAt(state: EditorState): Change | undefined {
  const { from, to } = state.selection;
  return changes(state.doc).find((c) => c.from <= to && c.to >= from);
}

/** Accepting keeps an insertion (drops its mark) and removes a deletion; rejecting does the opposite. */
function resolve(tr: Transaction, c: Change, accept: boolean) {
  const keep = (c.kind === "ins") === accept;
  if (c.paragraph !== undefined) {
    // A paragraph break: drop the revision from its mark; a break that goes joins the next paragraph on.
    const pos = tr.mapping.map(c.paragraph);
    const node = tr.doc.nodeAt(pos);
    if (!node) return;
    const markRun = node.attrs.props.markRun ?? {};
    const extra = (markRun.extra ?? []).filter((x: { name: string }) => x.name !== "w:ins" && x.name !== "w:del");
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: { ...node.attrs.props, markRun: { ...markRun, extra } } });
    const end = pos + node.nodeSize;
    const next = tr.doc.resolve(end).nodeAfter;
    if (!keep && next?.type.name === "paragraph") tr.join(end);
    return;
  }
  const from = tr.mapping.map(c.from);
  const to = tr.mapping.map(c.to);
  if (keep) tr.removeMark(from, to, tr.doc.type.schema.marks.revision);
  else tr.delete(from, to);
}

export function acceptOrReject(accept: boolean, all = false): Command {
  return (state, dispatch) => {
    const list = all ? changes(state.doc) : [changeAt(state)].filter((c): c is Change => !!c);
    if (!list.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const c of list) resolve(tr, c, accept);
      // Never recorded as a change itself (pw-notrack: trackchanges.ts).
      dispatch(tr.setMeta("pw-notrack", true).scrollIntoView());
    }
    return true;
  };
}

/** Selects the next (or previous) change after the selection; wraps around. */
export function gotoChange(direction: 1 | -1): Command {
  return (state, dispatch) => {
    const list = changes(state.doc);
    if (!list.length) return false;
    const { from, to } = state.selection;
    const next = direction > 0 ? list.find((c) => c.from >= to) ?? list[0] : [...list].reverse().find((c) => c.to <= from) ?? list[list.length - 1];
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, next.from, next.to)).scrollIntoView());
    return true;
  };
}

// --- comments ----------------------------------------------------------------------------------

export interface Comment {
  id: string;
  author: string;
  initials?: string;
  date?: string;
  text: string;
}

const COMMENTS = "word/comments.xml";
const COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function decode(b64: string): string {
  return new TextDecoder().decode(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
}
function encode(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
const unescape = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** The comments in word/comments.xml (all of them, including any no longer anchored in the text). */
export function readComments(base: PrimeDocument): Comment[] {
  const part = base.package.parts.find((p) => p.name === COMMENTS);
  if (!part) return [];
  const xml = decode(part.data);
  return [...xml.matchAll(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g)].map(([, attrs, body]) => {
    const a = (n: string) => { const v = new RegExp(`w:${n}="([^"]*)"`).exec(attrs)?.[1]; return v === undefined ? undefined : unescape(v); };
    const text = [...body.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
      .map((p) => [...p[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => unescape(t[1])).join(""))
      .join("\n");
    return { id: a("id") ?? "", author: a("author") ?? "", initials: a("initials"), date: a("date"), text };
  });
}

/** Adds a comment to word/comments.xml, creating the part (and its relationship) if needed. */
export function addCommentEntry(base: PrimeDocument, c: Comment) {
  const pkg = base.package;
  const paragraphs = c.text.split("\n").map((line, i) =>
    `<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${i === 0 ? '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' : ""}` +
    `<w:r><w:t xml:space="preserve">${escape(line)}</w:t></w:r></w:p>`).join("");
  const entry = `<w:comment w:id="${c.id}" w:author="${escape(c.author)}"${c.date ? ` w:date="${c.date}"` : ""}${c.initials ? ` w:initials="${escape(c.initials)}"` : ""}>${paragraphs}</w:comment>`;
  const part = pkg.parts.find((p) => p.name === COMMENTS);
  if (part) {
    part.data = encode(decode(part.data).replace(/<\/w:comments>\s*$/, `${entry}</w:comments>`));
    return;
  }
  pkg.parts.push({ name: COMMENTS, data: encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W}">${entry}</w:comments>`) });
  pkg.contentTypes.overrides[COMMENTS] = COMMENTS_TYPE;
  if (!pkg.preservedRels.some((r) => r.type === COMMENTS_REL)) {
    let n = 1;
    while (pkg.preservedRels.some((r) => r.id === `rIdC${n}`)) n++;
    pkg.preservedRels.push({ id: `rIdC${n}`, type: COMMENTS_REL, target: "comments.xml" });
  }
}

/** Where each comment sits in the text: its range (between the markers) and its reference. */
export function commentRanges(doc: Node): Map<string, { from: number; to: number; ref?: number }> {
  const out = new Map<string, { from: number; to: number; ref?: number }>();
  const idOf = (xml: string) => /w:id="([^"]*)"/.exec(xml)?.[1];
  doc.descendants((node, pos) => {
    const p = (node.type.name === "inline_preserved" || node.type.name === "run_preserved") ? node.attrs.preserved as { name: string; xml: string } : undefined;
    if (!p) return true;
    const id = idOf(p.xml);
    if (id === undefined) return false;
    const r = out.get(id) ?? { from: pos, to: pos };
    if (p.name === "w:commentRangeStart") r.from = pos + node.nodeSize;
    else if (p.name === "w:commentRangeEnd") r.to = pos;
    else if (p.name === "w:commentReference") { r.ref = pos; if (r.to < r.from) r.to = pos; }
    out.set(id, r);
    return false;
  });
  return out;
}

/** Comments that are anchored in the text, in reading order. */
export function anchoredComments(doc: Node, base: PrimeDocument): (Comment & { from: number; to: number })[] {
  const ranges = commentRanges(doc);
  return readComments(base).flatMap((c) => {
    const r = ranges.get(c.id);
    return r ? [{ ...c, from: r.from, to: Math.max(r.from, r.to) }] : [];
  }).sort((a, b) => a.from - b.from);
}

/** Comments on the selection: markers around it, a reference after it, and the entry in comments.xml. */
export function addComment(base: () => PrimeDocument, author: string, text: string): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty || !text.trim()) return false;
    if (dispatch) {
      const doc = base();
      const ids = [...readComments(doc).map((c) => Number(c.id)), ...[...commentRanges(state.doc).keys()].map(Number)].filter((n) => !Number.isNaN(n));
      const id = String(ids.length ? Math.max(...ids) + 1 : 0);
      const initials = author.split(/\s+/).filter(Boolean).map((w) => w[0]).join("").slice(0, 3).toUpperCase();
      addCommentEntry(doc, { id, author, initials, date: new Date().toISOString().replace(/\.\d+Z$/, "Z"), text });
      const { schema } = state;
      const marker = (name: string) => schema.nodes.inline_preserved.create({ preserved: { name, xml: `<${name} w:id="${id}"/>` } });
      const tr = state.tr;
      // End first, so the start position stays valid.
      tr.insert(to, [marker("w:commentRangeEnd"),
        schema.nodes.run_preserved.create({ props: { style: "CommentReference" }, preserved: { name: "w:commentReference", xml: `<w:commentReference w:id="${id}"/>` } })]);
      tr.insert(from, marker("w:commentRangeStart"));
      dispatch(tr);
    }
    return true;
  };
}

/** Removes a comment's markers and reference (its entry is dropped from comments.xml on save). */
export function deleteComment(id: string): Command {
  return (state, dispatch) => {
    const spots: [number, number][] = [];
    state.doc.descendants((node, pos) => {
      const p = (node.type.name === "inline_preserved" || node.type.name === "run_preserved") ? node.attrs.preserved as { name: string; xml: string } : undefined;
      if (p && /^w:comment(RangeStart|RangeEnd|Reference)$/.test(p.name) && new RegExp(`w:id="${id}"`).test(p.xml)) spots.push([pos, pos + node.nodeSize]);
      return !p;
    });
    if (!spots.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const [a, b] of spots.reverse()) tr.delete(a, b);
      dispatch(tr);
    }
    return true;
  };
}

// --- highlighting ------------------------------------------------------------------------------

interface ReviewState {
  active?: string;
  decorations: DecorationSet;
}

export const reviewKey = new PluginKey<ReviewState>("review");

function highlights(doc: Node, active?: string): DecorationSet {
  const out: Decoration[] = [];
  for (const [id, r] of commentRanges(doc)) {
    if (r.to > r.from) out.push(Decoration.inline(r.from, r.to, { class: id === active ? "pw-commented pw-comment-active" : "pw-commented" }));
  }
  return DecorationSet.create(doc, out);
}

/** Whether a change adds or removes kept items (comment markers among them). */
function touchesMarkers(tr: Transaction): boolean {
  const kept = (n: Node) => n.type.name === "inline_preserved" || n.type.name === "run_preserved";
  return tr.steps.some((step, i) => {
    const json = step.toJSON() as { from?: number; to?: number };
    let hit = false;
    const slice = (step as unknown as { slice?: { content: { descendants(f: (n: Node) => boolean | void): void } } }).slice;
    slice?.content.descendants((n) => { if (kept(n)) hit = true; return !hit; });
    if (!hit && json.from !== undefined && json.to !== undefined && json.to > json.from) {
      tr.docs[i].nodesBetween(json.from, json.to, (n) => { if (kept(n)) hit = true; return !hit; });
    }
    return hit;
  });
}

/** Highlights commented text (the active comment more strongly); set it with setMeta(reviewKey, { active }). */
export function reviewPlugin(): Plugin<ReviewState> {
  return new Plugin<ReviewState>({
    key: reviewKey,
    state: {
      init: (_, state) => ({ decorations: highlights(state.doc) }),
      apply(tr, old) {
        const meta = tr.getMeta(reviewKey) as { active?: string } | undefined;
        const active = meta ? meta.active : old.active;
        if (meta || (tr.docChanged && touchesMarkers(tr))) return { active, decorations: highlights(tr.doc, active) };
        return tr.docChanged ? { active, decorations: old.decorations.map(tr.mapping, tr.doc) } : old;
      },
    },
    props: { decorations: (state) => reviewKey.getState(state)?.decorations },
  });
}
