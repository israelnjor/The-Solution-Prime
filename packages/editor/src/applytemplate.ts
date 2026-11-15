/**
 * Applying a template or letterhead to the open document (Word's "attach template", made simple):
 * its header and footer, artwork pinned to its first page, and, if wanted, its page size, margins and
 * styles. The document's own text is never touched.
 *
 * Letterheads come in a few shapes, and each is applied the way it was designed:
 *  - on every page (an ordinary header and footer);
 *  - on the first page only ("Different first page", the following pages blank or with their own,
 *    smaller follow-on header);
 *  - as artwork in the body of the first page (a logo or a full-page background pinned to the page, as
 *    many Word letterheads and converted PDFs have it): it becomes the first page's header, so it stays
 *    on page 1 whatever is typed.
 */
import type { Block, HeaderFooterKind, Inline, PassthroughPart, PrimeDocument, Relationship, SectionProps, Story, Style, StyleSheet } from "@prime/pdm";
import type { EditorState, Transaction } from "prosemirror-state";
import { type SectionAttrs, sections, setSection } from "./docsetup.js";
import { stylesOf } from "./styleedit.js";

/** How the template's letterhead is laid out. */
export type LetterheadPattern = "every" | "first" | "firstAndRest" | "none";

export interface TemplateLook {
  pattern: LetterheadPattern;
  /** The template's own setting: on page 1 (or every page), and on the following pages. */
  first: { header?: Story; footer?: Story };
  rest: { header?: Story; footer?: Story };
  /** Artwork from the template's first page (pinned to the page) moved into its first-page header. */
  artwork: number;
  page: SectionProps;
}

/** Where the letterhead goes. */
export type ApplyPages = "first" | "every";
/** The following pages, when the letterhead is on the first page only. */
export type FollowingPages = "template" | "keep" | "none";

export interface ApplyOptions {
  pages: ApplyPages;
  following: FollowingPages;
  /** The template's page size and margins (all sections). */
  pageSetup: boolean;
  /** The template's styles replace the document's (fonts, headings, colours). */
  styles: boolean;
}

/** Whether a story shows anything: text, a picture, a shape, a field (page number), a table with text. */
export function storyShows(story: Story | undefined): boolean {
  if (!story) return false;
  let found = false;
  const walk = (v: unknown) => {
    if (found || !v || typeof v !== "object") return;
    if (Array.isArray(v)) { for (const x of v) walk(x); return; }
    const o = v as Record<string, unknown>;
    if (o.type === "text" && typeof o.text === "string" && /\S/.test(o.text)) { found = true; return; }
    if (o.type === "image" || o.type === "runPreserved" || o.type === "preserved" || o.type === "field" || o.type === "fieldChar") { found = true; return; }
    for (const k in o) walk(o[k]);
  };
  walk(story.blocks);
  return found;
}

/** A floating object placed by the page (or its margins), not by the text around it. */
const pinned = (i: Inline): boolean => {
  const f = (i.type === "image" || i.type === "runPreserved") ? i.float : undefined;
  if (!f) return false;
  const byText = (from: string) => /^(paragraph|line|character)$/i.test(from);
  return !byText(f.x.from) && !byText(f.y.from);
};

/** The rIds a piece of preserved XML points at (pictures, linked parts). */
const relIds = (xml: string) => [...xml.matchAll(/\br:(?:embed|id|link|pict|dm|lo|qs|cs)="([^"]+)"/g)].map((m) => m[1]);

/**
 * Artwork pinned to the template's first page, in its body (before the first page break): taken out of
 * the body, as a story of its own (one paragraph), with the relationships its preserved XML needs.
 */
function bodyArtwork(t: PrimeDocument): Story | undefined {
  const art: Inline[] = [];
  const rels: Relationship[] = [];
  for (const b of t.sections[0]?.blocks ?? []) {
    if (b.type !== "paragraph") continue;
    for (const i of b.content) {
      if (i.type === "break" && i.kind === "page") return finish();
      if (!pinned(i)) continue;
      art.push(structuredClone(i));
      if (i.type === "runPreserved") {
        for (const id of relIds(i.preserved.xml)) {
          const r = t.package.preservedRels.find((x) => x.id === id);
          if (r && !rels.some((x) => x.id === id)) rels.push(r);
        }
      }
    }
    if (b.props.pageBreakBefore && art.length) break;
  }
  return finish();
  function finish(): Story | undefined {
    if (!art.length) return undefined;
    const para: Block = { type: "paragraph", props: { spacing: { before: 0, after: 0, line: 12, lineRule: "exact" } }, content: art };
    return rels.length ? { blocks: [para], preservedRels: rels } : { blocks: [para] };
  }
}

/** Two stories as one (the second's blocks after the first's). */
function joined(a: Story | undefined, b: Story | undefined): Story | undefined {
  if (!a) return b;
  if (!b) return a;
  const rels = [...(a.preservedRels ?? [])];
  for (const r of b.preservedRels ?? []) if (!rels.some((x) => x.id === r.id)) rels.push(r);
  return { blocks: [...b.blocks, ...a.blocks], ...(rels.length ? { preservedRels: rels } : {}) };
}

/** What the template's letterhead is, read from its first section. */
export function templateLook(t: PrimeDocument): TemplateLook {
  const s = t.sections[0];
  const titled = !!s?.props.titlePage;
  const art = bodyArtwork(t);
  // Page 1 shows the first-page header (with "Different first page") or the ordinary one, plus any artwork.
  const first = {
    header: joined(titled ? s?.headers.first : s?.headers.default, art),
    footer: titled ? s?.footers.first : s?.footers.default,
  };
  const rest = { header: s?.headers.default, footer: s?.footers.default };
  const firstShows = storyShows(first.header) || storyShows(first.footer);
  const restShows = storyShows(rest.header) || storyShows(rest.footer);
  const same = JSON.stringify(first) === JSON.stringify(rest);
  const pattern: LetterheadPattern = !firstShows && !restShows ? "none"
    : !restShows ? "first"
    : !firstShows || same ? "every"
    : "firstAndRest";
  return { pattern, first, rest, artwork: art ? art.blocks[0].type === "paragraph" ? art.blocks[0].content.length : 0 : 0, page: s?.props ?? {} };
}

/** The usual choice for a template: as it was designed. */
export function defaultOptions(look: TemplateLook): ApplyOptions {
  return {
    pages: look.pattern === "every" ? "every" : "first",
    following: look.pattern === "firstAndRest" ? "template" : "keep",
    pageSetup: false,
    styles: false,
  };
}

/** What applying adds to the document's own resources (pictures and passthrough parts). */
interface Additions {
  images: PrimeDocument["images"];
  parts: PassthroughPart[];
  contentTypes: { defaults: Record<string, string>; overrides: Record<string, string> };
}

/** Copies a story from the template, its pictures and parts renamed where the document already has the name. */
function carry(story: Story | undefined, t: PrimeDocument, doc: PrimeDocument, add: Additions, renamed: Map<string, string>): Story | undefined {
  if (!story) return undefined;
  const out: Story = structuredClone(story);
  const taken = (name: string) => !!doc.images[name] || !!add.images[name] || doc.package.parts.some((p) => p.name === name) || add.parts.some((p) => p.name === name);
  const fresh = (name: string) => {
    const ext = /\.([A-Za-z0-9]+)$/.exec(name)?.[1] ?? "bin";
    let n = 1;
    while (taken(`word/media/prime-template${n}.${ext}`)) n++;
    return `word/media/prime-template${n}.${ext}`;
  };
  /** The name a template picture or part has in the document (copied over once). */
  const nameFor = (name: string): string => {
    let to = renamed.get(name);
    if (to) return to;
    const image = t.images[name];
    const part = t.package.parts.find((p) => p.name === name);
    if (!image && !part) return name;
    // The same picture already in the document (applied before): used as it is.
    const same = image ? doc.images[name]?.data === image.data : doc.package.parts.find((p) => p.name === name)?.data === part!.data;
    to = same ? name : taken(name) ? fresh(name) : name;
    if (!same) {
      if (image) add.images[to] = image;
      if (part) {
        add.parts.push({ name: to, data: part.data });
        const ext = /\.([A-Za-z0-9]+)$/.exec(to)?.[1]?.toLowerCase();
        const ct = t.package.contentTypes.overrides[name] ?? (ext ? t.package.contentTypes.defaults[ext] : undefined);
        if (t.package.contentTypes.overrides[name]) add.contentTypes.overrides[to] = t.package.contentTypes.overrides[name];
        else if (ext && ct && !doc.package.contentTypes.defaults[ext]) add.contentTypes.defaults[ext] = ct;
      }
    }
    renamed.set(name, to);
    return to;
  };
  const walk = (v: unknown) => {
    if (!v || typeof v !== "object") return;
    if (Array.isArray(v)) { for (const x of v) walk(x); return; }
    const o = v as Record<string, unknown>;
    if (typeof o.imageId === "string") o.imageId = nameFor(o.imageId);
    if (typeof o.imagePart === "string") o.imagePart = nameFor(o.imagePart);
    for (const k in o) if (k !== "imageId" && k !== "imagePart") walk(o[k]);
  };
  walk(out.blocks);
  if (out.preservedRels) out.preservedRels = out.preservedRels.map((r) => (r.external ? r : { ...r, target: nameFor(r.target) }));
  return out;
}

/** Every style id a value uses (paragraph, character and table styles). */
function usedStyles(v: unknown, into: Set<string>) {
  if (!v || typeof v !== "object") return;
  if (Array.isArray(v)) { for (const x of v) usedStyles(x, into); return; }
  const o = v as Record<string, unknown>;
  if (typeof o.style === "string") into.add(o.style);
  for (const k in o) usedStyles(o[k], into);
}

/** The document's styles with the template's: all of them (`all`), or only those the letterhead needs that the document lacks. */
function mergedStyles(sheet: StyleSheet, t: PrimeDocument, stories: (Story | undefined)[], all: boolean): StyleSheet | undefined {
  const theirs = new Map(t.styles.styles.map((s) => [s.id, s]));
  if (all) {
    const ids = new Set(t.styles.styles.map((s) => s.id));
    return { ...sheet, defaults: t.styles.defaults, styles: [...sheet.styles.filter((s) => !ids.has(s.id)), ...t.styles.styles] };
  }
  const need = new Set<string>();
  usedStyles(stories, need);
  const add: Style[] = [];
  const have = new Set(sheet.styles.map((s) => s.id));
  const take = (id: string | undefined) => {
    if (!id || have.has(id)) return;
    const s = theirs.get(id);
    if (!s) return;
    have.add(id);
    add.push(s);
    take(s.basedOn);
    take(s.link);
  };
  for (const id of need) take(id);
  return add.length ? { ...sheet, styles: [...sheet.styles, ...add] } : undefined;
}

export interface Application {
  tr: Transaction;
  additions: Additions;
}

/**
 * The transaction that applies template `t` to the document (headers, footers, page setup, styles),
 * and the pictures and parts it brings, which `commitApplication` adds to the document.
 */
export function applicationOf(t: PrimeDocument, doc: PrimeDocument, state: EditorState, opts: ApplyOptions): Application {
  const look = templateLook(t);
  const add: Additions = { images: {}, parts: [], contentTypes: { defaults: {}, overrides: {} } };
  const renamed = new Map<string, string>();
  const take = (s: Story | undefined) => carry(s, t, doc, add, renamed);
  const blank = (): Story => ({ blocks: [{ type: "paragraph", props: {}, content: [] }] });
  const parts = (["header", "footer"] as const).map((which) => {
    const firstShows = storyShows(look.first[which]);
    const restShows = storyShows(look.rest[which]);
    const designedFirst = take(look.first[which]);
    const designedRest = take(look.rest[which]);
    // The letterhead itself: page 1's design, or the following pages' when page 1 has none of this part.
    return { which, any: firstShows || restShows, firstShows, restShows, designedFirst, designedRest, main: firstShows ? designedFirst : designedRest };
  });

  let tr = state.tr;
  const all = sections(state);
  const s0 = all[0].attrs;
  all.forEach((s, i) => {
    const attrs: SectionAttrs = { props: { ...s.attrs.props }, headers: { ...s.attrs.headers }, footers: { ...s.attrs.footers } };
    for (const p of parts) {
      const stories = p.which === "header" ? attrs.headers : attrs.footers;
      const before = p.which === "header" ? s.attrs.headers : s.attrs.footers;
      const set = (kind: HeaderFooterKind, v: Story | undefined) => { if (v) stories[kind] = v; else delete stories[kind]; };
      if (opts.pages === "every") {
        if (!p.any) continue;
        // Every page: the first section, and every later one with its own, show the letterhead.
        if (i === 0 || before.default) set("default", p.main);
        if (before.first || (i === 0 && s.attrs.props.titlePage)) set("first", p.main);
        if (before.even) set("even", p.main);
        continue;
      }
      if (i === 0) {
        // Page 1 as it shows now (its own first-page part, or the ordinary one).
        const page1Now = s.attrs.props.titlePage ? before.first : before.default;
        if (!p.any) {
          // The template has none of this part: the document's own stays on every page.
          if (!s.attrs.props.titlePage && before.default) set("first", before.default);
          continue;
        }
        // Page 1: the letterhead. A template that leaves this part blank on page 1 by design keeps it
        // blank when its follow-on pages are used too; otherwise the document's own stays.
        set("first", p.firstShows ? p.designedFirst : opts.following === "template" ? blank() : page1Now ?? blank());
        if (opts.following === "template") { set("default", p.restShows ? p.designedRest : undefined); if (before.even) set("even", p.restShows ? p.designedRest : undefined); }
        else if (opts.following === "none") { set("default", undefined); set("even", undefined); }
        continue;
      }
      if (i === 1 && p.any && opts.following !== "keep") {
        // A later section that inherited section 1's header keeps what it showed before.
        const was = p.which === "header" ? s0.headers : s0.footers;
        if (!before.default && was.default) set("default", was.default);
        if (s.attrs.props.titlePage && !before.first && (was.first ?? was.default)) set("first", was.first ?? was.default);
      }
    }
    if (i === 0 && opts.pages === "first" && parts.some((p) => p.any)) attrs.props.titlePage = true;
    if (opts.pageSetup) {
      const pg = look.page;
      for (const k of ["pageWidth", "pageHeight", "orientation", "paperCode"] as const) {
        if (pg[k] !== undefined) (attrs.props as Record<string, unknown>)[k] = pg[k];
        else delete attrs.props[k];
      }
      if (pg.margins) attrs.props.margins = { ...pg.margins };
    }
    tr = setSection(tr, s.pos, attrs);
  });

  const styles = mergedStyles(stylesOf(state, doc), t, parts.flatMap((p) => [p.designedFirst, p.designedRest]), opts.styles);
  if (styles) tr.setDocAttribute("styles", styles);
  return { tr, additions: add };
}

/** Adds the pictures and parts an application brings to the document (before its transaction is dispatched). */
export function commitApplication(doc: PrimeDocument, a: Application) {
  Object.assign(doc.images, a.additions.images);
  doc.package.parts.push(...a.additions.parts);
  Object.assign(doc.package.contentTypes.defaults, a.additions.contentTypes.defaults);
  Object.assign(doc.package.contentTypes.overrides, a.additions.contentTypes.overrides);
}

/** The document with an application's resources, without changing it (for a preview). */
export function withAdditions(doc: PrimeDocument, a: Application): PrimeDocument {
  return {
    ...doc,
    images: { ...doc.images, ...a.additions.images },
    package: { ...doc.package, parts: [...doc.package.parts, ...a.additions.parts] },
  };
}
