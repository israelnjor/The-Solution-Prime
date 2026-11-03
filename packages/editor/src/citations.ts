/**
 * Citations and a bibliography, in Word's own format: the sources live in the document's
 * bibliography part (customXml, <b:Sources>), citations are CITATION fields and the reference list is
 * a BIBLIOGRAPHY field — so Word's Manage Sources sees the same sources and can update the fields.
 * Formatting styles: APA (7th), Harvard and IEEE.
 */
import type { Block, Inline, PrimeDocument, RunProps } from "@prime/pdm";
import type { Command, EditorState } from "prosemirror-state";
import { storyToProseMirror } from "./convert.js";

export type CitationStyle = "APA" | "Harvard" | "IEEE";
export type SourceType = "Book" | "JournalArticle" | "InternetSite" | "Report";

export interface Person { last: string; first?: string }

export interface Source {
  tag: string;
  type: SourceType;
  authors: Person[];
  title: string;
  year?: string;
  /** Book, report: publisher (report: institution); website: site name. */
  publisher?: string;
  city?: string;
  /** Journal article: the journal's name, volume, issue and pages ("12–19"). */
  journal?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  url?: string;
  /** Website: when it was read, as YYYY-MM-DD. */
  accessed?: string;
}

const B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
const STYLE_FILES: Record<CitationStyle, { file: string; name: string; version: string }> = {
  APA: { file: "\\APASixthEditionOfficeOnline.xsl", name: "APA", version: "6" },
  Harvard: { file: "\\HarvardAnglia2008OfficeOnline.xsl", name: "Harvard - Anglia", version: "2008" },
  IEEE: { file: "\\IEEE2006OfficeOnline.xsl", name: "IEEE", version: "2006" },
};

// --- the sources part -------------------------------------------------------------------------

const decode = (b64: string) => new TextDecoder().decode(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
const encode = (text: string) => {
  const bytes = new TextEncoder().encode(text);
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
};
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** The document's bibliography part, if it has one. */
function sourcesPart(doc: PrimeDocument) {
  return doc.package.parts.find((p) => /^customXml\/item\d+\.xml$/.test(p.name) && decode(p.data).includes("<b:Sources"));
}

const tagText = (xml: string, tag: string) => {
  const m = new RegExp(`<b:${tag}>([\\s\\S]*?)</b:${tag}>`).exec(xml);
  return m ? unesc(m[1].trim()) : undefined;
};

/** The sources saved in the document, and its chosen style. */
export function readSources(doc: PrimeDocument): { style: CitationStyle; sources: Source[] } {
  const part = sourcesPart(doc);
  if (!part) return { style: "APA", sources: [] };
  const xml = decode(part.data);
  const name = /StyleName="([^"]*)"/.exec(xml)?.[1] ?? "";
  const style: CitationStyle = /IEEE/i.test(name) ? "IEEE" : /Harvard/i.test(name) ? "Harvard" : "APA";
  const sources: Source[] = [];
  for (const m of xml.matchAll(/<b:Source>([\s\S]*?)<\/b:Source>/g)) {
    const s = m[1];
    const authorBlock = /<b:Author>\s*<b:Author>([\s\S]*?)<\/b:Author>\s*<\/b:Author>/.exec(s)?.[1] ?? "";
    const authors = [...authorBlock.matchAll(/<b:Person>([\s\S]*?)<\/b:Person>/g)].map((p) => ({ last: tagText(p[1], "Last") ?? "", first: tagText(p[1], "First") }));
    const corporate = tagText(authorBlock, "Corporate");
    const type = (tagText(s, "SourceType") ?? "Book") as SourceType;
    const y = tagText(s, "YearAccessed");
    const accessed = y ? `${y}-${String(monthNumber(tagText(s, "MonthAccessed") ?? "1")).padStart(2, "0")}-${(tagText(s, "DayAccessed") ?? "1").padStart(2, "0")}` : undefined;
    sources.push({
      tag: tagText(s, "Tag") ?? "", type: ["Book", "JournalArticle", "InternetSite", "Report"].includes(type) ? type : "Book",
      authors: corporate ? [{ last: corporate }] : authors, title: tagText(s, "Title") ?? "", year: tagText(s, "Year"),
      publisher: tagText(s, "Publisher") ?? tagText(s, "InternetSiteTitle") ?? tagText(s, "Institution"), city: tagText(s, "City"),
      journal: tagText(s, "JournalName"), volume: tagText(s, "Volume"), issue: tagText(s, "Issue"), pages: tagText(s, "Pages"),
      url: tagText(s, "URL"), accessed,
    });
  }
  return { style, sources };
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
function monthNumber(m: string): number {
  const n = Number(m);
  if (n >= 1 && n <= 12) return n;
  const i = MONTHS.findIndex((x) => x.toLowerCase().startsWith(m.toLowerCase().slice(0, 3)));
  return i >= 0 ? i + 1 : 1;
}

function sourceXml(s: Source, i: number): string {
  const t = (tag: string, v?: string) => (v ? `<b:${tag}>${esc(v)}</b:${tag}>` : "");
  const people = s.authors.map((p) => `<b:Person><b:Last>${esc(p.last)}</b:Last>${p.first ? `<b:First>${esc(p.first)}</b:First>` : ""}</b:Person>`).join("");
  const [ay, am, ad] = (s.accessed ?? "").split("-");
  const guid = `{${(i + 1).toString(16).padStart(8, "0").toUpperCase()}-5052-494D-4557-4F5244534F55}`;
  return `<b:Source>${t("Tag", s.tag)}${t("SourceType", s.type)}<b:Guid>${guid}</b:Guid>`
    + (people ? `<b:Author><b:Author><b:NameList>${people}</b:NameList></b:Author></b:Author>` : "")
    + t("Title", s.title) + t("Year", s.year) + t("City", s.city)
    + (s.type === "InternetSite" ? t("InternetSiteTitle", s.publisher) : s.type === "Report" ? t("Institution", s.publisher) : t("Publisher", s.publisher))
    + t("JournalName", s.journal) + t("Volume", s.volume) + t("Issue", s.issue) + t("Pages", s.pages) + t("URL", s.url)
    + (ay ? t("YearAccessed", ay) + t("MonthAccessed", MONTHS[Number(am) - 1]) + t("DayAccessed", String(Number(ad))) : "")
    + `<b:RefOrder>${i + 1}</b:RefOrder></b:Source>`;
}

/** Saves the sources and style into the document's bibliography part (created the first time). */
export function writeSources(doc: PrimeDocument, style: CitationStyle, sources: Source[]) {
  const st = STYLE_FILES[style];
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources SelectedStyle="${esc(st.file)}" StyleName="${esc(st.name)}" Version="${st.version}" xmlns:b="${B}" xmlns="${B}">${sources.map(sourceXml).join("")}</b:Sources>`;
  const part = sourcesPart(doc);
  if (part) { part.data = encode(xml); return; }
  // A new bibliography part, as Word makes it: the item, its properties, and the link from the document.
  let n = 1;
  while (doc.package.parts.some((p) => p.name === `customXml/item${n}.xml`)) n++;
  const item = `customXml/item${n}.xml`;
  const props = `customXml/itemProps${n}.xml`;
  doc.package.parts.push(
    { name: item, data: encode(xml) },
    { name: props, data: encode(`<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{${crypto.randomUUID().toUpperCase()}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${B}"/></ds:schemaRefs></ds:datastoreItem>`) },
    { name: `customXml/_rels/item${n}.xml.rels`, data: encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${n}.xml"/></Relationships>`) },
  );
  doc.package.contentTypes.overrides[props] = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
  let id = 1;
  while (doc.package.preservedRels.some((r) => r.id === `rIdBib${id}`)) id++;
  doc.package.preservedRels.push({ id: `rIdBib${id}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", target: item });
}

/** A short unique tag for a new source, Word-style ("Men24"). */
export function newTag(s: Pick<Source, "authors" | "year" | "title">, existing: Source[]): string {
  const base = `${(s.authors[0]?.last || s.title || "Src").replace(/[^A-Za-z]/g, "").slice(0, 3)}${(s.year ?? "").slice(-2)}` || "Src";
  let tag = base;
  for (let i = 2; existing.some((x) => x.tag === tag); i++) tag = `${base}${i}`;
  return tag;
}

// --- formatting ---------------------------------------------------------------------------------

const t = (text: string, props?: RunProps): Inline => ({ type: "text", text, ...(props ? { props } : {}) });
const it = (text: string) => t(text, { italic: true });
const initials = (first?: string) => (first ?? "").split(/[\s-]+/).filter(Boolean).map((w) => `${w[0].toUpperCase()}.`).join(" ");

function namesInText(a: Person[], style: CitationStyle): string {
  if (!a.length) return "";
  const and = style === "APA" ? "&" : "and";
  if (a.length === 1) return a[0].last;
  if (a.length === 2) return `${a[0].last} ${and} ${a[1].last}`;
  return `${a[0].last} et al.`;
}

/** The text of a citation: (Mensah & Owusu, 2024, p. 12), or [3] in IEEE. */
export function citationText(s: Source | undefined, style: CitationStyle, number: number, page?: string): string {
  if (!s) return style === "IEEE" ? `[${number}]` : "(source missing)";
  if (style === "IEEE") return page ? `[${number}, p. ${page}]` : `[${number}]`;
  const who = namesInText(s.authors, style) || s.title;
  const year = s.year || "n.d.";
  return `(${who}, ${year}${page ? `, p. ${page}` : ""})`;
}

function authorsApa(a: Person[]): string {
  const one = (p: Person) => (p.first ? `${p.last}, ${initials(p.first)}` : p.last);
  if (a.length === 1) return one(a[0]);
  if (a.length <= 20) return `${a.slice(0, -1).map(one).join(", ")}, & ${one(a[a.length - 1])}`;
  return `${a.slice(0, 19).map(one).join(", ")}, … ${one(a[a.length - 1])}`;
}
function authorsHarvard(a: Person[]): string {
  const one = (p: Person) => (p.first ? `${p.last}, ${initials(p.first)}` : p.last);
  if (a.length === 1) return one(a[0]);
  if (a.length <= 3) return `${a.slice(0, -1).map(one).join(", ")} and ${one(a[a.length - 1])}`;
  return `${one(a[0])} et al.`;
}
function authorsIeee(a: Person[]): string {
  const one = (p: Person) => (p.first ? `${initials(p.first)} ${p.last}` : p.last);
  if (a.length <= 2) return a.map(one).join(" and ");
  if (a.length <= 6) return `${a.slice(0, -1).map(one).join(", ")}, and ${one(a[a.length - 1])}`;
  return `${one(a[0])} et al.`;
}
const accessedLong = (d?: string) => {
  if (!d) return "";
  const [y, m, day] = d.split("-").map(Number);
  return `${day} ${MONTHS[m - 1]} ${y}`;
};
const end = (s: string) => (/[.?!]$/.test(s) ? s : `${s}.`);

/** One reference-list entry. */
export function referenceInlines(s: Source, style: CitationStyle, number: number): Inline[] {
  const year = s.year || "n.d.";
  const vol = s.volume ? `${s.volume}${s.issue ? `(${s.issue})` : ""}` : "";
  if (style === "APA") {
    // The author element ends in a full stop (a person's initials already do).
    const who = s.authors.length ? `${end(authorsApa(s.authors))} ` : "";
    if (s.type === "JournalArticle") {
      return [t(`${who}(${year}). ${end(s.title)} `), it(s.journal ?? ""), ...(vol ? [t(", "), it(s.volume ?? ""), t(s.issue ? `(${s.issue})` : "")] : []), t(`${s.pages ? `, ${s.pages}` : ""}.${s.url ? ` ${s.url}` : ""}`)];
    }
    if (s.type === "InternetSite") return [t(`${who}(${year}). `), it(s.title), t(`. ${s.publisher ? `${s.publisher}. ` : ""}${s.url ?? ""}`)];
    return [t(`${who}(${year}). `), it(s.title), t(`. ${s.publisher ? `${end(s.publisher)}` : ""}`)];
  }
  if (style === "Harvard") {
    const who = s.authors.length ? `${authorsHarvard(s.authors)} ` : "";
    if (s.type === "JournalArticle") return [t(`${who}(${year}) '${s.title}', `), it(s.journal ?? ""), t(`${vol ? `, ${vol}` : ""}${s.pages ? `, pp. ${s.pages}` : ""}.`)];
    if (s.type === "InternetSite") return [t(`${who}(${year}) `), it(s.title), t(`. ${s.url ? `Available at: ${s.url} ` : ""}${s.accessed ? `(Accessed: ${accessedLong(s.accessed)})` : ""}.`.replace(/ \.$/, "."))];
    return [t(`${who}(${year}) `), it(s.title), t(`. ${s.city ? `${s.city}: ` : ""}${s.publisher ? `${s.publisher}.` : ""}`)];
  }
  // IEEE: numbered, in the order cited.
  const who = s.authors.length ? `${authorsIeee(s.authors)}, ` : "";
  if (s.type === "JournalArticle") {
    return [t(`[${number}]\t${who}"${s.title}," `), it(s.journal ?? ""), t(`${s.volume ? `, vol. ${s.volume}` : ""}${s.issue ? `, no. ${s.issue}` : ""}${s.pages ? `, pp. ${s.pages}` : ""}, ${year}.`)];
  }
  if (s.type === "InternetSite") return [t(`[${number}]\t${who}"${s.title}." ${s.publisher ? `${s.publisher}. ` : ""}${s.url ?? ""}${s.accessed ? ` (accessed ${accessedLong(s.accessed)})` : ""}.`)];
  return [t(`[${number}]\t${who}`), it(s.title), t(`. ${s.city ? `${s.city}: ` : ""}${s.publisher ? `${s.publisher}, ` : ""}${year}.`)];
}

// --- fields in the document ---------------------------------------------------------------------

/** Tags cited in the document, in order of first citation (IEEE numbers). */
export function citedTags(state: EditorState): string[] {
  const out: string[] = [];
  let instr = "";
  state.doc.descendants((n) => {
    if (n.type.name === "field_char" && n.attrs.kind === "begin") instr = "";
    if (n.type.name === "instr") instr += n.attrs.text;
    if (n.type.name === "field_char" && n.attrs.kind === "separate") {
      const m = /^\s*CITATION\s+(\S+)/.exec(instr);
      if (m && !out.includes(m[1])) out.push(m[1]);
    }
    return true;
  });
  return out;
}

/** Inserts a citation of `tag` at the cursor (with a page, if given). */
export function insertCitation(doc: () => PrimeDocument, tag: string, page?: string): Command {
  return (state, dispatch) => {
    const { style, sources } = readSources(doc());
    const source = sources.find((s) => s.tag === tag);
    if (!source) return false;
    if (!dispatch) return true;
    const cited = citedTags(state);
    const number = cited.includes(tag) ? cited.indexOf(tag) + 1 : cited.length + 1;
    const instr = `CITATION ${tag} \\l 2057${page ? ` \\p ${page}` : ""}`;
    const marks = state.selection.$from.marks().filter((m) => m.type.name !== "field");
    const s = state.schema;
    const nodes = [
      s.nodes.field_char.create({ kind: "begin", props: {} }, null, marks), s.nodes.instr.create({ text: ` ${instr} `, props: {} }, null, marks),
      s.nodes.field_char.create({ kind: "separate", props: {} }, null, marks), s.text(citationText(source, style, number, page), marks),
      s.nodes.field_char.create({ kind: "end", props: {} }, null, marks),
    ];
    // Right after a word, a space first (Word separates citations from the text the same way).
    const before = state.doc.textBetween(Math.max(state.selection.$from.start(), state.selection.from - 1), state.selection.from);
    if (before && !/[\s(\[]/.test(before)) nodes.unshift(s.text(" ", marks));
    dispatch(state.tr.replaceWith(state.selection.from, state.selection.to, nodes).scrollIntoView());
    return true;
  };
}

/** The reference list's paragraphs: the sources cited (all sources if none are cited yet), in the style's order. */
export function bibliographyBlocks(state: EditorState, doc: PrimeDocument): Block[] {
  const { style, sources } = readSources(doc);
  const cited = citedTags(state);
  const list = (cited.length ? cited.map((tag) => sources.find((s) => s.tag === tag)).filter((s): s is Source => !!s) : sources);
  const ordered = style === "IEEE" ? list : [...list].sort((a, b) =>
    (a.authors[0]?.last ?? a.title).localeCompare(b.authors[0]?.last ?? b.title) || (a.year ?? "").localeCompare(b.year ?? ""));
  ensureBibliographyStyle(doc);
  const entries = ordered.map((s): Block => ({
    type: "paragraph",
    props: style === "IEEE"
      ? { style: "Bibliography", indent: { left: 36, hanging: 36 }, extra: [{ name: "w:tabs", xml: '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>' }] }
      : { style: "Bibliography", indent: { left: 36, hanging: 36 } },
    content: referenceInlines(s, style, cited.indexOf(s.tag) + 1 || list.indexOf(s) + 1),
  }));
  if (!entries.length) entries.push({ type: "paragraph", props: { style: "Bibliography" }, content: [t("There are no sources in the current document.")] });
  // The whole list is one BIBLIOGRAPHY field, from the first entry to the last.
  const first = entries[0] as Extract<Block, { type: "paragraph" }>;
  const last = entries[entries.length - 1] as Extract<Block, { type: "paragraph" }>;
  first.content = [{ type: "fieldChar", kind: "begin" }, { type: "instr", text: " BIBLIOGRAPHY " }, { type: "fieldChar", kind: "separate" }, ...first.content];
  last.content = [...last.content, { type: "fieldChar", kind: "end" }];
  return entries;
}

function ensureBibliographyStyle(doc: PrimeDocument) {
  if (!doc.styles.styles.some((s) => s.id === "Bibliography")) {
    doc.styles.styles.push({ id: "Bibliography", type: "paragraph", name: "Bibliography", basedOn: "Normal", next: "Normal" });
  }
}

/** Inserts the reference list, with a "References" heading, at the cursor. */
export function insertBibliography(doc: () => PrimeDocument): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const heading: Block = { type: "paragraph", props: { style: "Heading1" }, content: [t("References")] };
    const nodes = storyToProseMirror([heading, ...bibliographyBlocks(state, doc())], state.schema).content;
    const $c = state.selection.$from;
    const tr = state.tr;
    if ($c.depth === 0) tr.insert($c.pos, nodes);
    else if ($c.parent.content.size === 0) tr.replaceWith($c.before($c.depth), $c.after($c.depth), nodes);
    else tr.insert($c.after($c.depth), nodes);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

