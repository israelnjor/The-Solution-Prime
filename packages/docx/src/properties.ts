/**
 * Document properties (File → Info): title, subject, author, keywords and the rest of the package's
 * core properties (docProps/core.xml), which is otherwise carried through as it is. Reading takes
 * them from that part; writing replaces the part (or adds it, with its package relationship and
 * content type, to a document that has none).
 */
import type { PrimeDocument } from "@prime/pdm";
import { fromBase64, toBase64 } from "./xml.js";

export interface DocumentProperties {
  title?: string;
  subject?: string;
  /** dc:creator */
  author?: string;
  keywords?: string;
  /** dc:description (Word's "Comments") */
  comments?: string;
  category?: string;
  status?: string;
  lastModifiedBy?: string;
  /** ISO dates (dcterms:created, dcterms:modified). */
  created?: string;
  modified?: string;
  revision?: number;
}

export const CORE_PART = "docProps/core.xml";
const CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
export const CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

const FIELDS: [keyof DocumentProperties, string][] = [
  ["title", "dc:title"], ["subject", "dc:subject"], ["author", "dc:creator"], ["keywords", "cp:keywords"],
  ["comments", "dc:description"], ["lastModifiedBy", "cp:lastModifiedBy"], ["revision", "cp:revision"],
  ["created", "dcterms:created"], ["modified", "dcterms:modified"], ["category", "cp:category"], ["status", "cp:contentStatus"],
];

const decode = (b64: string) => new TextDecoder().decode(fromBase64(b64));
const encode = (s: string) => toBase64(new TextEncoder().encode(s));
const unescape = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** The document's core properties (none: {}). */
export function readProperties(doc: PrimeDocument): DocumentProperties {
  const part = doc.package.parts.find((p) => p.name === CORE_PART);
  if (!part) return {};
  const xml = decode(part.data);
  const out: DocumentProperties = {};
  for (const [key, tag] of FIELDS) {
    const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`).exec(xml);
    if (!m) continue;
    const v = unescape(m[1]).trim();
    if (key === "revision") { if (Number(v) > 0) out.revision = Number(v); } else (out as Record<string, string>)[key] = v;
  }
  return out;
}

/** The document with these core properties (empty strings remove a property); other properties kept. */
export function withProperties(doc: PrimeDocument, props: DocumentProperties): PrimeDocument {
  const parts = doc.package.parts;
  const part = parts.find((p) => p.name === CORE_PART);
  let xml = part ? decode(part.data)
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`;
  for (const [key, tag] of FIELDS) {
    if (!(key in props)) continue;
    const value = props[key];
    const re = new RegExp(`<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>|<${tag}(?:\\s[^>]*)?/>`);
    const text = value === undefined || value === "" ? "" : String(value);
    const dated = tag.startsWith("dcterms:") ? ' xsi:type="dcterms:W3CDTF"' : "";
    const element = text ? `<${tag}${dated}>${escape(text)}</${tag}>` : "";
    if (re.test(xml)) xml = xml.replace(re, element);
    else if (element) xml = xml.replace("</cp:coreProperties>", `${element}</cp:coreProperties>`);
  }
  const nextParts = part ? parts.map((p) => (p === part ? { ...p, data: encode(xml) } : p)) : [...parts, { name: CORE_PART, data: encode(xml) }];
  // A new part: its package relationship (in a package relationships part carried through) and content type.
  const rels = nextParts.find((p) => p.name === "_rels/.rels");
  if (!part && rels) {
    const r = decode(rels.data);
    if (!r.includes(CORE_REL)) {
      let n = 1;
      while (r.includes(`Id="rIdCore${n}"`)) n++;
      const updated = r.replace("</Relationships>", `<Relationship Id="rIdCore${n}" Type="${CORE_REL}" Target="${CORE_PART}"/></Relationships>`);
      nextParts.splice(nextParts.indexOf(rels), 1, { ...rels, data: encode(updated) });
    }
  }
  return {
    ...doc,
    package: {
      ...doc.package,
      parts: nextParts,
      contentTypes: { ...doc.package.contentTypes, overrides: { ...doc.package.contentTypes.overrides, [CORE_PART]: doc.package.contentTypes.overrides[CORE_PART] ?? CORE_TYPE } },
    },
  };
}
