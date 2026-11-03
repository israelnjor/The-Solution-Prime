/** Minimal XML helpers over @xmldom/xmldom. Pure: no Node or browser-only APIs. */
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type { Attrs, Preserved } from "@prime/pdm";

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  v: "urn:schemas-microsoft-com:vml",
  xml: "http://www.w3.org/XML/1998/namespace",
  pr: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

const PREFIX_BY_NS: Record<string, string> = Object.fromEntries(Object.entries(NS).map(([p, u]) => [u, p]));

export function parseXml(text: string): Document {
  const errors: string[] = [];
  const doc = new DOMParser({
    errorHandler: { warning: () => undefined, error: (m) => errors.push(m), fatalError: (m) => errors.push(m) },
  }).parseFromString(text.replace(/^﻿/, ""), "application/xml");
  if (!doc.documentElement || errors.length) throw new Error(`XML parse error: ${errors[0] ?? "no root element"}`);
  return doc;
}

/** Canonical qualified name ("w:p"), independent of the prefix the file happened to use. */
export function qname(el: Element): string {
  const prefix = el.namespaceURI ? PREFIX_BY_NS[el.namespaceURI] : undefined;
  return prefix ? `${prefix}:${el.localName}` : el.nodeName;
}

export function elements(el: Element | Document): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

export function child(el: Element, name: string): Element | undefined {
  return elements(el).find((c) => qname(c) === name);
}

export function attr(el: Element | undefined, local: string, ns: string = NS.w): string | undefined {
  const node = el?.getAttributeNodeNS(ns, local);
  return node ? node.value : undefined;
}

/** All attributes keyed by local name (w: namespace) or qualified name (others). */
export function attrs(el: Element): Attrs {
  const out: Attrs = {};
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.name.startsWith("xmlns")) continue;
    out[a.namespaceURI === NS.w ? a.localName : a.name] = a.value;
  }
  return out;
}

/** True when the element carries no attributes other than the given w: local names, and no children. */
export function onlyAttrs(el: Element, allowed: string[]): boolean {
  if (elements(el).length) return false;
  return Object.keys(attrs(el)).every((k) => allowed.includes(k));
}

export function onOff(el: Element): boolean {
  const v = attr(el, "val");
  return v === undefined || !["0", "false", "off"].includes(v);
}

const serializer = new XMLSerializer();

export function preserve(el: Element): Preserved {
  const xml = serializer.serializeToString(el);
  // The serializer re-declares namespaces on each fragment; drop those every generated part root declares.
  const end = xml.indexOf(">");
  const open = xml.slice(0, end).replace(/ xmlns:(\w+)="([^"]*)"/g, (m, p, uri) => (ROOT_NAMESPACES[`xmlns:${p}`] === uri ? "" : m));
  return { name: qname(el), xml: open + xml.slice(end) };
}

export function textContent(el: Element): string {
  return el.textContent ?? "";
}

// --- writing -------------------------------------------------------------------------------

/** Characters XML 1.0 does not allow (Word refuses a file containing any). */
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function escText(s: string): string {
  return s.replace(INVALID_XML, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function escAttr(s: string): string {
  return escText(s).replace(/"/g, "&quot;");
}

/** Attribute string from a bag of local names (written in the w: namespace unless already qualified). */
export function attrString(a: Attrs | undefined, prefix = "w"): string {
  if (!a) return "";
  return Object.entries(a)
    .map(([k, v]) => ` ${k.includes(":") ? k : `${prefix}:${k}`}="${escAttr(v)}"`)
    .join("");
}

export function el(name: string, a?: Attrs, inner?: string): string {
  const at = attrString(a);
  return inner ? `<${name}${at}>${inner}</${name}>` : `<${name}${at}/>`;
}

export type Item = [name: string, xml: string];

/** Join child items in schema order. Unknown names keep their relative position at the end. */
export function ordered(order: readonly string[], items: Item[]): string {
  const rank = (n: string) => {
    const i = order.indexOf(n);
    return i < 0 ? order.length : i;
  };
  return items
    .map((it, i) => ({ it, i }))
    .sort((x, y) => rank(x.it[0]) - rank(y.it[0]) || x.i - y.i)
    .map((x) => x.it[1])
    .join("");
}

export function extraItems(extra: Preserved[] | undefined): Item[] {
  return (extra ?? []).map((p) => [p.name, p.xml]);
}

/** Namespace declarations written on every part root we generate. */
export const ROOT_NAMESPACES: Attrs = {
  "xmlns:w": NS.w,
  "xmlns:r": NS.r,
  "xmlns:wp": NS.wp,
  "xmlns:a": NS.a,
  "xmlns:pic": NS.pic,
  "xmlns:mc": NS.mc,
  "xmlns:v": NS.v,
  "xmlns:o": "urn:schemas-microsoft-com:office:office",
  "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
  "xmlns:w10": "urn:schemas-microsoft-com:office:word",
  "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
  "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
  "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
  "mc:Ignorable": "w14 w15 wp14",
};

// --- base64 (btoa/atob exist in browsers and Node >= 16) ------------------------------------

export function toBase64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function fromBase64(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
