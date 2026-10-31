/** styles.xml and numbering.xml <-> PDM StyleSheet and Lists. */
import type { ListDefinition, ListInstance, ListLevel, Lists, Style, StyleSheet } from "@prime/pdm";
import { ABSTRACT_NUM_ORDER, LVL_ORDER, NUMBERING_ORDER, STYLE_ORDER } from "./ooxml.js";
import { readParagraphProps, readRunProps, readTableProps, writeParagraphProps, writeRunProps, writeTableProps } from "./props.js";
import {
  ROOT_NAMESPACES, attr, attrString, attrs, child, el, elements, escAttr, extraItems, type Item, onlyAttrs,
  ordered, preserve, qname,
} from "./xml.js";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// --- styles --------------------------------------------------------------------------------

export function readStyles(root: Element | undefined): StyleSheet {
  const sheet: StyleSheet = { defaults: {}, styles: [] };
  if (!root) return sheet;
  for (const c of elements(root)) {
    const n = qname(c);
    if (n === "w:docDefaults" && elements(c).every((d) => ["w:rPrDefault", "w:pPrDefault"].includes(qname(d)))) {
      const rPr = child(c, "w:rPrDefault");
      const pPr = child(c, "w:pPrDefault");
      if (rPr) sheet.defaults.run = readRunProps(child(rPr, "w:rPr")) ?? {};
      if (pPr) sheet.defaults.paragraph = readParagraphProps(child(pPr, "w:pPr")).props;
    } else if (n === "w:style") sheet.styles.push(readStyle(c));
    else (sheet.extra ??= []).push(preserve(c));
  }
  return sheet;
}

function readStyle(e: Element): Style {
  const s: Style = { id: attr(e, "styleId") ?? "", type: (attr(e, "type") ?? "paragraph") as Style["type"] };
  if (attr(e, "default") !== undefined) s.isDefault = ["1", "true", "on"].includes(attr(e, "default")!);
  if (attr(e, "customStyle") !== undefined) s.custom = ["1", "true", "on"].includes(attr(e, "customStyle")!);
  for (const c of elements(e)) {
    const n = qname(c);
    const simple = onlyAttrs(c, ["val"]);
    if (n === "w:name" && simple) s.name = attr(c, "val");
    else if (n === "w:basedOn" && simple) s.basedOn = attr(c, "val");
    else if (n === "w:next" && simple) s.next = attr(c, "val");
    else if (n === "w:link" && simple) s.link = attr(c, "val");
    else if (n === "w:pPr" && !child(c, "w:sectPr")) s.paragraph = readParagraphProps(c).props;
    else if (n === "w:rPr") s.run = readRunProps(c);
    else if (n === "w:tblPr") s.table = readTableProps(c);
    else (s.extra ??= []).push(preserve(c));
  }
  return s;
}

export function writeStyles(sheet: StyleSheet): string {
  let defaults = "";
  if (sheet.defaults.run || sheet.defaults.paragraph) {
    const r = sheet.defaults.run ? `<w:rPrDefault>${writeRunProps(sheet.defaults.run) || "<w:rPr/>"}</w:rPrDefault>` : "";
    const p = sheet.defaults.paragraph
      ? `<w:pPrDefault>${writeParagraphProps(sheet.defaults.paragraph) || "<w:pPr/>"}</w:pPrDefault>` : "";
    defaults = `<w:docDefaults>${r}${p}</w:docDefaults>`;
  }
  const extras = (sheet.extra ?? []).map((x) => x.xml).join("");
  const styles = sheet.styles.map(writeStyle).join("");
  return `${XML_DECL}<w:styles${attrString(ROOT_NAMESPACES, "")}>${defaults}${extras}${styles}</w:styles>`;
}

function writeStyle(s: Style): string {
  const items: Item[] = [];
  const v = (name: string, value: string | undefined) => value !== undefined && items.push([name, el(name, { val: value })]);
  v("w:name", s.name);
  v("w:basedOn", s.basedOn);
  v("w:next", s.next);
  v("w:link", s.link);
  if (s.paragraph) items.push(["w:pPr", writeParagraphProps(s.paragraph) || "<w:pPr/>"]);
  if (s.run) items.push(["w:rPr", writeRunProps(s.run) || "<w:rPr/>"]);
  if (s.table) items.push(["w:tblPr", writeTableProps(s.table)]);
  items.push(...extraItems(s.extra));
  let a = ` w:type="${escAttr(s.type)}"`;
  if (s.isDefault !== undefined) a += ` w:default="${s.isDefault ? 1 : 0}"`;
  if (s.custom !== undefined) a += ` w:customStyle="${s.custom ? 1 : 0}"`;
  a += ` w:styleId="${escAttr(s.id)}"`;
  return `<w:style${a}>${ordered(STYLE_ORDER, items)}</w:style>`;
}

// --- numbering -----------------------------------------------------------------------------

export function readNumbering(root: Element | undefined, warn: (m: string) => void): Lists {
  const lists: Lists = { definitions: [], instances: [] };
  if (!root) return lists;
  for (const c of elements(root)) {
    const n = qname(c);
    if (n === "w:abstractNum") lists.definitions.push(readDefinition(c));
    else if (n === "w:num") lists.instances.push(readInstance(c));
    else {
      if (n === "w:numPicBullet") warn("Picture bullets are kept but shown as their fallback symbol.");
      (lists.extra ??= []).push(preserve(c));
    }
  }
  return lists;
}

function readDefinition(e: Element): ListDefinition {
  const d: ListDefinition = { id: attr(e, "abstractNumId") ?? "", levels: [] };
  for (const c of elements(e)) {
    if (qname(c) === "w:lvl") d.levels.push(readLevel(c));
    else (d.extra ??= []).push(preserve(c));
  }
  return d;
}

function readLevel(e: Element): ListLevel {
  const { ilvl, ...rest } = attrs(e);
  const l: ListLevel = { level: Number(ilvl ?? 0) };
  if (Object.keys(rest).length) l.attrs = rest;
  for (const c of elements(e)) {
    const n = qname(c);
    const simple = onlyAttrs(c, ["val"]);
    if (n === "w:start" && simple) l.start = Number(attr(c, "val"));
    else if (n === "w:numFmt" && simple) l.format = attr(c, "val");
    else if (n === "w:lvlText" && simple) l.text = attr(c, "val");
    else if (n === "w:lvlJc" && simple) l.align = attr(c, "val");
    else if (n === "w:pPr") l.paragraph = readParagraphProps(c).props;
    else if (n === "w:rPr") l.run = readRunProps(c);
    else (l.extra ??= []).push(preserve(c));
  }
  return l;
}

function readInstance(e: Element): ListInstance {
  const inst: ListInstance = { id: attr(e, "numId") ?? "", definitionId: "" };
  for (const c of elements(e)) {
    if (qname(c) === "w:abstractNumId" && onlyAttrs(c, ["val"])) inst.definitionId = attr(c, "val") ?? "";
    else (inst.extra ??= []).push(preserve(c));
  }
  return inst;
}

export function writeNumbering(lists: Lists): string {
  const items: Item[] = [...extraItems(lists.extra)];
  for (const d of lists.definitions) {
    const inner: Item[] = [...extraItems(d.extra), ...d.levels.map((l): Item => ["w:lvl", writeLevel(l)])];
    items.push(["w:abstractNum", `<w:abstractNum w:abstractNumId="${escAttr(d.id)}">${ordered(ABSTRACT_NUM_ORDER, inner)}</w:abstractNum>`]);
  }
  for (const i of lists.instances) {
    const inner = el("w:abstractNumId", { val: i.definitionId }) + (i.extra ?? []).map((x) => x.xml).join("");
    items.push(["w:num", `<w:num w:numId="${escAttr(i.id)}">${inner}</w:num>`]);
  }
  return `${XML_DECL}<w:numbering${attrString(ROOT_NAMESPACES, "")}>${ordered(NUMBERING_ORDER, items)}</w:numbering>`;
}

function writeLevel(l: ListLevel): string {
  const items: Item[] = [];
  if (l.start !== undefined) items.push(["w:start", el("w:start", { val: String(l.start) })]);
  if (l.format !== undefined) items.push(["w:numFmt", el("w:numFmt", { val: l.format })]);
  if (l.text !== undefined) items.push(["w:lvlText", el("w:lvlText", { val: l.text })]);
  if (l.align !== undefined) items.push(["w:lvlJc", el("w:lvlJc", { val: l.align })]);
  if (l.paragraph) items.push(["w:pPr", writeParagraphProps(l.paragraph) || "<w:pPr/>"]);
  if (l.run) items.push(["w:rPr", writeRunProps(l.run) || "<w:rPr/>"]);
  items.push(...extraItems(l.extra));
  return `<w:lvl${attrString({ ilvl: String(l.level), ...l.attrs })}>${ordered(LVL_ORDER, items)}</w:lvl>`;
}

