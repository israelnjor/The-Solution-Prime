/**
 * Property codecs: w:rPr, w:pPr, w:sectPr, w:tblPr, w:trPr, w:tcPr <-> PDM.
 *
 * Rule: an element becomes a typed field only when we understand all of it. Anything else
 * (unknown elements, or known elements with attributes we do not model) is kept as a
 * Preserved fragment and written back in schema order — so nothing is silently dropped.
 */
import type {
  Attrs, CellProps, FontSpec, Indent, ParagraphProps, Preserved, RowProps, RunProps, SectionProps, Spacing,
  TableProps,
} from "@prime/pdm";
import { PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TBLPR_ORDER, TCPR_ORDER, TRPR_ORDER, pt, twips } from "./ooxml.js";
import { attr, attrs, el, elements, extraItems, type Item, onlyAttrs, onOff, ordered, preserve, qname } from "./xml.js";

const val = (e: Element) => attr(e, "val");
const simple = (e: Element) => onlyAttrs(e, ["val"]);
const leaf = (e: Element) => elements(e).length === 0;
const push = (list: Preserved[] | undefined, e: Element): Preserved[] => [...(list ?? []), preserve(e)];
const onOffXml = (name: string, v: boolean): Item => [name, v ? el(name) : el(name, { val: "0" })];
const valXml = (name: string, v: string): Item => [name, el(name, { val: v })];

function wrap(name: string, order: readonly string[], items: Item[], always = false): string {
  if (!items.length) return always ? `<${name}/>` : "";
  return `<${name}>${ordered(order, items)}</${name}>`;
}

// --- run properties ------------------------------------------------------------------------

const RUN_TOGGLES: Record<string, keyof RunProps> = {
  "w:b": "bold", "w:i": "italic", "w:caps": "caps", "w:smallCaps": "smallCaps",
  "w:strike": "strike", "w:dstrike": "doubleStrike", "w:vanish": "hidden",
};

export function readRunProps(rPr: Element | undefined): RunProps | undefined {
  if (!rPr) return undefined;
  const p: RunProps = {};
  const set = p as Record<string, unknown>;
  for (const c of elements(rPr)) {
    const n = qname(c);
    if (n in RUN_TOGGLES && simple(c)) set[RUN_TOGGLES[n]] = onOff(c);
    else if (n === "w:rStyle" && simple(c)) p.style = val(c);
    else if (n === "w:rFonts" && leaf(c)) p.font = attrs(c) as FontSpec;
    else if (n === "w:color" && simple(c) && val(c)) p.color = val(c)!.toUpperCase();
    else if (n === "w:sz" && simple(c) && val(c)) p.size = Number(val(c)) / 2;
    else if (n === "w:highlight" && simple(c)) p.highlight = val(c);
    else if (n === "w:u" && leaf(c)) p.underline = attrs(c);
    else if (n === "w:shd" && leaf(c)) p.shading = attrs(c);
    else if (n === "w:vertAlign" && simple(c)) p.vertAlign = val(c) as RunProps["vertAlign"];
    else p.extra = push(p.extra, c);
  }
  return p;
}

export function writeRunProps(p: RunProps | undefined): string {
  if (!p) return "";
  const items: Item[] = [];
  if (p.style !== undefined) items.push(valXml("w:rStyle", p.style));
  if (p.font) items.push(["w:rFonts", el("w:rFonts", p.font as Attrs)]);
  for (const [name, key] of Object.entries(RUN_TOGGLES)) {
    const v = p[key];
    if (typeof v === "boolean") items.push(onOffXml(name, v));
  }
  if (p.color !== undefined) items.push(valXml("w:color", p.color));
  if (p.size !== undefined) items.push(valXml("w:sz", String(Math.round(p.size * 2))));
  if (p.highlight !== undefined) items.push(valXml("w:highlight", p.highlight));
  if (p.underline) items.push(["w:u", el("w:u", p.underline)]);
  if (p.shading) items.push(["w:shd", el("w:shd", p.shading)]);
  if (p.vertAlign !== undefined) items.push(valXml("w:vertAlign", p.vertAlign));
  items.push(...extraItems(p.extra));
  return wrap("w:rPr", RPR_ORDER, items);
}

// --- paragraph properties ------------------------------------------------------------------

const PARA_TOGGLES: Record<string, keyof ParagraphProps> = {
  "w:keepNext": "keepNext", "w:keepLines": "keepLines", "w:pageBreakBefore": "pageBreakBefore",
  "w:widowControl": "widowControl",
};

function readSpacing(e: Element): Spacing {
  const s: Spacing = {};
  for (const [k, v] of Object.entries(attrs(e))) {
    if (k === "before" || k === "after") s[k] = pt(v);
    else if (k === "line") s.line = Number(v);
    else if (k === "lineRule") s.lineRule = v;
    else (s.other ??= {})[k] = v;
  }
  return s;
}

function writeSpacing(s: Spacing): string {
  const a: Attrs = {};
  if (s.before !== undefined) a.before = twips(s.before);
  if (s.after !== undefined) a.after = twips(s.after);
  if (s.line !== undefined) a.line = String(s.line);
  if (s.lineRule !== undefined) a.lineRule = s.lineRule;
  return el("w:spacing", { ...a, ...s.other });
}

function readIndent(e: Element): Indent {
  const ind: Indent = {};
  for (const [k, v] of Object.entries(attrs(e))) {
    if (k === "left" || k === "right" || k === "firstLine" || k === "hanging") ind[k] = pt(v);
    else (ind.other ??= {})[k] = v;
  }
  return ind;
}

function writeIndent(ind: Indent): string {
  const a: Attrs = {};
  for (const k of ["left", "right", "firstLine", "hanging"] as const) if (ind[k] !== undefined) a[k] = twips(ind[k]!);
  return el("w:ind", { ...a, ...ind.other });
}

/** Reads w:pPr. A nested w:sectPr is returned separately: it belongs to the section, not the paragraph. */
export function readParagraphProps(pPr: Element | undefined): { props: ParagraphProps; sectPr?: Element } {
  const p: ParagraphProps = {};
  const set = p as Record<string, unknown>;
  let sectPr: Element | undefined;
  for (const c of pPr ? elements(pPr) : []) {
    const n = qname(c);
    if (n in PARA_TOGGLES && simple(c)) set[PARA_TOGGLES[n]] = onOff(c);
    else if (n === "w:pStyle" && simple(c)) p.style = val(c);
    else if (n === "w:jc" && simple(c)) p.align = val(c);
    else if (n === "w:outlineLvl" && simple(c)) p.outlineLevel = Number(val(c));
    else if (n === "w:spacing" && leaf(c)) p.spacing = readSpacing(c);
    else if (n === "w:ind" && leaf(c)) p.indent = readIndent(c);
    else if (n === "w:numPr" && elements(c).every((x) => ["w:ilvl", "w:numId"].includes(qname(x)) && simple(x))) {
      const ilvl = elements(c).find((x) => qname(x) === "w:ilvl");
      const numId = elements(c).find((x) => qname(x) === "w:numId");
      p.list = { numId: numId ? val(numId) ?? "0" : "", level: ilvl ? Number(val(ilvl)) : undefined };
    } else if (n === "w:rPr") p.markRun = readRunProps(c);
    else if (n === "w:sectPr") sectPr = c;
    else p.extra = push(p.extra, c);
  }
  return { props: p, sectPr };
}

export function writeParagraphProps(p: ParagraphProps | undefined, sectPrXml = ""): string {
  const items: Item[] = [];
  if (p) {
    if (p.style !== undefined) items.push(valXml("w:pStyle", p.style));
    for (const [name, key] of Object.entries(PARA_TOGGLES)) {
      const v = p[key];
      if (typeof v === "boolean") items.push(onOffXml(name, v));
    }
    if (p.list) {
      const inner = (p.list.level !== undefined ? el("w:ilvl", { val: String(p.list.level) }) : "")
        + (p.list.numId !== "" ? el("w:numId", { val: p.list.numId }) : "");
      items.push(["w:numPr", `<w:numPr>${inner}</w:numPr>`]);
    }
    if (p.spacing) items.push(["w:spacing", writeSpacing(p.spacing)]);
    if (p.indent) items.push(["w:ind", writeIndent(p.indent)]);
    if (p.align !== undefined) items.push(valXml("w:jc", p.align));
    if (p.outlineLevel !== undefined) items.push(valXml("w:outlineLvl", String(p.outlineLevel)));
    if (p.markRun) items.push(["w:rPr", writeRunProps(p.markRun) || "<w:rPr/>"]);
    items.push(...extraItems(p.extra));
  }
  if (sectPrXml) items.push(["w:sectPr", sectPrXml]);
  return wrap("w:pPr", PPR_ORDER, items);
}

// --- section properties --------------------------------------------------------------------

/** Reads w:sectPr, except header/footer references, which the reader resolves into stories. */
export function readSectionProps(sectPr: Element): SectionProps {
  const s: SectionProps = {};
  for (const c of elements(sectPr)) {
    const n = qname(c);
    if (n === "w:headerReference" || n === "w:footerReference") continue;
    if (n === "w:pgSz" && onlyAttrs(c, ["w", "h", "orient", "code"])) {
      s.pageWidth = pt(attr(c, "w"));
      s.pageHeight = pt(attr(c, "h"));
      s.orientation = attr(c, "orient") as SectionProps["orientation"];
      s.paperCode = attr(c, "code");
    } else if (n === "w:pgMar" && onlyAttrs(c, ["top", "right", "bottom", "left", "header", "footer", "gutter"])) {
      const m: NonNullable<SectionProps["margins"]> = {};
      for (const [k, v] of Object.entries(attrs(c))) m[k as keyof typeof m] = pt(v);
      s.margins = m;
    } else if (n === "w:type" && simple(c)) s.breakType = val(c);
    else if (n === "w:titlePg" && simple(c)) s.titlePage = onOff(c);
    else s.extra = push(s.extra, c);
  }
  return s;
}

export function writeSectionProps(s: SectionProps, references: Item[]): string {
  const items: Item[] = [...references];
  if (s.breakType !== undefined) items.push(valXml("w:type", s.breakType));
  if (s.pageWidth !== undefined || s.pageHeight !== undefined) {
    const a: Attrs = {};
    if (s.pageWidth !== undefined) a.w = twips(s.pageWidth);
    if (s.pageHeight !== undefined) a.h = twips(s.pageHeight);
    if (s.orientation) a.orient = s.orientation;
    if (s.paperCode) a.code = s.paperCode;
    items.push(["w:pgSz", el("w:pgSz", a)]);
  }
  if (s.margins) {
    const a: Attrs = {};
    for (const [k, v] of Object.entries(s.margins)) if (v !== undefined) a[k] = twips(v);
    items.push(["w:pgMar", el("w:pgMar", a)]);
  }
  if (s.titlePage !== undefined) items.push(onOffXml("w:titlePg", s.titlePage));
  items.push(...extraItems(s.extra));
  return wrap("w:sectPr", SECTPR_ORDER, items, true);
}

// --- table, row and cell properties --------------------------------------------------------

export function readTableProps(tblPr: Element | undefined): TableProps {
  const t: TableProps = {};
  for (const c of tblPr ? elements(tblPr) : []) {
    const n = qname(c);
    if (n === "w:tblStyle" && simple(c)) t.style = val(c);
    else if (n === "w:tblW" && leaf(c)) t.width = attrs(c);
    else if (n === "w:jc" && simple(c)) t.align = val(c);
    else t.extra = push(t.extra, c);
  }
  return t;
}

export function writeTableProps(t: TableProps): string {
  const items: Item[] = [];
  if (t.style !== undefined) items.push(valXml("w:tblStyle", t.style));
  if (t.width) items.push(["w:tblW", el("w:tblW", t.width)]);
  if (t.align !== undefined) items.push(valXml("w:jc", t.align));
  items.push(...extraItems(t.extra));
  return wrap("w:tblPr", TBLPR_ORDER, items, true);
}

export function readRowProps(trPr: Element | undefined): RowProps {
  const r: RowProps = {};
  for (const c of trPr ? elements(trPr) : []) {
    const n = qname(c);
    if (n === "w:tblHeader" && simple(c)) r.header = onOff(c);
    else if (n === "w:cantSplit" && simple(c)) r.cantSplit = onOff(c);
    else if (n === "w:trHeight" && leaf(c)) r.height = attrs(c);
    else r.extra = push(r.extra, c);
  }
  return r;
}

export function writeRowProps(r: RowProps): string {
  const items: Item[] = [];
  if (r.cantSplit !== undefined) items.push(onOffXml("w:cantSplit", r.cantSplit));
  if (r.height) items.push(["w:trHeight", el("w:trHeight", r.height)]);
  if (r.header !== undefined) items.push(onOffXml("w:tblHeader", r.header));
  items.push(...extraItems(r.extra));
  return wrap("w:trPr", TRPR_ORDER, items);
}

export function readCellProps(tcPr: Element | undefined): CellProps {
  const t: CellProps = {};
  for (const c of tcPr ? elements(tcPr) : []) {
    const n = qname(c);
    if (n === "w:tcW" && leaf(c)) t.width = attrs(c);
    else if (n === "w:gridSpan" && simple(c)) t.span = Number(val(c));
    else if (n === "w:vMerge" && simple(c)) t.vMerge = val(c) === "restart" ? "restart" : "continue";
    else if (n === "w:shd" && leaf(c)) t.shading = attrs(c);
    else if (n === "w:vAlign" && simple(c)) t.vAlign = val(c);
    else t.extra = push(t.extra, c);
  }
  return t;
}

export function writeCellProps(t: CellProps): string {
  const items: Item[] = [];
  if (t.width) items.push(["w:tcW", el("w:tcW", t.width)]);
  if (t.span !== undefined) items.push(valXml("w:gridSpan", String(t.span)));
  if (t.vMerge) items.push(["w:vMerge", t.vMerge === "restart" ? el("w:vMerge", { val: "restart" }) : el("w:vMerge")]);
  if (t.shading) items.push(["w:shd", el("w:shd", t.shading)]);
  if (t.vAlign !== undefined) items.push(valXml("w:vAlign", t.vAlign));
  items.push(...extraItems(t.extra));
  return wrap("w:tcPr", TCPR_ORDER, items);
}
