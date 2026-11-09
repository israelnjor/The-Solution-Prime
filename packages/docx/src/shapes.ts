/**
 * Shapes and text boxes (DrawingML wps:wsp in a floating wp:anchor): how they are drawn (preset
 * geometry, fill, line — from wps:spPr, or from the theme style Word gives new shapes), new ones as
 * Word writes them, and changes to their fill and line, made in their XML (which is what is saved).
 */
import type { ShapeStyle } from "@prime/pdm";
import { EMU_PER_PT } from "./ooxml.js";
import { child, elements, qname } from "./xml.js";

/** The Office theme's colours (when a document has no theme of its own). */
export const DEFAULT_THEME: Record<string, string> = {
  dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6", accent1: "4472C4", accent2: "ED7D31",
  accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", hlink: "0563C1", folHlink: "954F72",
};
const ALIAS: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };

/** The theme's colour scheme (a:clrScheme), from the theme part's XML. */
export function themeColors(themeXml: string | undefined): Record<string, string> {
  const out = { ...DEFAULT_THEME };
  if (!themeXml) return out;
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(themeXml)?.[0] ?? "";
  for (const m of scheme.matchAll(/<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>([\s\S]*?)<\/a:\1>/g)) {
    const v = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[2])?.[1] ?? /lastClr="([0-9A-Fa-f]{6})"/.exec(m[2])?.[1];
    if (v) out[m[1]] = v.toUpperCase();
  }
  return out;
}

const hex = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0").toUpperCase();

/** A colour element's RRGGBB (srgbClr, schemeClr, sysClr, prstClr), with its shade / tint / luminance changes. */
function colorOf(el: Element | undefined, theme: Record<string, string>): string | undefined {
  if (!el) return undefined;
  const c = elements(el).find((e) => /^a:(srgbClr|schemeClr|sysClr|prstClr)$/.test(qname(e)));
  if (!c) return undefined;
  const name = qname(c);
  let rgb = name === "a:srgbClr" ? c.getAttribute("val")
    : name === "a:schemeClr" ? theme[ALIAS[c.getAttribute("val") ?? ""] ?? c.getAttribute("val") ?? ""]
    : name === "a:sysClr" ? c.getAttribute("lastClr") ?? (c.getAttribute("val") === "window" ? "FFFFFF" : "000000")
    : ({ black: "000000", white: "FFFFFF", red: "FF0000", blue: "0000FF", green: "008000", yellow: "FFFF00", gray: "808080" } as Record<string, string>)[c.getAttribute("val") ?? ""];
  if (!rgb || !/^[0-9A-Fa-f]{6}$/.test(rgb)) return undefined;
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(rgb!.slice(i, i + 2), 16));
  for (const m of elements(c)) {
    const v = Number(m.getAttribute("val") ?? 100000) / 100000;
    const n = qname(m);
    if (n === "a:shade") [r, g, b] = [r * v, g * v, b * v];
    else if (n === "a:tint") [r, g, b] = [r + (255 - r) * (1 - v), g + (255 - g) * (1 - v), b + (255 - b) * (1 - v)];
    else if (n === "a:lumMod") [r, g, b] = [r * v, g * v, b * v];
    else if (n === "a:lumOff") [r, g, b] = [r + 255 * v, g + 255 * v, b + 255 * v];
  }
  rgb = hex(r) + hex(g) + hex(b);
  return rgb;
}

/**
 * How a shape is drawn: its spPr's geometry, fill and line; where spPr says nothing about the fill or
 * line, the theme style (wps:style fillRef / lnRef) Word gives new shapes (accent 1 fill, a darker line).
 */
export function readShapeStyle(wsp: Element, theme: Record<string, string>): ShapeStyle {
  const spPr = child(wsp, "wps:spPr");
  const style = child(wsp, "wps:style");
  const geometry = (spPr && child(spPr, "a:prstGeom")?.getAttribute("prst")) || "rect";
  const out: ShapeStyle = { geometry };
  // Fill.
  if (spPr && child(spPr, "a:noFill")) { /* none */ }
  else if (spPr && child(spPr, "a:solidFill")) out.fill = colorOf(child(spPr, "a:solidFill"), theme);
  else if (spPr && (child(spPr, "a:gradFill") || child(spPr, "a:pattFill"))) {
    // A gradient or pattern: its first colour.
    const g = child(spPr, "a:gradFill") ?? child(spPr, "a:pattFill")!;
    const first = g.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "gs")[0] ?? child(g, "a:fgClr");
    out.fill = colorOf(first, theme);
  } else if (style) {
    const ref = child(style, "a:fillRef");
    if (ref && ref.getAttribute("idx") !== "0") out.fill = colorOf(ref, theme);
  }
  // Line.
  const ln = spPr && child(spPr, "a:ln");
  const width = ln?.getAttribute("w") ? Number(ln.getAttribute("w")) / EMU_PER_PT : 0.75;
  const dash = ln && child(ln, "a:prstDash")?.getAttribute("val");
  if (ln && child(ln, "a:noFill")) { /* none */ }
  else if (ln && child(ln, "a:solidFill")) {
    const color = colorOf(child(ln, "a:solidFill"), theme);
    if (color) out.line = { color, width, ...(dash && dash !== "solid" ? { dash } : {}) };
  } else if (style) {
    const ref = child(style, "a:lnRef");
    const color = ref && ref.getAttribute("idx") !== "0" ? colorOf(ref, theme) : undefined;
    if (color) out.line = { color, width: ln?.getAttribute("w") ? width : 1, ...(dash && dash !== "solid" ? { dash } : {}) };
  }
  return out;
}

// --- new shapes and text boxes ------------------------------------------------------------------------

const emu = (pt: number) => Math.round(pt * EMU_PER_PT);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

export interface NewShape {
  /** A text box (with text) or a shape (preset geometry). */
  kind: "textbox" | "shape";
  geometry: string;
  width: number;
  height: number;
  /** From the column's left and the anchor paragraph's top, in pt. */
  x: number;
  y: number;
  fill?: string;
  line?: { color: string; width: number };
  /** Text box: its paragraphs' XML (w:p…). */
  textXml?: string;
  /** Unique among the document's drawings (wp:docPr id). */
  id: number;
  name: string;
}

/** A floating shape or text box as Word writes it (in front of text, square wrapping off: no wrap). */
export function newShapeXml(s: NewShape): string {
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : "<a:noFill/>";
  const line = s.line ? `<a:ln w="${emu(s.line.width)}"><a:solidFill><a:srgbClr val="${s.line.color}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  const isLine = s.geometry === "line" || s.geometry === "straightConnector1";
  const text = s.kind === "textbox"
    ? `<wps:txbx><w:txbxContent>${s.textXml ?? "<w:p/>"}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`
    : `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
  return `<w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + s.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(s.x)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(s.y)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${emu(s.width)}" cy="${emu(s.height)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
    + `<wp:docPr id="${s.id}" name="${esc(s.name)}"/><wp:cNvGraphicFramePr/>`
    + `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>`
    + `<wps:cNvSpPr${s.kind === "textbox" ? ' txBox="1"' : ""}/>`
    + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(s.width)}" cy="${emu(s.height)}"/></a:xfrm><a:prstGeom prst="${s.geometry}"><a:avLst/></a:prstGeom>${isLine ? "" : fill}${line}</wps:spPr>`
    + `${text}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
}

/** The shape's XML with a new fill (undefined: no fill). */
export function withShapeFill(xml: string, fill: string | undefined): string {
  const value = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
  return xml.replace(/(<wps:spPr\b[^>]*>)([\s\S]*?)(<\/wps:spPr>)/, (_, open: string, body: string, close: string) => {
    // The fill comes after the geometry and before the line (CT_ShapeProperties order).
    const cleaned = body.replace(/<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*?(\/>|>[\s\S]*?<\/a:\1>)/, "");
    const at = cleaned.search(/<a:ln\b|<a:effectLst|<a:effectDag|<a:scene3d|<a:sp3d|<a:extLst/);
    return open + (at < 0 ? cleaned + value : cleaned.slice(0, at) + value + cleaned.slice(at)) + close;
  });
}

/** The shape's XML with a new line (undefined: no line). */
export function withShapeLine(xml: string, line: { color: string; width: number } | undefined): string {
  const value = line ? `<a:ln w="${emu(line.width)}"><a:solidFill><a:srgbClr val="${line.color}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  return xml.replace(/(<wps:spPr\b[^>]*>)([\s\S]*?)(<\/wps:spPr>)/, (_, open: string, body: string, close: string) => {
    const cleaned = body.replace(/<a:ln\b[^>]*?(\/>|>[\s\S]*?<\/a:ln>)/, "");
    const at = cleaned.search(/<a:effectLst|<a:effectDag|<a:scene3d|<a:sp3d|<a:extLst/);
    return open + (at < 0 ? cleaned + value : cleaned.slice(0, at) + value + cleaned.slice(at)) + close;
  });
}
