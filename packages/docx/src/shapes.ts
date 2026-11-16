/**
 * Shapes and text boxes (DrawingML wps:wsp in a floating wp:anchor): how they are drawn (preset
 * geometry, fill, line — from wps:spPr, or from the theme style Word gives new shapes), new ones as
 * Word writes them, and changes to their fill and line, made in their XML (which is what is saved).
 */
import type { LineEnd, ShapeLine, ShapeStyle } from "@prime/pdm";
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
  const prstGeom = spPr && child(spPr, "a:prstGeom");
  const geometry = prstGeom?.getAttribute("prst") || "rect";
  const out: ShapeStyle = { geometry };
  // Its adjustments (a:avLst a:gd name="adj" fmla="val 25000"), where set.
  const avLst = prstGeom && child(prstGeom, "a:avLst");
  for (const gd of avLst ? elements(avLst) : []) {
    const m = /^val\s+(-?\d+)$/.exec(gd.getAttribute("fmla") ?? "");
    const name = gd.getAttribute("name");
    if (m && name) (out.adj ??= {})[name] = Number(m[1]);
  }
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
  // Line ends (arrow heads): none, or one of the kinds drawn.
  const end = (name: string): LineEnd | undefined => {
    const type = ln && child(ln, name)?.getAttribute("type");
    return type && LINE_ENDS.has(type as LineEnd) ? (type as LineEnd) : undefined;
  };
  const [head, tail] = [end("a:headEnd"), end("a:tailEnd")];
  const extra = { ...(dash && dash !== "solid" ? { dash } : {}), ...(head ? { head } : {}), ...(tail ? { tail } : {}) };
  if (ln && child(ln, "a:noFill")) { /* none */ }
  else if (ln && child(ln, "a:solidFill")) {
    const color = colorOf(child(ln, "a:solidFill"), theme);
    if (color) out.line = { color, width, ...extra };
  } else if (style) {
    const ref = child(style, "a:lnRef");
    const color = ref && ref.getAttribute("idx") !== "0" ? colorOf(ref, theme) : undefined;
    if (color) out.line = { color, width: ln?.getAttribute("w") ? width : 1, ...extra };
  }
  return out;
}

const LINE_ENDS = new Set<LineEnd>(["triangle", "arrow", "stealth", "diamond", "oval"]);

/** A shape's turn (degrees clockwise) and mirroring, from its spPr's a:xfrm. */
export function readShapeTransform(wsp: Element): { rotation?: number; flipH?: boolean; flipV?: boolean } {
  const xfrm = child(child(wsp, "wps:spPr") ?? wsp, "a:xfrm");
  if (!xfrm) return {};
  const rot = Number(xfrm.getAttribute("rot") ?? 0) / 60000;
  return {
    ...(rot ? { rotation: ((rot % 360) + 360) % 360 } : {}),
    ...(xfrm.getAttribute("flipH") === "1" ? { flipH: true } : {}),
    ...(xfrm.getAttribute("flipV") === "1" ? { flipV: true } : {}),
  };
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
  line?: ShapeLine;
  /** Text box: its paragraphs' XML (w:p…). */
  textXml?: string;
  /** Its geometry's adjustments (a:avLst), where not Word's defaults. */
  adj?: Record<string, number>;
  /** Turned (degrees clockwise) and mirrored. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Unique among the document's drawings (wp:docPr id). */
  id: number;
  name: string;
}

/** A shape's line as DrawingML (a:ln): colour, width, dash, line ends; undefined: no line. */
export function lineXml(line: ShapeLine | undefined): string {
  if (!line) return "<a:ln><a:noFill/></a:ln>";
  return `<a:ln w="${emu(line.width)}"><a:solidFill><a:srgbClr val="${line.color}"/></a:solidFill>`
    + (line.dash ? `<a:prstDash val="${line.dash}"/>` : "")
    + (line.head ? `<a:headEnd type="${line.head}"/>` : "")
    + (line.tail ? `<a:tailEnd type="${line.tail}"/>` : "")
    + "</a:ln>";
}

/** The a:xfrm attributes for a turn and mirroring. */
const transformAttrs = (t: { rotation?: number; flipH?: boolean; flipV?: boolean }) =>
  (t.rotation ? ` rot="${Math.round((((t.rotation % 360) + 360) % 360) * 60000)}"` : "") + (t.flipH ? ' flipH="1"' : "") + (t.flipV ? ' flipV="1"' : "");

/** Text in a shape (a text box's paragraphs). */
const textBodyXml = (textXml: string) => `<wps:txbx><w:txbxContent>${textXml}</w:txbxContent></wps:txbx>`;

/** A floating drawing's wrapper (wp:anchor) as Word writes it, around its graphic (`uri`, `body`). */
export function anchorXml(a: {
  id: number; name: string; x: number; y: number; width: number; height: number; z?: number;
  uri: string; body: string;
  /** Its wrapping element (wp:wrapNone, wp:wrapSquare…) and whether it is behind the text. */
  wrap?: string; behind?: boolean;
}): string {
  return `<w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="${Math.round(a.z ?? 251659264 + a.id)}" behindDoc="${a.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(a.x)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(a.y)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${emu(a.width)}" cy="${emu(a.height)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${a.wrap ?? "<wp:wrapNone/>"}`
    + `<wp:docPr id="${a.id}" name="${esc(a.name)}"/><wp:cNvGraphicFramePr/>`
    + `<a:graphic><a:graphicData uri="${a.uri}">${a.body}</a:graphicData></a:graphic></wp:anchor></w:drawing>`;
}

const URI = {
  shape: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
};

/** A floating shape or text box as Word writes it (in front of text, square wrapping off: no wrap). */
export function newShapeXml(s: NewShape): string {
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : "<a:noFill/>";
  const isLine = s.geometry === "line" || s.geometry === "straightConnector1";
  const text = s.kind === "textbox"
    ? `${textBodyXml(s.textXml ?? "<w:p/>")}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`
    : `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
  const body = `<wps:wsp>`
    + (isLine ? "<wps:cNvCnPr/>" : `<wps:cNvSpPr${s.kind === "textbox" ? ' txBox="1"' : ""}/>`)
    + `<wps:spPr><a:xfrm${transformAttrs(s)}><a:off x="0" y="0"/><a:ext cx="${emu(s.width)}" cy="${emu(s.height)}"/></a:xfrm><a:prstGeom prst="${s.geometry}">${avLstXml(s.adj, s.geometry)}</a:prstGeom>${isLine ? "" : fill}${lineXml(s.line)}</wps:spPr>`
    + `${text}</wps:wsp>`;
  return anchorXml({ id: s.id, name: s.name, x: s.x, y: s.y, width: s.width, height: s.height, uri: URI.shape, body });
}

/**
 * Word's adjustments for each geometry, in its order, with their defaults (presetShapeDefinitions). Word
 * will not open a shape whose a:avLst gives only some of them: all are written when any is changed.
 */
const ADJ_DEFAULTS: Record<string, [string, number][]> = (() => {
  const arrow: [string, number][] = [["adj1", 50000], ["adj2", 50000]];
  const callout: [string, number][] = [["adj1", -20833], ["adj2", 62500]];
  return {
    roundRect: [["adj", 16667]], snip1Rect: [["adj", 16667]], round1Rect: [["adj", 16667]],
    snip2SameRect: [["adj1", 16667], ["adj2", 0]], round2SameRect: [["adj1", 16667], ["adj2", 0]], snipRoundRect: [["adj1", 16667], ["adj2", 16667]],
    triangle: [["adj", 50000]], parallelogram: [["adj", 25000]], trapezoid: [["adj", 25000]], hexagon: [["adj", 25000], ["vf", 115470]],
    octagon: [["adj", 29289]], plus: [["adj", 25000]], frame: [["adj1", 12500]], bevel: [["adj", 12500]], donut: [["adj", 25000]],
    noSmoking: [["adj", 18750]], can: [["adj", 25000]], cube: [["adj", 25000]], foldedCorner: [["adj", 16667]],
    homePlate: [["adj", 50000]], chevron: [["adj", 50000]], moon: [["adj", 50000]], smileyFace: [["adj", 4653]],
    wave: [["adj1", 12500], ["adj2", 0]], doubleWave: [["adj1", 6250], ["adj2", 0]],
    leftBracket: [["adj", 8333]], rightBracket: [["adj", 8333]], bracketPair: [["adj", 16667]],
    leftBrace: [["adj1", 8333], ["adj2", 50000]], rightBrace: [["adj1", 8333], ["adj2", 50000]],
    rightArrow: arrow, leftArrow: arrow, upArrow: arrow, downArrow: arrow, leftRightArrow: arrow, upDownArrow: arrow,
    notchedRightArrow: arrow, stripedRightArrow: arrow,
    star4: [["adj", 12500]], star5: [["adj", 19098], ["hf", 105146], ["vf", 110557]], star6: [["adj", 28868], ["hf", 115470]],
    star7: [["adj", 34601], ["hf", 102572], ["vf", 105210]], star8: [["adj", 37500]], star10: [["adj", 42533], ["hf", 105146]],
    star12: [["adj", 37500]], star16: [["adj", 37500]], star24: [["adj", 37500]], star32: [["adj", 37500]],
    wedgeRectCallout: callout, wedgeEllipseCallout: callout, cloudCallout: callout,
    wedgeRoundRectCallout: [...callout, ["adj3", 16667]],
  };
})();

/** A geometry's adjustments as DrawingML (a:avLst): all of the geometry's, in Word's order, when any is set. */
export function avLstXml(adj: Record<string, number> | undefined, geometry = ""): string {
  if (!adj || !Object.keys(adj).length) return "<a:avLst/>";
  const known = ADJ_DEFAULTS[geometry] ?? [];
  const names = [...known.map(([n]) => n), ...Object.keys(adj).filter((n) => !known.some(([k]) => k === n))];
  const value = (n: string) => adj[n] ?? known.find(([k]) => k === n)?.[1] ?? 0;
  return `<a:avLst>${names.map((n) => `<a:gd name="${esc(n)}" fmla="val ${Math.round(value(n))}"/>`).join("")}</a:avLst>`;
}

/** The shape's XML with new adjustments (its prstGeom's a:avLst). */
export function withShapeAdjust(xml: string, adj: Record<string, number> | undefined): string {
  return inSpPr(xml, (body) => body.replace(/<a:prstGeom\b([^>]*?)\s*\/>/, "<a:prstGeom$1></a:prstGeom>").replace(/(<a:prstGeom\b[^>]*>)([\s\S]*?)(<\/a:prstGeom>)/, (_, open: string, inner: string, close: string) =>
    open + inner.replace(/<a:avLst\s*\/>|<a:avLst\b[^>]*>[\s\S]*?<\/a:avLst>/, "").replace(/^/, avLstXml(adj, /prst="([^"]+)"/.exec(open)?.[1])) + close));
}

/** Edits the shape's spPr body. */
const inSpPr = (xml: string, edit: (body: string) => string) =>
  xml.replace(/(<wps:spPr\b[^>]*>)([\s\S]*?)(<\/wps:spPr>)/, (_, open: string, body: string, close: string) => open + edit(body) + close);

/** The shape's XML with a new fill (undefined: no fill). */
export function withShapeFill(xml: string, fill: string | undefined): string {
  const value = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
  return inSpPr(xml, (body) => {
    // The fill comes after the geometry and before the line (CT_ShapeProperties order).
    const cleaned = body.replace(/<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*?(\/>|>[\s\S]*?<\/a:\1>)/, "");
    const at = cleaned.search(/<a:ln\b|<a:effectLst|<a:effectDag|<a:scene3d|<a:sp3d|<a:extLst/);
    return at < 0 ? cleaned + value : cleaned.slice(0, at) + value + cleaned.slice(at);
  });
}

/** The shape's XML with a new line (colour, width, dash, line ends; undefined: no line). */
export function withShapeLine(xml: string, line: ShapeLine | undefined): string {
  const value = lineXml(line);
  return inSpPr(xml, (body) => {
    const cleaned = body.replace(/<a:ln\b[^>]*?(\/>|>[\s\S]*?<\/a:ln>)/, "");
    const at = cleaned.search(/<a:effectLst|<a:effectDag|<a:scene3d|<a:sp3d|<a:extLst/);
    return at < 0 ? cleaned + value : cleaned.slice(0, at) + value + cleaned.slice(at);
  });
}

/** The shape's XML turned and mirrored (its a:xfrm; made from its size if it had none). */
export function withShapeTransform(xml: string, t: { rotation?: number; flipH?: boolean; flipV?: boolean }): string {
  return inSpPr(xml, (body) => {
    if (/<a:xfrm\b/.test(body)) return body.replace(/<a:xfrm\b[^>]*?(\/?)>/, (_, close: string) => `<a:xfrm${transformAttrs(t)}${close}>`);
    const ext = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
    return `<a:xfrm${transformAttrs(t)}><a:off x="0" y="0"/><a:ext cx="${ext?.[1] ?? 0}" cy="${ext?.[2] ?? 0}"/></a:xfrm>${body}`;
  });
}

/** The XML with the text's place between the shape's top and bottom (wps:bodyPr anchor). */
export function withTextAnchor(xml: string, anchor: "t" | "ctr" | "b"): string {
  return xml.replace(/<wps:bodyPr\b[^>]*?(\/?)>/, (tag: string, close: string) =>
    / anchor="/.test(tag) ? tag.replace(/ anchor="[^"]*"/, ` anchor="${anchor}"`) : `${tag.slice(0, tag.length - (close ? 2 : 1))} anchor="${anchor}"${close}>`);
}

/** The XML with a new stacking order (wp:anchor relativeHeight). */
export function withStacking(xml: string, z: number): string {
  return xml.replace(/(<wp:anchor\b[^>]*?)relativeHeight="\d+"/, `$1relativeHeight="${Math.max(0, Math.round(z))}"`);
}

/**
 * A shape given text (Word's Add Text): its text body before its bodyPr, which keeps the shape's
 * outline (the text sits inside it). The old VML copy, which has no text, is dropped.
 */
export function withShapeText(xml: string, textXml: string): string {
  if (/<wps:txbx\b/.test(xml)) return xml;
  return xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/, "").replace(/<wps:bodyPr\b/, `${textBodyXml(textXml)}<wps:bodyPr`);
}
