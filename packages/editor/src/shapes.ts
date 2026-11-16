/**
 * Insert → Shapes and Text Box, and Shape Format: new shapes and text boxes as Word writes them (a
 * floating drawing kept as its XML, with its description for layout and display), and changes to the
 * fill and outline of the selected one — made in its XML (what is saved) and its description alike.
 */
import { newShapeXml, withShapeAdjust, withShapeFill, withShapeLine, withShapeText, withShapeTransform, withStacking, withTextAnchor } from "@prime/docx";
import type { Block, FloatInfo, LineEnd, ShapeLine, ShapeStyle } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection } from "prosemirror-state";

/** A shape to insert: its geometry, name, gallery group, and for lines, their arrow heads. */
export interface ShapeChoice { geometry: string; label: string; group: string; id?: string; ends?: { head?: LineEnd; tail?: LineEnd } }

const g = (group: string, list: [string, string][]): ShapeChoice[] => list.map(([geometry, label]) => ({ geometry, label, group }));

/** Word's shapes (Insert → Shapes), in its groups. */
export const SHAPES: ShapeChoice[] = [
  { geometry: "line", label: "Line", group: "Lines" },
  { geometry: "straightConnector1", id: "arrow", label: "Line Arrow", group: "Lines", ends: { tail: "triangle" } },
  { geometry: "straightConnector1", id: "doubleArrow", label: "Line Arrow: Double", group: "Lines", ends: { head: "triangle", tail: "triangle" } },
  ...g("Lines", [["arc", "Arc"], ["leftBracket", "Left Bracket"], ["rightBracket", "Right Bracket"], ["leftBrace", "Left Brace"], ["rightBrace", "Right Brace"], ["bracketPair", "Double Bracket"], ["bracePair", "Double Brace"]]),
  ...g("Rectangles", [["rect", "Rectangle"], ["roundRect", "Rectangle: Rounded Corners"], ["snip1Rect", "Rectangle: Single Corner Snipped"], ["snip2SameRect", "Rectangle: Top Corners Snipped"], ["round1Rect", "Rectangle: Single Corner Rounded"], ["round2SameRect", "Rectangle: Top Corners Rounded"], ["snipRoundRect", "Rectangle: Top Corners One Rounded and One Snipped"]]),
  ...g("Basic Shapes", [["ellipse", "Oval"], ["triangle", "Isosceles Triangle"], ["rtTriangle", "Right Triangle"], ["parallelogram", "Parallelogram"], ["trapezoid", "Trapezoid"], ["diamond", "Diamond"], ["pentagon", "Regular Pentagon"], ["hexagon", "Hexagon"], ["heptagon", "Heptagon"], ["octagon", "Octagon"], ["decagon", "Decagon"], ["dodecagon", "Dodecagon"], ["pie", "Partial Circle"], ["chord", "Chord"], ["teardrop", "Teardrop"], ["frame", "Frame"], ["plus", "Cross"], ["foldedCorner", "Rectangle: Folded Corner"], ["can", "Cylinder"], ["cube", "Cube"], ["bevel", "Rectangle: Beveled"], ["donut", "Circle: Hollow"], ["noSmoking", "\"No\" Symbol"], ["smileyFace", "Smiley Face"], ["heart", "Heart"], ["lightningBolt", "Lightning Bolt"], ["sun", "Sun"], ["moon", "Moon"], ["cloud", "Cloud"]]),
  ...g("Block Arrows", [["rightArrow", "Arrow: Right"], ["leftArrow", "Arrow: Left"], ["upArrow", "Arrow: Up"], ["downArrow", "Arrow: Down"], ["leftRightArrow", "Arrow: Left-Right"], ["upDownArrow", "Arrow: Up-Down"], ["quadArrow", "Arrow: Quad"], ["stripedRightArrow", "Arrow: Striped Right"], ["notchedRightArrow", "Arrow: Notched Right"], ["homePlate", "Arrow: Pentagon"], ["chevron", "Arrow: Chevron"]]),
  ...g("Equation Shapes", [["mathPlus", "Plus Sign"], ["mathMinus", "Minus Sign"], ["mathMultiply", "Multiplication Sign"], ["mathDivide", "Division Sign"], ["mathEqual", "Equals"], ["mathNotEqual", "Not Equal"]]),
  ...g("Flowchart", [["flowChartProcess", "Flowchart: Process"], ["flowChartAlternateProcess", "Flowchart: Alternate Process"], ["flowChartDecision", "Flowchart: Decision"], ["flowChartInputOutput", "Flowchart: Data"], ["flowChartPredefinedProcess", "Flowchart: Predefined Process"], ["flowChartInternalStorage", "Flowchart: Internal Storage"], ["flowChartDocument", "Flowchart: Document"], ["flowChartMultidocument", "Flowchart: Multidocument"], ["flowChartTerminator", "Flowchart: Terminator"], ["flowChartPreparation", "Flowchart: Preparation"], ["flowChartManualInput", "Flowchart: Manual Input"], ["flowChartManualOperation", "Flowchart: Manual Operation"], ["flowChartConnector", "Flowchart: Connector"], ["flowChartOffpageConnector", "Flowchart: Off-page Connector"], ["flowChartPunchedCard", "Flowchart: Card"], ["flowChartPunchedTape", "Flowchart: Punched Tape"], ["flowChartSummingJunction", "Flowchart: Summing Junction"], ["flowChartOr", "Flowchart: Or"], ["flowChartCollate", "Flowchart: Collate"], ["flowChartSort", "Flowchart: Sort"], ["flowChartExtract", "Flowchart: Extract"], ["flowChartMerge", "Flowchart: Merge"], ["flowChartOnlineStorage", "Flowchart: Stored Data"], ["flowChartDelay", "Flowchart: Delay"], ["flowChartMagneticTape", "Flowchart: Sequential Access Storage"], ["flowChartMagneticDisk", "Flowchart: Magnetic Disk"], ["flowChartMagneticDrum", "Flowchart: Direct Access Storage"], ["flowChartDisplay", "Flowchart: Display"]]),
  ...g("Stars and Banners", [["irregularSeal1", "Explosion: 8 Points"], ["irregularSeal2", "Explosion: 14 Points"], ["star4", "Star: 4 Points"], ["star5", "Star: 5 Points"], ["star6", "Star: 6 Points"], ["star7", "Star: 7 Points"], ["star8", "Star: 8 Points"], ["star10", "Star: 10 Points"], ["star12", "Star: 12 Points"], ["star16", "Star: 16 Points"], ["star24", "Star: 24 Points"], ["star32", "Star: 32 Points"], ["wave", "Wave"], ["doubleWave", "Double Wave"]]),
  ...g("Callouts", [["wedgeRectCallout", "Speech Bubble: Rectangle"], ["wedgeRoundRectCallout", "Speech Bubble: Rectangle with Corners Rounded"], ["wedgeEllipseCallout", "Speech Bubble: Oval"], ["cloudCallout", "Thought Bubble: Cloud"]]),
];

/** The gallery's groups, in order. */
export const SHAPE_GROUPS = [...new Set(SHAPES.map((s) => s.group))];

/** A line: drawn from one end to the other (nothing to fill, no text). */
export const isLine = (geometry: string) => geometry === "line" || geometry === "straightConnector1";

/** Where a shape drawn on the page goes: its anchor paragraph, its box (pt), and which way a line runs. */
export interface ShapePlace {
  /** The paragraph it is anchored to (its position in the document). */
  paraPos: number;
  /** From the column's left and the paragraph's top (pt). */
  x: number;
  y: number;
  width: number;
  height: number;
  flipH?: boolean;
  flipV?: boolean;
}

/** The next free drawing id in the document (wp:docPr ids must be unique). */
export function nextDrawingId(doc: Node): number {
  let max = 0;
  doc.descendants((n) => {
    const xml = n.attrs?.preserved?.xml as string | undefined;
    if (xml) for (const m of xml.matchAll(/<wp:docPr id="(\d+)"/g)) max = Math.max(max, Number(m[1]));
    return true;
  });
  return max + 1;
}

/**
 * The soft colours new shapes take (the highlighter's soft shades), each with a deeper outline of its
 * own colour: one after another, so the shapes of a diagram are told apart. Dark text reads on all.
 */
export const SOFT_LOOKS: { fill: string; line: string }[] = [
  { fill: "DDEBF7", line: "5B9BD5" }, // blue
  { fill: "E2F0D9", line: "70AD47" }, // green
  { fill: "FCE4D6", line: "ED7D31" }, // peach
  { fill: "FFF2CC", line: "BF9000" }, // yellow
  { fill: "EDE2F6", line: "8E6BBF" }, // lavender
  { fill: "FFD7E6", line: "D6608C" }, // pink
  { fill: "D9F2F2", line: "3A9C9C" }, // teal
];

/**
 * The look of a new shape: a soft fill and a deeper outline (the n-th of the soft colours); a line in
 * the deeper colour; a text box white with a thin black border.
 */
function newStyle(kind: "textbox" | "shape", choice: ShapeChoice, n: number): ShapeStyle {
  if (kind === "textbox") return { geometry: "rect", fill: "FFFFFF", line: { color: "000000", width: 0.75 } };
  const look = SOFT_LOOKS[n % SOFT_LOOKS.length];
  if (isLine(choice.geometry)) return { geometry: choice.geometry, line: { color: look.line, width: 1.25, ...choice.ends } };
  if (/^(arc|leftBracket|rightBracket|leftBrace|rightBrace|bracketPair|bracePair)$/.test(choice.geometry)) return { geometry: choice.geometry, line: { color: look.line, width: 1.25 } };
  return { geometry: choice.geometry, fill: look.fill, line: { color: look.line, width: 1 } };
}

/** How many shapes (not text boxes) the document has: the next new one takes the next soft colour. */
function shapeCount(doc: Node): number {
  let n = 0;
  doc.descendants((node) => {
    const f = node.attrs?.float as FloatInfo | null | undefined;
    if (f?.kind === "shape" || (f?.kind === "textbox" && f.shape && f.shape.geometry !== "rect")) n++;
    return true;
  });
  return n;
}

/**
 * Inserts a text box (white, a thin black border, for typing in) or a shape (a soft fill and a deeper
 * outline; a line just the line), in front of the text, and selects it: where it was drawn on the page
 * (`place`), or at the cursor's paragraph at Word's default size.
 */
export function insertShape(kind: "textbox" | "shape", choice: ShapeChoice | string = "rect", place?: ShapePlace): Command {
  const pick: ShapeChoice = typeof choice === "string" ? SHAPES.find((s) => s.geometry === choice) ?? { geometry: choice, label: "Shape", group: "" } : choice;
  return (state, dispatch) => {
    const type = state.schema.nodes.run_preserved;
    if (!type) return false;
    if (dispatch) {
      const id = nextDrawingId(state.doc);
      const line = isLine(pick.geometry);
      const width = place?.width ?? (kind === "textbox" ? 180 : line ? 144 : 72);
      const height = place ? Math.max(place.height, line ? 0.01 : 1) : kind === "textbox" ? 72 : line ? 0.01 : 72;
      const style = newStyle(kind, pick, shapeCount(state.doc));
      const x = place?.x ?? 0;
      const y = place?.y ?? 6;
      const turn = { ...(place?.flipH ? { flipH: true } : {}), ...(place?.flipV ? { flipV: true } : {}) };
      const xml = newShapeXml({
        kind, geometry: style.geometry, width, height, x, y, id, ...turn,
        name: kind === "textbox" ? `Text Box ${id}` : `${pick.label} ${id}`,
        ...(style.fill ? { fill: style.fill } : {}), ...(style.line ? { line: style.line } : {}),
      });
      const float: FloatInfo = {
        kind, width, height,
        x: { from: "column", offset: x }, y: { from: "paragraph", offset: y },
        wrap: "none", dist: { top: 3.6, bottom: 3.6, left: 9, right: 9 },
        shape: style, z: 251659264 + id, ...turn,
        ...(kind === "textbox" ? {
          blocks: [{ type: "paragraph" as const, props: {}, content: [] }],
          insets: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
        } : {}),
      };
      const node = type.create({ props: {}, preserved: { name: "w:drawing", xml }, float });
      // At the start of the paragraph it was drawn beside, or at the cursor.
      const at = place ? place.paraPos + 1 : state.selection.from;
      const tr = state.tr.insert(at, node);
      tr.setSelection(NodeSelection.create(tr.doc, at));
      dispatch(place ? tr : tr.scrollIntoView());
    }
    return true;
  };
}

/** The selected shape or text box (a floating drawing with a shape description), if any. */
export function shapeAt(state: EditorState): { pos: number; node: Node; float: FloatInfo & { shape: ShapeStyle } } | undefined {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return undefined;
  const f = sel.node.attrs.float as FloatInfo | null;
  if (sel.node.type.name !== "run_preserved" || !f?.shape) return undefined;
  return { pos: sel.from, node: sel.node, float: f as FloatInfo & { shape: ShapeStyle } };
}

/** Changes the selected shape: its XML (what is saved) and its description (what is drawn) together. */
function updateShape(change: (xml: string, float: FloatInfo & { shape: ShapeStyle }) => { xml: string; float: FloatInfo }): Command {
  return (state, dispatch) => {
    const s = shapeAt(state);
    if (!s) return false;
    if (dispatch) {
      const { xml, float } = change(s.node.attrs.preserved.xml, s.float);
      const tr = state.tr.setNodeMarkup(s.pos, undefined, { ...s.node.attrs, preserved: { ...s.node.attrs.preserved, xml }, float });
      dispatch(tr.setSelection(NodeSelection.create(tr.doc, s.pos)));
    }
    return true;
  };
}
const updateStyle = (change: (xml: string, style: ShapeStyle) => { xml: string; style: ShapeStyle }) =>
  updateShape((xml, f) => { const c = change(xml, f.shape); return { xml: c.xml, float: { ...f, shape: c.style } }; });

/** Shape Fill: a colour (RRGGBB), or none. */
export function setShapeFill(fill: string | undefined): Command {
  return updateStyle((xml, style) => ({ xml: withShapeFill(xml, fill), style: { ...style, fill } }));
}

/**
 * Shape Outline: a colour, width (pt), dash or line ends, changed on the line as it is (the rest kept);
 * undefined: no outline.
 */
export function setShapeLine(line: Partial<ShapeLine> | undefined): Command {
  return updateStyle((xml, style) => {
    const next: ShapeLine | undefined = line ? { color: "000000", width: 0.75, ...style.line, ...line } : undefined;
    // Solid, and no arrow heads: none written.
    if (next && !next.dash) delete next.dash;
    if (next && !next.head) delete next.head;
    if (next && !next.tail) delete next.tail;
    return { xml: withShapeLine(xml, next), style: { ...style, line: next } };
  });
}

/** Shape Styles: a fill and an outline together (a preset look). */
export function setShapeLook(look: { fill?: string; line?: { color: string; width: number } }): Command {
  return updateStyle((xml, style) => {
    const line = look.line ? { ...style.line, ...look.line } : undefined;
    const fill = isLine(style.geometry) ? undefined : look.fill;
    return { xml: withShapeLine(withShapeFill(xml, fill), line), style: { ...style, fill, line } };
  });
}

/** The shape's adjustments (its yellow handles: corner roundness, an arrow's head, a callout's pointer). */
export function setShapeAdjust(adj: Record<string, number>): Command {
  return updateStyle((xml, style) => ({ xml: withShapeAdjust(xml, adj), style: { ...style, adj } }));
}

/** Rotate: turned to `rotation` degrees, and/or mirrored (Flip Horizontal / Vertical); the rest kept. */
export function setShapeTurn(turn: { rotation?: number; flipH?: boolean; flipV?: boolean }): Command {
  return updateShape((xml, f) => {
    const next = {
      rotation: ((Math.round((turn.rotation ?? f.rotation ?? 0) * 10) / 10) % 360 + 360) % 360,
      flipH: turn.flipH ?? f.flipH ?? false,
      flipV: turn.flipV ?? f.flipV ?? false,
    };
    const { rotation: _r, flipH: _h, flipV: _v, ...rest } = f;
    return {
      xml: withShapeTransform(xml, next),
      float: { ...rest, ...(next.rotation ? { rotation: next.rotation } : {}), ...(next.flipH ? { flipH: true } : {}), ...(next.flipV ? { flipV: true } : {}) },
    };
  });
}

/** Rotate Right 90° / Left 90°. */
export const rotateShape = (by: number): Command => (state, dispatch) => {
  const s = shapeAt(state);
  return s ? setShapeTurn({ rotation: (s.float.rotation ?? 0) + by })(state, dispatch) : false;
};
/** Flip Horizontal / Vertical. */
export const flipShape = (way: "h" | "v"): Command => (state, dispatch) => {
  const s = shapeAt(state);
  return s ? setShapeTurn(way === "h" ? { flipH: !s.float.flipH } : { flipV: !s.float.flipV })(state, dispatch) : false;
};

/** Align Text: the text at the top, in the middle or at the bottom of its shape or text box. */
export function setTextAnchor(anchor: "t" | "ctr" | "b"): Command {
  return updateShape((xml, f) => {
    const { anchor: _a, ...rest } = f;
    return { xml: withTextAnchor(xml, anchor), float: { ...rest, ...(anchor === "t" ? {} : { anchor }) } };
  });
}

/** The stacking orders of the document's floating objects. */
function stackingOf(doc: Node): number[] {
  const out: number[] = [];
  doc.descendants((n) => {
    const f = n.attrs?.float as FloatInfo | null | undefined;
    if (f) {
      const xml = n.attrs.preserved?.xml as string | undefined;
      out.push(f.z ?? Number(/relativeHeight="(\d+)"/.exec(xml ?? "")?.[1] ?? 0));
    }
    return true;
  });
  return out;
}

/** Bring to Front / Send to Back: above or below every other floating object. */
export function stackShape(where: "front" | "back"): Command {
  return (state, dispatch) => {
    const s = shapeAt(state);
    if (!s) return false;
    const all = stackingOf(state.doc);
    const z = where === "front" ? Math.max(0, ...all) + 1 : Math.max(0, Math.min(...all) - 1);
    return updateShape((xml, f) => ({ xml: withStacking(xml, z), float: { ...f, z } }))(state, dispatch);
  };
}

/** Size: the shape's width and height (pt). */
export function setShapeSize(width: number, height: number): Command {
  const emu = (pt: number) => Math.round(pt * 12700);
  return updateShape((xml, f) => ({
    xml: xml
      .replace(/<wp:extent cx="\d+" cy="\d+"\s*\/>/, `<wp:extent cx="${emu(width)}" cy="${emu(height)}"/>`)
      .replace(/(<a:xfrm\b[^>]*>\s*<a:off\b[^>]*\/>\s*)<a:ext cx="\d+" cy="\d+"\s*\/>/, `$1<a:ext cx="${emu(width)}" cy="${emu(height)}"/>`),
    float: { ...f, width, height },
  }));
}

/** Whether text on the fill reads best in white (a dark fill); undefined: the usual black. */
function textOn(fill: string | undefined): string | undefined {
  if (!fill) return undefined;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(fill.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 150 ? "FFFFFF" : undefined;
}

/**
 * Add Text: the selected shape takes text, as in Word: it keeps its outline and fill, and its text is
 * centred across it and in its middle (white on a dark fill). Lines take none.
 */
export function addShapeText(): Command {
  return (state, dispatch) => {
    const s = shapeAt(state);
    if (!s || s.float.kind !== "shape" || isLine(s.float.shape.geometry)) return false;
    if (dispatch) {
      const color = textOn(s.float.shape.fill);
      const block: Block = { type: "paragraph", props: { align: "center", ...(color ? { markRun: { color } } : {}) }, content: [] };
      const textXml = `<w:p><w:pPr><w:jc w:val="center"/>${color ? `<w:rPr><w:color w:val="${color}"/></w:rPr>` : ""}</w:pPr></w:p>`;
      updateShape((xml, f) => ({
        xml: withTextAnchor(withShapeText(xml, textXml), "ctr"),
        float: { ...f, kind: "textbox", blocks: [block], insets: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }, anchor: "ctr", edited: true },
      }))(state, dispatch);
    }
    return true;
  };
}
