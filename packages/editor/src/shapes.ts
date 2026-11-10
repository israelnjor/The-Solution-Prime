/**
 * Insert → Shapes and Text Box, and Shape Format: new shapes and text boxes as Word writes them (a
 * floating drawing kept as its XML, with its description for layout and display), and changes to the
 * fill and outline of the selected one — made in its XML (what is saved) and its description alike.
 */
import { newShapeXml, withShapeFill, withShapeLine } from "@prime/docx";
import type { FloatInfo, ShapeStyle } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection } from "prosemirror-state";

export interface ShapeChoice { geometry: string; label: string }

/** Word's everyday shapes (Insert → Shapes). */
export const SHAPES: ShapeChoice[] = [
  { geometry: "rect", label: "Rectangle" },
  { geometry: "roundRect", label: "Rounded rectangle" },
  { geometry: "ellipse", label: "Oval" },
  { geometry: "triangle", label: "Triangle" },
  { geometry: "diamond", label: "Diamond" },
  { geometry: "rightArrow", label: "Right arrow" },
  { geometry: "leftArrow", label: "Left arrow" },
  { geometry: "upArrow", label: "Up arrow" },
  { geometry: "downArrow", label: "Down arrow" },
  { geometry: "pentagon", label: "Arrow pentagon" },
  { geometry: "hexagon", label: "Hexagon" },
  { geometry: "star5", label: "Star" },
  { geometry: "line", label: "Line" },
];

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
 * Inserts a text box (white, a thin black border, for typing in) or a shape (Word's blue fill and darker
 * outline; a line just the line) at the cursor's paragraph, in front of the text, and selects it.
 */
export function insertShape(kind: "textbox" | "shape", geometry = "rect"): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.run_preserved;
    if (!type) return false;
    if (dispatch) {
      const id = nextDrawingId(state.doc);
      const isLine = geometry === "line";
      const width = kind === "textbox" ? 180 : isLine ? 144 : 108;
      const height = kind === "textbox" ? 72 : isLine ? 0.01 : 72;
      const style: ShapeStyle = kind === "textbox"
        ? { geometry: "rect", fill: "FFFFFF", line: { color: "000000", width: 0.75 } }
        : isLine ? { geometry, line: { color: "4472C4", width: 1.5 } }
        : { geometry, fill: "4472C4", line: { color: "2F528F", width: 1 } };
      const x = 0;
      const y = 6;
      const xml = newShapeXml({
        kind, geometry: style.geometry, width, height, x, y, id,
        name: kind === "textbox" ? `Text Box ${id}` : `${SHAPES.find((s) => s.geometry === geometry)?.label ?? "Shape"} ${id}`,
        ...(style.fill ? { fill: style.fill } : {}), ...(style.line ? { line: style.line } : {}),
      });
      const float: FloatInfo = {
        kind, width, height,
        x: { from: "column", offset: x }, y: { from: "paragraph", offset: y },
        wrap: "none", dist: { top: 3.6, bottom: 3.6, left: 9, right: 9 },
        shape: style,
        ...(kind === "textbox" ? {
          blocks: [{ type: "paragraph" as const, props: {}, content: [] }],
          insets: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
        } : {}),
      };
      const node = type.create({ props: {}, preserved: { name: "w:drawing", xml }, float });
      const at = state.selection.from;
      const tr = state.tr.insert(at, node);
      tr.setSelection(NodeSelection.create(tr.doc, at));
      dispatch(tr.scrollIntoView());
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

function updateShape(change: (xml: string, style: ShapeStyle) => { xml: string; style: ShapeStyle }): Command {
  return (state, dispatch) => {
    const s = shapeAt(state);
    if (!s) return false;
    if (dispatch) {
      const { xml, style } = change(s.node.attrs.preserved.xml, s.float.shape);
      const tr = state.tr.setNodeMarkup(s.pos, undefined, {
        ...s.node.attrs, preserved: { ...s.node.attrs.preserved, xml }, float: { ...s.float, shape: style },
      });
      dispatch(tr.setSelection(NodeSelection.create(tr.doc, s.pos)));
    }
    return true;
  };
}

/** Shape Fill: a colour (RRGGBB), or none. */
export function setShapeFill(fill: string | undefined): Command {
  return updateShape((xml, style) => ({ xml: withShapeFill(xml, fill), style: { ...style, fill } }));
}

/** Shape Outline: a colour and width (pt), or none. */
export function setShapeLine(line: { color: string; width: number } | undefined): Command {
  return updateShape((xml, style) => ({ xml: withShapeLine(xml, line), style: { ...style, line } }));
}
