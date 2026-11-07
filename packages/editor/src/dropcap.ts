/**
 * Drop caps (Insert → Drop Cap), made as Word makes them: the paragraph's first letter moves into a
 * paragraph of its own, framed (w:framePr w:dropCap) to sit beside the paragraph's first lines; its
 * line height is exactly those lines', its size makes the capital reach from the first line's cap
 * height to the last line's baseline, and it is lowered as Word lowers it (w:position).
 */
import type { Measurer } from "@prime/layout";
import { Resolver } from "@prime/layout";
import type { ParagraphProps, PrimeDocument, RunProps } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, TextSelection } from "prosemirror-state";

export type DropCapKind = "none" | "drop" | "margin";

const kindOf = (node: Node | null | undefined): "drop" | "margin" | undefined => {
  const xml = (node?.attrs.props as ParagraphProps | undefined)?.extra?.find((x) => x.name === "w:framePr")?.xml;
  return (xml && (/w:dropCap="(drop|margin)"/.exec(xml)?.[1] as "drop" | "margin" | undefined)) || undefined;
};

/** The paragraph at the cursor and its drop cap paragraph (before it), as top-level positions. */
function find(state: EditorState): { target: number; cap?: number } | undefined {
  const $from = state.selection.$from;
  if ($from.depth < 1) return undefined;
  const index = $from.index(0);
  const doc = state.doc;
  const posOf = (i: number) => { let p = 0; for (let k = 0; k < i; k++) p += doc.child(k).nodeSize; return p; };
  const here = doc.child(index);
  // The cursor in the drop cap itself: the paragraph is the next one.
  if (kindOf(here) && index + 1 < doc.childCount) return { target: posOf(index + 1), cap: posOf(index) };
  if (here.type.name !== "paragraph") return undefined;
  const before = index > 0 ? doc.child(index - 1) : undefined;
  return { target: posOf(index), ...(kindOf(before) ? { cap: posOf(index - 1) } : {}) };
}

/** The drop cap of the paragraph at the cursor. */
export function dropCapOf(state: EditorState): DropCapKind {
  const f = find(state);
  return f?.cap !== undefined ? kindOf(state.doc.nodeAt(f.cap)) ?? "none" : "none";
}

/** Sets the paragraph's drop cap: dropped or in the margin, `lines` lines tall; "none" puts the letter back. */
export function setDropCap(base: () => PrimeDocument, measurer: Measurer, kind: DropCapKind, lines = 3): Command {
  return (state, dispatch) => {
    const f = find(state);
    if (!f) return false;
    const doc = state.doc;
    const target = doc.nodeAt(f.target)!;
    const tr = state.tr;
    if (f.cap !== undefined) {
      const cap = doc.nodeAt(f.cap)!;
      if (kind !== "none") {
        // Another kind: only the frame changes.
        const props = cap.attrs.props as ParagraphProps;
        const extra = (props.extra ?? []).map((x) => (x.name === "w:framePr" ? { ...x, xml: x.xml.replace(/w:dropCap="\w+"/, `w:dropCap="${kind}"`) } : x));
        if (dispatch) dispatch(tr.setNodeMarkup(f.cap, undefined, { ...cap.attrs, props: { ...props, extra } }));
        return true;
      }
      // None: the letter back at the start of its paragraph, in the paragraph's own formatting.
      if (dispatch) {
        const letter = cap.textContent;
        const marks = target.firstChild?.marks ?? [];
        tr.insert(f.target + 1, state.schema.text(letter, marks));
        tr.delete(f.cap, f.cap + cap.nodeSize);
        tr.setSelection(TextSelection.create(tr.doc, f.cap + 1));
        dispatch(tr);
      }
      return true;
    }
    if (kind === "none") return true;
    const first = target.firstChild;
    if (!first?.isText || !first.text) return false;
    if (!dispatch) return true;
    // The body's line height and size, as the layout measures them.
    const sheet = { ...base(), styles: (doc.attrs.styles as PrimeDocument["styles"] | null) ?? base().styles };
    const resolver = new Resolver(sheet);
    const rp = resolver.paragraph({ type: "paragraph", props: target.attrs.props, content: [] });
    const runProps = first.marks.find((m) => m.type.name === "run")?.attrs.props as RunProps | undefined;
    const run = resolver.run(rp.styleId, runProps);
    const metrics = measurer.metrics({ family: run.font, bold: run.bold, italic: run.italic });
    const single = metrics.lineHeight * run.size;
    const line = rp.lineRule === "exact" ? rp.line / 20 : rp.lineRule === "atLeast" ? Math.max(rp.line / 20, single) : single * (rp.line / 240);
    // Capital height ~0.64 of the font size (Word: 57.5pt for three 14.82pt lines of 11pt text).
    const size = Math.round((((lines - 1) * line + 0.64 * run.size) / 0.64) * 2) / 2;
    const position = -Math.round(size * 0.1565); // half-points, as Word lowers it
    const positionXml = { name: "w:position", xml: `<w:position w:val="${position}"/>` };
    const letterProps: RunProps = { ...(runProps ?? {}), size, extra: [...(runProps?.extra ?? []).filter((x) => x.name !== "w:position"), positionXml] };
    const props: ParagraphProps = {
      keepNext: true,
      spacing: { after: 0, line: Math.round(lines * line * 20), lineRule: "exact" },
      markRun: { size, extra: [positionXml] },
      extra: [
        { name: "w:framePr", xml: `<w:framePr w:dropCap="${kind}" w:lines="${lines}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>` },
        { name: "w:textAlignment", xml: '<w:textAlignment w:val="baseline"/>' },
      ],
      ...(((target.attrs.props as ParagraphProps | null)?.style) ? { style: (target.attrs.props as ParagraphProps).style } : {}),
    };
    const letter = [...first.text][0];
    const runMark = state.schema.marks.run.create({ props: letterProps });
    const capNode = state.schema.nodes.paragraph.create({ ...target.attrs, props }, state.schema.text(letter, [...first.marks.filter((m) => m.type.name !== "run"), runMark]));
    tr.delete(f.target + 1, f.target + 1 + letter.length);
    tr.insert(f.target, capNode);
    tr.setSelection(TextSelection.create(tr.doc, f.target + capNode.nodeSize + 1));
    dispatch(tr);
    return true;
  };
}
