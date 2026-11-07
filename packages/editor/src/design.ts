/**
 * Design tab: page colour, page borders and watermark — written as Word writes them, so Word shows
 * (and edits) them as its own: the document's w:background, each section's w:pgBorders, and Word's
 * watermark shape (VML WordArt "PowerPlusWaterMarkObject") in every header.
 */
import type { Block, HeaderFooterKind, Inline, Paragraph, Story } from "@prime/pdm";
import { watermarkShape } from "@prime/layout";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import type { SectionAttrs } from "./docsetup.js";

const EMPTY: SectionAttrs = { props: {}, headers: {}, footers: {} };

/** Every section's attributes and where they live (-1: the document's own, for the last section). */
function sections(state: EditorState): { pos: number; attrs: SectionAttrs }[] {
  const out: { pos: number; attrs: SectionAttrs }[] = [];
  state.doc.forEach((node, offset) => {
    if (node.type.name === "section_break") out.push({ pos: offset, attrs: node.attrs.section as SectionAttrs });
  });
  out.push({ pos: -1, attrs: (state.doc.attrs.section as SectionAttrs | null) ?? EMPTY });
  return out;
}

function setSection(tr: Transaction, pos: number, attrs: SectionAttrs): Transaction {
  return pos === -1 ? tr.setDocAttribute("section", attrs) : tr.setNodeMarkup(pos, undefined, { section: attrs });
}

// --- page colour ---------------------------------------------------------------------------------

/** The page colour (hex RGB, no '#'), if any. */
export function pageColorOf(state: EditorState): string | undefined {
  const xml = (state.doc.attrs.background as { xml: string } | null)?.xml;
  return xml ? /w:color="([0-9A-Fa-f]{6})"/.exec(xml)?.[1]?.toUpperCase() : undefined;
}

/** Sets the page colour (hex RGB); undefined removes it. On screen and in PDF, as in Word. */
export function setPageColor(hex: string | undefined): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const value = hex ? { name: "w:background", xml: `<w:background w:color="${hex.replace("#", "").toUpperCase()}"/>` } : null;
      dispatch(state.tr.setDocAttribute("background", value));
    }
    return true;
  };
}

// --- page borders --------------------------------------------------------------------------------

export interface PageBorderSpec {
  /** Word's line style: single, double, dotted, dashed, thick… */
  style: string;
  /** Line width in pt (Word allows 0.25–6). */
  width: number;
  /** Hex RGB, no '#'; undefined: automatic (black). */
  color?: string;
  /** Measured from the page edge (Word's default for a box) or from the text. */
  offsetFrom?: "page" | "text";
  /** Distance in pt from the page edge or the text (Word: 0–31). */
  space?: number;
}

/** The page border of the section at the cursor (the first side found), if any. */
export function pageBordersOf(state: EditorState): PageBorderSpec | undefined {
  const all = sections(state);
  const index = all.findIndex((s) => s.pos === -1 || s.pos >= state.selection.from);
  const xml = all[index].attrs.props.extra?.find((x) => x.name === "w:pgBorders")?.xml;
  const side = xml && /<w:(?:top|left|bottom|right)\b([^>]*?)\/?>/.exec(xml)?.[1];
  if (!side) return undefined;
  const style = /w:val="([^"]+)"/.exec(side)?.[1] ?? "single";
  if (style === "none" || style === "nil") return undefined;
  const color = /w:color="([0-9A-Fa-f]{6})"/.exec(side)?.[1]?.toUpperCase();
  return {
    style, width: Number(/w:sz="(\d+)"/.exec(side)?.[1] ?? 4) / 8, space: Number(/w:space="(\d+)"/.exec(side)?.[1] ?? 24),
    offsetFrom: /w:offsetFrom="text"/.test(xml) ? "text" : "page", ...(color ? { color } : {}),
  };
}

/** A box border round every page of the whole document; undefined removes page borders. */
export function setPageBorders(spec: PageBorderSpec | undefined): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    let xml: string | undefined;
    if (spec) {
      const sz = Math.round(Math.min(6, Math.max(0.25, spec.width)) * 8);
      const fromPage = (spec.offsetFrom ?? "page") === "page";
      const space = Math.round(Math.min(31, Math.max(0, spec.space ?? (fromPage ? 24 : 4))));
      const side = (name: string) => `<w:${name} w:val="${spec.style}" w:sz="${sz}" w:space="${space}" w:color="${spec.color ?? "auto"}"/>`;
      xml = `<w:pgBorders${fromPage ? ' w:offsetFrom="page"' : ""}>${["top", "left", "bottom", "right"].map(side).join("")}</w:pgBorders>`;
    }
    let tr = state.tr;
    for (const s of sections(state)) {
      const extra = (s.attrs.props.extra ?? []).filter((x) => x.name !== "w:pgBorders");
      if (xml) extra.push({ name: "w:pgBorders", xml });
      tr = setSection(tr, s.pos, { ...s.attrs, props: { ...s.attrs.props, extra } });
    }
    dispatch(tr);
    return true;
  };
}

// --- watermark -----------------------------------------------------------------------------------

export interface WatermarkSpec {
  text: string;
  /** Hex RGB, no '#' (Word's default: silver, C0C0C0). */
  color?: string;
  font?: string;
  /** Diagonal (Word's default) or horizontal. */
  layout?: "diagonal" | "horizontal";
  /** Semi-transparent (Word's default). */
  semitransparent?: boolean;
}

/** Word's WordArt shape type (text along a path), which its watermark shapes refer to. */
const SHAPETYPE = '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Word's watermark: a WordArt shape centred on the page's margins, behind the text. */
export function watermarkXml(spec: WatermarkSpec, id: number): string {
  const diagonal = (spec.layout ?? "diagonal") === "diagonal";
  const font = spec.font ?? "Calibri";
  // The shape's box: Word sizes it to the text (about half an em per character, an em and a fifth high).
  const width = diagonal ? 468 : 415;
  const height = Math.round((width / Math.max(1, spec.text.length * 0.5)) * 10) / 10;
  const color = spec.color ? `#${spec.color.replace("#", "").toLowerCase()}` : "silver";
  const style = `position:absolute;margin-left:0;margin-top:0;width:${width}pt;height:${Math.min(height, width / 2)}pt;${diagonal ? "rotation:315;" : ""}z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin`;
  return `<w:pict>${SHAPETYPE}<v:shape id="PowerPlusWaterMarkObject${id}" o:spid="_x0000_s${2049 + id}" type="#_x0000_t136" style="${style}" o:allowincell="f" fillcolor="${color}" stroked="f">`
    + `${spec.semitransparent === false ? "" : '<v:fill opacity=".5"/>'}<v:textpath style="font-family:&quot;${esc(font)}&quot;;font-size:1pt" string="${esc(spec.text)}"/>`
    + `<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict>`;
}

const isWatermark = (i: Inline) => (i.type === "runPreserved" || i.type === "preserved") && /PowerPlusWaterMarkObject/.test(i.preserved.xml);

/** A header without any watermark (paragraphs that held only one are kept, empty). */
function withoutWatermark(story: Story): Story {
  const clean = (blocks: Block[]): Block[] => blocks.map((b) => (b.type === "paragraph" ? { ...b, content: b.content.filter((i) => !isWatermark(i)) } : b));
  return { ...story, blocks: clean(story.blocks) };
}

/** The watermark of the section at the cursor's headers (or the ones it inherits), if any. */
export function watermarkOf(state: EditorState): WatermarkSpec | undefined {
  for (const s of sections(state)) {
    for (const story of Object.values(s.attrs.headers ?? {})) {
      for (const b of story?.blocks ?? []) {
        if (b.type !== "paragraph") continue;
        for (const i of b.content) {
          if (!isWatermark(i)) continue;
          const shape = watermarkShape((i as { preserved: { xml: string } }).preserved.xml);
          if (shape) return { text: shape.text, color: shape.color, font: shape.family, layout: shape.rotation ? "diagonal" : "horizontal", semitransparent: shape.opacity < 1 };
        }
      }
    }
  }
  return undefined;
}

/**
 * Puts a watermark on every page (in every header of every section, as Word does); undefined removes
 * it. A document without headers gets an empty one to hold it.
 */
export function setWatermark(spec: WatermarkSpec | undefined): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    let tr = state.tr;
    let id = 1;
    const all = sections(state);
    all.forEach((s, index) => {
      const headers: Partial<Record<HeaderFooterKind, Story>> = {};
      for (const [kind, story] of Object.entries(s.attrs.headers ?? {}) as [HeaderFooterKind, Story][]) headers[kind] = withoutWatermark(story);
      // The first section always has a default header to hold it; later ones inherit it.
      if (spec && index === 0 && !headers.default) headers.default = { blocks: [{ type: "paragraph", props: { style: "Header" }, content: [] } as Paragraph] };
      if (spec) {
        for (const kind of Object.keys(headers) as HeaderFooterKind[]) {
          const story = headers[kind]!;
          const blocks = [...story.blocks];
          const at = blocks.findIndex((b) => b.type === "paragraph");
          const mark: Inline = { type: "runPreserved", preserved: { name: "w:pict", xml: watermarkXml(spec, id++) } };
          if (at >= 0) blocks[at] = { ...(blocks[at] as Paragraph), content: [mark, ...(blocks[at] as Paragraph).content] };
          else blocks.unshift({ type: "paragraph", props: {}, content: [mark] } as Paragraph);
          headers[kind] = { ...story, blocks };
        }
      }
      tr = setSection(tr, s.pos, { ...s.attrs, headers });
    });
    dispatch(tr);
    return true;
  };
}
