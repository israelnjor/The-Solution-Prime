/**
 * Prime Document Model (PDM) — ADR-004.
 *
 * The editor-independent representation of a document. DOCX, layout, PDF, storage and the
 * editor all speak this model. It is plain JSON: binary data is base64, so a document can be
 * snapshotted with JSON.stringify and restored with JSON.parse.
 *
 * Units: lengths are points (pt). Colors are uppercase hex without '#', e.g. "C00000".
 * Content the model does not understand is carried as `Preserved` raw OOXML so it can be
 * written back unchanged (ADR-006).
 */

export const SCHEMA_VERSION = 1;

/** A raw OOXML fragment kept verbatim. `name` is its qualified element name, e.g. "w:bookmarkStart". */
export interface Preserved {
  name: string;
  xml: string;
}

/** Attribute bag for elements we keep losslessly but do not interpret. Keys are local names. */
export type Attrs = Record<string, string>;

// --- character and paragraph properties ----------------------------------------------------

export interface FontSpec {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
  hint?: string;
}

export interface RunProps {
  style?: string;
  font?: FontSpec;
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  hidden?: boolean;
  color?: string;
  /** Font size in pt. */
  size?: number;
  highlight?: string;
  underline?: Attrs;
  shading?: Attrs;
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Extra space after each character, in pt (negative: condensed). w:spacing. */
  charSpacing?: number;
  extra?: Preserved[];
}

export interface Indent {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
  other?: Attrs;
}

export interface Spacing {
  before?: number;
  after?: number;
  /** Raw line value: 240ths of a line when lineRule is "auto", otherwise twips. */
  line?: number;
  lineRule?: string;
  other?: Attrs;
}

export interface ParagraphProps {
  style?: string;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  list?: { numId: string; level?: number };
  spacing?: Spacing;
  indent?: Indent;
  align?: string;
  outlineLevel?: number;
  /** Formatting of the paragraph mark itself. */
  markRun?: RunProps;
  extra?: Preserved[];
}

// --- inline content ------------------------------------------------------------------------

export interface TextInline {
  type: "text";
  text: string;
  props?: RunProps;
}

export interface TabInline {
  type: "tab";
  props?: RunProps;
}

export interface BreakInline {
  type: "break";
  kind: "line" | "page" | "column";
  clear?: string;
  props?: RunProps;
}

export interface ImageInline {
  type: "image";
  imageId: string;
  /** Display size in pt. */
  width: number;
  height: number;
  name?: string;
  description?: string;
  props?: RunProps;
  /** A floating picture (wrap, position): written as wp:anchor. Absent: in line with text. */
  float?: FloatInfo;
}

export interface LinkInline {
  type: "link";
  href?: string;
  anchor?: string;
  tooltip?: string;
  content: Inline[];
}

/** w:fldSimple — a field whose instruction and cached result travel together. */
export interface FieldInline {
  type: "field";
  instr: string;
  content: Inline[];
}

/** Pieces of a complex field (w:fldChar / w:instrText). Kept flat; consumers pair them up. */
export interface FieldCharInline {
  type: "fieldChar";
  kind: "begin" | "separate" | "end";
  props?: RunProps;
}

export interface InstrInline {
  type: "instr";
  text: string;
  props?: RunProps;
}

/** Unsupported run content (e.g. footnote references, floating drawings), kept inside a run. */
export interface RunPreservedInline {
  type: "runPreserved";
  preserved: Preserved;
  props?: RunProps;
  /** For a floating picture or text box: where it sits and what it shows (for layout and display). */
  float?: FloatInfo;
}

/** Where a floating object's position is measured from (wp:positionH / wp:positionV relativeFrom). */
export interface FloatPosition {
  from: string;
  /** Offset in pt, or an alignment ("left", "center", "right", "top", "bottom", "inside", "outside"). */
  offset?: number;
  align?: string;
}

/**
 * A floating object (wp:anchor), read for layout and display only: the object itself is written back
 * verbatim from its preserved XML, so nothing about it can be lost.
 */
export interface FloatInfo {
  kind: "picture" | "textbox" | "ink" | "shape";
  /** Size in pt. A text box that grows with its text (`autoHeight`) gets its height from its content. */
  width: number;
  height: number;
  x: FloatPosition;
  y: FloatPosition;
  wrap: "square" | "tight" | "through" | "topAndBottom" | "none";
  /** Which sides text wraps on (wp:wrapSquare wrapText). */
  side?: "bothSides" | "left" | "right" | "largest";
  /** Drawn behind the text (only matters without wrapping). */
  behind?: boolean;
  /** Distance between the object and wrapped text, in pt. */
  dist: { top: number; bottom: number; left: number; right: number };
  /** Picture: its image part in the package (e.g. "word/media/image3.png"), or its image resource. */
  imagePart?: string;
  imageId?: string;
  /** Text box: its content and inner margins (pt). */
  blocks?: Block[];
  insets?: { top: number; bottom: number; left: number; right: number };
  autoHeight?: boolean;
  /** Ink (pen strokes): x, y pairs relative to the box (0–1 each way), pen width in pt, colour. */
  ink?: { strokes: number[][]; width: number; color: string };
}

/** Unsupported paragraph-level content (e.g. bookmarks), kept verbatim. */
export type NoteKind = "footnote" | "endnote";

/** A footnote or endnote reference in the text (w:footnoteReference / w:endnoteReference); shown as its number. */
export interface NoteRefInline {
  type: "noteRef";
  kind: NoteKind;
  id: string;
  props?: RunProps;
}

/** The note's own number at the start of its text (w:footnoteRef / w:endnoteRef). */
export interface NoteMarkInline {
  type: "noteMark";
  kind: NoteKind;
  props?: RunProps;
}

/** A footnote's or endnote's text. */
export interface Note {
  id: string;
  blocks: Block[];
}

/** The footnotes (or endnotes) part: the notes, plus Word's separator notes kept verbatim. */
export interface Notes {
  items: Note[];
  /** w:footnote elements with a w:type (separator, continuation separator…), as they were. */
  special: Preserved[];
  part?: string;
  rootAttrs?: Record<string, string>;
  preservedRels?: Relationship[];
}

export interface PreservedInline {
  type: "preserved";
  preserved: Preserved;
}

/**
 * A content control (w:sdt) around inline content. The content is ordinary model content;
 * the control's own settings (w:sdtPr, w:sdtEndPr) are kept verbatim.
 */
export interface ContentControlInline {
  type: "contentControl";
  settings: Preserved[];
  content: Inline[];
}

/**
 * A tracked change around inline content. Deleted content stays in the model (Word shows it
 * struck through in markup view) until the change is accepted or rejected.
 */
export interface RevisionInline {
  type: "revision";
  kind: "ins" | "del";
  /** id, author, date and any other attributes, verbatim (qualified names kept). */
  attrs: Attrs;
  content: Inline[];
}

export type Inline =
  | NoteRefInline
  | NoteMarkInline
  | TextInline
  | TabInline
  | BreakInline
  | ImageInline
  | LinkInline
  | FieldInline
  | FieldCharInline
  | InstrInline
  | RunPreservedInline
  | PreservedInline
  | ContentControlInline
  | RevisionInline;

// --- blocks --------------------------------------------------------------------------------

export interface Paragraph {
  type: "paragraph";
  props: ParagraphProps;
  content: Inline[];
}

export interface TableProps {
  style?: string;
  width?: Attrs;
  align?: string;
  extra?: Preserved[];
}

export interface RowProps {
  header?: boolean;
  cantSplit?: boolean;
  height?: Attrs;
  extra?: Preserved[];
}

export interface CellProps {
  width?: Attrs;
  span?: number;
  vMerge?: "restart" | "continue";
  shading?: Attrs;
  vAlign?: string;
  extra?: Preserved[];
}

export interface TableCell {
  props: CellProps;
  blocks: Block[];
}

export interface TableRow {
  props: RowProps;
  cells: TableCell[];
  /** Row-level exceptions and other row content we keep verbatim (e.g. w:tblPrEx). */
  extra?: Preserved[];
}

export interface Table {
  type: "table";
  props: TableProps;
  /** Column widths in pt. */
  grid: number[];
  rows: TableRow[];
}

export interface PreservedBlock {
  type: "preserved";
  preserved: Preserved;
}

/** A content control around whole paragraphs or tables. */
export interface ContentControlBlock {
  type: "contentControl";
  settings: Preserved[];
  blocks: Block[];
}

export type Block = Paragraph | Table | ContentControlBlock | PreservedBlock;

// --- sections, headers, footers ------------------------------------------------------------

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

/** A header or footer. Relationships are kept for preserved content that references them. */
export interface Story {
  blocks: Block[];
  preservedRels?: Relationship[];
}

export type HeaderFooterKind = "default" | "first" | "even";

export interface SectionProps {
  pageWidth?: number;
  pageHeight?: number;
  orientation?: "portrait" | "landscape";
  /** Printer paper code (w:pgSz/@w:code). */
  paperCode?: string;
  margins?: { top?: number; right?: number; bottom?: number; left?: number; header?: number; footer?: number; gutter?: number };
  breakType?: string;
  titlePage?: boolean;
  extra?: Preserved[];
}

export interface Section {
  props: SectionProps;
  /** A missing kind means "inherit from the previous section", as in Word. */
  headers: Partial<Record<HeaderFooterKind, Story>>;
  footers: Partial<Record<HeaderFooterKind, Story>>;
  blocks: Block[];
}

// --- styles and lists ----------------------------------------------------------------------

export interface Style {
  id: string;
  type: "paragraph" | "character" | "table" | "numbering";
  name?: string;
  basedOn?: string;
  next?: string;
  link?: string;
  isDefault?: boolean;
  custom?: boolean;
  paragraph?: ParagraphProps;
  run?: RunProps;
  table?: TableProps;
  extra?: Preserved[];
}

export interface StyleSheet {
  defaults: { run?: RunProps; paragraph?: ParagraphProps };
  styles: Style[];
  extra?: Preserved[];
}

export interface ListLevel {
  level: number;
  start?: number;
  format?: string;
  text?: string;
  align?: string;
  paragraph?: ParagraphProps;
  run?: RunProps;
  attrs?: Attrs;
  extra?: Preserved[];
}

export interface ListDefinition {
  id: string;
  levels: ListLevel[];
  extra?: Preserved[];
}

export interface ListInstance {
  id: string;
  definitionId: string;
  extra?: Preserved[];
}

export interface Lists {
  definitions: ListDefinition[];
  instances: ListInstance[];
  extra?: Preserved[];
}

// --- resources and package -----------------------------------------------------------------

export interface ImageResource {
  contentType: string;
  extension: string;
  /** base64 */
  data: string;
}

export interface PassthroughPart {
  name: string;
  /** base64 */
  data: string;
}

/** Package-level content carried through unchanged (theme, settings, document properties...). */
export interface PackageInfo {
  mainPart: string;
  parts: PassthroughPart[];
  /** Relationships of the main part that point at passthrough content, with original ids. */
  preservedRels: Relationship[];
  contentTypes: { defaults: Record<string, string>; overrides: Record<string, string> };
  /** Namespace declarations and other attributes of the original document root. */
  rootAttrs: Record<string, string>;
}

export interface PrimeDocument {
  schemaVersion: typeof SCHEMA_VERSION;
  styles: StyleSheet;
  lists: Lists;
  sections: Section[];
  images: Record<string, ImageResource>;
  /** Footnotes and endnotes (absent: none). */
  notes?: { footnote?: Notes; endnote?: Notes };
  package: PackageInfo;
}

// --- helpers -------------------------------------------------------------------------------

export function isParagraph(b: Block): b is Paragraph {
  return b.type === "paragraph";
}

export function isTable(b: Block): b is Table {
  return b.type === "table";
}

/** All blocks in reading order, descending into table cells. */
export function* walkBlocks(blocks: Block[]): Generator<Block> {
  for (const b of blocks) {
    yield b;
    if (b.type === "table") {
      for (const row of b.rows) for (const cell of row.cells) yield* walkBlocks(cell.blocks);
    } else if (b.type === "contentControl") yield* walkBlocks(b.blocks);
  }
}

/** Plain text of a paragraph, including tracked insertions and deletions (markup view). */
export function paragraphText(p: Paragraph): string {
  const out: string[] = [];
  const visit = (items: Inline[]) => {
    for (const i of items) {
      if (i.type === "text") out.push(i.text);
      else if (i.type === "tab") out.push("\t");
      else if (i.type === "break" && i.kind === "line") out.push("\n");
      else if (i.type === "link" || i.type === "field" || i.type === "contentControl" || i.type === "revision") visit(i.content);
    }
  };
  visit(p.content);
  return out.join("");
}
