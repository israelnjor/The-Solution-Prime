/**
 * Creating and modifying styles (Home → Styles, as Word's Modify Style and Create a Style). The
 * stylesheet being edited lives in the editor document (its `styles` attribute, which overrides the
 * file's), so a change is one undoable step and every paragraph of the style is redrawn and laid out.
 */
import type { ParagraphProps, PrimeDocument, RunProps, Style, StyleSheet } from "@prime/pdm";
import type { Command, EditorState } from "prosemirror-state";

/** The stylesheet in effect: the editor's own (after a change) or the file's. */
export function stylesOf(state: EditorState, base: PrimeDocument): StyleSheet {
  return (state.doc.attrs.styles as StyleSheet | null) ?? base.styles;
}

/** What the Styles dialog can change in a paragraph style (all optional: only what is given changes). */
export interface StyleChange {
  name?: string;
  basedOn?: string | null;
  next?: string | null;
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Hex RGB, no '#'. */
  color?: string;
  align?: "left" | "center" | "right" | "both";
  /** pt */
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing as a multiple (1, 1.15, 1.5, 2). */
  lineSpacing?: number;
  /** Left indent and first-line indent (negative: hanging), pt. */
  indentLeft?: number;
  firstLine?: number;
}

/** A style with a change applied (paragraph and run properties merged; null removes based-on / next). */
export function changedStyle(style: Style, c: StyleChange): Style {
  const out: Style = { ...style };
  if (c.name !== undefined) out.name = c.name;
  if (c.basedOn !== undefined) { if (c.basedOn) out.basedOn = c.basedOn; else delete out.basedOn; }
  if (c.next !== undefined) { if (c.next) out.next = c.next; else delete out.next; }
  const run: RunProps = { ...(style.run ?? {}) };
  if (c.font !== undefined) run.font = { ...(run.font ?? {}), ascii: c.font, hAnsi: c.font, cs: c.font, asciiTheme: undefined, hAnsiTheme: undefined };
  if (c.size !== undefined) run.size = c.size;
  if (c.bold !== undefined) run.bold = c.bold;
  if (c.italic !== undefined) run.italic = c.italic;
  if (c.underline !== undefined) run.underline = c.underline ? { val: "single" } : { val: "none" };
  if (c.color !== undefined) run.color = c.color;
  if (run.font) for (const k of Object.keys(run.font) as (keyof NonNullable<RunProps["font"]>)[]) if (run.font[k] === undefined) delete run.font[k];
  out.run = run;
  if (style.type === "paragraph") {
    const p: ParagraphProps = { ...(style.paragraph ?? {}) };
    if (c.align !== undefined) p.align = c.align;
    if (c.spaceBefore !== undefined || c.spaceAfter !== undefined || c.lineSpacing !== undefined) {
      p.spacing = { ...(p.spacing ?? {}) };
      if (c.spaceBefore !== undefined) p.spacing.before = c.spaceBefore;
      if (c.spaceAfter !== undefined) p.spacing.after = c.spaceAfter;
      if (c.lineSpacing !== undefined) { p.spacing.line = Math.round(c.lineSpacing * 240); p.spacing.lineRule = "auto"; }
    }
    if (c.indentLeft !== undefined || c.firstLine !== undefined) {
      p.indent = { ...(p.indent ?? {}) };
      if (c.indentLeft !== undefined) p.indent.left = c.indentLeft;
      if (c.firstLine !== undefined) {
        if (c.firstLine < 0) { p.indent.hanging = -c.firstLine; delete p.indent.firstLine; }
        else { p.indent.firstLine = c.firstLine; delete p.indent.hanging; }
      }
    }
    out.paragraph = p;
  }
  return out;
}

/** Modifies style `id` (Word's Modify Style): every paragraph of the style changes with it. */
export function modifyStyle(base: PrimeDocument, id: string, change: StyleChange): Command {
  return (state, dispatch) => {
    const sheet = stylesOf(state, base);
    const style = sheet.styles.find((s) => s.id === id);
    if (!style) return false;
    if (dispatch) {
      const styles = sheet.styles.map((s) => (s.id === id ? changedStyle(s, change) : s));
      dispatch(state.tr.setDocAttribute("styles", { ...sheet, styles }));
    }
    return true;
  };
}

/** A style id for a new style's name: letters and digits only, unique in the sheet. */
function newId(sheet: StyleSheet, name: string): string {
  const stem = name.replace(/[^A-Za-z0-9]/g, "") || "Style";
  let id = stem;
  for (let n = 1; sheet.styles.some((s) => s.id === id); n++) id = `${stem}${n}`;
  return id;
}

/**
 * A new paragraph style from the paragraph at the cursor (Word's Create a Style): based on its
 * style, with its own paragraph formatting and the formatting of the text at the cursor; the
 * paragraph then takes the new style.
 */
export function createStyle(base: PrimeDocument, name: string): Command {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    const para = $from.parent;
    if (para.type.name !== "paragraph" || !name.trim()) return false;
    if (dispatch) {
      const sheet = stylesOf(state, base);
      const props = { ...((para.attrs.props as ParagraphProps | null) ?? {}) };
      const basedOn = props.style ?? sheet.styles.find((s) => s.type === "paragraph" && s.isDefault)?.id;
      const run = ($from.marks().find((m) => m.type.name === "run")?.attrs.props as RunProps | undefined) ?? (para.firstChild?.marks.find((m) => m.type.name === "run")?.attrs.props as RunProps | undefined);
      const id = newId(sheet, name);
      const { style: _old, list: _list, markRun: _mark, ...paragraph } = props;
      const { style: _char, ...ownRun } = run ?? {};
      const style: Style = {
        id, type: "paragraph", name: name.trim(), custom: true, ...(basedOn ? { basedOn } : {}), next: id,
        ...(Object.keys(paragraph).length ? { paragraph } : {}), ...(Object.keys(ownRun).length ? { run: ownRun } : {}),
      };
      const tr = state.tr.setDocAttribute("styles", { ...sheet, styles: [...sheet.styles, style] });
      // The paragraph takes the new style; its direct formatting now lives in the style.
      tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, props: { style: id, ...(props.list ? { list: props.list } : {}) } });
      dispatch(tr);
    }
    return true;
  };
}

/** Modifies several styles in one step (styles missing from the document are left out). */
export function modifyStyles(base: PrimeDocument, changes: Record<string, StyleChange>): Command {
  return (state, dispatch) => {
    const sheet = stylesOf(state, base);
    const ids = new Set(sheet.styles.map((s) => s.id));
    if (!Object.keys(changes).some((id) => ids.has(id))) return false;
    if (dispatch) {
      const styles = sheet.styles.map((s) => (changes[s.id] ? changedStyle(s, changes[s.id]) : s));
      dispatch(state.tr.setDocAttribute("styles", { ...sheet, styles }));
    }
    return true;
  };
}

// --- Design → Document Formatting: style sets, fonts and colours ------------------------------------

/** The headings a style set, font pair or colour palette changes (Word's built-in style ids). */
const HEADINGS = ["Heading1", "Heading2", "Heading3"];

/** The id of the default paragraph style (usually "Normal"). */
function normalId(state: EditorState, base: PrimeDocument): string {
  return stylesOf(state, base).styles.find((s) => s.type === "paragraph" && s.isDefault)?.id ?? "Normal";
}

/** Heading and body fonts. */
export const FONT_PAIRS: { name: string; heading: string; body: string }[] = [
  { name: "Office", heading: "Calibri Light", body: "Calibri" },
  { name: "Aptos", heading: "Aptos Display", body: "Aptos" },
  { name: "Cambria and Calibri", heading: "Cambria", body: "Calibri" },
  { name: "Georgia", heading: "Georgia", body: "Georgia" },
  { name: "Arial", heading: "Arial", body: "Arial" },
  { name: "Times New Roman", heading: "Times New Roman", body: "Times New Roman" },
  { name: "Garamond", heading: "Garamond", body: "Garamond" },
  { name: "Century Gothic", heading: "Century Gothic", body: "Century Gothic" },
  { name: "Segoe UI", heading: "Segoe UI Semibold", body: "Segoe UI" },
];

/** Heading colours: Title, Heading 1, Headings 2 and 3 (hex RGB). */
export const COLOR_SETS: { name: string; title: string; h1: string; h2: string }[] = [
  { name: "Blue", title: "1F3864", h1: "2F5496", h2: "1F3763" },
  { name: "Blue II", title: "17365D", h1: "1F4E79", h2: "2E74B5" },
  { name: "Green", title: "375623", h1: "538135", h2: "385623" },
  { name: "Teal", title: "0B5563", h1: "13808F", h2: "0E6672" },
  { name: "Red", title: "7B1E1E", h1: "C00000", h2: "8B1A1A" },
  { name: "Orange", title: "843C0C", h1: "C55A11", h2: "833C0B" },
  { name: "Purple", title: "3F1D5C", h1: "7030A0", h2: "4B2373" },
  { name: "Grayscale", title: "262626", h1: "404040", h2: "595959" },
];

/** A whole look: fonts, sizes, colours and spacing for the body, title and headings. */
export const STYLE_SETS: { name: string; body: StyleChange; title: StyleChange; subtitle: StyleChange; h1: StyleChange; h2: StyleChange; h3: StyleChange }[] = [
  {
    name: "Basic (Simple)",
    body: { font: "Calibri", size: 11, spaceAfter: 8, lineSpacing: 1.08 },
    title: { font: "Calibri Light", size: 28, color: "000000", bold: false, align: "left", spaceAfter: 0 },
    subtitle: { font: "Calibri", size: 11, color: "5A5A5A", italic: false },
    h1: { font: "Calibri Light", size: 16, color: "2F5496", bold: false, spaceBefore: 12, spaceAfter: 0 },
    h2: { font: "Calibri Light", size: 13, color: "2F5496", bold: false, spaceBefore: 2, spaceAfter: 0 },
    h3: { font: "Calibri Light", size: 12, color: "1F3763", bold: false, spaceBefore: 2, spaceAfter: 0 },
  },
  {
    name: "Basic (Elegant)",
    body: { font: "Garamond", size: 12, spaceAfter: 10, lineSpacing: 1.15 },
    title: { font: "Garamond", size: 30, color: "262626", bold: false, align: "center", spaceAfter: 6 },
    subtitle: { font: "Garamond", size: 13, color: "595959", italic: true },
    h1: { font: "Garamond", size: 18, color: "262626", bold: false, spaceBefore: 18, spaceAfter: 6 },
    h2: { font: "Garamond", size: 14, color: "404040", bold: false, italic: true, spaceBefore: 12, spaceAfter: 4 },
    h3: { font: "Garamond", size: 12, color: "404040", bold: true, spaceBefore: 10, spaceAfter: 2 },
  },
  {
    name: "Basic (Stylish)",
    body: { font: "Century Gothic", size: 10, spaceAfter: 8, lineSpacing: 1.15 },
    title: { font: "Century Gothic", size: 26, color: "1F4E79", bold: true, align: "left", spaceAfter: 4 },
    subtitle: { font: "Century Gothic", size: 12, color: "2E74B5", italic: false },
    h1: { font: "Century Gothic", size: 15, color: "1F4E79", bold: true, spaceBefore: 16, spaceAfter: 4 },
    h2: { font: "Century Gothic", size: 12, color: "2E74B5", bold: true, spaceBefore: 10, spaceAfter: 2 },
    h3: { font: "Century Gothic", size: 11, color: "2E74B5", bold: false, spaceBefore: 8, spaceAfter: 2 },
  },
  {
    name: "Formal Report",
    body: { font: "Times New Roman", size: 12, spaceAfter: 6, lineSpacing: 1.5, align: "both" },
    title: { font: "Times New Roman", size: 20, color: "000000", bold: true, align: "center", spaceAfter: 12 },
    subtitle: { font: "Times New Roman", size: 14, color: "000000", italic: true },
    h1: { font: "Times New Roman", size: 14, color: "000000", bold: true, spaceBefore: 18, spaceAfter: 6 },
    h2: { font: "Times New Roman", size: 12, color: "000000", bold: true, spaceBefore: 12, spaceAfter: 4 },
    h3: { font: "Times New Roman", size: 12, color: "000000", bold: true, italic: true, spaceBefore: 10, spaceAfter: 2 },
  },
  {
    name: "Modern",
    body: { font: "Aptos", size: 11, spaceAfter: 8, lineSpacing: 1.15 },
    title: { font: "Aptos Display", size: 28, color: "0B5563", bold: false, align: "left", spaceAfter: 4 },
    subtitle: { font: "Aptos", size: 12, color: "13808F", italic: false },
    h1: { font: "Aptos Display", size: 18, color: "13808F", bold: false, spaceBefore: 18, spaceAfter: 4 },
    h2: { font: "Aptos Display", size: 14, color: "0E6672", bold: false, spaceBefore: 8, spaceAfter: 4 },
    h3: { font: "Aptos", size: 12, color: "0E6672", bold: true, spaceBefore: 8, spaceAfter: 2 },
  },
  {
    name: "Minimalist",
    body: { font: "Segoe UI", size: 10.5, spaceAfter: 6, lineSpacing: 1.2 },
    title: { font: "Segoe UI Light", size: 26, color: "404040", bold: false, align: "left", spaceAfter: 2 },
    subtitle: { font: "Segoe UI", size: 11, color: "7F7F7F", italic: false },
    h1: { font: "Segoe UI Semibold", size: 14, color: "262626", bold: false, spaceBefore: 14, spaceAfter: 2 },
    h2: { font: "Segoe UI Semibold", size: 12, color: "404040", bold: false, spaceBefore: 10, spaceAfter: 2 },
    h3: { font: "Segoe UI", size: 11, color: "404040", bold: true, spaceBefore: 8, spaceAfter: 0 },
  },
];

/** Applies a style set to the body, title and headings (one step). */
export function applyStyleSet(base: PrimeDocument, name: string): Command {
  return (state, dispatch) => {
    const set = STYLE_SETS.find((s) => s.name === name);
    if (!set) return false;
    return modifyStyles(base, {
      [normalId(state, base)]: set.body, Title: set.title, Subtitle: set.subtitle, Heading1: set.h1, Heading2: set.h2, Heading3: set.h3,
    })(state, dispatch);
  };
}

/** Sets the heading and body fonts (headings and title: the heading font; the body style: the body font). */
export function applyFontPair(base: PrimeDocument, name: string): Command {
  return (state, dispatch) => {
    const pair = FONT_PAIRS.find((p) => p.name === name);
    if (!pair) return false;
    const changes: Record<string, StyleChange> = { [normalId(state, base)]: { font: pair.body }, Title: { font: pair.heading }, Subtitle: { font: pair.body } };
    for (const id of HEADINGS) changes[id] = { font: pair.heading };
    return modifyStyles(base, changes)(state, dispatch);
  };
}

/** Colours the title and headings from a palette. */
export function applyColors(base: PrimeDocument, name: string): Command {
  return (state, dispatch) => {
    const c = COLOR_SETS.find((s) => s.name === name);
    if (!c) return false;
    return modifyStyles(base, { Title: { color: c.title }, Heading1: { color: c.h1 }, Heading2: { color: c.h2 }, Heading3: { color: c.h2 } })(state, dispatch);
  };
}
