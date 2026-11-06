/**
 * Style resolution: computes the effective formatting of paragraphs and runs the way Word does —
 * document defaults -> paragraph style chain -> numbering level -> direct formatting.
 */
import type { FontSpec, Indent, ListLevel, Paragraph, ParagraphProps, PrimeDocument, RunProps, Spacing, Style } from "@prime/pdm";

export interface ResolvedRun {
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  caps: boolean;
  hidden: boolean;
  vertAlign?: string;
  /** Extra space after each character, pt. */
  charSpacing?: number;
  color?: string;
  underline: boolean;
  /** Word's underline style (w:u w:val) when not a single line: "double", "dotted", "wave"… */
  underlineStyle?: string;
  strike: boolean;
  highlight?: string;
}

export interface BorderLine {
  width: number;
  space: number;
  color?: string;
}

export interface ResolvedParagraph {
  styleId: string;
  align: string;
  left: number;
  right: number;
  /** Positive: first-line indent. Negative: hanging indent. */
  firstLine: number;
  spaceBefore: number;
  spaceAfter: number;
  line: number;
  lineRule: string;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
  contextualSpacing: boolean;
  /** Bottom border height (line width + padding) in pt, if any. */
  borderBottom: number;
  /** The bottom border line itself, for drawing. */
  borderLine?: { width: number; space: number; color?: string };
  /** Height a top border adds above the first line (line width + padding). */
  borderTop: number;
  /** Paragraph border sides (line width and padding in pt). */
  borders: Partial<Record<"top" | "bottom" | "left" | "right", BorderLine>>;
  listLabel?: string;
  /** Formatting of the list label: the paragraph mark's, overridden by the list level's (e.g. Symbol bullets). */
  labelRun?: ResolvedRun;
  mark: ResolvedRun;
  /** Table style in effect when the paragraph is inside a table cell. */
  tableStyle?: string;
  /** Text formatting from the table style's conditional formats (the header row's bold…), for its runs. */
  cond?: RunProps;
  /** Background colour (w:shd fill), RRGGBB. */
  shading?: string;
  /** Custom tab stops, left to right: position in pt from the text column's left edge. */
  tabs: TabStop[];
}

export interface TabStop {
  pos: number;
  align: "left" | "center" | "right" | "decimal" | "bar";
  leader?: "dot" | "hyphen" | "underscore" | "middleDot" | "heavy";
}

/** Tab stops from w:tabs XML, applied to `stops` (a w:tab with w:val="clear" removes one). */
function applyTabs(stops: TabStop[], xml: string | undefined): TabStop[] {
  if (!xml) return stops;
  let out = [...stops];
  for (const m of xml.matchAll(/<w:tab\b[^>]*\/?>/g)) {
    const val = /w:val="([^"]+)"/.exec(m[0])?.[1] ?? "left";
    const pos = Number(/w:pos="(-?\d+)"/.exec(m[0])?.[1] ?? NaN) / 20;
    if (Number.isNaN(pos)) continue;
    out = out.filter((t) => Math.abs(t.pos - pos) > 0.5);
    if (val === "clear") continue;
    const leader = /w:leader="([^"]+)"/.exec(m[0])?.[1];
    const align = (val === "end" ? "right" : val === "start" ? "left" : val) as TabStop["align"];
    out.push({ pos, align: ["left", "center", "right", "decimal", "bar"].includes(align) ? align : "left",
      ...(leader && leader !== "none" ? { leader: leader as TabStop["leader"] } : {}) });
  }
  return out.sort((a, b) => a.pos - b.pos);
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];

export function roman(n: number): string {
  let out = "";
  for (const [v, s] of ROMAN) while (n >= v) { out += s; n -= v; }
  return out;
}

function letters(n: number): string {
  const ch = String.fromCharCode(97 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

function formatNumber(n: number, format = "decimal"): string {
  switch (format) {
    case "lowerLetter": return letters(n);
    case "upperLetter": return letters(n).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "none": return "";
    default: return String(n);
  }
}

function mergeFont(a: FontSpec | undefined, b: FontSpec | undefined): FontSpec | undefined {
  if (!b) return a;
  const out: FontSpec = { ...a };
  // An explicit font at a later level overrides an inherited theme font for the same slot.
  if (b.ascii !== undefined && b.asciiTheme === undefined) delete out.asciiTheme;
  if (b.hAnsi !== undefined && b.hAnsiTheme === undefined) delete out.hAnsiTheme;
  return { ...out, ...b };
}

function mergeRun(a: RunProps, b: RunProps | undefined): RunProps {
  if (!b) return a;
  const { font, extra, ...rest } = b;
  const defined = Object.fromEntries(Object.entries(rest).filter(([, v]) => v !== undefined));
  // Theme colours are kept verbatim (w:color with w:themeColor); their w:val holds the computed colour.
  const themed = extra?.find((x) => x.name === "w:color")?.xml.match(/w:val="([0-9A-Fa-f]{6})"/)?.[1]?.toUpperCase();
  return {
    ...a, ...(themed ? { color: themed } : {}), ...defined,
    font: mergeFont(a.font, font), extra: [...(a.extra ?? []), ...(extra ?? [])],
  };
}

function mergeSpacing(a: Spacing | undefined, b: Spacing | undefined): Spacing | undefined {
  if (!b) return a;
  return { ...a, ...b, other: { ...a?.other, ...b.other } };
}

function mergeIndent(a: Indent | undefined, b: Indent | undefined): Indent | undefined {
  if (!b) return a;
  const out: Indent = { ...a };
  if (b.firstLine !== undefined) delete out.hanging;
  if (b.hanging !== undefined) delete out.firstLine;
  return { ...out, ...b };
}

function mergePara(a: ParagraphProps, b: ParagraphProps | undefined): ParagraphProps {
  if (!b) return a;
  const { spacing, indent, extra, markRun, ...rest } = b;
  const defined = Object.fromEntries(Object.entries(rest).filter(([, v]) => v !== undefined));
  return {
    ...a, ...defined,
    spacing: mergeSpacing(a.spacing, spacing),
    indent: mergeIndent(a.indent, indent),
    markRun: markRun ? mergeRun(a.markRun ?? {}, markRun) : a.markRun,
    extra: [...(a.extra ?? []), ...(extra ?? [])],
  };
}

const hasExtra = (p: ParagraphProps, name: string) => (p.extra ?? []).some((x) => x.name === name && !/w:val="(0|false|off)"/.test(x.xml));

export class Resolver {
  private styles = new Map<string, Style>();
  private defaultParagraph?: string;
  private themeFonts: { major?: string; minor?: string } = {};
  private counters = new Map<string, number[]>();

  constructor(private doc: PrimeDocument) {
    for (const s of doc.styles.styles) {
      this.styles.set(s.id, s);
      if (s.type === "paragraph" && s.isDefault) this.defaultParagraph = s.id;
    }
    const theme = doc.package.parts.find((p) => /theme\/theme\d*\.xml$/.test(p.name));
    if (theme) {
      const xml = new TextDecoder().decode(Uint8Array.from(atob(theme.data), (c) => c.charCodeAt(0)));
      this.themeFonts.major = /<a:majorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
      this.themeFonts.minor = /<a:minorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
    }
  }

  /** A style's name ("toc 2", "heading 1"), or its id when it has none. */
  styleName(id: string): string {
    return this.styles.get(id)?.name ?? id;
  }

  private chain(id: string | undefined): Style[] {
    const out: Style[] = [];
    const seen = new Set<string>();
    while (id && this.styles.has(id) && !seen.has(id)) {
      seen.add(id);
      const s = this.styles.get(id)!;
      out.unshift(s);
      id = s.basedOn;
    }
    return out;
  }

  private fontName(f: FontSpec | undefined): string {
    const theme = f?.asciiTheme ?? f?.hAnsiTheme;
    if (theme) return (theme.startsWith("major") ? this.themeFonts.major : this.themeFonts.minor) ?? "Times New Roman";
    return f?.ascii ?? f?.hAnsi ?? "Times New Roman";
  }

  private toRun(p: RunProps): ResolvedRun {
    return {
      font: this.fontName(p.font),
      size: p.size ?? 10,
      bold: !!p.bold,
      italic: !!p.italic,
      caps: !!p.caps,
      hidden: !!p.hidden,
      vertAlign: p.vertAlign && p.vertAlign !== "baseline" ? p.vertAlign : undefined,
      color: p.color && p.color !== "AUTO" ? p.color : undefined,
      underline: !!p.underline && p.underline.val !== "none",
      ...(p.underline?.val && !["single", "none", "words"].includes(p.underline.val) ? { underlineStyle: p.underline.val } : {}),
      strike: !!p.strike || !!p.doubleStrike,
      highlight: p.highlight && p.highlight !== "none" ? p.highlight : undefined,
      ...(p.charSpacing ? { charSpacing: p.charSpacing } : {}),
    };
  }

  /** The table style applied to cells: the table's own, or the document's default table style. */
  tableStyle(id: string | undefined): string | undefined {
    return id ?? this.doc.styles.styles.find((s) => s.type === "table" && s.isDefault)?.id;
  }

  /**
   * Formatting a run sets itself: its character style chain plus direct formatting (nothing inherited
   * from the paragraph). Undefined fields mean "inherit".
   */
  direct(props: RunProps | undefined): Partial<ResolvedRun> {
    if (!props) return {};
    let r: RunProps = {};
    for (const s of this.chain(props.style)) r = mergeRun(r, s.run);
    r = mergeRun(r, props);
    const out: Partial<ResolvedRun> = {};
    // Only a Latin font (ascii/hAnsi or their theme fonts) changes the displayed font; a run that sets
    // just its East Asian or complex-script font keeps the paragraph's.
    const f = r.font;
    if (f && (f.ascii || f.hAnsi || f.asciiTheme || f.hAnsiTheme)) out.font = this.fontName(f);
    if (r.size !== undefined) out.size = r.size;
    if (r.bold !== undefined) out.bold = r.bold;
    if (r.italic !== undefined) out.italic = r.italic;
    if (r.caps !== undefined) out.caps = r.caps;
    if (r.hidden !== undefined) out.hidden = r.hidden;
    if (r.color !== undefined) out.color = r.color === "AUTO" ? undefined : r.color;
    if (r.underline !== undefined) {
      out.underline = r.underline.val !== "none";
      out.underlineStyle = r.underline.val && !["single", "none", "words"].includes(r.underline.val) ? r.underline.val : undefined;
    }
    if (r.strike !== undefined || r.doubleStrike !== undefined) out.strike = !!r.strike || !!r.doubleStrike;
    if (r.highlight !== undefined) out.highlight = r.highlight === "none" ? undefined : r.highlight;
    if (r.vertAlign !== undefined) out.vertAlign = r.vertAlign === "baseline" ? undefined : r.vertAlign;
    if (r.charSpacing !== undefined) out.charSpacing = r.charSpacing || undefined;
    return out;
  }

  /**
   * Effective run formatting inside a paragraph of the given style. Precedence (ECMA-376 §17.7.2):
   * defaults -> table style -> paragraph style -> character style -> direct.
   */
  run(paragraphStyle: string, props: RunProps | undefined, tableStyle?: string, cond?: RunProps): ResolvedRun {
    let r: RunProps = mergeRun({}, this.doc.styles.defaults.run);
    for (const s of this.chain(tableStyle)) r = mergeRun(r, s.run);
    // A table style's conditional formats (header row, banded rows…) sit at the table style's level.
    r = mergeRun(r, cond);
    for (const s of this.chain(paragraphStyle)) r = mergeRun(r, s.run);
    if (props?.style) for (const s of this.chain(props.style)) r = mergeRun(r, s.run);
    return this.toRun(mergeRun(r, props));
  }

  /**
   * Effective paragraph formatting. `countLists`: advance list numbering (a document-order pass);
   * pass false to style a single paragraph without disturbing the counters.
   */
  paragraph(p: Paragraph, tableStyle?: string, countLists = true, cond?: RunProps): ResolvedParagraph {
    const styleId = p.props.style ?? this.defaultParagraph ?? "";
    let props: ParagraphProps = mergePara({}, this.doc.styles.defaults.paragraph);
    for (const s of this.chain(tableStyle)) props = mergePara(props, s.paragraph);
    for (const s of this.chain(styleId)) props = mergePara(props, s.paragraph);
    const list = mergePara(props, p.props).list;
    let label: string | undefined;
    let level: ListLevel | undefined;
    if (list && list.numId && list.numId !== "0") {
      level = this.listLevel(list.numId, list.level ?? 0);
      if (level) {
        props = mergePara(props, level.paragraph);
        if (countLists) label = this.nextLabel(list.numId, list.level ?? 0, level);
      }
    }
    props = mergePara(props, p.props);
    // Tab stops accumulate down the chain (defaults, table style, paragraph style, list level, direct).
    const tabsOf = (x: ParagraphProps | undefined) => x?.extra?.find((e) => e.name === "w:tabs")?.xml;
    let tabs: TabStop[] = applyTabs([], tabsOf(this.doc.styles.defaults.paragraph));
    for (const st of this.chain(tableStyle)) tabs = applyTabs(tabs, tabsOf(st.paragraph));
    for (const st of this.chain(styleId)) tabs = applyTabs(tabs, tabsOf(st.paragraph));
    tabs = applyTabs(tabs, tabsOf(level?.paragraph));
    tabs = applyTabs(tabs, tabsOf(p.props));
    const shd = [...(props.extra ?? [])].reverse().find((x) => x.name === "w:shd")?.xml;
    const shading = shd && /w:fill="([0-9A-Fa-f]{6})"/.exec(shd)?.[1]?.toUpperCase();
    const ind = props.indent ?? {};
    const sp = props.spacing ?? {};
    // The most specific border definition wins (direct formatting after the style's).
    const border = [...(props.extra ?? [])].reverse().find((x) => x.name === "w:pBdr")?.xml;
    const side = (name: string): BorderLine | undefined => {
      const tag = border && new RegExp(`<w:${name}\\b[^>]*>`).exec(border)?.[0];
      if (!tag || /w:val="(none|nil)"/.test(tag)) return undefined;
      return {
        width: Number(/w:sz="(\d+)"/.exec(tag)?.[1] ?? 4) / 8,
        space: Number(/w:space="(\d+)"/.exec(tag)?.[1] ?? 0),
        color: /w:color="([0-9A-Fa-f]{6})"/.exec(tag)?.[1]?.toUpperCase(),
      };
    };
    const borders: ResolvedParagraph["borders"] = {};
    for (const name of ["top", "bottom", "left", "right"] as const) {
      const b = side(name);
      if (b) borders[name] = b;
    }
    const bottom = borders.bottom;
    return {
      styleId,
      align: props.align ?? "left",
      left: ind.left ?? 0,
      right: ind.right ?? 0,
      firstLine: ind.firstLine ?? -(ind.hanging ?? 0),
      spaceBefore: sp.before ?? 0,
      spaceAfter: sp.after ?? 0,
      line: sp.line ?? 240,
      lineRule: sp.lineRule ?? "auto",
      keepNext: !!props.keepNext,
      keepLines: !!props.keepLines,
      pageBreakBefore: !!props.pageBreakBefore,
      // Word applies widow/orphan control unless a document explicitly turns it off.
      widowControl: props.widowControl ?? true,
      contextualSpacing: hasExtra(props, "w:contextualSpacing"),
      borderBottom: bottom ? bottom.width + bottom.space : 0,
      ...(bottom ? { borderLine: bottom } : {}),
      borderTop: borders.top ? borders.top.width + borders.top.space : 0,
      borders,
      listLabel: label,
      labelRun: level && label !== undefined ? this.run(styleId, mergeRun(p.props.markRun ?? {}, level.run), tableStyle, cond) : undefined,
      mark: this.run(styleId, p.props.markRun, tableStyle, cond),
      tableStyle,
      ...(cond ? { cond } : {}),
      ...(shading && shading !== "FFFFFF" ? { shading } : {}),
      tabs,
    };
  }

  private listLevel(numId: string, ilvl: number): ListLevel | undefined {
    const inst = this.doc.lists.instances.find((i) => i.id === numId);
    const def = inst && this.doc.lists.definitions.find((d) => d.id === inst.definitionId);
    return def?.levels.find((l) => l.level === ilvl);
  }

  private nextLabel(numId: string, ilvl: number, level: ListLevel): string {
    const counts = this.counters.get(numId) ?? [];
    for (let i = 0; i <= ilvl; i++) counts[i] ??= (this.listLevel(numId, i)?.start ?? 1) - 1;
    counts[ilvl] += 1;
    counts.length = ilvl + 1; // deeper levels restart
    this.counters.set(numId, counts);
    if (level.format === "bullet") return level.text ?? "•";
    return (level.text ?? `%${ilvl + 1}.`).replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      return formatNumber(counts[i] ?? 1, this.listLevel(numId, i)?.format);
    });
  }
}
