/**
 * What a page shows besides its text (Design tab): page colour, page borders and a watermark — read as
 * Word writes them: the document's w:background, each section's w:pgBorders, and Word's watermark
 * shape (a VML WordArt "PowerPlusWaterMarkObject") in the page's header.
 */
import type { Inline, PrimeDocument, SectionProps, Story } from "@prime/pdm";
import type { FontKey, Measurer } from "./measure.js";

/** One side of a page border: a line from (x1, y1) to (x2, y2), in pt from the page's top-left. */
export interface PageBorderLine {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  /** Line width in pt (for "double": both lines and the space between). */
  width: number;
  color?: string;
  /** single, double, dotted, dashed (others are drawn as single). */
  style: string;
}

/** Text across the page behind the body text (Word's text watermark). */
export interface Watermark {
  text: string;
  font: FontKey;
  /** Font size (pt) that makes the text fill its box's width. */
  size: number;
  /** Hex RGB, no '#'. */
  color: string;
  opacity: number;
  /** Clockwise, in degrees (Word's diagonal watermark: 315). */
  rotation: number;
  /** Centre of the text, in pt from the page's top-left. */
  x: number;
  y: number;
}

export interface PageDecor {
  /** Page colour (hex RGB), on screen and in PDF (as Word's own PDFs). */
  background?: string;
  borders?: PageBorderLine[];
  watermark?: Watermark;
}

/** The document's page colour (w:background/@w:color), if any. */
export function pageColor(doc: PrimeDocument): string | undefined {
  const c = doc.background && /w:color="([0-9A-Fa-f]{6})"/.exec(doc.background.xml)?.[1];
  return c ? c.toUpperCase() : undefined;
}

/**
 * The section's page borders on one of its pages: measured from the page edge or from the text
 * (the margins), on all pages, the first only, or all but the first (w:pgBorders/@w:display).
 */
export function pageBorders(props: SectionProps, width: number, height: number, firstOfSection: boolean): PageBorderLine[] | undefined {
  const xml = props.extra?.find((x) => x.name === "w:pgBorders")?.xml;
  if (!xml) return undefined;
  const display = /w:display="(\w+)"/.exec(xml)?.[1] ?? "allPages";
  if ((display === "firstPage" && !firstOfSection) || (display === "notFirstPage" && firstOfSection)) return undefined;
  const fromPage = /w:offsetFrom="page"/.test(xml);
  const mg = props.margins ?? {};
  const sides: Partial<Record<"top" | "left" | "bottom" | "right", { width: number; space: number; color?: string; style: string }>> = {};
  for (const m of xml.matchAll(/<w:(top|left|bottom|right)\b([^>]*?)\/?>/g)) {
    const a = m[2];
    const style = /w:val="([^"]+)"/.exec(a)?.[1] ?? "single";
    if (style === "none" || style === "nil") continue;
    const color = /w:color="([0-9A-Fa-f]{6})"/.exec(a)?.[1]?.toUpperCase();
    sides[m[1] as "top"] = {
      style, width: Math.max(0.25, Number(/w:sz="(\d+)"/.exec(a)?.[1] ?? 4) / 8), space: Number(/w:space="(\d+)"/.exec(a)?.[1] ?? 0),
      ...(color ? { color } : {}),
    };
  }
  // Where each side's line runs: `space` pt in from the page edge, or out from the text.
  const top = sides.top ? (fromPage ? sides.top.space : (mg.top ?? 72) - sides.top.space) : undefined;
  const bottom = sides.bottom ? (fromPage ? height - sides.bottom.space : height - (mg.bottom ?? 72) + sides.bottom.space) : undefined;
  const left = sides.left ? (fromPage ? sides.left.space : (mg.left ?? 72) - sides.left.space) : undefined;
  const right = sides.right ? (fromPage ? width - sides.right.space : width - (mg.right ?? 72) + sides.right.space) : undefined;
  // A side without a border: the box still ends where it would have been (its default distance).
  const y1 = top ?? (fromPage ? 24 : (mg.top ?? 72) - 1);
  const y2 = bottom ?? (fromPage ? height - 24 : height - (mg.bottom ?? 72) + 1);
  const x1 = left ?? (fromPage ? 24 : (mg.left ?? 72) - 4);
  const x2 = right ?? (fromPage ? width - 24 : width - (mg.right ?? 72) + 4);
  const out: PageBorderLine[] = [];
  const line = (s: (typeof sides)["top"], a: [number, number], b: [number, number]) => {
    if (s) out.push({ x1: a[0], y1: a[1], x2: b[0], y2: b[1], width: s.width, style: s.style, ...(s.color ? { color: s.color } : {}) });
  };
  line(sides.top, [x1, y1], [x2, y1]);
  line(sides.bottom, [x1, y2], [x2, y2]);
  line(sides.left, [x1, y1], [x1, y2]);
  line(sides.right, [x2, y1], [x2, y2]);
  return out.length ? out : undefined;
}

const VML_COLORS: Record<string, string> = {
  silver: "C0C0C0", gray: "808080", grey: "808080", black: "000000", white: "FFFFFF", red: "FF0000", maroon: "800000",
  blue: "0000FF", navy: "000080", green: "008000", lime: "00FF00", yellow: "FFFF00", olive: "808000", teal: "008080",
  aqua: "00FFFF", fuchsia: "FF00FF", purple: "800080",
};

/** The text of Word's watermark shape in a header's VML, or undefined. */
export function watermarkShape(xml: string): { text: string; family: string; color: string; opacity: number; rotation: number; width: number; height: number } | undefined {
  if (!/PowerPlusWaterMarkObject/.test(xml)) return undefined;
  const text = /<v:textpath\b[^>]*\bstring="([^"]*)"/.exec(xml)?.[1];
  if (!text) return undefined; // a picture watermark
  const decode = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const style = decode(/<v:shape\b[^>]*\bstyle="([^"]*)"/.exec(xml)?.[1] ?? "");
  const pt = (name: string, fallback: number) => {
    const m = new RegExp(`(?:^|;)\\s*${name}:\\s*([\\d.]+)(pt|in)?`).exec(style);
    return m ? Number(m[1]) * (m[2] === "in" ? 72 : 1) : fallback;
  };
  const fill = (/<v:shape\b[^>]*\bfillcolor="([^"]*)"/.exec(xml)?.[1] ?? "silver").trim().toLowerCase();
  const color = /^#?[0-9a-f]{6}$/.test(fill) ? fill.replace("#", "").toUpperCase() : VML_COLORS[fill.replace(/\s.*$/, "")] ?? "C0C0C0";
  const opacityRaw = /<v:fill\b[^>]*\bopacity="([^"]*)"/.exec(xml)?.[1];
  const opacity = opacityRaw ? (opacityRaw.endsWith("f") ? Number(opacityRaw.slice(0, -1)) / 65536 : Number(opacityRaw)) : 1;
  const textStyle = decode(/<v:textpath\b[^>]*\bstyle="([^"]*)"/.exec(xml)?.[1] ?? "");
  const family = /font-family:\s*"?([^";]+)"?/.exec(textStyle)?.[1]?.trim() ?? "Calibri";
  return { text: decode(text), family, color, opacity, rotation: pt("rotation", 0), width: pt("width", 468), height: pt("height", 117) };
}

/**
 * The watermark in a header story, placed as Word places its own: centred on the margins' box,
 * the text sized to fill the shape's width.
 */
export function watermarkOf(story: Story | undefined, props: SectionProps, width: number, height: number, measurer: Measurer): Watermark | undefined {
  if (!story) return undefined;
  let found: ReturnType<typeof watermarkShape>;
  const visit = (inlines: Inline[]) => {
    for (const i of inlines) {
      if (found) return;
      if ((i.type === "runPreserved" || i.type === "preserved") && i.preserved) found = watermarkShape(i.preserved.xml);
      else if ("content" in i && Array.isArray(i.content)) visit(i.content as Inline[]);
    }
  };
  for (const b of story.blocks) if (b.type === "paragraph") visit(b.content);
  if (!found) return undefined;
  const font: FontKey = { family: found.family, bold: false, italic: false };
  const atOne = measurer.width(found.text, font) || 1;
  const mg = props.margins ?? {};
  return {
    text: found.text, font, size: Math.min(found.width / atOne, found.height * 1.2), color: found.color, opacity: found.opacity, rotation: found.rotation,
    x: ((mg.left ?? 72) + width - (mg.right ?? 72)) / 2,
    y: ((mg.top ?? 72) + height - (mg.bottom ?? 72)) / 2,
  };
}
