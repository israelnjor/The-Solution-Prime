/**
 * Rebuilding an editable document from PDF page content (ADR-009): text pieces become lines, lines
 * become paragraphs (alignment, indents, spacing), ruled grids become tables, repeated top and bottom
 * lines become headers and footers, bullets and numbers become lists, larger type becomes headings,
 * pictures stay pictures. What cannot be rebuilt faithfully is reported, never silently dropped.
 */
import { blankDocument } from "@prime/editor";
import type {
  Block, Inline, ListDefinition, ListLevel, Paragraph, ParagraphProps, PrimeDocument, RunProps, Style, Table, TableCell,
} from "@prime/pdm";
import type { Fill, PageContent, Picture, Rule, TextPiece } from "./extract.js";

export interface ConversionResult {
  doc: PrimeDocument;
  /** What may need checking, for the conversion preview. */
  warnings: string[];
}

// --- lines ------------------------------------------------------------------------------------

interface Line {
  page: number;
  pieces: TextPiece[];
  /** Baseline, left and right edges, font size (largest on the line). */
  y: number;
  x0: number;
  x1: number;
  size: number;
  /** The previous line ended on purpose (a line break, Shift+Enter in Word), within the same paragraph. */
  breakBefore?: boolean;
}

const top = (l: Line) => l.y - l.size * 0.8;
const bottom = (l: Line) => l.y + l.size * 0.25;
const lineText = (l: Line) => l.pieces.map((p) => p.text).join("");

/** Spacing accents (as TeX draws them: the accent, then the letter) and their combining forms. */
const ACCENTS: Record<string, string> = {
  "¨": "̈", "´": "́", "˝": "̋", "ˆ": "̂", "˜": "̃", "¸": "̧",
  "ˇ": "̌", "˘": "̆", "˙": "̇", "˚": "̊", "¯": "̄", "`": "̀",
};
const ACCENT_THEN_LETTER = new RegExp(`([${Object.keys(ACCENTS).join("")}])(\\p{L})`, "gu");

/**
 * Accented letters drawn as an accent plus a letter ("m¨" "ussen" → "müssen"): pieces that end with an
 * accent join the piece right after them, and accent + letter become one character.
 */
function composeAccents(pieces: TextPiece[]): TextPiece[] {
  const out: TextPiece[] = [];
  for (const p of pieces) {
    const last = out[out.length - 1];
    if (last && /[¨´˝ˆ˜¸ˇ˘˙˚¯]$/.test(last.text) && /^\p{L}/u.test(p.text) &&
        p.x - (last.x + last.width) < p.size * 0.4) {
      out[out.length - 1] = { ...last, text: last.text + p.text, width: p.x + p.width - last.x };
      continue;
    }
    out.push(p);
  }
  return out.map((p) => {
    const text = p.text.replace(ACCENT_THEN_LETTER, (_, a: string, l: string) => (l + ACCENTS[a]).normalize("NFC"));
    return text === p.text ? p : { ...p, text };
  });
}

/** Left and right edges of what is visible: spaces before and after the text do not count. */
function extent(pieces: TextPiece[]): { x0: number; x1: number } | undefined {
  const ink = pieces.filter((p) => p.text.trim());
  if (!ink.length) return undefined;
  const first = ink[0];
  const last = ink[ink.length - 1];
  const lead = first.text.length - first.text.trimStart().length;
  const trail = last.text.length - last.text.trimEnd().length;
  return {
    x0: Math.min(...ink.map((p) => p.x)) + (first.width * lead) / first.text.length,
    x1: Math.max(...ink.map((p) => p.x + p.width)) - (last.width * trail) / last.text.length,
  };
}

/** Pieces sharing a baseline (within a fraction of their size) form a line, left to right. */
function buildLines(page: PageContent, pieces: TextPiece[]): Line[] {
  // Full-size text first, so a line's baseline and size are its text's; smaller pieces (superscripts,
  // subscripts) then join the line they sit on, slightly above or below its baseline.
  const sorted = [...pieces].filter((p) => p.text.length).sort((a, b) => b.size - a.size || a.y - b.y || a.x - b.x);
  const lines: Line[] = [];
  for (const p of sorted) {
    const line = lines.find((l) => Math.abs(l.y - p.y) < Math.max(l.size, p.size) * 0.45);
    if (line) line.pieces.push(p);
    else lines.push({ page: page.number, pieces: [p], y: p.y, x0: p.x, x1: p.x + p.width, size: p.size });
  }
  for (const l of lines) {
    l.pieces.sort((a, b) => a.x - b.x);
    l.pieces = composeAccents(l.pieces);
    Object.assign(l, extent(l.pieces) ?? {});
  }
  // A line made only of spaces carries nothing.
  return lines.filter((l) => lineText(l).trim().length).sort((a, b) => a.y - b.y || a.x0 - b.x0);
}

// --- tables -----------------------------------------------------------------------------------

interface Grid {
  xs: number[];
  ys: number[];
  /** Border presence: h[r][c] = line above row r over column c; v[r][c] = line left of column c in row r. */
  h: boolean[][];
  v: boolean[][];
  fills: Fill[];
}

const cluster = (values: number[], tol: number): number[] => {
  const out: number[] = [];
  for (const v of [...values].sort((a, b) => a - b)) if (!out.length || v - out[out.length - 1] > tol) out.push(v);
  return out;
};

/** Ruled grids on a page (table borders), each with its rows and columns. */
function findGrids(rules: Rule[], fills: Fill[]): Grid[] {
  const hs = rules.filter((r) => Math.abs(r.y2 - r.y1) < 1 && r.x2 - r.x1 > 4);
  const vs = rules.filter((r) => Math.abs(r.x2 - r.x1) < 1 && r.y2 - r.y1 > 4);
  // Connected groups of crossing or touching rules.
  const all = [...hs.map((r) => ({ r, h: true })), ...vs.map((r) => ({ r, h: false }))];
  const parent = all.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const touch = (a: Rule, b: Rule) => a.x1 <= b.x2 + 2 && b.x1 <= a.x2 + 2 && a.y1 <= b.y2 + 2 && b.y1 <= a.y2 + 2;
  for (let i = 0; i < all.length; i++) for (let j = i + 1; j < all.length; j++) if (touch(all[i].r, all[j].r)) parent[find(i)] = find(j);
  const groups = new Map<number, typeof all>();
  all.forEach((x, i) => groups.set(find(i), [...(groups.get(find(i)) ?? []), x]));
  const grids: Grid[] = [];
  for (const g of groups.values()) {
    const gh = g.filter((x) => x.h).map((x) => x.r);
    const gv = g.filter((x) => !x.h).map((x) => x.r);
    if (gh.length < 2 || gv.length < 2) continue;
    const xs = cluster(gv.map((r) => r.x1), 2);
    const ys = cluster(gh.map((r) => r.y1), 2);
    // One box (two lines each way) is a bordered paragraph, not a table.
    if (xs.length < 2 || ys.length < 2 || (xs.length === 2 && ys.length === 2)) continue;
    const near = (a: number, b: number) => Math.abs(a - b) <= 2;
    const covers = (r: Rule, from: number, to: number, horizontal: boolean) =>
      horizontal ? r.x1 <= from + 2 && r.x2 >= to - 2 : r.y1 <= from + 2 && r.y2 >= to - 2;
    const h = ys.map((y) => xs.slice(0, -1).map((x, c) => gh.some((r) => near(r.y1, y) && covers(r, x, xs[c + 1], true))));
    const v = ys.slice(0, -1).map((y, r) => xs.map((x) => gv.some((s) => near(s.x1, x) && covers(s, y, ys[r + 1], false))));
    const box = { x0: xs[0], x1: xs[xs.length - 1], y0: ys[0], y1: ys[ys.length - 1] };
    grids.push({ xs, ys, h, v, fills: fills.filter((f) => f.x >= box.x0 - 2 && f.x + f.w <= box.x1 + 2 && f.y >= box.y0 - 2 && f.y + f.h <= box.y1 + 2) });
  }
  return grids;
}

const inGrid = (g: Grid, x: number, y: number) => x >= g.xs[0] - 1 && x <= g.xs[g.xs.length - 1] + 1 && y >= g.ys[0] - 1 && y <= g.ys[g.ys.length - 1] + 1;

// --- the conversion ---------------------------------------------------------------------------

type Item =
  | { kind: "lines"; lines: Line[]; y: number }
  | { kind: "table"; grid: Grid; lines: Line[]; pictures: Picture[]; y: number; page: number }
  | { kind: "picture"; picture: Picture; y: number; page: number };

const BULLETS = /^[•●○◦▪▫■□‣⁃➢✓✔·o∙]$/;
const NUMBERED = /^(\(?(\d{1,3}|[a-zA-Z]|[ivxlcIVXLC]{1,5})[.)])$/;

/** Text the line starts with if it is a list label (a bullet or "1." "a)" "iv."), else undefined. */
function listLabel(l: Line): { label: string; kind: "bullet" | "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman"; textX: number } | undefined {
  const first = l.pieces[0];
  const text = first.text.trim();
  const words = text.split(/\s+/);
  const label = words[0];
  if (!label) return undefined;
  // Where the text after the label starts: later in the same piece, or the next piece with text.
  const after = words.length > 1 ? first.x + first.width * (text.indexOf(words[1]) / Math.max(text.length, 1)) : l.pieces.slice(1).find((p) => p.text.trim())?.x;
  if (after === undefined || after - first.x < first.size * 0.5) return undefined;
  if (BULLETS.test(label)) return { label, kind: "bullet", textX: after };
  if (!NUMBERED.test(label)) return undefined;
  const core = label.replace(/[().]/g, "");
  const kind = /^\d+$/.test(core) ? "decimal" : /^[ivxlc]+$/.test(core) && core.length > 1 ? "lowerRoman" : /^[IVXLC]+$/.test(core) && core.length > 1 ? "upperRoman"
    : /^[a-z]$/.test(core) ? "lowerLetter" : "upperLetter";
  return { label, kind, textX: after };
}

function romanValue(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  const d = s.toLowerCase().split("").map((c) => v[c] ?? 0);
  return d.reduce((sum, x, k) => sum + (x < (d[k + 1] ?? 0) ? -x : x), 0);
}

/** Whether most of a line's characters are bold (a bold lead-in like "Key point:" does not count). */
function mostlyBold(l: Line): boolean {
  let bold = 0;
  let all = 0;
  for (const p of l.pieces) { const n = p.text.trim().length; all += n; if (p.font.bold) bold += n; }
  return bold * 2 > all;
}

/** Width of a line's first word (estimated from its piece's width). */
function firstWordWidth(l: Line): number {
  const p = l.pieces.find((x) => x.text.trim()) ?? l.pieces[0];
  const text = p.text.trimStart();
  const word = text.split(/\s/)[0] ?? "";
  return text.length ? (p.width * word.length) / text.length : 0;
}

const LINE_FACTOR: Record<string, number> = { calibri: 1.22, "calibri light": 1.22, cambria: 1.17, aptos: 1.2, "times new roman": 1.15, arial: 1.15, georgia: 1.14, verdana: 1.22, "segoe ui": 1.33 };
const factor = (family: string) => LINE_FACTOR[family.toLowerCase()] ?? 1.17;

const round = (v: number, step = 0.5) => Math.round(v / step) * step;

export function reconstruct(pages: PageContent[]): ConversionResult {
  const warnings: string[] = [];
  const doc = blankDocument();
  const section = doc.sections[0];
  if (!pages.length) return { doc, warnings: ["The PDF has no pages."] };

  // --- headers and footers: lines repeated near the top or bottom of most pages --------------
  const pageLines = pages.map((p) => buildLines(p, p.texts));
  // The same text at the same height on most pages; only the page's own number may differ.
  // Only page counters may differ from page to page: the page number, "3/39", "3 of 39".
  const keyOf = (l: Line) => `${Math.round(l.y)}|${lineText(l)
    .replace(/\d+(\s*(\/|of)\s*\d+)/g, "#$1")
    .replace(new RegExp(`(^|\\D)${l.page}(\\D|$)`), "$1#$2")
    .replace(/\s+/g, " ").trim()}`;
  const repeated = (zone: (l: Line, p: PageContent) => boolean) => {
    const counts = new Map<string, number>();
    pageLines.forEach((ls, i) => new Set(ls.filter((l) => zone(l, pages[i])).map(keyOf)).forEach((k) => counts.set(k, (counts.get(k) ?? 0) + 1)));
    const need = Math.max(2, Math.ceil(pages.length * 0.6));
    return new Set([...counts.entries()].filter(([, n]) => n >= need).map(([k]) => k));
  };
  const headerKeys = pages.length > 1 ? repeated((l, p) => l.y < p.height * 0.12) : new Set<string>();
  const footerKeys = pages.length > 1 ? repeated((l, p) => l.y > p.height * 0.88) : new Set<string>();
  const storyFrom = (lines: Line[], page: PageContent): Block[] => lines.map((l) => {
    const text = lineText(l).trim();
    // The page's own number becomes a PAGE field.
    const content: Inline[] = [];
    const m = new RegExp(`(^|\\D)(${page.number})(\\D|$)`).exec(text);
    const props = runProps(l.pieces[0], undefined);
    if (m) {
      const at = m.index + m[1].length;
      if (at) content.push({ type: "text", text: text.slice(0, at), ...(props ? { props } : {}) });
      content.push({ type: "field", instr: " PAGE ", content: [{ type: "text", text: String(page.number), ...(props ? { props } : {}) }] });
      const rest = text.slice(at + String(page.number).length);
      if (rest) content.push({ type: "text", text: rest, ...(props ? { props } : {}) });
    } else content.push({ type: "text", text, ...(props ? { props } : {}) });
    const shade = shadingBehind(l);
    return {
      type: "paragraph", props: {
        align: alignOfLine(l, page),
        ...(shade ? { extra: [{ name: "w:shd", xml: `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` }] } : {}),
      }, content,
    } as Paragraph;
  });

  // --- background areas behind text (banners, highlighted boxes) become paragraph shading ----------
  const backgrounds = new Map(pages.map((pg) => {
    const grids = findGrids(pg.rules, pg.fills);
    return [pg.number, pg.fills.filter((f) => !grids.some((g) => inGrid(g, f.x + f.w / 2, f.y + f.h / 2)))] as const;
  }));
  /** The colour of the smallest background area behind a line, if any. */
  function shadingBehind(l: Line): string | undefined {
    const all = (backgrounds.get(l.page) ?? []).filter((f) => f.y <= top(l) + 2 && f.y + f.h >= l.y - 1);
    let behind = all.filter((f) => f.x <= l.x0 + 1 && f.x + f.w >= l.x1 - 1);
    // Light text with no single box behind it (a bar made of several boxes): the box under most of it.
    const light = l.pieces.some((p) => p.text.trim() && parseInt(p.color, 16) > 0xd0d0d0);
    if (!behind.length && light) {
      const overlap = (f: Fill) => Math.max(0, Math.min(f.x + f.w, l.x1) - Math.max(f.x, l.x0));
      behind = all.filter((f) => overlap(f) > 0).sort((a, b) => overlap(b) - overlap(a)).slice(0, 1);
    }
    // The one drawn last (on top: a box over its own shadow), among the smallest that cover the line.
    const smallest = Math.min(...behind.map((f) => f.w * f.h));
    return behind.filter((f) => f.w * f.h <= smallest * 1.15).pop()?.color;
  }

  // --- page setup: the body's usual edges --------------------------------------------------------
  const first = pages[0];
  const bodyLines = pageLines.map((ls) => ls.filter((l) => !headerKeys.has(keyOf(l)) && !footerKeys.has(keyOf(l))));
  const allBody = bodyLines.flat();
  // The text column: the leftmost and rightmost text (a few lines' margin of error: indented and centred
  // lines start further in), on pages the size of the first (a landscape page is wider).
  const sameSize = bodyLines.filter((_, i) => Math.abs(pages[i].width - first.width) < 1).flat();
  const lefts = sameSize.map((l) => l.x0).sort((a, b) => a - b);
  const left = round(lefts[Math.floor(lefts.length * 0.03)] ?? 72, 1);
  const rights = sameSize.map((l) => l.x1).sort((a, b) => b - a);
  const right = round(rights[Math.floor(rights.length * 0.03)] ?? first.width - 72, 1);
  const tops = bodyLines.map((ls) => (ls.length ? top(ls[0]) : Infinity)).filter(Number.isFinite);
  const bottoms = bodyLines.map((ls) => (ls.length ? bottom(ls[ls.length - 1]) : -Infinity)).filter(Number.isFinite);
  const marginTop = Math.max(18, Math.floor(Math.min(...tops, first.height) - 2));
  const marginBottom = Math.max(18, Math.floor(first.height - Math.max(...bottoms, 0) - 2));
  section.props = {
    pageWidth: round(first.width, 0.05), pageHeight: round(first.height, 0.05),
    orientation: first.width > first.height ? "landscape" : "portrait",
    margins: { top: marginTop, bottom: Math.min(marginBottom, 72), left, right: Math.max(18, round(first.width - right, 1)), header: Math.min(36, marginTop / 2), footer: Math.min(36, marginBottom / 2), gutter: 0 },
  };

  // --- body text style: the most common font and size ----------------------------------------
  const pieces = allBody.flatMap((l) => l.pieces).filter((p) => p.text.trim());
  const weight = (p: TextPiece) => p.text.length;
  const tally = <K>(key: (p: TextPiece) => K) => {
    const m = new Map<K, number>();
    for (const p of pieces) m.set(key(p), (m.get(key(p)) ?? 0) + weight(p));
    return [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
  };
  const bodyFont = tally((p) => p.font.family) ?? "Calibri";
  const bodySize = tally((p) => round(p.size)) ?? 11;
  doc.styles.defaults.run = { font: { ascii: bodyFont, hAnsi: bodyFont, eastAsia: bodyFont, cs: bodyFont }, size: bodySize };

  // Header and footer stories (their formatting is relative to the body text's).
  const firstWithHeader = pageLines.findIndex((ls) => ls.some((l) => headerKeys.has(keyOf(l))));
  if (firstWithHeader >= 0) section.headers.default = { blocks: storyFrom(pageLines[firstWithHeader].filter((l) => headerKeys.has(keyOf(l))), pages[firstWithHeader]) };
  const firstWithFooter = pageLines.findIndex((ls) => ls.some((l) => footerKeys.has(keyOf(l))));
  if (firstWithFooter >= 0) section.footers.default = { blocks: storyFrom(pageLines[firstWithFooter].filter((l) => footerKeys.has(keyOf(l))), pages[firstWithFooter]) };

  // Pictures at the same place near the top or bottom of most pages (a logo): the header's or footer's.
  const picKey = (p: Picture) => [p.x, p.y, p.w, p.h].map((v) => Math.round(v)).join(",");
  const picCounts = new Map<string, number>();
  for (const pg of pages) new Set(pg.pictures.map(picKey)).forEach((k) => picCounts.set(k, (picCounts.get(k) ?? 0) + 1));
  const chromePictures = new Set<string>();
  if (pages.length > 1) {
    for (const [k, n] of picCounts) {
      if (n < Math.max(2, Math.ceil(pages.length * 0.6))) continue;
      const pg = pages.find((x) => x.pictures.some((p) => picKey(p) === k))!;
      const pic = pg.pictures.find((p) => picKey(p) === k)!;
      const inTop = pic.y + pic.h < pg.height * 0.3;
      if (!inTop && pic.y < pg.height * 0.7) continue;
      const story = inTop ? (section.headers.default ??= { blocks: [] }) : (section.footers.default ??= { blocks: [] });
      story.blocks.unshift(pictureParagraph(pic, left, right));
      chromePictures.add(k);
    }
  }

  /** Run formatting that differs from a paragraph's style (base). */
  function runProps(p: TextPiece, base: { family: string; size: number; bold: boolean; italic: boolean; color: string } | undefined): RunProps | undefined {
    const b = base ?? { family: bodyFont, size: bodySize, bold: false, italic: false, color: "000000" };
    const out: RunProps = {};
    if (p.font.family !== b.family) out.font = { ascii: p.font.family, hAnsi: p.font.family, cs: p.font.family };
    if (Math.abs(round(p.size) - b.size) > 0.2) out.size = round(p.size);
    if (p.font.bold !== b.bold) out.bold = p.font.bold;
    if (p.font.italic !== b.italic) out.italic = p.font.italic;
    if (p.color !== b.color) out.color = p.color;
    if (p.underline) out.underline = { val: "single" };
    return Object.keys(out).length ? out : undefined;
  }

  // --- headings: paragraph sizes clearly above the body text ------------------------------
  const headingSizes = [...new Set(allBody.filter((l) => l.size > bodySize + 1 && lineText(l).length < 120).map((l) => round(l.size)))].sort((a, b) => b - a);
  const titleSize = headingSizes[0] !== undefined && headingSizes[0] >= 20 && headingSizes.length > 1 ? headingSizes[0] : undefined;
  const levels = headingSizes.filter((s) => s !== titleSize).slice(0, 3);
  const styleFor = (size: number) => (titleSize !== undefined && size === titleSize ? "Title" : levels.includes(size) ? `Heading${levels.indexOf(size) + 1}` : undefined);
  const styleLook = new Map<string, { family: string; size: number; bold: boolean; italic: boolean; color: string }>();

  // --- items in reading order per page: paragraphs' lines, tables, pictures --------------------
  const items: Item[] = [];
  pages.forEach((page, i) => {
    const grids = findGrids(page.rules, page.fills);
    const lines = bodyLines[i];
    const used = new Set<Line>();
    for (const g of grids) {
      const inside = lines.filter((l) => inGrid(g, (l.x0 + l.x1) / 2, l.y - l.size * 0.3));
      inside.forEach((l) => used.add(l));
      const pics = page.pictures.filter((pic) => inGrid(g, pic.x + pic.w / 2, pic.y + pic.h / 2));
      items.push({ kind: "table", grid: g, lines: inside, pictures: pics, y: g.ys[0], page: page.number });
    }
    for (const pic of page.pictures) {
      if (chromePictures.has(picKey(pic)) || grids.some((g) => inGrid(g, pic.x + pic.w / 2, pic.y + pic.h / 2))) continue;
      items.push({ kind: "picture", picture: pic, y: pic.y, page: page.number });
    }
    let run: Line[] = [];
    const flush = () => { if (run.length) items.push({ kind: "lines", lines: run, y: top(run[0]) }); run = []; };
    const pageItems = items.filter((x) => (x.kind === "table" || x.kind === "picture") && x.page === page.number);
    for (const l of lines) {
      if (used.has(l)) continue;
      // A table or picture between two lines ends the run of lines.
      if (run.length && pageItems.some((x) => x.y > top(run[run.length - 1]) && x.y < top(l))) flush();
      run.push(l);
    }
    flush();
    // A scanned page: no text, one picture covering most of the page (a full-page screenshot in a typed
    // document has a caption or margins around it).
    const scan = page.pictures.some((pic) => pic.w > page.width * 0.85 && pic.h > page.height * 0.85);
    if (!page.texts.some((t) => t.text.trim()) && scan) {
      warnings.push(`Page ${page.number} is a scanned image: its text cannot be edited until text recognition is added.`);
    }
  });
  items.sort((a, b) => pageOf(a) - pageOf(b) || a.y - b.y);
  function pageOf(x: Item) { return x.kind === "lines" ? x.lines[0].page : x.page; }

  // --- paragraphs from lines --------------------------------------------------------------------
  const blocks: Block[] = [];
  const lists = new Map<string, string>(); // list kind → numId
  const pageHeight = first.height;
  let previousBottom: { page: number; y: number } | undefined;

  function alignOfLine(l: Line, page: PageContent): string | undefined {
    const center = (l.x0 + l.x1) / 2;
    const mid = (left + right) / 2;
    if (l.x0 > left + 12 && Math.abs(center - mid) < 6) return "center";
    if (l.x0 > left + 24 && Math.abs(l.x1 - right) < 4) return "right";
    void page;
    return undefined;
  }

  /** Where each paragraph's first line starts on the page (contents levels are read from it). */
  const firstX = new WeakMap<Paragraph, number>();
  const paragraphOf = (lines: Line[], label?: ReturnType<typeof listLabel>, styleId?: string): Paragraph => {
    const firstLine = lines[0];
    const look = styleId ? styleLook.get(styleId) : undefined;
    const content: Inline[] = [];
    const push = (text: string, p: TextPiece) => {
      if (!text) return;
      const props = runProps(p, look);
      // Text in a link area goes inside a link (consecutive pieces with the same address share one).
      let into = content;
      if (p.link) {
        const lastLink = content[content.length - 1];
        if (lastLink?.type === "link" && lastLink.href === p.link) into = lastLink.content;
        else { const link: Inline = { type: "link", href: p.link, content: [] }; content.push(link); into = (link as { content: Inline[] }).content; }
      }
      const last = into[into.length - 1];
      if (last?.type === "text" && JSON.stringify(last.props ?? null) === JSON.stringify(props ?? null)) last.text += text;
      else into.push({ type: "text", text, ...(props ? { props } : {}) });
    };
    lines.forEach((l, li) => {
      let pieces = l.pieces;
      if (li === 0 && label) {
        // Drop the label (the list draws it) and any spaces after it.
        const [head, ...rest] = pieces;
        const remainder = head.text.trim().slice(label.label.length).replace(/^\s+/, "");
        pieces = remainder ? [{ ...head, text: remainder, x: label.textX }, ...rest] : rest;
        while (pieces.length && !pieces[0].text.trim()) pieces = pieces.slice(1);
        if (pieces.length) pieces = [{ ...pieces[0], text: pieces[0].text.replace(/^\s+/, "") }, ...pieces.slice(1)];
      }
      if (li > 0 && l.breakBefore) {
        const last = content[content.length - 1];
        if (last?.type === "text") last.text = last.text.replace(/\s+$/, "");
        content.push({ type: "break", kind: "line" });
      } else if (li > 0) {
        const prev = content[content.length - 1];
        const ends = prev?.type === "text" ? prev.text : "";
        // Lines join with a space, except after a hyphen or one already there.
        if (ends && !/[\s\-‐‑]$/.test(ends)) push(" ", pieces[0] ?? l.pieces[0]);
      }
      pieces.forEach((p, k) => {
        const prev = pieces[k - 1];
        if (prev) {
          const gap = p.x - (prev.x + prev.width);
          const spaced = /\s$/.test(prev.text) || /^\s/.test(p.text);
          if (gap > p.size * 2.5) content.push({ type: "tab" });
          else if (gap > p.size * 0.15 && !spaced) push(" ", p);
        }
        // Raised or lowered smaller text: superscript / subscript.
        const shifted = p.size < l.size * 0.8 ? (p.y < l.y - l.size * 0.2 ? "superscript" : p.y > l.y + l.size * 0.1 ? "subscript" : undefined) : undefined;
        if (shifted) {
          const props = { ...(runProps({ ...p, size: l.size }, look) ?? {}), vertAlign: shifted } as RunProps;
          content.push({ type: "text", text: p.text, props });
        } else push(p.text, p);
      });
    });
    // Paragraph layout.
    const props: ParagraphProps = {};
    if (styleId) props.style = styleId;
    const rest = lines.slice(1);
    const textLeft = label ? label.textX : rest.length ? Math.min(...rest.map((l) => l.x0)) : firstLine.x0;
    // List items can be justified too (their label does not count as centred or right-aligned text).
    const found = alignOfParagraph(lines);
    const align = label && found !== "both" ? undefined : found;
    if (align) props.align = align;
    if (!align || align === "both" || align === "left") {
      const indentLeft = round(textLeft - left, 0.5);
      const firstLineIndent = label ? round(firstLine.x0 - textLeft, 0.5) : round(firstLine.x0 - textLeft, 0.5);
      const indent: ParagraphProps["indent"] = {};
      if (indentLeft > 1) indent.left = indentLeft;
      if (Math.abs(firstLineIndent) > 1) {
        if (firstLineIndent > 0) indent.firstLine = firstLineIndent;
        else indent.hanging = -firstLineIndent;
      }
      if (Object.keys(indent).length) props.indent = indent;
    }
    if (lines.length > 1) {
      const pitch = (lines[lines.length - 1].y - firstLine.y) / (lines.length - 1);
      const natural = firstLine.size * factor(firstLine.pieces[0].font.family);
      const multiple = pitch / natural;
      if (Math.abs(multiple - 1) > 0.04) props.spacing = { line: Math.round(240 * multiple), lineRule: "auto" };
      else props.spacing = { line: 240, lineRule: "auto" };
    }
    if (label) props.list = { numId: label.kind === "bullet" ? listId("bullet") : numberedList(label.label), level: 0 };
    const shade = shadingBehind(firstLine);
    if (shade) props.extra = [...(props.extra ?? []), { name: "w:shd", xml: `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` }];
    const paragraph: Paragraph = { type: "paragraph", props, content };
    firstX.set(paragraph, firstLine.x0);
    return paragraph;
  };

  function alignOfParagraph(lines: Line[]): string | undefined {
    if (lines.length === 1) return alignOfLine(lines[0], first);
    const body = lines.slice(0, -1);
    const full = body.every((l) => Math.abs(l.x1 - right) < 3);
    if (lines.every((l) => Math.abs((l.x0 + l.x1) / 2 - (left + right) / 2) < 6) && lines.some((l) => l.x0 > left + 12)) return "center";
    if (lines.every((l) => Math.abs(l.x1 - right) < 3) && lines.some((l) => l.x0 > left + 12)) return "right";
    if (full && body.length >= 1) {
      // Justified lines stretch their spaces; left-aligned text that happens to fill lines does not.
      const spread = body.some((l) => l.pieces.length > 1 && l.pieces.some((p, k) => k > 0 && p.x - (l.pieces[k - 1].x + l.pieces[k - 1].width) > p.size * 0.6));
      return spread || lines.length > 2 ? "both" : undefined;
    }
    return undefined;
  }

  function listId(kind: NonNullable<ReturnType<typeof listLabel>>["kind"]): string {
    let id = lists.get(kind);
    if (id) return id;
    id = String(doc.lists.instances.length + 1);
    lists.set(kind, id);
    const def: ListDefinition = {
      id: `${id}`,
      levels: Array.from({ length: 9 }, (_, level) => ({
        level, start: 1, format: kind,
        text: kind === "bullet" ? ["•", "o", "▪"][level % 3] : `%${level + 1}.`,
        align: "left",
        paragraph: { indent: { left: 36 * (level + 1), hanging: 18 } },
        ...(kind === "bullet" && level % 3 === 1 ? { run: { font: { ascii: "Courier New", hAnsi: "Courier New" } } } : {}),
      })),
    };
    doc.lists.definitions.push(def);
    doc.lists.instances.push({ id, definitionId: def.id });
    return id;
  }

  /**
   * A numbered list item: its label's style ("(a)", "ii.", "3)") and number decide the list. The next
   * number in the same style continues the list; anything else starts a new one at its own number.
   * A lone "i" (or "v", "x") is a letter after "h" (or "u", "w"), otherwise a roman numeral.
   */
  let lastNumbered: { pattern: string; kind: string; value: number; numId: string } | undefined;
  function numberedList(label: string): string {
    const core = /[0-9]+|[a-zA-Z]+/.exec(label)?.[0] ?? "1";
    const pattern = label.replace(core, "%1");
    const letterValue = core.length === 1 ? core.toLowerCase().charCodeAt(0) - 96 : 0;
    let kind: ListLevel["format"];
    let value: number;
    if (/^\d+$/.test(core)) { kind = "decimal"; value = Number(core); }
    else if (/^[ivxlc]+$/i.test(core) && !(core.length === 1 && lastNumbered?.pattern === pattern &&
        /Letter$/.test(lastNumbered.kind) && lastNumbered.value === letterValue - 1) && (core.length > 1 || /^[ivx]$/i.test(core))) {
      kind = core === core.toLowerCase() ? "lowerRoman" : "upperRoman";
      value = romanValue(core);
    } else { kind = core === core.toLowerCase() ? "lowerLetter" : "upperLetter"; value = letterValue; }
    if (lastNumbered && lastNumbered.pattern === pattern && lastNumbered.kind === kind && value === lastNumbered.value + 1) {
      lastNumbered.value = value;
      return lastNumbered.numId;
    }
    const numId = String(doc.lists.instances.length + 1);
    doc.lists.definitions.push({
      id: numId,
      levels: Array.from({ length: 9 }, (_, level) => ({
        level, start: level === 0 ? value : 1, format: kind, text: level === 0 ? pattern : `%${level + 1}.`, align: "left",
        paragraph: { indent: { left: 36 * (level + 1), hanging: 18 } },
      })),
    });
    doc.lists.instances.push({ id: numId, definitionId: numId });
    lastNumbered = { pattern, kind: kind ?? "decimal", value, numId };
    return numId;
  }

  /** Space between the end of what came before and this block's first line (pt). */
  const gapBefore = (page: number, y: number) =>
    previousBottom && previousBottom.page === page ? Math.max(0, y - previousBottom.y) : 0;

  // Paragraph styles: each heading style looks like its first use; spacing comes from the paragraphs.
  const paragraphs: { lines: Line[]; label?: ReturnType<typeof listLabel>; style?: string }[] = [];
  const pushParagraph = (lines: Line[], label?: ReturnType<typeof listLabel>) => {
    const style = label ? undefined : styleFor(round(lines[0].size));
    if (style && !styleLook.has(style)) {
      const p = lines[0].pieces[0];
      styleLook.set(style, { family: p.font.family, size: round(lines[0].size), bold: p.font.bold, italic: p.font.italic, color: p.color });
    }
    paragraphs.push({ lines, label, style });
  };

  for (const item of items) {
    if (item.kind !== "lines") continue;
    // Group lines into paragraphs.
    let cur: Line[] = [];
    let label: ReturnType<typeof listLabel>;
    const close = () => { if (cur.length) pushParagraph(cur, label); cur = []; label = undefined; };
    for (const l of item.lines) {
      // A heading that starts "1." is a numbered heading, not a list.
      const lab = styleFor(round(l.size)) ? undefined : listLabel(l);
      const prev = cur[cur.length - 1];
      if (prev) {
        const pitch = l.y - prev.y;
        const expected = cur.length > 1 ? (prev.y - cur[0].y) / (cur.length - 1) : prev.size * 1.25;
        const sameLook = Math.abs(l.size - prev.size) < 0.6 && mostlyBold(l) === mostlyBold(prev);
        // Word moves a word to the next line only when it does not fit: if the next line's first word
        // would have fitted after the previous line (within this paragraph's right edge), that line
        // ended its paragraph.
        const edge = Math.max(...cur.map((c) => c.x1), l.x1);
        // Also a short line (an address, a sign-off) whose next word would fit before the margin.
        const fits = (limit: number) => prev.x1 + firstWordWidth(l) + l.size * 0.25 < limit - 1;
        const prevEndsShort = fits(edge) || (prev.x1 - left < (right - left) * 0.6 && fits(right));
        // Room left on the previous line, whatever its alignment: if the next word would have fitted there,
        // the line ended on purpose.
        // (For centred and right-aligned lines, measured against the whole column: their left edges vary.)
        const room = (right - left) - (prev.x1 - prev.x0);
        const alignedRightOrCentre = prev.x0 > left + 12 && (Math.abs(prev.x1 - right) < 4 || Math.abs((prev.x0 + prev.x1) / 2 - (left + right) / 2) < 6);
        const endedEarly = prevEndsShort || (alignedRightOrCentre && firstWordWidth(l) + l.size * 0.25 < room - 1);
        const close_ = pitch > 0 && pitch < Math.max(expected * 1.25, prev.size * 1.5);
        const continues = sameLook && !lab && close_ && (label ? Math.abs(l.x0 - label.textX) < 3 : !endedEarly);
        // Ended early but no further away than the lines of a paragraph: a line break, not a new paragraph.
        const tight = cur.length > 1 ? pitch < expected * 1.08 : pitch < prev.size * 1.32;
        const lineBreak = !continues && sameLook && !lab && !label && close_ && endedEarly && tight;
        if (lineBreak) l.breakBefore = true;
        else if (!continues) close();
      }
      if (!cur.length) label = lab;
      cur.push(l);
    }
    close();
    void item;
  }

  // Styles for headings (and the Title), from their first use.
  for (const [id, look] of styleLook) {
    const style = doc.styles.styles.find((s) => s.id === id) as Style | undefined;
    if (!style) continue;
    style.run = {
      font: { ascii: look.family, hAnsi: look.family, cs: look.family }, size: look.size,
      bold: look.bold, italic: look.italic, ...(look.color !== "000000" ? { color: look.color } : {}),
    };
  }

  // Blocks in order, with the space between them; a page's paragraph that runs on joins the next page's.
  const byFirstLine = new Map(paragraphs.map((p) => [p.lines[0], p]));
  let pendingBreak = false;
  let lastPage = 0;
  for (const item of items) {
    const page = pageOf(item);
    if (page !== lastPage) {
      // Much empty space at the bottom of the previous page: the next page starts with a break.
      if (lastPage && previousBottom && previousBottom.y < pageHeight * 0.75) pendingBreak = true;
      lastPage = page;
      previousBottom = undefined;
    }
    let block: Block | undefined;
    let y = item.y;
    let endY = item.y;
    if (item.kind === "lines") {
      for (const l of item.lines) {
        const para = byFirstLine.get(l);
        if (!para) continue;
        const p = paragraphOf(para.lines, para.label, para.style);
        const startY = top(para.lines[0]);
        const gap = gapBefore(page, startY);
        const last = blocks[blocks.length - 1];
        // Continuation across a page: no break pending, same look, previous paragraph ran to the margin.
        if (!previousBottom && last?.type === "paragraph" && !pendingBreak && !para.label && !para.style && !last.props.style && !last.props.list) {
          const prevText = last.content.map((c) => (c.type === "text" ? c.text : "")).join("");
          if (prevText && !/[.!?:]["')\]]?$/.test(prevText.trim())) {
            last.content.push({ type: "text", text: " " }, ...p.content);
            previousBottom = { page, y: bottom(para.lines[para.lines.length - 1]) };
            continue;
          }
        }
        if (gap > 0) setSpaceBefore(p, gap, para.lines[0]);
        if (pendingBreak) { p.props.pageBreakBefore = true; pendingBreak = false; }
        blocks.push(p);
        previousBottom = { page, y: bottom(para.lines[para.lines.length - 1]) };
      }
      continue;
    }
    if (item.kind === "table") {
      const table = tableOf(item.grid, item.lines, item.pictures);
      endY = item.grid.ys[item.grid.ys.length - 1];
      // A table that runs on from the previous page (same columns, nothing between): one table again,
      // without the header row Word repeats on each page.
      const last = blocks[blocks.length - 1];
      if (!previousBottom && !pendingBreak && last?.type === "table" && last.grid.length === table.grid.length &&
          last.grid.every((w, k) => Math.abs(w - table.grid[k]) < 2)) {
        const text = (r: Table["rows"][number]) => JSON.stringify(r.cells.map((c) => c.blocks));
        const rows = table.rows.length > 1 && text(table.rows[0]) === text(last.rows[0]) ? table.rows.slice(1) : table.rows;
        if (rows !== table.rows) last.rows[0].props = { ...last.rows[0].props, header: true };
        last.rows.push(...rows);
        previousBottom = { page, y: endY };
        continue;
      }
      block = table;
    } else {
      const pic = item.picture;
      block = pictureParagraph(pic, left, right);
      endY = pic.y + pic.h;
      y = pic.y;
    }
    if (block.type === "paragraph") {
      const gap = gapBefore(page, y);
      if (gap > 0) block.props.spacing = { ...block.props.spacing, before: round(gap, 0.5) };
      if (pendingBreak) { block.props.pageBreakBefore = true; pendingBreak = false; }
    } else if (pendingBreak) {
      blocks.push({ type: "paragraph", props: { pageBreakBefore: true, spacing: { after: 0, line: 240, lineRule: "auto" } }, content: [] });
      pendingBreak = false;
    }
    blocks.push(block);
    previousBottom = { page, y: endY };
  }

  /** A picture in a paragraph of its own, aligned as it sits between `from` and `to`. */
  function pictureParagraph(pic: Picture, from: number, to: number): Paragraph {
    const extension = pic.contentType === "image/jpeg" ? "jpeg" : "png";
    const id = `media/import${Object.keys(doc.images).length + 1}.${extension}`;
    doc.images[id] = { contentType: pic.contentType, extension, data: base64(pic.data) };
    const center = pic.x + pic.w / 2;
    const align = Math.abs(center - (from + to) / 2) < 12 && pic.x > from + 12 ? "center" : pic.x > from + 24 && Math.abs(pic.x + pic.w - to) < 6 ? "right" : undefined;
    return { type: "paragraph", props: align ? { align } : {}, content: [{ type: "image", imageId: id, width: round(pic.w, 0.1), height: round(pic.h, 0.1) }] };
  }

  /** Space before a paragraph: what the gap leaves beyond the line's own height. */
  function setSpaceBefore(p: Paragraph, gap: number, l: Line) {
    const natural = l.size * factor(l.pieces[0].font.family);
    const extra = gap - (natural - l.size * 1.05);
    if (extra > 1) p.props.spacing = { ...p.props.spacing, before: round(extra, 0.5) };
  }

  // Word's default space after paragraphs would add to the measured gaps: none by default here.
  doc.styles.defaults.paragraph = { spacing: { after: 0, line: 240, lineRule: "auto" } };
  rebuildContents(blocks);
  section.blocks = blocks.length ? blocks : [{ type: "paragraph", props: {}, content: [] }];

  /**
   * Tables of contents: three or more paragraphs in a row that end in a page number after a leader
   * (dots) or a wide gap become Word contents entries — styles TOC 1–3 (by indent), a right tab stop
   * with dots at the margin — inside a TOC field, so Word can update the table from the headings.
   */
  function rebuildContents(list: Block[]) {
    const ENTRY = /^(.*?\S)(?:[\s.·…_]*?(?:\.{3,}|·{3,}|…+|_{3,})[\s.·…_]*|\t[\s.]*)([0-9]{1,4}|[ivxlc]{1,6})\s*$/i;
    const plain = (b: Block) => (b.type === "paragraph" ? b.content.map((c) => (c.type === "text" ? c.text : c.type === "tab" ? "\t" : c.type === "link" ? c.content.map((x) => (x.type === "text" ? x.text : x.type === "tab" ? "\t" : "")).join("") : "")).join("") : "");
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j < list.length && list[j].type === "paragraph" && ENTRY.test(plain(list[j]))) j++;
      if (j - i >= 3) {
        const entries = list.slice(i, j) as Paragraph[];
        // Level by where each entry's first line starts (left indent plus first-line indent), shallowest first.
        const startOf = (e: Paragraph) => Math.round((firstX.get(e) ?? 0) / 3);
        const indents = [...new Set(entries.map(startOf))].sort((a, b) => a - b);
        // A list of figures or tables ("Figure 2 …") is a table of figures: Word's TOC field with \c.
        const caption = /^(Figure|Table|Fig\.)\s+\d+/i.exec(plain(entries[0]))?.[1];
        entries.forEach((e, k) => {
          const m = ENTRY.exec(plain(e))!;
          const level = caption ? 1 : Math.min(3, indents.indexOf(startOf(e)) + 1);
          const firstText = [e.content, ...e.content.map((c) => (c.type === "link" ? c.content : []))].flat().find((c): c is Extract<Inline, { type: "text" }> => c.type === "text");
          const props = firstText?.props ? { props: firstText.props } : {};
          const content: Inline[] = [];
          if (k === 0) {
            const instr = caption ? ` TOC \\h \\z \\c "${caption === "Fig." ? "Figure" : caption}" ` : ' TOC \\o "1-3" \\h \\z \\u ';
            content.push({ type: "fieldChar", kind: "begin" }, { type: "instr", text: instr }, { type: "fieldChar", kind: "separate" });
          }
          content.push({ type: "text", text: m[1].trim(), ...props }, { type: "tab" }, { type: "text", text: m[2], ...props });
          if (k === entries.length - 1) content.push({ type: "fieldChar", kind: "end" });
          const tabPos = Math.round((right - left) * 20);
          e.props = {
            style: `TOC${level}`,
            ...(e.props.spacing ? { spacing: e.props.spacing } : {}),
            extra: [{ name: "w:tabs", xml: `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tabPos}"/></w:tabs>` }],
          };
          e.content = content;
          ensureTocStyle(level);
        });
      }
      i = Math.max(j, i + 1);
    }
  }

  function ensureTocStyle(level: number) {
    const id = `TOC${level}`;
    if (doc.styles.styles.some((st) => st.id === id)) return;
    doc.styles.styles.push({
      id, type: "paragraph", name: `toc ${level}`, basedOn: "Normal", next: "Normal",
      paragraph: { spacing: { after: 5 }, ...(level > 1 ? { indent: { left: 11 * (level - 1) } } : {}) },
    });
  }

  // --- tables ---------------------------------------------------------------------------------
  function tableOf(g: Grid, lines: Line[], pictures: Picture[]): Table {
    const cols = g.xs.length - 1;
    const rows = g.ys.length - 1;
    const grid = g.xs.slice(1).map((x, c) => round(x - g.xs[c], 0.05));
    const cellOf = (x: number, y: number) => {
      const c = g.xs.findIndex((v, k) => k < cols && x >= v - 1 && x < g.xs[k + 1] + 1);
      const r = g.ys.findIndex((v, k) => k < rows && y >= v - 1 && y < g.ys[k + 1] + 1);
      return [r, c];
    };
    const owned = new Map<string, Line[]>();
    for (const l of lines) {
      // Pieces may fall in different cells (a borderless line spanning a row): split by cell.
      for (const p of l.pieces) {
        const [r, c] = cellOf(p.x + Math.min(p.width, 4) / 2, p.y - p.size * 0.3);
        if (r < 0 || c < 0) continue;
        const key = `${r}:${c}`;
        const list = owned.get(key) ?? [];
        const same = list.find((x) => x.y === l.y);
        if (same) { same.pieces.push(p); Object.assign(same, extent(same.pieces) ?? {}); }
        else list.push({ ...l, pieces: [p], ...(extent([p]) ?? { x0: p.x, x1: p.x + p.width }) });
        owned.set(key, list);
      }
    }
    const fillAt = (r: number, c: number) => g.fills.find((f) => f.x <= g.xs[c] + 3 && f.x + f.w >= g.xs[c + 1] - 3 && f.y <= g.ys[r] + 3 && f.y + f.h >= g.ys[r + 1] - 3)?.color;
    const out: Table = {
      type: "table",
      props: { style: "TableGrid", width: { w: "0", type: "auto" }, extra: [{ name: "w:tblBorders", xml: '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>' }] },
      grid,
      rows: [],
    };
    // Merges: no line between two grid cells means one cell spans both.
    const spanRight = (r: number, c: number) => { let n = 1; while (c + n < cols && !g.v[r][c + n]) n++; return n; };
    // A cell merges into the one above only when no rule separates them and it has nothing of its own.
    const hasText = (r: number, c: number) => (owned.get(`${r}:${c}`) ?? []).length > 0;
    const mergedDown = (r: number, c: number) => r > 0 && !g.h[r][c] && !hasText(r, c);
    for (let r = 0; r < rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < cols;) {
        const span = spanRight(r, c);
        const cellLines = Array.from({ length: span }, (_, k) => owned.get(`${r}:${c + k}`) ?? []).flat().sort((a, b) => a.y - b.y || a.x0 - b.x0);
        const fill = fillAt(r, c);
        const vMerge = mergedDown(r, c) ? "continue" : r + 1 < rows && mergedDown(r + 1, c) ? "restart" : undefined;
        const blocksIn: Block[] = cellLines.length ? cellParagraphs(cellLines, g.xs[c], g.xs[c + span]) : [];
        // Pictures in the cell, in order with its text.
        for (const pic of pictures) {
          const [pr, pc] = cellOf(pic.x + pic.w / 2, pic.y + pic.h / 2);
          if (pr !== r || pc < c || pc >= c + span) continue;
          const at = cellLines.filter((l) => l.y < pic.y + pic.h / 2).length;
          blocksIn.splice(Math.min(at, blocksIn.length), 0, pictureParagraph(pic, g.xs[c] + 5.4, g.xs[c + span] - 5.4));
        }
        if (!blocksIn.length) blocksIn.push({ type: "paragraph", props: {}, content: [] });
        cells.push({
          props: {
            width: { w: String(Math.round(grid.slice(c, c + span).reduce((a, b) => a + b, 0) * 20)), type: "dxa" },
            ...(span > 1 ? { span } : {}), ...(vMerge ? { vMerge } : {}),
            ...(fill ? { shading: { val: "clear", color: "auto", fill } } : {}),
          },
          blocks: vMerge === "continue" ? [{ type: "paragraph", props: {}, content: [] }] : blocksIn,
        });
        c += span;
      }
      out.rows.push({ props: {}, cells });
    }
    return out;
  }

  function cellParagraphs(lines: Line[], cellLeft: number, cellRight: number): Block[] {
    const out: Block[] = [];
    let cur: Line[] = [];
    // Word's cell margins: 5.4pt each side.
    const [cl, cr] = [cellLeft + 5.4, cellRight - 5.4];
    const flush = () => {
      if (!cur.length) return;
      const p = paragraphOf(cur);
      // Alignment within a cell is relative to the cell, not the page.
      delete p.props.align;
      delete p.props.indent;
      const room = cr - cl;
      if (cur.every((l) => l.x0 > cl + 6 && Math.abs((l.x0 + l.x1) / 2 - (cl + cr) / 2) < Math.max(3, room * 0.03))) p.props.align = "center";
      else if (cur.every((l) => l.x0 > cl + 6 && Math.abs(l.x1 - cr) < 3)) p.props.align = "right";
      out.push(p);
      cur = [];
    };
    for (const l of lines) {
      const prev = cur[cur.length - 1];
      if (prev && (l.y - prev.y > prev.size * 1.6 || Math.abs(l.size - prev.size) > 0.6)) flush();
      cur.push(l);
    }
    flush();
    return out;
  }

  if (pages.some((p) => p.texts.some((t) => Math.abs(t.size) > 0 && t.text.trim()) && p.pictures.some((pic) => {
    const ls = buildLines(p, p.texts);
    return ls.some((l) => l.y > pic.y && l.y < pic.y + pic.h && (l.x0 > pic.x + pic.w || l.x1 < pic.x));
  }))) warnings.push("Some pictures had text beside them; they were placed on their own line.");
  return { doc, warnings };
}

function base64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
