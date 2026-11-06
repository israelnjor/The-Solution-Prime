/**
 * How each table cell looks, as Word resolves it: borders side by side, shading, and the text
 * formatting a table style gives parts of the table (a bold or white header row, banded rows, a
 * total row, first/last column) — switched on and off by the table's w:tblLook.
 *
 * Precedence, lowest first: the table style (and the styles it is based on), the table's own
 * borders, the style's conditional formats (whole table, banded columns, banded rows, first/last
 * column, header/total row, corner cells), then the cell's own formatting.
 */
import type { PrimeDocument, RunProps, Style, Table } from "@prime/pdm";

export interface Border { style: string; width: number; color?: string }
export type Sides = { top?: Border; bottom?: Border; left?: Border; right?: Border };
export interface CellLook {
  /** Sides with a visible border. */
  borders: Sides;
  fill?: string;
  /** Text formatting from the table style's conditional formats (below paragraph styles in precedence). */
  run?: RunProps;
}

type Side = "top" | "bottom" | "left" | "right" | "insideH" | "insideV";
/** An explicit "no border" overrides what is below it, so it is kept until the end. */
type Spec = Partial<Record<Side, Border | null>>;

const ORDER = ["wholeTable", "band1Vert", "band2Vert", "band1Horz", "band2Horz", "firstCol", "lastCol", "firstRow", "lastRow", "neCell", "nwCell", "seCell", "swCell"];

/** Borders in a w:tblBorders / w:tcBorders element (start/end read as left/right). */
export function parseBorders(xml: string | undefined): Spec {
  const out: Spec = {};
  if (!xml) return out;
  for (const m of xml.matchAll(/<w:(top|bottom|left|right|start|end|insideH|insideV)\b([^>]*?)\/?>/g)) {
    const side = (m[1] === "start" ? "left" : m[1] === "end" ? "right" : m[1]) as Side;
    const a = m[2];
    const val = /w:val="([^"]+)"/.exec(a)?.[1] ?? "single";
    if (val === "none" || val === "nil") { out[side] = null; continue; }
    const sz = Number(/w:sz="(\d+)"/.exec(a)?.[1] ?? 4);
    const color = /w:color="([0-9A-Fa-f]{6})"/.exec(a)?.[1]?.toUpperCase();
    out[side] = { style: val, width: Math.max(0.25, sz / 8), ...(color ? { color } : {}) };
  }
  return out;
}

const fillOf = (xml: string | undefined) => {
  const f = xml && /<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/.exec(xml)?.[1]?.toUpperCase();
  return f && f !== "AUTO" ? f : undefined;
};

/** The run formatting in a conditional format's w:rPr that matters for drawing. */
function parseRun(xml: string | undefined): RunProps | undefined {
  if (!xml) return undefined;
  const on = (tag: string) => {
    const m = new RegExp(`<w:${tag}(?:\\s[^>]*)?/>`).exec(xml);
    if (!m) return undefined;
    const v = /w:val="([^"]+)"/.exec(m[0])?.[1];
    return !(v === "0" || v === "false" || v === "off");
  };
  const r: RunProps = {};
  const b = on("b");
  if (b !== undefined) r.bold = b;
  const i = on("i");
  if (i !== undefined) r.italic = i;
  const caps = on("caps");
  if (caps !== undefined) r.caps = caps;
  const color = /<w:color\b[^>]*w:val="([0-9A-Fa-f]{6}|auto)"/.exec(xml)?.[1]?.toUpperCase();
  if (color) r.color = color;
  const sz = /<w:sz\b[^>]*w:val="(\d+)"/.exec(xml)?.[1];
  if (sz) r.size = Number(sz) / 2;
  return Object.keys(r).length ? r : undefined;
}

interface Conditional { borders: Spec; fill?: string; run?: RunProps }

/** One table style's parts: its own borders and cell shading, and its conditional formats by type. */
function styleParts(s: Style) {
  const xml = (s.extra ?? []).map((x) => x.xml).join("");
  const own: Conditional = {
    borders: parseBorders(s.table?.extra?.find((x) => x.name === "w:tblBorders")?.xml),
    fill: fillOf((s.extra ?? []).find((x) => x.name === "w:tcPr")?.xml),
  };
  const conditional = new Map<string, Conditional>();
  for (const m of xml.matchAll(/<w:tblStylePr\b[^>]*w:type="([^"]+)"[^>]*>([\s\S]*?)<\/w:tblStylePr>/g)) {
    const body = m[2];
    const tcPr = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/.exec(body)?.[1];
    conditional.set(m[1], {
      borders: { ...parseBorders(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/.exec(body)?.[0]), ...parseBorders(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/.exec(tcPr ?? "")?.[0]) },
      fill: fillOf(tcPr),
      run: parseRun(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(body)?.[0]),
    });
  }
  const band = (name: string) => Number(new RegExp(`<w:${name}\\b[^>]*w:val="(\\d+)"`).exec(xml)?.[1] ?? 0) || undefined;
  return { own, conditional, rowBand: band("tblStyleRowBandSize"), colBand: band("tblStyleColBandSize") };
}

/** Which conditional formats the table switches on (w:tblLook; Word's default when absent). */
function looks(t: Table) {
  const xml = t.props.extra?.find((x) => x.name === "w:tblLook")?.xml ?? "";
  const hex = /w:val="([0-9A-Fa-f]{4})"/.exec(xml)?.[1];
  const bits = hex ? parseInt(hex, 16) : 0x04a0;
  const flag = (name: string, bit: number, invert = false) => {
    const m = new RegExp(`w:${name}="(\\w+)"`).exec(xml)?.[1];
    const v = m !== undefined ? m === "1" || m === "true" || m === "on" : !!(bits & bit);
    return invert ? !v : v;
  };
  return {
    firstRow: flag("firstRow", 0x0020), lastRow: flag("lastRow", 0x0040), firstCol: flag("firstColumn", 0x0080),
    lastCol: flag("lastColumn", 0x0100), hBand: flag("noHBand", 0x0200, true), vBand: flag("noVBand", 0x0400, true),
  };
}

const merge = (a: Spec, b: Spec): Spec => ({ ...a, ...b });

/** Every cell's look (rows × cells, as in the table). */
export function cellLooks(doc: PrimeDocument, t: Table, tableStyle: string | undefined): CellLook[][] {
  // The style chain, base first.
  const chain: Style[] = [];
  const seen = new Set<string>();
  for (let id = tableStyle; id && !seen.has(id);) {
    seen.add(id);
    const s = doc.styles.styles.find((x) => x.id === id);
    if (!s) break;
    chain.unshift(s);
    id = s.basedOn;
  }
  const parts = chain.map(styleParts);
  let base: Spec = {};
  let baseFill: string | undefined;
  for (const p of parts) { base = merge(base, p.own.borders); baseFill = p.own.fill ?? baseFill; }
  base = merge(base, parseBorders(t.props.extra?.find((x) => x.name === "w:tblBorders")?.xml));
  const conditional = (type: string): Conditional | undefined => {
    let out: Conditional | undefined;
    for (const p of parts) {
      const c = p.conditional.get(type);
      if (!c) continue;
      out = { borders: merge(out?.borders ?? {}, c.borders), fill: c.fill ?? out?.fill, run: c.run || out?.run ? { ...out?.run, ...c.run } : undefined };
    }
    return out;
  };
  const rowBand = parts.map((p) => p.rowBand).filter(Boolean).pop() ?? 1;
  const colBand = parts.map((p) => p.colBand).filter(Boolean).pop() ?? 1;
  const on = looks(t);
  const columns = t.grid.length;
  const rows = t.rows.length;
  // Header rows: the first row, and any rows marked to repeat directly after it.
  let headers = 0;
  while (headers < rows && (headers === 0 || t.rows[headers].props.header)) headers++;
  const headerRows = on.firstRow ? headers : 0;

  return t.rows.map((row, r) => {
    let col = 0;
    return row.cells.map((cell) => {
      const start = col;
      const span = cell.props.span ?? 1;
      col += span;
      const end = col - 1;
      // The last row this cell reaches (a vertical merge continues below).
      let bottomRow = r;
      if (cell.props.vMerge === "restart") {
        while (bottomRow + 1 < rows && t.rows[bottomRow + 1].cells.some((c, k) => cellStart(t.rows[bottomRow + 1], k) === start && c.props.vMerge === "continue")) bottomRow++;
      }
      const types: string[] = ["wholeTable"];
      const bodyRow = r - headerRows;
      const bodyCol = start - (on.firstCol ? 1 : 0);
      if (on.vBand && bodyCol >= 0) types.push(Math.floor(bodyCol / colBand) % 2 === 0 ? "band1Vert" : "band2Vert");
      // Banding skips the header rows and the total row (as Word does).
      if (on.hBand && bodyRow >= 0 && !(on.lastRow && bottomRow === rows - 1)) types.push(Math.floor(bodyRow / rowBand) % 2 === 0 ? "band1Horz" : "band2Horz");
      if (on.firstCol && start === 0) types.push("firstCol");
      if (on.lastCol && end === columns - 1) types.push("lastCol");
      if (on.firstRow && r < headerRows) types.push("firstRow");
      if (on.lastRow && bottomRow === rows - 1) types.push("lastRow");
      if (on.firstRow && r === 0 && on.lastCol && end === columns - 1) types.push("neCell");
      if (on.firstRow && r === 0 && on.firstCol && start === 0) types.push("nwCell");
      if (on.lastRow && bottomRow === rows - 1 && on.lastCol && end === columns - 1) types.push("seCell");
      if (on.lastRow && bottomRow === rows - 1 && on.firstCol && start === 0) types.push("swCell");
      types.sort((a, b) => ORDER.indexOf(a) - ORDER.indexOf(b));

      // Outer edges take the table's outside borders; the rest its inside lines.
      const edge = { top: r === 0, bottom: bottomRow === rows - 1, left: start === 0, right: end === columns - 1 };
      const pick = (spec: Spec, side: "top" | "bottom" | "left" | "right"): Border | null | undefined =>
        edge[side] ? spec[side] : spec[side === "top" || side === "bottom" ? "insideH" : "insideV"];
      const sides: Record<"top" | "bottom" | "left" | "right", Border | null | undefined> = {
        top: pick(base, "top"), bottom: pick(base, "bottom"), left: pick(base, "left"), right: pick(base, "right"),
      };
      let fill = baseFill;
      let run: RunProps | undefined;
      for (const type of types) {
        const c = conditional(type);
        if (!c) continue;
        // A part's own outside sides apply where the cell is on that part's edge; its inside lines elsewhere.
        for (const side of ["top", "bottom", "left", "right"] as const) {
          const own = c.borders[side];
          const inside = c.borders[side === "top" || side === "bottom" ? "insideH" : "insideV"];
          // Row-shaped parts (header, total row, row bands) have their own top and bottom; column-shaped
          // ones their own left and right; corner cells all four; the whole table only its outside.
          const rowPart = type === "firstRow" || type === "lastRow" || type.endsWith("Horz");
          const colPart = type === "firstCol" || type === "lastCol" || type.endsWith("Vert");
          const horizontal = side === "top" || side === "bottom";
          const atEdge = type.endsWith("Cell") || (rowPart && horizontal) || (colPart && !horizontal) || edge[side];
          if (atEdge) { if (own !== undefined) sides[side] = own; }
          else if (inside !== undefined) sides[side] = inside;
        }
        if (c.fill) fill = c.fill;
        if (c.run) run = { ...run, ...c.run };
      }
      // The cell's own formatting.
      const direct = parseBorders(cell.props.extra?.find((x) => x.name === "w:tcBorders")?.xml);
      for (const side of ["top", "bottom", "left", "right"] as const) if (direct[side] !== undefined) sides[side] = direct[side];
      const own = cell.props.shading?.fill?.toUpperCase();
      if (own) fill = own === "AUTO" ? undefined : own;
      const borders: Sides = {};
      for (const side of ["top", "bottom", "left", "right"] as const) if (sides[side]) borders[side] = sides[side]!;
      return { borders, ...(fill ? { fill } : {}), ...(run ? { run } : {}) };
    });
  });
}

/** The grid column a row's cell `k` starts in. */
function cellStart(row: Table["rows"][number], k: number): number {
  let col = 0;
  for (let i = 0; i < k; i++) col += row.cells[i].props.span ?? 1;
  return col;
}

