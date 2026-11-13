/**
 * Table editing, as on Word's Table Design and Table Layout tabs: merge and split cells, borders and
 * shading for the selected cells, cell alignment, table styles and their options (header row, banded
 * rows…), column widths (drag, distribute, fit the window), repeating header rows, sorting, and
 * converting between text and tables.
 *
 * The cells the selection covers form a rectangle of rows × grid columns, as in Word.
 */
import type { Block, CellProps, ParagraphProps, PrimeDocument, Style } from "@prime/pdm";
import type { Node, ResolvedPos } from "prosemirror-model";
import { type Command, type EditorState, Plugin, TextSelection } from "prosemirror-state";
import { storyFromProseMirror, storyToProseMirror } from "./convert.js";

const twips = (pt: number) => String(Math.round(pt * 20));

interface Cell { node: Node; start: number; span: number }
interface Grid { table: Node; pos: number; rows: { node: Node; cells: Cell[] }[] }

function tableAt($pos: ResolvedPos): { table: Node; pos: number; row: number; cell: number } | undefined {
  for (let d = $pos.depth; d > 2; d--) {
    if ($pos.node(d).type.name === "table_cell" && $pos.node(d - 2).type.name === "table") {
      return { table: $pos.node(d - 2), pos: $pos.before(d - 2), row: $pos.index(d - 2), cell: $pos.index(d - 1) };
    }
  }
  return undefined;
}

function gridOf(table: Node, pos: number): Grid {
  const rows: Grid["rows"] = [];
  table.forEach((row) => {
    const cells: Cell[] = [];
    let col = 0;
    row.forEach((cell) => {
      const span = (cell.attrs.props as CellProps).span ?? 1;
      cells.push({ node: cell, start: col, span });
      col += span;
    });
    rows.push({ node: row, cells });
  });
  return { table, pos, rows };
}

export interface CellRange { grid: Grid; r1: number; r2: number; c1: number; c2: number }

/** The selected cells: rows r1…r2 and grid columns c1…c2 of the table at the selection. */
export function selectedCells(state: EditorState): CellRange | undefined {
  const a = tableAt(state.selection.$from);
  const b = tableAt(state.selection.$to);
  if (!a) return undefined;
  const grid = gridOf(a.table, a.pos);
  const end = b && b.pos === a.pos ? b : a;
  const ca = grid.rows[a.row].cells[a.cell];
  const cb = grid.rows[end.row].cells[end.cell];
  return {
    grid, r1: Math.min(a.row, end.row), r2: Math.max(a.row, end.row),
    c1: Math.min(ca.start, cb.start), c2: Math.max(ca.start + ca.span, cb.start + cb.span) - 1,
  };
}

/** The table node rebuilt from rows of cells (and a grid, if it changed). */
function rebuild(state: EditorState, g: Grid, rows: { node: Node; cells: Node[] }[], grid?: number[]) {
  const s = state.schema;
  const table = s.nodes.table.create({ ...g.table.attrs, ...(grid ? { grid } : {}) },
    rows.map((r) => s.nodes.table_row.create(r.node.attrs, r.cells)));
  return state.tr.replaceWith(g.pos, g.pos + g.table.nodeSize, table);
}

const withProps = (cell: Node, props: CellProps) => cell.type.create({ ...cell.attrs, props }, cell.content, cell.marks);
const isEmpty = (cell: Node) => cell.childCount === 1 && cell.firstChild!.content.size === 0;

/** Every cell in the range (with its row), for formatting commands. */
function cellsIn(range: CellRange): { r: number; k: number; cell: Cell }[] {
  const out: { r: number; k: number; cell: Cell }[] = [];
  for (let r = range.r1; r <= range.r2; r++) {
    range.grid.rows[r].cells.forEach((cell, k) => {
      if (cell.start + cell.span - 1 >= range.c1 && cell.start <= range.c2) out.push({ r, k, cell });
    });
  }
  return out;
}

/** Applies `change` to the props of every selected cell (r, k: row and cell index). */
function updateCells(change: (props: CellProps, at: { r: number; cell: Cell; range: CellRange }) => CellProps): Command {
  return (state, dispatch) => {
    const range = selectedCells(state);
    if (!range) return false;
    if (!dispatch) return true;
    const picked = new Set(cellsIn(range).map((x) => `${x.r}:${x.k}`));
    const rows = range.grid.rows.map((row, r) => ({
      node: row.node,
      cells: row.cells.map((c, k) => (picked.has(`${r}:${k}`) ? withProps(c.node, change({ ...c.node.attrs.props }, { r, cell: c, range })) : c.node)),
    }));
    const sel = state.selection;
    const tr = rebuild(state, range.grid, rows);
    tr.setSelection(TextSelection.between(tr.doc.resolve(Math.min(sel.from, tr.doc.content.size)), tr.doc.resolve(Math.min(sel.to, tr.doc.content.size))));
    dispatch(tr);
    return true;
  };
}

// --- merge and split -----------------------------------------------------------------------------

/**
 * Merges the selected cells into one (their text kept, one paragraph after another). Across rows, the
 * merged cell spans them vertically (as Word does with w:vMerge).
 */
export const mergeCells: Command = (state, dispatch) => {
  const range = selectedCells(state);
  if (!range) return false;
  const { grid, r1, r2, c1, c2 } = range;
  // Each row must have cells that start at c1 and end at c2 (no cell sticking out of the rectangle).
  for (let r = r1; r <= r2; r++) {
    const inside = grid.rows[r].cells.filter((c) => c.start + c.span - 1 >= c1 && c.start <= c2);
    if (!inside.length || inside[0].start !== c1 || inside[inside.length - 1].start + inside[inside.length - 1].span - 1 !== c2) return false;
  }
  if (r1 === r2 && grid.rows[r1].cells.filter((c) => c.start >= c1 && c.start <= c2).length < 2) return false;
  if (!dispatch) return true;
  // All the text goes into the first cell (empty cells add nothing).
  const content: Node[] = [];
  for (let r = r1; r <= r2; r++) for (const c of grid.rows[r].cells) {
    if (c.start >= c1 && c.start <= c2 && !isEmpty(c.node)) c.node.forEach((b) => content.push(b));
  }
  if (!content.length) content.push(state.schema.nodes.paragraph.create());
  const width = (state.doc.nodeAt(grid.pos)!.attrs.grid as number[]).slice(c1, c2 + 1).reduce((a, b) => a + b, 0);
  const rows = grid.rows.map((row, r) => {
    if (r < r1 || r > r2) return { node: row.node, cells: row.cells.map((c) => c.node) };
    const out: Node[] = [];
    for (const c of row.cells) {
      if (c.start < c1 || c.start > c2) { out.push(c.node); continue; }
      if (c.start !== c1) continue;
      const { vMerge: _v, ...props } = c.node.attrs.props as CellProps;
      const merged: CellProps = { ...props, span: c2 - c1 + 1, width: { w: twips(width), type: "dxa" } };
      if (merged.span === 1) delete merged.span;
      if (r2 > r1) merged.vMerge = r === r1 ? "restart" : "continue";
      out.push(c.node.type.create({ ...c.node.attrs, props: merged }, r === r1 ? content : [state.schema.nodes.paragraph.create()]));
    }
    return { node: row.node, cells: out };
  });
  const tr = rebuild(state, grid, rows);
  // The cursor in the merged cell.
  const cellPos = (() => {
    const t = tr.doc.nodeAt(grid.pos)!;
    let p = grid.pos + 1;
    for (let r = 0; r < r1; r++) p += t.child(r).nodeSize;
    p += 1;
    const row = t.child(r1);
    for (let k = 0; k < row.childCount; k++) {
      if (gridOf(t, grid.pos).rows[r1].cells[k].start === c1) return p;
      p += row.child(k).nodeSize;
    }
    return p;
  })();
  tr.setSelection(TextSelection.near(tr.doc.resolve(cellPos + 2)));
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * Splits the cell at the cursor: a merged cell back into its parts; an ordinary cell into two
 * columns (its grid column halved; the cells above and below it span both halves).
 */
export const splitCell: Command = (state, dispatch) => {
  const at = tableAt(state.selection.$from);
  if (!at) return false;
  const g = gridOf(at.table, at.pos);
  const target = g.rows[at.row].cells[at.cell];
  const props = target.node.attrs.props as CellProps;
  if (!dispatch) return true;
  const grid = [...(at.table.attrs.grid as number[])];
  const cellWidth = (start: number, span: number) => ({ w: twips(grid.slice(start, start + span).reduce((a, b) => a + b, 0)), type: "dxa" });
  if ((props.span ?? 1) > 1 || props.vMerge === "restart") {
    // Undo the merge: one cell per grid column, and the rows below stop continuing it.
    const rows = g.rows.map((row, r) => ({
      node: row.node,
      cells: row.cells.flatMap((c) => {
        const p = c.node.attrs.props as CellProps;
        const same = c.start === target.start && c.span === target.span;
        const inMerge = same && (r === at.row || (r > at.row && p.vMerge === "continue" && g.rows.slice(at.row + 1, r).every((x) => x.cells.some((y) => y.start === c.start && (y.node.attrs.props as CellProps).vMerge === "continue"))));
        if (!inMerge) return [c.node];
        const { span: _s, vMerge: _v, ...rest } = p;
        return Array.from({ length: c.span }, (_, i) => c.node.type.create(
          { ...c.node.attrs, props: { ...rest, width: cellWidth(c.start + i, 1) } },
          i === 0 && r === at.row ? c.node.content : [state.schema.nodes.paragraph.create()]));
      }),
    }));
    dispatch(keepCursor(state, rebuild(state, g, rows)).scrollIntoView());
    return true;
  }
  // An ordinary cell: its grid column becomes two.
  const col = target.start;
  const half = grid[col] / 2;
  grid.splice(col, 1, half, half);
  const rows = g.rows.map((row, r) => ({
    node: row.node,
    cells: row.cells.flatMap((c) => {
      const p = c.node.attrs.props as CellProps;
      if (r === at.row && c === target) {
        return [withProps(c.node, { ...p, width: { w: twips(half), type: "dxa" } }), c.node.type.create({ props: { ...p, width: { w: twips(half), type: "dxa" } } }, state.schema.nodes.paragraph.create())];
      }
      // Other rows: the cell over that column now spans one more.
      if (c.start <= col && c.start + c.span - 1 >= col) return [withProps(c.node, { ...p, span: c.span + 1 })];
      return [c.node];
    }),
  }));
  dispatch(keepCursor(state, rebuild(state, g, rows, grid)).scrollIntoView());
  return true;
};

/** The cursor back where it was (the table was replaced; the split cell starts where it did). */
function keepCursor(state: EditorState, tr: ReturnType<typeof rebuild>) {
  return tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(state.selection.from, tr.doc.content.size))));
}

// --- borders, shading, alignment -----------------------------------------------------------------

export type BorderChoice = "all" | "outside" | "inside" | "none" | "top" | "bottom" | "left" | "right" | "insideH" | "insideV";
export interface Pen { style?: string; width?: number; color?: string }

const side = (name: string, pen: Pen | null) => (pen
  ? `<w:${name} w:val="${pen.style ?? "single"}" w:sz="${Math.round((pen.width ?? 0.5) * 8)}" w:space="0" w:color="${pen.color ?? "auto"}"/>`
  : `<w:${name} w:val="nil"/>`);

/** The cell's own borders (w:tcBorders), side by side: undefined keeps a side as it is. */
function setTcBorders(props: CellProps, sides: Partial<Record<"top" | "left" | "bottom" | "right", Pen | null>>): CellProps {
  const old = props.extra?.find((x) => x.name === "w:tcBorders")?.xml ?? "";
  const keep = (name: string) => new RegExp(`<w:${name}\\b[^>]*/>`).exec(old)?.[0] ?? "";
  const xml = (["top", "left", "bottom", "right"] as const)
    .map((n) => (sides[n] !== undefined ? side(n, sides[n]!) : keep(n))).join("");
  const others = (props.extra ?? []).filter((x) => x.name !== "w:tcBorders");
  return { ...props, extra: xml ? [...others, { name: "w:tcBorders", xml: `<w:tcBorders>${xml}</w:tcBorders>` }] : others };
}

/**
 * Borders on the selected cells, as Word's Borders menu: all, outside, inside, none, one side, or the
 * inside horizontal/vertical lines — relative to the selected rectangle.
 */
export function setCellBorders(choice: BorderChoice, pen: Pen = {}): Command {
  return updateCells((props, { r, cell, range }) => {
    const top = r === range.r1;
    const bottom = r === range.r2;
    const left = cell.start === range.c1;
    const right = cell.start + cell.span - 1 === range.c2;
    const on = (b: boolean) => (b ? pen : undefined);
    const sides: Partial<Record<"top" | "left" | "bottom" | "right", Pen | null>> =
      choice === "all" ? { top: pen, bottom: pen, left: pen, right: pen }
      : choice === "none" ? { top: null, bottom: null, left: null, right: null }
      : choice === "outside" ? { top: on(top), bottom: on(bottom), left: on(left), right: on(right) }
      : choice === "inside" ? { top: on(!top), bottom: on(!bottom), left: on(!left), right: on(!right) }
      : choice === "insideH" ? { top: on(!top), bottom: on(!bottom) }
      : choice === "insideV" ? { left: on(!left), right: on(!right) }
      : { [choice]: (choice === "top" && top) || (choice === "bottom" && bottom) || (choice === "left" && left) || (choice === "right" && right) ? pen : undefined };
    return setTcBorders(props, sides);
  });
}

/** Shading (a fill colour, RRGGBB) on the selected cells; undefined removes it. */
export function setCellShading(fill: string | undefined): Command {
  return updateCells((props) => {
    const { shading: _s, ...rest } = props;
    return fill ? { ...rest, shading: { val: "clear", color: "auto", fill: fill.toUpperCase() } } : rest;
  });
}

/** Vertical alignment in the selected cells (top, center, bottom), and optionally the text's horizontal alignment. */
export function setCellAlign(vertical: "top" | "center" | "bottom", horizontal?: "left" | "center" | "right"): Command {
  return (state, dispatch) => {
    const range = selectedCells(state);
    if (!range) return false;
    if (!dispatch) return true;
    const picked = new Set(cellsIn(range).map((x) => `${x.r}:${x.k}`));
    const s = state.schema;
    const rows = range.grid.rows.map((row, r) => ({
      node: row.node,
      cells: row.cells.map((c, k) => {
        if (!picked.has(`${r}:${k}`)) return c.node;
        const { vAlign: _v, ...rest } = c.node.attrs.props as CellProps;
        const props = vertical === "top" ? rest : { ...rest, vAlign: vertical };
        const content: Node[] = [];
        c.node.forEach((b) => {
          if (horizontal && b.type.name === "paragraph") {
            const p = { ...(b.attrs.props as ParagraphProps) };
            if (horizontal === "left") delete p.align; else p.align = horizontal;
            content.push(b.type.create({ ...b.attrs, props: p }, b.content, b.marks));
          } else content.push(b);
        });
        return s.nodes.table_cell.create({ ...c.node.attrs, props }, content);
      }),
    }));
    const sel = state.selection;
    const tr = rebuild(state, range.grid, rows);
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(sel.from, tr.doc.content.size))));
    dispatch(tr);
    return true;
  };
}

// --- table styles ------------------------------------------------------------------------------

const B = (name: string, color: string, sz = 4, val = "single") => `<w:${name} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
const all = (color: string, sz = 4) => ["top", "left", "bottom", "right", "insideH", "insideV"].map((n) => B(n, color, sz)).join("");
const part = (type: string, body: { run?: string; fill?: string; borders?: string }) =>
  `<w:tblStylePr w:type="${type}">${body.run ? `<w:rPr>${body.run}</w:rPr>` : ""}<w:tblPr/><w:tcPr>${body.borders ? `<w:tcBorders>${body.borders}</w:tcBorders>` : ""}${body.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${body.fill}"/>` : ""}</w:tcPr></w:tblStylePr>`;

/** Table styles offered in the gallery: Word's names and looks, plus an APA style for academic work. */
export const TABLE_STYLES: { id: string; name: string; borders: string; parts: string[] }[] = [
  { id: "TableGrid", name: "Table Grid", borders: all("auto"), parts: [] },
  { id: "PlainTable1", name: "Plain Table 1", borders: all("BFBFBF"), parts: [part("firstRow", { run: "<w:b/>" }), part("lastRow", { run: "<w:b/>", borders: B("top", "BFBFBF", 4, "double") }), part("firstCol", { run: "<w:b/>" }), part("band1Horz", { fill: "F2F2F2" })] },
  { id: "PlainTable2", name: "Plain Table 2", borders: B("top", "7F7F7F") + B("bottom", "7F7F7F"), parts: [part("firstRow", { run: "<w:b/>", borders: B("bottom", "7F7F7F") }), part("lastRow", { run: "<w:b/>", borders: B("top", "7F7F7F") }), part("firstCol", { run: "<w:b/>" }), part("band1Horz", { borders: B("top", "7F7F7F") + B("bottom", "7F7F7F") })] },
  ...([["1", "156082", "C1E4F5", "45B0E1"], ["2", "E97132", "FAE2D5", "F1A983"], ["6", "4EA72E", "D9F2D0", "8DD873"]] as const).map(([n, dark, light, line]) => ({
    id: `GridTable4-Accent${n}`, name: `Grid Table 4 – Accent ${n}`, borders: all(line),
    parts: [part("firstRow", { run: '<w:b/><w:color w:val="FFFFFF"/>', fill: dark, borders: ["top", "left", "bottom", "right"].map((s) => B(s, dark)).join("") + '<w:insideH w:val="nil"/><w:insideV w:val="nil"/>' }),
      part("lastRow", { run: "<w:b/>", borders: B("top", dark, 4, "double") }), part("firstCol", { run: "<w:b/>" }), part("lastCol", { run: "<w:b/>" }), part("band1Horz", { fill: light })],
  })),
  { id: "ListTable3", name: "List Table 3", borders: B("top", "000000") + B("left", "000000") + B("bottom", "000000") + B("right", "000000"), parts: [part("firstRow", { run: '<w:b/><w:color w:val="FFFFFF"/>', fill: "000000" }), part("lastRow", { run: "<w:b/>", borders: B("top", "000000", 4, "double") }), part("firstCol", { run: "<w:b/>" }), part("band1Horz", { borders: B("top", "000000") + B("bottom", "000000") })] },
  { id: "APATable", name: "APA (academic)", borders: B("top", "000000", 8) + B("bottom", "000000", 8), parts: [part("firstRow", { borders: B("bottom", "000000", 8) }), part("lastRow", { borders: B("bottom", "000000", 8) })] },
];

function ensureTableStyle(doc: PrimeDocument, id: string) {
  const def = TABLE_STYLES.find((s) => s.id === id);
  if (!def || doc.styles.styles.some((s) => s.id === id)) return;
  const style: Style = {
    id, type: "table", name: def.name, ...(doc.styles.styles.some((s) => s.id === "TableNormal") ? { basedOn: "TableNormal" } : {}),
    paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } },
    table: { extra: [{ name: "w:tblBorders", xml: `<w:tblBorders>${def.borders}</w:tblBorders>` }] },
    extra: def.parts.map((xml) => ({ name: "w:tblStylePr", xml })),
  };
  doc.styles.styles.push(style);
}

export interface TableLook { firstRow: boolean; lastRow: boolean; firstColumn: boolean; lastColumn: boolean; bandedRows: boolean; bandedColumns: boolean }

const lookXml = (l: TableLook) => `<w:tblLook w:val="${(((l.firstRow ? 0x20 : 0) | (l.lastRow ? 0x40 : 0) | (l.firstColumn ? 0x80 : 0) | (l.lastColumn ? 0x100 : 0) | (l.bandedRows ? 0 : 0x200) | (l.bandedColumns ? 0 : 0x400)).toString(16).toUpperCase()).padStart(4, "0")}" w:firstRow="${+l.firstRow}" w:lastRow="${+l.lastRow}" w:firstColumn="${+l.firstColumn}" w:lastColumn="${+l.lastColumn}" w:noHBand="${+!l.bandedRows}" w:noVBand="${+!l.bandedColumns}"/>`;

/** The table's style options (Header Row, Total Row, Banded Rows…). */
export function tableLookAt(state: EditorState): TableLook | undefined {
  const at = tableAt(state.selection.$from);
  if (!at) return undefined;
  const xml = (at.table.attrs.props.extra ?? []).find((x: { name: string }) => x.name === "w:tblLook")?.xml ?? "";
  const flag = (n: string, def: boolean) => { const m = new RegExp(`w:${n}="(\\w+)"`).exec(xml)?.[1]; return m === undefined ? def : m === "1" || m === "true"; };
  return {
    firstRow: flag("firstRow", true), lastRow: flag("lastRow", false), firstColumn: flag("firstColumn", true),
    lastColumn: flag("lastColumn", false), bandedRows: !flag("noHBand", false), bandedColumns: !flag("noVBand", true),
  };
}

/** Sets the table's style (from the gallery) and, optionally, its style options. */
export function setTableStyle(doc: () => PrimeDocument, styleId: string | undefined, look?: TableLook): Command {
  return (state, dispatch) => {
    const at = tableAt(state.selection.$from);
    if (!at) return false;
    if (!dispatch) return true;
    if (styleId) ensureTableStyle(doc(), styleId);
    const props = { ...at.table.attrs.props };
    if (styleId !== undefined) props.style = styleId || undefined;
    if (look) props.extra = [...(props.extra ?? []).filter((x: { name: string }) => x.name !== "w:tblLook"), { name: "w:tblLook", xml: lookXml(look) }];
    // A style's borders show only where the table does not set its own.
    if (styleId) props.extra = (props.extra ?? []).filter((x: { name: string }) => x.name !== "w:tblBorders");
    dispatch(state.tr.setNodeMarkup(at.pos, undefined, { ...at.table.attrs, props }));
    return true;
  };
}

// --- column widths ------------------------------------------------------------------------------

/** New column widths (pt) for the table at the cursor; the cells' own widths follow. */
export function setColumnWidths(tablePos: number, grid: number[]): Command {
  return (state, dispatch) => {
    const table = state.doc.nodeAt(tablePos);
    if (table?.type.name !== "table") return false;
    if (!dispatch) return true;
    const g = gridOf(table, tablePos);
    const rows = g.rows.map((row) => ({
      node: row.node,
      cells: row.cells.map((c) => withProps(c.node, { ...(c.node.attrs.props as CellProps), width: { w: twips(grid.slice(c.start, c.start + c.span).reduce((a, b) => a + b, 0)), type: "dxa" } })),
    }));
    const sel = state.selection;
    const tr = rebuild(state, g, rows, grid);
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(sel.from, tr.doc.content.size))));
    dispatch(tr);
    return true;
  };
}

/** Distribute Columns: all columns of the table as wide as each other. */
export const distributeColumns: Command = (state, dispatch) => {
  const at = tableAt(state.selection.$from);
  if (!at) return false;
  const grid = at.table.attrs.grid as number[];
  const total = grid.reduce((a, b) => a + b, 0);
  return setColumnWidths(at.pos, grid.map(() => total / grid.length))(state, dispatch);
};

/** AutoFit Window: the table as wide as the text column, columns keeping their proportions. */
export function autofitWindow(width: number): Command {
  return (state, dispatch) => {
    const at = tableAt(state.selection.$from);
    if (!at) return false;
    const grid = at.table.attrs.grid as number[];
    const total = grid.reduce((a, b) => a + b, 0) || 1;
    return setColumnWidths(at.pos, grid.map((w) => (w / total) * width))(state, dispatch);
  };
}

// --- rows ---------------------------------------------------------------------------------------

/** Repeat as Header Row: the selected rows (from the top) repeat at the top of every page the table runs onto. */
export const toggleHeaderRows: Command = (state, dispatch) => {
  const range = selectedCells(state);
  if (!range) return false;
  if (!dispatch) return true;
  const on = !range.grid.rows[range.r1].node.attrs.props.header;
  const rows = range.grid.rows.map((row, r) => {
    const props = { ...row.node.attrs.props };
    // Header rows are the rows at the top: rows 0…r2 when turning on.
    if (on ? r <= range.r2 : r >= range.r1 && r <= range.r2) {
      if (on) props.header = true; else delete props.header;
    }
    return { node: row.node.type.create({ ...row.node.attrs, props }, row.node.content), cells: row.cells.map((c) => c.node) };
  });
  const sel = state.selection;
  const tr = rebuild(state, range.grid, rows);
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(sel.from, tr.doc.content.size))));
  dispatch(tr);
  return true;
};

/** Whether the row at the cursor repeats as a header row. */
export const isHeaderRow = (state: EditorState) => {
  const at = tableAt(state.selection.$from);
  return !!at && !!at.table.child(at.row).attrs.props.header;
};

// --- sorting ------------------------------------------------------------------------------------

export interface SortOptions { column: number; descending?: boolean; type?: "text" | "number" | "date"; header?: boolean }

/** Sorts the table's rows by a column's text (header rows, and merged rows, stay where they are). */
export function sortTable(o: SortOptions): Command {
  return (state, dispatch) => {
    const at = tableAt(state.selection.$from);
    if (!at) return false;
    const g = gridOf(at.table, at.pos);
    const fixedTop = o.header ? Math.max(1, g.rows.findIndex((r) => !r.node.attrs.props.header)) : 0;
    const body = g.rows.slice(fixedTop);
    if (body.some((r) => r.cells.some((c) => c.node.attrs.props.vMerge))) return false;
    if (!dispatch) return true;
    const key = (row: Grid["rows"][number]) => row.cells.find((c) => c.start <= o.column && c.start + c.span > o.column)?.node.textContent.trim() ?? "";
    const value = (t: string) => {
      if (o.type === "number") { const n = Number(t.replace(/[^0-9.\-]/g, "")); return Number.isFinite(n) && t ? n : -Infinity; }
      if (o.type === "date") { const d = Date.parse(t); return Number.isNaN(d) ? -Infinity : d; }
      return t.toLocaleLowerCase();
    };
    const sorted = [...body].sort((a, b) => {
      const x = value(key(a));
      const y = value(key(b));
      const c = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
      return o.descending ? -c : c;
    });
    const rows = [...g.rows.slice(0, fixedTop), ...sorted].map((r) => ({ node: r.node, cells: r.cells.map((c) => c.node) }));
    dispatch(rebuild(state, g, rows).scrollIntoView());
    return true;
  };
}

// --- text and tables ----------------------------------------------------------------------------

/** Convert Text to Table: the selected paragraphs, split at tabs (or commas), become rows. */
export function textToTable(doc: () => PrimeDocument, separator: "tab" | "comma" = "tab"): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const $from = state.doc.resolve(from);
    const $to = state.doc.resolve(to);
    if ($from.depth !== 1 || $to.depth !== 1) return false;
    const start = $from.before(1);
    const end = $to.after(1);
    const paras: Node[] = [];
    state.doc.nodesBetween(start, end, (n) => { if (n.type.name === "paragraph") paras.push(n); return false; });
    if (!paras.length) return false;
    if (!dispatch) return true;
    const s = state.schema;
    const rows = paras.map((p) => {
      // Each piece keeps its formatting: the paragraph is cut at its tabs (or commas).
      const cells: Node[][] = [[]];
      p.forEach((n) => {
        if (separator === "tab" && n.type.name === "tab") { cells.push([]); return; }
        if (separator === "comma" && n.isText) {
          const parts = n.text!.split(",");
          parts.forEach((t, i) => { if (i > 0) cells.push([]); const v = i > 0 ? t.replace(/^ /, "") : t; if (v) cells[cells.length - 1].push(s.text(v, n.marks)); });
          return;
        }
        cells[cells.length - 1].push(n);
      });
      return { props: p.attrs.props as ParagraphProps, cells };
    });
    const cols = Math.max(...rows.map((r) => r.cells.length));
    const section = doc().sections[0]?.props ?? {};
    const width = (section.pageWidth ?? 612) - (section.margins?.left ?? 72) - (section.margins?.right ?? 72);
    const col = width / cols;
    const table = s.nodes.table.create(
      { props: { style: "TableGrid", width: { w: "0", type: "auto" } }, grid: Array(cols).fill(col) },
      rows.map((r) => s.nodes.table_row.create({ props: {} }, Array.from({ length: cols }, (_, i) => s.nodes.table_cell.create(
        { props: { width: { w: twips(col), type: "dxa" } } },
        s.nodes.paragraph.create({ props: (({ list: _x, ...rest }) => rest)(r.props) }, r.cells[i] ?? []))))),
    );
    ensureTableStyle(doc(), "TableGrid");
    const tr = state.tr.replaceWith(start, end, table);
    tr.setSelection(TextSelection.near(tr.doc.resolve(start + 4)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Convert Table to Text: each row a paragraph, cells separated by tabs (or commas). */
export function tableToText(separator: "tab" | "comma" = "tab"): Command {
  return (state, dispatch) => {
    const at = tableAt(state.selection.$from);
    if (!at) return false;
    if (!dispatch) return true;
    const s = state.schema;
    const paras: Node[] = [];
    at.table.forEach((row) => {
      const content: Node[] = [];
      let first: ParagraphProps | undefined;
      row.forEach((cell, _o, k) => {
        if ((cell.attrs.props as CellProps).vMerge === "continue") return;
        if (k > 0) content.push(separator === "tab" ? s.nodes.tab.create() : s.text(", "));
        // A cell's paragraphs join with spaces (a row becomes one paragraph).
        const blocks: Block[] = storyFromProseMirror(cell);
        blocks.forEach((b, i) => {
          if (b.type !== "paragraph") return;
          first ??= b.props;
          if (i > 0) content.push(s.text(" "));
          storyToProseMirror([b], s).firstChild?.forEach((n) => content.push(n));
        });
      });
      paras.push(s.nodes.paragraph.create({ props: first ?? {} }, content));
    });
    const tr = state.tr.replaceWith(at.pos, at.pos + at.table.nodeSize, paras);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// --- resizing columns with the mouse ------------------------------------------------------------

/**
 * Column borders can be dragged: near a cell's right edge the pointer becomes a resize arrow; dragging
 * moves that border (the next column gives or takes the difference; the last border changes the
 * table's width). A guide line shows where it will go.
 */
export function columnResizing(): Plugin {
  return new Plugin({
    view(view) {
      let hover: { tablePos: number; col: number; edge: number } | undefined;
      let drag: { tablePos: number; col: number; startX: number; zoom: number; guide: HTMLElement } | undefined;
      const zoomOf = () => {
        const w = parseFloat(getComputedStyle(view.dom).width);
        return w ? view.dom.getBoundingClientRect().width / w : 1;
      };
      const find = (e: MouseEvent) => {
        const td = (e.target as HTMLElement).closest?.("td");
        if (!td || !view.dom.contains(td)) return undefined;
        const r = td.getBoundingClientRect();
        // Within 4px of the cell's right edge (or the previous cell's, at the left edge).
        let cell: Element | null = td;
        if (Math.abs(e.clientX - r.left) <= 4) cell = td.previousElementSibling;
        else if (Math.abs(e.clientX - r.right) > 4) return undefined;
        if (!cell) return undefined;
        const pos = view.posAtDOM(cell, 0);
        const at = tableAt(view.state.doc.resolve(pos));
        if (!at) return undefined;
        const g = gridOf(at.table, at.pos);
        const c = g.rows[at.row].cells[at.cell];
        return { tablePos: at.pos, col: c.start + c.span - 1, edge: (cell as HTMLElement).getBoundingClientRect().right };
      };
      const move = (e: MouseEvent) => {
        if (drag) {
          drag.guide.style.left = `${e.clientX}px`;
          return;
        }
        hover = find(e);
        view.dom.classList.toggle("pw-col-resize", !!hover);
      };
      const down = (e: MouseEvent) => {
        if (!hover || e.button !== 0) return;
        e.preventDefault();
        const guide = document.createElement("div");
        guide.className = "pw-col-guide";
        guide.style.left = `${e.clientX}px`;
        document.body.appendChild(guide);
        drag = { tablePos: hover.tablePos, col: hover.col, startX: e.clientX, zoom: zoomOf(), guide };
      };
      const up = (e: MouseEvent) => {
        if (!drag) return;
        const d = drag;
        drag = undefined;
        d.guide.remove();
        const delta = ((e.clientX - d.startX) / d.zoom) * (72 / 96);
        const table = view.state.doc.nodeAt(d.tablePos);
        if (!table || Math.abs(delta) < 0.5) return;
        const grid = [...(table.attrs.grid as number[])];
        const min = 12;
        if (d.col + 1 < grid.length) {
          // The border between two columns: one grows, the other shrinks.
          const shift = Math.max(min - grid[d.col], Math.min(grid[d.col + 1] - min, delta));
          grid[d.col] += shift;
          grid[d.col + 1] -= shift;
        } else grid[d.col] = Math.max(min, grid[d.col] + delta);
        setColumnWidths(d.tablePos, grid)(view.state, view.dispatch);
      };
      view.dom.addEventListener("mousedown", down, true);
      window.addEventListener("mousemove", move);
      window.addEventListener("mouseup", up);
      return {
        destroy() {
          view.dom.removeEventListener("mousedown", down, true);
          window.removeEventListener("mousemove", move);
          window.removeEventListener("mouseup", up);
          drag?.guide.remove();
        },
      };
    },
  });
}
