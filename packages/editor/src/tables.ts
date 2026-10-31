/**
 * Table commands: insert a table, add/delete rows and columns, move between cells.
 * Column widths live in the table's grid; inserting a column scales the grid to keep the table's width.
 */
import type { PrimeDocument } from "@prime/pdm";
import type { Node, ResolvedPos } from "prosemirror-model";
import { type Command, type EditorState, TextSelection, type Transaction } from "prosemirror-state";

const BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';

const twips = (pt: number) => String(Math.round(pt * 20));

/** Inserts a rows × cols table with single borders, filling the text width. */
export function insertTable(doc: () => PrimeDocument, rows: number, cols: number): Command {
  return (state, dispatch) => {
    const { table, table_row, table_cell, paragraph } = state.schema.nodes;
    if (!dispatch) return true;
    const section = doc().sections[0]?.props ?? {};
    const width = (section.pageWidth ?? 612) - (section.margins?.left ?? 72) - (section.margins?.right ?? 72);
    const col = width / cols;
    const cell = () => table_cell.create({ props: { width: { w: twips(col), type: "dxa" } } }, paragraph.create());
    const node = table.create(
      { props: { width: { w: "0", type: "auto" }, extra: [{ name: "w:tblBorders", xml: BORDERS }] }, grid: Array(cols).fill(col) },
      Array.from({ length: rows }, () => table_row.create({ props: {} }, Array.from({ length: cols }, cell))),
    );
    const tr = state.tr.replaceSelectionWith(node);
    // Put the cursor in the first cell.
    const start = tr.mapping.map(state.selection.from) - node.nodeSize;
    const inside = Math.max(0, start) + 4;
    if (inside < tr.doc.content.size) tr.setSelection(TextSelection.near(tr.doc.resolve(inside)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

interface CellContext {
  table: Node;
  tablePos: number;
  rowIndex: number;
  cellIndex: number;
}

function cellContext($pos: ResolvedPos): CellContext | undefined {
  for (let d = $pos.depth; d > 2; d--) {
    if ($pos.node(d).type.name === "table_cell" && $pos.node(d - 2).type.name === "table") {
      return { table: $pos.node(d - 2), tablePos: $pos.before(d - 2), rowIndex: $pos.index(d - 2), cellIndex: $pos.index(d - 1) };
    }
  }
  return undefined;
}

export const inTable = (state: EditorState) => !!cellContext(state.selection.$from);

function rebuild(state: EditorState, ctx: CellContext, rows: Node[], grid: number[]): Transaction {
  const node = state.schema.nodes.table.create({ ...ctx.table.attrs, grid }, rows);
  return state.tr.replaceWith(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, node);
}

const emptyCell = (state: EditorState, like: Node) => {
  const { vMerge: _v, ...props } = like.attrs.props;
  return state.schema.nodes.table_cell.create({ props }, state.schema.nodes.paragraph.create());
};

/** Adds a row above (-1) or below (+1) the current row, with cells shaped like the current row's. */
export function addRow(where: -1 | 1): Command {
  return (state, dispatch) => {
    const ctx = cellContext(state.selection.$from);
    if (!ctx) return false;
    if (!dispatch) return true;
    const rows: Node[] = [];
    ctx.table.forEach((r) => rows.push(r));
    const like = rows[ctx.rowIndex];
    const cells: Node[] = [];
    like.forEach((c) => cells.push(emptyCell(state, c)));
    const { header: _h, ...rowProps } = like.attrs.props;
    rows.splice(ctx.rowIndex + (where > 0 ? 1 : 0), 0, state.schema.nodes.table_row.create({ props: rowProps }, cells));
    dispatch(rebuild(state, ctx, rows, ctx.table.attrs.grid).scrollIntoView());
    return true;
  };
}

/** Adds a column left (-1) or right (+1) of the current cell; the grid is rescaled to keep the table's width. */
export function addColumn(where: -1 | 1): Command {
  return (state, dispatch) => {
    const ctx = cellContext(state.selection.$from);
    if (!ctx) return false;
    if (!dispatch) return true;
    const at = ctx.cellIndex + (where > 0 ? 1 : 0);
    const grid: number[] = [...ctx.table.attrs.grid];
    const total = grid.reduce((a, b) => a + b, 0);
    grid.splice(Math.min(at, grid.length), 0, grid[Math.min(ctx.cellIndex, grid.length - 1)] ?? total / (grid.length || 1));
    const scale = total / grid.reduce((a, b) => a + b, 0);
    const scaled = grid.map((w) => w * scale);
    const rows: Node[] = [];
    ctx.table.forEach((r) => {
      const cells: Node[] = [];
      r.forEach((c) => cells.push(c));
      const like = cells[Math.min(ctx.cellIndex, cells.length - 1)];
      cells.splice(Math.min(at, cells.length), 0, emptyCell(state, like));
      rows.push(state.schema.nodes.table_row.create(r.attrs, withWidths(state, cells, scaled)));
    });
    dispatch(rebuild(state, ctx, rows, scaled).scrollIntoView());
    return true;
  };
}

/** Gives cells their grid column widths (Word honours cell widths, so they must follow the grid). */
function withWidths(state: EditorState, cells: Node[], grid: number[]): Node[] {
  let col = 0;
  return cells.map((c) => {
    const span = c.attrs.props.span ?? 1;
    const w = grid.slice(col, col + span).reduce((a, b) => a + b, 0);
    col += span;
    return state.schema.nodes.table_cell.create({ ...c.attrs, props: { ...c.attrs.props, width: { w: twips(w), type: "dxa" } } }, c.content);
  });
}

/** Deletes the current row (the table, if it was the last row). */
export const deleteRow: Command = (state, dispatch) => {
  const ctx = cellContext(state.selection.$from);
  if (!ctx) return false;
  if (!dispatch) return true;
  if (ctx.table.childCount === 1) return deleteTable(state, dispatch);
  const rows: Node[] = [];
  ctx.table.forEach((r, _, i) => { if (i !== ctx.rowIndex) rows.push(r); });
  dispatch(rebuild(state, ctx, rows, ctx.table.attrs.grid).scrollIntoView());
  return true;
};

/** Deletes the current column (the table, if it was the last column). */
export const deleteColumn: Command = (state, dispatch) => {
  const ctx = cellContext(state.selection.$from);
  if (!ctx) return false;
  if (!dispatch) return true;
  const grid: number[] = [...ctx.table.attrs.grid];
  if (grid.length <= 1) return deleteTable(state, dispatch);
  const total = grid.reduce((a, b) => a + b, 0);
  grid.splice(Math.min(ctx.cellIndex, grid.length - 1), 1);
  const scale = total / grid.reduce((a, b) => a + b, 0);
  const scaled = grid.map((w) => w * scale);
  const rows: Node[] = [];
  ctx.table.forEach((r) => {
    const cells: Node[] = [];
    r.forEach((c, _, i) => { if (i !== ctx.cellIndex) cells.push(c); });
    if (cells.length) rows.push(state.schema.nodes.table_row.create(r.attrs, withWidths(state, cells, scaled)));
  });
  dispatch(rebuild(state, ctx, rows, scaled).scrollIntoView());
  return true;
};

/** Deletes the whole table, leaving an empty paragraph in its place. */
export const deleteTable: Command = (state, dispatch) => {
  const ctx = cellContext(state.selection.$from);
  if (!ctx) return false;
  if (dispatch) {
    const tr = state.tr.replaceWith(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, state.schema.nodes.paragraph.create());
    dispatch(tr.setSelection(TextSelection.near(tr.doc.resolve(ctx.tablePos + 1))).scrollIntoView());
  }
  return true;
};

/** Tab / Shift+Tab inside a table: move to the next / previous cell (Tab in the last cell adds a row). */
export function moveCell(dir: 1 | -1): Command {
  return (state, dispatch, view) => {
    const ctx = cellContext(state.selection.$from);
    if (!ctx) return false;
    const cells: number[] = [];
    state.doc.nodesBetween(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, (node, pos) => {
      if (node.type.name === "table_cell") cells.push(pos);
      return node.type.name !== "table_cell";
    });
    const here = cells.findIndex((p) => p <= state.selection.from && state.selection.from < p + state.doc.nodeAt(p)!.nodeSize);
    const next = here + dir;
    if (next >= cells.length) {
      if (dir > 0 && dispatch) {
        addRow(1)(state, dispatch, view);
        return true;
      }
      return false;
    }
    if (next < 0) return false;
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(cells[next] + 1))).scrollIntoView());
    return true;
  };
}
