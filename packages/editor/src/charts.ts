/**
 * Insert → Chart, and editing a chart's data: a chart added to the document (its part and spreadsheet
 * in the package, its drawing at the cursor), and the data of a chart already there, rewritten.
 */
import { addChart, type ChartData, updateChart } from "@prime/docx";
import { chartTheme, type ChartSpec, parseChart } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection } from "prosemirror-state";
import { nextDrawingId } from "./shapes.js";

/** Word's sample chart: four categories, three series. */
export const SAMPLE_CHART: ChartData = {
  kind: "column",
  title: "Chart Title",
  categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
  series: [
    { name: "Series 1", values: [4.3, 2.5, 3.5, 4.5] },
    { name: "Series 2", values: [2.4, 4.4, 1.8, 2.8] },
    { name: "Series 3", values: [2, 2, 3, 5] },
  ],
};

/** A chart (w × h pt, 6" × 3.5" by default) at the cursor, selected. */
export function insertChart(base: () => PrimeDocument, data: ChartData, w = 432, h = 252): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.run_preserved;
    if (!type) return false;
    if (dispatch) {
      const xml = addChart(base(), data, w, h, nextDrawingId(state.doc));
      const node = type.create({ props: {}, preserved: { name: "w:drawing", xml }, float: null });
      // After a selected object (the chart just inserted, a picture), not in its place; else at the cursor.
      const sel = state.selection;
      const at = sel instanceof NodeSelection ? sel.to : sel.from;
      const tr = sel instanceof NodeSelection ? state.tr.insert(at, node) : state.tr.replaceSelectionWith(node, false);
      const pos = tr.mapping.map(at, -1);
      if (tr.doc.nodeAt(pos)?.type === type) tr.setSelection(NodeSelection.create(tr.doc, pos));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** The chart drawing at `pos` (or the selected one): its relationship and what it shows. */
export function chartAt(doc: Node, pos: number, base: PrimeDocument): { pos: number; node: Node; relId: string; spec?: ChartSpec } | undefined {
  const node = doc.nodeAt(pos);
  const xml = node?.attrs?.preserved?.xml as string | undefined;
  const relId = xml && /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(xml)?.[1];
  if (!node || !relId) return undefined;
  const rel = base.package.preservedRels.find((r) => r.id === relId);
  const part = rel && base.package.parts.find((p) => p.name === rel.target);
  const text = (b64: string) => new TextDecoder().decode(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
  const theme = base.package.parts.find((p) => /^word\/theme\/theme\d*\.xml$/.test(p.name));
  const spec = part ? parseChart(text(part.data), chartTheme(theme ? text(theme.data) : undefined)) : undefined;
  return { pos, node, relId, ...(spec ? { spec } : {}) };
}

/** New data for the chart at `pos` (Edit Data): its part rewritten, the chart drawn again. */
export function setChartData(base: () => PrimeDocument, pos: number, data: ChartData): Command {
  return (state, dispatch) => {
    const c = chartAt(state.doc, pos, base());
    if (!c) return false;
    if (dispatch) {
      if (!updateChart(base(), c.relId, data)) return false;
      // The drawing is unchanged; a new revision mark makes it drawn again (and undo restores the old look).
      const preserved = { ...c.node.attrs.preserved, rev: (Number(c.node.attrs.preserved.rev) || 0) + 1 };
      const tr = state.tr.setNodeMarkup(pos, undefined, { ...c.node.attrs, preserved });
      dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
    }
    return true;
  };
}

/** A number as people type it in a table or spreadsheet: "1,234", "45%", "$12.50", "(3)" for -3. */
export function parseChartNumber(s: string): number {
  let t = s.trim().replace(/[\s\u00a0,]/g, "").replace(/^[$€£¥₦]|[$€£¥₦%]$/g, "");
  if (/^\(.*\)$/.test(t)) t = `-${t.slice(1, -1)}`;
  return t === "" ? NaN : Number(t);
}

/**
 * Chart data from a grid of cells (rows of text: pasted from Excel, or a table's cells): a header row
 * of series names and a first column of categories when they hold words; numbers otherwise. A scatter
 * chart's first column is its x values. Undefined when the grid holds no numbers.
 */
export function chartFromGrid(grid: string[][], kind: ChartData["kind"] = "column", title?: string): ChartData | undefined {
  const rows = grid.map((r) => r.map((c) => c.trim())).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return undefined;
  const width = Math.max(...rows.map((r) => r.length));
  const cell = (i: number, j: number) => rows[i]?.[j] ?? "";
  const isNum = (s: string) => Number.isFinite(parseChartNumber(s));
  // Labels down the first column: any word there (below a header), or a scatter chart's x values.
  const labels = kind === "scatter" || rows.slice(1).some((r) => r[0] && !isNum(r[0])) || (rows.length === 1 && !isNum(cell(0, 0)));
  const from = labels ? 1 : 0;
  // A header row: words across the top (the top-left corner may hold anything).
  const header = rows.length > 1 && Array.from({ length: width - from }, (_, j) => cell(0, j + from)).some((c) => c && !isNum(c));
  const body = header ? rows.slice(1) : rows;
  if (from >= width || !body.length) return undefined;
  const series = Array.from({ length: width - from }, (_, j) => ({
    name: header && cell(0, j + from) ? cell(0, j + from) : `Series ${j + 1}`,
    values: body.map((r) => parseChartNumber(r[j + from] ?? "")),
  })).filter((s) => s.values.some((v) => Number.isFinite(v)));
  if (!series.length) return undefined;
  const categories = body.map((r, i) => (labels ? r[0] || `${i + 1}` : `${i + 1}`));
  return { kind, ...(title ? { title } : {}), categories, series };
}

/** The table around the cursor: its cells' text, row by row (merged cells once), and where it ends. */
export function tableAtCursor(state: EditorState): { grid: string[][]; end: number } | undefined {
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name !== "table") continue;
    const grid: string[][] = [];
    node.forEach((row) => {
      const cells: string[] = [];
      row.forEach((cell) => {
        const texts: string[] = [];
        cell.forEach((block) => texts.push(block.textContent));
        cells.push(texts.join(" "));
      });
      grid.push(cells);
    });
    return { grid, end: $from.after(d) };
  }
  return undefined;
}

/** A chart in a paragraph of its own after the table that ends at `end` (Chart from Table). */
export function insertChartAfter(base: () => PrimeDocument, data: ChartData, end: number, w = 432, h = 252): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.run_preserved;
    const para = state.schema.nodes.paragraph;
    if (!type || !para) return false;
    if (dispatch) {
      const xml = addChart(base(), data, w, h, nextDrawingId(state.doc));
      const node = type.create({ props: {}, preserved: { name: "w:drawing", xml }, float: null });
      const tr = state.tr.insert(end, para.create(null, node));
      tr.setSelection(NodeSelection.create(tr.doc, end + 1));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}
