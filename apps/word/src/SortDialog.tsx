import { selectedCells, sortTable } from "@prime/editor";
import type { EditorView } from "prosemirror-view";
import { useState } from "react";

/** Table Layout → Sort: by a column, ascending or descending, as text, numbers or dates. */
export function SortDialog({ view, onClose }: { view: EditorView; onClose(): void }) {
  const range = selectedCells(view.state);
  // Column names from the first row (or "Column 1, 2…").
  const first = range?.grid.rows[0];
  const columns = (range?.grid.table.attrs.grid as number[] | undefined)?.map((_, i) => {
    const text = first?.cells.find((c) => c.start <= i && c.start + c.span > i)?.node.textContent.trim();
    return text ? `${text} (column ${i + 1})` : `Column ${i + 1}`;
  }) ?? [];
  const [column, setColumn] = useState(range?.c1 ?? 0);
  const [type, setType] = useState<"text" | "number" | "date">("text");
  const [descending, setDescending] = useState(false);
  const [header, setHeader] = useState(true);
  const [error, setError] = useState("");
  const sort = () => {
    if (!sortTable({ column, type, descending, header })(view.state, view.dispatch)) {
      setError("This table has cells merged across rows, which cannot be sorted. Split them first.");
      return;
    }
    onClose();
    view.focus();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter") sort(); }}>
      <div className="hf-dialog sort-dialog" role="dialog" aria-label="Sort">
        <h3>Sort</h3>
        <label>Sort by<select value={column} onChange={(e) => setColumn(Number(e.target.value))}>{columns.map((c, i) => <option key={i} value={i}>{c}</option>)}</select></label>
        <label>Type<select value={type} onChange={(e) => setType(e.target.value as "text")}>
          <option value="text">Text</option><option value="number">Number</option><option value="date">Date</option>
        </select></label>
        <label>Order<select value={descending ? "d" : "a"} onChange={(e) => setDescending(e.target.value === "d")}>
          <option value="a">Ascending (A to Z, 0 to 9)</option><option value="d">Descending (Z to A, 9 to 0)</option>
        </select></label>
        <label className="choice"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> My table has a header row (it stays at the top)</label>
        {error && <p className="eq-error">{error}</p>}
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={sort}>OK</button>
        </div>
      </div>
    </div>
  );
}
