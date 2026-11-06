// Table editing through the commands: merge across and down, split, borders, shading, a table style
// with options, sorting, text ↔ table. Saved for Word to check.
// Usage: node tables.mjs <out.docx>
import { writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import {
  blankDocument, createSchema, fromProseMirror, insertTable, mergeCells, setCellBorders, setCellShading, setTableStyle, sortTable,
  splitCell, tableToText, textToTable, toggleHeaderRows, toProseMirror,
} from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const base = blankDocument();
base.sections[0].blocks = [
  { type: "paragraph", props: {}, content: [] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Name" }, { type: "tab" }, { type: "text", text: "Score" }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Ama" }, { type: "tab" }, { type: "text", text: "71" }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Kofi" }, { type: "tab" }, { type: "text", text: "88" }] },
];
let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const doc = () => base;
// A 5 × 3 table at the top, filled in.
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
run(insertTable(doc, 5, 3));
const table = () => { let t; state.doc.forEach((n, pos) => { if (!t && n.type.name === "table") t = { node: n, pos }; }); return t; };
const cellPos = (r, k) => { const { node, pos } = table(); let p = pos + 1; for (let i = 0; i < r; i++) p += node.child(i).nodeSize; p += 1; for (let i = 0; i < k; i++) p += node.child(r).child(i).nodeSize; return p; };
const data = [["Results", "", ""], ["Region", "Rainfall", "Yield"], ["Accra", "810", "1.9"], ["Ashanti", "1420", "2.6"], ["Volta", "1180", "2.2"]];
for (let r = data.length - 1; r >= 0; r--) for (let k = 2; k >= 0; k--) if (data[r][k]) state = state.apply(state.tr.insertText(data[r][k], cellPos(r, k) + 2));
const select = (r1, k1, r2, k2) => { state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cellPos(r1, k1) + 2, cellPos(r2, k2) + 2))); };
const shape = () => { const out = []; table().node.forEach((row) => { const cells = []; row.forEach((c) => cells.push(`${c.textContent}${c.attrs.props.span ? `[${c.attrs.props.span}]` : ""}${c.attrs.props.vMerge ? `{${c.attrs.props.vMerge[0]}}` : ""}`)); out.push(cells.join(" | ")); }); return out; };

select(0, 0, 0, 2); run(mergeCells);                       // the title across the table
console.log("merged across:", shape()[0]);
select(0, 0, 0, 0); run(splitCell); console.log("split back:   ", shape()[0]);
select(0, 0, 0, 2); run(mergeCells);
select(2, 0, 3, 0); run(mergeCells);                       // Accra and Ashanti share a region cell
console.log("merged down:  ", shape().slice(2, 4).join("  //  "));
select(2, 0, 3, 0); run(splitCell);                        // and back, for sorting
select(1, 0, 4, 2); run(setCellBorders("outside", { width: 1.5, color: "1F3864" }));
select(1, 0, 1, 2); run(setCellShading("D9E2F3"));
select(2, 1, 2, 1); run(setTableStyle(doc, "GridTable4-Accent1", { firstRow: true, lastRow: true, firstColumn: true, lastColumn: false, bandedRows: true, bandedColumns: false }));
select(0, 0, 1, 0); run(toggleHeaderRows);                   // title and column headings repeat on every page
select(2, 1, 2, 1); run(sortTable({ column: 1, descending: true, type: "number", header: true }));
console.log("table:        ", JSON.stringify(shape()));
// The tab-separated lines below become a table; then the first table back to text.
let p1 = -1; state.doc.forEach((n, pos) => { if (p1 < 0 && n.textContent === "Name\tScore".replace("\t", "")) p1 = pos; });
let last = 0; state.doc.forEach((n, pos) => { last = pos + n.nodeSize; });
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, p1 + 1, last - 1)));
run(textToTable(doc));
const tables = []; state.doc.forEach((n) => { if (n.type.name === "table") tables.push(n.childCount + "×" + n.child(0).childCount); });
console.log("tables:       ", tables.join(", "));
writeFileSync(process.argv[2], writeDocx(fromProseMirror(state.doc, base)));
// A copy with the first table turned back into text.
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cellPos(1, 0) + 2)));
run(tableToText());
const lines = []; state.doc.forEach((n) => { if (n.type.name === "paragraph" && n.textContent) lines.push(n.textContent.replace(/\t/g, " ⇥ ")); });
console.log("as text:      ", JSON.stringify(lines.slice(0, 4)));
