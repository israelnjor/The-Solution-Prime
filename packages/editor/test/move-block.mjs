// Alt+Shift+↑/↓ (moveBlock): paragraphs and a table move past their neighbours; text is unchanged.
import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { createSchema, moveBlock, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
let state = EditorState.create({ doc: toProseMirror(doc, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const tops = () => { const out = []; state.doc.forEach((n) => out.push(n.type.name === "table" ? "[table]" : n.textContent.slice(0, 18))); return out; };
const before = tops();
// Cursor in the third top-level block, move it up twice, then down once.
let pos = 0; for (let i = 0; i < 2; i++) pos += state.doc.child(i).nodeSize;
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + 2)));
run(moveBlock(-1)); run(moveBlock(-1));
const up = tops();
run(moveBlock(1));
console.log("before:", JSON.stringify(before.slice(0, 4)));
console.log("up x2: ", JSON.stringify(up.slice(0, 4)));
console.log("down:  ", JSON.stringify(tops().slice(0, 4)));
// In a table, the row moves (as in Word).
const t = tops().indexOf("[table]");
pos = 0; for (let i = 0; i < t; i++) pos += state.doc.child(i).nodeSize;
const table = () => state.doc.child(t);
const rows = () => { const out = []; table().forEach((r) => out.push(r.textContent.slice(0, 12))); return out; };
const firstRow = table().child(0).nodeSize;
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + firstRow + 4)));
const rowsBefore = rows();
run(moveBlock(-1));
console.log("rows before:", JSON.stringify(rowsBefore.slice(0, 3)), "after:", JSON.stringify(rows().slice(0, 3)));
console.log("same text:", state.doc.textContent.length === toProseMirror(doc, createSchema()).textContent.length);
