// Wave 1 formatting in a .docx for Word to check: character spacing, double underline, superscript,
// tab stops set by setTabStops (with a style stop cleared), a DATE field, keep with next.
// Usage: node wave1-word.mjs <out.docx>
import { writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import { blankDocument, createSchema, fromProseMirror, insertFieldResult, setParagraph, setTabStops, toProseMirror, updateRun } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const base = blankDocument();
base.sections[0].blocks = [
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Expanded by 2pt" }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Double underlined E=mc2" }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Name" }, { type: "tab" }, { type: "text", text: "Centre" }, { type: "tab" }, { type: "text", text: "12.50" }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Dated: " }] },
];
let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
const run = (cmd) => cmd(state, (tr) => { state = state.apply(tr); });
const starts = []; state.doc.forEach((n, pos) => starts.push(pos));
const sel = (i, a, b) => { state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, starts[i] + 1 + a, starts[i] + 1 + b))); };
sel(0, 0, 8); run(updateRun((p) => ({ ...p, charSpacing: 2 })));
sel(1, 0, 17); run(updateRun((p) => ({ ...p, underline: { val: "double" } })));
sel(1, 22, 23); run(updateRun((p) => ({ ...p, vertAlign: "superscript" })));
sel(2, 0, 0); run(setTabStops(() => base, [{ pos: 5 * 72 / 2.54, align: "center" }, { pos: 12 * 72 / 2.54, align: "right", leader: "dot" }]));
run(setParagraph("keepNext", true));
sel(3, 7, 7); run(insertFieldResult(String.raw`DATE \@ "d MMMM yyyy"`, "3 November 2026"));
writeFileSync(process.argv[2], writeDocx(fromProseMirror(state.doc, base)));
console.log("wrote", process.argv[2]);
