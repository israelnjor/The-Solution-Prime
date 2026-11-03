// Picture wrapping through the real commands: in line → square (moved) → behind; Word verifies.
// Usage: node picture-wrap.mjs <in.docx (with an in-line picture)> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { createSchema, floatAt, fromProseMirror, pictureWrap, setFloatBox, setPictureWrap, toProseMirror } from "@prime/editor";
import { EditorState, NodeSelection } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
let state = EditorState.create({ doc: toProseMirror(doc, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
let pos = -1;
state.doc.descendants((n, p) => { if (pos < 0 && n.type.name === "image") pos = p; return pos < 0; });
state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, pos)));
console.log("before:", pictureWrap(state.doc.nodeAt(pos)));
run(setPictureWrap("square"));
console.log("after:", pictureWrap(state.doc.nodeAt(pos)));
const $p = state.doc.resolve(pos);
const para = $p.before($p.depth);
const { node } = floatAt(state.doc, para, 0);
run(setFloatBox(para, 0, { x: 150, y: 20, width: node.attrs.width / 2, height: node.attrs.height / 2 }));
writeFileSync(process.argv[3], writeDocx(fromProseMirror(state.doc, doc)));
console.log("saved");
