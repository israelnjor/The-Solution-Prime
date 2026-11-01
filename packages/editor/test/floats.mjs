// Moves and resizes a floating picture and a text box through the real commands; Word verifies.
// Usage: node floats.mjs <in.docx (with floating objects)> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { createSchema, deleteFloat, floatAt, fromProseMirror, setFloatBox, toProseMirror } from "@prime/editor";
import { EditorState } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };

// Floating objects by paragraph: [paragraph position, how many].
const anchors = [];
state.doc.descendants((node, pos) => {
  if (node.type.name !== "paragraph") return true;
  let n = 0;
  node.descendants((c) => { if (c.type.name === "run_preserved" && c.attrs.float) n++; });
  if (n) anchors.push([pos, n]);
  return false;
});
console.log("anchors", JSON.stringify(anchors));
const [para] = anchors[0];
const kinds = Array.from({ length: anchors[0][1] }, (_, i) => floatAt(state.doc, para, i).node.attrs.float.kind);
const pic = kinds.indexOf("picture");
const box = kinds.indexOf("textbox");
console.log("kinds", kinds.join(","));
// Picture: to 40pt right of the column's left edge, 50pt below the paragraph, 200pt wide (proportional).
const f = floatAt(state.doc, para, pic).node.attrs.float;
run(setFloatBox(para, pic, { x: 40, y: 50, width: 200, height: (200 * f.height) / f.width }));
// Text box: 300pt wide, at x 10, y 400.
run(setFloatBox(para, box, { x: 10, y: 400, width: 300, height: 0 }));
// The last floating object in the document is removed.
const [lastPara, lastN] = anchors[anchors.length - 1];
run(deleteFloat(lastPara, lastN - 1));
writeFileSync(process.argv[3], writeDocx(fromProseMirror(state.doc, doc)));
console.log("saved", process.argv[3]);
