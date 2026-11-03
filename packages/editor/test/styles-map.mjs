// Cost of typing for the paragraph-styles decorations: map vs rebuild, on a long document.
// Usage: node styles-map.mjs <docx>
import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { createRender, createSchema, paragraphStyles, toProseMirror } from "@prime/editor";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
import { EditorState } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
const measurer = new SystemFontMeasurer();
const schema = createSchema(createRender(doc, measurer));
const plugin = paragraphStyles(doc, measurer);
let t = performance.now();
let state = EditorState.create({ doc: toProseMirror(doc, schema), plugins: [plugin] });
console.log("build ms", Math.round(performance.now() - t));
const set = plugin.getState(state).decorations;
let n = 0, local = 0; set.find().forEach(() => n++);
local = set.local?.length;
console.log("decorations", n, "at top level", local, "children", set.children?.length / 3);
let target = Math.floor(state.doc.content.size / 2);
state.doc.nodesBetween(target, state.doc.content.size, (node, pos) => { if (node.isTextblock && node.content.size && target === Math.floor(state.doc.content.size / 2)) { target = pos + 1; return false; } return true; });
const times = [];
for (let i = 0; i < 40; i++) {
  const tr = state.tr.insertText("x", target + i);
  t = performance.now();
  state = state.apply(tr);
  times.push(performance.now() - t);
}
times.sort((a, b) => a - b);
console.log("apply per key ms: median", times[20].toFixed(2), "max", times.at(-1).toFixed(2));
// The map alone, and with no decorations (the state's own cost).
const tr = state.tr.insertText("x", target);
const m = [];
for (let i = 0; i < 40; i++) { t = performance.now(); plugin.getState(state).decorations.map(tr.mapping, tr.doc); m.push(performance.now() - t); }
m.sort((a, b) => a - b);
console.log("map alone ms: median", m[20].toFixed(2));
const bare = EditorState.create({ doc: state.doc });
const b = [];
for (let i = 0; i < 40; i++) { const tr2 = bare.tr.insertText("x", target); t = performance.now(); bare.apply(tr2); b.push(performance.now() - t); }
b.sort((a, c) => a - c);
console.log("apply without plugin ms: median", b[20].toFixed(3));
{
  const { DecorationSet } = await import("prosemirror-view");
  const all = plugin.getState(state).decorations.find();
  const c = [];
  for (let i = 0; i < 20; i++) { t = performance.now(); DecorationSet.create(tr.doc, all.map((d) => d)); c.push(performance.now() - t); }
  c.sort((a, b) => a - b);
  console.log("create from", all.length, "decorations ms: median", c[10].toFixed(2));
}
