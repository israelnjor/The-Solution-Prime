// Paste Options: paste formatted text, then switch it to Merge, Text only, and back to Keep.
import { blankDocument, createSchema, pasteAs, pasteOptions, toProseMirror } from "@prime/editor";
import { Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

const base = blankDocument();
base.sections[0].blocks = [{ type: "paragraph", props: {}, content: [{ type: "text", text: "Before  after", props: { font: { ascii: "Georgia" }, size: 14 } }] }];
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(base, schema), plugins: [pasteOptions()] });
const src = blankDocument();
src.sections[0].blocks = [
  { type: "paragraph", props: { style: "Heading1", align: "center" }, content: [{ type: "text", text: "Big title", props: { bold: true, size: 28, color: "FF0000" } }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "second line", props: { italic: true, font: { ascii: "Arial" } } }] },
];
const pasted = toProseMirror(src, schema);
const slice = new Slice(pasted.content, 1, 1);
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 8)));
state = state.apply(state.tr.replaceSelection(slice).setMeta("paste", true));
const run = (cmd) => cmd(state, (tr) => { state = state.apply(tr); });
const show = () => { const out = []; state.doc.forEach((p) => { const r = []; p.forEach((t) => r.push(`${t.text}${JSON.stringify(t.marks[0]?.attrs.props ?? {})}`)); out.push(`${JSON.stringify(p.attrs.props)} ${r.join(" ")}`); }); return out.join("\n   "); };
console.log("keep:  ", show());
run(pasteAs("merge")); console.log("merge: ", show());
run(pasteAs("text")); console.log("text:  ", show());
run(pasteAs("keep")); console.log("keep:  ", show());
