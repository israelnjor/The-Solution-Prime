// Edits a document through the real editor commands, saves it, and reports what was written.
// Usage: node edit-roundtrip.mjs <in.docx> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { createSchema, fromProseMirror, setParagraph, toggleRun, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const [input, output] = process.argv.slice(2);
const { doc } = readDocx(readFileSync(input));
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema) });
const run = (cmd) => cmd(state, (tr) => { state = state.apply(tr); });

// 1. Find the first plain paragraph and type a sentence at its start.
let target;
state.doc.descendants((node, pos) => {
  if (!target && node.type.name === "paragraph" && node.textContent.length > 20 && !node.attrs.props.style) target = pos;
});
state = state.apply(state.tr.insertText("EDITED IN PRIME WORD. ", target + 1));
// 2. Bold the word "EDITED".
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, target + 1, target + 7)));
run(toggleRun("bold"));
// 3. Make the paragraph a centred Heading 1.
run(setParagraph("style", "Heading1"));
run(setParagraph("align", "center"));

writeFileSync(output, writeDocx(fromProseMirror(state.doc, doc)));
console.log("saved", output);
