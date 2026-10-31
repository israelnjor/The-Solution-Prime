// Different first page + per-section headers through the real commands; Word verifies.
// Usage: node sections.mjs <in.docx (with 2 sections)> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { createSchema, fromProseMirror, sectionInfo, setStory, setTitlePage, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const story = (t) => ({ blocks: [{ type: "paragraph", props: {}, content: [{ type: "text", text: t }] }] });

// Section 1 (cursor at start): different first page, with its own first-page header.
state = state.apply(state.tr.setSelection(TextSelection.atStart(state.doc)));
console.log("cursor in", JSON.stringify(sectionInfo(state)));
run(setTitlePage(true));
run(setStory("header", story("COVER PAGE HEADER"), "first"));
run(setStory("header", story("Section one header"), "default"));
// Section 2 (cursor at end): its own header.
state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
console.log("cursor in", JSON.stringify(sectionInfo(state)));
run(setStory("header", story("Section two header"), "default"));
writeFileSync(process.argv[3], writeDocx(fromProseMirror(state.doc, doc)));
