// Footnotes through the commands: insert one between existing ones (it takes number 2, the rest move
// up), give it text, delete the first, add an endnote; save for Word to check.
// Usage: node notes.mjs <in.docx (Word document with footnotes)> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { createSchema, deleteNote, fromProseMirror, insertNote, noteList, setNoteText, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const { doc } = readDocx(readFileSync(process.argv[2]));
let state = EditorState.create({ doc: toProseMirror(doc, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const show = () => noteList(state).map((e) => `${e.kind[0]}${e.number}:${e.note.blocks.map((b) => b.content.filter((i) => i.type === "text").map((i) => i.text).join("")).join("").trim().slice(0, 28)}`);
console.log("before:", JSON.stringify(show()));

// After "warming" (between footnotes 1 and 2).
let at = -1;
state.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text.includes("Rainfall")) at = pos + n.text.indexOf("Rainfall") + "Rainfall".length; return at < 0; });
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)));
let created;
run(insertNote(() => doc, "footnote", (id) => { created = id; }));
const text = [{ type: "paragraph", props: { style: "FootnoteText" }, content: [{ type: "noteMark", kind: "footnote", props: { style: "FootnoteReference" } }, { type: "text", text: " Added in Prime Word: rainfall records since 1960." }] }];
run(setNoteText("footnote", created, text));
console.log("inserted:", JSON.stringify(show()));
const first = noteList(state).find((e) => e.kind === "footnote" && e.number === "1");
run(deleteNote("footnote", first.id));
console.log("deleted 1:", JSON.stringify(show()));
// An endnote at the very end.
state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
let end;
run(insertNote(() => doc, "endnote", (id) => { end = id; }));
run(setNoteText("endnote", end, [{ type: "paragraph", props: { style: "EndnoteText" }, content: [{ type: "noteMark", kind: "endnote" }, { type: "text", text: " Endnote added in Prime Word." }] }]));
console.log("final:", JSON.stringify(show()));
writeFileSync(process.argv[3], writeDocx(fromProseMirror(state.doc, doc)));
console.log("wrote", process.argv[3]);
