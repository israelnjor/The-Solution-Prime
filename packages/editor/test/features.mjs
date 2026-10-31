// Exercises the editing commands the way a user would, then saves a .docx for Word to verify.
// Usage: node features.mjs <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import {
  addColumn, addRow, blankDocument, changeLevel, createSchema, deleteColumn, findAll, findPlugin, fromProseMirror,
  insertImage, insertPageBreak, insertTable, replaceAll, setColor, setFindQuery, setFont, setFontSize, setHighlight,
  setLink, setParagraph, toggleList, toggleRun, toProseMirror,
} from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const doc = blankDocument();
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema), plugins: [findPlugin()] });
const run = (cmd) => { const ok = cmd(state, (tr) => { state = state.apply(tr); }); if (!ok) throw new Error("command refused"); };
const type = (text) => { state = state.apply(state.tr.insertText(text)); };
const enter = () => { state = state.apply(state.tr.split(state.selection.from)); };
const selectText = (text) => {
  let at;
  state.doc.descendants((n, pos) => { if (at === undefined && n.isText && n.text.includes(text)) at = pos + n.text.indexOf(text); });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at, at + text.length)));
};
const end = () => { state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc))); };

type("Prime Word feature test"); run(setParagraph("style", "Title")); enter(); run(setParagraph("style", undefined));
type("Shopping list:"); enter();
type("Apples"); run(toggleList(() => doc, "bullet")); enter(); type("Green apples"); run(changeLevel(1)); enter();
type("Oranges"); run(changeLevel(-1)); enter(); run(toggleList(() => doc, "bullet")); // leave the list
type("Steps:"); enter();
type("Open the file"); run(toggleList(() => doc, "number")); enter(); type("Edit it"); enter(); type("Save it"); enter();
run(toggleList(() => doc, "number"));
type("Styled words: RED ARIAL fourteen and highlighted."); enter();
selectText("RED ARIAL"); run(setFont("Arial")); run(setFontSize(14)); run(setColor("C00000")); run(toggleRun("bold"));
selectText("highlighted"); run(setHighlight("yellow"));
end(); type("Visit the site"); selectText("the site"); run(setLink("https://example.com/prime"));
end(); enter(); type("Before the break."); run(insertPageBreak); type("After the break, on page 2.");
end(); enter(); run(insertTable(() => doc, 2, 3)); type("A1");
run(addRow(1)); run(addColumn(1)); run(addColumn(1)); run(deleteColumn);
end(); type("Picture below:"); enter();
run(insertImage(() => doc, { data: readFileSync(new URL("./sample.png.b64", import.meta.url), "utf8"), contentType: "image/png", width: 200, height: 100, name: "Sample" }));
end(); enter(); type("Colour colour COLOUR.");
run(setFindQuery({ text: "colour" }));
const before = findAll(state.doc, { text: "colour" }).length;
let replaced = 0; run(replaceAll("color", (n) => { replaced = n; }));

state.doc.check();
writeFileSync(process.argv[2], writeDocx(fromProseMirror(state.doc, doc)));
console.log(JSON.stringify({ findMatches: before, replaced, lists: doc.lists.definitions.length, images: Object.keys(doc.images).length }));
