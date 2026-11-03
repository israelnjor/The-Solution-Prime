// Blocks (quick parts): a signature block and an address inserted at the cursor (each alignment), into
// the header and footer, and a selection saved and inserted again; the result survives a .docx round trip.
// Usage: node parts.mjs [out.docx]
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import {
  addressBlock, blankDocument, createSchema, fromProseMirror, insertPart, partFromSelection, picturePart, signatureBlock, storyAt, toProseMirror,
} from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const png = readFileSync(new URL("./sample.png.b64", import.meta.url), "utf8").trim();
const base = blankDocument();
base.sections[0].blocks = [
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Dear Sir/Madam," }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Thank you for your letter." }] },
  { type: "paragraph", props: {}, content: [] },
];
let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const end = () => state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
const doc = () => base;

const sig = signatureBlock({ closing: "Yours faithfully,", signature: { png, width: 200, height: 100 }, line: true, name: "Ama Mensah", title: "Head Librarian", organisation: "Accra Community Library", date: "3 November 2026" });
state = end(); run(insertPart(sig, doc, "cursor", "left"));
state = end(); run(insertPart(sig, doc, "cursor", "right"));
const addr = addressBlock(["Accra Community Library", "12 Independence Avenue", "Accra, Ghana"], { bold: true });
run(insertPart(addr, doc, "header", "right"));
run(insertPart(addr, doc, "footer", "center"));
run(insertPart(picturePart({ data: png, width: 200, height: 100 }, { width: 90, name: "Stamp", float: true }), doc, "cursor", "right"));

// Save "Thank you for your letter." as a part and insert it at the end.
let from = state.doc.child(0).nodeSize + 1;
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, from + state.doc.child(1).content.size)));
const saved = partFromSelection(doc)(state);
state = end(); run(insertPart(saved, doc, "cursor", "center"));

const out = fromProseMirror(state.doc, base);
const text = (blocks) => blocks.map((b) => b.type === "paragraph" ? `${b.props.align ?? "-"}:${b.content.map((i) => i.type === "text" ? i.text : `[${i.float ? "float" : i.type}]`).join("")}` : `[${b.type}]`);
console.log("body:  ", JSON.stringify(text(out.sections[0].blocks)));
console.log("header:", JSON.stringify(text(storyAt(state, "header").blocks)));
console.log("footer:", JSON.stringify(text(storyAt(state, "footer").blocks)));
const images = new Set(); JSON.stringify(out.sections[0].blocks, (k, v) => { if (k === "imageId") images.add(v); return v; });
console.log("pictures:", [...images].map((id) => `${id}:${!!out.images[id]}`).join(" "));
const bytes = writeDocx(out);
const back = readDocx(bytes).doc;
console.log("round trip:", JSON.stringify(text(back.sections[0].blocks)) === JSON.stringify(text(out.sections[0].blocks)) ? "same" : "DIFFERENT");
if (process.argv[2]) writeFileSync(process.argv[2], bytes);
