// Citations and a reference list: three sources, citations in APA, the list; then IEEE after Update
// Field. Saved twice for Word: APA (<out>-apa.docx) and IEEE (<out>-ieee.docx).
// Usage: node citations.mjs <out-prefix>
import { writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import {
  blankDocument, createSchema, fromProseMirror, insertBibliography, insertCitation, newTag, readSources, toProseMirror, updateFields, writeSources,
} from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const base = blankDocument();
const sources = [];
const add = (s) => { sources.push({ ...s, tag: newTag(s, sources) }); };
add({ type: "Book", authors: [{ last: "Mensah", first: "Ama" }, { last: "Owusu", first: "Kwame" }], title: "Climate and Farming in Ghana", year: "2024", publisher: "Sub-Saharan Publishers", city: "Accra" });
add({ type: "JournalArticle", authors: [{ last: "Boateng", first: "Yaw" }, { last: "Asante", first: "Efua" }, { last: "Darko", first: "Kofi" }], title: "Rainfall variability in the Volta Basin", year: "2021", journal: "Journal of African Earth Sciences", volume: "174", issue: "2", pages: "104–117" });
add({ type: "InternetSite", authors: [{ last: "Ghana Statistical Service" }], title: "Agricultural census 2023", year: "2023", publisher: "GSS", url: "https://statsghana.gov.gh/census", accessed: "2026-11-03" });
writeSources(base, "APA", sources);
base.sections[0].blocks = [
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Farmers plant earlier now " }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Rainfall varies widely " }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Census figures confirm it " }] },
  { type: "paragraph", props: {}, content: [] },
];
let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const endOf = (i) => { let p = 0; for (let k = 0; k <= i; k++) p += state.doc.child(k).nodeSize; return p - 1; };
const doc = () => base;
const tags = readSources(base).sources.map((s) => s.tag);
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, endOf(1)))); run(insertCitation(doc, tags[1]));
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, endOf(0)))); run(insertCitation(doc, tags[0], "12"));
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, endOf(2)))); run(insertCitation(doc, tags[2]));
state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc))); run(insertBibliography(doc));
run(updateFields(doc)); // numbers and order as cited
const text = () => { const out = []; state.doc.forEach((n) => out.push(n.textContent)); return out; };
console.log("tags:", tags.join(", "));
console.log("APA:\n  " + text().join("\n  "));
writeFileSync(`${process.argv[2]}-apa.docx`, writeDocx(fromProseMirror(state.doc, base)));
writeSources(base, "IEEE", readSources(base).sources);
run(updateFields(doc));
console.log("IEEE:\n  " + text().join("\n  "));
writeFileSync(`${process.argv[2]}-ieee.docx`, writeDocx(fromProseMirror(state.doc, base)));
