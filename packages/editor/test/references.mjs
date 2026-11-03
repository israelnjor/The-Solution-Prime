// References: a table of contents, captions, a table of figures and cross-references, inserted
// through the real commands; page numbers from our layout; Update Field; saved for Word to check.
// Usage: node references.mjs <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import {
  blankDocument, captions, createSchema, fromProseMirror, headings, insertCaption, insertCrossReference, insertToc, refreshPageRefs,
  toProseMirror, updateFields,
} from "@prime/editor";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
import { EditorState, TextSelection } from "prosemirror-state";

const png = readFileSync(new URL("./sample.png.b64", import.meta.url), "utf8").trim();
const base = blankDocument();
base.images["word/media/chart.png"] = { contentType: "image/png", extension: "png", data: png };
const para = (text, style) => ({ type: "paragraph", props: style ? { style } : {}, content: text ? [{ type: "text", text }] : [] });
const filler = (n) => Array.from({ length: n }, (_, i) => para(`Body text ${i + 1}. Prime Word keeps the meaning and structure of documents through every edit, and the layout matches Word's.`));
const picture = { type: "paragraph", props: {}, content: [{ type: "image", imageId: "word/media/chart.png", width: 200, height: 100 }] };
base.sections[0].blocks = [
  para(""),
  para("Introduction", "Heading1"), ...filler(14),
  para("Background", "Heading2"), ...filler(10), picture,
  para("Method", "Heading1"), ...filler(18),
  para("Data collection", "Heading2"), ...filler(6), picture,
  para("Results", "Heading1"), para("See the results below."), ...filler(8),
];
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(base, schema) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const doc = () => base;
const at = (pos) => { state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos))); };
const nth = (pred, k = 0) => { const hits = []; state.doc.forEach((n, pos) => { if (pred(n)) hits.push(pos); }); return hits[k]; };
const isPicture = (n) => n.firstChild?.type.name === "image";

// Captions under both pictures.
at(nth(isPicture, 0) + 1); run(insertCaption(doc, "Figure", "Monthly sales by region"));
at(nth(isPicture, 1) + 1); run(insertCaption(doc, "Figure", "Survey responses"));
// Contents and a table of figures at the top (in the empty first paragraph, and after it).
at(1); run(insertToc(doc, { levels: 3 }));
const afterToc = nth((n) => n.attrs.props?.style === "Heading1");
at(afterToc + 1); run(insertToc(doc, { caption: "Figure" }, {}, false));
// Cross-references in "See the results below.": to Figure 2 (label and number) and its page, and to a heading.
const fig2 = captions(state, "Figure")[1];
const results = nth((n) => n.textContent === "See the results below.");
at(results + 1 + "See ".length); run(insertCrossReference(fig2, "labelNumber"));
state = state.apply(state.tr.insertText(" on page ")); run(insertCrossReference(fig2, "page"));
state = state.apply(state.tr.insertText(" and the section ")); run(insertCrossReference(headings(state, base).find((h) => h.text === "Method"), "text"));
state = state.apply(state.tr.insertText(", "));

// Pages from our layout, then Update Field, twice (the contents can move pages itself).
for (let i = 0; i < 2; i++) {
  const model = fromProseMirror(state.doc, base);
  const { bookmarks } = layoutDocument(model, new SystemFontMeasurer());
  run(updateFields(doc, bookmarks));
}
const model = fromProseMirror(state.doc, base);
const layout = layoutDocument(model, new SystemFontMeasurer());
refreshPageRefs(model, layout.bookmarks);
const entries = [];
state.doc.forEach((n) => { if (/^(TOC|TableofFigures)/.test(n.attrs.props?.style ?? "")) entries.push(n.textContent.replace(/\t/g, " … ")); });
console.log("pages:", layout.pages.length);
console.log("entries:", JSON.stringify(entries));
console.log("captions:", captions(state, "Figure").map((c) => c.text).join(" | "));
console.log("results line:", state.doc.nodeAt(nth((n) => n.textContent.startsWith("See ")))?.textContent);
writeFileSync(process.argv[2], writeDocx(model));
console.log("wrote", process.argv[2]);
