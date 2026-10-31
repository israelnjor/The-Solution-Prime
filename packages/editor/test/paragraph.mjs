// Paragraph formatting through the real commands; saves a .docx and a PDF for verification.
// Usage: node paragraph.mjs <out.docx> <out.pdf>
import { writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import { blankDocument, createSchema, fromProseMirror, setBorders, setIndent, setLineSpacing, setParagraphSpacing, toProseMirror } from "@prime/editor";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
import { renderPdf } from "@prime/pdf";
import { EditorState, TextSelection } from "prosemirror-state";

const doc = blankDocument();
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema) });
const run = (cmd) => { if (!cmd(state, (tr) => { state = state.apply(tr); })) throw new Error("refused"); };
const para = (text) => { state = state.apply(state.tr.insertText(text)); };
const enter = () => { state = state.apply(state.tr.split(state.selection.from)); };
const selectPara = (n) => { // n-th paragraph (0-based)
  let i = 0, pos;
  state.doc.forEach((node, offset) => { if (i++ === n) pos = offset + 1; });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
};
const long = "Prime Word paragraph formatting test. ".repeat(6);
para("Double spaced. " + long); enter();
para("Spacing 18pt before and 12pt after. " + long); enter();
para("Indented 1 cm left and right, first line 1.27 cm. " + long); enter();
para("Boxed paragraph. " + long); enter();
para("Bottom border only."); enter();
para("Exactly 20pt lines. " + long);
selectPara(0); run(setLineSpacing("multiple", 2));
selectPara(1); run(setParagraphSpacing({ before: 18, after: 12 }));
selectPara(2); run(setIndent({ left: 28.35, right: 28.35, firstLine: 36 }));
selectPara(3); run(setBorders("box"));
selectPara(4); run(setBorders("bottom", "C00000"));
selectPara(5); run(setLineSpacing("exact", 20));
const out = fromProseMirror(state.doc, doc);
writeFileSync(process.argv[2], writeDocx(out));
const m = new SystemFontMeasurer();
writeFileSync(process.argv[3], await renderPdf(out, layoutDocument(out, m), m));
