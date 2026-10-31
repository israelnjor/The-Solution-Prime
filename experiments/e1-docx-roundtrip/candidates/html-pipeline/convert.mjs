// Candidate: DOCX -> HTML (mammoth) -> DOCX (html-to-docx).
// This is the path a typical web/contenteditable editor takes when it has no native DOCX model,
// so it is a useful floor for what "naive" import/export loses.
// Usage: node convert.mjs <in.docx> <out.docx>   (writes <out.docx>.warnings.json)
import { readFile, writeFile } from "node:fs/promises";
import mammoth from "mammoth";
import HTMLtoDOCX from "html-to-docx";

const [input, output] = process.argv.slice(2);
const { value: html, messages } = await mammoth.convertToHtml({ path: input });
const docx = await HTMLtoDOCX(`<!DOCTYPE html><html><body>${html}</body></html>`, null, {
  table: { row: { cantSplit: true } },
  footer: true,
  pageNumber: true,
});
await writeFile(output, docx);
await writeFile(`${output}.warnings.json`, JSON.stringify(messages.map((m) => `${m.type}: ${m.message}`), null, 2));
