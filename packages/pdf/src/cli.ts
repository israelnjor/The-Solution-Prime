#!/usr/bin/env node
/** prime-pdf <in.docx> <out.pdf> — read a Word document, lay it out, and export it as PDF. */
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
import { renderPdf } from "./index.js";

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error("usage: prime-pdf <in.docx> <out.pdf>");
  process.exit(2);
}
const { doc } = readDocx(readFileSync(input));
const fonts = new SystemFontMeasurer();
const layout = layoutDocument(doc, fonts);
writeFileSync(output, await renderPdf(doc, layout, fonts));
console.log(`${output}: ${layout.pages.length} pages`);
