#!/usr/bin/env node
/** PDF → Word from the command line: prime-pdf-import in.pdf out.docx */
import { readFileSync, writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import { convertPdf } from "./index.js";

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error("Usage: prime-pdf-import <in.pdf> <out.docx>");
  process.exit(2);
}
const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
const { doc, warnings } = await convertPdf(new Uint8Array(readFileSync(input)), pdfjs as never);
writeFileSync(output, writeDocx(doc));
console.log(`${output}: ${doc.sections[0].blocks.length} blocks`);
for (const w of warnings) console.log(`warning: ${w}`);
