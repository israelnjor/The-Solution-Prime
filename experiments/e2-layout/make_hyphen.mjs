// Builds the Wave 4 hyphenation fixture: fixture D with automatic hyphenation on (Word's defaults:
// zone 0.25", no limit), to compare our line breaks with Word's.
// Usage: node make_hyphen.mjs   (writes ../e1-docx-roundtrip/fixtures-wave4/K_hyphen.docx)
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "../../packages/docx/dist/index.js";

const fixtures = new URL("../e1-docx-roundtrip/", import.meta.url);
const doc = readDocx(readFileSync(new URL("fixtures/D_complex_report.docx", fixtures))).doc;
doc.hyphenation = { auto: true };
writeFileSync(new URL("fixtures-wave4/K_hyphen.docx", fixtures), writeDocx(doc));
console.log("ok");
