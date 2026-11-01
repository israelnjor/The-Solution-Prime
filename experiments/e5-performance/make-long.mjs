// Builds a long test document by repeating the fixtures' content (headings, lists, tables, images).
// Usage: node make-long.mjs <out.docx> [copies]
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "../../packages/docx/dist/index.js";

const [out, copies = "12"] = process.argv.slice(2);
const fixtures = new URL("../e1-docx-roundtrip/fixtures/", import.meta.url);
const base = readDocx(readFileSync(new URL("D_complex_report.docx", fixtures))).doc;
const extra = ["A_basic_formatting", "B_structured_content"].map((f) => readDocx(readFileSync(new URL(`${f}.docx`, fixtures))).doc);
const body = (d) => d.sections.flatMap((s) => s.blocks);
const unit = [...body(base), ...extra.flatMap((d) => d.styles === base.styles ? body(d) : body(d).filter((b) => b.type === "paragraph"))];
const last = base.sections.at(-1);
base.sections = [{ ...last, blocks: Array.from({ length: Number(copies) }, () => structuredClone(unit)).flat() }];
writeFileSync(out, writeDocx(base));
console.log(`blocks: ${base.sections[0].blocks.length}`);
