// Builds the Wave 4 layout fixture from fixture D's content: section breaks of every kind (continuous,
// odd page in landscape, even page, next page) and a column break, to compare our pages with Word's.
// Usage: node make_wave4.mjs   (writes ../e1-docx-roundtrip/fixtures-wave4/H_sections.docx)
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "../../packages/docx/dist/index.js";

const fixtures = new URL("../e1-docx-roundtrip/", import.meta.url);
const doc = readDocx(readFileSync(new URL("fixtures/D_complex_report.docx", fixtures))).doc;
const s0 = doc.sections[0];
const blocks = doc.sections.flatMap((s) => s.blocks);
const cut = (a, b) => structuredClone(blocks.slice(a, b));
const portrait = { ...s0.props };
const landscape = { ...s0.props, pageWidth: s0.props.pageHeight ?? 792, pageHeight: s0.props.pageWidth ?? 612, orientation: "landscape" };
const n = blocks.length;
const at = (f) => Math.round(n * f);
const sections = [
  { ...s0, props: { ...portrait }, blocks: cut(0, at(0.15)) },
  { props: { ...portrait, breakType: "continuous" }, headers: {}, footers: {}, blocks: cut(at(0.15), at(0.3)) },
  { props: { ...landscape, breakType: "oddPage" }, headers: {}, footers: {}, blocks: cut(at(0.3), at(0.5)) },
  { props: { ...portrait, breakType: "evenPage" }, headers: {}, footers: {}, blocks: cut(at(0.5), at(0.7)) },
  { props: { ...portrait }, headers: {}, footers: {}, blocks: cut(at(0.7), n) },
];
// A column break in the middle of a long paragraph of the last section (a page break without columns).
const last = sections[4].blocks;
const para = last.find((b) => b.type === "paragraph" && b.content.filter((i) => i.type === "text").reduce((s, i) => s + i.text.length, 0) > 150);
const run = para.content.findIndex((i) => i.type === "text" && i.text.length > 40);
const text = para.content[run];
para.content.splice(run, 1, { ...text, text: text.text.slice(0, 20) }, { type: "break", kind: "column" }, { ...text, text: text.text.slice(20) });
doc.sections = sections;
writeFileSync(new URL("fixtures-wave4/H_sections.docx", fixtures), writeDocx(doc));
console.log("sections:", sections.map((s) => `${s.props.breakType ?? "nextPage"}:${s.blocks.length}`).join(" "));
