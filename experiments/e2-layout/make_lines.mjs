// Builds the Wave 4 line-numbering fixture from fixture C (long text over 5 pages): numbers every
// line, restarting on each page, starting at 5 (w:start="4"); one paragraph left out.
// Usage: node make_lines.mjs   (writes ../e1-docx-roundtrip/fixtures-wave4/J_lines.docx)
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "../../packages/docx/dist/index.js";

const fixtures = new URL("../e1-docx-roundtrip/", import.meta.url);
const doc = readDocx(readFileSync(new URL("fixtures/C_multipage_layout.docx", fixtures))).doc;
for (const s of doc.sections) {
  s.props.extra = [...(s.props.extra ?? []).filter((x) => x.name !== "w:lnNumType"), { name: "w:lnNumType", xml: '<w:lnNumType w:countBy="1" w:start="4"/>' }];
}
const para = doc.sections[0].blocks.filter((b) => b.type === "paragraph")[3];
para.props = { ...para.props, extra: [...(para.props.extra ?? []), { name: "w:suppressLineNumbers", xml: "<w:suppressLineNumbers/>" }] };
writeFileSync(new URL("fixtures-wave4/J_lines.docx", fixtures), writeDocx(doc));
console.log("ok");
