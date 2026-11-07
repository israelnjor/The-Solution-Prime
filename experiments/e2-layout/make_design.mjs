// Builds the Wave 4 Design fixture from fixture D: a text watermark (as Word writes it), a page colour
// and a double box border on every page, to check that Word shows them as its own.
// Usage: node make_design.mjs   (writes ../e1-docx-roundtrip/fixtures-wave4/I_design.docx)
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "../../packages/docx/dist/index.js";
import { watermarkXml } from "../../packages/editor/dist/design.js";

const fixtures = new URL("../e1-docx-roundtrip/", import.meta.url);
const doc = readDocx(readFileSync(new URL("fixtures/D_complex_report.docx", fixtures))).doc;
doc.background = { name: "w:background", xml: '<w:background w:color="FFF8E1"/>' };
const side = (n) => `<w:${n} w:val="double" w:sz="6" w:space="24" w:color="1F4E79"/>`;
const borders = { name: "w:pgBorders", xml: `<w:pgBorders w:offsetFrom="page">${["top", "left", "bottom", "right"].map(side).join("")}</w:pgBorders>` };
let id = 1;
for (const s of doc.sections) {
  s.props.extra = [...(s.props.extra ?? []).filter((x) => x.name !== "w:pgBorders"), borders];
  for (const story of Object.values(s.headers)) {
    const p = story.blocks.find((b) => b.type === "paragraph");
    p.content.unshift({ type: "runPreserved", preserved: { name: "w:pict", xml: watermarkXml({ text: "CONFIDENTIAL" }, id++) } });
  }
}
writeFileSync(new URL("fixtures-wave4/I_design.docx", fixtures), writeDocx(doc));
console.log("headers with a watermark:", id - 1);
