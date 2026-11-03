import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
const { doc } = readDocx(readFileSync(process.argv[2]));
const r = layoutDocument(doc, new SystemFontMeasurer());
const out = { pages: r.pages.length, foot: [], end: [] };
for (const p of r.pages) for (const l of p.lines) {
  const t = l.segments.map((s) => s.text).join("");
  if (l.story === "footnote" && /Note \d+:/.test(t)) out.foot.push([p.number, Math.round(l.y * 10) / 10]);
  if (l.story === "endnote" && /Endnote at/.test(t)) out.end.push([p.number, Math.round(l.y * 10) / 10]);
}
console.log(JSON.stringify(out));
