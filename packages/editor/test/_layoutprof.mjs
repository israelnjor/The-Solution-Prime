import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";
const { doc } = readDocx(readFileSync(process.argv[2]));
const m = new SystemFontMeasurer();
for (let i = 0; i < 3; i++) {
  const t = performance.now();
  const r = layoutDocument(doc, m);
  console.log(`layout ${i + 1}: ${Math.round(performance.now() - t)} ms, ${r.pages.length} pages`);
}
