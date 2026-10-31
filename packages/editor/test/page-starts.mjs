// Checks that every engine page start lands on the right text in the editor document.
// Usage: node page-starts.mjs <in.docx>
import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { createSchema, resolvePageStarts, toProseMirror } from "@prime/editor";
import { layoutDocument } from "@prime/layout";
import { SystemFontMeasurer } from "@prime/layout/dist/node/fonts.js";

const { doc } = readDocx(readFileSync(process.argv[2]));
const layout = layoutDocument(doc, new SystemFontMeasurer());
const pm = toProseMirror(doc, createSchema());
let ok = 0;
const starts = resolvePageStarts(pm, layout.pageStarts);
for (const s of starts) {
  const page = layout.pages[s.page - 1];
  const firstBody = page.lines.find((l) => !l.story);
  const engineText = (firstBody?.segments ?? []).map((x) => x.text).join("").trim().slice(0, 24);
  const end = Math.min(pm.content.size, s.pos + 60);
  const editorText = pm.textBetween(s.kind === "row" ? s.pos : s.pos, end, " ", " ").replace(/\s+/g, " ").trim().slice(0, 24);
  // The editor text at the page start must appear on that page in the engine's layout (the first
  // line may be a repeated table header row; picture-only lines have no text).
  const pageText = page.lines.filter((l) => !l.story).map((l) => l.segments.map((x) => x.text).join("")).join(" ").replace(/\s+/g, " ");
  const probe = editorText.slice(0, 12);
  const match = probe === "" || pageText.includes(probe) || engineText === "";
  if (match) ok++;
  console.log(`page ${s.page} ${s.kind.padEnd(6)} engine «${engineText}»  editor «${editorText}» ${match ? "✓" : "✗"}`);
}
console.log(`${ok}/${starts.length} page starts match (engine pages: ${layout.pages.length})`);
