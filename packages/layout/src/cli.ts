#!/usr/bin/env node
/**
 * prime-layout — run the layout engine on a .docx.
 *
 *   prime-layout markers <in.docx>      JSON: page count and the position of every content marker
 *   prime-layout paragraphs <in.docx>   JSON: page count and the position of every paragraph
 *   prime-layout pages <in.docx>        JSON: full page/line/segment output
 */
import { readFileSync } from "node:fs";
import { readDocx } from "@prime/docx";
import { layoutDocument } from "./engine.js";
import { SystemFontMeasurer } from "./node/fonts.js";

const [command, input] = process.argv.slice(2);
if (!input || !["markers", "paragraphs", "pages"].includes(command)) {
  console.error("usage: prime-layout markers|paragraphs|pages <in.docx>");
  process.exit(2);
}
const { doc } = readDocx(readFileSync(input));
const result = layoutDocument(doc, new SystemFontMeasurer());
const pages = result.pages.length;
const out = command === "markers" ? { pages, markers: result.markers }
  : command === "paragraphs" ? { pages, paragraphs: result.paragraphs }
  : result;
process.stdout.write(JSON.stringify(out));
