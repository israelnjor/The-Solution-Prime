#!/usr/bin/env node
/**
 * prime-editor roundtrip <in.docx> <out.docx>
 *   DOCX -> document model -> editor document (serialised to JSON and back) -> model -> DOCX.
 *   Gate E3: opening a document in the editor must not change it.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import type { PrimeDocument } from "@prime/pdm";
import { fromProseMirror, toProseMirror } from "./convert.js";
import { createSchema } from "./schema.js";

const [command, input, output] = process.argv.slice(2);
if (command !== "roundtrip" || !input || !output) {
  console.error("usage: prime-editor roundtrip <in.docx> <out.docx>");
  process.exit(2);
}
const { doc, warnings } = readDocx(readFileSync(input));
const schema = createSchema();
const pm = schema.nodeFromJSON(JSON.parse(JSON.stringify(toProseMirror(doc, schema).toJSON())));
pm.check();
const back = JSON.parse(JSON.stringify(fromProseMirror(pm, doc))) as PrimeDocument;
writeFileSync(output, writeDocx(back));
writeFileSync(`${output}.warnings.json`, JSON.stringify(warnings, null, 2));
