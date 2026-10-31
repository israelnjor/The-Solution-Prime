#!/usr/bin/env node
/**
 * prime-docx — command-line access to the DOCX engine.
 *
 *   prime-docx roundtrip <in.docx> <out.docx>   read -> PDM (JSON round trip) -> write; warnings to <out>.warnings.json
 *   prime-docx dump <in.docx> [model.json]       write the PDM model as JSON
 */
import { readFileSync, writeFileSync } from "node:fs";
import type { PrimeDocument } from "@prime/pdm";
import { readDocx } from "./reader.js";
import { writeDocx } from "./writer.js";

const [command, input, output] = process.argv.slice(2);

function main(): number {
  if (command === "roundtrip" && input && output) {
    const { doc, warnings } = readDocx(readFileSync(input));
    // Through JSON on purpose: proves the model is a complete, serialisable snapshot (autosave format).
    const snapshot = JSON.parse(JSON.stringify(doc)) as PrimeDocument;
    writeFileSync(output, writeDocx(snapshot));
    writeFileSync(`${output}.warnings.json`, JSON.stringify(warnings, null, 2));
    for (const w of warnings) console.error(`warning: ${w}`);
    return 0;
  }
  if (command === "dump" && input) {
    const { doc, warnings } = readDocx(readFileSync(input));
    const json = JSON.stringify({ warnings, doc }, (k, v) => (k === "data" ? `<${v.length} base64 chars>` : v), 2);
    if (output) writeFileSync(output, json);
    else console.log(json);
    return 0;
  }
  console.error("usage: prime-docx roundtrip <in.docx> <out.docx> | prime-docx dump <in.docx> [out.json]");
  return 2;
}

process.exit(main());
