// Damaged and odd files: the reader must fail fast with a clear message, never crash or hang.
// Usage: node robust.mjs <good.docx>
import { readFileSync } from "node:fs";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { readDocx, writeDocx } from "@prime/docx";

const good = readFileSync(process.argv[2]);
const files = unzipSync(good);
const rezip = (change) => { const f = { ...files }; change(f); return zipSync(f); };
const cases = {
  "empty file": new Uint8Array(0),
  "plain text": strToU8("Hello, this is not a Word document."),
  "PDF renamed .docx": strToU8("%PDF-1.7\n1 0 obj << >> endobj\n%%EOF"),
  "old .doc (binary)": Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Array(500).fill(0)]),
  "truncated download": good.subarray(0, Math.floor(good.length / 2)),
  "zip without the document": rezip((f) => { delete f["word/document.xml"]; }),
  "malformed document XML": rezip((f) => { f["word/document.xml"] = strToU8("<w:document><w:body><w:p>"); }),
  "no styles part": rezip((f) => { delete f["word/styles.xml"]; }),
  "list without numbering part": rezip((f) => { delete f["word/numbering.xml"]; }),
  "no content types": rezip((f) => { delete f["[Content_Types].xml"]; }),
};
let bad = 0;
for (const [name, bytes] of Object.entries(cases)) {
  const t0 = Date.now();
  try {
    const { doc, warnings } = readDocx(bytes);
    // If it opens, it must also save.
    writeDocx(doc);
    console.log(`opens   ${name} (${Date.now() - t0}ms)${warnings.length ? `: ${warnings.slice(0, 2).join("; ")}` : ""}`);
  } catch (e) {
    const msg = String(e?.message ?? e);
    const clear = msg.length < 200 && !/undefined|null|cannot read|is not a function/i.test(msg);
    if (!clear) bad++;
    console.log(`refused ${name} (${Date.now() - t0}ms): ${msg}${clear ? "" : "   <-- unclear message"}`);
  }
}
process.exitCode = bad ? 1 : 0;
