// Letterheads in all three layouts, with a logo; Word verifies each opens with the header content.
// Usage: node letterhead.mjs <logo.png> <out-folder>
import { readFileSync, writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import { letterheadDocument } from "@prime/editor";

const [logoPath, out] = process.argv.slice(2);
const png = readFileSync(logoPath);
const w = png.readUInt32BE(16), h = png.readUInt32BE(20);
for (const layout of ["centre", "logoLeft", "logoRight"]) {
  const doc = letterheadDocument({
    name: "Accra Community Library", tagline: "Knowledge for everyone", layout, color: "1F3864", rule: true, footer: layout === "logoRight",
    address: ["12 Independence Avenue", "Accra, Ghana"], contacts: ["+233 30 222 0000", "info@acl.org.gh", "www.acl.org.gh"],
    logo: { data: png.toString("base64"), contentType: "image/png", width: w, height: h },
  });
  doc.sections[0].blocks = [{ type: "paragraph", props: {}, content: [{ type: "text", text: "Dear Sir/Madam," }] }];
  writeFileSync(`${out}/letterhead-${layout}.docx`, writeDocx(doc));
  console.log("wrote", layout);
}
