// Copies pdf.js's support files (character maps, standard fonts, image decoders, ICC profiles) into the
// app's public folder, so PDF → Word works offline. Run before `dev` and `build` of apps/word.
import { cpSync, existsSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const from = join(root, "node_modules", "pdfjs-dist");
const to = join(root, "apps", "word", "public", "pdfjs");
mkdirSync(to, { recursive: true });
for (const dir of ["cmaps", "standard_fonts", "wasm", "iccs"]) {
  if (existsSync(join(from, dir))) cpSync(join(from, dir), join(to, dir), { recursive: true });
}
cpSync(join(from, "LICENSE"), join(to, "LICENSE"));
console.log("pdf.js files copied");
