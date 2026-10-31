// Copies the spelling dictionaries (Hunspell .aff/.dic + licence) into the app's public folder, so they
// ship with Prime Word and load offline. Run automatically before `dev` and `build` of apps/word.
import { copyFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
for (const [pkg, lang] of [["dictionary-en", "en-US"], ["dictionary-en-gb", "en-GB"]]) {
  const from = join(root, "node_modules", pkg);
  const to = join(root, "apps", "word", "public", "dictionaries", lang);
  mkdirSync(to, { recursive: true });
  for (const f of ["index.aff", "index.dic", "license"]) copyFileSync(join(from, f), join(to, f));
}
console.log("dictionaries copied");
