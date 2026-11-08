// Usage: node scripts/build-thesaurus.mjs <WordNet 3.0 dict folder> node_modules/dictionary-en/index.dic apps/word/public/thesaurus
// Builds Prime Word's thesaurus (apps/word/public/thesaurus/) from WordNet 3.0's data files: synonym
// sets that hold a word of the English dictionary (adjectives with their "similar to" words merged in),
// with antonyms. Reads the WordNet files as data only.
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
const [wnDir, dicPath, outDir] = process.argv.slice(2);
const common = new Set(readFileSync(dicPath, "utf8").split(/\r?\n/).slice(1).map((l) => l.split("/")[0].toLowerCase()).filter((w) => /^[a-z][a-z'-]*$/.test(w)));
const POS = { noun: "n", verb: "v", adj: "a", adv: "r" };
const synsets = new Map(); // "pos:offset" -> { pos, words, similar: [], antonyms: [] }
for (const [file, pos] of Object.entries(POS)) {
  for (const line of readFileSync(`${wnDir}/data.${file}`, "latin1").split("\n")) {
    if (!line || line.startsWith("  ")) continue;
    const parts = line.split(" | ")[0].split(" ");
    const offset = parts[0];
    const type = parts[2];
    const wcnt = parseInt(parts[3], 16);
    const words = [];
    let i = 4;
    for (let k = 0; k < wcnt; k++, i += 2) words.push(parts[i].replace(/\(.*\)$/, "").replace(/_/g, " "));
    const pcnt = Number(parts[i]); i++;
    const similar = [];
    const antonyms = [];
    for (let k = 0; k < pcnt; k++, i += 4) {
      const [sym, off, p] = [parts[i], parts[i + 1], parts[i + 2] === "s" ? "a" : parts[i + 2]];
      if (sym === "&") similar.push(`${p}:${off}`);
      if (sym === "!") antonyms.push(`${p}:${off}`);
    }
    synsets.set(`${pos}:${offset}`, { pos, type, words, similar, antonyms });
  }
}
const isCommon = (w) => common.has(w.toLowerCase());
// Each word's meanings, most used first (the index files list a word's synsets in that order).
const rank = new Map(); // "pos:word" -> offsets
for (const [file, pos] of Object.entries(POS)) {
  for (const line of readFileSync(`${wnDir}/index.${file}`, "latin1").split("\n")) {
    if (!line || line.startsWith("  ")) continue;
    const parts = line.trim().split(" ");
    const cnt = Number(parts[2]);
    rank.set(`${pos}:${parts[0].replace(/_/g, " ")}`, parts.slice(parts.length - cnt));
  }
}
/** A word with how used this meaning of it is (":3": its fourth meaning; none: its first, or not its own). */
const ranked = (pos, offset, w, own) => { if (!own) return `${w}:9`; const r = (rank.get(`${pos}:${w.toLowerCase()}`) ?? []).indexOf(offset); return r > 0 ? `${w}:${Math.min(r, 9)}` : w; };
const out = [];
const keep = new Map(); // key -> index in out
for (const [key, s] of synsets) {
  // A satellite adjective is shown under its head (merged there).
  if (s.type === "s") continue;
  let words = [...s.words];
  if (s.pos === "a") for (const sim of s.similar) words.push(...(synsets.get(sim)?.words ?? []));
  words = [...new Set(words)];
  if (words.length < 2 && !s.antonyms.length) continue;
  if (!words.some(isCommon)) continue;
  keep.set(key, out.length);
  const off = key.split(":")[1];
  out.push({ pos: s.pos, words: words.map((w) => ranked(s.pos, off, w, s.words.includes(w))), antonyms: s.antonyms });
}
// Satellites become groups of their own too (the word "capable" looks up its own meaning, with its head's words).
for (const [key, s] of synsets) {
  if (s.type !== "s") continue;
  const head = s.similar.map((k) => synsets.get(k)).find(Boolean);
  const words = [...new Set([...s.words, ...(head?.words ?? [])])];
  if (words.length < 2 || !s.words.some(isCommon)) continue;
  keep.set(key, out.length);
  const off = key.split(":")[1];
  out.push({ pos: "a", words: words.map((w) => ranked("a", off, w, s.words.includes(w))), antonyms: head?.antonyms ?? [] });
}
const lines = out.map((s) => {
  const ants = [...new Set(s.antonyms.flatMap((k) => synsets.get(k)?.words.slice(0, 2) ?? []))];
  return `${s.pos}|${s.words.join(",")}${ants.length ? `|${ants.join(",")}` : ""}`;
});
mkdirSync(outDir, { recursive: true });
writeFileSync(`${outDir}/thesaurus.txt`, lines.join("\n") + "\n");
console.log("groups", lines.length, "bytes", Buffer.byteLength(lines.join("\n")));
