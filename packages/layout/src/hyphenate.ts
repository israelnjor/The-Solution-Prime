/**
 * Where a word may be hyphenated (Liang's algorithm with US English patterns, as TeX does). Word uses
 * its own hyphenator, which agrees on most words.
 */
import { EXCEPTIONS, PATTERNS } from "./hyphen-en.js";

/** Pattern letters → the digits between them (".ach4" → "ach": [0, 0, 0, 4]). */
let patterns: Map<string, number[]> | undefined;
let exceptions: Map<string, number[]> | undefined;

function load() {
  patterns = new Map();
  for (const p of PATTERNS.split(" ")) {
    const letters = p.replace(/\d/g, "");
    const digits: number[] = [0];
    for (const ch of p) {
      if (/\d/.test(ch)) digits[digits.length - 1] = Number(ch);
      else digits.push(0);
    }
    patterns.set(letters, digits);
  }
  exceptions = new Map();
  for (const e of EXCEPTIONS.split(" ")) {
    const points: number[] = [];
    let n = 0;
    for (const ch of e) {
      if (ch === "-") points.push(n);
      else n++;
    }
    exceptions.set(e.replace(/-/g, ""), points);
  }
}

const LEFT = 2; // letters kept before a hyphen
const RIGHT = 3; // and after it
const cache = new Map<string, number[]>();

/** The positions (letter counts) after which `word` may break, e.g. "hyphenation" → [2, 6] (hy-phen-ation). */
export function hyphenPoints(word: string): number[] {
  if (word.length < LEFT + RIGHT) return [];
  const lower = word.toLowerCase();
  if (!/^[a-z]+$/.test(lower)) return [];
  const known = cache.get(lower);
  if (known) return known;
  if (!patterns) load();
  const exception = exceptions!.get(lower);
  if (exception) { cache.set(lower, exception); return exception; }
  const text = `.${lower}.`;
  const score = new Array<number>(text.length + 1).fill(0);
  for (let i = 0; i < text.length; i++) {
    for (let j = i + 1; j <= text.length; j++) {
      const digits = patterns!.get(text.slice(i, j));
      if (!digits) continue;
      for (let k = 0; k < digits.length; k++) score[i + k] = Math.max(score[i + k], digits[k]);
    }
  }
  // score[i + 1] is the value between letters i-1 and i of the word (the dot shifts by one).
  const points: number[] = [];
  for (let n = LEFT; n <= lower.length - RIGHT; n++) if (score[n + 1] % 2 === 1) points.push(n);
  cache.set(lower, points);
  return points;
}
