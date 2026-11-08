/**
 * Basic grammar checking, offline: a few rules that are rarely wrong (a false alarm costs more trust
 * than a missed one). Each finds a stretch of a paragraph's text and offers what to put instead.
 * Rules: a word repeated ("the the"), a / an before the wrong sound, a sentence starting in lower case,
 * "i" for "I", a space before or none after a comma or full stop, two spaces, and a handful of common
 * confusions ("could of", "more then", "its a", "your welcome").
 */

export interface GrammarIssue {
  index: number;
  length: number;
  rule: string;
  message: string;
  replacements: string[];
}

/** Words that may stand twice in a row ("had had", "that that"). */
const MAY_REPEAT = new Set(["had", "that"]);
/** Vowel letters sounding like a consonant: "a university", "a one-off", "a European". */
const CONSONANT_SOUND = /^(uni(?!nt|mp|nd|nh|ns|nf)|use|usu|uti|ura|ure|uro|eu|ewe|one\b|once|ubiq|uga|uku)/i;
/** A silent h: "an hour", "an honest". */
const VOWEL_SOUND = /^(hour|honest|honou?r|heir|herb\b)/i;
/** Abbreviations after which a lower-case word is no new sentence. */
const ABBREVIATIONS = new Set(["e.g", "i.e", "etc", "vs", "cf", "al", "approx", "fig", "no", "vol", "p", "pp", "ch", "sec", "dept", "est", "inc", "ltd", "co", "mr", "mrs", "ms", "dr", "prof", "st", "jr", "sr", "a.m", "p.m", "viz", "ca"]);

const CONFUSIONS: { re: RegExp; message: string; fix: (m: RegExpExecArray) => string }[] = [
  { re: /\b(could|should|would|must|might)\s+of\b/giu, message: "Did you mean \"have\"?", fix: (m) => `${m[1]} have` },
  { re: /\b(more|less|rather|better|worse|other|greater|fewer|larger|smaller|higher|lower)\s+then\b/giu, message: "Did you mean \"than\" (a comparison)?", fix: (m) => `${m[1]} than` },
  { re: /\bits\s+(a|an|the|not|been|going|time|true|important|clear|possible)\b/giu, message: "Did you mean \"it's\" (it is)?", fix: (m) => `${m[0][0] === "I" ? "It" : "it"}'s ${m[1]}` },
  { re: /\byour\s+(welcome|right|wrong|not|going|being)\b/giu, message: "Did you mean \"you're\" (you are)?", fix: (m) => `${m[0][0] === "Y" ? "You" : "you"}'re ${m[1]}` },
  { re: /\bthere\s+(own|selves)\b/giu, message: "Did you mean \"their\"?", fix: (m) => `their ${m[1]}` },
];

/** The grammar issues in a paragraph's text. */
export function grammarIssues(text: string): GrammarIssue[] {
  const out: GrammarIssue[] = [];
  const add = (index: number, length: number, rule: string, message: string, replacements: string[]) => {
    if (!out.some((o) => index < o.index + o.length && o.index < index + length)) out.push({ index, length, rule, message, replacements });
  };

  // A word twice in a row.
  for (const m of text.matchAll(/\b(\p{L}+)(\s+)(\1)\b/giu)) {
    if (MAY_REPEAT.has(m[1].toLowerCase()) || m[1].length < 2 && !/^[aI]$/i.test(m[1])) continue;
    if (m[1] !== m[3] && m[1].toLowerCase() !== m[3].toLowerCase()) continue;
    add(m.index!, m[0].length, "repeat", `"${m[1]}" twice in a row.`, [m[1]]);
  }
  // a / an.
  for (const m of text.matchAll(/\b(a|an|A|An)\s+([\p{L}][\p{L}'’-]*)/gu)) {
    const word = m[2];
    if (word === word.toUpperCase() && word.length > 1) continue; // acronyms: "a UN report", "an FBI agent"
    if (/^[A-Z]$/.test(word)) continue; // "a B grade"
    const vowelSound = (/^[aeiou]/i.test(word) && !CONSONANT_SOUND.test(word)) || VOWEL_SOUND.test(word);
    const article = m[1].toLowerCase();
    if (article === "a" && vowelSound) add(m.index!, m[1].length, "article", `Use "an" before a vowel sound ("${word}").`, [m[1] === "A" ? "An" : "an"]);
    if (article === "an" && !vowelSound) add(m.index!, m[1].length, "article", `Use "a" before a consonant sound ("${word}").`, [m[1] === "An" ? "A" : "a"]);
  }
  // A sentence starting in lower case (after a full stop, ! or ?).
  for (const m of text.matchAll(/([\p{L}\d)'"’”]+)([.!?])(\s+)(\p{Ll}[\p{L}'’]*)/gu)) {
    const before = m[1].toLowerCase().replace(/^[("']+/, "");
    // (An abbreviation, or a single letter: "e.g. this", "U.S. army".)
    if (m[2] === "." && (ABBREVIATIONS.has(before) || before.length < 2)) continue;
    if (/\d$/.test(before) && m[2] === ".") continue; // "version 2. then" — numbering
    const word = m[4];
    if (/^(e\.g|i\.e|etc|vs)$/i.test(word)) continue;
    const start = m.index! + m[1].length + m[2].length + m[3].length;
    add(start, word.length, "capital", "A sentence starts with a capital letter.", [word[0].toUpperCase() + word.slice(1)]);
  }
  // "i" for "I".
  for (const m of text.matchAll(/(?<![\p{L}'’.-])i(?=(?:['’](?:m|ve|d|ll))?(?![\p{L}.-]))/gu)) {
    add(m.index!, 1, "i", "\"I\" is always a capital letter.", ["I"]);
  }
  // A space before a comma, full stop, semicolon, colon, ! or ?.
  for (const m of text.matchAll(/(\p{L})( +)([,.;:!?])(?=\s|$)/gu)) {
    if (m[3] === "." && text.slice(m.index! + m[0].length - 1, m.index! + m[0].length + 2) === "...") continue;
    add(m.index! + 1, m[2].length + 1, "space-before", `No space before "${m[3]}".`, [m[3]]);
  }
  // No space after a comma between words.
  for (const m of text.matchAll(/(\p{L}),(\p{L}{2,})/gu)) add(m.index! + 1, 1, "space-after", "A space after a comma.", [", "]);
  // Two or more spaces between words.
  for (const m of text.matchAll(/(?<=\S)( {2,})(?=\S)/g)) add(m.index!, m[1].length, "spaces", "Two spaces between words.", [" "]);
  // Common confusions.
  for (const c of CONFUSIONS) for (const m of text.matchAll(c.re)) add(m.index!, m[0].length, "confusion", c.message, [c.fix(m as RegExpExecArray)]);
  return out.sort((a, b) => a.index - b.index);
}
