/**
 * AutoCorrect and AutoFormat as you type, as in Word: curly quotes, dashes, symbols ((c) → ©, --> → →),
 * common misspellings, a capital at the start of a sentence, "TWo INitial CApitals", and automatic
 * lists ("1. " or "- " at the start of a paragraph). Each change is its own undo step, made after the
 * typed character: Ctrl+Z, or Backspace straight away, takes back only the change.
 */
import type { PrimeDocument } from "@prime/pdm";
import { closeHistory, undo } from "prosemirror-history";
import { Plugin, PluginKey, type Transaction } from "prosemirror-state";
import { ensureList, type ListKind } from "./lists.js";

export interface AutoFormatOptions {
  quotes: boolean;
  dashes: boolean;
  symbols: boolean;
  corrections: boolean;
  capitalise: boolean;
  twoCapitals: boolean;
  lists: boolean;
  /** The person's own replacements ("pw" → "Prime Word"), applied when the text ends. */
  custom?: Record<string, string>;
}
export const AUTOFORMAT_DEFAULTS: AutoFormatOptions = {
  quotes: true, dashes: true, symbols: true, corrections: true, capitalise: true, twoCapitals: true, lists: true,
};

/** Replaced as soon as the last character is typed. */
const SYMBOLS: [string, string][] = [
  ["(c)", "©"], ["(C)", "©"], ["(r)", "®"], ["(R)", "®"], ["(tm)", "™"], ["(TM)", "™"], ["...", "…"],
  ["-->", "→"], ["<--", "←"], ["==>", "⇒"], ["<==", "⇐"], ["<=>", "⇔"],
];

/** Common misspellings (lower case), corrected when the word ends. */
export const CORRECTIONS: Record<string, string> = {
  teh: "the", hte: "the", adn: "and", nad: "and", taht: "that", thier: "their", waht: "what", wiht: "with",
  yuor: "your", jsut: "just", abotu: "about", becuase: "because", becasue: "because", recieve: "receive",
  recieved: "received", seperate: "separate", definately: "definitely", occured: "occurred",
  occurence: "occurrence", untill: "until", wich: "which", beleive: "believe", accomodate: "accommodate",
  acheive: "achieve", alot: "a lot", begining: "beginning", calender: "calendar", comming: "coming",
  existance: "existence", freind: "friend", goverment: "government", happend: "happened",
  independant: "independent", knowlege: "knowledge", neccessary: "necessary", necesary: "necessary",
  occassion: "occasion", posession: "possession", publically: "publicly", recomend: "recommend",
  refered: "referred", sucess: "success", succesful: "successful", tommorow: "tomorrow",
  tomorow: "tomorrow", truely: "truly", wierd: "weird", realy: "really", allready: "already",
  arguement: "argument", enviroment: "environment", foward: "forward", relevent: "relevant",
  responsability: "responsibility", supercede: "supersede", tounge: "tongue", dont: "don't",
  doesnt: "doesn't", didnt: "didn't", cant: "can't", wont: "won't", isnt: "isn't", wasnt: "wasn't",
  couldnt: "couldn't", shouldnt: "shouldn't", wouldnt: "wouldn't", havent: "haven't", i: "I",
};

/** Abbreviations after which the next word does not start a sentence. */
const ABBREVIATIONS = new Set(["e.g.", "i.e.", "etc.", "vs.", "approx.", "dr.", "mr.", "mrs.", "ms.", "prof.", "no.", "fig.", "cf.", "st.", "ave.", "jan.", "feb.", "aug.", "sept.", "oct.", "nov.", "dec.", "p.", "pp.", "vol.", "ed.", "et al."]);

const BOUNDARY = /[\s.,;:!?)\]}"”’]/;
const OPENS_QUOTE = /[\s([{—–"“‘'-]/;

/**
 * Adds to `tr` the automatic change after `typed` was typed, the cursor now at `cursor` in `tr.doc`;
 * false if there is none. `doc` gives the document, for automatic lists.
 */
export function autoFormatAfter(tr: Transaction, cursor: number, typed: string, opts: AutoFormatOptions, doc?: () => PrimeDocument): boolean {
  const $c = tr.doc.resolve(cursor);
  const para = $c.parent;
  if (!para.isTextblock) return false;
  const start = cursor - $c.parentOffset;
  const text = para.textBetween(0, $c.parentOffset, undefined, "\ufffc"); // up to the cursor, typed character included
  const before = text.slice(0, -typed.length);
  const replace = (from: number, to: number, by: string) => { tr.insertText(by, start + from, start + to); return true; };

  // Curly quotes and apostrophes.
  if (opts.quotes && (typed === '"' || typed === "'")) {
    const prev = before.slice(-1);
    const open = !prev || OPENS_QUOTE.test(prev);
    return replace(text.length - 1, text.length, typed === '"' ? (open ? "“" : "”") : open ? "‘" : "’");
  }

  // Symbols, replaced as soon as they are complete.
  if (opts.symbols) {
    const hit = SYMBOLS.find(([k]) => text.endsWith(k));
    if (hit) return replace(text.length - hit[0].length, text.length, hit[1]);
  }

  // Lists: "1. ", "1) ", "* ", "- " or "• " at the start of a paragraph that is not in a list yet.
  if (opts.lists && doc && typed === " " && para.type.name === "paragraph" && !para.attrs.props?.list) {
    const kind: ListKind | undefined = /^1[.)] $/.test(text) ? "number" : /^[*\-•] $/.test(text) ? "bullet" : undefined;
    if (kind) {
      tr.delete(start, cursor);
      const props = { ...para.attrs.props, list: { numId: ensureList(doc(), kind), level: 0 } };
      tr.setNodeMarkup(start - 1, undefined, { ...para.attrs, props });
      return true;
    }
  }

  if (!BOUNDARY.test(typed)) return false;

  // Dashes: "word -- word" → en dash, "word--word" → em dash (once the word after it ends).
  if (opts.dashes) {
    const m = /(\S)( ?)--( ?)([^\s-]+)$/.exec(before);
    if (m && m[1] !== "-") {
      const at = before.length - m[0].length + 1 + m[2].length;
      return replace(at, at + 2, m[2] && m[3] ? "–" : "—");
    }
  }

  // The person's own replacements.
  for (const [k, v] of Object.entries(opts.custom ?? {})) {
    if (!k || !before.endsWith(k)) continue;
    const at = before.length - k.length;
    if (at === 0 || /[\s(\["“‘]/.test(before[at - 1])) return replace(at, before.length, v);
  }

  // The word just finished.
  const m = /(^|[\s(\["“‘])([A-Za-z][A-Za-z']*)$/.exec(before);
  if (!m) return false;
  const word = m[2];
  const wordStart = before.length - word.length;
  let fixed = word;
  if (opts.corrections) {
    const c = CORRECTIONS[word.toLowerCase()];
    if (c && (word === word.toLowerCase() || (word[0] === word[0].toUpperCase() && word.slice(1) === word.slice(1).toLowerCase()))) {
      fixed = word[0] === word[0].toUpperCase() ? c[0].toUpperCase() + c.slice(1) : c;
    }
  }
  if (opts.twoCapitals && /^[A-Z]{2}[a-z]+$/.test(fixed) && fixed.length >= 3) fixed = fixed[0] + fixed.slice(1).toLowerCase();
  if (opts.capitalise && /^[a-z]/.test(fixed) && startsSentence(before.slice(0, wordStart))) fixed = fixed[0].toUpperCase() + fixed.slice(1);
  return fixed !== word && replace(wordStart, before.length, fixed);
}

/** Whether text that comes before a word ends a sentence (or is the start of the paragraph). */
function startsSentence(lead: string): boolean {
  const t = lead.replace(/[\s"“‘(\[]+$/, "");
  if (!t) return true; // the start of the paragraph
  if (!/[.!?]$/.test(t) || !/\s$/.test(lead)) return false;
  const last = /(\S+)$/.exec(t)![1].toLowerCase();
  return !(ABBREVIATIONS.has(last) || /^[a-z]\.$/.test(last) || /\.\S+\./.test(last)); // not after "e.g.", "U.S."
}

const key = new PluginKey<boolean>("autoformat");

/** The AutoFormat plugin: list it before the keymaps (it handles Backspace straight after a change). */
export function autoFormat(options: () => AutoFormatOptions, doc?: () => PrimeDocument): Plugin<boolean> {
  return new Plugin<boolean>({
    key,
    state: {
      init: () => false,
      // True straight after an automatic change, until anything else happens.
      apply: (tr, was) => (tr.getMeta(key) ? true : tr.docChanged || tr.selectionSet ? false : was),
    },
    props: {
      handleTextInput(view, from, to, text) {
        if (text.length !== 1 || view.composing) return false;
        const after = key.getState(view.state);
        const typed = view.state.tr.insertText(text, from, to);
        view.dispatch(after ? closeHistory(typed) : typed); // what is typed after a change is not undone with it
        const change = view.state.tr;
        if (autoFormatAfter(change, from + text.length, text, options(), doc)) view.dispatch(closeHistory(change).setMeta(key, true));
        return true;
      },
      handleKeyDown(view, event) {
        if (event.key === "Backspace" && key.getState(view.state) && !event.ctrlKey && !event.altKey) {
          return undo(view.state, view.dispatch);
        }
        if (event.key === "Enter" && !event.shiftKey && view.state.selection.empty) {
          // Enter ends a word too: correct it (with a stand-in space, taken out again), then Enter goes on.
          const { from } = view.state.selection;
          const tr = view.state.tr.insertText(" ", from);
          if (autoFormatAfter(tr, from + 1, " ", { ...options(), lists: false })) {
            const space = tr.mapping.slice(1).map(from + 1);
            view.dispatch(closeHistory(tr.delete(space - 1, space)));
          }
        }
        return false;
      },
    },
  });
}
