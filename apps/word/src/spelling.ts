/**
 * Spell checking, offline: Hunspell dictionaries (SCOWL, en-GB / en-US) shipped with the app, checked
 * with nspell. Misspelled words get a red wavy underline; right-click offers suggestions.
 */
import { timed } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import nspell from "nspell";
import type { Node } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";

export type Speller = ReturnType<typeof nspell>;

const PERSONAL = "prime-word:dictionary";

/** Words the user added ("Add to dictionary"), kept in this app's local storage. */
function personalWords(): string[] {
  try { return JSON.parse(localStorage.getItem(PERSONAL) ?? "[]"); } catch { return []; }
}

export function addToDictionary(speller: Speller, word: string) {
  speller.add(word);
  try { localStorage.setItem(PERSONAL, JSON.stringify([...new Set([...personalWords(), word])])); } catch { /* storage unavailable */ }
}

/** The dictionary for a document: US English when the document says so, otherwise British English. */
export function documentLanguage(doc: PrimeDocument): "en-US" | "en-GB" {
  const lang = doc.styles.defaults.run?.extra?.find((x) => x.name === "w:lang")?.xml ?? "";
  return /w:val="en-US"/i.test(lang) ? "en-US" : "en-GB";
}

const loaded = new Map<string, Promise<Speller>>();

export function loadSpeller(lang: "en-US" | "en-GB"): Promise<Speller> {
  let p = loaded.get(lang);
  if (!p) {
    p = Promise.all([
      fetch(`./dictionaries/${lang}/index.aff`).then((r) => r.text()),
      fetch(`./dictionaries/${lang}/index.dic`).then((r) => r.text()),
    ]).then(([aff, dic]) => {
      const speller = nspell(aff, dic);
      for (const w of personalWords()) speller.add(w);
      return speller;
    });
    loaded.set(lang, p);
  }
  return p;
}

export interface Misspelling {
  from: number;
  to: number;
  word: string;
}

const WORD = /\p{L}[\p{L}'’]*\p{L}|\p{L}/gu;

/** Misspelled words in the document (skipping acronyms, words with digits, web and e-mail addresses). */
export function findMisspellings(doc: Node, speller: Speller, ignored: Set<string>, cache: Map<string, boolean>,
  paragraphs?: WeakMap<Node, Misspelling[]>): Misspelling[] {
  const out: Misspelling[] = [];
  const correct = (w: string) => {
    let ok = cache.get(w);
    if (ok === undefined) {
      ok = speller.correct(w);
      cache.set(w, ok);
    }
    return ok;
  };
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // A paragraph checked before and unchanged since (the same node): its results, moved to where it is now.
    const known = paragraphs?.get(node);
    if (known) {
      for (const m of known) out.push({ from: pos + m.from, to: pos + m.to, word: m.word });
      return false;
    }
    const first = out.length;
    let text = "";
    const at: number[] = [];
    // Text not to check: links and fields (addresses, references), and web/e-mail addresses typed as text.
    const skip: boolean[] = [];
    node.forEach((child, offset) => {
      const s = child.isText ? child.text! : " ";
      const linked = child.marks.some((m) => m.type.name === "link" || m.type.name === "field");
      for (let i = 0; i < s.length; i++) { at.push(pos + 1 + offset + i); skip.push(linked); }
      text += s;
    });
    for (const m of text.matchAll(/\S*(?:https?:\/\/|www\.|@)\S*/giu)) {
      for (let i = m.index!; i < m.index! + m[0].length; i++) skip[i] = true;
    }
    for (const m of text.matchAll(WORD)) {
      const raw = m[0];
      const word = raw.replace(/’/g, "'").replace(/^'+|'+$/g, "");
      if (word.length < 2 || ignored.has(word) || skip[m.index!]) continue;
      if (word === word.toUpperCase()) continue; // acronyms: GDP, NASA
      if (!/^[\p{Script=Latin}']+$/u.test(word)) continue; // other scripts (Greek, Cyrillic…): not an English dictionary's call
      // Part of a file name, path or code (report.docx, C:\path, snake_case)?
      const before = text[m.index! - 1] ?? " ";
      const after = text[m.index! + raw.length] ?? " ";
      if (/[@./\\_:]/.test(before) || /[@\\/_]/.test(after) || (after === "." && /\p{L}/u.test(text[m.index! + raw.length + 1] ?? ""))) continue;
      if (correct(word) || correct(word.toLowerCase())) continue;
      out.push({ from: at[m.index!], to: at[m.index! + raw.length - 1] + 1, word });
    }
    paragraphs?.set(node, out.slice(first).map((m) => ({ from: m.from - pos, to: m.to - pos, word: m.word })));
    return false;
  });
  return out;
}

interface SpellState {
  issues: Misspelling[];
  decorations: DecorationSet;
}

export const spellKey = new PluginKey<SpellState>("spelling");

/** Underlines misspellings; re-checks the document shortly after each change. */
export function spellPlugin(getSpeller: () => Speller | undefined, ignored: Set<string>): Plugin<SpellState> {
  const cache = new Map<string, boolean>();
  // Each paragraph's misspellings (relative to it), until the dictionary or the ignored words change.
  let paragraphs = new WeakMap<Node, Misspelling[]>();
  const build = (doc: Node): SpellState => {
    const speller = getSpeller();
    if (!speller) return { issues: [], decorations: DecorationSet.empty };
    const issues = timed("spelling", () => findMisspellings(doc, speller, ignored, cache, paragraphs));
    return { issues, decorations: DecorationSet.create(doc, issues.map((i) => Decoration.inline(i.from, i.to, { class: "pw-misspelled" }))) };
  };
  return new Plugin<SpellState>({
    key: spellKey,
    state: {
      init: (_, state) => build(state.doc),
      apply(tr, old) {
        // A full re-check (a new dictionary, a word ignored or added); after typing, only what changed.
        if (tr.getMeta(spellKey) === "recheck") { cache.clear(); paragraphs = new WeakMap(); return build(tr.doc); }
        if (tr.getMeta(spellKey) === "update") return build(tr.doc);
        if (!tr.docChanged) return old;
        // Keep underlines in place while typing; the full re-check follows when typing pauses.
        return { issues: old.issues, decorations: old.decorations.map(tr.mapping, tr.doc) };
      },
    },
    props: { decorations: (state) => spellKey.getState(state)?.decorations },
    view(view) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      return {
        update(v, prev) {
          if (v.state.doc === prev.doc) return;
          clearTimeout(timer);
          timer = setTimeout(() => v.dispatch(v.state.tr.setMeta(spellKey, "update").setMeta("addToHistory", false)), 400);
        },
        destroy() { clearTimeout(timer); },
      };
    },
  });
}

export function recheck(view: EditorView) {
  view.dispatch(view.state.tr.setMeta(spellKey, "recheck").setMeta("addToHistory", false));
}

/** The misspelling at a document position, if any. */
export function misspellingAt(view: EditorView, pos: number): Misspelling | undefined {
  return spellKey.getState(view.state)?.issues.find((i) => i.from <= pos && pos <= i.to);
}
