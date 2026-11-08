import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useState } from "react";

/** Groups of words of like meaning (from WordNet 3.0), by part of speech, with antonyms. */
interface Group { pos: "n" | "v" | "a" | "r"; words: string[]; antonyms: string[]; /** How used this meaning is, per word (0: the word's main meaning). */ ranks: Map<string, number> }
let loaded: Promise<Map<string, Group[]>> | undefined;

/** The thesaurus, fetched from the app's own files the first time it is used. */
export function thesaurus(): Promise<Map<string, Group[]>> {
  loaded ??= fetch(new URL("./thesaurus/thesaurus.txt", location.href).href).then((r) => r.text()).then((text) => {
    const index = new Map<string, Group[]>();
    for (const line of text.split("\n")) {
      if (!line) continue;
      const [pos, words, antonyms] = line.split("|");
      const ranks = new Map<string, number>();
      const list = words.split(",").map((w) => { const [word, r] = w.split(":"); ranks.set(word.toLowerCase(), r ? Number(r) : 0); return word; });
      const g: Group = { pos: pos as Group["pos"], words: list, antonyms: antonyms ? antonyms.split(",") : [], ranks };
      for (const w of g.words) {
        const k = w.toLowerCase();
        const groups = index.get(k);
        if (groups) groups.push(g); else index.set(k, [g]);
      }
    }
    return index;
  });
  return loaded;
}

/** A word's dictionary forms to look up: itself, then without its endings (WordNet's own rules). */
export function baseForms(word: string): string[] {
  const w = word.toLowerCase().replace(/[’']s$/, "");
  const out = [w];
  const rules: [RegExp, string][] = [
    [/ies$/, "y"], [/ches$/, "ch"], [/shes$/, "sh"], [/sses$/, "ss"], [/xes$/, "x"], [/zes$/, "z"], [/men$/, "man"], [/es$/, "e"], [/es$/, ""], [/s$/, ""],
    [/ied$/, "y"], [/ed$/, "e"], [/ed$/, ""], [/(.)\1ed$/, "$1"], [/ying$/, "ie"], [/ing$/, "e"], [/ing$/, ""], [/(.)\1ing$/, "$1"],
    [/ier$/, "y"], [/iest$/, "y"], [/er$/, "e"], [/er$/, ""], [/est$/, "e"], [/est$/, ""], [/(.)\1er$/, "$1"], [/(.)\1est$/, "$1"], [/ily$/, "y"], [/ly$/, ""],
  ];
  for (const [re, to] of rules) if (re.test(w)) { const b = w.replace(re, to); if (b.length > 1 && !out.includes(b)) out.push(b); }
  return out;
}

const POS: Record<Group["pos"], string> = { n: "noun", v: "verb", a: "adjective", r: "adverb" };

/** The word at (or selected around) the cursor, with where it is. */
function wordAt(view: EditorView): { word: string; from: number; to: number } | undefined {
  const { $from, from, to, empty } = view.state.selection;
  if (!empty) {
    const t = view.state.doc.textBetween(from, to, " ").trim();
    return t && !/\s/.test(t) ? { word: t, from, to } : t ? { word: t, from, to } : undefined;
  }
  const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, "￼");
  const at = $from.parentOffset;
  let a = at;
  let b = at;
  while (a > 0 && /[\p{L}'’-]/u.test(text[a - 1])) a--;
  while (b < text.length && /[\p{L}'’-]/u.test(text[b])) b++;
  if (a === b) return undefined;
  return { word: text.slice(a, b), from: $from.start() + a, to: $from.start() + b };
}

/** A replacement in the original word's case: "Happy" → "Glad", "HAPPY" → "GLAD". */
const inCase = (like: string, word: string) =>
  like === like.toUpperCase() && like.length > 1 ? word.toUpperCase() : like[0] === like[0].toUpperCase() ? word[0].toUpperCase() + word.slice(1) : word;

/**
 * Review → Thesaurus (Shift+F7): words of like meaning for the word at the cursor, by meaning and part
 * of speech, and its opposites; a word clicked is looked up, Insert puts it in place of the word.
 */
export function ThesaurusPane({ view, onClose }: { view: () => EditorView | undefined; onClose(): void }) {
  const [index, setIndex] = useState<Map<string, Group[]>>();
  const [query, setQuery] = useState(() => { const v = view(); return (v && wordAt(v)?.word) ?? ""; });
  const [history, setHistory] = useState<string[]>([]);
  useEffect(() => { void thesaurus().then(setIndex); }, []);
  const found = useMemo(() => {
    if (!index || !query.trim()) return undefined;
    // The word's meanings, most used first.
    for (const form of baseForms(query.trim())) { const g = index.get(form); if (g) return { form, groups: [...g].sort((a, b) => (a.ranks.get(form) ?? 9) - (b.ranks.get(form) ?? 9)) }; }
    return { form: query.trim().toLowerCase(), groups: [] as Group[] };
  }, [index, query]);
  const look = (w: string) => { setHistory((h) => [...h, query]); setQuery(w); };
  const insert = (w: string) => {
    const v = view();
    if (!v) return;
    const at = wordAt(v);
    if (!at) return;
    v.dispatch(v.state.tr.insertText(inCase(at.word, w), at.from, at.to).scrollIntoView());
    v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, at.from + w.length)));
    v.focus();
  };
  return (
    <aside className="thesaurus-pane" aria-label="Thesaurus">
      <header>
        <strong>Thesaurus</strong>
        <button className="link" onClick={onClose} aria-label="Close thesaurus">✕</button>
      </header>
      <div className="thes-search">
        <button onClick={() => { const h = [...history]; const prev = h.pop(); if (prev !== undefined) { setHistory(h); setQuery(prev); } }} disabled={!history.length} aria-label="Back" title="Back">←</button>
        <input value={query} placeholder="Look up a word" aria-label="Look up" onChange={(e) => setQuery(e.target.value)} />
        <button onClick={() => { const v = view(); const w = v && wordAt(v)?.word; if (w) look(w); }} title="Look up the word at the cursor">At cursor</button>
      </div>
      {!index ? <p className="thes-note">Loading the thesaurus…</p>
        : !found ? <p className="thes-note">Put the cursor in a word, or type one above.</p>
        : !found.groups.length ? <p className="thes-note">No words found for “{query}”.</p>
        : (
          <div className="thes-groups">
            {found.form !== query.trim().toLowerCase() && <p className="thes-note">Showing “{found.form}”.</p>}
            {found.groups.map((g, i) => {
              const others = g.words.filter((w) => w.toLowerCase() !== found.form);
              return (
                <section key={i}>
                  <h4>{(others[0] ?? g.words[0])} <span>({POS[g.pos]})</span></h4>
                  <ul>
                    {others.map((w) => (
                      <li key={w}><button className="thes-word" onClick={() => look(w)} title={`Look up “${w}”`}>{w}</button>
                        <button className="thes-insert" onClick={() => insert(w)} title={`Put “${w}” in place of the word`} aria-label={`Insert ${w}`}>Insert</button></li>
                    ))}
                  </ul>
                  {g.antonyms.length > 0 && (
                    <p className="thes-ant">Opposite: {g.antonyms.map((w, k) => <span key={w}>{k ? ", " : ""}<button className="thes-word" onClick={() => look(w)}>{w}</button></span>)}</p>
                  )}
                </section>
              );
            })}
          </div>
        )}
      <p className="thes-credit">From WordNet 3.0, Princeton University.</p>
    </aside>
  );
}
