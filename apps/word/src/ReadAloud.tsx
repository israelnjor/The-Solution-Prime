import type { Node } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { useEffect, useRef, useState } from "react";
import { scrollToPos } from "./NavPane";

/** The word being read, highlighted in the document. */
export const speakKey = new PluginKey<DecorationSet>("speak");
export function speakPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: speakKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set) {
        const m = tr.getMeta(speakKey) as { from: number; to: number } | null | undefined;
        if (m === null) return DecorationSet.empty;
        if (m) return DecorationSet.create(tr.doc, [Decoration.inline(m.from, m.to, { class: "pw-speaking" })]);
        return set.map(tr.mapping, tr.doc);
      },
    },
    props: { decorations: (state) => speakKey.getState(state) },
  });
}

/** A paragraph's text to speak, with each character's document position. */
function spoken(node: Node, pos: number): { text: string; at: number[] } {
  let text = "";
  const at: number[] = [];
  node.forEach((child, offset) => {
    // Deleted text (tracked) is not read.
    if (child.marks.some((m) => m.type.name === "revision" && m.attrs.kind === "del")) return;
    const s = child.isText ? child.text! : " ";
    for (let i = 0; i < s.length; i++) at.push(pos + 1 + offset + i);
    text += s;
  });
  return { text, at };
}

/** The next paragraph with words in it, from `from` on. */
function nextParagraph(doc: Node, from: number): { node: Node; pos: number } | undefined {
  let found: { node: Node; pos: number } | undefined;
  doc.nodesBetween(from, doc.content.size, (node, pos) => {
    if (found) return false;
    if (node.isTextblock) {
      if (pos + node.nodeSize > from && /\S/.test(node.textContent)) found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

const VOICE = "prime-word:voice";
const RATE = "prime-word:speech-rate";
const stored = (k: string) => { try { return localStorage.getItem(k) ?? ""; } catch { return ""; } };
const store = (k: string, v: string) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } };

/**
 * Review → Read Aloud (Ctrl+Alt+Space): reads the document from the cursor, paragraph by paragraph,
 * the word being read highlighted. Only the voices installed on this computer are offered, so the text
 * never leaves it.
 */
export function ReadAloud({ view, onClose }: { view: () => EditorView | undefined; onClose(): void }) {
  const synth = typeof speechSynthesis !== "undefined" ? speechSynthesis : undefined;
  const [voices, setVoices] = useState<SpeechSynthesisVoice[]>(() => synth?.getVoices().filter((v) => v.localService) ?? []);
  const [voice, setVoice] = useState(() => stored(VOICE));
  const [rate, setRate] = useState(() => Number(stored(RATE)) || 1);
  const [state, setState] = useState<"stopped" | "playing" | "paused">("stopped");
  // The paragraph being read (where reading goes on from); a counter to ignore utterances of a stopped run.
  const at = useRef(-1);
  const run = useRef(0);

  useEffect(() => {
    if (!synth) return;
    const load = () => setVoices(synth.getVoices().filter((v) => v.localService));
    load();
    synth.addEventListener("voiceschanged", load);
    return () => synth.removeEventListener("voiceschanged", load);
  }, [synth]);

  const highlight = (range: { from: number; to: number } | null) => {
    const v = view();
    if (v && !v.isDestroyed) v.dispatch(v.state.tr.setMeta(speakKey, range).setMeta("addToHistory", false));
  };

  const stop = () => {
    run.current++;
    synth?.cancel();
    highlight(null);
    setState("stopped");
  };
  useEffect(() => () => { run.current++; synth?.cancel(); const v = view(); if (v && !v.isDestroyed) v.dispatch(v.state.tr.setMeta(speakKey, null)); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const speakFrom = (from: number, ticket: number) => {
    const v = view();
    if (!synth || !v || ticket !== run.current) return;
    const para = nextParagraph(v.state.doc, from);
    if (!para) { stop(); return; }
    at.current = para.pos;
    const { text, at: positions } = spoken(para.node, para.pos);
    const u = new SpeechSynthesisUtterance(text);
    const chosen = voices.find((x) => x.voiceURI === voice);
    if (chosen) { u.voice = chosen; u.lang = chosen.lang; }
    u.rate = rate;
    u.onboundary = (e) => {
      if (ticket !== run.current || e.name !== "word") return;
      const len = e.charLength || (/^\S+/.exec(text.slice(e.charIndex))?.[0].length ?? 1);
      const a = positions[e.charIndex];
      const b = positions[Math.min(positions.length - 1, e.charIndex + len - 1)];
      if (a === undefined || b === undefined) return;
      highlight({ from: a, to: b + 1 });
      // Keep the word in sight.
      const cur = view();
      const canvas = cur?.dom.closest(".canvas");
      if (cur && canvas) {
        try {
          const c = cur.coordsAtPos(a);
          const box = canvas.getBoundingClientRect();
          if (c.top < box.top + 40 || c.bottom > box.bottom - 40) scrollToPos(cur, a, "center");
        } catch { /* not drawn */ }
      }
    };
    u.onend = () => { if (ticket === run.current) speakFrom(para.pos + para.node.nodeSize, ticket); };
    u.onerror = (e) => { if (ticket === run.current && e.error !== "interrupted" && e.error !== "canceled") stop(); };
    synth.speak(u);
  };

  const play = () => {
    if (!synth) return;
    if (state === "paused") { synth.resume(); setState("playing"); return; }
    const v = view();
    if (!v) return;
    const ticket = ++run.current;
    synth.cancel();
    setState("playing");
    speakFrom(v.state.selection.from, ticket);
  };
  const pause = () => { synth?.pause(); setState("paused"); };
  // A new voice or speed: the paragraph being read starts again with it.
  const restart = () => {
    if (state !== "playing" || at.current < 0) return;
    const ticket = ++run.current;
    synth?.cancel();
    speakFrom(at.current, ticket);
  };
  useEffect(restart, [voice, rate]); // eslint-disable-line react-hooks/exhaustive-deps
  // Start at once when opened.
  useEffect(() => { play(); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className="read-aloud" role="toolbar" aria-label="Read aloud">
      {!synth ? <span>Read Aloud is not available here.</span> : (
        <>
          {state === "playing"
            ? <button onClick={pause} title="Pause" aria-label="Pause">❚❚</button>
            : <button onClick={play} title={state === "paused" ? "Resume" : "Read from the cursor"} aria-label="Play">▶</button>}
          <button onClick={stop} disabled={state === "stopped"} title="Stop" aria-label="Stop">■</button>
          <label className="field">Voice{" "}
            <select value={voice} aria-label="Voice" onChange={(e) => { setVoice(e.target.value); store(VOICE, e.target.value); }}>
              <option value="">Default</option>
              {voices.map((x) => <option key={x.voiceURI} value={x.voiceURI}>{x.name.replace(/^Microsoft /, "").replace(/ - .*$/, "")} ({x.lang})</option>)}
            </select>
          </label>
          <label className="field">Speed{" "}
            <select value={String(rate)} aria-label="Reading speed" onChange={(e) => { setRate(Number(e.target.value)); store(RATE, e.target.value); }}>
              {[0.6, 0.8, 1, 1.2, 1.5, 1.8].map((r) => <option key={r} value={r}>{r === 1 ? "Normal" : `${r}×`}</option>)}
            </select>
          </label>
        </>
      )}
      <button className="link" onClick={() => { stop(); onClose(); }} aria-label="Close read aloud">✕</button>
    </div>
  );
}
