/**
 * Show ¶ (formatting marks): paragraph ends, tabs and breaks are drawn by CSS on the `pw-show-marks`
 * class; spaces need a mark each, so only the paragraphs on screen get them (refreshed on scroll),
 * which keeps long documents fast.
 */
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";

interface MarksState { on: boolean; decorations: DecorationSet }
export const showMarksKey = new PluginKey<MarksState>("showMarks");

/** Turns formatting marks on or off. */
export function setShowMarks(view: EditorView, on: boolean) {
  view.dispatch(view.state.tr.setMeta(showMarksKey, { on }).setMeta("addToHistory", false));
}

/** The visible part of the document: spaces there, as decorations. */
function spacesOnScreen(view: EditorView): DecorationSet {
  const box = view.dom.getBoundingClientRect();
  const top = Math.max(box.top, 0);
  const bottom = Math.min(box.bottom, innerHeight);
  if (bottom <= top) return DecorationSet.empty;
  const x = box.left + Math.min(40, box.width / 2);
  const from = view.posAtCoords({ left: x, top: top + 1 })?.pos ?? 0;
  const to = view.posAtCoords({ left: x, top: bottom - 1 })?.pos ?? view.state.doc.content.size;
  const decos: Decoration[] = [];
  // A little beyond the screen either way, so short scrolls do not show unmarked text.
  const a = Math.max(0, from - 3000);
  const b = Math.min(view.state.doc.content.size, to + 3000);
  view.state.doc.nodesBetween(a, b, (node, pos) => {
    if (!node.isText) return true;
    const t = node.text!;
    for (let i = t.indexOf(" "); i >= 0; i = t.indexOf(" ", i + 1)) decos.push(Decoration.inline(pos + i, pos + i + 1, { class: "pw-space" }));
    return false;
  });
  return DecorationSet.create(view.state.doc, decos);
}

export function showMarks(): Plugin<MarksState> {
  return new Plugin<MarksState>({
    key: showMarksKey,
    state: {
      init: () => ({ on: false, decorations: DecorationSet.empty }),
      apply(tr, old) {
        const meta = tr.getMeta(showMarksKey) as { on?: boolean; decorations?: DecorationSet } | undefined;
        const on = meta?.on ?? old.on;
        if (!on) return { on, decorations: DecorationSet.empty };
        if (meta?.decorations) return { on, decorations: meta.decorations };
        return { on, decorations: tr.docChanged ? old.decorations.map(tr.mapping, tr.doc) : old.decorations };
      },
    },
    props: {
      decorations: (state) => showMarksKey.getState(state)?.decorations,
      attributes: (state): Record<string, string> => (showMarksKey.getState(state)?.on ? { class: "pw-show-marks" } : {}),
    },
    view(view) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      const refresh = () => {
        clearTimeout(timer);
        timer = setTimeout(() => {
          if (!showMarksKey.getState(view.state)?.on) return;
          view.dispatch(view.state.tr.setMeta(showMarksKey, { decorations: spacesOnScreen(view) }).setMeta("addToHistory", false));
        }, 80);
      };
      // Scrolling happens in the window's scrolling ancestor; listen in the capture phase to catch it.
      window.addEventListener("scroll", refresh, true);
      return {
        update(v, prev) {
          const now = showMarksKey.getState(v.state);
          const was = showMarksKey.getState(prev);
          // Turned on, or text typed (new spaces appear marked as soon as the typing pauses).
          if (now?.on && (!was?.on || v.state.doc !== prev.doc)) refresh();
        },
        destroy() { clearTimeout(timer); window.removeEventListener("scroll", refresh, true); },
      };
    },
  });
}
