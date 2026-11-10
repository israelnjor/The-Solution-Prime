import { storyFromProseMirror, storyToProseMirror, toggleRun, typingParser } from "@prime/editor";
import { DOMParser as PmDOMParser } from "prosemirror-model";
import type { FloatInfo } from "@prime/pdm";
import { baseKeymap } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { EditorState, type Plugin, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { useEffect, useRef } from "react";

/**
 * Typing in a text box (double-click it): its text in an editor laid exactly over the box, in the
 * document's fonts and styles. Esc, a click outside or scrolling ends it; the text is then the box's
 * (marked edited, so it is written into the box when the document is saved). One undo step.
 */
export function TextBoxEditor({ view, pos, box, zoom, plugins = [], onDone }: {
  view: EditorView; pos: number; box: HTMLElement; zoom: number; plugins?: Plugin[]; onDone(): void;
}) {
  const mount = useRef<HTMLDivElement>(null);
  const holder = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const node = view.state.doc.nodeAt(pos);
    const float = node?.attrs.float as FloatInfo | null;
    if (!node || !float || !mount.current || !holder.current) { onDone(); return; }
    const insets = float.insets ?? { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 };
    const r = box.getBoundingClientRect();
    Object.assign(holder.current.style, { left: `${r.left}px`, top: `${r.top}px` });
    Object.assign(mount.current.style, {
      zoom: String(zoom / 100), width: `${float.width}pt`, minHeight: `${Math.max(12, float.height)}pt`,
      padding: `${insets.top}pt ${insets.right}pt ${insets.bottom}pt ${insets.left}pt`,
    });
    box.classList.add("pw-float-editing");
    let done = false;
    const editor = new EditorView(mount.current, {
      // Typing read with the document's own rules; pasting with all of them (web formatting too).
      domParser: typingParser(view.state.schema), clipboardParser: PmDOMParser.fromSchema(view.state.schema),
      state: EditorState.create({
        doc: storyToProseMirror(float.blocks ?? [], view.state.schema),
        plugins: [
          history(),
          keymap({
            "Mod-z": undo, "Mod-y": redo, "Mod-Shift-z": redo,
            "Mod-b": toggleRun("bold"), "Mod-i": toggleRun("italic"), "Mod-u": toggleRun("underline"),
            Escape: () => { finish(); return true; },
          }),
          keymap(baseKeymap),
          // The document's paragraph styles (fonts, spacing), as in the document.
          ...plugins,
        ],
      }),
      attributes: { class: "pw-editor pw-textbox-editor", spellcheck: "false", "aria-label": "Text box" },
    });
    editor.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)));
    editor.focus();
    // The text goes into the box (the document keeps it as one change).
    const save = () => {
      if (done || view.isDestroyed) return;
      done = true;
      const current = view.state.doc.nodeAt(pos);
      if (current && current.attrs.float) {
        const blocks = storyFromProseMirror(editor.state.doc);
        const before = JSON.stringify((current.attrs.float as FloatInfo).blocks ?? []);
        if (JSON.stringify(blocks) !== before) {
          view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, float: { ...current.attrs.float, blocks, edited: true } }));
        }
      }
    };
    // Done typing (Esc, a click outside, scrolling): saved, and the editor goes.
    const finish = () => {
      save();
      onDone();
      if (!view.isDestroyed) view.focus();
    };
    const outside = (e: MouseEvent) => { if (!holder.current?.contains(e.target as Node)) finish(); };
    const scrolled = (e: Event) => { if (!holder.current?.contains(e.target as Node)) finish(); };
    const arm = setTimeout(() => window.addEventListener("mousedown", outside, true), 0);
    window.addEventListener("scroll", scrolled, true);
    // Taken away (another document shown, or React's double run in development): what was typed is kept.
    return () => {
      clearTimeout(arm);
      window.removeEventListener("mousedown", outside, true);
      window.removeEventListener("scroll", scrolled, true);
      box.classList.remove("pw-float-editing");
      save();
      editor.destroy();
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div ref={holder} className="textbox-editor" role="dialog" aria-label="Text box">
      <div ref={mount} />
    </div>
  );
}
