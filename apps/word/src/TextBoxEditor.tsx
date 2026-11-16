import { floatElementAt, storyFromProseMirror, storyToProseMirror, toggleRun, typingParser } from "@prime/editor";
import { textArea } from "@prime/layout";
import { DOMParser as PmDOMParser } from "prosemirror-model";
import type { FloatInfo } from "@prime/pdm";
import { baseKeymap } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { EditorState, NodeSelection, type Plugin, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { useEffect, useRef } from "react";

/**
 * Typing in a text box (double-click it): its text in an editor laid exactly over the box, in the
 * document's fonts and styles. Esc, a click outside or scrolling ends it; the text is then the box's
 * (marked edited, so it is written into the box when the document is saved). One undo step.
 */
export function TextBoxEditor({ view, pos, box, zoom, plugins = [], initialText, onDone }: {
  view: EditorView; pos: number; box: HTMLElement; zoom: number; plugins?: Plugin[];
  /** Typed first (the key that started typing in a selected shape). */
  initialText?: string;
  onDone(): void;
}) {
  const mount = useRef<HTMLDivElement>(null);
  const holder = useRef<HTMLDivElement>(null);
  // Typed once (the editor may be set up twice: React's development run, a document shown again).
  const pending = useRef(initialText);
  useEffect(() => {
    const node = view.state.doc.nodeAt(pos);
    const float = node?.attrs.float as FloatInfo | null;
    if (!node || !float || !mount.current || !holder.current) { onDone(); return; }
    // As the page lays it out: inside the shape's own text area (an oval's middle) and inner margins,
    // at the top, in the middle or at the bottom.
    const inner = float.insets ?? { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 };
    const area = float.shape ? textArea(float.shape.geometry, float.width, float.height) : { left: 0, top: 0, right: 0, bottom: 0 };
    const insets = {
      left: inner.left + area.left * float.width, right: inner.right + area.right * float.width,
      top: inner.top + area.top * float.height, bottom: inner.bottom + area.bottom * float.height,
    };
    // Its unturned box (a turned shape is typed in straight).
    const r = box.getBoundingClientRect();
    const [w, h] = [box.offsetWidth * (zoom / 100), box.offsetHeight * (zoom / 100)];
    const turned = !!float.rotation;
    Object.assign(holder.current.style, { left: `${turned ? r.left + r.width / 2 - w / 2 : r.left}px`, top: `${turned ? r.top + r.height / 2 - h / 2 : r.top}px` });
    Object.assign(mount.current.style, {
      zoom: String(zoom / 100), width: `${float.width}pt`, minHeight: `${Math.max(12, float.height)}pt`,
      padding: `${insets.top}pt ${insets.right}pt ${insets.bottom}pt ${insets.left}pt`,
      boxSizing: "border-box", display: "flex", flexDirection: "column",
      justifyContent: float.anchor === "ctr" ? "center" : float.anchor === "b" ? "flex-end" : "flex-start",
    });
    // The box's own drawn text hidden while typing over it (also when the page draws it again meanwhile).
    const hide = () => {
      const now = floatElementAt(view, pos) ?? box;
      if (!now.classList.contains("pw-float-editing")) now.classList.add("pw-float-editing");
    };
    hide();
    const redrawn = new MutationObserver(hide);
    redrawn.observe(view.dom, { childList: true, subtree: true });
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
    // An empty box types in its paragraph mark's formatting (white text on a dark shape).
    const first = editor.state.doc.firstChild;
    const markRun = (first?.attrs.props as { markRun?: Record<string, unknown> } | undefined)?.markRun;
    const runMark = editor.state.schema.marks.run;
    if (first && !first.content.size && markRun && runMark) editor.dispatch(editor.state.tr.setStoredMarks([runMark.create({ props: markRun })]));
    if (pending.current) {
      editor.dispatch(editor.state.tr.insertText(pending.current));
      pending.current = undefined;
    }
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
          // The shape stays selected (as in Word), ready for its format.
          const tr = view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, float: { ...current.attrs.float, blocks, edited: true } });
          view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
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
      redrawn.disconnect();
      box.classList.remove("pw-float-editing");
      if (!view.isDestroyed) floatElementAt(view, pos)?.classList.remove("pw-float-editing");
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
